# Phase 3 — Unique placeholder reconciliation

**Created:** 2026-08-22
**Status:** Complete

**Feature:** Selection integrity hardening

**Plan slug:** `selection-integrity-hardening`

**Depends on:** Phase 2 — Unique import layout selection

**GitHub PR:** `gavin/placeholder-rebind-unique-selection`, targeting
`gavin/import-layout-unique-selection`

## Motivation

Placeholder reconciliation currently protects exact type-and-index matches, but its weaker
same-type and compatible-family tiers silently claim the first available target slot. Two content
placeholders of the same type commonly represent different columns or semantic roles, so collection
order is not a safe identity signal.

This phase makes those fallback tiers unique-only and gives adopt-theme import the explicit mapping
escape hatch already available to direct layout rebinding. It remains a surgical change to the
shared mapper and import boundary: exact matching, explicit-map validation, orphan handling, and
append behavior keep their existing contracts.

## Context

### What exists today

- `src/pptx/rebind.py` owns the shared placeholder matcher. `_compute_mapping()` first applies an
  explicit partial map, then performs global exact type-plus-index matching, same-type matching,
  and compatible-family matching over unclaimed layout slots.
- The global exact pass in `src/pptx/rebind.py` deliberately settles every exact match before any
  weaker pass. This prevents an earlier source placeholder from stealing a later placeholder's
  exact slot and must remain unchanged.
- The same-type and family passes in `src/pptx/rebind.py` currently take the first candidate from an
  ordered list. They neither prove uniqueness nor expose the alternatives in a refusal.
- `Slide.rebind_layout()` in `src/pptx/slide.py` already accepts a partial explicit map. Existing
  validation ensures source and target indexes exist, target claims are one-to-one, and a `None`
  target deliberately orphans that source placeholder. Unlisted source placeholders continue
  through automatic matching.
- Adopt-theme import in `src/pptx/compose.py` calls the same mapper during
  `_validate_mode_preparation()`, before the destination transaction begins, but always requests
  automatic matching. The resulting mapping is already carried in the preparation record and
  applied to the copied slide.
- `Presentation.import_slide()` in `src/pptx/presentation.py` exposes layout selection but no
  placeholder map. `append_deck()` stages every source slide before opening its destination
  transaction and intentionally has no per-slide override surface.
- `ImportReport` in `src/pptx/compose.py` is a deterministic fixed-field payload at
  `paper-import-report` version 1. It does not currently expose the mapping used. `RebindReport`
  already serializes that information and remains version 1 because its payload does not change.
- `tests/paper/test_rebind.py` already pins exact-pass precedence, explicit-map validation,
  orphan/bake behavior, save/reopen results, rollback, and changed-part budgets.
- `tests/paper/test_import.py` already covers all import modes, whole-deck preflight atomicity,
  deterministic reports, save/reopen behavior, package integrity, and LibreOffice loading. The
  frozen report is `tests/paper/goldens/import_beta_keep.import.json`.

### What this phase delivers

- Automatic fallback matching succeeds only when the current tier has exactly one unclaimed
  candidate. Multiple candidates refuse with `AmbiguousTargetError` and actionable source and
  target details.
- The global exact type-plus-index pass and the existing explicit-map validation and orphan policy
  remain authoritative.
- Adopt-theme `import_slide()` accepts the same partial explicit-map meaning as direct rebind and
  records the fully resolved mapping in its report.
- Keep-appearance and bake imports reject a supplied mapping because those modes do not perform
  placeholder reconciliation.
- `append_deck()` remains automatic-only and refuses the whole operation during staging if any
  source slide has ambiguous placeholder reconciliation.
- `paper-import-report` version 2 always serializes `placeholder_map_used`; modes without
  placeholder reconciliation use an empty list.
- Focused direct-rebind and import contracts, report goldens, and public documentation make the
  stricter behavior independently reviewable in the third PR of the stack.

### Reference files to study before starting

- `docs/plans/selection-integrity-hardening/spec.md` — authoritative LS-03 behavior, public
  interface, schema, atomicity, and stack contracts.
- `docs/plans/selection-integrity-hardening/build-sequence.md` — phase boundary, dependency order,
  and cross-cutting implementation rules.
- `docs/plans/lossy-selection-audit.md` — the first-compatible reproduction and why positional
  choice is unsafe.
- `CONTRIBUTING.md` — validate-before-mutate, typed-refusal, save/reopen, fixture, changed-part,
  and compatibility requirements.
- `src/pptx/rebind.py` — `_compute_mapping()`, placeholder-family definitions, explicit-map
  validation, orphan policy, transaction boundary, and `RebindReport` serialization.
- `src/pptx/compose.py` — import argument validation, adopt-theme preparation, copied-placeholder
  reconciliation, append staging, `ImportReport`, and schema version ownership.
- `src/pptx/presentation.py` — public `Presentation.import_slide()` and `append_deck()` surfaces and
  docstrings.
- `src/pptx/slide.py` — public `Slide.rebind_layout()` map semantics and documentation.
- `src/pptx/errors.py` — `AmbiguousTargetError` and the typed-refusal contract.
- `tests/paper/test_rebind.py` — primary direct-mapper and rebind regression home.
- `tests/paper/test_import.py` — primary import, append, report, atomicity, save/reopen, and
  LibreOffice regression home.
- `tests/paper/contract.py` — unchanged-package, changed-part, and save/reopen helpers.
- `tests/paper/goldens/import_beta_keep.import.json` — frozen fixed-field import-report contract.
- `tests/paper/_authoring/update_goldens.py` — approved mechanism for regenerating the import
  golden after review of the schema change.
- `docs/api/rebind.rst`, `docs/api/compose.rst`, and `docs/user/paper-additions.rst` — public
  guidance that must describe unique-only automatic reconciliation and the explicit import map.

## Steps

### Step 1 — Make automatic fallback tiers unique-only

**Goal:** Direct rebind and adopt-theme import never select a same-type or compatible-family
placeholder merely because it appears first, while exact and explicit bindings retain their current
behavior.

**Work:**

- Change the shared mapper in `src/pptx/rebind.py` so each same-type and family fallback evaluates
  all currently unclaimed target slots before deciding.
- Accept a fallback only when that tier has exactly one candidate. If a same-type tier has multiple
  candidates, refuse there rather than dropping to the family tier. If the same-type tier is empty,
  retain the existing transition to the family tier.
- Raise `AmbiguousTargetError` for a non-unique tier. Identify the source placeholder by its type
  and index, enumerate every candidate target type/index pair deterministically, name the tier that
  was ambiguous, and tell the caller to provide `placeholder_map`.
- Preserve the mapper's existing three global passes. Every exact type-plus-index match across the
  slide must settle before any same-type evaluation, and every same-type result must settle before
  family evaluation.
- Preserve explicit partial-map processing before automatic matching. Explicit target claims
  remain one-to-one, `None` continues to deliberately orphan a source placeholder, and only
  unlisted source placeholders enter the automatic passes.
- Preserve the existing behavior for zero candidates: the source placeholder remains unmatched and
  follows the caller's established orphan behavior.
- Keep candidate ordering deterministic by the layout placeholder index already used by the shared
  mapper. Do not use shape order beyond stable error presentation, geometry, names, similarity, or
  a new scoring rule to break ties.

**Constraints:**

- Do not alter the exact global pass, `_TYPE_FAMILIES`, explicit-map validation, or duplicate source
  index refusal except where wording must identify the new ambiguity escape hatch.
- Do not mutate slide or layout XML while computing a mapping. Both direct rebind and import must
  still fail before their first document write.
- Do not turn an ambiguous tier into an orphan. Ambiguity is a distinct `AmbiguousTargetError`, not
  `UnsupportedStructureError` and not a signal to try a weaker tier.
- Do not require a complete explicit map. Existing callers may continue supplying only the source
  entries that need overrides and let unique automatic matching settle the rest.
- Keep the selection rule inside `src/pptx/rebind.py`; import must continue using the same mapper
  rather than maintaining a second implementation.

**Verification:**

```bash
uv run pytest -q tests/paper/test_rebind.py -k "mapping or placeholder or family or exact"
uv run pytest -q tests/paper/test_import.py -k "adopt_theme and placeholder"
uv run ruff check src/pptx/rebind.py tests/paper/test_rebind.py tests/paper/test_import.py
```

### Step 2 — Expose import mapping and advance the report to version 2

**Goal:** A caller can resolve adopt-theme ambiguity explicitly and can audit the complete mapping
used by every successful import report without expanding `append_deck()` into a per-slide mapping
API.

**Work:**

- Extend the public method in `src/pptx/presentation.py` and the module entry point in
  `src/pptx/compose.py` with the same automatic-or-partial-explicit map contract used by
  `Slide.rebind_layout()`.
- Validate that an explicit map applies only to adopt-theme import. Reject a supplied map for
  keep-appearance and bake before transplant planning or destination mutation because those modes
  do not use placeholder reconciliation.
- Pass the caller's map into adopt-theme preparation in `src/pptx/compose.py`. Continue using the
  precomputed resolved mapping when reconciling the copied slide, so validation and mutation cannot
  independently choose different slots.
- Keep explicit-map validation owned by the shared mapper: source indexes must exist, target
  indexes must exist on the selected destination layout, target claims remain one-to-one, and
  `None` means the source placeholder is intentionally orphaned and follows adopt-theme's existing
  bake behavior.
- Leave `Presentation.append_deck()` and `pptx.compose.append_deck()` without a mapping parameter.
  Each staged slide must continue to request automatic mapping, and any ambiguity must abort
  staging before the transaction and before the first imported slide is written.
- Advance only `paper-import-report` from version 1 to version 2 in `src/pptx/compose.py`. Add a
  fixed `placeholder_map_used` field carrying the complete resolved source-to-target mapping,
  including deliberate or automatic `None` outcomes, in deterministic source-index order.
- Always serialize the new field. Adopt-theme reports contain the resolved entries; bake and
  keep-appearance reports serialize an empty list because no reconciliation map applies.
- Keep `paper-rebind-report` at version 1. Its existing `placeholder_map_used` field and serialized
  shape do not change.
- Regenerate `tests/paper/goldens/import_beta_keep.import.json` through the repository's authoring
  helper and review the diff so the only schema changes are version 2 and the required empty
  mapping list.

**Constraints:**

- Do not add an orphan-policy argument to import. Adopt-theme retains its current behavior of
  baking unmatched or explicitly orphaned placeholders from source-resolved values.
- Do not add a per-slide map collection to `append_deck()` or silently skip ambiguous slides.
- Do not reuse the version-2 number for later section identity work; phase 4 owns import-report
  version 3 and must retain this phase's field unchanged.
- Preserve report field ordering and deterministic list ordering. Do not omit the field for modes
  where it is empty.
- Preserve the source-nonmutation guarantee and the existing layout, notes, section, relationship,
  run-shift, and package-transaction behavior.

**Verification:**

```bash
uv run pytest -q tests/paper/test_import.py -k "placeholder_map or import_report or deterministic or append_deck"
uv run pytest -q tests/paper/test_rebind.py -k "explicit_map or report"
uv run ruff check src/pptx/compose.py src/pptx/presentation.py tests/paper/test_import.py
```

### Step 3 — Pin public behavior, persistence, and documentation

**Goal:** The third stack PR independently proves unique-only selection, explicit import recovery,
atomic refusal, report versioning, and persisted output, and explains those contracts without
requiring readers to consult the planning spec.

**Work:**

- Add direct mapper and `Slide.rebind_layout()` regressions in `tests/paper/test_rebind.py` for a
  unique same-type fallback, unique family fallback, multiple same-type targets, multiple family
  targets, exact-pass precedence, explicit disambiguation, partial-map automatic completion, and
  deliberate orphaning.
- Prove that ambiguity details include the source type/index, every candidate target type/index,
  the fallback tier, and `placeholder_map`, with deterministic ordering independent of layout
  traversal order.
- For every direct-rebind ambiguity, assert `AmbiguousTargetError` and exact refusal atomicity.
  Preserve the existing orphan refusal and bake tests to prove that no-candidate behavior was not
  conflated with multiple-candidate behavior.
- Add adopt-theme import regressions in `tests/paper/test_import.py` showing automatic ambiguity
  refusal, successful explicit disambiguation, partial explicit maps, explicit `None` orphaning,
  and rejection of maps in keep-appearance and bake modes.
- Save and reopen successful explicit imports, then assert the imported placeholder indexes,
  selected layout binding, text, and nearby unaffected placeholder state. Pin the exact package
  parts changed by the successful import and keep the existing relationship and section integrity
  checks green.
- Add an `append_deck()` regression whose later staged source slide has ambiguous reconciliation.
  Assert the entire destination remains byte-identical and no earlier source slide was imported.
- Assert import reports expose the complete resolved map for automatic and explicit adopt-theme
  imports, include `None` for deliberately orphaned sources, and expose an empty tuple/list through
  the object/serialized forms for keep-appearance and bake.
- Update the frozen report golden and deterministic comparison for schema version 2. Ensure every
  existing report consumer and assertion reads the new fixed field without depending on omitted
  keys.
- Update the `Slide.rebind_layout()` and `Presentation.import_slide()` docstrings plus
  `docs/api/rebind.rst`, `docs/api/compose.rst`, and the compose section of
  `docs/user/paper-additions.rst`. Document exact-first behavior, unique-only fallback tiers,
  partial explicit maps, `None` orphan meaning, adopt-theme-only applicability, automatic-only
  append, actionable ambiguity refusal, and report version 2.
- Reuse the existing template fixtures for focused candidate mutations. Add a new binary fixture
  and its provenance only if an ambiguity or persistence case cannot be represented safely by the
  existing fixture corpus.

**Constraints:**

- Persisted-content assertions must save, reopen, and inspect the reopened presentation.
- Every ambiguity refusal must assert unchanged in-memory state and unchanged serialized package
  bytes through `tests/paper/contract.py`.
- Successful mutation coverage must pin an exact changed-part budget rather than merely asserting
  that the output opens.
- Do not relax existing source-package byte equality, relationship integrity, section integrity,
  or LibreOffice smoke coverage.
- Do not update phase-4 section behavior, phase-2 layout matching, unrelated report schemas, or
  inherited python-pptx APIs in this PR.
- Keep public documentation in RST and docstrings; planning Markdown remains outside the Sphinx
  toctree.

**Verification:**

```bash
uv run pytest -q tests/paper/test_rebind.py tests/paper/test_import.py
uv run pytest -q -m lo_smoke tests/paper/test_rebind.py tests/paper/test_import.py
uv run ruff check src/pptx/rebind.py src/pptx/compose.py src/pptx/presentation.py src/pptx/slide.py tests/paper/test_rebind.py tests/paper/test_import.py
uv run make docs
```

## Files

| Action | Path | Purpose |
| --- | --- | --- |
| Edit | `src/pptx/rebind.py` | Require uniqueness in same-type and family fallback tiers while preserving exact and explicit matching. |
| Edit | `src/pptx/compose.py` | Accept and validate the import map, reuse it during preparation, serialize the resolved map, and advance import-report to version 2. |
| Edit | `src/pptx/presentation.py` | Expose and document the adopt-theme import mapping argument. |
| Edit | `src/pptx/slide.py` | Clarify that direct rebind's automatic fallback tiers are unique-only and ambiguity requires an explicit map. |
| Edit | `tests/paper/test_rebind.py` | Add direct matching, ambiguity, explicit recovery, atomicity, and persistence regressions. |
| Edit | `tests/paper/test_import.py` | Add import and append ambiguity, explicit map, report, save/reopen, part-budget, and atomicity regressions. |
| Edit | `tests/paper/goldens/import_beta_keep.import.json` | Pin `paper-import-report` version 2 and the always-present mapping field. |
| Edit | `docs/api/rebind.rst` | Publish exact-first and unique-only direct rebind behavior. |
| Edit | `docs/api/compose.rst` | Publish adopt-theme explicit mapping, automatic-only append, ambiguity, and report-v2 behavior. |
| Edit | `docs/user/paper-additions.rst` | Update the integrated compose/rebind workflow with the new safety and recovery contracts. |
| Reference; no planned edit | `src/pptx/errors.py` | Reuse the existing `AmbiguousTargetError` type. |
| Reference; no planned edit | `tests/paper/contract.py` | Reuse atomicity, changed-part, and save/reopen helpers. |
| Reference; generated use only | `tests/paper/_authoring/update_goldens.py` | Regenerate the reviewed import golden without changing the generator unless schema generation proves incomplete. |

## What this phase does NOT include

- Structural text anchors from phase 1 or unique layout candidate selection from phase 2.
- Section name or section GUID selection and import-report version 3 from phase 4.
- Shape or paragraph diff matching and deck-diff schema changes from phases 5 and 6.
- Changes to placeholder type families, exact type-plus-index precedence, explicit-map validation,
  or the meaning of `None`.
- A complete-map requirement; explicit maps remain partial overrides.
- A placeholder mapping surface on `append_deck()`.
- A new import orphan policy or a change to adopt-theme's existing orphan baking.
- Geometry-, proximity-, name-, order-, or fuzzy-based tie breaking.
- Footer-furniture canonicalization or inherited `SlideLayouts.get_by_name()` behavior.
- New runtime dependencies, a new matching module, or a general matching framework.
- Changes to inherited python-pptx public APIs or unrelated cleanup.

## Tests this phase must include

### Shared mapper and direct rebind

- The existing global exact type-plus-index regression remains unchanged and green.
- A sole unclaimed same-type target binds automatically.
- Multiple unclaimed same-type targets raise `AmbiguousTargetError` without trying the family tier.
- A sole unclaimed compatible-family target binds automatically when no same-type target exists.
- Multiple unclaimed compatible-family targets raise `AmbiguousTargetError`.
- Candidate enumeration is complete and deterministic and identifies source and target type/index
  pairs plus the `placeholder_map` recovery argument.
- A partial explicit map claims its requested target before automatic passes and leaves remaining
  sources eligible for exact or unique fallback matching.
- An explicit map disambiguates either fallback tier and preserves one-to-one target ownership.
- An explicit `None` retains deliberate orphan semantics under both refuse and bake policies.
- Invalid source indexes, invalid target indexes, duplicate target claims, wrong map types, and
  duplicate source placeholder indexes retain their existing errors and mutation-free behavior.
- Every new typed refusal preserves exact package bytes; every successful rebind is asserted after
  save/reopen and retains its exact changed-part budget.

### Adopt-theme import and append

- Automatic adopt-theme import with multiple same-type candidates refuses before adding a part or
  slide.
- Automatic adopt-theme import with multiple family candidates refuses with the same shared-mapper
  semantics.
- An explicit partial map selects the intended target slot; save/reopen proves the imported slide
  binds to that slot and preserves unaffected placeholder content.
- Explicit `None` bakes the selected source placeholder and records a `None` target in the report.
- Invalid map source/target indexes and duplicate target claims use the existing mapper validation
  and leave source and destination unchanged.
- A supplied map is rejected for keep-appearance and bake imports before mutation.
- `append_deck()` remains automatic-only; ambiguity on a later staged slide leaves the entire
  destination byte-identical and imports none of the earlier slides.
- The source presentation remains byte-identical across successful and refused imports.
- Successful explicit imports save/reopen cleanly, meet an exact changed-part budget, have no
  dangling relationships or section entries, and load in LibreOffice.

### Report and documentation compatibility

- Automatic and explicit adopt-theme reports expose every resolved source mapping in source-index
  order.
- Deliberate and automatic orphan outcomes serialize with a `null` target rather than disappearing.
- Keep-appearance and bake reports always expose an empty object-level mapping and an empty
  serialized list.
- The report schema is exactly `paper-import-report` version 2; all other version-1 fields retain
  their names, order, and meanings.
- The frozen keep-appearance golden changes only for version 2 and the always-present empty mapping
  field, and repeated imports remain deterministic.
- `RebindReport` remains version 1 with its existing serialized payload.
- Sphinx renders the updated direct-rebind and import behavior with warnings treated as errors.
- Ruff passes on every changed Python and test file.

## Done when

1. The shared mapper preserves its global exact pass and accepts a weaker tier only when exactly
   one unclaimed target candidate exists.
2. Every non-unique same-type or compatible-family tier raises actionable
   `AmbiguousTargetError` before document mutation and never falls through or selects by order.
3. Existing explicit-map validation, partial override behavior, one-to-one claims, and `None`
   orphan semantics remain unchanged.
4. Adopt-theme import accepts the direct-rebind map contract, applies the precomputed mapping to
   the copied slide, and refuses maps for modes that do not reconcile placeholders.
5. `append_deck()` exposes no mapping parameter and atomically refuses before its first write when
   any staged slide has ambiguous placeholder reconciliation.
6. Every import report serializes `placeholder_map_used`; adopt-theme records its complete resolved
   mapping, other modes record an empty list, and `paper-import-report` is version 2.
7. Ambiguous direct rebind and import cases preserve exact package bytes; successful mapped edits
   pass save/reopen assertions and exact changed-part budgets.
8. The focused phase gate passes:

   ```bash
   uv run pytest -q tests/paper/test_rebind.py tests/paper/test_import.py
   uv run pytest -q -m lo_smoke tests/paper/test_rebind.py tests/paper/test_import.py
   uv run ruff check src/pptx/rebind.py src/pptx/compose.py src/pptx/presentation.py src/pptx/slide.py tests/paper/test_rebind.py tests/paper/test_import.py
   uv run make docs
   ```

9. The complete independently green PR gate passes:

   ```bash
   uv run pytest
   uv run behave
   uv run pytest -m lo_smoke tests/paper
   uv run make docs
   uv run make build
   ```

10. The GitHub comparison from `gavin/import-layout-unique-selection` to
    `gavin/placeholder-rebind-unique-selection` contains only LS-03 production changes, focused
    tests and golden updates, and this PR's public documentation; its PR description links the
    surrounding stack layers and identifies LS-03 as closed.
