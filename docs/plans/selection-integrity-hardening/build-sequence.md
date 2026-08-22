# Build Sequence: Selection integrity hardening

**Created:** 2026-08-22
**Status:** Executed

## Dependency Graph

<!-- Machine-readable DAG. /execute-plan parses this block. -->
```yaml
phases:
  - id: 1
    name: Structural text anchors
    depends_on: []
  - id: 2
    name: Unique import layout selection
    depends_on: [1]
  - id: 3
    name: Unique placeholder reconciliation
    depends_on: [2]
  - id: 4
    name: Stable section selection
    depends_on: [3]
  - id: 5
    name: Stable shape diff identity
    depends_on: [4]
  - id: 6
    name: Exact text and table diff alignment
    depends_on: [5]
  - id: 7
    name: Documentation hygiene and compatibility
    depends_on: [6]
```

## Context

### What exists today

- `BlockAnchor` stores an OPC part, slide-global block ordinal, and an eight-character text hash.
  `replace_text_at()` resolves the ordinal first and validates only that hash, so duplicate text can
  redirect an edit after structural reorder.
- `inspect_text()` already records slide-unique shape IDs and display-oriented container details.
  Shape IDs are unique across the entire slide, including grouped descendants, so no additional
  group-lineage identity is needed. Table coordinates exist only inside a formatted string.
- Slide import chooses the first destination layout matching name, type, or blank fallback, and bake
  can fall back to the first layout overall.
- Placeholder reconciliation preserves exact type/index matches but chooses the first same-type or
  family-compatible fallback. Direct rebind accepts `placeholder_map`; adopt-theme import does not.
- Section import accepts a name, validates it with one first-match lookup, and performs a second
  first-match lookup during enrollment even though each section already has a stable GUID.
- `diff_decks()` matches slides by permanent slide ID but matches top-level shapes by unique name or
  kind ordinal. Text blocks match by shape ID plus paragraph ordinal; table cell coordinates are
  discarded. Notes are compared separately as one flat string.
- Public report payloads are deterministic and versioned. `ImportReport` serializes a fixed field
  set rather than omitting empty fields; `SlideChange` omits empty facets.

### What this feature requires

- Current text anchors resolve a stable structural container before validating a full content
  fingerprint, while three-field legacy anchors recover only through an exact unique hash match.
- Automatic layout, placeholder, and section matching succeeds only for a unique candidate;
  callers receive explicit `target_layout`, `placeholder_map`, and `section_id` escape hatches.
- Lineage deck diffing uses shape IDs and exact conservative paragraph alignment, reports table grid
  changes separately, and never replaces ambiguity with ordinal or similarity matching.
- Inspection, import, and deck-diff schema versions advance only in their owning PR layers.
- The work ships as six bottom-up GitHub PRs. Each layer is independently green and documented;
  phase 7 is committed into the sixth/top PR rather than creating a seventh PR.

### What this feature explicitly excludes

- Footer-furniture duplicate canonicalization and inherited `SlideLayouts.get_by_name()` behavior.
- Visual, geometric, fuzzy, nearest-neighbor, OCR, or independently authored deck matching.
- Durable identity across deleted-and-reused same-kind shape IDs or arbitrary structure changes
  where OOXML provides no stable identity.
- A per-slide mapping surface on `append_deck()`, a public section collection, or new runtime
  dependencies.
- Table-style effective-format resolution or paragraph-level notes diffing.
- Unrelated refactors, task-named helpers, and changes to upstream python-pptx APIs.

### Cross-cutting rules

- Follow `CONTRIBUTING.md`: validate before mutation, use typed Paper refusals, assert successful
  content after save/reopen, assert byte equality on refusal, and pin exact changed-part budgets.
- Keep helpers private to their consuming module unless a later phase demonstrates real reuse.
  Structural locator data belongs in `inspect.py`; exact paragraph alignment belongs in `diff.py`.
- Keep deck-diff v5's location-aware effective-shift representation inside `diff.py`; do not alter
  the shared rebind/import `RunShift` payload and silently change those earlier report schemas.
- Preserve `content_hash()` as the public eight-character legacy helper. Current anchors use a
  separate full-fingerprint path so existing callers and tests are not silently redefined.
- Use slide-unique shape ID directly for top-level and grouped shapes. Do not add group-name paths,
  z-order paths, or redundant shape-ID lineage.
- Every ambiguity error names the candidates and the explicit public argument that resolves it.
- Public schema changes are deterministic and golden-tested: text inspection v3 and replace result
  v2 in phase 1, import report v2 in phase 3 and v3 in phase 4, deck diff v4 in phase 5 and v5 in
  phase 6.
- Each PR contains the focused docs for its own behavior. Phase 7 performs the final cross-surface
  sweep and full compatibility gate; it does not defer missing API documentation from earlier PRs.
- No architecture, `AGENTS.md`, `CLAUDE.md`, `.claude/rules`, or `.claude/skills` files exist in this
  repository. Do not create them solely for this feature.
- CI compatibility spans Python 3.9 through 3.13. Run Ruff on changed Python, but do not introduce a
  new mandatory type-check gate or dependency.

### Reference files

- `docs/plans/selection-integrity-hardening/spec.md` — complete behavior, interface, stack, and
  schema contracts.
- `docs/plans/lossy-selection-audit.md` — reproductions, severity, provenance, and rejected lossy
  matching approaches.
- `CONTRIBUTING.md` — atomicity, fixture, reopen, changed-part, compatibility, and quality rules.
- `README.md` — public refusal and perceive/edit/compose/verify contract.
- `src/pptx/inspect.py` — text traversal, `BlockAnchor`, `TextBlock`, schemas, and content hashing.
- `src/pptx/edit.py` — deck-wide and anchored replacement, story traversal, and `refind()`.
- `src/pptx/compose.py` — import reports, validation, layout selection, placeholder preparation,
  section enrollment, and append staging.
- `src/pptx/presentation.py` — public import and append method surfaces.
- `src/pptx/rebind.py` — placeholder mapper, rebind reports, and effective-run comparison.
- `src/pptx/slide.py` — public rebind surface and shape/layout ownership.
- `src/pptx/shapes/base.py` — slide-wide shape-ID uniqueness contract.
- `src/pptx/diff.py` — report schemas, lineage matching, shape matching, text diff, effective shifts,
  bullets, notes, and existing slide-order LCS.
- `src/pptx/errors.py` — typed refusal hierarchy and error semantics.
- `tests/paper/test_effective_inspect.py`, `tests/paper/test_edit_text.py`, and
  `tests/paper/test_table_ops.py` — inspection/anchor contracts, wrong-write regression home, and
  table-cell editing coverage.
- `tests/paper/test_import.py` and `tests/paper/test_rebind.py` — import, append, layout,
  placeholder, section, atomicity, and report contracts.
- `tests/paper/test_diff.py` and `tests/paper/goldens/lineage_v1_v2.diff.json` — diff identity,
  schema, determinism, detail-level, and lineage contracts.
- `tests/paper/_authoring/build_fixtures.py`, `tests/paper/fixtures/README.md`, and
  `tests/paper/test_fixture_corpus.py` — frozen fixture generation, provenance, and manifests.
- `docs/api/inspect.rst`, `docs/api/edit.rst`, `docs/api/compose.rst`, `docs/api/rebind.rst`,
  `docs/api/diff.rst`, `docs/api/errors.rst`, and `docs/user/paper-additions.rst` — authoritative
  public documentation surfaces.
- `.github/workflows/test.yml`, `pyproject.toml`, and `Makefile` — exact repository quality gates.

## Phases

### Phase 1: Structural text anchors

Replace ordinal-first current anchors with slide-unique shape or table-cell addressing followed by a
full content fingerprint. Preserve the pinned legacy hash helper and accept old three-field anchors
only through exact unique recovery.

**Steps:**

- Step 1 — Carry structured container and paragraph coordinates through the shared text traversal
  and emit text-inspection v3 anchors without introducing redundant group lineage.
- Step 2 — Resolve current anchors structurally, validate their full fingerprint, and make legacy
  anchors recover conservatively before any edit. Require the fingerprint to be unique within the
  resolved container so a local ordinal cannot break ties between identical paragraphs; advance
  replace-result to v2 because it embeds those anchors.
- Step 3 — Add wrong-write, container, legacy, field, atomicity, schema, and save/reopen regressions;
  update inspect/edit documentation.

**Touches:** `src/pptx/inspect.py`, `src/pptx/edit.py`, focused inspection/edit/table tests, inspection
goldens, inspect/edit API docs, Paper additions overview

**Done when:** duplicate-text reorder cannot redirect an edit; identical paragraphs within one
container refuse rather than relying on their local ordinal; every supported container resolves by
stable identity; ambiguous legacy anchors refuse byte-identically; text-inspection v3 is
deterministic; replace-result v2 carries the same anchor schema; focused pytest, Ruff, Sphinx, and
the full PR quality gate pass.

**Depends on:** nothing

**GitHub PR:** `gavin/block-anchor-structural-identity` targeting PR #38's
`codex/pptx-doc-hygiene` branch at exact remote head `a3c1ae6e`

### Phase 2: Unique import layout selection

Make import layout inference unique at every existing tier and delete the arbitrary first-layout
fallback without changing explicit target or keep-appearance behavior.

**Steps:**

- Step 1 — Collect and classify layout candidates by name, type, and bake blank fallback; accept only
  a unique candidate and produce actionable ambiguity refusals.
- Step 2 — Cover single-slide and atomic whole-deck import across unique, absent, ambiguous,
  explicit, order-changed, and no-layout cases; update compose guidance.

**Touches:** `src/pptx/compose.py`, import tests, compose/import docs

**Done when:** no automatic import returns a collection-order choice; `append_deck()` refuses before
writing when any slide is ambiguous; explicit and unique paths retain their reports; focused pytest,
Ruff, Sphinx, and the full PR quality gate pass.

**Depends on:** phase 1 (GitHub stack order; runtime logic is otherwise independent)

**GitHub PR:** `gavin/import-layout-unique-selection` targeting phase 1's branch

### Phase 3: Unique placeholder reconciliation

Preserve the exact global placeholder pass, require uniqueness in weaker tiers, and expose the
existing explicit map contract through adopt-theme import with import-report v2.

**Steps:**

- Step 1 — Refuse multiple same-type or family candidates without changing exact matches,
  one-to-one explicit-map validation, or orphan policy.
- Step 2 — Thread `placeholder_map` through public import validation and preparation, serialize the
  resolved map in every import report, and keep `append_deck()` automatic-only.
- Step 3 — Add direct rebind, adopt-theme import, append atomicity, schema, save/reopen, and
  changed-part regressions; update rebind/compose guidance.

**Touches:** `src/pptx/rebind.py`, `src/pptx/compose.py`, `src/pptx/presentation.py`, direct rebind and
import tests, import goldens, rebind/compose docs

**Done when:** an ambiguous fallback never claims the first slot; callers can settle an import with
the same map semantics as direct rebind; import-report v2 always records the resolved map; focused
pytest, Ruff, Sphinx, and the full PR quality gate pass.

**Depends on:** phase 2

**GitHub PR:** `gavin/placeholder-rebind-unique-selection` targeting phase 2's branch

### Phase 4: Stable section selection

Resolve a destination section once by unique exact name or stable GUID and carry that resolved
identity through enrollment and import-report v3.

**Steps:**

- Step 1 — Add one conservative section resolver covering mutually exclusive name/ID selection,
  no-match behavior, and duplicate-name ambiguity.
- Step 2 — Pass the resolved section through preflight and enrollment, add the public `section_id`
  surface and report field, and preserve adjacent enrollment when neither selector is supplied.
- Step 3 — Add duplicate, GUID, validation, atomicity, enrollment, schema, and save/reopen tests;
  update compose guidance.

**Touches:** `src/pptx/compose.py`, `src/pptx/presentation.py`, import/section tests, import report
goldens, compose docs

**Done when:** validation and enrollment cannot resolve different sections; duplicate names refuse
before mutation; GUID selection is exact and auditable; import-report v3 preserves v2 fields;
focused pytest, Ruff, Sphinx, and the full PR quality gate pass.

**Depends on:** phase 3

**GitHub PR:** `gavin/section-selection-identity` targeting phase 3's branch

### Phase 5: Stable shape diff identity

Replace unique-name/kind-ordinal matching with slide-unique shape IDs and make every shape-oriented
deck-diff facet carry a structured identity in schema v4.

**Steps:**

- Step 1 — Match top-level lineage shapes by ID, treat incompatible-kind ID reuse as remove/add,
  and preserve the declared group-boundary behavior.
- Step 2 — Introduce one structured shape reference across additions, removals, geometry, images,
  and chart changes; update deterministic serialization and consumers.
- Step 3 — Add duplicate-name reorder, z-order, add/remove, renamed, reused-ID, schema, and golden
  regressions; update diff guidance.

**Touches:** `src/pptx/diff.py`, diff/walkthrough tests, deck-diff golden, diff docs

**Done when:** z-order and duplicate names cannot fabricate changes; all real shape facets retain
stable ID plus display name; deck-diff v4 is deterministic; focused pytest, Ruff, Sphinx, and the
full PR quality gate pass.

**Depends on:** phase 4

**GitHub PR:** `gavin/shape-diff-stable-identity` targeting phase 4's branch

### Phase 6: Exact text and table diff alignment

Replace block-ordinal pairing with a small exact alignment boundary in `diff.py`, using phase 1's
structured container data and phase 5's shape identity. Report table grid changes separately and
advance deck-diff to v5 without changing flat notes comparison or table-style resolution.

**Steps:**

- Step 1 — Align exact paragraph fingerprints within stable shape/table containers, leaving repeated
  ambiguous content unmatched rather than forcing ordinal or fuzzy pairs.
- Step 2 — Preserve table before/after coordinates, add structural grid-change reporting at the
  structure detail level, and reuse aligned paragraph pairs for existing supported text/effective/
  bullet facets without widening their domain. Use a deck-diff-owned location-aware effective
  shift record rather than changing the shared rebind/import payload.
- Step 3 — Add paragraph insertion/deletion, duplicate, field, group, and eval-observed table
  expansion regressions; update schema v5, goldens, consumers, and focused diff guidance.

**Touches:** `src/pptx/diff.py`, narrowly required inspection metadata from phase 1, diff/walkthrough
tests, deck-diff golden, diff docs

**Done when:** paragraph or table-grid insertion does not cascade into fictional replacements;
ambiguity is explicit; table structure appears at the correct detail level; notes and unsupported
table formatting remain unchanged; focused pytest, Ruff, Sphinx, and the full PR quality gate pass.

**Depends on:** phase 5 and phase 1 transitively

**GitHub PR:** `gavin/text-diff-exact-alignment` targeting phase 5's branch

### Phase 7: Documentation hygiene and compatibility

Finish the top PR by reconciling cross-surface language, removing stale claims about hash and
ordinal safety, validating stack scope, and running the complete compatibility and distribution
gates. This phase creates no seventh PR.

**Steps:**

- Step 1 — Sweep README, public docstrings, RST, cross-references, and planning status for stale
  first-match, content-hash-only, synthetic ordinal, and schema-version guidance.
- Step 2 — Verify every GitHub stack comparison contains only its owning change and that branch/base
  metadata and PR stack links are correct.
- Step 3 — Run full pytest, Behave, LibreOffice smoke, Ruff, Sphinx, build/twine, and diff checks;
  route failures back to their owning phase.

**Touches:** README and affected public docs/docstrings, planning status, GitHub PR metadata; no new
architecture or agent-instruction files

**Done when:** public docs explain the shipped system without the spec; stale search terms are gone
or explicitly historical; every stack layer is independently reviewable; all repository quality
gates pass.

**Depends on:** phase 6

**GitHub PR:** committed into `gavin/text-diff-exact-alignment` before the top PR is approved
