# Phase 1 — Structural text anchors

**Created:** 2026-08-22
**Status:** Complete

## Motivation

`BlockAnchor` currently treats a part-wide paragraph ordinal plus an eight-character text hash as
identity. If traversal order changes and the old ordinal now contains identical text, an anchored
edit can silently write to the wrong object. This phase makes current anchors identify their text
container structurally and validate its complete content fingerprint before mutation, while
retaining a conservative compatibility path for existing three-field anchors.

This is PR 1 of the stack: `gavin/block-anchor-structural-identity`, targeting PR #38's
`codex/pptx-doc-hygiene` branch at exact remote head `a3c1ae6e`. It must be correct, documented,
and green without relying on any later selection-hardening PR.

## Context

**What exists today:**

- `src/pptx/inspect.py` owns `BlockAnchor`, the public pinned eight-character `content_hash()`
  helper, the `paper-text-inspection` v2 payload, and the shared visibility-complete traversal of
  slide shapes, grouped shapes, tables, and blind regions.
- `TextBlock` already reports slide-unique `shape_id` values. Group paths and table coordinates
  are currently presentation strings for readers, not structured targeting data.
- `src/pptx/edit.py` shares inspection's story traversal, but its materialized replacement plan
  carries only part, global block ordinal, and paragraph. `replace_text_at()` resolves the ordinal
  before checking the short hash.
- `replace_text()` can traverse existing notes slides and returns anchors for changed notes
  paragraphs even though `inspect_text()` itself accepts a slide. Those returned notes anchors
  must receive the same current structural contract as slide anchors.
- `refind()` currently scans the entire named part for the short hash. This behavior is safe only
  when the match is exact and unique; direct ordinal resolution must not bypass that ambiguity
  check for legacy anchors.
- `src/pptx/shapes/base.py` documents that shape IDs are unique across all shapes on a slide,
  including descendants of groups. Group-name or z-order lineage would therefore add another
  mutable selector without adding identity.
- `ReplaceResult.to_dict()` embeds serialized anchors under `paper-replace-result` v1. Extending
  current anchor payloads consequently requires this phase to advance that schema as well as the
  inspection schema.
- `CONTRIBUTING.md` requires typed pre-mutation refusal, byte-identical refusal tests, assertions
  after save/reopen, exact changed-part budgets, a frozen regression fixture or golden, and no new
  runtime dependencies.

**What this phase delivers:**

- Current anchors with an explicit anchor version, a structured container locator, a diagnostic
  global block ordinal, and a full SHA-256 content fingerprint.
- Structural locators for top-level shapes, grouped leaf shapes, table cells, and notes text
  shapes, using slide-unique shape IDs and only the container-local coordinates needed to identify
  a paragraph.
- Structural resolution before fingerprint validation for current anchored edits, with no
  ordinal-first fallback.
- Exact-unique recovery for legacy three-field anchors and structurally scoped exact recovery for
  current anchors.
- `paper-text-inspection` v3 and `paper-replace-result` v2 payloads, with deterministic reviewed
  goldens and compatibility coverage.
- Focused regressions proving the original duplicate-text wrong-write cannot recur in shapes,
  groups, tables, or notes.
- Updated inspection, editing, and overview documentation describing identity, staleness,
  recovery, and compatibility accurately.

**Reference files to study before starting:**

- `docs/plans/selection-integrity-hardening/spec.md` — authoritative behavior, interface,
  compatibility, schema, and PR-stack contracts.
- `docs/plans/lossy-selection-audit.md` — LS-01 reproduction, impact, and provenance.
- `CONTRIBUTING.md` — atomicity, fixture, reopen, changed-part, and full quality-gate rules.
- `src/pptx/inspect.py` — `BlockAnchor`, `TextBlock`, inspection schemas, content hashing, and the
  shared story traversal.
- `src/pptx/edit.py` — anchor production, deck-wide planning, anchored resolution, `refind()`, and
  replacement transactions.
- `src/pptx/shapes/base.py` — the slide-wide shape-ID uniqueness contract.
- `src/pptx/errors.py` — existing typed Paper refusal hierarchy; reuse it instead of adding a
  phase-specific exception.
- `tests/paper/test_effective_inspect.py` — inspection schema, deterministic golden, anchor-order,
  normalization, and public hash contracts.
- `tests/paper/test_inspect_text.py` — field and blind-region inspection behavior.
- `tests/paper/test_edit_text.py` — anchored replacement, notes, atomicity, changed-part, recovery,
  and LibreOffice smoke coverage.
- `tests/paper/test_table_ops.py` — anchored table-cell writes and table save/reopen helpers.
- `tests/paper/test_fields_hf.py` — field display-text exclusion and current anchor stability.
- `tests/paper/_authoring/build_fixtures.py`, `tests/paper/fixtures/README.md`, and
  `tests/paper/test_fixture_corpus.py` — frozen fixture creation and corpus discipline.
- `tests/paper/_authoring/update_goldens.py` and `tests/paper/goldens/*.inspect.json` — the only
  supported inspection-golden update path and reviewed serialized contracts.
- `docs/api/inspect.rst`, `docs/api/edit.rst`, and `docs/user/paper-additions.rst` — authoritative
  public documentation for the affected APIs.
- `.github/workflows/test.yml`, `pyproject.toml`, and `Makefile` — CI matrix and repository quality
  commands.

## Steps

### Step 1 — Emit structurally addressed current anchors

**Goal:** Every anchor produced from current package traversal contains enough exact structural
information to locate its paragraph without using part-wide ordinal position, and its content is
represented by a full fingerprint with stable field semantics.

**Work:**

- Extend the public anchor value in `src/pptx/inspect.py` so existing positional construction with
  part, block index, and content hash remains accepted as a legacy anchor, while package-produced
  anchors identify themselves as the current version and serialize a structured locator.
- Keep `part`, `block_index`, and `content_hash` available on current serialized anchors during the
  compatibility period. Treat `block_index` as diagnostic output only for current anchors.
- Define the smallest locator that distinguishes each supported container:
  - top-level and grouped text shapes: owning shape ID plus paragraph index in that shape;
  - table-cell text: graphic-frame shape ID, row, column, and paragraph index in that cell;
  - notes text: notes part, owning shape ID, and paragraph index in that notes shape.
- Use the same slide-unique shape ID for grouped leaf shapes that inspection already reports. Do
  not add group-name paths, group ordinals, z-order, or shape-name matching to the locator.
- Preserve current reader-facing `container_detail` strings, but carry table row and column as
  structured traversal facts rather than reparsing a formatted display string during editing.
- Give blind-region anchors an explicit current container kind so an attempted edit continues to
  produce the existing typed unsupported-structure refusal rather than becoming a malformed or
  legacy anchor.
- Add one private, shared definition of the current fingerprint in `src/pptx/inspect.py`. It must
  hash the complete SHA-256 digest of NFC-normalized literal paragraph content and the ordered
  field-marker types at their content positions. It must retain hard line breaks and meaningful
  whitespace while excluding volatile rendered field values.
- Keep public `content_hash()` byte-for-byte compatible: eight lowercase hexadecimal characters,
  NFC normalization, and no whitespace trimming. Current anchors use the new full-fingerprint
  path; callers invoking `content_hash()` continue to get the legacy value.
- Advance `paper-text-inspection` from v2 to v3 and make current locator serialization
  deterministic. Do not change unrelated effective-font, paragraph-format, or deck-manifest
  schemas.

**Constraints:**

- Structural identity is part plus shape ID plus container-local coordinates. Shape names,
  traversal order, group names, geometry, and text similarity are never identity.
- Paragraph position is valid only within the identified shape or table cell; it must never be
  interpreted as a part-wide paragraph index.
- The fingerprint representation must distinguish fields of different types in different
  positions even when their cached display strings or surrounding literal text match.
- Do not expose a general-purpose matching framework or add a new module. The locator and
  fingerprint belong with inspection, their source of truth.
- Preserve the legacy three-argument `BlockAnchor` construction path without weakening the safety
  behavior delivered in Step 2.
- Do not add runtime dependencies.

**Verification:**

- `uv run pytest -q tests/paper/test_effective_inspect.py tests/paper/test_inspect_text.py tests/paper/test_fields_hf.py`
- `uv run ruff check src/pptx/inspect.py tests/paper/test_effective_inspect.py tests/paper/test_inspect_text.py tests/paper/test_fields_hf.py`
- `git diff --check`

### Step 2 — Resolve structurally and recover conservatively

**Goal:** Current anchored edits can mutate only the exact structurally identified paragraph after
its full fingerprint validates, while legacy anchors and explicit recovery refuse whenever exact
identity cannot be established uniquely.

**Work:**

- Carry the structured owner, table coordinates, and container-local paragraph index through the
  complete preflight plan in `src/pptx/edit.py`, so deck-wide replacement returns current anchors
  for slide shapes, grouped shapes, table cells, and existing notes shapes.
- Resolve a current anchor by named OPC part and supported container kind, require exactly one
  owner with its shape ID, validate table coordinates when present, and then locate the paragraph
  only within that owner. Treat a missing owner, duplicate shape ID, changed table topology,
  out-of-range local paragraph, or mismatched container kind as a typed refusal before any write.
- After structural resolution succeeds, calculate the current full fingerprint from live XML. A
  mismatch raises `StaleAnchorError`; it must not trigger a search or fall back to the diagnostic
  block index. Also require that fingerprint to identify only one paragraph within the resolved
  container. An identical paragraph at another local position makes the target ambiguous and must
  refuse rather than allowing paragraph reorder to recreate the original wrong-write at a smaller
  scope.
- Preserve the existing literal replacement, run-formatting, field-boundary, transaction, and
  changed-part behavior once the target passes structural and fingerprint validation.
- Return a fresh current anchor after every successful deck-wide or anchored edit. Its structural
  locator must still identify the edited paragraph, its diagnostic block index must reflect the
  post-edit traversal, and its fingerprint must represent the post-edit content.
- Recognize a three-field anchor as legacy. Ignore its block index for writes and search only the
  named part for its exact legacy hash: one match may proceed, no matches raise
  `TargetNotFoundError`, and multiple matches raise `AmbiguousTargetError` with candidate block
  indices. No legacy refusal may mutate the presentation.
- Make `refind()` honor the same distinction. For a current anchor, search only inside the
  structurally identified container for an exact, unique full-fingerprint match and return its
  refreshed local and diagnostic coordinates. For a legacy anchor, retain the exact-unique
  part-wide search. Neither path may use nearest text, edit distance, ordering preference, or
  geometry.
- Ensure notes anchors returned by `replace_text(..., include_notes=True)` can be used by
  `replace_text_at()` and `refind()` after save/reopen under the same structural and fingerprint
  rules as slide anchors.
- Advance `paper-replace-result` from v1 to v2 because its `blocks` collection now embeds current
  anchor payloads. Empty and populated results must remain fixed-field and deterministic.

**Constraints:**

- All target validation happens inside the existing package transaction and before the first XML
  mutation. Candidate discovery itself must remain read-only.
- Reuse `StaleAnchorError`, `TargetNotFoundError`, `AmbiguousTargetError`, and
  `UnsupportedStructureError` consistently; do not create an LS-01-specific error hierarchy.
- A matching structural locator never overrides changed content. It is stale, not “close enough.”
- A container-local paragraph ordinal is not sufficient when another paragraph in that container
  has the same full fingerprint. Refuse the duplicate instead of trusting either ordinal.
- A matching hash never overrides missing or ambiguous structural identity for a current anchor.
- Legacy support is intentionally conservative. Do not retain the unsafe ordinal-plus-hash fast
  path, even when the caller supplies an in-range block index.
- Do not broaden anchored editing into unsupported blind or chart content.
- Do not promise identity after a shape is deleted and its ID is reused, a table cell moves to new
  coordinates, or OOXML structure changes beyond what the locator can prove. The safe behavior is
  refusal and re-inspection.

**Verification:**

- `uv run pytest -q tests/paper/test_edit_text.py tests/paper/test_table_ops.py`
- `uv run pytest -q tests/paper/test_edit_text.py -k "anchor or refind or notes or payload"`
- `uv run ruff check src/pptx/edit.py tests/paper/test_edit_text.py tests/paper/test_table_ops.py`
- `git diff --check`

### Step 3 — Freeze the wrong-write regression and document the contract

**Goal:** The original silent wrong-write and every supported compatibility path are permanently
covered, the serialized schema changes are reviewed, and users can understand when an anchor
edits, becomes stale, can be recovered, or must be re-inspected.

**Work:**

- Add one minimal frozen self-generated fixture with two independently identifiable text shapes
  carrying identical literal content. Record honest provenance, duplicate-content ground truth,
  and its hashes in the fixture corpus. Use this fixture for the direct LS-01 reproduction instead
  of creating a large all-purpose anchor fixture.
- Add a regression that captures an anchor for one duplicate-text shape, changes shape traversal
  order so the former global ordinal points at the other shape, and proves the original shape is
  the only one edited after save/reopen. The pre-fix behavior must fail this test by editing the
  other shape.
- Cover current anchors for top-level shapes, grouped leaf shapes with duplicate group or shape
  names, multi-paragraph table cells, and notes shapes. Confirm owner reorder does not redirect an
  edit, while a changed owner, table coordinate, or local paragraph refuses rather than guessing.
- Cover current fingerprint behavior for NFC-equivalent literal text, meaningful whitespace, hard
  breaks, and ordered field markers. Prove cached field display changes do not rot an anchor and a
  field-type or marker-position change does.
- Cover malformed or stale current locators: unknown part, absent shape ID, duplicate shape ID,
  incompatible container kind, invalid table coordinates, out-of-range local paragraph, and full
  fingerprint mismatch. Every refusal test must assert the typed exception and byte-identical
  package state.
- Cover legacy construction and recovery explicitly: unique hash despite a wrong ordinal succeeds;
  absent hash refuses; duplicate hash refuses; unknown part refuses; and `refind()` follows the
  same exact-unique rules.
- Assert successful content only after save/reopen and retain exact changed-part budgets for slide
  and notes edits. Preserve the existing LibreOffice load smoke for replaced output.
- Regenerate all three reviewed text-inspection goldens through
  `tests/paper/_authoring/update_goldens.py`, then review that their only systematic changes are
  inspection v3 and the deterministic current anchor fields.
- Update `docs/api/inspect.rst` and public docstrings to describe current locator fields, the
  diagnostic-only global ordinal, full fingerprint semantics, field markers, schema v3, and the
  unchanged legacy `content_hash()` helper.
- Update `docs/api/edit.rst` and public docstrings to describe structural-first resolution, stale
  content handling, conservative legacy behavior, current and legacy `refind()` scope, notes
  anchors, and replace-result v2.
- Update `docs/user/paper-additions.rst` so the perceive/edit overview no longer claims a short
  content hash alone prevents wrong writes.
- Prepare PR 1 with only LS-01 production code, focused fixture/goldens, tests, and affected docs.
  Its description must identify the finding, state that it targets PR #38's
  `codex/pptx-doc-hygiene` branch, and link PR 2 as the next stack layer once that PR exists.

**Constraints:**

- The frozen fixture must remain narrowly about duplicate text and structural identity. Reuse the
  existing `nested_groups`, `tables_in_group`, and `chart_notes` fixtures for container-specific
  coverage rather than duplicating them.
- Fixture bytes and sidecar must be produced through the repository's authoring workflow and be
  added to `MANIFEST.sha256`; never hand-edit PPTX package bytes.
- Golden changes require human-readable review. Do not run a broad generated-file update unrelated
  to the text-inspection schema.
- Tests must assert behavior and atomicity, not dataclass construction or generic JSON round trips.
- Documentation for Phase 1 ships in PR 1. Phase 7 may reconcile cross-stack wording but is not a
  reason to defer current anchor documentation.
- Do not include layout, placeholder, section, or diff matching changes in this branch.

**Verification:**

- `uv run python tests/paper/_authoring/update_goldens.py`
- `uv run pytest -q tests/paper/test_fixture_corpus.py tests/paper/test_effective_inspect.py tests/paper/test_inspect_text.py tests/paper/test_edit_text.py tests/paper/test_table_ops.py tests/paper/test_fields_hf.py`
- `uv run ruff check src/pptx/inspect.py src/pptx/edit.py tests/paper/test_effective_inspect.py tests/paper/test_inspect_text.py tests/paper/test_edit_text.py tests/paper/test_table_ops.py tests/paper/test_fields_hf.py tests/paper/_authoring/build_fixtures.py`
- `uv run make docs`
- `git diff --check`

## Files

| Action | Path |
|---|---|
| Edit | `src/pptx/inspect.py` — current anchor model, structured traversal facts, full fingerprint, inspection v3, and public documentation |
| Edit | `src/pptx/edit.py` — current anchor production/resolution, conservative legacy recovery, notes support, replace-result v2, and public documentation |
| Edit | `tests/paper/test_effective_inspect.py` — pinned legacy hash, current anchor, schema v3, determinism, and golden contracts |
| Edit | `tests/paper/test_inspect_text.py` — structured locator and field-marker inspection cases |
| Edit | `tests/paper/test_edit_text.py` — wrong-write, structural resolution, legacy, notes, atomicity, payload v2, reopen, and changed-part regressions |
| Edit | `tests/paper/test_table_ops.py` — table-cell coordinate and local-paragraph anchor behavior |
| Edit | `tests/paper/test_fields_hf.py` — field display-value stability and field-marker fingerprint assertions that belong with field fixtures |
| Edit | `tests/paper/test_fixture_corpus.py` — verify the duplicate-text fixture's shape identities and literal-content ground truth |
| Edit | `tests/paper/_authoring/build_fixtures.py` — generate the minimal duplicate-text regression deck |
| Create | `tests/paper/fixtures/self_generated/anchor_duplicate_text.pptx` — frozen LS-01 wrong-write fixture |
| Create | `tests/paper/fixtures/self_generated/anchor_duplicate_text.json` — fixture provenance and ground truth |
| Edit | `tests/paper/fixtures/MANIFEST.sha256` — pin the new fixture and sidecar hashes |
| Edit | `tests/paper/fixtures/README.md` — list the new feature-isolated fixture |
| Edit | `tests/paper/goldens/branded_template.inspect.json` — reviewed text-inspection v3 payload |
| Edit | `tests/paper/goldens/clrmap_remap.inspect.json` — reviewed text-inspection v3 payload |
| Edit | `tests/paper/goldens/gauntlet_slide1.inspect.json` — reviewed text-inspection v3 payload |
| Edit | `docs/api/inspect.rst` — structural locator, fingerprint, compatibility, and schema guidance |
| Edit | `docs/api/edit.rst` — structural resolution, refusals, recovery, notes anchors, and result schema guidance |
| Edit | `docs/user/paper-additions.rst` — accurate perceive/edit safety overview |

## What this phase does NOT include

- Unique-only layout selection, placeholder reconciliation, or section selection; those are
  phases 2 through 4.
- Shape or paragraph matching in `diff_decks()`; those are phases 5 and 6.
- Group-name paths, z-order paths, geometry, fuzzy text, prefix/suffix windows, edit distance, or
  nearest-neighbor fallback in any anchor or recovery path.
- A stronger promise across deleted-and-reused shape IDs, moved table cells, arbitrary paragraph
  rewrites, or other structural changes OOXML cannot identify durably.
- Editing chart text, `mc:AlternateContent`, or any other region already reported as unsupported or
  blind.
- Changes to public `content_hash()` output, inherited python-pptx APIs, or any runtime dependency.
- Unrelated cleanup in the inspection/edit traversal or report types.

## Tests this phase must include

- The frozen duplicate-text reorder reproduction: a current anchor edits its original shape even
  when its old part-wide ordinal now points at an identical paragraph in another shape.
- Shape reorder and paragraph insertion cases proving current write resolution uses shape ID and
  the paragraph's local container coordinates, never the global block ordinal; duplicate exact
  fingerprints inside one container must refuse rather than trusting the local paragraph ordinal.
- Grouped leaf shapes with repeated display names, proving slide-unique leaf shape ID is sufficient
  and group-name lineage is not consulted.
- Table-cell anchors proving graphic-frame ID, row, column, and paragraph-within-cell are all
  enforced; coordinate or topology changes refuse rather than redirect.
- Notes anchors produced by deck-wide replacement, followed by anchored edit and `refind()` after
  save/reopen.
- Full-fingerprint cases for NFC normalization, whitespace, hard breaks, field types and marker
  order, and volatile cached field display text.
- Current-anchor refusal cases for unknown part, absent/duplicate shape identity, incompatible
  container, invalid cell, invalid local paragraph, and stale content.
- Legacy three-field compatibility: unique part-wide exact hash succeeds even with a stale ordinal;
  no match and duplicate match refuse atomically.
- `refind()` cases proving current searches stay within the structural container and legacy
  searches remain exact-unique within the named part.
- Post-edit anchors can immediately drive a follow-up edit and survive save/reopen.
- Inspection v3 and replace-result v2 keys, versions, fixed-field behavior, deterministic ordering,
  and reviewed goldens.
- Successful slide and notes content assertions after save/reopen, exact changed-part budgets, and
  byte-identical packages for every documented refusal.
- Existing formatting preservation, field boundaries, zero-match semantics, blind-region refusal,
  nested-group depth, fixture integrity, and LibreOffice load smoke remain green.

**Does NOT need tests:** generic dataclass equality, serialization round trips that do not assert a
public schema contract, SHA-256 library behavior, or shape-ID uniqueness already guaranteed by the
existing object model. Tests must focus on selection, validation, compatibility, and observable
mutation/refusal behavior.

## Done when

1. Every current anchor emitted by `inspect_text()`, `replace_text()`, `replace_text_at()`, and
   `refind()` contains a deterministic structural locator and full content fingerprint.
2. Top-level, grouped, table-cell, and notes targets resolve by part, slide-unique shape ID,
   container coordinates, and container-local paragraph index as applicable; current edits never
   resolve by global block ordinal.
3. The frozen duplicate-text reorder regression edits only the originally anchored shape after
   save/reopen and would fail against the pre-phase implementation.
4. Structural ambiguity, missing structure, unsupported container kinds, and stale fingerprints
   raise existing typed refusals before mutation, with byte-identical package assertions. An exact
   duplicate fingerprint inside the resolved container is included in structural ambiguity.
5. Legacy three-field anchors remain constructible and can edit only after an exact unique hash
   match in the named part; duplicate or absent matches refuse atomically.
6. `refind()` is exact and conservative for both current and legacy anchors and contains no fuzzy,
   ordinal-preference, shape-name, group-name, or geometric fallback.
7. Public `content_hash()` retains its pinned eight-character NFC SHA-256 behavior, while current
   fingerprints are full SHA-256 and include ordered field markers without volatile field display
   text.
8. `paper-text-inspection` is v3 and `paper-replace-result` is v2; both serialize the same current
   anchor shape deterministically, and all reviewed inspection goldens are updated only for this
   contract.
9. Inspection/edit API docs, public docstrings, and the Paper additions overview explain
   structural identity, diagnostic ordinals, staleness, legacy recovery, notes anchors, and schema
   versions accurately.
10. Focused checks pass:
    `uv run pytest -q tests/paper/test_fixture_corpus.py tests/paper/test_effective_inspect.py tests/paper/test_inspect_text.py tests/paper/test_edit_text.py tests/paper/test_table_ops.py tests/paper/test_fields_hf.py`.
11. Style and docs checks pass:
    `uv run ruff check src/pptx/inspect.py src/pptx/edit.py tests/paper/test_effective_inspect.py tests/paper/test_inspect_text.py tests/paper/test_edit_text.py tests/paper/test_table_ops.py tests/paper/test_fields_hf.py tests/paper/_authoring/build_fixtures.py`, `uv run make docs`, and `git diff --check`.
12. The complete PR gate passes independently at the head of
    `gavin/block-anchor-structural-identity`: `uv run pytest`, `uv run behave`,
    `uv run pytest -m lo_smoke tests/paper`, and `uv run make build`.
13. PR 1 contains only LS-01 production changes, its fixture/goldens/tests, and its affected
    documentation, targets PR #38's `codex/pptx-doc-hygiene` branch, and is ready to serve as the
    GitHub base for PR 2.
