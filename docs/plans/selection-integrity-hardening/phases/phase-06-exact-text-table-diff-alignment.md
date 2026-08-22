# Phase 6 — Exact Text and Table Diff Alignment

**Created:** 2026-08-22
**Status:** Complete

## Motivation

`diff_decks()` currently treats a paragraph's ordinal within its shape as identity, so inserting
one paragraph or one table row can turn unchanged content into a cascade of fictional
replacements. This phase makes the verification report conservative and exact: it aligns content
only inside a stable structural container, reports uncertainty instead of guessing, and gives
table-grid changes their own structural evidence without widening the package's formatting scope.

## Context

**What exists today:**

- `src/pptx/diff.py` matches slides by permanent slide ID, then `_diff_text()` keys inspected text
  blocks by `(shape_id, paragraph ordinal within shape)`. That shape-scoped key avoids slippage
  after a different shape is removed, but it still pairs unrelated paragraphs after an insertion,
  deletion, or reorder inside the same shape.
- `_field_markers()` already distinguishes literal text from field markers and preserves marker
  offsets for reporting. The Phase 1 inspection work provides a full structural paragraph
  fingerprint based on NFC-normalized literal text plus field marker types and order, without
  volatile rendered field values.
- Phase 1 also provides structured container data for ordinary and grouped leaf shapes and for
  table cells: the slide-unique owning shape ID, table row and column, and paragraph index within
  the container. The current display string in `TextBlock.container_detail` is not an identity
  surface and must not be parsed.
- Phase 5 advances `paper-deck-diff` to version 4, matches top-level shape facets by stable shape
  ID, and gives shape-oriented payloads a structured shape reference. Phase 6 builds on that
  identity rather than reintroducing names or synthetic ordinals.
- Full-detail effective shifts currently come from `pptx.rebind._resolution_state(...,
  align_by_content=True)`. That path still embeds the paragraph ordinal in its key, and its
  `RunShift` record is also used by rebind and import reports. Changing the shared `RunShift`
  payload here would silently change schemas owned by earlier PRs.
- Bullet shifts use unique paragraph text within a shape and skip duplicates. This is conservative
  for repeated text, but it is a separate matching rule and cannot follow a uniquely aligned
  paragraph across insertion or movement.
- Table-cell blocks are visibility-complete but deliberately blind for effective font and
  paragraph-format resolution because table-style inheritance is not implemented. Notes are not
  part of slide block alignment; `_notes_text()` compares the entire notes body and emits one flat
  `notes_change`.
- `SlideChange.to_dict()` omits empty facets, the deck-diff golden is regenerated only by
  `tests/paper/_authoring/update_goldens.py`, and the contract suite requires deterministic output
  and an empty self-diff across the entire non-corrupt fixture corpus.

**What this phase delivers:**

- One private, deterministic, exact sequence-alignment boundary in `src/pptx/diff.py`, partitioned
  by Phase 1 structural container identity and shared by all deck-diff paragraph consumers.
- Conservative paragraph outcomes that distinguish insertions, deletions, uniquely bounded
  replacements, exact moves/alignment, and ambiguous regions without using fuzzy similarity or an
  ordinal tie-break.
- Structured before/after paragraph locations for ordinary shapes, grouped leaves, and table
  cells, reusing Phase 5's shape-reference contract.
- A table-structure facet at `detail="structure"` that reports before/after grid dimensions and
  any exactly established row or column change without manufacturing cell replacements.
- Text, field, effective-value, and bullet reporting driven by the same aligned paragraph domain,
  while preserving the existing flat notes comparison and excluding table effective formatting.
- `paper-deck-diff` version 5, a reviewed deterministic golden, focused LS-05 fixtures and
  regressions, updated diff documentation, and an independently green top implementation PR.

**Reference files to study before starting:**

- `docs/plans/selection-integrity-hardening/spec.md` — authoritative LS-05 behavior, report,
  exclusions, and stack contracts.
- `docs/plans/selection-integrity-hardening/build-sequence.md` — Phase 1 and Phase 5 dependencies,
  schema ownership, cross-cutting rules, and the PR boundary.
- `docs/plans/lossy-selection-audit.md` — the ordinal-cascade reproduction, provenance, severity,
  and rejected fuzzy matching approaches.
- `src/pptx/diff.py` — `SlideChange`, `DeckDiff`, detail gates, current ordinal text matcher,
  effective shifts, bullet shifts, notes comparison, and the existing slide-order LCS.
- `src/pptx/inspect.py` at the Phase 1 head — structured text container and paragraph coordinates,
  full structural fingerprint, field representation, table traversal, and blind-region contract.
- `src/pptx/rebind.py` — existing `RunShift`, `_resolution_state()`, and `_shifts_between()`
  contracts that import/rebind consumers still own and that deck-diff v5 must not mutate.
- `src/pptx/shapes/base.py` — the slide-wide uniqueness guarantee for shape IDs, including grouped
  descendants.
- `tests/paper/test_diff.py` — self-diff, detail-level, deterministic report, field, effective,
  bullet, notes, and ordinal-slippage regressions.
- `tests/paper/test_walkthrough_qbr.py` — the release workflow's independent operation-report
  versus `diff_decks()` consistency check.
- `tests/paper/goldens/lineage_v1_v2.diff.json` — current deterministic deck-diff payload.
- `tests/paper/_authoring/build_fixtures.py`, `tests/paper/_authoring/update_goldens.py`,
  `tests/paper/fixtures/README.md`, and `tests/paper/fixtures/MANIFEST.sha256` — frozen-fixture and
  golden authoring discipline.
- `docs/api/diff.rst` and `docs/user/paper-additions.rst` — authoritative public diff and workflow
  documentation.
- `CONTRIBUTING.md` — fixture, deterministic-report, compatibility, and full quality-gate rules.

## Steps

### Step 1 — Establish one conservative paragraph alignment and the version-5 text event model

**Goal:** Every slide-text comparison is partitioned by stable container identity and produces
only exact, structurally justified paragraph associations. The version-5 report can describe what
is known without disguising an ambiguous region as a list of replacements.

**Work:**

Replace the current paragraph-ordinal map in `src/pptx/diff.py` with a small private alignment
boundary. Build an ordered sequence for each stable leaf container from Phase 1 inspection data:
one sequence for each ordinary or grouped text-bearing shape and one row-major sequence for each
table frame. Use the slide-unique shape ID as container identity, retain the container kind, and
carry each block's structured before-side or after-side location with it. Display names remain
diagnostic metadata only.

Align each pair of container sequences by the Phase 1 structural fingerprint. Correspondence must
be exact and order-aware. Where repeated fingerprints allow multiple equally valid exact
alignments, retain only correspondences that are invariant across those alternatives. An ordinal,
candidate iteration order, geometry, edit distance, or text similarity must never settle a tie.
Repeated content may still align when unique exact neighbors or container boundaries establish one
correspondence; otherwise the changed region remains unmatched or is represented as ambiguous.

Classify unmatched regions conservatively. A one-sided region is an insertion or deletion. A
single before block and single after block bounded by the same uniquely established exact context
may be reported as a replacement. Multiple unmatched blocks on both sides are one ambiguous region
unless exact evidence establishes smaller outcomes. An exact fingerprint that occurs once per
container may carry before and after locations as a move/alignment when order changes; do not call
similar-looking but non-identical paragraphs moves. If an ambiguous repeated region is structurally
and observably unchanged on both sides, suppress it without claiming instance-to-instance
identity, preserving the empty self-diff invariant. If content, fields, or full-detail formatting
inside such a region differs and attribution is uncertain, surface the ambiguity rather than
silently dropping it or pairing by position.

Advance `paper-deck-diff` from version 4 to version 5 in this PR. Give text events a stable shape
reference plus explicit `before_location` and `after_location` values, with absent locations for
insertions or deletions. Locations identify the container kind, paragraph position within the
container, and table row/column when applicable; slide-global block indexes and
`container_detail` strings are diagnostic only. Preserve literal before/after text and detailed
field-marker changes, but add an explicit event kind so insertion, deletion, replacement,
move/alignment, and ambiguity cannot be confused. Keep empty facets omitted and establish one
deterministic ordering based on stable container identity and structured before/after location.

Add focused tests alongside the production change for insertion, deletion, a genuinely bounded
single replacement, unique-content reorder, duplicates with unique context, an unresolved
duplicate region, NFC-equivalent text, whitespace, and field-only changes. Assert both the
high-level event kind and the exact before/after locations. Preserve the existing colliding-slide-ID
test: this phase improves attribution within a lineage-matched slide but does not broaden deck
lineage identity.

**Constraints:**

- The alignment implementation remains private to `src/pptx/diff.py`; do not create a generic
  matching module or a new runtime dependency.
- Reuse Phase 1's fingerprint semantics and structured locator data. Do not parse
  `container_detail`, recompute a weaker text-only hash, or create a second public fingerprint
  contract.
- Do not reuse `_longest_common_subsequence()` unchanged. It returns only a set of unique slide
  IDs and intentionally chooses one deterministic tie, while paragraph alignment must retain
  before/after positions and refuse ambiguous duplicate correspondence.
- Exact deterministic tie-breaking may order report entries, but it must never create an object
  association that exact evidence does not establish.
- Preserve whitespace as content. NFC normalization does not trim, collapse, or case-fold text.
- Field marker types and order participate in the alignment fingerprint; volatile displayed field
  values do not. Existing marker offsets remain available to report a field moving within an
  aligned or uniquely bounded paragraph.
- Do not change slide matching, Phase 5 shape matching, chart matching, picture matching, package
  comparison, or the permanent-slide-ID lineage limitation.
- Do not change `pptx.rebind.RunShift` or any import/rebind report schema. Deck diff owns any new
  location-aware record it needs.
- Do not report ambiguity for an unchanged self-diff merely because identical paragraphs have no
  distinguishable instance identity.

**Verification:**

- `uv run pytest -q tests/paper/test_diff.py -k "paragraph or text or field or whitespace or self_diff"`
- `uv run ruff check src/pptx/diff.py tests/paper/test_diff.py`
- `git diff --check`

### Step 2 — Apply the alignment to tables, effective values, and bullets without widening scope

**Goal:** A table-grid insertion is reported structurally and unchanged cell text stays aligned,
while effective-value and bullet changes use the same proven paragraph correspondence as text.
Notes and unsupported table formatting retain their current behavior exactly.

**Work:**

Add an omitted-when-empty table-structure facet to `SlideChange`, populated for stable table-frame
IDs on both sides. At `detail="structure"`, report the Phase 5 structured shape reference and the
before/after row and column counts. When the exact cell sequence and coordinates uniquely establish
a row or column insertion or deletion, include that structural location; when duplicate or empty
cells make the location uncertain, report the exact dimension change plus candidate or ambiguity
context instead of selecting an index. A table frame present on only one side remains governed by
Phase 5 shape addition/removal and must not generate a fictitious matched grid change.

For `detail="text"` and `detail="full"`, align table paragraphs across the entire stable frame in
row-major sequence while retaining each side's row, column, and paragraph-within-cell location.
This allows unchanged values after a row or column insertion to stay associated even though their
coordinates changed. New or removed cells produce only justified insert/delete events; multi-cell
regions containing repeated values remain ambiguous rather than becoming a cascade of positional
replacements. Tables inside groups use the same slide-unique leaf frame ID and must not depend on a
group-name path.

Make text changes, effective-value shifts, and bullet shifts consume the same paragraph-alignment
result instead of independently rebuilding paragraph identity. For aligned ordinary/grouped
paragraphs, compare effective runs only when exact run identity is unique within the paragraph,
preserving the existing policy of skipping repeated or empty run identities rather than guessing.
Use a deck-diff-owned, location-aware effective-shift payload so both paragraph locations are
auditable without changing the rebind/import `RunShift` contract. Resolve bullet values only for
the existing supported `p:sp` paragraph domain and emit bullet shifts only for aligned paragraphs.
Unchanged text that moved may therefore carry a real effective or bullet change at its actual
before/after locations.

Keep table-cell runs blind for effective fonts and paragraph formatting; do not report table
effective or bullet shifts until table-style inheritance exists. Keep speaker notes entirely on
the current `_notes_text()` path: notes still compare as one flat before/after string and emit only
`notes_change`. Preserve the existing detail gates: table dimensions at `structure`, text and
fields at `text`, and effective/bullet values only at `full`.

Add focused coverage for row insertion, column insertion, deletion, multiple paragraphs per cell,
duplicate and blank cells, and grouped tables. Assert that dimension changes appear at structure
detail, exact text insertions appear at text detail, and no unchanged cells become replacements.
At full detail, cover an ordinary paragraph insertion before an unchanged effective or bullet
change and verify that the change follows the aligned paragraph. Also prove that table formatting
stays absent and that an ordinary notes edit produces exactly the existing flat `notes_change`.

**Constraints:**

- Table structure is identity and grid evidence, not visual inference. Do not use cell geometry,
  estimated widths, merge appearance, neighboring pixel distance, or fuzzy cell text.
- Preserve before and after table coordinates separately. Do not collapse them into one ordinal or
  parse the legacy display label to recover coordinates.
- A row/column-count delta is certain even when its insertion location is not; represent those two
  confidence levels separately rather than withholding the structural change or inventing an
  exact index.
- The sequence domain is one stable table frame, never all table blocks on a slide and never cells
  from two different frames with similar names or content.
- Do not add table-style inheritance, table effective-font resolution, table bullet resolution,
  merge surgery, or changes to the existing table editing APIs.
- Do not refactor paragraph-level notes inspection or alignment. `notes_change` remains separate
  and flat.
- Do not change rebind's runtime matching or report behavior. Reuse the deck-diff paragraph
  alignment inside `diff.py`; do not move it into `rebind.py` merely because the old diff imported
  `_resolution_state()`.
- Preserve Phase 5's top-level shape-boundary contract. This phase may use grouped leaf IDs for
  text containers, but it does not recursively redefine shape additions/removals.

**Verification:**

- `uv run pytest -q tests/paper/test_diff.py -k "table or effective or bullet or notes or group"`
- `uv run pytest -q tests/paper/test_walkthrough_qbr.py -k "self_consistency_against_deck_diff"`
- `uv run ruff check src/pptx/diff.py tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py`
- `git diff --check`

### Step 3 — Freeze the regression, migrate consumers and documentation, and close PR 6 cleanly

**Goal:** LS-05 is preserved as reproducible corpus evidence, every in-repo deck-diff consumer and
golden understands schema version 5, and `gavin/text-diff-exact-alignment` is independently green
before the separate Phase 7 repository-wide hygiene sweep begins.

**Work:**

Add a focused self-generated lineage pair under `tests/paper/fixtures/self_generated/` that
captures the two failure shapes without depending on the code being fixed: an ordinary text body
with a paragraph insertion and a table whose later version has a row or column added before
unchanged content. Author the two sides independently in
`tests/paper/_authoring/build_fixtures.py` rather than generating the after deck by calling the
table-insertion operation under test. Include repeated or blank content where it is needed to pin
the conservative ambiguity outcome. Add honest sidecars, document the pair in the fixture taxonomy,
and regenerate `MANIFEST.sha256` exactly through the command documented in the fixture README.

Use that pair in `tests/paper/test_diff.py` to preserve the eval-observed regression: one structural
grid change and the real inserted text are reported, while every unchanged paragraph or cell is
free of fictional replacement events. Keep smaller in-memory cases for branch coverage that does
not merit extra binary fixtures. Re-run the full non-corrupt self-diff corpus, determinism checks,
field/effective/bullet regressions, and the QBR walkthrough consumer. Update the lineage deck-diff
golden through `tests/paper/_authoring/update_goldens.py` only, inspect its diff, and ensure the only
schema migration is version 4 to version 5 plus the intentional text/table payload changes.

Update `src/pptx/diff.py` public docstrings, `docs/api/diff.rst`, and the deck-diff section of
`docs/user/paper-additions.rst`. Document stable container partitioning, conservative exact
alignment, structured before/after locations, ambiguity semantics, table-structure detail gates,
and the version-5 migration. State explicitly that notes remain flat, table effective formatting
is unsupported, independently authored decks remain out of scope, and exact alignment does not
turn shape IDs into durable identity after deletion and reuse.

Review the Git comparison for `gavin/text-diff-exact-alignment` against
`gavin/shape-diff-stable-identity`. The comparison must contain LS-05 production code, its focused
fixtures/tests/golden, and its own diff documentation only. Record the branch base and the adjacent
PR links in the PR description's Stack section. Phase 7 will add the final cross-repository stale
language sweep and compatibility confirmation to this same branch; do not absorb that sweep here.

**Constraints:**

- Frozen fixture bytes, sidecars, and manifest updates travel together. Tests never regenerate a
  fixture at runtime.
- Goldens update only through `tests/paper/_authoring/update_goldens.py`; review and reject churn in
  inspection or import goldens unrelated to this phase.
- Keep the regression fixture feature-isolated. Do not copy a full evaluation task deck or include
  task-specific names, branding, or content.
- Update consumers to the version-5 contract rather than maintaining a hidden ordinal compatibility
  path inside `diff_decks()`.
- Documentation in this phase covers the public behavior changed by LS-05. The repository-wide
  stale-term sweep, planning-status reconciliation, stack-wide comparison audit, and final release
  gate belong to Phase 7.
- Do not mix LS-01 through LS-04/LS-06 fixes, unrelated fixture regeneration, broad formatting
  cleanup, or a seventh PR into this phase.

**Verification:**

- `uv run python tests/paper/_authoring/update_goldens.py`
- `uv run pytest -q tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py`
- `uv run pytest -q tests/paper/test_fixture_corpus.py -k "diff_alignment"`
- `uv run ruff check src/pptx/diff.py tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py tests/paper/_authoring/build_fixtures.py tests/paper/_authoring/update_goldens.py`
- `uv run sphinx-build -W -b html docs docs/.build/html`
- `git diff --check`
- `uv run pytest`
- `uv run behave`
- `uv run pytest -m lo_smoke tests/paper`
- `uv run make docs`
- `uv run make build`

## Files

| Action | Path |
|---|---|
| Edit | `src/pptx/diff.py` — add exact container-scoped paragraph alignment, version-5 text/table/effective/bullet payloads, and preserve flat notes behavior |
| Edit | `tests/paper/test_diff.py` — add paragraph, table, ambiguity, location, detail-gate, downstream-facet, schema, determinism, and regression assertions |
| Edit | `tests/paper/test_walkthrough_qbr.py` — consume and assert the structured version-5 text-change contract in the release workflow |
| Edit | `tests/paper/_authoring/build_fixtures.py` — independently author the focused before/after paragraph-and-table alignment pair |
| Create | `tests/paper/fixtures/self_generated/diff_alignment_v1.pptx` and `tests/paper/fixtures/self_generated/diff_alignment_v2.pptx` — frozen LS-05 lineage inputs |
| Create | `tests/paper/fixtures/self_generated/diff_alignment_v1.json` and `tests/paper/fixtures/self_generated/diff_alignment_v2.json` — provenance and exact ground truth for the pair |
| Edit | `tests/paper/fixtures/README.md` — add the pair to the feature-isolated taxonomy and explain its diff purpose |
| Edit | `tests/paper/fixtures/MANIFEST.sha256` — pin the new fixture and sidecar hashes |
| Edit | `tests/paper/goldens/lineage_v1_v2.diff.json` — record deterministic `paper-deck-diff` version 5 output |
| Edit | `docs/api/diff.rst` — document exact matching, event/location fields, table structure, detail gates, and schema migration |
| Edit | `docs/user/paper-additions.rst` — update the public verification workflow and limitations |

## What this phase does NOT include

- Structural anchor creation or editing behavior; Phase 1 owns `BlockAnchor`, inspection v3, and
  replace-result v2.
- Import layout, placeholder, or section selection; phases 2 through 4 own those changes.
- Shape matching or structured shape-facet migration; Phase 5 owns deck-diff v4 and is this phase's
  base.
- Fuzzy text similarity, edit-distance pairing, geometric proximity, nearest-neighbor scoring,
  render comparison, OCR, or matching independently authored decks.
- Durable identity after a shape ID is deleted and reused by a new same-kind shape.
- Table-style effective formatting, table bullets, merge inference, or changes to table edit APIs.
- Paragraph-level notes alignment; notes remain one flat `notes_change` payload.
- Changes to the public rebind/import `RunShift` contract or their report schema versions.
- A public alignment API, a shared matching framework, a new source module, or a runtime
  dependency.
- A cross-repository documentation sweep, planning-status update, stack-wide release audit, or a
  seventh GitHub PR; Phase 7 performs that final hygiene work in this same top branch.

## Tests this phase must include

- Inserting one paragraph at the beginning, middle, and end of a shape produces one insertion each
  and leaves every exact unchanged neighbor associated with its correct before/after location.
- Deleting one paragraph produces one deletion without cascading replacements.
- A single changed paragraph uniquely bounded by exact neighbors is one replacement with explicit
  before and after locations.
- Reordering paragraphs whose structural fingerprints are unique produces exact move/alignment
  evidence rather than replacements.
- Repeated identical paragraphs align only where exact surrounding sequence establishes a unique
  correspondence; unresolved duplicate regions expose ambiguity and never use ordinal order to
  pair instances.
- An unchanged shape containing repeated identical or empty paragraphs still self-diffs as empty,
  without pretending those individual instances have durable identity.
- NFC-equivalent literal text aligns, while leading, trailing, and repeated whitespace remains
  content and continues to produce real changes.
- Field type/order participates in alignment, and an aligned or uniquely bounded field-position or
  field-type change preserves detailed marker evidence without using rendered field values.
- Ordinary and nested-group text use the leaf shape ID as container identity; group names and group
  order never participate.
- Row and column count changes appear at `detail="structure"` with stable table-frame identity and
  exact before/after dimensions.
- A table row insertion and a table column insertion retain before/after cell coordinates, report
  the real inserted content, and produce no shifted-cell replacement cascade.
- Duplicate or blank table cells make an insertion location ambiguous when exact evidence cannot
  settle it; the dimension change remains present and no index is invented.
- Multiple paragraphs per table cell align within the stable frame's row-major sequence without
  colliding with paragraphs in another cell or another table.
- A grouped table uses its slide-unique leaf frame ID and has the same grid/text behavior as a
  top-level table.
- Full-detail effective shifts and bullet shifts follow the same aligned ordinary/grouped
  paragraph after an earlier insertion; repeated run or paragraph identity is skipped or marked
  ambiguous rather than guessed.
- Table-cell effective and bullet shifts remain absent, preserving the blind table-style boundary.
- Notes-only edits still emit exactly one flat `notes_change`; notes do not enter paragraph or
  table event lists.
- `structure`, `text`, and `full` preserve their gates: table grid at structure, paragraph/field
  events at text, and effective/bullet shifts only at full.
- A stable table frame removed or added as a shape does not also produce a matched grid change.
- `SlideChange.to_dict()` omits empty facets, report ordering is deterministic across runs, and the
  reviewed golden is byte-identical at schema version 5.
- Every non-corrupt corpus fixture has an empty self-diff, package-level changes retain their
  existing behavior, and the QBR walkthrough still agrees with the operation reports.
- The new fixture pair opens, survives save/reopen, passes relationship integrity and manifest
  checks, and loads under the existing LibreOffice smoke gate.

## Done when

1. No deck-diff paragraph consumer pairs blocks by slide-global or container-local ordinal.
2. One private exact alignment result in `src/pptx/diff.py` drives text, field, effective-value,
   and bullet comparisons for their shared supported paragraph domain.
3. Insertions, deletions, exact moves, and uniquely bounded replacements carry stable shape
   identity plus separate before/after locations; unresolved repeated regions are unmatched or
   explicit ambiguity, never forced pairs.
4. `paper-deck-diff` is schema version 5, empty facets remain omitted, serialization is
   deterministic, and every in-repo consumer and reviewed golden uses the new contract.
5. Table row/column-count changes appear at structure detail, exact cell text remains aligned
   across grid growth, and ambiguous duplicate cells never produce fictional replacement
   cascades or invented insertion indexes.
6. Notes remain one flat `notes_change`, and table effective-font/bullet reporting remains outside
   scope and absent.
7. The focused frozen fixture pair, sidecars, README entry, and manifest hashes land together and
   reproduce both the ordinary paragraph and table-grid failures independently of the diff code.
8. PR 6 contains only LS-05 production, regression, fixture/golden, and focused documentation
   changes, is based on `gavin/shape-diff-stable-identity`, and is independently reviewable as
   `gavin/text-diff-exact-alignment` before Phase 7 adds final hygiene to the same branch.
9. Focused checks pass:
   `uv run pytest -q tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py`,
   `uv run pytest -q tests/paper/test_fixture_corpus.py -k "diff_alignment"`,
   `uv run ruff check src/pptx/diff.py tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py tests/paper/_authoring/build_fixtures.py tests/paper/_authoring/update_goldens.py`,
   `uv run sphinx-build -W -b html docs docs/.build/html`, and `git diff --check`.
10. The complete PR quality gate passes:
    `uv run pytest`, `uv run behave`, `uv run pytest -m lo_smoke tests/paper`,
    `uv run make docs`, and `uv run make build`.
