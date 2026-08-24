# Phase 5 — Stable Shape Diff Identity

**Created:** 2026-08-22
**Status:** Complete

## Motivation

`diff_decks()` currently identifies a top-level shape by a unique display name and otherwise by
its element kind plus collection ordinal. That fallback turns a harmless z-order change among
duplicate-named or unnamed shapes into confident but fictional additions, removals, geometry
changes, image replacements, or chart changes. This phase makes the lineage diff use the same
stable shape identity that the deck already carries and gives every shape-oriented report facet a
consistent, auditable reference.

## Context

**What exists today:**

- `src/pptx/diff.py` pairs slides by permanent slide ID, so the report is explicitly for
  lineage-related decks rather than independently authored or visually similar presentations.
- `_shape_keys()` uses a unique shape name when available. Duplicate or empty names fall back to
  the shape element's local name plus its ordinal among shapes of that element kind.
- `_diff_slide()` compares only top-level shapes. It derives additions, removals, geometry,
  picture replacement, and chart-data changes from the synthetic keys produced by
  `_shape_keys()`.
- Shape additions, removals, and image replacements serialize as strings. Geometry entries use a
  string-valued `shape` field, and chart entries use a string-valued `chart` field. Those labels
  can be a display name or a manufactured `kind#ordinal`, so consumers cannot recover stable
  identity.
- `src/pptx/shapes/base.py` defines `shape_id` as a positive integer unique among every shape on a
  slide, including grouped descendants. The diff's existing top-level boundary therefore needs no
  group-name path or z-order lineage.
- Group membership is intentionally structural at this layer: moving a leaf into or out of a
  group currently appears as a top-level removal/addition rather than a recursively matched move.
- `paper-deck-diff` is currently schema version 3. `SlideChange.to_dict()` omits empty facets, and
  the reviewed `lineage_v1_v2.diff.json` golden pins deterministic field order and payload shape.
- `tests/paper/test_diff.py` covers the lineage edit list, frozen golden, detail gates,
  deterministic output, package-only z-order changes, image replacement, chart data, geometry,
  and self-diff across the fixture corpus. `tests/paper/test_walkthrough_qbr.py` consumes the
  string-valued shape fields in its end-to-end self-consistency assertions.

**What this phase delivers:**

- Top-level shape pairing by slide-wide shape ID for permanent-slide-ID pairs.
- Conservative handling of incompatible shape-kind reuse as a removal plus an addition, and an
  explicit typed refusal for a malformed slide that repeats a shape ID within one side rather than
  silently overwriting a candidate.
- One structured shape-reference value, carrying stable shape ID and current display name, across
  shape additions, removals, geometry changes, image replacements, and chart changes.
- `paper-deck-diff` schema version 4 with deterministic serialization and no synthetic
  `kind#ordinal` identities.
- Regressions for the eval-observed duplicate-name reorder failure and for rename, z-order,
  add/remove, kind-reuse, grouping, schema, and consumer behavior.
- Public documentation of shape-ID matching, the top-level group boundary, and the remaining
  same-kind deleted-ID reuse limitation.

**Reference files to study before starting:**

- `docs/plans/selection-integrity-hardening/spec.md` — authoritative LS-04 behavior, report, and
  GitHub stack contracts.
- `docs/plans/selection-integrity-hardening/build-sequence.md` — phase dependencies, schema
  ownership, cross-cutting rules, and PR boundaries.
- `docs/plans/lossy-selection-audit.md` — the eval-triggered duplicate-name/z-order reproduction,
  impact, provenance, and rejected approximate matchers.
- `src/pptx/diff.py` — current schema values, `SlideChange`, name/ordinal shape keys, all
  shape-oriented facets, deterministic serialization, and detail-level gates.
- `src/pptx/shapes/base.py` — the slide-wide uniqueness contract for `shape_id`.
- `src/pptx/shapes/graphfrm.py`, `src/pptx/shapes/group.py`, and
  `src/pptx/shapes/picture.py` — shape-kind and grouped/top-level behavior that the matcher must
  classify without probing visual similarity.
- `src/pptx/errors.py` — existing typed refusal for unsupported or internally ambiguous document
  structure.
- `tests/paper/test_diff.py` — focused lineage, deterministic-golden, shape-facet, detail-level,
  package-only, stream, and corpus regressions.
- `tests/paper/test_walkthrough_qbr.py` — end-to-end consumers of removal and image-replacement
  report values.
- `tests/paper/goldens/lineage_v1_v2.diff.json` — reviewed deck-diff schema golden.
- `tests/paper/_authoring/update_goldens.py` — the only approved command for regenerating the
  deck-diff golden.
- `tests/paper/fixtures/self_generated/gauntlet.pptx` and
  `tests/paper/fixtures/self_generated/lineage_v1.pptx` — frozen baselines suitable for focused
  duplicate-name, reorder, and shape-facet regressions.
- `tests/paper/fixtures/README.md` and `tests/paper/test_fixture_corpus.py` — fixture provenance and
  manifest requirements if the existing corpus cannot express one required case.
- `docs/api/diff.rst`, `docs/user/paper-additions.rst`, and `README.md` — public deck-diff contract
  and lineage-limit documentation.
- `CONTRIBUTING.md` — fixture, golden, compatibility, and complete quality-gate requirements.

## Steps

### Step 1 — Replace positional shape pairing with stable lineage identity

**Goal:** Within each permanent-slide-ID pair, every top-level shape is matched only by its
slide-wide shape ID and compatible structural kind; collection order and display name no longer
participate in identity.

**Work:**

Replace the name/element-kind-ordinal key map in `src/pptx/diff.py` with a top-level shape inventory
keyed by `shape_id`. Retain shape name only as display metadata captured separately for each side.
Validate the per-slide uniqueness assumption while building each inventory. If malformed input
contains the same shape ID more than once on one side, refuse through the existing typed
unsupported-structure boundary with deterministic candidate diagnostics rather than allowing a
dictionary overwrite or inventing a secondary match.

For IDs present on both sides, compare the same structural shape-kind classification already used
by the current matcher. Compatible kinds proceed to the existing facet comparisons. Incompatible
kinds do not become one changed object: place the before shape in removals and the after shape in
additions, and skip geometry, image, and chart comparisons for that ID. An ID found on only one
side remains a straightforward addition or removal.

Keep matching at the top-level shape tree. A child moved into a group disappears from the
top-level inventory and is reported as removed; a child moved out appears as added. Do not descend
into groups to rescue that association. Confirm that merely changing z-order or a shape's display
name does not change the matched pair or generate any shape facet when its substantive values are
unchanged.

**Constraints:**

- Shape identity is `shape_id`, scoped to one already-matched slide. Do not include slide position,
  z-order, name, geometry, element ordinal, or group traversal path in the key.
- Do not add fuzzy names, bounding boxes, pixel distance, overlap, nearest-neighbor scoring, image
  similarity, or any visual fallback.
- Preserve the lineage-only contract. This phase does not attempt to associate shapes in
  independently authored decks.
- Treat incompatible structural kinds as two honest objects, even when their ID and name are the
  same. Do not compare their facets as though one had transformed into the other.
- Do not recursively match grouped descendants. Preserve the documented top-level add/remove
  semantics and the existing detail-level gates.
- Detect duplicate IDs before materializing the map. A malformed side must not silently keep the
  first or last candidate.
- Keep the matching implementation private to `src/pptx/diff.py`; do not create a package-wide
  identity framework or change upstream shape collection behavior.
- Do not change text, effective-value, bullet, or notes alignment in this step. Phase 6 owns those
  consumers.

**Verification:**

- `uv run pytest -q tests/paper/test_diff.py -k "shape or duplicate or z_order or geometry or image or chart or group"`
- `uv run ruff check src/pptx/diff.py tests/paper/test_diff.py`
- `git diff --check`

### Step 2 — Carry one structured shape reference through every shape facet

**Goal:** Every shape-oriented entry in the public deck-diff report exposes stable shape ID and
human-readable name in one consistent structure, while empty-facet omission and detail gating
remain unchanged.

**Work:**

Add one immutable report value in `src/pptx/diff.py` for a shape reference, following the existing
`SlideRef` pattern rather than introducing unrelated schema machinery. Its serialized form carries
the slide-scoped stable shape ID and the display name from the side represented by the entry.
Names may be duplicate, empty, or different across sides; they are labels only and are never
promoted back into matching identity.

Use that same reference shape in all five affected facets. Additions carry the after-side
reference, removals carry the before-side reference, and matched geometry, image, and chart changes
carry the stable ID with the after-side display name. A rename by itself remains visible through
the package-level diff but does not manufacture a specialized shape change. Ensure
`SlideChange.to_dict()` recursively converts these values to JSON-ready dictionaries and retains
its current deterministic facet order and empty-facet omission. Leave slide references, text
entries, notes, effective shifts, bullet shifts, and package changes unchanged in this schema
layer.

Advance `paper-deck-diff` from version 3 to version 4 in the same change. Update in-repository
consumers to inspect the structured reference rather than compare the former string label. Refresh
the reviewed lineage diff golden only through `tests/paper/_authoring/update_goldens.py`, and review
that the deck-diff changes are limited to version 4 and the intended five shape facets. The helper
also rewrites unrelated goldens, so restore or exclude outputs whose bytes did not need to change
rather than committing incidental churn.

**Constraints:**

- Use a single structured representation for additions, removals, geometry, images, and charts;
  do not define one near-duplicate reference format per facet.
- Stable ID and display name are required. Do not serialize synthetic `kind#ordinal` labels.
- Before-only and after-only entries use names from their actual side. Matched entries must not
  imply that a display-name change changed the object's identity.
- Preserve chart delta contents, geometry facet/value contents, image hash semantics, and current
  structure/text/full gates; only their shape reference changes here.
- Do not add a shape-move facet for z-order. The package-level semantic diff may remain non-empty
  when XML order changes, but specialized shape additions, removals, geometry, images, and charts
  must remain empty for a pure z-order edit.
- Do not preserve the old string-valued shape payload under a second compatibility field. The
  report's version increment is the migration boundary.
- Serialization must be deterministic across runs and must remain directly JSON-serializable.
- Do not advance beyond deck-diff version 4. Phase 6 owns version 5.

**Verification:**

- `uv run python tests/paper/_authoring/update_goldens.py`
- `uv run pytest -q tests/paper/test_diff.py -k "golden or deterministic or shape or geometry or image or chart"`
- `uv run pytest -q tests/paper/test_walkthrough_qbr.py -k "self_consistency"`
- `uv run ruff check src/pptx/diff.py tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py`
- `git diff --check`

### Step 3 — Lock the regression cases and document schema v4

**Goal:** Focused contract tests prove that stable IDs eliminate the observed false matches without
hiding real lineage changes, and public documentation gives callers enough information to migrate
to the version-4 report.

**Work:**

Extend `tests/paper/test_diff.py` with the eval-observed case: begin from a frozen deck, give two
same-kind shapes the same display name, reverse only their z-order on the after side, and assert
that the specialized shape facets report no additions, removals, geometry changes, image
replacements, or chart changes. Assert separately that the package-level diff still records the
real slide XML ordering change.

Add focused cases showing that duplicate or empty names do not affect matching; a rename alone
does not become removal/addition; a real geometry, image, or chart edit remains attributed to the
same stable ID; and an actual shape addition or deletion exposes the correct after-side or
before-side reference. Cover an ID reused with an incompatible structural kind and require exactly
one removal plus one addition with no cross-kind facet comparison. Cover a malformed duplicate ID
within one side and assert a typed, deterministic refusal rather than candidate loss. Preserve one
group-boundary case proving that moving a leaf into or out of a group retains the declared
top-level add/remove interpretation.

Pin version 4, structured values, field order, empty-facet omission, and deterministic ordering in
the frozen golden and direct report assertions. Update `tests/paper/test_walkthrough_qbr.py` so its
self-consistency checks assert both ID and name for the deleted picture and replaced image. Retain
the full corpus self-diff test to detect regressions on producer-authored packages.

Update `docs/api/diff.rst`, `docs/user/paper-additions.rst`, the `diff_decks()` docstring, and the
relevant README deck-diff summary. Document that top-level shapes in lineage pairs match by
slide-wide shape ID; duplicate, empty, or renamed display names are harmless; z-order alone does
not generate specialized shape changes; group-boundary moves remain remove/add; and schema v4 uses
structured references. State the remaining hazard plainly: deleting a shape and later reusing its
ID for a new same-kind shape is not distinguishable without a persistent identifier absent from
general PPTX files. Include migration guidance for consumers that previously compared string
labels in shape facets.

**Constraints:**

- The regression must exercise the public `diff_decks()` entry point, not only a private inventory
  helper.
- Use existing frozen fixtures plus narrow, deterministic in-memory edits when they can represent
  the case. If a new binary fixture is unavoidable, add its truthful sidecar and manifest entry and
  keep the fixture limited to the LS-04 reproduction.
- Keep the expected report exact enough to catch wrong attribution: assert stable IDs, correct
  side-specific names, facet counts, and absence of fictional facets.
- Do not weaken package-level z-order reporting to make the specialized shape facet empty.
- Do not claim that shape ID solves deleted-and-reused same-kind identity or cross-deck visual
  matching.
- Do not change paragraph/table matching, notes comparison, effective-value alignment, or bullet
  alignment; phase 6 owns those changes.
- Do not regenerate unrelated fixtures or goldens, and do not fold documentation hygiene from
  phase 7 into this PR beyond the authoritative LS-04/API material required for independent review.

**Verification:**

- `uv run pytest -q tests/paper/test_diff.py`
- `uv run pytest -q tests/paper/test_walkthrough_qbr.py -k "self_consistency"`
- `uv run ruff check src/pptx/diff.py tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py`
- `uv run sphinx-build -W -b html docs docs/.build/html`
- `git diff --check`

## Files

| Action | Path |
|---|---|
| Edit | `src/pptx/diff.py` — replace name/ordinal pairing, add structured shape references, update all shape facets, and advance deck-diff to v4 |
| Edit | `tests/paper/test_diff.py` — add duplicate-name/z-order, rename, add/remove, incompatible-kind, malformed-ID, group-boundary, schema, and determinism regressions |
| Edit | `tests/paper/test_walkthrough_qbr.py` — migrate shape-facet consumers to stable structured references |
| Edit | `tests/paper/goldens/lineage_v1_v2.diff.json` — reviewed schema-v4 deck-diff payload generated by the approved script |
| Reference; generated use only | `tests/paper/_authoring/update_goldens.py` — regenerate the deck-diff golden without changing the generator unless schema generation itself proves incomplete |
| Edit | `docs/api/diff.rst` — document shape-ID matching, structured references, group scope, and lineage hazards |
| Edit | `docs/user/paper-additions.rst` — update verification workflow and report migration guidance |
| Edit | `README.md` — correct the public deck-diff identity and limitation summary |
| Conditional edit | `tests/paper/fixtures/README.md` and `tests/paper/fixtures/MANIFEST.sha256` — only if a new focused frozen fixture is strictly necessary |
| Conditional create | `tests/paper/fixtures/self_generated/<focused-shape-diff-fixture>.pptx` and its sidecar — only if existing frozen fixtures plus deterministic edits cannot express the regression |

## What this phase does NOT include

- Structural text anchors or inspection/replace-result schema changes from phase 1.
- Import layout, placeholder, or section matching from phases 2 through 4.
- Paragraph or table-cell sequence alignment, text-location restructuring, table grid reporting,
  effective-value alignment, or bullet alignment; phase 6 owns deck-diff version 5.
- Recursive matching of shapes inside groups or a new report interpretation for moving a shape
  across the top-level/group boundary.
- Z-order move reporting as a specialized shape facet; the package-level XML change remains the
  evidence for that operation.
- Matching independently authored decks, geometric association, visual similarity, rendering,
  OCR, fuzzy name comparison, or nearest-neighbor selection.
- Detection of deleted-and-reused same-kind shape IDs beyond documenting the lineage limitation.
- Changes to upstream shape IDs, shape collections, mutation APIs, or any runtime dependency.
- Broad report refactors, unrelated fixture regeneration, or phase-6 schema design.

## Tests this phase must include

- Two same-kind, duplicate-named shapes whose z-order is reversed remain paired by their original
  IDs and produce no specialized additions, removals, geometry, image, or chart changes.
- A z-order-only edit remains observable through `package_changes`; the fix does not erase the
  actual XML-order delta.
- Duplicate names, empty names, and a display-name change do not alter shape identity or create a
  removal/addition pair.
- A real geometry edit after duplicate-name or z-order changes is attributed to the correct stable
  ID and carries its display name.
- A real image replacement and a real chart-data change retain the correct stable ID through their
  structured references.
- A shape present only before serializes as one removal using the before-side ID/name; a shape
  present only after serializes as one addition using the after-side ID/name.
- An identical ID with incompatible structural kinds serializes as one removal plus one addition
  and never produces matched geometry, image, or chart facets.
- A malformed slide containing duplicate shape IDs on one side raises the chosen existing typed
  unsupported-structure refusal with deterministic identity diagnostics rather than overwriting a
  candidate or pairing by ordinal.
- Moving a leaf into or out of a group preserves the documented top-level removal/addition
  behavior and does not trigger recursive or geometric matching.
- Shape additions, removals, geometry changes, image replacements, and chart changes all serialize
  the same stable ID/name reference shape; synthetic `kind#ordinal` labels are absent.
- `paper-deck-diff` reports version 4, remains deterministic, omits empty facets, and matches the
  reviewed `lineage_v1_v2.diff.json` golden.
- `structure` still gates additions, removals, geometry, and image changes; chart data still starts
  at `text`; all pre-existing text, notes, effective, bullet, slide, and package facets retain their
  current behavior.
- The QBR walkthrough's self-consistency assertions consume structured references and still
  identify the intended deleted and image-replaced shapes after save/reopen.
- Self-diff remains empty for every non-corrupt fixture, including producer-authored decks.

## Done when

1. Every top-level shape in a permanent-slide-ID pair is considered for matching only by its
   slide-wide shape ID and compatible structural kind.
2. Reordering duplicate-named or unnamed shapes cannot fabricate additions, removals, geometry,
   image, or chart changes, while the real package-level z-order delta remains visible.
3. Incompatible-kind ID reuse is represented as removal plus addition; malformed duplicate IDs
   refuse explicitly; neither case falls back to name, ordinal, geometry, or similarity.
4. Every affected shape facet exposes one structured stable ID/display-name reference, and all
   in-repository consumers have migrated from string comparison.
5. `paper-deck-diff` is version 4, its serialization is deterministic, empty facets remain omitted,
   and the reviewed lineage golden changes only as required by this schema layer.
6. Public docs describe shape-ID matching, duplicate-name safety, top-level group semantics,
   schema-v4 migration, and the unresolved same-kind ID-reuse limitation without claiming visual
   or independently authored deck matching.
7. PR 5 contains only LS-04 production, regression, golden, and documentation changes; it is based
   on `gavin/section-selection-identity` and independently reviewable as
   `gavin/shape-diff-stable-identity`.
8. Focused checks pass:
   `uv run pytest -q tests/paper/test_diff.py`,
   `uv run pytest -q tests/paper/test_walkthrough_qbr.py -k "self_consistency"`,
   `uv run ruff check src/pptx/diff.py tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py`,
   `uv run sphinx-build -W -b html docs docs/.build/html`, and `git diff --check`.
9. The complete PR quality gate passes:
   `uv run pytest`, `uv run behave`, `uv run pytest -m lo_smoke tests/paper`,
   `uv run make docs`, and `uv run make build`.
