---
title: Selection integrity hardening
author: Paper Instruments
created: 2026-08-22
status: executed
---

# Selection integrity hardening

## Motivation

`paper-pptx` promises that an uncertain edit refuses rather than silently changing the wrong
object. Six Paper-added selectors violate that contract by treating traversal order as identity or
by choosing the first candidate in an ambiguous set. The same weakness also makes `diff_decks()`
produce misleading shape and text changes, undermining the verification layer used to evaluate all
other package behavior.

## Goals

- An inspected text anchor can never edit a different shape, group descendant, table cell, notes
  shape, or paragraph merely because that object occupies the anchor's former ordinal and has the
  same text.
- Automatic slide import and placeholder reconciliation proceed only through explicit or unique
  matches; ambiguous inputs refuse atomically with actionable candidate details.
- `diff_decks()` uses lineage identity for shapes and exact, conservative sequence alignment for
  paragraphs and table cells, without manufacturing changes after reorder or insertion.
- Section-name selection requires uniqueness, with stable section ID available when names collide.
- Existing unambiguous workflows retain their behavior, and compatibility paths for legacy anchors
  become conservative rather than unsafe.
- No selector introduced by this work uses fuzzy text, geometry, bounding boxes, pixel distance, or
  nearest-neighbor scoring.
- The work ships as a bottom-up GitHub PR stack whose layers are independently testable,
  documented, and small enough to review without understanding an unrelated finding.

## Approach

Extend inspection in `src/pptx/inspect.py` to emit a versioned structural text locator alongside a
full content fingerprint, and make anchored editing in `src/pptx/edit.py` resolve that identity
before validating content. Change the candidate tiers in `src/pptx/compose.py` and
`src/pptx/rebind.py` to accept only a single candidate, reusing the existing typed-refusal and
explicit-override conventions. Replace ordinal association in `src/pptx/diff.py` with stable shape
IDs and one shared exact sequence-alignment boundary for text-bearing containers, retaining table
coordinates as structured context rather than flattening them into a global ordinal. Strengthen
section selection at the compose boundary and expose the section GUID as the explicit identity
path. Deliver these changes as six dependent GitHub PRs, one cohesive finding per layer, with
schema migrations occurring in the layer that changes each public payload.

**Rejected alternatives:**

- Lengthen the anchor hash without adding structural identity: identical text in different objects
  would still collide by design.
- Resolve ambiguous layouts, placeholders, shapes, or paragraphs by visual proximity or fuzzy
  similarity: presentation geometry and text resemblance are not stable identity.
- Preserve first-match behavior and only add warnings to reports: the deck would already be
  mutated before the caller could act.
- Pair duplicate diff objects by ordinal and label the result “best effort”: a precise-looking false
  association is less useful than an explicit unmatched or ambiguous result.
- Build six independent matching frameworks: the package already has reusable exact-candidate,
  typed-refusal, and sequence-alignment patterns.

## Scope

### In scope

- LS-01: structurally addressed, content-validated `BlockAnchor` records and safe legacy-anchor
  handling.
- LS-02: unique-only automatic layout selection for `import_slide()` and `append_deck()`.
- LS-03: unique-only placeholder fallback and an explicit placeholder mapping path for
  adopt-theme import.
- LS-04: shape-ID matching for lineage-related deck diffs.
- LS-05: exact paragraph and table-cell alignment, including honest reporting of table grid
  insertions without cascades of fictional replacements.
- LS-06: unique section-name lookup and stable section-ID selection.
- Typed, atomic refusals; versioned report payloads; public RST documentation; regression fixtures
  and goldens.

### Out of scope

- LS-07 footer-furniture canonicalization; it requires a separate product decision about whether
  duplicate removal is intentional default behavior.
- LS-08 changes to upstream `SlideLayouts.get_by_name()`; changing inherited first-match semantics
  requires a separate compatibility decision. Paper-owned composition code will not use it as an
  ambiguity bypass.
- Cross-deck visual matching, rendering, OCR, or comparison of independently authored decks;
  `diff_decks()` remains lineage-only.
- Geometric or fuzzy fallback matching of any kind.
- Durable identity for a deleted shape ID that is later reused by a new same-kind shape. The known
  lineage hazard must be documented; detection requires a persistent identifier not present in all
  PPTX files.
- Preserving a text anchor across arbitrary structural changes when OOXML provides no durable
  paragraph or cell identity. The safe outcome in that case is a typed refusal and re-inspection.
- Changing `append_deck()` into a per-slide mapping API. If an ambiguous source slide needs an
  explicit layout or placeholder map, callers use ordered `import_slide()` calls.
- New runtime dependencies.

## Delivery as a GitHub PR stack

All implementation changes must be opened as a dependent GitHub PR stack, reviewed and merged from
the bottom upward. Per the approved implementation base, the first PR targets PR #38's
`codex/pptx-doc-hygiene` branch at exact remote head `a3c1ae6e`; each later PR targets the branch
directly below it rather than `main`.

| Order | Branch | GitHub base | Scope | Public schema after this PR |
|---:|---|---|---|---|
| 1 | `gavin/block-anchor-structural-identity` | PR #38 (`codex/pptx-doc-hygiene` at `a3c1ae6e`) | LS-01 structural anchors, conservative legacy resolution, inspection/edit docs and regressions | `paper-text-inspection` v3; `paper-replace-result` v2 |
| 2 | `gavin/import-layout-unique-selection` | PR 1 branch | LS-02 unique layout tiers and removal of first-layout fallback | unchanged |
| 3 | `gavin/placeholder-rebind-unique-selection` | PR 2 branch | LS-03 unique placeholder fallback, import `placeholder_map`, reports and regressions | `paper-import-report` v2 |
| 4 | `gavin/section-selection-identity` | PR 3 branch | LS-06 unique section names, `section_id`, reports and regressions | `paper-import-report` v3 |
| 5 | `gavin/shape-diff-stable-identity` | PR 4 branch | LS-04 stable shape-ID diffing and structured shape references | `paper-deck-diff` v4 |
| 6 | `gavin/text-diff-exact-alignment` | PR 5 branch | LS-05 paragraph/table alignment and structural table-change reporting | `paper-deck-diff` v5 |

### Stack rules

1. Every PR contains its own production change, focused regression fixture or golden, authoritative
   documentation, and any schema migration caused by that layer. Documentation and tests must not
   be deferred to a final cleanup PR.
2. Every PR must pass the complete quality gate independently when checked out at its head. A later
   layer may reuse an earlier abstraction but may not be required to make an earlier layer correct.
3. Each PR description includes a **Stack** section linking the immediately preceding and following
   PRs, states its GitHub base branch, and identifies which LS finding it closes.
4. Reviewers can approve each layer on its own contract. A PR must not mix drive-by refactors,
   formatting sweeps, generated-file churn unrelated to its schema, or fixes belonging to another
   layer.
5. When an earlier layer changes during review, rebase or restack every descendant and rerun its
   focused tests plus the full quality gate. Do not merge descendants out of order.
6. Merge bottom-up. After each base PR merges, retarget the next PR to `main`, confirm GitHub shows
   only that layer's intended diff, and continue until the stack is exhausted.
7. Public schema versions advance in the PR that first changes their serialized shape. The final
   stack therefore ends at text-inspection v3, replace-result v2, import-report v3, and deck-diff
   v5; version numbers must not be retroactively reused for a second payload shape.
8. Shared helpers belong in the earliest PR that needs them and must earn their place in that PR.
   Do not create a foundation-only PR containing unused abstractions.

## Requirements

### 1. Structural text anchors

#### Functional requirements

1. Every anchor emitted by `inspect_text()`, `replace_text()`, `replace_text_at()`, or `refind()`
   must carry a versioned structural locator in addition to its part, diagnostic block ordinal, and
   content fingerprint.
2. The locator must distinguish these container forms without relying on display names:
   - a top-level shape by slide-unique shape ID;
   - a grouped leaf shape by its slide-unique shape ID; group-name or group-order paths are not
     identity;
   - a table cell by the graphic-frame shape ID, row, column, and paragraph index within the cell;
   - a notes text body by notes part, owning shape ID, and paragraph index within that shape.
3. A paragraph ordinal is valid only within its structurally identified container. A slide-global
   block index remains diagnostic and must not be the write target for a current-version anchor.
4. Anchor resolution must locate exactly one supported container and exactly one paragraph within
   it. A missing locator, duplicate identity, incompatible container kind, or structural change
   that prevents unique resolution raises a typed refusal before mutation.
5. Only after structural resolution succeeds may the editor compare the current content
   fingerprint. A mismatch raises `StaleAnchorError` and leaves the presentation unchanged.
6. Current-version fingerprints use full SHA-256 rather than the existing eight-character prefix.
   They include NFC-normalized literal text plus field marker types and order, while excluding
   volatile rendered field values.
   The existing public `content_hash()` helper retains its pinned eight-character legacy behavior;
   current anchors use a separate full-fingerprint path.
7. Structural identity never weakens content validation. A matching shape or cell with changed text
   is stale, not a valid target.
8. A current anchor's fingerprint must identify exactly one paragraph inside its structurally
   identified container before a write is allowed. If identical paragraphs make the local ordinal
   unstable under insertion or reorder, resolution raises `AmbiguousTargetError`; the ordinal must
   not break that tie.
9. A legacy three-field anchor remains accepted for compatibility, but it must not be resolved by
   ordinal plus hash. It may resolve only when its old fingerprint identifies exactly one block in
   the named part; otherwise it raises `AmbiguousTargetError` or `TargetNotFoundError`.
10. `refind()` may recover only an exact, unique target consistent with the anchor's available
   structural identity. It must never use nearest text, edit distance, shape order, or geometry.
11. Successful anchored edits retain the existing formatting-preservation, field-boundary,
    transaction, and changed-part contracts.

#### Interface contract

- `BlockAnchor` remains the public value passed to `replace_text_at()` and `refind()`.
- Current anchors identify their schema version and expose their structural locator in
  `to_dict()`; the existing `part`, `block_index`, and `content_hash` keys remain available during
  compatibility support.
- `paper-text-inspection` advances from version 2 to version 3 because its nested anchor payload is
  extended.
- `paper-replace-result` advances from version 1 to version 2 because its `blocks` collection
  serializes the same extended anchor payload.
- Existing code constructing `BlockAnchor(part, block_index, content_hash)` remains valid but gets
  the conservative legacy-resolution behavior above.

### 2. Unique-only import layout selection

#### Functional requirements

1. An explicit destination `target_layout` remains authoritative after existing ownership and
   enrollment validation.
2. Without an explicit target, automatic selection evaluates the existing tiers in order: exact
   layout name, exact non-custom layout type, and—only for bake mode—blank layout.
3. A tier succeeds only when it contains exactly one candidate.
4. Multiple candidates at a stronger tier raise `AmbiguousTargetError`; the resolver must not drop
   to a weaker tier after ambiguity has already been established.
5. An ambiguity refusal lists every candidate's layout name, type, part name, and owning master so
   the caller can supply `target_layout`.
6. Zero candidates at a tier may continue to the next documented tier. If no unique candidate
   exists, the operation retains the existing typed unsupported/no-match refusal.
7. Bake mode must not fall back to the first destination layout. No unique blank layout means the
   caller supplies `target_layout`.
8. `append_deck()` applies the same rules while staging every source slide. Any ambiguity refuses
   the whole append before the first destination write.
9. `ImportReport.layout_binding_method` continues to distinguish explicit, unique name, unique
   type, unique blank, and transplanted bindings. The obsolete `first-fallback` result is removed.

### 3. Unique-only placeholder reconciliation

#### Functional requirements

1. The existing global exact type-plus-index pass remains unchanged and runs before weaker tiers.
2. An unmatched source placeholder may bind automatically by exact type only when exactly one
   unclaimed target slot has that type.
3. If no same-type candidate exists, it may bind by the existing compatible type family only when
   exactly one unclaimed target slot belongs to that family.
4. More than one candidate in either fallback tier raises `AmbiguousTargetError` before mutation.
   The refusal identifies the source placeholder and all candidate target type/index pairs and
   instructs the caller to provide `placeholder_map`.
5. Existing explicit-map validation remains authoritative: source and target indexes must exist,
   target claims are one-to-one, and `None` deliberately orphans a source placeholder.
6. `Slide.rebind_layout()` keeps its existing `placeholder_map` interface and atomicity contract.
7. Adopt-theme `Presentation.import_slide()` gains the same optional `placeholder_map` contract and
   passes it through the shared mapper. The argument is rejected for modes where placeholder
   reconciliation does not use it.
8. Adopt-theme import reports the resolved placeholder mapping so an automatic or explicit binding
   can be audited after the operation.
9. `append_deck()` remains automatic-only and atomically refuses if any slide requires an explicit
   placeholder map.

#### Interface contract

- `Presentation.import_slide()` and `pptx.compose.import_slide()` accept optional
  `placeholder_map`, with the same `{source_idx: target_idx | None}` meaning as
  `Slide.rebind_layout()`.
- `paper-import-report` advances from version 1 to version 2 and adds
  `placeholder_map_used`. It is always serialized, using an empty list for modes without
  placeholder reconciliation, matching the report's existing fixed-field convention.

### 4. Stable shape matching in `diff_decks()`

#### Functional requirements

1. Within permanent-slide-ID pairs, top-level shapes match by slide-unique shape ID, not by name,
   element-kind ordinal, or z-order.
2. Shape names remain display metadata. Duplicate, empty, or changed names do not alter identity.
3. A z-order-only change must not produce shape additions, removals, geometry changes, image
   replacements, or chart changes.
4. A shape ID present on only one side is an addition or removal. A reused ID whose element kind is
   incompatible across sides is represented as removal plus addition rather than a matched change.
5. Moving a shape into or out of a group retains the existing documented top-level add/remove
   semantics; this feature does not introduce recursive visual matching.
6. Shape-related diff entries expose stable shape ID separately from the human-readable name. They
   must not manufacture synthetic `kind#ordinal` identity labels.
7. The lineage limitation and same-kind shape-ID reuse hazard are documented beside the existing
   permanent-slide-ID limitation.

#### Report contract

- `paper-deck-diff` advances from version 3 to version 4 in the LS-04 PR.
- Shape additions, removals, geometry changes, image replacements, and chart changes use structured
  shape references containing stable shape ID and display name.
- Existing detail-level gates and empty-facet omission remain unchanged.

### 5. Exact paragraph and table-cell diff alignment

#### Functional requirements

1. Text comparison first partitions slide blocks by stable structural container: leaf shape
   identity for ordinary/grouped text and table-frame shape identity for tables. Existing flat
   `notes_change` comparison remains unchanged.
2. Within each stable container, paragraphs align through one deterministic exact sequence
   algorithm over structural fingerprints rather than paragraph ordinal.
3. A paragraph fingerprint includes NFC-normalized literal text and field marker types/order. It
   does not include volatile rendered field values and does not use fuzzy similarity.
4. An insertion or deletion before unchanged paragraphs must align the unchanged paragraphs and
   report only the actual insertion/deletion.
5. Repeated identical fingerprints are matched only when their correspondence is uniquely
   established by exact surrounding sequence. Ambiguous repeated paragraphs remain unmatched or
   are explicitly marked ambiguous; they are never paired by ordinal.
6. Table blocks carry structured frame shape ID, row, column, and paragraph-within-cell data through
   inspection and diff reporting. Consumers must not parse the current display-oriented
   `container_detail` string to recover coordinates.
7. Table text alignment operates within the stable table frame across the row-major block sequence,
   retaining before/after cell coordinates as context. A row or column insertion must not turn
   unchanged cell values into a cascade of replacements.
8. Table row/column-count changes are reported as structural table changes. When exact evidence is
   insufficient to locate an insertion among duplicate cells, the report states that ambiguity
   rather than inventing cell replacements.
9. The same exact alignment primitive is reused by text changes, effective-value shifts, and bullet
   shifts wherever those facets compare the same paragraph domain. Table formatting remains outside
   effective-value resolution until table-style inheritance is supported.
10. Existing detail-level gates remain: structural table changes appear at `structure`; text and
    field changes require `text`; effective and bullet changes require `full`.

#### Report contract

- `paper-deck-diff` advances from version 4 to version 5 in the LS-05 PR.
- Shape and text entries contain structured identity and before/after location fields instead of
  relying on synthetic ordinal labels.
- Location-aware effective-shift serialization is owned by deck diff v5. It must not change the
  shared `pptx.rebind.RunShift` payload or the already-versioned import/rebind reports that embed it.
- The schema can represent an insertion, deletion, replacement, move/alignment, or ambiguity
  without encoding an uncertain association as a replacement.
- Empty facets remain omitted and serialization remains deterministic and golden-tested.

### 6. Unambiguous section selection

#### Functional requirements

1. Section lookup by name gathers all exact matches.
2. No match raises `TargetNotFoundError`; one match succeeds; multiple matches raise
   `AmbiguousTargetError` before import mutation.
3. The ambiguity refusal lists each matching section's GUID and deck order.
4. Callers can select an existing section by its `p14:section/@id` GUID. Section IDs are compared
   exactly and must resolve to one enrolled section.
5. Name and ID selection are mutually exclusive. Invalid types or specifying both are caller errors;
   a well-formed but missing ID is `TargetNotFoundError`.
6. Preflight and enrollment use the same already-resolved section element so the operation cannot
   validate one section and later mutate a different one.
7. `ImportReport` records both the enrolled section name and stable section ID.

#### Interface contract

- The existing `section=` name argument remains supported with unique-only semantics.
- `Presentation.import_slide()` and `pptx.compose.import_slide()` add optional `section_id=` for
  stable selection.
- `paper-import-report` advances from version 2 to version 3 and adds `section_id`; it retains the
  version-2 `placeholder_map_used` field unchanged.
- Section selection does not create, rename, reorder, or expose a new public section collection.

### 7. Error, atomicity, and documentation contracts

1. Multiple valid candidates use `AmbiguousTargetError`; absent targets use
   `TargetNotFoundError`; changed anchored content uses `StaleAnchorError`; unsupported structural
   identity uses `UnsupportedStructureError`.
2. Every mutating refusal occurs before mutation or inside the existing package transaction and
   leaves memory and output bytes unchanged.
3. Error messages name what matched, why automatic selection is unsafe, and which explicit argument
   resolves it.
4. Authoritative RST documentation for inspect/edit, compose/import, rebind, diff, errors, and the
   Paper additions overview reflects the new contracts and schema versions.
5. Documentation must not claim that content hashes alone provide object identity or that ordinal
   fallback is honest ambiguity handling.
6. No new behavior changes inherited upstream APIs outside the named Paper surfaces.

## Test strategy

### What earns its keep

- Preserve the minimal LS-01 wrong-write reproduction as a fixture-backed regression: two
  identical-text shapes reorder, and the old anchor must still identify the original shape or
  refuse without changing either shape.
- Cover current anchors in top-level shapes, nested groups, multi-paragraph text bodies, table
  cells, notes, and field-bearing paragraphs; cover missing identities, stale content, structural
  movement, duplicate text, and conservative legacy-anchor recovery.
- Exercise duplicate layout names across one and multiple masters, duplicate types, multiple blank
  layouts, explicit targets, layout reordering, and atomic `append_deck()` refusal.
- Exercise placeholder exact matches, unique same-type/family fallbacks, ambiguous fallback tiers,
  explicit maps, adopt-theme import maps, orphan behavior, and output-byte equality on refusal.
- Preserve the eval-observed LS-04 case as a regression: duplicate-named shapes reorder without
  fictional additions, removals, or geometry changes.
- Preserve the eval-observed LS-05 case as a regression: table row/column expansion reports the
  grid change without dozens of shifted text replacements. Add ordinary paragraph insertion,
  deletion, reorder, duplicate text, fields, and groups cases. Preserve the existing notes-diff
  contract unchanged.
- Cover duplicate section names, exact section IDs, name/ID exclusivity, missing IDs, report
  identity, and byte-identical refusal.
- For every successful mutation, save, reopen, and assert both the requested effect and nearby
  state that must remain stable.
- Pin exact changed-part budgets for successful mutators and byte equality for every new refusal.
- Update deterministic inspection, import, and deck-diff goldens for their schema-version changes.
- In each stack layer, inspect the GitHub comparison against its immediate base and assert that the
  diff contains only that PR's finding, tests, documentation, and required generated artifacts.
- Run the complete repository quality gate: `uv run pytest`, `uv run behave`,
  `uv run pytest -m lo_smoke tests/paper`, `uv run make docs`, and `uv run make build`.

### What does not need separate tests

- Dataclass construction or enum/value lists apart from behavior exercised through public reports.
- General SHA-256 correctness supplied by the standard library; test the package's normalization
  and fingerprint contents instead.
- Fuzzy or geometric matching, because those behaviors are explicitly prohibited.
- Independently authored deck matching, rendering, or visual similarity.

## Risks and mitigations

- **Risk: stricter matching causes previously successful ambiguous calls to refuse.** Mitigate with
  actionable errors, `target_layout`, `placeholder_map`, and `section_id`, while preserving every
  unique automatic case.
- **Risk: anchor schema changes break persisted payload consumers.** Retain the three legacy fields,
  accept legacy construction, version inspection output, and resolve legacy anchors only through
  exact unique recovery.
- **Risk: stable shape IDs can be reused after deletion.** Detect incompatible-kind reuse, document
  the remaining same-kind lineage hazard, and never add approximate matching to conceal it.
- **Risk: exact sequence alignment under-reports repeated text.** Represent ambiguity or unmatched
  blocks explicitly; conservative under-association is preferable to fictional replacements.
- **Risk: report schema changes break evaluators and golden consumers.** Version the inspection,
  import, and deck-diff payloads, update all in-repo consumers and goldens together, and document
  the migration.
- **Risk: multiple consumers implement different paragraph identity rules.** Keep one shared exact
  alignment/fingerprint boundary for text, effective-value, and bullet reporting.

## Open questions

- None. Structural locator ownership remains in `inspect.py`, paragraph alignment remains private
  to `diff.py`, and import reports always serialize `placeholder_map_used`.
