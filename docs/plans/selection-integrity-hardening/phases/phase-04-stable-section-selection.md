# Phase 4 — Stable Section Selection

**Created:** 2026-08-22
**Status:** Complete

## Motivation

Named slide import currently validates the first section with a matching name and then performs a
second first-match lookup during enrollment. Duplicate names can therefore direct a slide into an
arbitrary section, and resolving the target twice leaves validation and mutation dependent on the
same unstable collection order. This phase makes explicit section enrollment exact and auditable:
names must be unique, callers can select by stable section GUID, and one preflight result is carried
unchanged into enrollment.

## Context

**What exists today:**

- `src/pptx/compose.py` accepts `section` as an optional name. `_validate_arguments()` calls
  `_find_section()` before mutation, but `_enroll_in_section()` calls `_find_section()` again after
  the imported slide has been added to the destination sequence.
- `_find_section()` returns the first exact name match in section-list order. It does not detect
  duplicate names or expose the existing `p14:section/@id` GUID as an identity path.
- `_enroll_in_section()` has two deliberate modes: an explicitly named section receives the new
  slide at its end, while an import with no selector enrolls beside the preceding slide or in the
  first section when no predecessor is enrolled. The adjacent mode must remain unchanged.
- The frozen `tests/paper/fixtures/self_generated/sections.pptx` deck contains deterministic
  section names, GUIDs, and memberships. `tests/paper/test_import.py` already covers named,
  adjacent, missing-name, package-integrity, and save/reopen behavior against that fixture.
- Phase 3 leaves `paper-import-report` at version 2 with a fixed-field
  `placeholder_map_used` payload. Its existing `section` field records the actual enrolled name,
  including a section chosen by adjacent enrollment.

**What this phase delivers:**

- One conservative preflight resolver for explicit section name or exact section GUID selection.
- Typed, actionable refusal for missing and ambiguous selectors before any destination mutation.
- A public `section_id` import argument that is mutually exclusive with `section`.
- Enrollment that consumes the already-resolved explicit section rather than repeating lookup.
- `paper-import-report` version 3 with the actual enrolled section GUID, while retaining every
  version-2 field and the exact `placeholder_map_used` representation.
- Focused regression coverage and public compose documentation in PR 4 of the GitHub stack.

**Reference files to study before starting:**

- `docs/plans/selection-integrity-hardening/spec.md` — authoritative LS-06 behavior, interface,
  schema, and stack contract.
- `docs/plans/lossy-selection-audit.md` — duplicate-section reproduction, impact, provenance, and
  recommended unique-only remedy.
- `src/pptx/compose.py` — `ImportReport`, import preflight, transaction boundary, section lookup,
  section enrollment, and report construction.
- `src/pptx/presentation.py` — public `Presentation.import_slide()` arguments, forwarding, and
  documentation.
- `src/pptx/errors.py` — `AmbiguousTargetError`, `TargetNotFoundError`, and refusal semantics.
- `tests/paper/test_import.py` — existing section, atomicity, report-golden, and save/reopen tests.
- `tests/paper/contract.py` — byte-equality, changed-part, and save/reopen helpers.
- `tests/paper/fixtures/self_generated/sections.pptx` and
  `tests/paper/_authoring/build_fixtures.py` — deterministic section names, GUIDs, ordering, and
  membership used by the regressions.
- `tests/paper/goldens/import_beta_keep.import.json` and
  `tests/paper/_authoring/update_goldens.py` — deterministic import-report schema contract and its
  sole update path.
- `docs/api/compose.rst` and `docs/user/paper-additions.rst` — authoritative import API and
  workflow guidance.
- `CONTRIBUTING.md` — pre-mutation validation, refusal atomicity, reopen, fixture, changed-part,
  compatibility, and quality requirements.

## Steps

### Step 1 — Resolve an explicit section exactly once during preflight

**Goal:** A caller-supplied section name or GUID produces exactly one destination section element
before mutation, or a typed refusal with enough information to choose safely.

**Work:**

Replace the first-match section lookup in `src/pptx/compose.py` with one private resolution boundary
that considers only sections enrolled in the destination presentation's section list. Preserve
exact name comparison, but gather every matching section: no matches raise `TargetNotFoundError`,
one match resolves, and multiple matches raise `AmbiguousTargetError`. Name ambiguity diagnostics
must list each candidate's GUID and zero-based section-list order, deterministically, and tell the
caller to pass `section_id`.

Add exact GUID selection through `section_id`. A GUID must identify exactly one enrolled section;
no match raises `TargetNotFoundError`, and a duplicated ID refuses as ambiguous rather than taking
the first element. Comparison uses the stored string exactly—including braces and case—and does
not normalize, infer, or fuzzy-match identifiers.

Validate `section` and `section_id` together at the public import preflight boundary. Each accepts
only a string or `None`, and specifying both is a caller error. Resolve an explicit selector once,
before transplant planning or destination writes, and carry that element as the selection result
for the rest of the operation. When both selectors are absent, retain a distinct adjacent-
enrollment path; do not preselect the first section.

**Constraints:**

- Do not expose a public section collection, proxy type, or section-management API.
- Do not search by substring, case folding, GUID normalization, slide proximity, or collection
  position. Name and ID comparison are exact.
- Duplicate names are legal input to the resolver and must produce a typed ambiguity, not a caller
  validation error or silent choice.
- Invalid selector types or supplying both selectors remain `ValueError`-style caller mistakes;
  well-typed selectors that match nothing use `TargetNotFoundError`.
- Candidate diagnostics must be deterministic and must not disclose only the first candidate.
- Keep resolution private to composition; do not create a general selection framework.
- An explicit target must be resolved before the existing import transaction begins so refusal is
  naturally byte-identical.

**Verification:**

- `uv run pytest -q tests/paper/test_import.py -k "section and (duplicate or missing or id or argument)"`
- `uv run ruff check src/pptx/compose.py tests/paper/test_import.py`

### Step 2 — Carry resolved identity through enrollment and report version 3

**Goal:** Mutation uses the exact section selected in preflight, and every successful import reports
the actual enrolled section's name and GUID without changing adjacent-enrollment semantics.

**Work:**

Extend `Presentation.import_slide()` in `src/pptx/presentation.py` and the composition entry point in
`src/pptx/compose.py` with the optional `section_id` selector, forwarding it without reinterpretation.
Pass the preflight's resolved explicit section element through import preparation and into
enrollment. Explicit enrollment must never look the section up again by either name or ID.

Preserve existing enrollment placement: an explicitly selected section receives the imported
slide using the current explicit-section insertion behavior; with neither selector, enrollment
still follows the preceding slide's section and falls back to the first section only when no
preceding section entry can be found. Have enrollment identify the actual destination section it
used so both explicit and adjacent paths report the element's current name and GUID. A destination
without sections continues to report `None` for both fields.

Advance `paper-import-report` from version 2 to version 3 and add a fixed `section_id` field beside
the existing `section` identity. Preserve all version-2 fields, ordering conventions, and values,
especially `placeholder_map_used`; the new field serializes as `null` when the slide was not
enrolled in a section. Update the reviewed import golden through the repository's existing golden
generator so schema version and fixed fields are explicit and deterministic.

**Constraints:**

- Validation and enrollment must share the same resolved explicit XML element; no second name or
  GUID lookup is allowed after mutation starts.
- Do not alter the relative placement of a slide within an explicitly selected or adjacent
  section.
- Do not change `append_deck()`. It supplies no explicit section selector and must retain adjacent
  destination-section behavior.
- Do not create, rename, delete, reorder, copy, or normalize sections or their IDs.
- Report the actual enrolled element, not merely the caller's input. This matters when adjacent
  enrollment is selected and when a section name could change inside a surrounding transaction.
- `placeholder_map_used` remains always serialized exactly as phase 3 defines it; report version 3
  must not reshape or reorder its entries.
- Keep the schema change limited to `paper-import-report` v3; no other public schema advances in
  this phase.

**Verification:**

- `uv run pytest -q tests/paper/test_import.py -k "section or report"`
- `uv run ruff check src/pptx/compose.py src/pptx/presentation.py tests/paper/test_import.py`
- `git diff -- tests/paper/goldens/import_beta_keep.import.json`

### Step 3 — Lock down atomicity, persistence, and public guidance

**Goal:** Regression tests and documentation make unique section selection, stable-ID recovery,
adjacent behavior, and report compatibility independently reviewable in PR 4.

**Work:**

Extend `tests/paper/test_import.py` around the existing `sections.pptx` corpus. Create duplicate-name
and, separately, duplicate-ID conditions narrowly in the destination presentation used by each
test. Prove that ambiguous names enumerate every matching GUID and deck order, missing names and
IDs use `TargetNotFoundError`, selector type and exclusivity errors occur before mutation, and an
exact `section_id` succeeds even when names collide. Include an exactness regression showing that a
case- or brace-altered ID is not silently normalized to a match.

For every refusal, compare serialized destination bytes before and after. For successful name,
GUID, and adjacent imports, save and reopen before checking section membership, new slide position,
and neighboring section entries. Pin the affected package-member budget and verify the source deck
remains unchanged. Assert that reports contain both the actual section name and GUID for explicit
and adjacent enrollment, contain both as `None` when the destination has no sections, remain
deterministic, identify schema version 3, and retain the phase-3 `placeholder_map_used` payload
unchanged.

Update the method documentation in `src/pptx/presentation.py`, the report field documentation in
`src/pptx/compose.py`, and the public guidance in `docs/api/compose.rst` and
`docs/user/paper-additions.rst`. Explain unique-only name selection, exact GUID selection, mutual
exclusivity, adjacent behavior when neither is supplied, typed recovery, and report v3. Keep the
guidance about section identity within slide import; do not imply the package now offers general
section authoring.

**Constraints:**

- Tests must exercise `Presentation.import_slide()` and real section XML, not only a private
  resolver or mocked section list.
- Use the existing deterministic section fixture and report golden. Modify fixture authoring or
  add another binary fixture only if the duplicate condition cannot be represented reliably in a
  focused test; document any fixture change in `tests/paper/fixtures/README.md`.
- Successful assertions follow save → reopen → assert. Refusal assertions prove byte equality,
  not only exception type.
- Error assertions must verify candidate identity and recovery guidance without pinning incidental
  prose that would make harmless wording changes costly.
- Do not fold layout, placeholder, diff, footer, or general section-management changes into this
  PR.

**Verification:**

- `uv run pytest -q tests/paper/test_import.py`
- `uv run ruff check src/pptx/compose.py src/pptx/presentation.py tests/paper/test_import.py`
- `uv run sphinx-build -W -b html docs docs/.build/html`
- `git diff --check`

## Files

| Action | Path |
|---|---|
| Edit | `src/pptx/compose.py` — resolve explicit sections once, enroll the resolved element, expose `section_id`, and advance `ImportReport` to v3 |
| Edit | `src/pptx/presentation.py` — add and document the public `section_id` selector and forward it to composition |
| Edit | `tests/paper/test_import.py` — add name/ID uniqueness, exclusivity, exactness, atomicity, enrollment, persistence, and report regressions |
| Edit | `tests/paper/goldens/import_beta_keep.import.json` — record import-report v3 and the fixed `section_id` field while retaining v2 fields |
| Edit | `docs/api/compose.rst` — document unique name selection, stable ID selection, adjacent behavior, errors, and schema v3 |
| Edit | `docs/user/paper-additions.rst` — summarize safe section targeting and the explicit ID recovery path |
| Conditional edit | `tests/paper/_authoring/build_fixtures.py` and `tests/paper/fixtures/README.md` — only if the existing section fixture cannot express a required regression |
| Conditional edit/create | `tests/paper/fixtures/self_generated/<focused-section-ambiguity-fixture>.pptx` — only if focused in-test setup is insufficient |

## What this phase does NOT include

- Structural anchors or inspection/edit schema changes from phase 1.
- Layout selection changes from phase 2.
- Placeholder matching semantics, import `placeholder_map`, or the shape of
  `placeholder_map_used` from phase 3.
- Shape or paragraph/table-cell matching in `diff_decks()` from phases 5 and 6.
- Section creation, renaming, deletion, reordering, copying from the source deck, or a public
  section collection.
- Section-name case folding, GUID normalization, fuzzy matching, ordinal fallback, or geometry-
  based selection.
- New section selectors on `append_deck()` or per-slide append mappings.
- Changes to upstream python-pptx APIs, new runtime dependencies, broad compose refactors, or
  unrelated fixture/golden regeneration.

## Tests this phase must include

- A unique exact `section` name enrolls the slide in that section, survives save/reopen, and reports
  its actual name and GUID.
- Duplicate exact names raise `AmbiguousTargetError` before mutation, enumerate every candidate's
  GUID and section-list order, mention `section_id`, and leave destination bytes unchanged.
- A missing name raises `TargetNotFoundError` and leaves destination bytes unchanged.
- An exact `section_id` selects the intended section even when another section has the same name,
  survives save/reopen, and reports the selected name and ID.
- A well-typed but missing ID raises `TargetNotFoundError`; an ID with altered case or braces does
  not match by normalization.
- A duplicated section ID raises `AmbiguousTargetError` rather than selecting the first element.
- Non-string selector values and specifying both `section` and `section_id` produce caller errors
  before mutation.
- Supplying neither selector preserves existing adjacent enrollment, including insertion after the
  preceding section entry and first-section fallback when no predecessor is enrolled; the report
  contains that actual section's name and GUID.
- A destination with no sections preserves existing behavior and reports `section=None` and
  `section_id=None`.
- Successful section enrollment has an exact changed-part budget, does not mutate the source deck,
  and leaves section slide IDs valid and non-duplicated after save/reopen.
- Every new refusal is byte-identical, and existing named, adjacent, append, relationship-integrity,
  and section-cleanliness tests remain green.
- `ImportReport.to_dict()` emits schema `paper-import-report`, version 3, a fixed `section_id` field,
  and the unchanged version-2 `placeholder_map_used` structure.
- The reviewed import golden and repeated imports prove deterministic serialization.

## Done when

1. Explicit section selection resolves exactly one enrolled destination section by exact name or
   exact GUID before import mutation.
2. Duplicate names and IDs raise actionable `AmbiguousTargetError`; missing well-typed selectors
   raise `TargetNotFoundError`; invalid argument combinations remain caller errors.
3. Import validation and enrollment consume the same resolved explicit section element, with no
   second selector lookup after mutation begins.
4. Imports without an explicit selector retain adjacent enrollment and first-section fallback
   behavior exactly, including section membership order after save/reopen.
5. `Presentation.import_slide()` and `pptx.compose.import_slide()` expose mutually exclusive
   `section` and `section_id` paths, while `append_deck()` remains unchanged.
6. `paper-import-report` version 3 always serializes the actual `section` and `section_id`, retains
   every version-2 field, and preserves `placeholder_map_used` unchanged.
7. All new refusals leave destination bytes unchanged; successful imports preserve the source,
   reopen cleanly, satisfy section-ID integrity checks, and remain inside their changed-part
   budgets.
8. Current public compose/import documentation explains unique-only names, exact IDs, adjacent
   fallback, typed errors, and the report migration without promising general section management.
9. PR 4 contains only LS-06 production, regression, schema-golden, and documentation changes, is
   based on `gavin/placeholder-rebind-unique-selection`, and is independently reviewable as
   `gavin/section-selection-identity`.
10. Focused checks pass:
    `uv run pytest -q tests/paper/test_import.py`,
    `uv run ruff check src/pptx/compose.py src/pptx/presentation.py tests/paper/test_import.py`,
    `uv run sphinx-build -W -b html docs docs/.build/html`, and `git diff --check`.
11. The complete PR quality gate passes:
    `uv run pytest`, `uv run behave`, `uv run pytest -m lo_smoke tests/paper`,
    `uv run make docs`, and `uv run make build`.
