# Phase 2 — Unique Import Layout Selection

**Created:** 2026-08-22
**Status:** Complete

## Motivation

Automatic slide import currently treats the first layout in collection order as the intended
layout when several layouts have the same name or type, and bake mode ultimately accepts the first
destination layout even when it has no semantic relationship to the source. This phase makes those
choices safe and reviewable: automatic binding succeeds only at an exact, unique tier, while an
ambiguous or unmatched deck refuses before mutation and tells the caller to provide an explicit
`target_layout`.

## Context

**What exists today:**

- `src/pptx/compose.py` validates an explicit `target_layout` for destination ownership and
  enrollment before `_resolve_layout_binding()` is called. The resolver then checks exact name,
  exact non-custom type, bake-mode blank, and finally the first destination layout in order.
- `_resolve_layout_binding()` returns immediately on the first candidate at each automatic tier.
  It therefore cannot distinguish a unique match from duplicate layouts on one master or across
  several masters.
- `append_deck()` stages the transplant plan, layout binding, and mode preparation for every source
  slide before opening its write transaction. This preflight boundary already provides the right
  place to refuse an ambiguous later slide without partially appending earlier slides.
- `ImportReport.layout_binding_method` records `name-match`, `type-match`, `explicit`, `transplant`,
  `blank-fallback`, or `first-fallback`. The report schema remains version 1 in this phase because
  the serialized field set does not change.
- `tests/paper/test_import.py` exercises explicit, name, type, transplant, unmatched, whole-deck
  preflight, report, save/reopen, and relationship-integrity behavior against the frozen fixture
  corpus.

**What this phase delivers:**

- Unique-only automatic layout binding at the existing name, type, and bake blank tiers.
- An actionable `AmbiguousTargetError` when a matching tier contains multiple layouts, without
  falling through to a weaker tier.
- Removal of bake mode's arbitrary first-layout fallback.
- Preservation of explicit-target and keep-appearance behavior, and byte-identical whole-deck
  refusal when any staged slide has an ambiguous layout.
- Updated import documentation and report vocabulary with no public schema-version change.

**Reference files to study before starting:**

- `docs/plans/selection-integrity-hardening/spec.md` — authoritative LS-02 behavior and stack
  contract.
- `docs/plans/lossy-selection-audit.md` — concrete layout ambiguity reproduction, impact, and
  rejected geometric matching approaches.
- `src/pptx/compose.py` — argument validation, layout resolution, `ImportReport`, import preflight,
  and append staging.
- `src/pptx/presentation.py` — public `Presentation.import_slide()` documentation and argument
  forwarding.
- `src/pptx/errors.py` — existing `AmbiguousTargetError` and typed-refusal semantics.
- `tests/paper/test_import.py` — established import fixtures, atomicity assertions, report checks,
  and save/reopen validation.
- `tests/paper/contract.py` — byte-equality, changed-part, and save/reopen test helpers.
- `tests/paper/fixtures/self_generated/template_alpha.pptx` and
  `tests/paper/fixtures/self_generated/template_beta.pptx` — existing corpus inputs from which
  duplicate-name, duplicate-type, blank, and no-match conditions can be constructed without a new
  broad fixture family.
- `docs/api/compose.rst` and `docs/user/paper-additions.rst` — authoritative import and Paper
  workflow documentation.
- `CONTRIBUTING.md` — pre-mutation validation, refusal atomicity, reopen, fixture, changed-part,
  and compatibility requirements.

## Steps

### Step 1 — Make each automatic binding tier exact and unique

**Goal:** Automatic adopt-theme and bake imports bind only when the strongest applicable tier has
exactly one destination layout; neither mode ever selects a layout because it happens to appear
first.

**Work:**

Update the layout resolution boundary in `src/pptx/compose.py` to evaluate all destination layouts
at each existing tier. Preserve the current precedence: explicit target, exact source layout name,
exact non-custom source layout type, and bake-only blank layout. A tier with one candidate returns
that binding; an empty tier may continue; a tier with multiple candidates refuses immediately with
`AmbiguousTargetError`.

Make each ambiguity message identify every candidate by layout name, layout type, layout part name,
and owning master part name, and direct the caller to `target_layout`. Candidate diagnostics must
be deterministic so reordered presentation collections do not produce unstable errors. When every
automatic tier is empty, retain the existing typed no-match behavior and explicit-target guidance.
For bake mode, delete `first-fallback`: a unique blank layout is the final automatic option.

Keep the existing binding-method values for successful paths. `name-match`, `type-match`, and
`blank-fallback` now mean unique matches; `explicit` and `transplant` retain their meanings.
Remove `first-fallback` from runtime and report documentation without changing the report's fields
or schema version.

**Constraints:**

- An ambiguous stronger tier must refuse; it must not be bypassed by a unique weaker tier.
- Layout matching remains exact. Do not introduce geometry, placeholder similarity, fuzzy names,
  master order, layout order, or other scoring.
- Do not change inherited `SlideLayouts.get_by_name()`; LS-08 is explicitly outside this feature.
- Explicit `target_layout` remains authoritative after existing type, package-ownership, and
  enrollment validation.
- `keep_appearance` continues to return a transplanted binding without inspecting destination
  layout candidates.
- Keep the change local to composition. Do not create a general matching framework or a new public
  helper for one resolver.
- Do not advance `paper-import-report`; this phase changes the allowed values of an existing field,
  not its serialized shape.

**Verification:**

- `uv run pytest -q tests/paper/test_import.py -k "layout or adopt_theme or bake"`
- `uv run ruff check src/pptx/compose.py tests/paper/test_import.py`

### Step 2 — Prove unique selection, actionable refusal, and append atomicity

**Goal:** Focused regressions demonstrate that unique candidates remain compatible, every
ambiguous tier refuses with enough identity to recover explicitly, and no import mutates either
deck before a decision is complete.

**Work:**

Extend `tests/paper/test_import.py` using the existing frozen import corpus and narrowly scoped
in-memory layout metadata changes. Cover a unique name match, name absent followed by unique type,
and bake absent-name/type followed by unique blank. Add duplicate-name cases both within one master
and across masters, duplicate-type and duplicate-blank cases, and verify that a duplicate at a
stronger tier refuses rather than falling through. Assert ambiguity diagnostics contain every
candidate's identifying fields and the `target_layout` recovery instruction.

For each ambiguity class, prove that supplying an enrolled destination `target_layout` succeeds,
survives save/reopen, and reports `explicit`. Add an order-variation regression demonstrating that
reordering candidate layouts cannot turn ambiguity into a silent choice. Add a destination with no
usable automatic candidate, including the no-enrolled-layout edge when representable through the
public collection contract, and assert a typed no-match refusal rather than an index error.

Add an `append_deck()` regression whose later source slide encounters layout ambiguity. Assert the
destination package is byte-identical to its pre-call state and no earlier source slide was
appended. Preserve the existing successful append and keep-appearance tests to demonstrate that
preflight and transplant behavior did not regress.

Update `src/pptx/presentation.py`, `docs/api/compose.rst`, and
`docs/user/paper-additions.rst` so callers know automatic matches must be unique, ambiguity is a
typed pre-write refusal, bake has no first-layout fallback, and explicit `target_layout` is the
recovery path. Update the `ImportReport` documentation in `src/pptx/compose.py` to remove the
obsolete method value.

**Constraints:**

- Refusal tests must compare the destination before and after, not merely assert an exception.
- Successful mutation assertions must save and reopen the presentation before checking the bound
  layout and slide content.
- The tests must exercise real layout/master collections and the import entry points; do not test
  only a detached helper or replace package behavior with mocks.
- Prefer the existing fixture corpus with deliberate in-memory ambiguity construction. Add or
  regenerate a binary fixture only if the relevant master/layout structure cannot be represented
  reliably that way, and document any new fixture in `tests/paper/fixtures/README.md` and its
  manifest.
- Keep report-schema goldens unchanged unless a successful fixture's binding method legitimately
  changes from an obsolete fallback. Do not introduce unrelated golden churn.
- Do not expand this PR into placeholder reconciliation, section identity, report v2, or upstream
  layout-collection semantics.

**Verification:**

- `uv run pytest -q tests/paper/test_import.py`
- `uv run ruff check src/pptx/compose.py src/pptx/presentation.py tests/paper/test_import.py`
- `uv run sphinx-build -W -b html docs docs/.build/html`
- `git diff --check`

## Files

| Action | Path |
|---|---|
| Edit | `src/pptx/compose.py` — require unique candidates, emit ambiguity diagnostics, remove the first-layout fallback, and refresh report documentation |
| Edit | `src/pptx/presentation.py` — document the unique automatic-selection and explicit-target contract on the public method |
| Edit | `tests/paper/test_import.py` — add unique, ambiguous, unmatched, explicit-recovery, order, and append-atomicity regressions |
| Edit | `docs/api/compose.rst` — document tier precedence, typed ambiguity, and bake behavior |
| Edit | `docs/user/paper-additions.rst` — summarize safe automatic binding and the explicit recovery path |
| Conditional edit | `tests/paper/fixtures/README.md` and fixture manifest — only if a new binary fixture is strictly necessary |
| Conditional create | `tests/paper/fixtures/self_generated/<focused-layout-ambiguity-fixture>.pptx` — only if the existing corpus cannot express the regression reliably |

## What this phase does NOT include

- Structural text anchors or inspection schema v3; phase 1 owns those changes.
- Placeholder fallback uniqueness, import `placeholder_map`, or import-report v2; phase 3 owns
  those changes.
- Section-name or section-ID selection; phase 4 owns those changes.
- Shape or paragraph matching in `diff_decks()`; phases 5 and 6 own those changes.
- Changes to upstream `SlideLayouts.get_by_name()` or a new public strict layout-collection API.
- Layout selection by visual similarity, geometry, placeholder overlap, fuzzy name matching, or
  any other best-effort heuristic.
- Per-slide explicit mapping for `append_deck()`.
- Runtime dependencies, broad composition refactors, or unrelated fixture regeneration.

## Tests this phase must include

- A single exact name candidate binds and reports `name-match`.
- With no name candidate, a single exact non-custom type candidate binds and reports `type-match`.
- With no name or type candidate, bake mode binds a single blank layout and reports
  `blank-fallback`.
- Duplicate exact names on one master and across masters raise `AmbiguousTargetError`, enumerate
  every candidate, mention `target_layout`, and leave the destination byte-identical.
- Duplicate exact non-custom types raise the same typed, atomic refusal when the name tier is
  empty.
- Multiple blank layouts raise the same typed, atomic refusal when the earlier tiers are empty.
- Ambiguity at a stronger tier is not bypassed by a unique candidate at a weaker tier.
- Reordering masters or layouts does not silently change the selected layout; an ambiguous set
  remains ambiguous and its diagnostic ordering remains deterministic.
- A valid explicit destination layout succeeds despite ambiguous automatic candidates, survives
  save/reopen, and reports `explicit`.
- No automatic candidate produces the existing typed no-match refusal with `target_layout`
  guidance; a destination with no layouts never leaks `IndexError`.
- `keep_appearance` remains independent of destination layout ambiguity and reports `transplant`.
- `append_deck()` with an ambiguous later slide appends nothing and leaves destination bytes
  unchanged.
- Existing successful adopt-theme, bake, keep-appearance, report-golden, relationship-integrity,
  source-nonmutation, and changed-part-budget tests remain green.

## Done when

1. No automatic import path chooses a destination layout from a multi-candidate tier or from raw
   collection order.
2. Every ambiguity refusal identifies all candidates deterministically, tells the caller to pass
   `target_layout`, raises `AmbiguousTargetError`, and occurs before destination mutation.
3. Unique name, unique type, unique blank, explicit, and transplant paths retain their documented
   behavior and binding-method reports after save/reopen.
4. Bake mode has no first-layout fallback, and `first-fallback` is absent from production code and
   current public documentation.
5. Whole-deck append preflights all layout choices and remains byte-identically atomic when any
   source slide is ambiguous.
6. `paper-import-report` remains schema version 1 and existing golden payload shape is unchanged.
7. PR 2 contains only LS-02 production, regression, and documentation changes, is based on
   `gavin/block-anchor-structural-identity`, and is independently reviewable as
   `gavin/import-layout-unique-selection`.
8. Focused checks pass:
   `uv run pytest -q tests/paper/test_import.py`,
   `uv run ruff check src/pptx/compose.py src/pptx/presentation.py tests/paper/test_import.py`,
   `uv run sphinx-build -W -b html docs docs/.build/html`, and `git diff --check`.
9. The complete PR quality gate passes:
   `uv run pytest`, `uv run behave`, `uv run pytest -m lo_smoke tests/paper`,
   `uv run make docs`, and `uv run make build`.
