# Phase 7 — Documentation Hygiene and Compatibility

**Created:** 2026-08-22
**Status:** Complete

## Motivation

The six selection-integrity changes are not complete until the installed docstrings, public RST
guides, README, report contracts, and recovery guidance describe one coherent shipped system.
This final phase performs that reconciliation, proves that each layer of the GitHub stack remains
independently reviewable, and runs the complete compatibility and distribution gates after all
implementation work is present. It is committed into the sixth and top PR,
`gavin/text-diff-exact-alignment`; it does not create a seventh PR.

## Context

**What exists today:**

- Phases 1–6 each own one lossy-selection finding, its production behavior, focused regressions,
  public documentation, and any report-schema migration caused by that layer.
- Phase 1 replaces ordinal-first current anchors with structural container identity followed by a
  full content fingerprint. It retains the public eight-character `content_hash()` behavior for
  legacy consumers and ends at `paper-text-inspection` version 3 and `paper-replace-result`
  version 2.
- Phases 2–4 make automatic layout, placeholder, and section selection unique-only and add the
  explicit `target_layout`, `placeholder_map`, and `section_id` recovery paths. The import report
  advances in its owning layers and ends at `paper-import-report` version 3.
- Phases 5–6 make lineage deck diffs use stable shape IDs and exact conservative paragraph/table
  alignment. The deck-diff report ends at `paper-deck-diff` version 5; notes remain a flat text
  comparison and table-style effective formatting remains unsupported.
- `README.md`, public docstrings under `src/pptx/`, `docs/api/*.rst`, and
  `docs/user/paper-additions.rst` are the durable public surfaces. The planning spec and audit are
  implementation context, not substitutes for public documentation.
- `.github/workflows/test.yml`, `CONTRIBUTING.md`, `pyproject.toml`, and `Makefile` define the
  compatibility gate: Python 3.9–3.13 pytest and Behave, LibreOffice smoke coverage, warning-free
  Sphinx, and validated distribution artifacts with the package-ownership installation matrix.
- This repository has no `AGENTS.md`, `CLAUDE.md`, `.claude/rules`, `.claude/skills`, or
  `.claude/planning-context.md` surface to update or create.

**What this phase delivers:**

- Consistent public terminology for structural identity, content validation, exact matching,
  ambiguity, lineage, and explicit recovery across every affected API and workflow page.
- One accurate final schema narrative: text inspection v3, replacement result v2, import report
  v3, and deck diff v5, including the fields and migration boundaries introduced by each layer.
- Removal or explicit historical labeling of stale claims that content hashes alone identify an
  object, that collection order is an acceptable automatic selector, or that ordinal pairing is a
  reliable diff identity.
- Verification that all six PR branches have the intended immediate GitHub base, contain only
  their owning layer, link the surrounding stack, identify their LS finding, and pass checks at
  their own heads.
- Review of fixture, golden, report, and changed-part-budget changes against the layer that owns
  them, without regenerating or broadening them in this final phase.
- Complete repository, documentation, independent-loader, lint, and distribution validation, then
  final planning-status updates.

**Reference files to study before starting:**

- `docs/plans/selection-integrity-hardening/spec.md` — authoritative behavior, schema, error,
  stack, and definition-of-done contracts.
- `docs/plans/selection-integrity-hardening/build-sequence.md` — dependency order, branch/base
  mapping, cross-cutting rules, and completion gates.
- `docs/plans/selection-integrity-hardening/phases/phase-01-structural-text-anchors.md` through
  `docs/plans/selection-integrity-hardening/phases/phase-06-exact-text-table-diff-alignment.md` —
  the exact behavior, tests, schema changes, and documentation owned by each implementation layer.
- `docs/plans/lossy-selection-audit.md` — provenance and severity context; use only to detect stale
  claims, not as public API documentation.
- `README.md` — top-level perceive/edit/compose/verify promises, safety contract, and limitations.
- `src/pptx/inspect.py` and `src/pptx/edit.py` — installed anchor, fingerprint, legacy recovery,
  result, and refusal documentation.
- `src/pptx/compose.py`, `src/pptx/presentation.py`, `src/pptx/rebind.py`, and
  `src/pptx/slide.py` — installed import, explicit-override, placeholder, section, and report
  documentation.
- `src/pptx/diff.py` — installed lineage, shape identity, text/table alignment, detail-level, notes,
  and report-schema documentation.
- `src/pptx/errors.py` — public typed-refusal meanings.
- `docs/api/inspect.rst`, `docs/api/edit.rst`, `docs/api/compose.rst`,
  `docs/api/rebind.rst`, `docs/api/diff.rst`, and `docs/api/errors.rst` — authoritative API pages.
- `docs/user/paper-additions.rst` — authoritative end-to-end Paper workflow.
- `tests/paper/test_effective_inspect.py`, `tests/paper/test_edit_text.py`,
  `tests/paper/test_table_ops.py`, `tests/paper/test_import.py`, `tests/paper/test_rebind.py`,
  `tests/paper/test_diff.py`, and `tests/paper/test_walkthrough_qbr.py` — executable contracts for
  every claim being reconciled.
- `tests/paper/goldens/`, `tests/paper/fixtures/README.md`,
  `tests/paper/test_fixture_corpus.py`, and `tests/paper/contract.py` — deterministic payload,
  fixture provenance, atomicity, reopen, and changed-part-budget rules.
- `CONTRIBUTING.md`, `.github/workflows/test.yml`, `pyproject.toml`, and `Makefile` — exact local
  and CI quality gates.

## Steps

### Step 1 — Reconcile public terminology, schemas, refusals, and recovery guidance

**Goal:** A caller can understand the final selection-integrity behavior and choose the correct
explicit recovery path from the installed package and rendered docs, without consulting the spec,
audit, eval traces, or PR history.

**Work:**

Sweep `README.md`, the affected public docstrings under `src/pptx/`, the six relevant API pages,
and `docs/user/paper-additions.rst`. Reconcile each statement against the completed phase tests and
final serialized payloads. Keep detailed API contracts in docstrings and API pages, concise
workflow guidance in the Paper additions overview, and only top-level promises and limitations in
the README.

For anchored editing, distinguish a current anchor's structural locator and full fingerprint from
the retained eight-character legacy `content_hash()` helper. Explain that current resolution finds
a unique supported container and paragraph before validating content, while legacy three-field
anchors recover only by an exact unique match. Ensure the docs use the final typed refusals
consistently: missing identity or target, changed content, ambiguous identity, and unsupported
structure must not be collapsed into a generic error. Remove any implication that the diagnostic
slide-global block index is a current write target or that a longer hash alone establishes object
identity.

For composition, document the exact tier order and unique-only rule without promising a fallback
based on destination collection order. Make every ambiguity description name the corresponding
explicit recovery path: `target_layout` for layout selection, `placeholder_map` for adopt-theme
placeholder reconciliation, and `section_id` for duplicate section names. Keep `append_deck()`
automatic-only and direct callers to ordered single-slide imports when per-slide overrides are
needed.

For deck diffing, describe permanent slide IDs, top-level slide-unique shape IDs, structured shape
references, and exact conservative paragraph/table alignment as lineage mechanisms, not visual
matching. State that repeated content may remain unmatched or explicitly ambiguous, table grid
changes are structural, notes retain their existing flat comparison, and table-style effective
format resolution remains unsupported. Preserve the documented same-kind shape-ID reuse hazard
instead of implying perfect identity across deletion and recreation.

Verify that schema names, final versions, required fields, omitted-empty-field behavior, and
detail-level gates agree across production constants, serialized goldens, docstrings, and RST:
`paper-text-inspection` v3, `paper-replace-result` v2, `paper-import-report` v3, and
`paper-deck-diff` v5. Cross-reference the refusal page and public objects rather than duplicating
divergent definitions.

**Constraints:**

- Change only statements made incomplete or stale by phases 1–6. Do not broadly rewrite the README,
  inherited python-pptx documentation, or unrelated API pages.
- Documentation must describe tested public behavior, not internal helper choices, eval tasks,
  trace findings, or task-specific recipes.
- Do not promise geometric, fuzzy, visual, nearest-neighbor, or independently authored deck
  matching. Do not soften a conservative refusal into a best-effort success.
- Do not redefine the legacy public `content_hash()` helper or present its eight-character value as
  current structural identity.
- Do not change runtime behavior to make prose easier to write. Route a substantive mismatch back
  to the phase that owns it, fix and verify it there, restack descendants, then resume this phase.
- Keep notes comparison, table-style effective formatting, append behavior, schema omission rules,
  and the same-kind shape-ID reuse limitation within their tested boundaries.
- Do not add a new documentation format, generated documentation output, or dependency.

**Verification:**

- `rg -n -i "content[- ]hash anchors?|hash alone|block ordinal|slide-global block|first[- ]fallback|first (destination )?layout|kind ordinal|synthetic ordinal|pair(ed|ing)? by ordinal|fuzzy|nearest-neighbor" README.md docs/api/inspect.rst docs/api/edit.rst docs/api/compose.rst docs/api/rebind.rst docs/api/diff.rst docs/api/errors.rst docs/user/paper-additions.rst src/pptx/inspect.py src/pptx/edit.py src/pptx/compose.py src/pptx/presentation.py src/pptx/rebind.py src/pptx/slide.py src/pptx/diff.py`
- `rg -n "paper-text-inspection|paper-replace-result|paper-import-report|paper-deck-diff|SCHEMA_VERSION|RESULT_SCHEMA_VERSION|placeholder_map_used|section_id|shape_id" src/pptx/inspect.py src/pptx/edit.py src/pptx/compose.py src/pptx/diff.py docs/api/inspect.rst docs/api/edit.rst docs/api/compose.rst docs/api/diff.rst docs/user/paper-additions.rst`
- `uv run pytest -q tests/paper/test_effective_inspect.py tests/paper/test_edit_text.py tests/paper/test_table_ops.py tests/paper/test_import.py tests/paper/test_rebind.py tests/paper/test_diff.py tests/paper/test_walkthrough_qbr.py`
- `uv run sphinx-build -W -b html docs docs/.build/html`
- `git diff --check origin/codex/pptx-doc-hygiene...HEAD`

### Step 2 — Audit the six-layer GitHub stack and owned test artifacts

**Goal:** Each PR is based on the intended predecessor, contains only its own finding and required
supporting artifacts, communicates its place in the stack, and is independently green before the
top PR is approved.

**Work:**

Inspect the local branch ancestry and each immediate-base comparison for the six agreed branches:

1. `gavin/block-anchor-structural-identity` targets PR #38's
   `codex/pptx-doc-hygiene` branch at exact remote head `a3c1ae6e` and owns LS-01.
2. `gavin/import-layout-unique-selection` targets the phase 1 branch and owns LS-02.
3. `gavin/placeholder-rebind-unique-selection` targets the phase 2 branch and owns LS-03.
4. `gavin/section-selection-identity` targets the phase 3 branch and owns LS-06.
5. `gavin/shape-diff-stable-identity` targets the phase 4 branch and owns LS-04.
6. `gavin/text-diff-exact-alignment` targets the phase 5 branch, owns LS-05, and also contains this
   final documentation-hygiene work.

For each immediate-base comparison, review commits, file names, and generated artifacts. Confirm
that production code, focused tests, RST/docstring updates, and schema/golden changes live in the
layer that first changes their contract. Any fixture introduced for a regression must be narrowly
named, documented in `tests/paper/fixtures/README.md`, represented by the existing manifest
contract, and absent from unrelated layers. Golden changes must be deterministic and limited to
the inspection, replacement, import, or deck-diff version owned by that PR. Changed-part and
byte-equality expectations must remain explicit in the focused mutator/refusal tests.

Inspect every GitHub PR's head/base metadata and body. Each body must identify its LS finding and
include a **Stack** section linking its immediate predecessor and successor where one exists.
Confirm required GitHub checks pass on each PR head. If an earlier layer changed during review,
restack every descendant and rerun the affected focused checks and complete gate; do not accept an
out-of-date green result from before the restack.

Phase 7 creates no new branch or PR. Its documentation and status changes belong in
`gavin/text-diff-exact-alignment`; GitHub must still show that PR based on
`gavin/shape-diff-stable-identity`, not `main`.

**Constraints:**

- Do not squash layers together, reassign a schema migration to a later PR, or add a foundation-only
  abstraction layer.
- Do not use the cumulative `origin/codex/pptx-doc-hygiene...top` diff to judge an individual PR.
  Review every layer against its immediate base.
- Do not regenerate fixture binaries, manifests, or goldens in this phase. If an artifact is wrong,
  fix it on its owning branch and restack.
- Do not update expected changed-part budgets merely to make a failure pass. Investigate the
  unexpected package delta and return a behavioral defect to its owning phase.
- Do not merge descendants before their base, retarget an open layer away from its immediate
  predecessor, or create a seventh documentation PR.
- Do not treat a local pass as a replacement for the GitHub Python-version and package-ownership
  checks.

**Verification:**

- `git merge-base --is-ancestor origin/codex/pptx-doc-hygiene gavin/block-anchor-structural-identity`
- `git merge-base --is-ancestor gavin/block-anchor-structural-identity gavin/import-layout-unique-selection`
- `git merge-base --is-ancestor gavin/import-layout-unique-selection gavin/placeholder-rebind-unique-selection`
- `git merge-base --is-ancestor gavin/placeholder-rebind-unique-selection gavin/section-selection-identity`
- `git merge-base --is-ancestor gavin/section-selection-identity gavin/shape-diff-stable-identity`
- `git merge-base --is-ancestor gavin/shape-diff-stable-identity gavin/text-diff-exact-alignment`
- `git diff --name-status origin/codex/pptx-doc-hygiene...gavin/block-anchor-structural-identity`
- `git diff --name-status gavin/block-anchor-structural-identity...gavin/import-layout-unique-selection`
- `git diff --name-status gavin/import-layout-unique-selection...gavin/placeholder-rebind-unique-selection`
- `git diff --name-status gavin/placeholder-rebind-unique-selection...gavin/section-selection-identity`
- `git diff --name-status gavin/section-selection-identity...gavin/shape-diff-stable-identity`
- `git diff --name-status gavin/shape-diff-stable-identity...gavin/text-diff-exact-alignment`
- `git diff --name-status origin/codex/pptx-doc-hygiene...gavin/text-diff-exact-alignment -- tests/paper/goldens tests/paper/fixtures`
- `git diff --check origin/codex/pptx-doc-hygiene...gavin/text-diff-exact-alignment`
- `gh pr view gavin/block-anchor-structural-identity --json headRefName,baseRefName,url,body`
- `gh pr view gavin/import-layout-unique-selection --json headRefName,baseRefName,url,body`
- `gh pr view gavin/placeholder-rebind-unique-selection --json headRefName,baseRefName,url,body`
- `gh pr view gavin/section-selection-identity --json headRefName,baseRefName,url,body`
- `gh pr view gavin/shape-diff-stable-identity --json headRefName,baseRefName,url,body`
- `gh pr view gavin/text-diff-exact-alignment --json headRefName,baseRefName,url,body`
- `gh pr checks gavin/block-anchor-structural-identity`
- `gh pr checks gavin/import-layout-unique-selection`
- `gh pr checks gavin/placeholder-rebind-unique-selection`
- `gh pr checks gavin/section-selection-identity`
- `gh pr checks gavin/shape-diff-stable-identity`
- `gh pr checks gavin/text-diff-exact-alignment`

### Step 3 — Run final compatibility gates and record completion

**Goal:** The top of the stack is a tested release candidate that preserves inherited behavior,
Paper safety contracts, independent-loader compatibility, documentation quality, and distribution
ownership across the repository's supported environments.

**Work:**

Run Ruff across the production and Paper test surfaces changed by the stack. Run the complete
pytest suite and Behave suite, then the complete LibreOffice-marked Paper suite in an environment
with headless LibreOffice installed. Build Sphinx with warnings as errors and run the repository's
documented docs target. Build the sdist and wheel and require strict twine validation through the
existing build target.

Review the final working tree for unintended generated output. `docs/.build/`, `dist/`, coverage
files, caches, and compiled Python files must remain ignored and uncommitted. Confirm the source
fixture corpus and deterministic goldens pass their own tests, and that the complete focused suite
from step 1 still passes after the cross-document edits.

Require GitHub CI on the top PR to pass pytest and Behave for Python 3.9, 3.10, 3.11, 3.12, and
3.13; the LibreOffice contract; warning-free documentation; distribution validation; and the
package-ownership installation matrix. Do not substitute the local interpreter's result for the
version matrix.

Only after all local and GitHub checks pass, mark the selection-integrity spec `executed`, the
build sequence `Executed`, and every phase plan `Complete`. Keep those status-only changes in the
top PR. Confirm once more that the top PR targets the phase 5 branch and contains no unrelated
runtime or generated-file change.

**Constraints:**

- Do not weaken, skip, deselect, or rewrite a failing test or compatibility gate to finish the
  phase. Route a behavior failure to its owning phase and rerun descendants after restacking.
- Do not add a mandatory type-check gate or dependency. Pyright remains outside this repository's
  required quality gate.
- Do not add or update release versions, tags, release notes, or publish artifacts.
- Do not commit `docs/.build/`, `dist/`, coverage output, caches, or other generated build products.
- Do not create `AGENTS.md`, `CLAUDE.md`, `.claude/rules`, `.claude/skills`, or
  `.claude/planning-context.md`; the feature does not create those repository surfaces.
- Do not change runtime behavior, fixture content, goldens, manifests, or expected package-part
  budgets in this step. Return substantive defects to the owning phase.
- Do not mark planning status complete before both local gates and required GitHub checks pass.

**Verification:**

- `git diff --check origin/codex/pptx-doc-hygiene...HEAD`
- `uv run ruff check src/pptx tests/paper`
- `uv run pytest`
- `uv run behave`
- `uv run pytest -m lo_smoke tests/paper`
- `uv run sphinx-build -W -b html docs docs/.build/html`
- `uv run make docs`
- `uv run make build`
- `uv run pytest -q tests/paper/test_fixture_corpus.py`
- `git status --short`
- `! git diff --name-only origin/codex/pptx-doc-hygiene...HEAD | rg "(^|/)(docs/\.build|dist|\.coverage|__pycache__)(/|$)"`
- `gh pr checks gavin/text-diff-exact-alignment`
- `gh pr view gavin/text-diff-exact-alignment --json headRefName,baseRefName,state,url`

## Files

| Action | Path |
|---|---|
| Edit as needed | `README.md` — reconcile only the top-level selection, diff, recovery, and limitation summaries made stale by phases 1–6 |
| Edit documentation only | `src/pptx/inspect.py` — reconcile structural-anchor, full-fingerprint, legacy-hash, and inspection-schema docstrings |
| Edit documentation only | `src/pptx/edit.py` — reconcile structural resolution, content validation, legacy recovery, replacement-result, and refusal docstrings |
| Edit documentation only | `src/pptx/compose.py` — reconcile unique layout/placeholder/section selection and final import-report documentation |
| Edit documentation only | `src/pptx/presentation.py` — reconcile public import arguments and explicit recovery guidance |
| Edit documentation only | `src/pptx/rebind.py` — reconcile placeholder tier, explicit-map, ambiguity, and report documentation |
| Edit documentation only | `src/pptx/slide.py` — reconcile the public rebind surface if its installed documentation is stale |
| Edit documentation only | `src/pptx/diff.py` — reconcile lineage identity, exact alignment, table structure, detail-level, and deck-diff v5 docstrings |
| Edit documentation only | `src/pptx/errors.py` — reconcile typed-refusal descriptions only if earlier phases left terminology inconsistent |
| Edit | `docs/api/inspect.rst` — publish current and legacy anchor identity/validation contracts |
| Edit | `docs/api/edit.rst` — publish anchored resolution, staleness, ambiguity, and recovery contracts |
| Edit | `docs/api/compose.rst` — publish unique-only tiers, explicit overrides, section identity, atomicity, and report v3 |
| Edit | `docs/api/rebind.rst` — publish unique-only placeholder reconciliation and explicit-map behavior |
| Edit | `docs/api/diff.rst` — publish shape identity, exact text/table alignment, ambiguity, lineage limits, and report v5 |
| Edit as needed | `docs/api/errors.rst` — make the public refusal taxonomy consistent across selection surfaces |
| Edit | `docs/user/paper-additions.rst` — reconcile the end-to-end perceive/edit/compose/verify workflow and recovery paths |
| Status-only edit after all checks pass | `docs/plans/selection-integrity-hardening/spec.md` — mark the implemented specification executed |
| Status-only edit after all checks pass | `docs/plans/selection-integrity-hardening/build-sequence.md` — mark the confirmed sequence executed |
| Status-only edits after all checks pass | `docs/plans/selection-integrity-hardening/phases/phase-01-structural-text-anchors.md` through `phase-07-documentation-hygiene-compatibility.md` — mark every completed phase |
| Operational edit | GitHub PR metadata for the six agreed branches — correct immediate bases and Stack/LS descriptions without a seventh PR |

## What this phase does NOT include

- New selection, inspection, import, rebind, section, diff, or report runtime behavior.
- A seventh implementation branch or documentation PR; all phase 7 changes belong in
  `gavin/text-diff-exact-alignment`.
- New report fields, schema-version changes, migrations, public APIs, helper abstractions, or
  runtime dependencies.
- Fuzzy, geometric, visual, nearest-neighbor, OCR, or independently authored deck matching.
- Changes to flat notes comparison, table-style effective-value support, append per-slide mapping,
  or same-kind deleted/reused shape-ID detection.
- Fixture or golden regeneration, manifest rewrites, expected changed-part-budget changes, or
  binary churn. Defects in those artifacts return to their owning phase.
- Broad README rewriting, changes to unrelated upstream manuals, or duplication of the entire spec
  into public docs.
- Generated Sphinx HTML, built distributions, coverage data, caches, or compiled artifacts.
- New `AGENTS.md`, `CLAUDE.md`, `.claude/rules`, `.claude/skills`, or
  `.claude/planning-context.md` files.
- A new type-check requirement, package-version bump, release PR, tag, publication, or merge of the
  stacked PRs.

## Tests this phase must include

- The focused inspection, editing, table, import, rebind, diff, and walkthrough tests that pin every
  public statement changed during the documentation sweep.
- Deterministic schema/golden checks proving the final public versions and payload shapes:
  text-inspection v3, replace-result v2, import-report v3, and deck-diff v5.
- `tests/paper/test_fixture_corpus.py` to verify every fixture and manifest remains intentional and
  hash-pinned; phase 7 adds no new fixture content.
- Existing refusal tests that assert the expected typed error and byte equality for ambiguous,
  missing, stale, and unsupported selections.
- Existing successful-mutator tests that save, reopen, assert the requested effect and nearby
  stability, and pin exact changed-part budgets.
- The complete pytest suite, covering inherited python-pptx behavior and all Paper contracts.
- The complete Behave suite, preserving inherited acceptance behavior.
- The complete `lo_smoke` Paper suite in an environment with headless LibreOffice installed. A
  local skip does not replace the CI LibreOffice job.
- Sphinx HTML with warnings treated as errors, plus the documented docs target.
- Ruff across `src/pptx` and `tests/paper`.
- The sdist/wheel build and strict twine validation, followed by the existing CI
  package-ownership installation matrix.
- Required GitHub checks at each of the six PR heads after the final restack.
- No new documentation-only unit test by default. Add one only in the owning implementation phase
  if a public behavior lacks an executable contract.

## Done when

1. README, installed docstrings, API RST, and the Paper workflow use the same terms for structural
   identity, full content validation, exact unique selection, ambiguity, lineage, and recovery.
2. No current public statement claims that a content hash alone identifies an object, that a
   diagnostic block ordinal is a current write target, or that collection/ordinal order is a safe
   ambiguity resolver. Historical limitations are clearly labeled as such.
3. Anchor documentation distinguishes current structural locators from conservative legacy
   recovery and maps missing, stale, ambiguous, and unsupported conditions to the tested typed
   refusals.
4. Composition documentation describes unique-only name/type/family tiers and directs callers to
   `target_layout`, `placeholder_map`, or `section_id` without widening `append_deck()`.
5. Diff documentation describes slide and shape lineage identity, exact conservative text/table
   alignment, table structure, duplicate ambiguity, the same-kind ID-reuse hazard, flat notes, and
   unsupported table-style effective values without promising visual matching.
6. Production constants, golden payloads, docstrings, and RST agree on text-inspection v3,
   replace-result v2, import-report v3, and deck-diff v5, including their final fields and
   omission rules.
7. All public cross-references render without warnings, and the shipped behavior is understandable
   without reading the planning spec, lossy-selection audit, eval traces, or downstream skills.
8. Every local branch is an ancestor of the next branch in the six-layer order, and every GitHub PR
   has the intended immediate base, LS finding, Stack links, and independently green required
   checks.
9. Each immediate-base comparison contains only its owning behavior, focused regressions, public
   documentation, schema migration, and narrowly required artifacts; phase 7 appears only in the
   sixth PR and creates no seventh PR.
10. Fixture, manifest, golden, byte-equality, save/reopen, and changed-part-budget changes are
    deterministic, pass their contract tests, and remain in their owning layers with no phase 7
    regeneration or expectation loosening.
11. `git diff --check origin/codex/pptx-doc-hygiene...HEAD` and
    `uv run ruff check src/pptx tests/paper` pass.
12. The complete pytest and Behave suites pass locally, and required GitHub checks are green for
    Python 3.9 through 3.13.
13. The complete LibreOffice-marked Paper suite passes where headless LibreOffice is installed.
14. Sphinx builds with warnings as errors, the documented docs target passes, and no generated
    documentation output is committed.
15. The sdist and wheel pass strict twine validation, the CI package-ownership installation matrix
    is green, and no built artifact is committed.
16. No new runtime/documentation dependency, type-check requirement, release version, fixture or
    golden churn, agent instruction file, planning-context file, or unrelated refactor is introduced
    by this phase.
17. Only after criteria 1–16 pass, the spec and build sequence record execution, every phase plan
    records completion, and the top PR still targets `gavin/shape-diff-stable-identity`.
