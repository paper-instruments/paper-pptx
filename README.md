<div align="center">
  <a href="https://github.com/paper-instruments/paper-pptx">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset=".github/assets/logo-dark.svg">
      <img alt="paper-pptx logo" src=".github/assets/logo-light.svg" height="128">
    </picture>
  </a>
  <h1>paper-pptx</h1>

[![PyPI](https://img.shields.io/pypi/v/paper-pptx.svg)](https://pypi.org/project/paper-pptx/)
[![Python versions](https://img.shields.io/pypi/pyversions/paper-pptx.svg)](https://pypi.org/project/paper-pptx/)
[![License](https://img.shields.io/pypi/l/paper-pptx.svg)](LICENSE)
[![CI](https://github.com/paper-instruments/paper-pptx/actions/workflows/test.yml/badge.svg)](https://github.com/paper-instruments/paper-pptx/actions/workflows/test.yml)
[![Downloads](https://img.shields.io/pypi/dm/paper-pptx.svg)](https://pypi.org/project/paper-pptx/)

</div>

An agent-first structure editor for PowerPoint files.

`paper-pptx` is a Python library for safely inspecting, editing, composing, and
verifying existing PowerPoint (`.pptx`) presentations. It is a drop-in hard fork
of [`python-pptx`](https://github.com/scanny/python-pptx) `v1.0.2`: the
distribution is renamed, the import name stays `pptx`, and every existing call
keeps working unchanged.

```python
from pptx import Presentation          # the import name is unchanged
```

The fork exists to prevent **silent corruption**: a deck that opens fine and is
quietly wrong. Automated systems cannot eyeball a slide, so every added
operation returns its outcome as typed, machine-readable data — and when an
operation cannot proceed safely, it raises a typed refusal and leaves the
presentation byte-for-byte unchanged rather than guessing.

## Table of contents

1. [Installation](#installation)
1. [Quick start](#quick-start)
1. [What we changed from python-pptx, and why](#what-we-changed-from-python-pptx-and-why)
1. [The safety contract](#the-safety-contract)
1. [API surface map](#api-surface-map)
1. [Everything from python-pptx still works](#everything-from-python-pptx-still-works)
1. [Documentation](#documentation)
1. [Roadmap and known limitations](#roadmap-and-known-limitations)
1. [How it's tested](#how-its-tested)
1. [Contributing](#contributing)
1. [Community and support](#community-and-support)
1. [Citation](#citation)
1. [License](#license)
1. [Acknowledgments](#acknowledgments)

## Installation

Requires Python 3.9+.

```bash
python -m pip uninstall -y python-pptx paper-pptx
python -m pip install paper-pptx
```

> [!WARNING]
> The clean uninstall is required when migrating from `python-pptx`. Both
> distributions own the same `pptx` import package (the same
> distribution/import split as Pillow: `pip install pillow`, `import PIL`), and
> pip cannot safely overlay or uninstall two distributions that own the same
> files. If both are installed, `import pptx` refuses with an `ImportError`
> rather than running an unverifiable mix of the two.

Confirm the install:

```bash
paper-pptx-doctor
# paper-pptx-doctor: OK (paper-pptx 0.1.2)
```

The doctor verifies that `paper-pptx` metadata is present, `python-pptx` is
absent, and the installed `pptx` files match the distribution's RECORD hashes —
so "I am actually running paper-pptx" is provable, not assumed.

## Quick start

```python
from pptx import Presentation
from pptx.diff import diff_decks
from pptx.edit import replace_text

prs = Presentation("deck.pptx")

run = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
font = run.effective_font()
print(font.size.value_pt)              # resolved through layout/master/theme, e.g. 36.0

replace_text(prs, "FY25", "FY26")      # preserves untouched run formatting
prs.save("deck.v2.pptx")               # atomic: the old file survives any failure

delta = diff_decks("deck.pptx", "deck.v2.pptx", detail="text")
print(len(delta.slide_changes), "slides changed")
```

A fuller composition workflow (import a slide from another deck, apply live
footers, scrub, and verify) is in
[A complete example](#compose-assemble-decks-across-files) below and in
[`docs/user/paper-additions.rst`](docs/user/paper-additions.rst).

## What we changed from python-pptx, and why

`python-pptx` is excellent at *creating* presentations. Its lossless package
layer, disciplined XML mapping, and a decade of absorbed real-world edge cases
are why this fork builds on it rather than starting over.

The harder problem is changing a live, branded, template-driven deck. Upstream's
editing surface stalled short of that: you cannot copy, delete, or reorder a
slide through the public API; you cannot make a real bullet in an arbitrary
text box; you cannot learn what font size a shape actually renders at when the
value is inherited through placeholder → layout → master → theme — the API
returns `None`. Production automation therefore falls back to raw XML and
ZIP-package surgery for exactly the operations template work needs most, and
the dominant failure mode of that surgery is silent corruption: decks that open
in python-pptx but not in PowerPoint, cloned slides that secretly share an
editable chart with their original.

paper-pptx extends the fork into a superset of upstream's public Python API —
every inherited call keeps working, and upstream's own pytest and behave suites
run green on every change — while adding the missing operations as safe,
first-class APIs. The entire fork is one reviewed change on top of upstream
`v1.0.2` (git tag `paper-base`): 180 files changed, 27,546 insertions. The
additions group into four verbs: **perceive**, **edit**, **compose**,
**verify**. A handful of existing behaviors are deliberately stricter; those
are listed honestly in
[What is deliberately not additive](#what-is-deliberately-not-additive).

### Perceive: read what the deck actually renders

**Effective formatting, with provenance.** Stock `run.font.size`, `.name`, and
`.color` return `None` for any value inherited through the placeholder →
layout → master → theme chain — on a branded template, that is nearly
everything, so callers are blind to what the deck looks like. paper-pptx adds
`pptx.inspect.effective_font()`, `effective_paragraph_format()`, and
`effective_shape_format()` (plus `run.effective_font()` as a convenience). Each
value reports whether it resolved, resolves theme colors through the master
color map, and carries the ordered provenance chain of sources consulted.
Values the resolver does not support — gradients, East Asian typefaces, style
modulations — report `resolved=False` instead of a guess.

```python
font = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0].effective_font()
font.size.value_pt        # e.g. 26.0, even though run.font.size is None
for step in font.size.provenance:
    if step.supplied:
        print("supplied by", step.level)   # e.g. "master txStyles titleStyle lvl1"
```

**Visibility-complete text inspection.** Iterating top-level shapes misses text
inside nested groups and table cells, and `shape.text` flattens slide-number
fields and line breaks into plain characters. `pptx.inspect.inspect_text()`
traverses shapes, groups (recursively), and table cells; fields keep their
field type; every block carries a content-hash `BlockAnchor` for stale-safe
edits; and regions the library cannot read (chart text, SmartArt) are counted
as blind regions instead of silently vanishing from "all the text".

**A deterministic deck manifest.** `pptx.inspect.inspect_deck()` emits a
versioned, JSON-friendly structural manifest — slides, shapes, z-order,
geometry, placeholder bindings, layouts, masters — deterministic enough to
golden-file, so an agent can fingerprint structure before and after an edit
without crawling OOXML.

### Edit: change one deck without flattening it

**Anchored, formatting-preserving text replacement.** The only stock write path
(`shape.text = ...`) flattens every run's formatting in the paragraph.
`pptx.edit.replace_text()` replaces matches deck-wide — through groups and
table cells, never across field or line-break boundaries — while untouched runs
keep their exact XML. `replace_text_at()` applies an edit at a `BlockAnchor`
and raises `StaleAnchorError` if the content changed since inspection, and
`refind()` is the explicit recovery path.

```python
from pptx.edit import replace_text, replace_text_at

result = replace_text(prs, "ACME", "NORTHSTAR", include_notes=True)
anchor = result.blocks[0]
replace_text_at(prs, anchor, "NORTHSTAR", "NS")   # StaleAnchorError if the deck moved on
```

**Relationship-safe slide lifecycle.** Upstream has no public clone, delete,
move, or reorder for slides; the folk XML recipes leave sections and custom
shows pointing at deleted slides and produce clones that secretly share an
editable chart and workbook with their original. `Slides.clone()` deep-copies
charts *with their embedded workbooks* and notes under an explicit
`SlideClonePolicy`; `delete()` purges section and custom-show references and
never strands orphaned parts; `reorder()` requires an exact permutation.
Relationship types the library cannot clone safely (OLE objects, ActiveX)
refuse before anything changes.

```python
prs.slides.clone(3)
prs.slides.move(prs.slides[3], 0)
prs.slides.delete(1)
```

**Real bullets, not glyph hacks.** Upstream has no bullet API, so automation
fakes lists by typing `- ` or `•` into the text — which changes the text
itself and is invisible to PowerPoint's list semantics. `paragraph.bullet`
writes genuine `a:buChar` / `a:buAutoNum` / `a:buNone` markup with
hanging-indent geometry.

```python
paragraph.bullet.set_numbered("arabicPeriod", start_at=1)
paragraph.bullet.set_character("•")
paragraph.bullet.set_none()
```

**Autofit made explicit.** PowerPoint stores shrink-to-fit as invisible
`normAutofit` scale percentages; editing text or flipping the autofit flag
silently changes how the deck renders at next open. `TextFrame.font_scale` and
`line_space_reduction` expose the scaling, and `normalize_autofit()` bakes the
currently-rendered sizes into explicit formatting before disabling autofit —
refusing (rather than guessing) when a size cannot be resolved.

**Notes without side effects.** Merely *reading* `slide.notes_slide` upstream
creates a notes part — reconnaissance mutates the package.
`Slide.read_notes_text()` never creates anything, and `replace_notes_text()`
edits only an existing notes body, refusing on slides that have none.

**Typed, group-aware shape lookup.** `shapes.shape_by_name()`,
`picture_by_name()`, `table_by_name()`, and `chart_by_name()` recurse into
groups, return the right object type, and raise `AmbiguousTargetError` on
duplicate names instead of silently taking the first match — the canonical way
an automated edit lands on the wrong object with no error at all.

**Shape and table surgery that keeps the package consistent.**
`SlideShapes.delete()` / `move()` / `add_copy()` manage relationships
correctly (an image relationship is dropped only when nothing else uses it;
copied charts get their own workbook). `Table.insert_row()` / `delete_row()` /
`insert_column()` / `delete_column()` keep the grid and frame geometry
consistent and guard merged regions cell-wise — only an operation that would
actually split a merge refuses.

```python
table.insert_row(after=2, copy_format_from=2)
table.insert_column(after=4)
table.delete_row(0)     # refuses if a vertical merge spans past row 0
```

**Image and chart replacement without collateral damage.**
`Picture.replace_image()` swaps image bytes while leaving position, size,
rotation, and crop untouched — and when two pictures share one image part, only
the targeted picture changes. `Chart.replace_data_safe()` validates the target
chart and workbook structure fully before writing, refuses shared chart parts
and unsupported chart families, and supports workbook-less charts; upstream's
unguarded `replace_data()` remains available.

```python
pic = slide.shapes.picture_by_name("Shared Logo A")
pic.replace_image("new_logo.png")

chart = slide.shapes.chart_by_name("Quarterly Chart")
chart.replace_data_safe(["North", "South"], [("FY27", [10, 20])])
```

### Compose: assemble decks across files

Nobody builds a pitch book from scratch: overview pages come from the master
deck, tombstones from the credentials library, sector pages from the sector
team. The job *is* composition, and composition is relationship-and-inheritance
surgery — the corruption-prone, package-level mechanics this fork exists to
own. A slide's appearance half lives outside the slide, in its layout, master,
and theme; copying slide XML alone produces the failure everyone has seen —
fonts snap to defaults, brand colors revert, charts detach from their data.

**Cross-deck import with explicit fidelity modes.**
`Presentation.import_slide()` and `append_deck()` make the inheritance
trade-off explicit — `mode` is required, with no default: `"adopt_theme"`
rebinds to a destination layout, `"keep_appearance"` transplants the source
layout/master/theme chain (fingerprint-deduplicated, so ten slides from one
source share one master), and `"bake"` freezes resolved formatting into the
slide. Every import returns an `ImportReport`, and the source deck is never
mutated.

**Layout rebind with a mandatory shift report.** `Slide.rebind_layout()` moves
a slide to another layout under explicit placeholder and orphan policies, runs
the effective-value resolver before and after, and reports every run whose
resolved appearance changed — a rebind never shifts appearance silently.

**Real fields, not static text.** `Presentation.apply_footers()` reproduces
PowerPoint's Insert → Header & Footer behavior: genuine `a:fld` slide-number
and date fields that renumber on reorder, bound to the layout's footer
placeholders, applied idempotently.

**A scrub gate before the deck leaves.** Speaker notes, comments, and metadata
leaking in an externally sent deck is a compliance failure, not a cosmetic one.
`Presentation.scrub()` removes selected notes, comments, metadata, hidden
slides, unused layouts/masters, unreachable media, and embedded fonts — by
relationship-graph reachability, so anything a live slide still uses
structurally cannot be removed — and returns a `ScrubReport` listing exactly
which parts were removed or modified.

A complete example:

```python
from pptx import Presentation
from pptx.diff import diff_decks

prs = Presentation("house_deck.pptx")
source = Presentation("sector_team_deck.pptx")

report = prs.import_slide(source, 0, mode="adopt_theme")
for shift in report.run_shifts:
    print(shift.text, shift.before["name"]["value"], "->", shift.after["name"]["value"])

prs.apply_footers(footer="Confidential", slide_number=True)
prs.scrub(metadata=True, comments=True)
prs.save("house_deck.v2.pptx")

delta = diff_decks("house_deck.pptx", "house_deck.v2.pptx", detail="text")
print("slides added:", [s.slide_id for s in delta.slides_added])
```

### Verify: prove what changed

A deck is amnesiac: the format carries no revision markup, so "what changed
between v3 and v4" has no programmable answer anywhere in the ecosystem —
until you diff it yourself.

**A semantic deck diff.** `pptx.diff.diff_decks()` matches slides by permanent
slide ID, so a reorder reports as a *move* rather than a delete-plus-add, and
reports shape, text, chart-data, image, and notes changes within matched
slides. `detail="full"` adds per-run effective-value shifts, which lets a
pipeline compare an operation's report against what actually changed in the
saved file.

**Byte-minimal saves and a package-level oracle.** A stock save rewrites every
ZIP member, so even a no-op looks like total churn. `pptx.package.patch_save()`
writes semantically-unchanged parts back with their original bytes — a
one-line edit to a sixty-slide deck diffs as one part, not sixty — and
`diff_package()` reports exactly which parts differ. Meaningful whitespace,
including a trailing space inside a run, is treated as content.

```python
from pptx.package import patch_save

residual = patch_save("deck.pptx", prs, "out.pptx")   # PackageDiff of real changes only
```

### What is deliberately not additive

paper-pptx's public Python API is a superset of upstream's, but a few narrow
existing behaviors changed on purpose. Each trades edge-case permissiveness for
corruption prevention:

- **Guarded package intake.** Opening a `.pptx` now rejects ambiguous or unsafe
  ZIP archives — duplicate or case-colliding member names, path traversal,
  encrypted or exotically-compressed members, lying size headers,
  resource-exhaustion bombs — with a typed `PackageLimitError` *before* any
  editable object exists. A permissive reader "successfully" opens an ambiguous
  archive and then faithfully edits the wrong interpretation of it. Some
  odd-but-consumer-readable files that upstream accepted are now refused.
- **Atomic save.** `save()` keeps its signature, but saving to a path now
  writes a sibling temporary file and atomically replaces the destination only
  after serialization succeeds, preserving the existing file's permission bits;
  stream saves snapshot and restore the destination on failure. Upstream wrote
  directly to the destination, so a mid-save failure destroyed the only copy.
- **`SlideLayouts.remove()` hardening.** Same signature, stricter semantics:
  stale or foreign proxies and unsafe states now refuse atomically instead of
  partially mutating.
- **Python 3.9+ floor.** Upstream `v1.0.2` supported Python 3.8.
- **Distribution identity.** The distribution is renamed `paper-pptx`, and
  `import pptx` fails loudly when both `python-pptx` and `paper-pptx` metadata
  are installed, because a mixed site-packages can silently run a blend of the
  two libraries. `pptx.__version__` stays `"1.0.2"` (the upstream API surface),
  and `pptx.__paper_version__` (`"0.1.2"`) identifies the fork release — so
  callers can distinguish "which upstream surface" from "which paper release".

## The safety contract

Every added operation either does exactly what it claims or refuses
atomically. Mutating operations validate fully before they change anything —
never mutate-then-validate. When an operation cannot proceed safely, it raises
a typed refusal from the hierarchy rooted at `pptx.errors.PaperRefusal`
(`PackageLimitError`, `TargetNotFoundError`, `StaleAnchorError`,
`AmbiguousTargetError`, `UnsupportedStructureError`, `RelationshipPolicyError`,
`BoundaryViolationError`) and leaves the presentation byte-for-byte unchanged
in memory and on disk. Programmer mistakes — a bad type, an out-of-range
index — remain plain `ValueError` or `TypeError`, so callers can catch
`PaperRefusal` separately:

```python
from pptx import Presentation
from pptx.errors import PaperRefusal

prs = Presentation("deck.pptx")
try:
    prs.slides.clone(3)                 # slide contains an embedded OLE object
except PaperRefusal:
    ...                                 # document is untouched; handle or report
```

A refused edit is a success mode; a quietly wrong file is the worst outcome
this library can produce. Held proxies survive a refusal too: after a refused
multi-part operation, existing slide and shape objects still point at valid,
unchanged content, and stale handles (a slide proxy held across that slide's
deletion) raise `TargetNotFoundError` instead of silently editing a neighbor.

Refusal atomicity is enforced operation by operation: each documented refusal
condition has a test asserting both that the typed refusal is raised and that
output bytes equal input bytes. It is a per-operation contract proven by the
test harness, not a blanket guarantee over every code path.

## API surface map

New modules, plus methods added to existing classes. Nothing is re-exported
from `pptx` itself; import from the module named below.

| Module | What it does | Reference |
|---|---|---|
| `pptx.inspect` | Effective (rendered) values with provenance; text inspection with content-hash anchors; deck manifest | [docs](docs/api/inspect.rst) |
| `pptx.edit` | Deck-wide and anchored text replacement that preserves run formatting | [docs](docs/api/edit.rst) |
| `pptx.diff` | Semantic deck-to-deck diff (`diff_decks`) | [docs](docs/api/diff.rst) |
| `pptx.compose` | Cross-deck slide import and deck append (via `Presentation.import_slide` / `append_deck`) | [docs](docs/api/compose.rst) |
| `pptx.rebind` | Layout rebinding with shift reports (via `Slide.rebind_layout`) | [docs](docs/api/rebind.rst) |
| `pptx.scrub` | Reported removal of notes, comments, metadata, unused parts (via `Presentation.scrub`) | [docs](docs/api/scrub.rst) |
| `pptx.hf` | Real slide-number/date/footer fields (via `apply_footers`) | [docs](docs/api/hf.rst) |
| `pptx.package` | Semantic package diff and byte-minimal `patch_save` | [docs](docs/api/package.rst) |
| `pptx.errors` | The `PaperRefusal` typed-refusal hierarchy | [docs](docs/api/errors.rst) |

Methods added to inherited classes: `Slides.clone` / `delete` / `reorder` /
`move`; `SlideShapes.delete` / `move` / `add_copy` and the `*_by_name`
lookups; `Table.insert_row` / `delete_row` / `insert_column` /
`delete_column`; `Picture.replace_image`; `Chart.replace_data_safe`;
`TextFrame.normalize_autofit`, `font_scale`, `line_space_reduction`;
`_Paragraph.bullet` and field helpers; `Slide.read_notes_text` /
`replace_notes_text` / `apply_footers` / `rebind_layout`.

## Everything from python-pptx still works

The import surface is unchanged: `from pptx import Presentation`, `pptx.util`,
`pptx.chart.data`, placeholder access, shape trees — every documented upstream
API behaves as upstream documents it, and upstream's own pytest and behave
suites run on every change to keep it that way. Existing code, snippets, and
model priors that say `import pptx` work as-is; `pptx.__version__` still
reports the upstream base (`"1.0.2"`).

New upstream releases are merged, never rebased, so the fork retains its
history and compatibility.

## Documentation

There is no hosted documentation site yet. The Sphinx docs build in CI and
extend the upstream python-pptx documentation:

- [`docs/user/paper-additions.rst`](docs/user/paper-additions.rst) — the
  narrative guide to everything the fork adds (perceive / edit / compose /
  verify), with worked examples.
- [`docs/api/`](docs/api/) — reference pages for each added module.
- The remaining documentation is inherited from python-pptx and describes the
  shared, unchanged foundation.

Build locally with `uv sync --group docs` and `uv run make docs`.

## Roadmap and known limitations

A great many refusals are deliberate scope statements, and converting a
documented typed refusal into a correct operation is the sanctioned growth
path of this package. Known gaps, honestly:

- **Notes parts are never created.** `read_notes_text` / `replace_notes_text`
  work only on existing notes; creating the notes-part graph (notes master
  included) is a recorded candidate, not yet built.
- **Table-cell effective values refuse** until the table-style resolution walk
  lands.
- **Chart data replacement supports single-plot category charts.** XY, bubble,
  stock, surface, radar, 3-D, and multi-plot combos refuse rather than risk
  cache/workbook desynchronization.
- **Fixture provenance.** The test corpus is generated and
  LibreOffice-round-tripped, with provenance recorded truthfully;
  PowerPoint-authored and Google-exported fixtures are still pending, and
  release sign-off includes a manual desktop-PowerPoint checklist.
- **Deck diff assumes lineage.** `diff_decks` matches slides by permanent ID,
  which serves decks derived from a common ancestor (v4 saved from v3);
  matching independently-authored decks is out of scope today.

Deliberate non-goals: no rendering or layout-geometry computation (appearance
verification belongs to a harness, not this library), no SmartArt authoring
(opaque preservation only), no animations/transitions editing, no aspect-ratio
migration, no bulk template-migration orchestration (`rebind_layout` is the
primitive), no document-QA `check()` API, and no new runtime dependencies.

## How it's tested

- Upstream's pytest and behave suites run on every change to check
  compatibility with existing behavior.
- A frozen, hash-pinned fixture corpus includes generated presentations and
  LibreOffice round-trips with provenance sidecars.
- The contract harness in [`tests/paper/`](tests/paper/) saves and reopens
  before asserting, enforces exact changed-part budgets and refusal atomicity
  (output bytes equal input bytes on every documented refusal), and validates
  selected XML fragments against their schemas.
- Release verification requires a headless LibreOffice load smoke
  (`pytest -m lo_smoke tests/paper`) and the manual checklist in
  [`tests/paper/RELEASE-CHECKLIST.md`](tests/paper/RELEASE-CHECKLIST.md).
- House rule: no fix without a fixture.

## Contributing

Contributions are welcome — see [`CONTRIBUTING.md`](CONTRIBUTING.md) for dev
setup, the test matrix, and the safety conventions every mutating API must
satisfy (validate-fully-then-mutate, save→reopen assertions, changed-part
budgets). The short version:

```bash
uv sync --group dev
uv run pytest
uv run behave
```

## Community and support

- Bugs and feature requests: [GitHub Issues](https://github.com/paper-instruments/paper-pptx/issues)
- Release history: [GitHub releases](https://github.com/paper-instruments/paper-pptx/releases)

## Citation

If you reference paper-pptx in research or writing:

```bibtex
@software{paper_pptx,
  title   = {paper-pptx: an agent-first structure editor for PowerPoint files},
  author  = {{Paper Instruments, Inc.}},
  year    = {2026},
  version = {0.1.2},
  url     = {https://github.com/paper-instruments/paper-pptx}
}
```

## License

MIT, inherited from python-pptx. Original work © 2013 Steve Canny and the
python-pptx contributors; fork additions © 2026 Paper Instruments, Inc. This
fork preserves the upstream license and attribution. See
[`LICENSE`](LICENSE).

## Acknowledgments

paper-pptx exists because [python-pptx](https://github.com/scanny/python-pptx)
is excellent. Steve Canny and the python-pptx contributors built the lossless
package layer, the disciplined XML mapping, and a decade of absorbed edge
cases that make safe deck editing possible at all — this fork stands on that
work and gratefully keeps their API intact.
