<div align="center">
  <a href="https://github.com/paper-instruments/paper-pptx">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/paper-instruments/paper-pptx/main/.github/assets/logo-dark.svg">
      <img alt="paper-pptx logo" src="https://raw.githubusercontent.com/paper-instruments/paper-pptx/main/.github/assets/logo-light.svg" height="128">
    </picture>
  </a>
  <h1>paper-pptx</h1>

[![PyPI](https://img.shields.io/pypi/v/paper-pptx.svg)](https://pypi.org/project/paper-pptx/)
[![Python versions](https://img.shields.io/pypi/pyversions/paper-pptx.svg)](https://pypi.org/project/paper-pptx/)
[![Test](https://github.com/paper-instruments/paper-pptx/actions/workflows/test.yml/badge.svg)](https://github.com/paper-instruments/paper-pptx/actions/workflows/test.yml)

</div>

**An import-compatible, agent-first structure editor for PowerPoint files, designed to prevent silent corruption when editing existing decks.**

`paper-pptx` is a drop-in hard fork of [python-pptx](https://github.com/scanny/python-pptx) `v1.0.2` for safely inspecting, editing, composing, and verifying existing PowerPoint (`.pptx`) presentations. It keeps python-pptx's package layer, XML mapping, and object model. It adds the rendered values a deck shows, edits that survive PowerPoint's run fragmentation, and refusals in place of guesses.

```python
from pptx import Presentation   # the import name is unchanged (see "Drop-in by design")
```

The fork exists to prevent **silent corruption**: a deck that opens without error but is wrong. Automated systems cannot eyeball a slide, so every added operation returns its outcome as typed, machine-readable data, and an operation that cannot proceed safely raises a typed refusal and leaves the presentation byte-for-byte unchanged.

---

## Why paper-pptx exists

`python-pptx` is excellent at *creating* decks. Its lossless package layer, disciplined XML mapping, and a decade of absorbed edge cases are why this fork builds on it.

The harder problem is changing an existing deck without flattening formatting, stranding relationships, or losing content outside the slide body. Stock `run.font.size` returns `None` for any value inherited through the placeholder chain, the only stock write path flattens every run in a paragraph, and there is no public clone, delete, or reorder for slides. An agent cannot look at the result, so it needs the deck's structure and every edit outcome as typed data, and it needs the library to refuse unsafe edits.

## Quick start

```python
from pptx import Presentation
from pptx.diff import diff_decks
from pptx.edit import replace_text

prs = Presentation("deck.pptx")

run = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
font = run.effective_font()
print(font.size.value_pt)              # resolved through layout/master/theme

replace_text(prs, "FY25", "FY26")      # preserves untouched run formatting
prs.save("deck.v2.pptx")               # atomic on a path: the old file survives any failure

delta = diff_decks("deck.pptx", "deck.v2.pptx", detail="text")
print(len(delta.slide_changes), "slides changed")
```

## What paper-pptx adds

### Perceive: read what the deck renders

- **`pptx.inspect` effective values, with provenance.** `effective_font()`, `effective_paragraph_format()`, and `effective_shape_format()` resolve size, typeface, color, alignment, line spacing, and bullets through the run, paragraph, placeholder, layout, master, and theme chain, and report which rung supplied each value. Bullet typeface and size resolve on their own chains, because the schema inherits them separately from the glyph. Unresolved values are reported as unresolved.
- **Visibility-complete text inspection.** `inspect_text()` reaches nested groups and table cells, which iterating top-level shapes misses, and returns content-hash anchors that survive an edit. Regions it cannot survey are reported as blind blocks rather than skipped.
- **A deterministic deck manifest.** `inspect_deck()` emits a versioned, JSON-friendly structural manifest: slides, shapes, z-order, geometry, and placeholder identity.

### Edit: change one deck without flattening it

- **Anchored, formatting-preserving replacement.** `pptx.edit.replace_text` and `replace_text_at` rewrite text while leaving untouched runs byte-identical. The stock `shape.text = ...` flattens every run's formatting in the paragraph.
- **Relationship-safe slide lifecycle.** `slides.clone/delete/move/reorder` maintain sections, custom shows, and relationships. Upstream has no public equivalent, and the folk XML recipes strand both.
- **Real bullets.** `paragraph.bullet` authors genuine `a:buChar` and `a:buAutoNum` state, including bullet typeface and size, instead of typing `-` or `•` into the text.
- **Autofit made explicit.** `text_frame.normalize_autofit()` freezes PowerPoint's invisible `normAutofit` scale percentages so an edit does not resize text without warning.
- **Notes without side effects.** `slide.read_notes_text()` never creates a notes part; reading `slide.notes_slide` upstream does.
- **Typed, group-aware lookup.** `shape_by_name()`, `picture_by_name()`, `table_by_name()`, and `chart_by_name()` recurse into groups and refuse a duplicate rather than returning the first match.
- **Surgery that keeps the package consistent.** `SlideShapes.delete/move/add_copy`, table row and column insert and delete, `Picture.replace_image()`, and `Chart.replace_data_safe()` maintain relationships and owned parts.
- **Batched validation.** `with prs.batch():` validates once at block exit instead of once per mutating call, and discards every edit in the block if that check fails.
- **`SlideLayouts.remove()` hardening.** Same signature, stricter semantics: stale or foreign proxies and unsafe states refuse atomically instead of partially mutating.

### Compose: assemble decks across files

- **Cross-deck import with explicit fidelity modes.** `Presentation.import_slide()` and `append_deck()` make the inheritance trade-off a required argument: `adopt_theme`, `keep_appearance`, or `bake`. Each returns an `ImportReport` naming every part added, reused, and dropped.
- **Layout rebind with a shift report.** `Slide.rebind_layout()` moves a slide under explicit placeholder and orphan policies and reports every run whose resolved font changed. Placeholder geometry and text direction are inherited from the layout and sit outside that comparison.
- **Real fields, not static text.** `apply_footers()` writes genuine `a:fld` slide-number and date fields, so PowerPoint refreshes them.

### Verify: prove what changed

- **A semantic deck diff.** `pptx.diff.diff_decks()` matches slides by permanent slide ID, so a reorder reports as a move rather than a delete plus an add. `detail="text"` adds chart data, text, and notes; `detail="full"` adds per-run and bullet shifts.
- **Byte-minimal saves and a package oracle.** `pptx.package.patch_save()` writes semantically unchanged parts back with their original bytes, so a one-line edit diffs as a few parts rather than all of them, and `diff_package()` reports exactly which parts differ.

### Package intake and save

- **Guarded package intake.** Opening a `.pptx` rejects ambiguous or unsafe ZIP archives: duplicate or case-colliding member names, path traversal, encrypted or exotically-compressed members, lying size headers, an archive that does not span its file exactly, and any member that resolves to no content type. A part that nothing references is kept; PowerPoint opens that package and drops the part on its next save, and so does `save()`.
- **Atomic save.** Saving to a path writes a sibling temporary file and replaces the destination only after serialization succeeds, preserving the existing file's permission bits and resolving symlinks. Stream saves stage the whole package first, so a serialization failure emits nothing, and restore the destination on failure when it can be read and rewound; a write-only sink keeps whatever landed.

## Safety contract

Every added operation either does exactly what it claims or refuses atomically. Mutating operations run inside a package transaction: whatever the operation touched is restored if it refuses, and the deck-wide check runs before anything commits. Some operations, including `apply_footers()`, `append_deck()`, `import_slide()` and slide clone, also validate in full before touching anything.

A refusal raises a typed error from the hierarchy rooted at `pptx.errors.PaperRefusal` and leaves the presentation byte-for-byte unchanged in memory and on disk. The subclasses say what went wrong: `PackageLimitError`, `TargetNotFoundError`, `StaleAnchorError`, `AmbiguousTargetError`, `UnsupportedStructureError`, `RelationshipPolicyError`, and `BoundaryViolationError`. Programmer mistakes remain plain `ValueError` or `TypeError`, so callers can catch `PaperRefusal` separately.

A refused edit is a success mode; the worst outcome this library can produce is a file that opens without error but is wrong. Held proxies survive a refusal, and a stale handle raises `TargetNotFoundError` instead of editing a neighbor. Each documented refusal condition has a test asserting both that the typed refusal is raised and that output bytes equal input bytes.

## Drop-in by design

Only the distribution and repository are renamed. The importable package stays `pptx`. This is the same distribution/import split as Pillow (`pip install pillow`, `import PIL`), and it preserves existing code, snippets, and model priors.

- GitHub repository / PyPI distribution: **`paper-pptx`**
- Python import: **`pptx`**
- Fork sentinel: `pptx.__paper_version__`

Every documented upstream API behaves as upstream documents it: `from pptx import Presentation`, `pptx.util`, `pptx.chart.data`, placeholder access, shape trees. New upstream releases are merged, never rebased, so the fork retains its history and compatibility.

## Installation

Requires Python 3.9+. Upstream `v1.0.2` supported Python 3.8.

```bash
python -m pip uninstall -y python-pptx paper-pptx
python -m pip install paper-pptx
```

The clean uninstall is required when migrating from `python-pptx`. Both distributions own the same `pptx` import package, and pip cannot safely overlay or uninstall two distributions that own the same files. If both are installed, `import pptx` refuses with an `ImportError` rather than running an unverifiable mix of the two.

Confirm the install:

```bash
paper-pptx-doctor
```

## Documentation

The Sphinx docs extend the upstream python-pptx documentation to cover the fork's additions: start with `docs/user/paper-additions.rst` and the `docs/api/*.rst` reference pages. Everything inherited from python-pptx works as documented at the [python-pptx documentation](https://python-pptx.readthedocs.io/).

## Current limitations

Converting a documented typed refusal into a correct operation is the sanctioned growth path of this package. The known gaps:

- **Notes parts are never created.** `read_notes_text` and `replace_notes_text` work only on existing notes.
- **Table-cell effective values refuse** until the table-style resolution walk is built.
- **Chart data replacement supports single-plot category charts.** XY, bubble, stock, surface, radar, 3-D, and multi-plot combos refuse.
- **Bullet color is not resolved**, and bullet diffing needs `detail="full"` and skips table cells.
- **`diff_decks` assumes lineage**, matching slides by permanent ID. It is not a visual diff.
- **Fixture provenance.** The test corpus is generated and LibreOffice round-tripped, not authored by PowerPoint.

Deliberate non-goals: no rendering or layout-geometry computation, no SmartArt authoring (opaque preservation only), and no animation or transition authoring.

## Testing

- Upstream's pytest and behave suites run on every change to check compatibility with existing behavior.
- A frozen, hash-pinned fixture corpus includes generated presentations and LibreOffice round-trips.
- The contract harness checks refusal atomicity and validates the fixture corpus with a headless LibreOffice load smoke.

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](https://github.com/paper-instruments/paper-pptx/blob/main/CONTRIBUTING.md) for the engineering discipline this fork runs on. The short version: the upstream suite must remain green; persistence changes need saved-and-reopened assertions and exact package-delta checks; guarded refusals must be atomic and leave bytes unchanged; and a refusal message must name what was found, why it is unsafe, and what to do about it.

## Community

- **Bugs and feature requests**: [GitHub Issues](https://github.com/paper-instruments/paper-pptx/issues)
- **Questions and ideas**: [GitHub Discussions](https://github.com/paper-instruments/paper-pptx/discussions)

## Citation

If you reference paper-pptx in research or writing:

```bibtex
@software{paper_pptx,
  title   = {paper-pptx: an agent-first structure editor for PowerPoint files},
  author  = {{Paper Instruments, Inc.}},
  year    = {2026},
  url     = {https://github.com/paper-instruments/paper-pptx}
}
```

## Acknowledgments

paper-pptx exists because [python-pptx](https://github.com/scanny/python-pptx) is excellent. Steve Canny and the python-pptx contributors built the lossless package layer, the disciplined XML mapping, and a decade of absorbed edge cases that make safe deck editing possible at all. This fork stands on that work and keeps their API intact.

## License

MIT, inherited from python-pptx. Original work © 2013 Steve Canny and the python-pptx contributors; fork additions © 2026 Paper Instruments, Inc. This fork preserves the upstream license and attribution. See [LICENSE](https://github.com/paper-instruments/paper-pptx/blob/main/LICENSE).
