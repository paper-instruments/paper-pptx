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

Inspect, edit, compose, and compare PowerPoint files with Python.

`paper-pptx` is a fork of
[`python-pptx`](https://github.com/scanny/python-pptx) `v1.0.2`. The PyPI
package has a new name, but the Python import stays `pptx`.

```python
from pptx import Presentation
```

Existing `python-pptx` code can use this fork without import changes. The fork
also adds APIs for common edits to existing presentations.

## Installation

`paper-pptx` requires Python 3.9 or later.

Remove both packages before you install `paper-pptx`:

```bash
python -m pip uninstall -y python-pptx paper-pptx
python -m pip install paper-pptx
```

> [!WARNING]
> Do not install `python-pptx` and `paper-pptx` together. Both packages use the
> same `pptx` import directory. `import pptx` raises `ImportError` if it finds
> metadata for both packages.

Check the installation:

```bash
paper-pptx-doctor
# paper-pptx-doctor: OK (paper-pptx 0.1.2)
```

The command checks the package metadata and installed file hashes. It also
checks that `python-pptx` is not installed.

## Quick start

```python
from pptx import Presentation
from pptx.diff import diff_decks
from pptx.edit import replace_text

prs = Presentation("deck.pptx")

run = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
font = run.effective_font()
print(font.size.value_pt)

replace_text(prs, "FY25", "FY26")
prs.save("deck.v2.pptx")

delta = diff_decks("deck.pptx", "deck.v2.pptx", detail="text")
print(len(delta.slide_changes), "slides changed")
```

## What we changed, and why

`python-pptx` provides a stable package and XML model. However, its public API
does not support some common edits to existing presentations.

Some workflows therefore edit OOXML files and package relationships directly.
These edits can cause silent corruption: the file opens, but its structure or
content is wrong. For example, a copied slide can share an editable chart with
the original.

This fork adds validated APIs for those operations. It reports changes and
refuses operations that it cannot complete safely.

"New" means that this fork adds the API. "Edited" means that this fork changes
an existing `python-pptx` API or behavior.

| Function or API | Status | Description |
|---|---|---|
| `pptx.inspect.effective_font()` and `run.effective_font()` | New | Get the font values that PowerPoint will use and report the source of each value. |
| `pptx.inspect.effective_paragraph_format()` | New | Gets inherited paragraph values and reports the source of each value. |
| `pptx.inspect.effective_shape_format()` | New | Gets inherited shape values and reports the source of each value. |
| `pptx.inspect.inspect_text()` | New | Finds text in shapes, groups, and table cells. It preserves fields and reports unreadable regions. |
| `pptx.inspect.inspect_deck()` | New | Returns a versioned structural record of the presentation. |
| `pptx.edit.replace_text()` | New | Replaces text across a presentation without changing unrelated run formatting. |
| `pptx.edit.replace_text_at()` and `refind()` | New | Apply an anchored edit or find a new anchor when the old content has changed. |
| `Slides.clone()` | New | Copies a slide and gives copied charts and workbooks their own package parts. |
| `Slides.delete()` | New | Deletes a slide and removes its references from sections and custom shows. |
| `Slides.reorder()` and `Slides.move()` | New | Change slide order and require valid slide references. |
| `SlideShapes.delete()`, `move()`, and `add_copy()` | New | Edit shapes and update their package relationships. |
| `shape_by_name()`, `picture_by_name()`, `table_by_name()`, and `chart_by_name()` | New | Find shapes inside groups and reject names that match more than one shape. |
| `Table.insert_row()` and `delete_row()` | New | Change table rows while preserving the grid and valid merged cells. |
| `Table.insert_column()` and `delete_column()` | New | Change table columns while preserving the grid and valid merged cells. |
| `paragraph.bullet` | New | Creates PowerPoint bullets and numbered lists instead of adding bullet characters to text. |
| `TextFrame.normalize_autofit()` | New | Stores the current fitted font sizes before it turns off automatic fitting. |
| `TextFrame.font_scale` and `line_space_reduction` | New | Expose the scale values that PowerPoint stores for automatic fitting. |
| `Slide.read_notes_text()` | New | Reads notes without creating a notes part. |
| `Slide.replace_notes_text()` | New | Changes text in an existing notes part. |
| `Picture.replace_image()` | New | Replaces one picture without changing another picture that shares the old image. It preserves the picture geometry and crop. |
| `Chart.replace_data_safe()` | New | Validates a chart and its workbook before it replaces data. It refuses shared or unsupported chart structures. |
| `Presentation.import_slide()` and `append_deck()` | New | Require `adopt_theme`, `keep_appearance`, or `bake` to control theme and formatting changes. |
| `Slide.rebind_layout()` | New | Applies a different layout and reports changes to resolved text formatting. |
| `Presentation.apply_footers()` and `Slide.apply_footers()` | New | Add PowerPoint fields for footers, dates, and slide numbers. |
| `Presentation.scrub()` | New | Removes selected private data and unused parts. It preserves parts that a live slide uses. |
| `pptx.diff.diff_decks()` | New | Matches slides by permanent ID and reports moves and content changes. |
| `pptx.package.patch_save()` and `diff_package()` | New | Preserve the original bytes for unchanged parts and report parts that changed. |
| `pptx.errors.PaperRefusal` | New | Identifies edits that the library cannot complete without risking an invalid result. |
| `Presentation()` package checks | Edited | Rejects unsafe or ambiguous ZIP members before it creates editable objects. |
| `Presentation.save()` | Edited | Writes to a temporary file and replaces the destination only after a successful save. |
| `SlideLayouts.remove()` | Edited | Rejects stale, foreign, or unsafe layout references before it changes the presentation. |
| `import pptx` | Edited | Rejects an environment that contains metadata for both distributions. |

See [`docs/user/paper-additions.rst`](docs/user/paper-additions.rst) for
examples and detailed behavior.

## Safety contract

Each added edit validates its inputs before it changes the presentation. A
documented refusal leaves the presentation unchanged in memory and on disk.

The library raises subclasses of `pptx.errors.PaperRefusal` for these
refusals. Examples include `StaleAnchorError`, `AmbiguousTargetError`, and
`UnsupportedStructureError`.

```python
from pptx import Presentation
from pptx.errors import PaperRefusal

prs = Presentation("deck.pptx")

try:
    prs.slides.clone(3)
except PaperRefusal as error:
    print(error)
```

Invalid Python types and values still raise `TypeError` or `ValueError`. The
safety contract applies to documented refusal conditions for each operation.

## Compatibility

The public Python API is a superset of the `python-pptx` `v1.0.2` API. The
project runs the upstream pytest and behave suites for each change.

The fork has these intentional differences:

- It requires Python 3.9 or later.
- It rejects unsafe or ambiguous ZIP packages.
- It saves files atomically.
- It applies stricter checks when it removes slide layouts.
- It rejects an installation that also contains `python-pptx` metadata.

`pptx.__version__` reports the upstream API version. `pptx.__paper_version__`
reports the fork version.

The project merges new upstream releases. It does not rebase its fork commits.

## Documentation

The Sphinx documentation builds in CI. There is no hosted documentation site
at this time.

- [`docs/user/paper-additions.rst`](docs/user/paper-additions.rst) explains the
  APIs that this fork adds.
- [`docs/api/`](docs/api/) contains reference pages for the new modules.
- The other pages come from `python-pptx` and describe the shared API.

Build the documentation locally:

```bash
uv sync --group docs
uv run make docs
```

## Known limits

- The notes APIs do not create notes parts. They only read or edit existing
  notes parts.
- The effective-format APIs do not resolve table styles.
- Safe chart replacement supports single-plot category charts.
- The fixture set does not yet include files from PowerPoint or Google Slides.
- `diff_decks()` compares presentations that have a common source. It does not
  match slides from unrelated presentations.
- The library does not render slides or calculate layout geometry.
- The library does not edit SmartArt, animations, or transitions.

Unsupported operations raise a typed refusal when the library can detect them.

## Testing

The project uses these checks:

- The upstream pytest and behave suites check API compatibility.
- The tests save and reopen changed presentations.
- The tests compare changed package parts with expected part lists.
- Refusal tests confirm that output bytes equal input bytes.
- Schema tests validate selected XML fragments.
- LibreOffice tests confirm that saved presentations can open.

The fixture sources and hashes are in
[`tests/paper/fixtures/`](tests/paper/fixtures/). The release steps are in
[`tests/paper/RELEASE-CHECKLIST.md`](tests/paper/RELEASE-CHECKLIST.md).

## Contributing

See [`CONTRIBUTING.md`](CONTRIBUTING.md) for the development setup and test
requirements.

```bash
uv sync --group dev
uv run pytest
uv run behave
```

## Support

- Report bugs and request features in [GitHub Issues](https://github.com/paper-instruments/paper-pptx/issues).
- See published versions in [GitHub Releases](https://github.com/paper-instruments/paper-pptx/releases).

## Citation

```bibtex
@software{paper_pptx,
  title   = {paper-pptx: a structure editor for PowerPoint files},
  author  = {{Paper Instruments, Inc.}},
  year    = {2026},
  version = {0.1.2},
  url     = {https://github.com/paper-instruments/paper-pptx}
}
```

## License

This project uses the MIT License. See [`LICENSE`](LICENSE).

Original work © 2013 Steve Canny and the `python-pptx` contributors. Fork
changes © 2026 Paper Instruments, Inc.

## Acknowledgments

This project uses the package model and public API from
[`python-pptx`](https://github.com/scanny/python-pptx). Thanks to Steve Canny
and the `python-pptx` contributors for their work.
