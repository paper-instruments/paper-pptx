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

**An import-compatible, agent-safe fork of python-pptx for creating and editing PowerPoint presentations without silent corruption.**

`paper-pptx` is an import-compatible hard fork of [python-pptx](https://github.com/scanny/python-pptx) for creating and editing PowerPoint (`.pptx`) presentations. It keeps the `pptx` import name and upstream object model, and adds structure-aware operations and machine-readable results for automated workflows.

## Installation

```bash
python -m pip uninstall -y python-pptx paper-pptx
python -m pip install paper-pptx
paper-pptx-doctor
```

Both distributions provide the `pptx` import package. Do not install `python-pptx` and `paper-pptx` in the same environment.

## Quick start

Create a presentation and update its text without flattening untouched run formatting:

```python
from pptx import Presentation
from pptx.edit import replace_text

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "FY25 plan"
prs.save("deck.pptx")

prs = Presentation("deck.pptx")
result = replace_text(prs, "FY25", "FY26")
prs.save("deck-v2.pptx")
print(result.replacements)  # 1
```

`replace_text` returns a machine-readable result and refuses unsupported edits before changing the presentation.

## Documentation

Read the [paper-pptx documentation](https://docs.paperinstruments.com/).

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](https://github.com/paper-instruments/paper-pptx/blob/main/CONTRIBUTING.md).

## Acknowledgments

paper-pptx builds on python-pptx by Steve Canny and contributors. This fork preserves their API, license, and attribution.

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

Cite it as a fork of *python-pptx* by Steve Canny and contributors.

## License

MIT, inherited from python-pptx. Original work © 2013 Steve Canny and the python-pptx contributors; fork additions © 2026 Paper Instruments, Inc. See [LICENSE](https://github.com/paper-instruments/paper-pptx/blob/main/LICENSE).
