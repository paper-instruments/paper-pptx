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

**An import-compatible, agent-first structure editor that helps prevent silent corruption in existing PowerPoint decks.**

`paper-pptx` is a hard fork of [python-pptx](https://github.com/scanny/python-pptx) `v1.0.2` for programs that inspect and change existing PowerPoint (`.pptx`) presentations. It keeps the `pptx` import name and the upstream object model. The fork adds structure-aware operations and machine-readable results for existing decks.

Automated editors cannot inspect the rendered result of each change. A deck can open without errors after an edit has flattened formatting or damaged relationships. `paper-pptx` exposes the deck's effective values and package structure. Its added APIs raise typed refusals when they cannot select one target or preserve package consistency.

## Installation

`paper-pptx` requires Python 3.9 or later.

```bash
python -m pip uninstall -y python-pptx paper-pptx
python -m pip install paper-pptx
```

Uninstall `python-pptx` before installing this fork because both distributions provide the `pptx` import package. Verify the installation with:

```bash
paper-pptx-doctor
```

## Documentation

The [Sphinx documentation](https://github.com/paper-instruments/paper-pptx/tree/main/docs) covers the fork APIs. The [python-pptx documentation](https://python-pptx.readthedocs.io/) covers the inherited API.

## Contributing

See [CONTRIBUTING.md](https://github.com/paper-instruments/paper-pptx/blob/main/CONTRIBUTING.md).

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

## License and attribution

Steve Canny and the python-pptx contributors created the package and XML foundation. Paper Instruments maintains this fork under the inherited MIT license. See [LICENSE](https://github.com/paper-instruments/paper-pptx/blob/main/LICENSE).
