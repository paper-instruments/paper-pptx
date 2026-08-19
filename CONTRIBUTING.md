# Contributing to paper-pptx

Thanks for your interest in improving paper-pptx. This project is a hard fork
of [python-pptx](https://github.com/scanny/python-pptx) with a strict
compatibility discipline, so contributions are judged on two axes: does the
change do what it claims, and does everything that worked before still work.

## Development setup

Requires Python 3.9+ and a POSIX-ish environment (macOS or Linux; CI runs
Ubuntu).

```bash
git clone https://github.com/paper-instruments/paper-pptx.git
cd paper-pptx
uv sync --group dev
```

Dev, test, and docs dependencies are declared as PEP 735 dependency groups in
`pyproject.toml` and pinned by the committed `uv.lock`. If you prefer pip
(25.1+), `pip install -e . --group dev` works equally well.

## Running the tests

CI (`.github/workflows/test.yml`) runs all of the following; running them
locally before opening a PR saves a round trip:

```bash
uv run pytest                          # upstream unit suite + tests/paper contract harness
uv run behave                          # upstream acceptance features (or: uv run make accept)
uv run pytest -m lo_smoke tests/paper  # LibreOffice load smoke; needs headless LibreOffice
uv run make docs                       # Sphinx build; CI builds with warnings as errors
uv run make build                      # sdist/wheel build + twine check
```

The pytest matrix in CI covers Python 3.9 through 3.13.

Useful Makefile targets: `accept`, `build`, `clean`, `cleandocs`, `coverage`,
`docs`, `opendocs`.

## The rules that get PRs merged

1. **Upstream stays green.** Zero changes to the behavior of existing public
   APIs. The mechanical proof is that upstream's own pytest and behave suites
   pass unmodified on every PR.
2. **Refusal atomicity.** Every mutating operation validates fully before it
   changes anything. Every documented refusal condition needs a test asserting
   both that the typed refusal (a `pptx.errors.PaperRefusal` subclass) is
   raised and that output bytes equal input bytes.
3. **The reopen rule.** Test assertions about document content go
   save → reopen → assert. Never assert on the in-memory object you just
   mutated.
4. **Changed-part budgets.** Contract tests in `tests/paper/` assert exactly
   which package parts an operation may touch; a one-slide edit that dirties
   sixty parts is a bug even if the output renders correctly.
5. **No fix without a fixture.** Bug fixes come with a fixture (or golden)
   under `tests/paper/fixtures/` / `tests/paper/goldens/` that fails before
   the fix and passes after.
6. **No new runtime dependencies.** Dev/test dependencies are negotiable;
   runtime dependencies are not.
7. **Whitespace is content.** Meaningful text whitespace — including a
   trailing space inside a run — must survive every code path.

Converting a documented typed refusal into a correct, tested operation is the
sanctioned way to grow this package. Silent fallbacks and best-effort partial
edits are not.

## Releases

Releases are cut from `v*` tags by `.github/workflows/release.yaml` (build,
quality gate, PyPI trusted publishing). Release sign-off also includes the
manual verification checklist in `tests/paper/RELEASE-CHECKLIST.md`.

## Proposing changes

Open a [GitHub Issue](https://github.com/paper-instruments/paper-pptx/issues)
first for anything beyond a small fix — especially new API surface, which must
fit the perceive / edit / compose / verify model and the refusal contract
described in [`docs/user/paper-additions.rst`](docs/user/paper-additions.rst).
