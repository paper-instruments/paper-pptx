"""Symbols named in the hand-written documentation must still exist and still be public.

The generated API reference cannot drift: it is rebuilt from these docstrings. The hand-written
pages can. They name `pptx.*` symbols in prose and code samples, and nothing connects those
mentions to the code - rename a function and the site keeps presenting the old name as live,
with no build failing anywhere.

That is the drift that actually costs someone time: `griffe check` catches the rename, but only
this catches the documentation still describing what was renamed.

The pages live in the `paper-office-docs` repository, so this test needs a path to it and skips
when it has no way to find one. Skipping rather than failing is deliberate: a test that fails
because a sibling checkout is missing gets marked xfail and stops protecting anything.

    PAPER_OFFICE_DOCS=../paper-office-docs uv run pytest tests/paper/test_docs_symbols.py

This test failing regularly is also the signal that the hand-written pages belong in this
repository, so a rename and its documentation land in one reviewable change.
"""

from __future__ import annotations

import importlib
import os
import pathlib
import re

import pytest

# -- only the hand-written pages; the api/ tree is generated from the code and cannot drift --
HAND_WRITTEN = ("index.mdx", "reference.mdx", "vs-python-pptx.mdx")

# -- A dotted pptx path. Deliberately narrow: a looser pattern matches prose and produces
# -- failures nobody trusts. The negative lookbehind keeps hostnames out - `pptx.readthedocs.io`
# -- inside `python-pptx.readthedocs.io` is a URL, not a symbol.
SYMBOL = re.compile(r"(?<![\w./-])pptx(?:\.[A-Za-z_][A-Za-z0-9_]*)+")

# -- documented as removed, or named only to say it no longer exists --
IGNORED = {"pptx.scrub"}


def _docs_root() -> pathlib.Path | None:
    """Locate the paper-office-docs checkout, or None when it cannot be found."""

    def usable(path: pathlib.Path) -> bool:
        return (path / "content" / "docs" / "pptx").is_dir()

    configured = os.environ.get("PAPER_OFFICE_DOCS")
    if configured and usable(pathlib.Path(configured)):
        return pathlib.Path(configured)

    # -- search upward for a sibling checkout rather than counting parents: this repo is often
    # -- worked on from a git worktree, which nests the root several levels deeper --
    for ancestor in pathlib.Path(__file__).resolve().parents:
        candidate = ancestor / "paper-office-docs"
        if usable(candidate):
            return candidate
    return None


def _referenced_symbols() -> dict:
    """Return {symbol: [pages naming it]} across the hand-written pptx pages."""
    root = _docs_root()
    if root is None:
        return {}
    found: dict = {}
    for name in HAND_WRITTEN:
        page = root / "content" / "docs" / "pptx" / name
        if not page.is_file():
            continue
        for match in SYMBOL.findall(page.read_text(encoding="utf-8")):
            if match not in IGNORED:
                found.setdefault(match, []).append(name)
    return found


def _resolves(dotted: str) -> bool:
    """True when `dotted` names something importable and public."""
    parts = dotted.split(".")
    for split in range(len(parts), 0, -1):
        module_path = ".".join(parts[:split])
        try:
            module = importlib.import_module(module_path)
        except ImportError:
            continue
        target = module
        for attribute in parts[split:]:
            # -- dunders such as __version__ are public by convention; a single leading
            # -- underscore is not --
            if attribute.startswith("_") and not attribute.startswith("__"):
                return False
            target = getattr(target, attribute, None)
            if target is None:
                return False
        return True
    return False


requires_docs = pytest.mark.skipif(
    _docs_root() is None,
    reason="paper-office-docs checkout not found; set PAPER_OFFICE_DOCS to enable",
)


class DescribeTheHandWrittenDocs:
    """Every pptx symbol the site names by hand must still exist and still be public."""

    @requires_docs
    def test_it_names_symbols_at_all(self):
        """Guards the extraction itself: a regex that matches nothing proves nothing."""
        assert _referenced_symbols(), (
            "no pptx.* symbols were found in the hand-written pages. Either the pages moved or "
            "the extraction pattern stopped matching - either way this test is now vacuous."
        )

    @requires_docs
    def test_every_named_symbol_still_resolves(self):
        unresolved = {
            symbol: pages
            for symbol, pages in sorted(_referenced_symbols().items())
            if not _resolves(symbol)
        }
        assert not unresolved, (
            "the documentation site names symbols that no longer resolve or are no longer "
            f"public: {unresolved}. Either restore them, or update those pages in "
            "paper-office-docs."
        )
