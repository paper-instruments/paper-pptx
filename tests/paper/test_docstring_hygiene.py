"""Docstrings must stay self-contained and MDX-safe.

Sphinx let docstrings depend on `docs/conf.py` for meaning: 169 `rst_epilog` substitutions
turned `|None|` into a cross-reference at build time. That tree is gone, so a docstring
carrying `|None|` now shows literal pipes to every reader - an editor, an agent reading the
installed wheel, and the fumadocs MDX pipeline, where a raw `<a:p>` is parsed as a JSX
component and fails the build outright.

This breakage class is silent. Nothing imports a docstring, no test exercises one, and a
paragraph can lose its meaning without a single failure anywhere. A check is the only thing
that catches it, which is why it lives here rather than in a lint script nobody runs.

The rewriter itself is `tools/docstring_hazards.py`.
"""

from __future__ import annotations

import pathlib
import sys

import pytest

TOOLS = pathlib.Path(__file__).resolve().parents[2] / "tools"
sys.path.insert(0, str(TOOLS))

from docstring_hazards import HAZARDS, SRC, _rewrite, scan  # noqa: E402


def _sites() -> dict:
    """Return every remaining hazard site, keyed by class, as a list of file paths."""
    found = {name: [] for name in HAZARDS}
    for path in sorted(SRC.rglob("*.py")):
        _, counts = scan(path)
        for name, count in counts.items():
            if count:
                found[name].append(f"{path.relative_to(SRC.parents[1])} ({count})")
    return found


class DescribeDocstringHygiene:
    """The four hazard classes must be absent from every docstring under src/pptx."""

    @pytest.mark.parametrize("hazard", HAZARDS)
    def test_no_hazard_survives_in_src(self, hazard: str):
        offenders = _sites()[hazard]
        assert not offenders, (
            f"{len(offenders)} file(s) still carry a '{hazard}' hazard in a docstring. "
            f"Run `uv run python tools/docstring_hazards.py --list` to see every site, then "
            f"`--fix` to rewrite them. Offenders: {', '.join(offenders[:10])}"
        )


class DescribeTheRewriter:
    """The rewriter converts markup and leaves everything else alone."""

    @pytest.mark.parametrize(
        ("before", "after"),
        [
            # -- substitutions lose their pipes --
            ("returns |None| when absent", "returns `None` when absent"),
            # -- roles keep the target, drop the role and its display prefixes --
            ("see :meth:`~Slide.rebind_layout`", "see `Slide.rebind_layout`"),
            ("see :class:`.None`", "see `None`"),
            (":ref:`the label <PpActionType>`", "`the label`"),
            # -- OOXML tags are content: backticked, never removed --
            ("appends <a:p> to the body", "appends `<a:p>` to the body"),
            ("emits <c:delete val='0'/>", "emits `<c:delete val='0'/>`"),
            # -- a bare brace is a JSX expression opener to MDX --
            ("keys are {facet: value}", "keys are `{facet: value}`"),
        ],
    )
    def test_it_converts_each_hazard_class(self, before: str, after: str):
        assert _rewrite(before)[0] == after

    @pytest.mark.parametrize(
        "text",
        [
            "already `<a:p>` backticked",
            "already `None` backticked",
            "plain prose with no markup at all",
        ],
    )
    def test_it_leaves_safe_text_untouched(self, text: str):
        assert _rewrite(text)[0] == text

    def test_it_does_not_touch_code_outside_docstrings(self, tmp_path: pathlib.Path):
        """A hazard in an ordinary string literal is code, not documentation."""
        module = tmp_path / "sample.py"
        module.write_text(
            '"""Docstring with |None| in it."""\n'
            'SEPARATOR = "|None|"\n'
            'TAG = "<a:p>"\n',
            encoding="utf-8",
        )
        rewritten, counts = scan(module)

        assert counts["substitution"] == 1, "only the docstring occurrence converts"
        assert '"""Docstring with `None` in it."""' in rewritten
        assert 'SEPARATOR = "|None|"' in rewritten
        assert 'TAG = "<a:p>"' in rewritten
