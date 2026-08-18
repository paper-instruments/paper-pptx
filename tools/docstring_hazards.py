"""Find and remove reST/MDX hazards from docstrings.

The docstrings in `src/pptx/` were written for Sphinx: they leaned on 169 `rst_epilog`
substitutions defined in the since-deleted `docs/conf.py`, on Sphinx cross-reference roles, and
on raw OOXML tags. With that configuration gone, `|None|` is literal pipe soup wherever the
docstring is read - in an editor, in the wheel an agent reads, and in the fumadocs MDX pipeline,
where a raw `<a:p>` is parsed as a JSX component and fails the build.

This ran once as a sweep; what keeps it true is `tests/paper/test_docstring_hygiene.py`.

Four hazard classes, all rewritten to inline code spans:

    |Name|              ->  `Name`
    :meth:`target`      ->  `target`
    <a:p>               ->  `<a:p>`      (legitimate content - backtick, never delete)
    {...}               ->  `{...}`      (MDX reads a bare brace as a JSX expression)

Only docstring nodes are touched, reached through the AST. A text-level pass would also
hit ordinary string literals, which in this codebase legitimately contain pipes, angle
brackets, and braces.

    uv run python tools/docstring_hazards.py            # report
    uv run python tools/docstring_hazards.py --list     # report with every site
    uv run python tools/docstring_hazards.py --fix      # rewrite in place
"""

from __future__ import annotations

import argparse
import ast
import pathlib
import re
import sys
from typing import Dict, Iterator, List, Tuple

SRC = pathlib.Path(__file__).resolve().parent.parent / "src" / "pptx"

SUBSTITUTION = re.compile(r"\|([A-Za-z_][A-Za-z0-9_.\-]*)\|")
ROLE = re.compile(r":(?:meth|ref|class|func|attr|mod|obj|exc|term|doc):`([^`]+)`")
XML_TAG = re.compile(r"<[a-zA-Z]+:[a-zA-Z0-9_]+[^>]*?/?>")
BRACE = re.compile(r"\{[^}\n]*\}")

# -- reST's double-backtick literal and Markdown's single-backtick span. Matched longest
# -- first so a ``literal`` is never mistaken for two adjacent `spans`.
CODE_SPAN = re.compile(r"``[^`]*``|`[^`\n]*`")

HAZARDS = ("substitution", "role", "xml_tag", "brace")


def _code_spans(text: str) -> List[Tuple[int, int]]:
    """Return the character ranges of `text` already inside a code span.

    A hazard inside one of these is already displayed verbatim, so wrapping it again would
    nest backticks - which is how ``./c:grouping{val=?}`` becomes ```./c:grouping`{val=?}```
    and stops rendering. Recomputed for every pass, because each pass moves the offsets.
    """
    return [match.span() for match in CODE_SPAN.finditer(text)]


def _role_target(target: str) -> str:
    """Reduce a Sphinx role target to the symbol it names.

    Handles the `text <target>` form by keeping the visible text, and strips the `~` and
    `.` prefixes Sphinx uses to control how much of a dotted path is displayed.
    """
    labelled = re.match(r"^(.*?)\s*<[^>]+>$", target)
    if labelled:
        target = labelled.group(1)
    return target.lstrip("~.").strip()


def _rewrite(text: str) -> Tuple[str, Dict[str, int]]:
    """Return `text` with every hazard converted, plus a per-class count of conversions."""
    counts = dict.fromkeys(HAZARDS, 0)

    def substitute(pattern: re.Pattern[str], name: str, replacement) -> None:
        nonlocal text
        protected = _code_spans(text)
        out: List[str] = []
        cursor = 0
        for match in pattern.finditer(text):
            start, end = match.span()
            # -- contained, not merely overlapping: a role's own `target` is a code span, but
            # -- the `:meth:` prefix sits outside it, so the role still needs rewriting.
            if any(span_start <= start and end <= span_end for span_start, span_end in protected):
                continue
            out.append(text[cursor:start])
            out.append(replacement(match))
            counts[name] += 1
            cursor = end
        out.append(text[cursor:])
        text = "".join(out)

    # -- roles first: they are the outermost construct, and their `target` would otherwise
    # -- read as a code span protecting whatever sits inside it from the later passes.
    substitute(ROLE, "role", lambda m: f"`{_role_target(m.group(1))}`")
    substitute(SUBSTITUTION, "substitution", lambda m: f"`{m.group(1)}`")
    substitute(XML_TAG, "xml_tag", lambda m: f"`{m.group(0)}`")
    substitute(BRACE, "brace", lambda m: f"`{m.group(0)}`")
    return text, counts


def _docstring_spans(source: str) -> Iterator[Tuple[int, int]]:
    """Yield (start, end) character offsets of every docstring literal in `source`."""
    tree = ast.parse(source)
    line_starts = [0]
    for line in source.splitlines(keepends=True):
        line_starts.append(line_starts[-1] + len(line))

    def offset(lineno: int, col: int) -> int:
        return line_starts[lineno - 1] + col

    for node in ast.walk(tree):
        if not isinstance(
            node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)
        ):
            continue
        if ast.get_docstring(node, clean=False) is None:
            continue
        literal = node.body[0].value  # -- the docstring Constant, per get_docstring
        yield (
            offset(literal.lineno, literal.col_offset),
            offset(literal.end_lineno, literal.end_col_offset),
        )


def scan(path: pathlib.Path) -> Tuple[str, Dict[str, int]]:
    """Return `path`'s source with hazards rewritten, plus per-class conversion counts."""
    source = path.read_text(encoding="utf-8")
    totals = dict.fromkeys(HAZARDS, 0)
    pieces: List[str] = []
    cursor = 0
    for start, end in sorted(_docstring_spans(source)):
        rewritten, counts = _rewrite(source[start:end])
        pieces.append(source[cursor:start])
        pieces.append(rewritten)
        for name, count in counts.items():
            totals[name] += count
        cursor = end
    pieces.append(source[cursor:])
    return "".join(pieces), totals


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--fix", action="store_true", help="rewrite files in place")
    parser.add_argument("--list", action="store_true", help="print every hazard site")
    args = parser.parse_args()

    totals = dict.fromkeys(HAZARDS, 0)
    touched: List[pathlib.Path] = []

    for path in sorted(SRC.rglob("*.py")):
        rewritten, counts = scan(path)
        if not any(counts.values()):
            continue
        touched.append(path)
        for name, count in counts.items():
            totals[name] += count
        if args.fix:
            path.write_text(rewritten, encoding="utf-8")
        if args.list:
            relative = path.relative_to(SRC.parent.parent)
            detail = " ".join(f"{n}={c}" for n, c in counts.items() if c)
            print(f"{relative}: {detail}")

    for name in HAZARDS:
        print(f"{name:13} {totals[name]}")
    print(f"{'total':13} {sum(totals.values())} across {len(touched)} files")

    if args.fix:
        print(f"\nrewrote {len(touched)} files")
        return 0
    return 1 if any(totals.values()) else 0


if __name__ == "__main__":
    sys.exit(main())
