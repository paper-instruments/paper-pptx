"""`docs/dev/analysis/` must survive future documentation cleanups.

The Sphinx tree was deleted and `docs/` now holds exactly one thing: 79 `.rst` files under
`docs/dev/analysis/`. They are python-pptx's own per-feature design notes, inherited with the
fork and never modified. Each pairs prose with a real XML specimen and the matching xsd
excerpt, so they document the OOXML format and how the oxml layer models it - the layer this
fork is forbidden to change. `src/pptx/oxml/AGENTS.md` requires reading the relevant one
before adding XML vocabulary.

The hazard this test exists for: those files sit *inside* the tree that was just emptied. The
next person tidying documentation finds 79 unpublished `.rst` files with no obvious reason to
exist, and `git rm -r docs/` looks like the obvious move. Nothing else in the repo would fail.
"There is always git history" does not save it - you cannot read a file you must first
resurrect, and the requirement is that an implementer reads it before writing code.

They stay as `.rst` deliberately. They were never rendered anywhere, so the format is inert,
and `origin` has no upstream remote, so converting them later carries no merge risk.
"""

from __future__ import annotations

import pathlib
import subprocess

REPO = pathlib.Path(__file__).resolve().parents[2]
ANALYSIS = REPO / "docs" / "dev" / "analysis"
EXPECTED_ANALYSIS_FILES = 79


def _tracked_rst() -> list:
    """Return every tracked `.rst` path in the repo, as posix strings."""
    out = subprocess.run(
        ["git", "ls-files", "*.rst"],
        cwd=REPO,
        capture_output=True,
        text=True,
        check=True,
    )
    return sorted(line for line in out.stdout.splitlines() if line)


class DescribeTheDocsTree:
    """After the Sphinx retirement, docs/ holds only the inherited analysis notes."""

    def test_the_analysis_notes_are_all_present(self):
        found = sorted(p.name for p in ANALYSIS.rglob("*.rst"))
        assert len(found) == EXPECTED_ANALYSIS_FILES, (
            f"expected {EXPECTED_ANALYSIS_FILES} .rst files under docs/dev/analysis/, found "
            f"{len(found)}. These are upstream's design notes and are required reading before "
            f"adding XML vocabulary - see src/pptx/oxml/AGENTS.md. If you are deleting docs, "
            f"delete siblings by explicit path; never `git rm -r docs/`."
        )

    def test_no_rst_survives_outside_the_analysis_notes(self):
        strays = [p for p in _tracked_rst() if not p.startswith("docs/dev/analysis/")]
        assert not strays, (
            f"reST was retired from this repo; {strays} reintroduces it. Documentation goes in "
            f"Markdown, or in docstrings, or on the docs site."
        )

    def test_sphinx_config_is_gone(self):
        """conf.py defined 169 substitutions that docstrings depended on to render."""
        assert not (REPO / "docs" / "conf.py").exists()
        assert not (REPO / "docs" / "Makefile").exists()
