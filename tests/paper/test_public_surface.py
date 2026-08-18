"""Each paper-added module must declare its public surface, accurately.

`__all__` is load-bearing for two consumers that fail quietly when it is wrong. `griffe check`
reads it to decide whether a change breaks the public API, so a name missing from it can be
removed without the gate objecting. The docs generator reads it to decide what to publish, so
a name missing from it silently vanishes from the reference instead of failing a build.

Both assertions matter, in opposite directions: a name in `__all__` that does not exist is a
lie, and a public name absent from `__all__` is an omission. Neither shows up anywhere else.

`pptx.package` is deliberately absent from this test. It is an upstream module that paper
extended - it exists at the `paper-base` tag - so declaring its public surface would be a
change to inherited code. Its paper-added members are named explicitly by the docs generator
instead.
"""

from __future__ import annotations

import ast
import importlib
import pathlib

import pytest

PAPER_MODULES = ("inspect", "edit", "diff", "compose", "rebind", "hf", "errors")
SRC = pathlib.Path(__file__).resolve().parents[2] / "src" / "pptx"

# -- Names that read as public but are documented internal helpers, so they are correctly
# -- absent from __all__. `materialize_slides` calls itself "paper-pptx internal helper" and is
# -- imported by six paper modules as a shared refusal guard; the tidy fix is to rename it with
# -- a leading underscore, which is an API change and does not belong in a docs sweep.
INTERNAL_HELPERS = {"errors": {"materialize_slides"}}


def _defined_public_names(module_name: str) -> set:
    """Return the public names `module_name` defines at module level, read from source.

    Read statically rather than through `dir()`, which cannot distinguish a name the module
    defines from one it imported.
    """
    tree = ast.parse((SRC / f"{module_name}.py").read_text(encoding="utf-8"))
    names = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            if not node.name.startswith("_"):
                names.add(node.name)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id.isupper() and target.id[0] != "_":
                    names.add(target.id)
    return names - INTERNAL_HELPERS.get(module_name, set())


class DescribeThePublicSurface:
    """Every paper module declares __all__, and it matches what the module defines."""

    @pytest.mark.parametrize("module_name", PAPER_MODULES)
    def test_it_declares_all(self, module_name: str):
        module = importlib.import_module(f"pptx.{module_name}")
        assert hasattr(module, "__all__"), (
            f"pptx.{module_name} declares no __all__. Both `griffe check` and the docs "
            f"generator read it; without one, neither knows what is public."
        )

    @pytest.mark.parametrize("module_name", PAPER_MODULES)
    def test_every_exported_name_resolves(self, module_name: str):
        module = importlib.import_module(f"pptx.{module_name}")
        missing = sorted(n for n in module.__all__ if not hasattr(module, n))
        assert not missing, (
            f"pptx.{module_name}.__all__ exports {missing}, which the module does not "
            f"define. Remove them or restore the names."
        )

    @pytest.mark.parametrize("module_name", PAPER_MODULES)
    def test_no_public_name_is_left_unexported(self, module_name: str):
        module = importlib.import_module(f"pptx.{module_name}")
        unexported = sorted(_defined_public_names(module_name) - set(module.__all__))
        assert not unexported, (
            f"pptx.{module_name} defines public {unexported} but does not export them. "
            f"An unexported public name is invisible to `griffe check` and absent from the "
            f"generated reference. Add them to __all__, or prefix them with an underscore "
            f"if they were never meant to be public."
        )

    def test_upstream_package_module_stays_undeclared(self):
        """`pptx.package` is inherited; its public surface is not ours to declare."""
        import pptx.package

        assert not hasattr(pptx.package, "__all__"), (
            "pptx.package is an upstream module (present at the paper-base tag). Declaring "
            "__all__ there changes inherited code. Its paper-added members - xml_equivalent, "
            "PartDelta, PackageDiff, diff_package, patch_save - are named explicitly by the "
            "docs generator instead."
        )
