"""Fork version and distribution-conflict guard."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, distribution

__paper_version__ = "0.2.0"


def assert_distribution_identity() -> None:
    """Fail early when both distributions claim the frozen ``pptx`` import."""
    try:
        distribution("python-pptx")
    except PackageNotFoundError:
        return
    try:
        distribution("paper-pptx")
    except PackageNotFoundError:
        return
    raise ImportError(
        "paper-pptx and python-pptx are both installed and cannot safely share the 'pptx' "
        "package; uninstall both, then install only paper-pptx"
    )
