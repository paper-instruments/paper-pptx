"""Typed safe-refusal exceptions for paper-pptx APIs.

A `PaperRefusal` means the operation declined to run and the document — the in-memory XML tree
and any file on disk — is exactly as it was. Refusals are a success mode: the API judged the
requested change unsafe or unsupported and said so instead of guessing. Callers can therefore
catch `PaperRefusal` distinctly from bugs; programmer errors (bad argument types or values)
remain `TypeError`/`ValueError` as usual.

Every paper mutating API is structured validate-fully-then-mutate, so a refusal can never
leave a partial edit behind.
"""

from __future__ import annotations

__all__ = [
    "AmbiguousTargetError",
    "BoundaryViolationError",
    "PackageLimitError",
    "PaperRefusal",
    "RelationshipPolicyError",
    "StaleAnchorError",
    "TargetNotFoundError",
    "UnsupportedStructureError",
]


class PaperRefusal(Exception):
    """Base class for all safe refusals raised by paper-pptx APIs."""


class PackageLimitError(PaperRefusal):
    """The package archive is ambiguous or unsafe to expand."""


class AmbiguousTargetError(PaperRefusal):
    """The addressing given matches more than one target; refusing to pick one."""


class TargetNotFoundError(PaperRefusal):
    """The addressing given matches nothing in this document."""


class StaleAnchorError(TargetNotFoundError):
    """The block at an anchor's position no longer matches the anchor's content hash.

    The document changed since the anchor was produced. Refusing beats guessing: use
    `pptx.edit.refind()` to recover a fresh anchor explicitly. (Subclass of
    |TargetNotFoundError| so existing handlers keep working.)
    """


class UnsupportedStructureError(PaperRefusal):
    """The document contains structure this API cannot operate on safely."""


class BoundaryViolationError(PaperRefusal):
    """The operation would cross a boundary it promised to stay inside."""


class RelationshipPolicyError(PaperRefusal):
    """The relationship graph cannot be honored under the requested policy."""


def materialize_slides(prs, operation: str):
    """Return `list(prs.slides)`, refusing typed when the relationship graph is broken.

    paper-pptx internal helper. Paper organs traverse the whole deck up front;
    corrupt input (a `p:sldId` referencing a missing relationship) must speak from those
    APIs as a typed, specific refusal - never a raw traceback. Upstream loader and
    traversal behavior on such files is unchanged (the additive contract): only the
    paper entry points route through this guard.
    """
    try:
        return list(prs.slides)
    except KeyError as exc:
        raise UnsupportedStructureError(
            "%s refused: the presentation's relationship graph is broken (%s); repair "
            "the package before operating on it" % (operation, exc)
        ) from exc
