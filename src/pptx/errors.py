"""Typed safe-refusal exceptions for paper-pptx APIs.

A `PaperRefusal` means the operation declined to run and the document — the in-memory XML tree
and any file on disk — is exactly as it was. Refusals are a success mode: the API judged the
requested change unsafe or unsupported and said so instead of guessing. Callers can therefore
catch `PaperRefusal` distinctly from bugs; programmer errors (bad argument types or values)
remain `TypeError`/`ValueError` as usual.

Paper mutating APIs run inside a package transaction: whatever the operation touched is
restored if it refuses, and the deck-wide check runs before anything commits. Some APIs
(`apply_footers`, `append_deck`, `import_slide`, clone) additionally validate in full before
touching anything. Either way a refusal never leaves a partial edit behind.
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
    """The package archive has no single unambiguous reading, or is unsafe to expand.

    Raised reading a package (ambiguous ZIP end records, duplicate or overlapping members,
    encryption, malformed `[Content_Types].xml`) and writing one (parts that would collide on a
    single ZIP member name).
    """


class AmbiguousTargetError(PaperRefusal):
    """The addressing given matches more than one target; refusing to pick one."""


class TargetNotFoundError(PaperRefusal):
    """The target could not be resolved in this document.

    Either the addressing given (name, index, id, section) matches nothing, or a proxy handed in has
    gone stale because the shape, part, or content it wrapped is no longer reachable. See
    |StaleAnchorError| for the content-hash case.
    """


class StaleAnchorError(TargetNotFoundError):
    """The block at an anchor's position no longer matches the anchor's content hash.

    The document changed since the anchor was produced. Refusing beats guessing: use
    `pptx.edit.refind()` to recover a fresh anchor explicitly. (Subclass of
    |TargetNotFoundError| so existing handlers keep working.)
    """


class UnsupportedStructureError(PaperRefusal):
    """This API cannot operate safely on the document as it stands.

    Covers input it will not touch (unsupported structure, an unreadable package, a signed deck a
    rewrite would invalidate) and, at commit or `batch()` exit, edits whose result would not reopen
    as a presentation. In that second case the edits roll back.
    """


class BoundaryViolationError(PaperRefusal):
    """An operation ran outside the scope where it is safe.

    Today that means one thing: saving while a `batch()` block is open on the package. Those edits
    are not validated yet and may still roll back, so save after the block closes.
    """


class RelationshipPolicyError(PaperRefusal):
    """The relationship graph cannot be carried across as asked.

    Either the source carries relationship types this operation's ledger does not support, or the
    graph itself is unusable: a malformed relationship collection, an invalid target mode, or an
    internal relationship pointing outside its own package.
    """


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
