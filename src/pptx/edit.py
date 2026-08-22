"""Anchor-consuming write APIs: formatting-preserving text replacement (paper-pptx addition).

The write half of the loop `pptx.inspect` opened. `inspect_text` produces content-hash
anchors; this module consumes them. The run-preservation semantics:

- literal, case-sensitive matching; matches never cross paragraph, line-break (`a:br`), or
  field (`a:fld`) boundaries;
- runs are split at match boundaries; boundary fragments keep their source run's `rPr`
  verbatim; replacement text inherits the `rPr` of the run where the match STARTS;
- runs the match does not touch stay byte-identical; runs consumed whole are removed;
- staleness refuses (|StaleAnchorError|) — recovery is the explicit `refind()`, never a
  silent re-find.

Traversal is visibility-complete via the same walker `inspect_text` uses (grouped shapes,
table cells), so text that inspection can see, replacement can reach.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterator, List, Tuple

from pptx.errors import (
    AmbiguousTargetError,
    StaleAnchorError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from pptx.inspect import BlockAnchor, content_hash, iter_text_bodies
from pptx.oxml.ns import qn

if TYPE_CHECKING:
    from pptx.presentation import Presentation

RESULT_SCHEMA_NAME = "paper-replace-result"
RESULT_SCHEMA_VERSION = 1


@dataclass(frozen=True)
class ReplaceResult:
    """Outcome of a text replacement.

    Fields:

    * ``replacements`` -- total number of occurrences replaced across the deck.
    * ``blocks`` -- POST-edit :class:`pptx.inspect.BlockAnchor` for each block that was
      touched (their content hashes reflect the new text).
    """

    replacements: int
    blocks: Tuple[BlockAnchor, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict:
        """Return the replace result as a JSON-ready dict stamped with its schema and version."""
        return {
            "schema": RESULT_SCHEMA_NAME,
            "version": RESULT_SCHEMA_VERSION,
            "replacements": self.replacements,
            "blocks": [anchor.to_dict() for anchor in self.blocks],
        }


def replace_text(
    prs: "Presentation", find: str, replace: str, *, include_notes: bool = False
) -> ReplaceResult:
    """Replace literal `find` with `replace` across `prs`; return a |ReplaceResult|.

    Matching is case-sensitive, left-to-right, and non-overlapping. A match may span consecutive
    text runs in one paragraph, but never a paragraph boundary or an intervening non-run element
    such as a line break or field. `find` must be a non-empty string; `replace` may be empty. Tabs
    are accepted, while line breaks, XML control characters, and non-encodable text are refused.

    Boundary fragments retain their original run properties, and replacement text receives the
    run properties of the run where its match starts. Untouched runs remain byte-identical. Any
    run left with no text after replacement is removed, so formatting found only on that run is
    not retained.

    Traversal covers slide shapes, grouped shapes, and table cells, plus existing notes slides
    when `include_notes=True`. An unsupported blind region refuses the whole operation before any
    write. Zero matches is a successful result with ``replacements == 0`` and no block anchors.
    Each returned block anchor describes the post-edit text of a block that changed.
    """
    _validate_find_replace(find, replace)
    _require_presentation_root(prs)
    from pptx._transaction import PackageTransaction

    with PackageTransaction(prs.part.package, prs):
        # -- validate-fully-then-mutate (§1.3): materialize the COMPLETE traversal first,
        # -- so any traversal refusal fires before the first write.
        plan = _materialize_blocks(prs, include_notes)
        total = 0
        touched: List[BlockAnchor] = []
        for partname, block_index, p in plan:
            count = _replace_in_paragraph(p, find, replace)
            if count:
                total += count
                touched.append(
                    BlockAnchor(partname, block_index, content_hash(_paragraph_text(p)))
                )
        return ReplaceResult(total, tuple(touched))


def replace_text_at(
    prs: "Presentation", anchor: BlockAnchor, find: str, replace: str
) -> ReplaceResult:
    """Replace literal `find` in the block at `anchor`; return a |ReplaceResult|.

    Matching, validation, run-boundary formatting, and non-run boundaries are the same as
    :func:`replace_text`. The block's current text must hash to `anchor.content_hash`; otherwise
    |StaleAnchorError| is raised and :func:`refind` is the explicit recovery path. An anchor that
    addresses an unsupported blind region, or a missing target, refuses without mutation; blind
    regions elsewhere do not block this anchored operation. `find` absent from the block, or
    present only across a boundary a match cannot cross, raises |TargetNotFoundError| rather than
    returning a zero count.

    On success the result contains the number of occurrences replaced in that block and its one
    post-edit anchor.
    """
    _validate_find_replace(find, replace)
    if not isinstance(anchor, BlockAnchor):
        raise ValueError("anchor must be a BlockAnchor, got %r" % (anchor,))
    _require_presentation_root(prs)
    from pptx._transaction import PackageTransaction

    with PackageTransaction(prs.part.package, prs):
        p = _block_paragraph(prs, anchor)
        current_text = _paragraph_text(p)
        current_hash = content_hash(current_text)
        if current_hash != anchor.content_hash:
            raise StaleAnchorError(
                "anchor is stale: block %d of %s now hashes %s (anchor says %s); the document"
                " changed since the anchor was produced — use pptx.edit.refind() to recover"
                % (anchor.block_index, anchor.part, current_hash, anchor.content_hash)
            )
        if find not in current_text:
            raise TargetNotFoundError(
                "%r does not occur in the anchored block (text %r)" % (find, current_text)
            )
        count = _replace_in_paragraph(p, find, replace)
        if count == 0:
            # -- `find` appears in the hash-text but every occurrence crosses a field or
            # -- line-break boundary, which matches never do.
            raise TargetNotFoundError(
                "%r occurs in the anchored block only across a field or line-break boundary;"
                " matches never cross a:fld/a:br" % (find,)
            )
        return ReplaceResult(
            count,
            (BlockAnchor(anchor.part, anchor.block_index, content_hash(_paragraph_text(p))),),
        )


def refind(prs: "Presentation", anchor: BlockAnchor) -> BlockAnchor:
    """Return a fresh anchor for the unique block still matching `anchor.content_hash`.

    The explicit recovery path for a stale anchor: searches every block of the anchor's
    part by content hash. No match → |TargetNotFoundError|; more than one →
    |AmbiguousTargetError| (the hash alone cannot say which block was meant).
    """
    if not isinstance(anchor, BlockAnchor):
        raise ValueError("anchor must be a BlockAnchor, got %r" % (anchor,))
    matches = []
    for partname, spTree in _iter_story_trees(prs, include_notes=True):
        if partname != anchor.part:
            continue
        block_index = 0
        for kind, _, txBody, _, _ in iter_text_bodies(spTree):
            if txBody is None:
                block_index += 1  # -- occupies an index but is never hash-matchable
                continue
            for p in txBody.findall(qn("a:p")):
                if content_hash(_paragraph_text(p)) == anchor.content_hash:
                    matches.append(BlockAnchor(partname, block_index, anchor.content_hash))
                block_index += 1
    if not matches:
        raise TargetNotFoundError(
            "no block in %s hashes to %s; the anchored content is gone"
            % (anchor.part, anchor.content_hash)
        )
    if len(matches) > 1:
        raise AmbiguousTargetError(
            "%d blocks in %s hash to %s (indices %s); refusing to pick one"
            % (
                len(matches),
                anchor.part,
                anchor.content_hash,
                [m.block_index for m in matches],
            )
        )
    return matches[0]


# ------------------------------------------------------------------------------- internals


def _require_presentation_root(prs) -> None:
    """Refuse a retained Presentation proxy whose root has been replaced."""
    if getattr(prs, "_element", None) is not getattr(prs.part, "_element", None):
        raise TargetNotFoundError(
            "presentation is stale: its XML root is no longer the live package root"
        )


def _validate_find_replace(find, replace) -> None:
    """Full validation before any mutation (§1.3)."""
    if not isinstance(find, str) or not find:
        raise ValueError("find must be a non-empty str, got %r" % (find,))
    if not isinstance(replace, str):
        raise ValueError("replace must be a str, got %r" % (replace,))
    for name, value in (("find", find), ("replace", replace)):
        try:
            value.encode("utf-8")
        except UnicodeEncodeError:
            raise ValueError("%s contains characters not encodable in XML: %r" % (name, value))
        if any(ch in value for ch in "\n\r\x0b"):
            raise ValueError(
                "%s must not contain line breaks (matches and replacements never cross"
                " paragraph or line-break boundaries); got %r" % (name, value)
            )
        # -- C0 controls (other than tab) are invalid in XML 1.0 text; the upstream setter
        # -- would silently rewrite them as _xHHHH_ escape literals visible to the reader
        if any(ch < " " and ch != "\t" for ch in value):
            raise ValueError(
                "%s contains control characters that cannot appear in XML text: %r"
                % (name, value)
            )


def _iter_story_trees(prs, include_notes) -> "Iterator[Tuple[str, object]]":
    """Yield (partname, spTree) for every slide (and notes slide when asked), deck order."""
    from pptx.errors import materialize_slides

    for slide in materialize_slides(prs, "replace_text"):
        yield str(slide.part.partname), slide._element.spTree
        if include_notes and slide.has_notes_slide:
            notes_part = slide.part.part_related_by(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
            )
            yield str(notes_part.partname), notes_part._element.spTree


def _materialize_blocks(prs, include_notes):
    """Return the complete [(partname, block_index, a:p element)] plan, refusing first.

    Exhausting the traversal before any mutation is what makes deck-wide replacement
    refusal-atomic: the depth guard and the markup-compatibility refusal below fire while
    the document is still untouched. `mc:AlternateContent` refuses because "replace every
    occurrence" cannot be honored over content this library cannot see into.
    """
    plan = []
    for partname, spTree in _iter_story_trees(prs, include_notes):
        block_index = 0
        for kind, _, txBody, _, _ in iter_text_bodies(spTree):
            if kind == "alternate-content":
                raise UnsupportedStructureError(
                    "%s contains mc:AlternateContent; deck-wide replacement cannot"
                    " guarantee every occurrence is reached inside markup-compatibility"
                    " branches (inspect_text reports these as blind regions)" % partname
                )
            if txBody is None:
                block_index += 1
                continue
            for p in txBody.findall(qn("a:p")):
                plan.append((partname, block_index, p))
                block_index += 1
    return plan


def _block_paragraph(prs, anchor: BlockAnchor):
    """Return the `a:p` element at `anchor.block_index` within `anchor.part`.

    Block indices align with `inspect_text`: an `mc:AlternateContent` subtree occupies
    exactly one (blind) index; anchoring it refuses.
    """
    for partname, spTree in _iter_story_trees(prs, include_notes=True):
        if partname != anchor.part:
            continue
        block_index = 0
        for kind, _, txBody, _, _ in iter_text_bodies(spTree):
            if txBody is None:
                if block_index == anchor.block_index:
                    detail = (
                        "mc:AlternateContent (a blind region)"
                        if kind == "alternate-content"
                        else "%s (a blind graphic region)" % kind
                    )
                    raise UnsupportedStructureError(
                        "the anchored block is %s; this content is not editable in v0.1"
                        % detail
                    )
                block_index += 1
                continue
            for p in txBody.findall(qn("a:p")):
                if block_index == anchor.block_index:
                    return p
                block_index += 1
        raise TargetNotFoundError(
            "block index %d is beyond the %d blocks of %s"
            % (anchor.block_index, block_index, anchor.part)
        )
    raise TargetNotFoundError("no slide or notes part named %r in this presentation" % anchor.part)


def _paragraph_text(p) -> str:
    """Concatenated `a:r` run text of `p` — same definition `inspect_text` hashes."""
    pieces = []
    for child in p:
        if child.tag == qn("a:r"):
            text = child.find(qn("a:t"))
            pieces.append((text.text or "") if text is not None else "")
        elif child.tag == qn("a:br"):
            pieces.append("\v")
    return "".join(pieces)


def _replace_in_paragraph(p, find: str, replace: str) -> int:
    """Apply the pinned run-preservation replacement inside one paragraph; return count."""
    total = 0
    for segment in _run_segments(p):
        total += _replace_in_segment(segment, find, replace)
    return total


def _run_segments(p) -> "List[List[object]]":
    """Split `p`'s children into maximal consecutive `a:r` sequences.

    `a:br`, `a:fld`, and any other intervening element end a segment: visible text is
    discontinuous there, so matches must not cross.
    """
    segments: "List[List[object]]" = []
    current: "List[object]" = []
    for child in p:
        if child.tag == qn("a:r"):
            current.append(child)
        elif current:
            segments.append(current)
            current = []
    if current:
        segments.append(current)
    return segments


def _replace_in_segment(runs, find: str, replace: str) -> int:
    """Replace every occurrence of `find` across `runs`.

    Characters the edit does not touch keep their own run, so formatting survives a replacement that
    straddles a run boundary.
    """
    texts = [_run_text_of(r) for r in runs]
    full = "".join(texts)
    matches = _find_occurrences(full, find)
    if not matches:
        return 0

    # -- piece stream: (source_run_index, text). Untouched characters keep their own run;
    # -- a replacement inherits the run where its match starts. Positions are monotone, so
    # -- pieces group contiguously per run and per-run reassembly below is order-safe.
    pieces: "List[Tuple[int, str]]" = []
    position = 0
    run_starts = []
    offset = 0
    for text in texts:
        run_starts.append(offset)
        offset += len(text)

    def run_index_at(char_index: int) -> int:
        for index in range(len(run_starts) - 1, -1, -1):
            if run_starts[index] <= char_index:
                return index
        return 0

    for start, end in matches:
        if start > position:
            _append_retained(pieces, full, position, start, run_starts, texts)
        pieces.append((run_index_at(start), replace))
        position = end
    if position < len(full):
        _append_retained(pieces, full, position, len(full), run_starts, texts)

    # -- reassemble per original run: unchanged text -> element untouched (byte-identical);
    # -- empty -> element removed; changed -> only the a:t text is rewritten (rPr untouched)
    new_texts = ["" for _ in runs]
    for run_index, text in pieces:
        new_texts[run_index] += text
    for run, old_text, new_text in zip(list(runs), texts, new_texts):
        if new_text == old_text:
            continue
        if new_text == "":
            run.getparent().remove(run)
        else:
            run.text = new_text  # -- CT_RegularTextRun.text setter; rPr untouched
    return len(matches)


def _append_retained(pieces, full, start, end, run_starts, texts) -> None:
    """Append retained (untouched) characters [start, end) split by their owning runs."""
    for run_index, run_start in enumerate(run_starts):
        run_end = run_start + len(texts[run_index])
        lo = max(start, run_start)
        hi = min(end, run_end)
        if lo < hi:
            pieces.append((run_index, full[lo:hi]))


def _find_occurrences(text: str, find: str) -> "List[Tuple[int, int]]":
    """Non-overlapping (start, end) occurrences of `find`, left to right."""
    occurrences = []
    position = 0
    while True:
        index = text.find(find, position)
        if index == -1:
            return occurrences
        occurrences.append((index, index + len(find)))
        position = index + len(find)


def _run_text_of(r) -> str:
    """Text of a run's `a:t` child, or an empty string when absent."""
    t = r.find(qn("a:t"))
    return (t.text or "") if t is not None else ""
