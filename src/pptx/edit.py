"""Structural-anchor write APIs: formatting-preserving text replacement (paper-pptx addition).

The write half of the loop `pptx.inspect` opened. `inspect_text` produces structural,
content-validated anchors; this module consumes them. The run-preservation semantics:

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
from pptx.inspect import (
    _ANCHOR_VERSION,
    BlockAnchor,
    _paragraph_fingerprint,
    _paragraph_literal_text,
    content_hash,
    iter_text_bodies,
)
from pptx.oxml.ns import qn

if TYPE_CHECKING:
    from pptx.presentation import Presentation

RESULT_SCHEMA_NAME = "paper-replace-result"
RESULT_SCHEMA_VERSION = 2


@dataclass(frozen=True)
class _BlockTarget:
    """One exact paragraph plus the structural facts needed to anchor it."""

    part: str
    block_index: int
    owner: object
    paragraph_index: int
    cell_coordinates: object
    paragraph: object


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
        for target in plan:
            count = _replace_in_paragraph(target.paragraph, find, replace)
            if count:
                total += count
                touched.append(_anchor_for_target(target))
        return ReplaceResult(total, tuple(touched))


def replace_text_at(
    prs: "Presentation", anchor: BlockAnchor, find: str, replace: str
) -> ReplaceResult:
    """Replace literal `find` in the block at `anchor`; return a |ReplaceResult|.

    Matching, validation, run-boundary formatting, and non-run boundaries are the same as
    :func:`replace_text`. Current anchors resolve the exact shape or table cell first, then require
    a unique matching full fingerprint in that container. Changed content raises
    |StaleAnchorError|; missing or ambiguous structure refuses without mutation. Legacy
    three-field anchors search only for an exact unique short hash in their named part. `find`
    absent from the resolved block, or present only across a boundary a match cannot cross, raises
    |TargetNotFoundError| rather than returning a zero count.

    On success the result contains the number of occurrences replaced in that block and its one
    post-edit anchor.
    """
    _validate_find_replace(find, replace)
    if not isinstance(anchor, BlockAnchor):
        raise ValueError("anchor must be a BlockAnchor, got %r" % (anchor,))
    _require_presentation_root(prs)
    from pptx._transaction import PackageTransaction

    with PackageTransaction(prs.part.package, prs):
        target = (
            _resolve_legacy_target(prs, anchor)
            if anchor.is_legacy
            else _resolve_current_target(prs, anchor)
        )
        p = target.paragraph
        current_text = _paragraph_literal_text(p)
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
            (_anchor_for_target(target),),
        )


def refind(prs: "Presentation", anchor: BlockAnchor) -> BlockAnchor:
    """Return a fresh current anchor for an exact unique fingerprint match.

    Current anchors search only their structurally identified shape or table cell. Legacy
    three-field anchors search the whole named part by their pinned short text hash. Neither path
    uses ordinal preference, shape names, geometry, or approximate text.
    """
    if not isinstance(anchor, BlockAnchor):
        raise ValueError("anchor must be a BlockAnchor, got %r" % (anchor,))
    _require_presentation_root(prs)
    if anchor.is_legacy:
        return _anchor_for_target(_resolve_legacy_target(prs, anchor))
    spTree, container = _resolve_current_container(prs, anchor)
    matches = [
        _target_from_container(anchor.part, spTree, container, index, p)
        for index, p in enumerate(container["txBody"].findall(qn("a:p")))
        if _paragraph_fingerprint(p) == anchor.content_hash
    ]
    if not matches:
        raise TargetNotFoundError(
            "no paragraph in the anchored container of %s has fingerprint %s; the anchored"
            " content is gone"
            % (anchor.part, anchor.content_hash)
        )
    if len(matches) > 1:
        raise AmbiguousTargetError(
            "%d paragraphs in the anchored container of %s have fingerprint %s"
            % (len(matches), anchor.part, anchor.content_hash)
        )
    return _anchor_for_target(matches[0])


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
    """Return the complete structural paragraph plan, refusing before any mutation.

    Exhausting the traversal before any mutation is what makes deck-wide replacement
    refusal-atomic: the depth guard and the markup-compatibility refusal below fire while
    the document is still untouched. `mc:AlternateContent` refuses because "replace every
    occurrence" cannot be honored over content this library cannot see into.
    """
    plan = []
    for partname, spTree in _iter_story_trees(prs, include_notes):
        plan.extend(_iter_block_targets(partname, spTree, refuse_alternate_content=True))
    return plan


def _resolve_legacy_target(prs, anchor: BlockAnchor) -> _BlockTarget:
    """Resolve a three-field anchor only by an exact unique part-wide legacy hash."""
    spTree = _story_tree(prs, anchor.part)
    matches = [
        target
        for target in _iter_block_targets(anchor.part, spTree)
        if content_hash(_paragraph_literal_text(target.paragraph)) == anchor.content_hash
    ]
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
                [target.block_index for target in matches],
            )
        )
    return matches[0]


def _resolve_current_target(prs, anchor: BlockAnchor) -> _BlockTarget:
    """Resolve current structural identity, then validate content and local uniqueness."""
    spTree, container = _resolve_current_container(prs, anchor)
    locator = anchor.locator
    paragraph_index = locator.get("paragraph_index")
    if not isinstance(paragraph_index, int) or isinstance(paragraph_index, bool):
        raise UnsupportedStructureError(
            "current anchor locator requires an integer paragraph_index"
        )
    paragraphs = container["txBody"].findall(qn("a:p"))
    if not 0 <= paragraph_index < len(paragraphs):
        raise TargetNotFoundError(
            "paragraph index %r is beyond the %d paragraphs in the anchored container of %s"
            % (paragraph_index, len(paragraphs), anchor.part)
        )
    paragraph = paragraphs[paragraph_index]
    current_fingerprint = _paragraph_fingerprint(paragraph)
    if current_fingerprint != anchor.content_hash:
        raise StaleAnchorError(
            "anchor is stale: structurally resolved paragraph %d of %s now fingerprints %s"
            " (anchor says %s); use pptx.edit.refind() to recover"
            % (paragraph_index, anchor.part, current_fingerprint, anchor.content_hash)
        )
    matching_indices = [
        index
        for index, candidate in enumerate(paragraphs)
        if _paragraph_fingerprint(candidate) == anchor.content_hash
    ]
    if len(matching_indices) > 1:
        raise AmbiguousTargetError(
            "%d paragraphs in the anchored container of %s share fingerprint %s"
            " (local indices %s); refusing to trust paragraph_index"
            % (len(matching_indices), anchor.part, anchor.content_hash, matching_indices)
        )
    return _target_from_container(anchor.part, spTree, container, paragraph_index, paragraph)


def _resolve_current_container(prs, anchor: BlockAnchor):
    """Resolve and return a current anchor's exact supported text container."""
    if anchor.version != _ANCHOR_VERSION or not isinstance(anchor.locator, dict):
        raise UnsupportedStructureError(
            "anchor version/locator is unsupported (expected current version %d)"
            % _ANCHOR_VERSION
        )
    locator = anchor.locator
    kind = locator.get("kind")
    if kind not in ("shape", "table-cell"):
        detail = "mc:AlternateContent" if kind == "alternate-content" else repr(kind)
        raise UnsupportedStructureError(
            "anchored container kind %s is a blind or unsupported text region" % detail
        )
    shape_id = locator.get("shape_id")
    if not isinstance(shape_id, int) or isinstance(shape_id, bool) or shape_id <= 0:
        raise UnsupportedStructureError("current anchor locator requires a positive shape_id")
    spTree = _story_tree(prs, anchor.part)

    matching_owners = [
        owner for owner in _iter_shape_owners(spTree) if _owner_shape_id(owner) == shape_id
    ]
    if not matching_owners:
        raise TargetNotFoundError(
            "no container with shape id %d exists in %s" % (shape_id, anchor.part)
        )
    if len(matching_owners) > 1:
        raise AmbiguousTargetError(
            "%d containers in %s claim shape id %d; refusing ambiguous structural identity"
            % (len(matching_owners), anchor.part, shape_id)
        )
    structural_owner = matching_owners[0]
    resolved = {"owner": structural_owner, "cells": {}}
    for actual_kind, owner, txBody, _, cell_coordinates in iter_text_bodies(spTree):
        if owner is not structural_owner:
            continue
        resolved["kind"] = (
            "shape" if actual_kind in ("shape", "group")
            else "table-cell" if actual_kind == "table-cell"
            else actual_kind
        )
        if resolved["kind"] == "shape":
            resolved["txBody"] = txBody
        elif resolved["kind"] == "table-cell":
            resolved["cells"][cell_coordinates] = txBody

    if "kind" not in resolved:
        raise UnsupportedStructureError(
            "shape id %d in %s is not a supported text container" % (shape_id, anchor.part)
        )
    if resolved["kind"] != kind:
        raise UnsupportedStructureError(
            "shape id %d in %s is a %s container, not the anchored %s container"
            % (shape_id, anchor.part, resolved["kind"], kind)
        )
    if kind == "shape":
        return spTree, {
            "kind": kind,
            "owner": structural_owner,
            "txBody": resolved["txBody"],
            "cell_coordinates": None,
        }

    row = locator.get("row")
    column = locator.get("column")
    if any(not isinstance(value, int) or isinstance(value, bool) for value in (row, column)):
        raise UnsupportedStructureError(
            "table-cell anchor locator requires integer row and column coordinates"
        )
    coordinates = row, column
    txBody = resolved["cells"].get(coordinates)
    if txBody is None:
        raise TargetNotFoundError(
            "table cell (%d, %d) does not exist in shape id %d of %s"
            % (row, column, shape_id, anchor.part)
        )
    return spTree, {
        "kind": kind,
        "owner": structural_owner,
        "txBody": txBody,
        "cell_coordinates": coordinates,
    }


def _story_tree(prs, partname):
    """Return the unique slide/notes story tree named `partname`."""
    matches = [
        spTree
        for candidate_partname, spTree in _iter_story_trees(prs, include_notes=True)
        if candidate_partname == partname
    ]
    if not matches:
        raise TargetNotFoundError(
            "no slide or notes part named %r in this presentation" % partname
        )
    if len(matches) > 1:
        raise AmbiguousTargetError(
            "%d slide or notes stories are named %r; refusing ambiguous part identity"
            % (len(matches), partname)
        )
    return matches[0]


def _iter_block_targets(partname, spTree, *, refuse_alternate_content=False):
    """Yield editable paragraph targets in the shared diagnostic traversal order."""
    block_index = 0
    for kind, owner, txBody, _, cell_coordinates in iter_text_bodies(spTree):
        if kind == "alternate-content" and refuse_alternate_content:
            raise UnsupportedStructureError(
                "%s contains mc:AlternateContent; deck-wide replacement cannot guarantee"
                " every occurrence is reached inside markup-compatibility branches"
                " (inspect_text reports these as blind regions)" % partname
            )
        if txBody is None:
            block_index += 1
            continue
        for paragraph_index, paragraph in enumerate(txBody.findall(qn("a:p"))):
            yield _BlockTarget(
                partname,
                block_index,
                owner,
                paragraph_index,
                cell_coordinates,
                paragraph,
            )
            block_index += 1


def _target_from_container(partname, spTree, container, paragraph_index, paragraph):
    """Build a target with its refreshed part-wide diagnostic block index."""
    for target in _iter_block_targets(partname, spTree):
        if target.paragraph is paragraph:
            return _BlockTarget(
                partname,
                target.block_index,
                container["owner"],
                paragraph_index,
                container["cell_coordinates"],
                paragraph,
            )
    raise TargetNotFoundError(
        "the structurally resolved paragraph is not reachable in %s" % partname
    )


def _anchor_for_target(target: _BlockTarget) -> BlockAnchor:
    """Return a fresh current structural anchor for `target`."""
    locator = {
        "kind": "table-cell" if target.cell_coordinates is not None else "shape",
        "shape_id": _owner_shape_id(target.owner),
    }
    if target.cell_coordinates is not None:
        locator["row"], locator["column"] = target.cell_coordinates
    locator["paragraph_index"] = target.paragraph_index
    return BlockAnchor(
        target.part,
        target.block_index,
        _paragraph_fingerprint(target.paragraph),
        _ANCHOR_VERSION,
        locator,
    )


def _owner_shape_id(owner) -> int:
    """Return an owner's cNvPr id, or zero for a blind owner without one."""
    cNvPr = owner.find(".//%s" % qn("p:cNvPr"))
    if cNvPr is None:
        return 0
    try:
        return int(cNvPr.get("id"))
    except (TypeError, ValueError):
        return 0


def _iter_shape_owners(container):
    """Yield every ordinary shape owner recursively, excluding blind AC branches."""
    for owner in container.iter_shape_elms():
        yield owner
        if owner.tag == qn("p:grpSp"):
            yield from _iter_shape_owners(owner)


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
