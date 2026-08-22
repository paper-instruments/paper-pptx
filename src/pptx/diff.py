"""Deck diff - the verification mirror (paper-pptx addition).

PPTX carries no revision markup: a deck is amnesiac about its own history, so "what
changed between v3 and v4?" has no programmable answer anywhere in the ecosystem. This
provides it as a typed report beside the file - the deck-format analogue of a
redline - assembled from the permanent slide ids, the visibility-complete text layer,
the effective-value resolver, and the kernel's semantic XML comparison.

The matching contract, declared honestly: slides match by their PERMANENT slide id and
top-level shapes within them match by slide-wide shape id, which serves lineage-derived
decks (v4 saved from v3 - the actual use case). Independently built decks can reuse the
same numeric ids, so unrelated decks are outside this matching contract. Deleting an id
and reusing it for a new same-kind object is likewise indistinguishable. One documented
hazard: deleting the max-id slide then adding a new one recycles the id (upstream allocates
max+1), which id-based matching reads as one edited slide - order add-before-delete when
building lineage.

Report-only: no annotated-copy rendering, no visual diffing, no similarity scoring -
those are harness products built ON this report.
"""

from __future__ import annotations

import hashlib
import io
import os
from collections import Counter
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

SCHEMA_NAME = "paper-deck-diff"
SCHEMA_VERSION = 5  # -- was 4: exact paragraph alignment and table structure

_DETAIL_LEVELS = ("structure", "text", "full")


@dataclass(frozen=True)
class BulletShift:
    """One paragraph whose resolved bullet changed between the two decks.

    Bullets live only in formatting, so a list losing its bullets changes no text and no
    field marker: `text_changes` stays empty while the slide visibly loses every glyph.
    This is the facet that reports it.

    Paragraphs use the same exact container-scoped alignment as text and effective-font
    reporting. Repeated ambiguity is skipped rather than settled by paragraph order.

    Fields:

    * ``part`` -- partname of the slide the paragraph lives on.
    * ``shape`` -- stable shape identity plus current display name.
    * ``before_location`` / ``after_location`` -- container-local paragraph locations.
    * ``text`` -- the paragraph's text, which is also half its identity.
    * ``before`` / ``after`` -- the resolved ``bullet``, ``bullet_font`` and ``bullet_size``
      payloads from :func:`pptx.inspect.effective_paragraph_format` on each side.
    """

    part: str
    shape: ShapeRef
    before_location: dict
    after_location: dict
    text: str
    before: dict
    after: dict

    def to_dict(self) -> dict:
        """Return this bullet change as a JSON-ready dict, before and after values included."""
        return {
            "part": self.part,
            "shape": self.shape.to_dict(),
            "before_location": self.before_location,
            "after_location": self.after_location,
            "text": self.text,
            "before": self.before,
            "after": self.after,
        }


@dataclass(frozen=True)
class SlideRef:
    """Identifies a slide by id, position, and title."""
    slide_id: int
    position: int
    title: Optional[str]

    def to_dict(self) -> dict:
        """Return this slide reference as a JSON-ready dict."""
        return {"slide_id": self.slide_id, "position": self.position, "title": self.title}


@dataclass(frozen=True)
class ShapeRef:
    """Identifies a slide-scoped shape by stable id and display name."""

    shape_id: int
    name: str

    def to_dict(self) -> dict:
        """Return this shape reference as a JSON-ready dict."""
        return {"shape_id": self.shape_id, "name": self.name}


@dataclass(frozen=True)
class EffectiveShift:
    """One exactly aligned run whose resolved font changed."""

    part: str
    shape: ShapeRef
    before_location: dict
    after_location: dict
    before_run_index: int
    after_run_index: int
    text: str
    before: dict
    after: dict

    def to_dict(self) -> dict:
        """Return this location-aware shift as a JSON-ready dictionary."""
        return {
            "part": self.part,
            "shape": self.shape.to_dict(),
            "before_location": self.before_location,
            "after_location": self.after_location,
            "before_run_index": self.before_run_index,
            "after_run_index": self.after_run_index,
            "text": self.text,
            "before": self.before,
            "after": self.after,
        }


@dataclass(frozen=True)
class MovedSlide:
    """A slide that kept its identity and changed position."""
    slide_id: int
    from_position: int
    to_position: int

    def to_dict(self) -> dict:
        """Return this move as a JSON-ready dict."""
        return {
            "slide_id": self.slide_id,
            "from_position": self.from_position,
            "to_position": self.to_position,
        }


@dataclass(frozen=True)
class SlideChange:
    """Within-slide deltas for one id-matched slide pair. Empty facets are omitted
    from the payload, so an all-empty change never appears in `slide_changes`."""

    slide_id: int
    shapes_added: Tuple[ShapeRef, ...] = ()
    shapes_removed: Tuple[ShapeRef, ...] = ()
    geometry_changes: Tuple[dict, ...] = ()
    images_replaced: Tuple[ShapeRef, ...] = ()
    chart_data_changes: Tuple[dict, ...] = ()
    table_structure_changes: Tuple[dict, ...] = ()
    text_changes: Tuple[dict, ...] = ()
    notes_change: Optional[dict] = None
    effective_shifts: tuple = ()
    bullet_shifts: tuple = ()

    @property
    def is_empty(self) -> bool:
        """True when nothing on this slide changed. Check it before adding the slide to a report."""
        return not (
            self.shapes_added
            or self.shapes_removed
            or self.geometry_changes
            or self.images_replaced
            or self.chart_data_changes
            or self.table_structure_changes
            or self.text_changes
            or self.notes_change
            or self.effective_shifts
            or self.bullet_shifts
        )

    def to_dict(self) -> dict:
        """Return this slide's changes as a JSON-ready dict, omitting the categories that are empty.
        """
        payload: dict = {"slide_id": self.slide_id}
        if self.shapes_added:
            payload["shapes_added"] = [ref.to_dict() for ref in self.shapes_added]
        if self.shapes_removed:
            payload["shapes_removed"] = [ref.to_dict() for ref in self.shapes_removed]
        if self.geometry_changes:
            payload["geometry_changes"] = [
                _shape_refs_to_dict(change) for change in self.geometry_changes
            ]
        if self.images_replaced:
            payload["images_replaced"] = [ref.to_dict() for ref in self.images_replaced]
        if self.chart_data_changes:
            payload["chart_data_changes"] = [
                _shape_refs_to_dict(change) for change in self.chart_data_changes
            ]
        if self.table_structure_changes:
            payload["table_structure_changes"] = [
                _shape_refs_to_dict(change) for change in self.table_structure_changes
            ]
        if self.text_changes:
            payload["text_changes"] = [_shape_refs_to_dict(change) for change in self.text_changes]
        if self.notes_change is not None:
            payload["notes_change"] = self.notes_change
        if self.effective_shifts:
            payload["effective_shifts"] = [s.to_dict() for s in self.effective_shifts]
        if self.bullet_shifts:
            payload["bullet_shifts"] = [s.to_dict() for s in self.bullet_shifts]
        return payload


def _shape_refs_to_dict(value):
    """Recursively replace |ShapeRef| values with JSON-ready dictionaries."""
    if isinstance(value, ShapeRef):
        return value.to_dict()
    if isinstance(value, dict):
        return {key: _shape_refs_to_dict(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_shape_refs_to_dict(item) for item in value]
    return value


@dataclass(frozen=True)
class DeckDiff:
    """The whole comparison. `.to_dict()` is deterministic and goldenable."""

    detail: str
    slides_added: Tuple[SlideRef, ...] = ()
    slides_removed: Tuple[SlideRef, ...] = ()
    slides_moved: Tuple[MovedSlide, ...] = ()
    slide_changes: Tuple[SlideChange, ...] = field(default_factory=tuple)
    package_changes: tuple = field(default_factory=tuple)

    @property
    def is_empty(self) -> bool:
        """True when the two decks match. Check it to skip reporting a no-op edit."""
        return not (
            self.slides_added
            or self.slides_removed
            or self.slides_moved
            or self.slide_changes
            or self.package_changes
        )

    def to_dict(self) -> dict:
        """Return the whole diff as a JSON-ready dict stamped with its schema and version."""
        return {
            "schema": SCHEMA_NAME,
            "version": SCHEMA_VERSION,
            "detail": self.detail,
            "slides_added": [ref.to_dict() for ref in self.slides_added],
            "slides_removed": [ref.to_dict() for ref in self.slides_removed],
            "slides_moved": [move.to_dict() for move in self.slides_moved],
            "slide_changes": [change.to_dict() for change in self.slide_changes],
            "package_changes": [change.to_dict() for change in self.package_changes],
        }


def diff_decks(path_a, path_b, *, detail: str = "structure") -> DeckDiff:
    """Compare two decks; return the typed |DeckDiff|.

    `path_a`/`path_b` accept a file path, a file-like object, or an already-open
    |Presentation|. `detail`: "structure" (slide add/remove/move, shape add/remove,
    geometry, image replacement, table-grid deltas), "text" (+ exact paragraph deltas,
    chart data per series/category, notes), "full" (+ per-run effective-value shifts via
    the resolver - expensive on large decks, deliberately opt-in).

    When either side is an open |Presentation|, both sides are normalized by serializing
    before the package-level comparison, because a live presentation has no on-disk package
    to read; when both sides are a path or a file-like object, the packages are compared
    exactly.

    Matching is by permanent slide id and, within matched slides, top-level shapes match by
    slide-wide shape id and compatible structural kind. Display names are labels only. Moving a
    shape across a group boundary remains a top-level removal or addition. Independently built
    decks can reuse ids and are outside this contract. A deleted shape id reused by a new
    same-kind shape is likewise indistinguishable without a persistent identifier not present in
    general PPTX files. Slide ids allocate as max+1, so deleting the highest-id slide and then
    adding a new one RECYCLES the id, and this diff will read that delete-plus-add as one edited
    slide - order add-before-delete when producing lineage decks you intend to diff.
    """
    if detail not in _DETAIL_LEVELS:
        raise ValueError(
            "detail must be one of %s, got %r" % (", ".join(_DETAIL_LEVELS), detail)
        )
    from pptx.errors import materialize_slides

    stream_positions = _stream_positions(path_a, path_b)
    try:
        prs_a = _open_deck(path_a)
        prs_b = _open_deck(path_b)
        package_changes = _package_changes(path_a, prs_a, path_b, prs_b)

        order_a = [(slide.slide_id, slide) for slide in materialize_slides(prs_a, "diff_decks")]
        order_b = [(slide.slide_id, slide) for slide in materialize_slides(prs_b, "diff_decks")]
        ids_a = [slide_id for slide_id, _ in order_a]
        ids_b = [slide_id for slide_id, _ in order_b]
        position_a = {slide_id: index for index, (slide_id, _) in enumerate(order_a)}
        position_b = {slide_id: index for index, (slide_id, _) in enumerate(order_b)}
        common = set(ids_a) & set(ids_b)

        slides_added = tuple(
            SlideRef(slide_id, position_b[slide_id], _title_of(slide))
            for slide_id, slide in order_b
            if slide_id not in common
        )
        slides_removed = tuple(
            SlideRef(slide_id, position_a[slide_id], _title_of(slide))
            for slide_id, slide in order_a
            if slide_id not in common
        )

        common_a = [slide_id for slide_id in ids_a if slide_id in common]
        common_b = [slide_id for slide_id in ids_b if slide_id in common]
        stationary = _longest_common_subsequence(common_a, common_b)
        slides_moved = tuple(
            MovedSlide(slide_id, position_a[slide_id], position_b[slide_id])
            for slide_id in common_b  # -- destination order, deterministic
            if slide_id not in stationary
        )

        slide_a_by_id = dict(order_a)
        slide_b_by_id = dict(order_b)
        changes = []
        for slide_id in common_b:
            change = _diff_slide(
                slide_id, slide_a_by_id[slide_id], slide_b_by_id[slide_id], detail
            )
            if not change.is_empty:
                changes.append(change)

        return DeckDiff(
            detail=detail,
            slides_added=slides_added,
            slides_removed=slides_removed,
            slides_moved=slides_moved,
            slide_changes=tuple(changes),
            package_changes=package_changes,
        )
    finally:
        _restore_stream_positions(stream_positions)


def _stream_positions(*sources):
    """Capture supported stream positions so diffing is observationally read-only."""
    from pptx.errors import UnsupportedStructureError
    from pptx.presentation import Presentation as _PresentationProxy

    positions = []
    seen = set()
    for source in sources:
        if isinstance(source, (_PresentationProxy, str, bytes, os.PathLike)):
            continue
        if id(source) in seen:
            continue
        seen.add(id(source))
        if not hasattr(source, "read"):
            raise UnsupportedStructureError(
                "diff_decks refused: inputs must be paths, Presentation objects, or "
                "seekable binary streams"
            )
        try:
            if hasattr(source, "seekable") and not source.seekable():
                raise OSError("stream reports that it is not seekable")
            position = source.tell()
            if type(position) is not int or position < 0:
                raise UnsupportedStructureError(
                    "diff_decks refused: stream tell() must return a nonnegative integer"
                )
            source.seek(position)
        except UnsupportedStructureError:
            raise
        except (AttributeError, OSError, TypeError, ValueError) as exc:
            raise UnsupportedStructureError(
                "diff_decks refused: non-seekable input streams are unsupported (%s)" % exc
            ) from exc
        positions.append((source, position))
    return tuple(positions)


def _restore_stream_positions(stream_positions) -> None:
    """Attempt every stream restore before reporting an aggregate typed failure."""
    from pptx.errors import UnsupportedStructureError

    failures = []
    for index, (stream, position) in enumerate(stream_positions):
        try:
            stream.seek(position)
            if stream.tell() != position:
                raise OSError("stream did not restore to %d" % position)
        except Exception as exc:  # noqa: BLE001 - aggregate all restore failures
            failures.append((index, exc))
    if failures:
        index, exc = failures[0]
        raise UnsupportedStructureError(
            "diff_decks could not restore %d input stream position(s); first failure was "
            "input %d (%s)" % (len(failures), index, type(exc).__name__)
        ) from exc


def _package_changes(source_a, prs_a, source_b, prs_b) -> tuple:
    """Return semantic deltas, reading both sides of the pair the same way."""
    from pptx.package import _diff_maps
    from pptx.presentation import Presentation as _PresentationProxy

    # -- the choice belongs to the PAIR, not to one input: reading one side exactly and
    # -- the other serialized compares two renderings of one document, and reports the
    # -- difference between the renderings as a change to the document
    normalized = any(isinstance(s, _PresentationProxy) for s in (source_a, source_b))
    map_a = _source_package_map(source_a, prs_a, "before", normalized)
    map_b = _source_package_map(source_b, prs_b, "after", normalized)
    return _diff_maps(map_a, map_b, "before", "after").deltas


def _source_package_map(source, prs, label: str, normalized: bool) -> dict:
    """Serialize the side when the pair is normalized; else read the package exactly."""
    from pptx.package import _read_zip_map, _read_zip_map_from_bytes

    if normalized:
        buffer = io.BytesIO()
        # -- serialize through the package, not `Presentation.save`: this is a read for
        # -- comparison, not a publish, so an open `batch()` block must not refuse it
        prs.part.package.save(buffer)
        return _read_zip_map_from_bytes(buffer.getvalue(), "%s normalized input" % label)
    if isinstance(source, (str, bytes, os.PathLike)):
        return _read_zip_map(os.fspath(source))

    position = source.tell()
    try:
        source.seek(0)
        data = _read_stream_package_bytes(source, label)
        return _read_zip_map_from_bytes(data, "%s input" % label)
    finally:
        source.seek(position)


def _read_stream_package_bytes(source, label: str) -> bytes:
    """Read a package stream in fixed-size chunks, with typed I/O failures."""
    from pptx.errors import UnsupportedStructureError

    chunks = []
    try:
        while True:
            chunk = source.read(1024 * 1024)
            if not isinstance(chunk, bytes):
                raise UnsupportedStructureError(
                    "diff_decks refused: %s input stream must return bytes" % label
                )
            if not chunk:
                return b"".join(chunks)
            chunks.append(chunk)
    except UnsupportedStructureError:
        raise
    except (OSError, TypeError, ValueError) as exc:
        raise UnsupportedStructureError(
            "diff_decks refused: cannot read %s input stream (%s)" % (label, exc)
        ) from exc


def _open_deck(source):
    """Return a |Presentation| for a path, stream, or Presentation, refusing typed."""
    from pptx import Presentation
    from pptx.errors import UnsupportedStructureError
    from pptx.presentation import Presentation as _PresentationProxy

    if isinstance(source, _PresentationProxy):
        return source
    try:
        return Presentation(source)
    except Exception as exc:
        raise UnsupportedStructureError(
            "diff_decks refused: %r is not a readable presentation package (%s)"
            % (source, exc)
        ) from exc


# -------------------------------------------------------------------------------- matching


def _longest_common_subsequence(sequence_a: "List[int]", sequence_b: "List[int]") -> set:
    """Ids that kept their relative order; everything else in common reads as moved.

    Between equally-minimal move explanations the DP backtrack is deterministic
    (preferring earlier pairs); move attribution in a genuine tie is arbitrary but
    stable, and the from/to positions in each entry show the actual displacement.
    """
    len_a, len_b = len(sequence_a), len(sequence_b)
    dp = [[0] * (len_b + 1) for _ in range(len_a + 1)]
    for i in range(len_a - 1, -1, -1):
        for j in range(len_b - 1, -1, -1):
            if sequence_a[i] == sequence_b[j]:
                dp[i][j] = dp[i + 1][j + 1] + 1
            else:
                dp[i][j] = max(dp[i + 1][j], dp[i][j + 1])
    kept = set()
    i = j = 0
    while i < len_a and j < len_b:
        if sequence_a[i] == sequence_b[j]:
            kept.add(sequence_a[i])
            i += 1
            j += 1
        elif dp[i + 1][j] >= dp[i][j + 1]:
            i += 1
        else:
            j += 1
    return kept


def _title_of(slide) -> "Optional[str]":
    """Title text of `slide`, or None when it has no title placeholder or the title is blank."""
    title = slide.shapes.title
    if title is None or not title.has_text_frame:
        return None
    return title.text_frame.text or None


# ---------------------------------------------------------------------- within one slide


def _shape_kind(shape) -> str:
    """Return the structural element kind used to decide shape compatibility."""
    from lxml import etree

    localname = etree.QName(shape._element.tag).localname
    if localname != "graphicFrame":
        return localname
    # -- p:graphicFrame is only the generic container. Its graphic-data URI distinguishes
    # -- tables, charts, OLE objects, SmartArt, and other extension payloads whose shape IDs
    # -- can be reused after delete-then-add just like any other top-level shape.
    graphic_data_uri = shape._element.graphicData_uri
    return "graphicFrame[%s]" % (graphic_data_uri or "unknown")


def _iter_shape_tree(shapes):
    """Yield shapes in deterministic tree order for slide-wide id validation."""
    for shape in shapes:
        yield shape
        nested = getattr(shape, "shapes", None)
        if nested is not None:
            for descendant in _iter_shape_tree(nested):
                yield descendant


def _shape_inventory(slide, slide_id: int, side: str) -> dict:
    """Return top-level shapes by id after validating slide-wide id uniqueness."""
    from pptx.errors import UnsupportedStructureError

    candidates: "Dict[int, List[object]]" = {}
    for shape in _iter_shape_tree(slide.shapes):
        candidates.setdefault(shape.shape_id, []).append(shape)
    duplicates = {shape_id: shapes for shape_id, shapes in candidates.items() if len(shapes) > 1}
    if duplicates:
        details = []
        for shape_id in sorted(duplicates):
            shapes = duplicates[shape_id]
            rendered = ", ".join(
                "%s name=%r" % (_shape_kind(shape), shape.name) for shape in shapes
            )
            details.append("id %d: %d candidates [%s]" % (shape_id, len(shapes), rendered))
        raise UnsupportedStructureError(
            "diff_decks refused: %s side of slide %d has duplicate shape IDs (%s); "
            "shape IDs must be unique across the slide"
            % (side, slide_id, "; ".join(details))
        )
    return {shape.shape_id: shape for shape in slide.shapes}


def _shape_ref(shape) -> ShapeRef:
    """Return the stable report reference for `shape` using its current-side name."""
    return ShapeRef(shape_id=shape.shape_id, name=shape.name)


def _diff_slide(slide_id, slide_a, slide_b, detail) -> SlideChange:
    """Compare one id-matched slide pair, gated by `detail`.

    Always: top-level shape add and remove, geometry, image replacement. `"text"` adds chart data,
    text blocks and notes. `"full"` adds per-run effective-value shifts and bullet shifts. Shapes
    match at the top level only, so a shape moved into a group reads as one removal plus one
    addition.
    """
    shapes_a = _shape_inventory(slide_a, slide_id, "before")
    shapes_b = _shape_inventory(slide_b, slide_id, "after")
    ids_a, ids_b = set(shapes_a), set(shapes_b)
    added_ids = ids_b - ids_a
    removed_ids = ids_a - ids_b
    matched_ids = []
    for shape_id in sorted(ids_a & ids_b):
        if _shape_kind(shapes_a[shape_id]) != _shape_kind(shapes_b[shape_id]):
            removed_ids.add(shape_id)
            added_ids.add(shape_id)
        else:
            matched_ids.append(shape_id)

    shapes_added = tuple(_shape_ref(shapes_b[shape_id]) for shape_id in sorted(added_ids))
    shapes_removed = tuple(_shape_ref(shapes_a[shape_id]) for shape_id in sorted(removed_ids))

    geometry_changes: "List[dict]" = []
    images_replaced: "List[ShapeRef]" = []
    chart_changes: "List[dict]" = []
    for shape_id in matched_ids:
        shape_a, shape_b = shapes_a[shape_id], shapes_b[shape_id]
        shape_ref = _shape_ref(shape_b)
        for facet in ("left", "top", "width", "height", "rotation"):
            value_a = getattr(shape_a, facet, None)
            value_b = getattr(shape_b, facet, None)
            if value_a != value_b:
                geometry_changes.append(
                    {"shape": shape_ref, "facet": facet, "before": value_a, "after": value_b}
                )
        image_hash_a = _image_hash(shape_a)
        if image_hash_a is not None and image_hash_a != _image_hash(shape_b):
            images_replaced.append(shape_ref)
        if detail in ("text", "full") and getattr(shape_a, "has_chart", False) and getattr(
            shape_b, "has_chart", False
        ):
            chart_changes.extend(_diff_chart(shape_ref, shape_a.chart, shape_b.chart))

    table_structure_changes = _table_structure_changes(slide_a, slide_b)
    alignments = ()
    text_changes: "List[dict]" = []
    notes_change = None
    if detail in ("text", "full"):
        alignments = _align_slide_text(slide_a, slide_b)
        text_changes = _text_events(alignments, detail, slide_a, slide_b)
        notes_a = _notes_text(slide_a)
        notes_b = _notes_text(slide_b)
        if notes_a != notes_b:
            notes_change = {"before": notes_a, "after": notes_b}

    effective_shifts: tuple = ()
    bullet_shifts: tuple = ()
    if detail == "full":
        effective_shifts = _effective_shifts(alignments)
        bullet_shifts = _bullet_shifts(slide_a, slide_b, alignments)

    return SlideChange(
        slide_id=slide_id,
        shapes_added=shapes_added,
        shapes_removed=shapes_removed,
        geometry_changes=tuple(geometry_changes),
        images_replaced=tuple(images_replaced),
        chart_data_changes=tuple(chart_changes),
        table_structure_changes=table_structure_changes,
        text_changes=tuple(text_changes),
        notes_change=notes_change,
        effective_shifts=effective_shifts,
        bullet_shifts=bullet_shifts,
    )


def _image_hash(shape) -> "Optional[str]":
    """Short content hash of a picture's image bytes.

    None when `shape` is not a picture, so non-pictures never enter the image comparison. An image
    reference that will not load hashes as the sentinel "unreadable" rather than raising mid-diff:
    two unreadable sides compare equal, but unreadable against readable reports as a replacement.
    """
    from pptx.shapes.picture import Picture

    if not isinstance(shape, Picture):
        return None
    try:
        return hashlib.sha256(shape.image.blob).hexdigest()[:16]
    except (KeyError, ValueError):  # pragma: no cover - unloadable image reference
        return "unreadable"


def _diff_chart(key, chart_a, chart_b) -> "List[dict]":
    """Per-series/category data deltas; falls back to an honest opaque flag when the
    chart family cannot be read as category data."""
    from pptx.package import xml_equivalent

    try:
        data_a = _chart_data(chart_a)
        data_b = _chart_data(chart_b)
    except Exception:  # noqa: BLE001 - non-category families: honest opaque comparison
        from lxml import etree

        equivalent = xml_equivalent(
            etree.tostring(chart_a._chartSpace), etree.tostring(chart_b._chartSpace)
        )
        if equivalent:
            return []
        return [
            {
                "chart": key,
                "opaque": True,
                "note": "chart XML differs (family not comparable per-point)",
            }
        ]

    deltas = []
    categories_a, series_a = data_a
    categories_b, series_b = data_b
    if categories_a != categories_b:
        deltas.append(
            {"chart": key, "categories_before": categories_a, "categories_after": categories_b}
        )
    identity_a = [(item[0], item[1], item[2]) for item in series_a]
    identity_b = [(item[0], item[1], item[2]) for item in series_b]
    if identity_a != identity_b:
        deltas.append(
            {
                "chart": key,
                "series_order_before": [item[2] for item in series_a],
                "series_order_after": [item[2] for item in series_b],
            }
        )
    for index in range(max(len(series_a), len(series_b))):
        if index >= len(series_a):
            deltas.append({"chart": key, "series_added": series_b[index][2]})
            continue
        if index >= len(series_b):
            deltas.append({"chart": key, "series_removed": series_a[index][2]})
            continue
        _, _, series_name_a, values_a = series_a[index]
        _, _, series_name_b, values_b = series_b[index]
        series_name = series_name_b
        if series_name_a != series_name_b:
            continue
        length = max(len(values_a), len(values_b))
        for index in range(length):
            value_a = values_a[index] if index < len(values_a) else None
            value_b = values_b[index] if index < len(values_b) else None
            if value_a != value_b:
                category = None
                if index < len(categories_b):
                    category = categories_b[index]
                elif index < len(categories_a):
                    category = categories_a[index]
                deltas.append(
                    {
                        "chart": key,
                        "series": series_name,
                        "category": category,
                        "before": value_a,
                        "after": value_b,
                    }
                )
    return deltas


def _chart_data(chart):
    """Return `(categories, series)` for `chart`, comparable by value across decks.

    `categories` comes from `plots[0]`; `series` spans every plot as `(plot_index, series_index,
    name, values)`. Raises IndexError on a chart with no plots, which is the signal `_diff_chart`
    uses to compare chart XML opaquely instead.
    """
    categories = [str(category) for category in chart.plots[0].categories]
    series = []
    for plot_index, plot in enumerate(chart.plots):
        for series_index, one_series in enumerate(plot.series):
            series.append(
                (plot_index, series_index, one_series.name, tuple(one_series.values))
            )
    return categories, series


@dataclass(frozen=True)
class _ContainerAlignment:
    """Exact paragraph evidence for one stable leaf shape or table frame."""

    shape: ShapeRef
    before: tuple
    after: tuple
    pairs: tuple
    ordered_pairs: tuple
    move_pairs: tuple


def _block_location(block) -> dict:
    """Return the structured container-local location carried by inspection v3."""
    locator = block.anchor.locator
    location = {
        "container": block.container,
        "paragraph_index": locator["paragraph_index"],
    }
    if block.container == "table-cell":
        location["row"] = locator["row"]
        location["column"] = locator["column"]
    return location


def _container_key(block) -> tuple:
    domain = "table" if block.container == "table-cell" else "shape"
    return domain, block.shape_id


def _container_blocks(slide) -> dict:
    """Ordered inspected paragraphs partitioned by stable structural container."""
    from pptx.inspect import inspect_text

    result: dict = {}
    for block in inspect_text(slide).blocks:
        if block.container not in ("shape", "group", "table-cell"):
            continue
        result.setdefault(_container_key(block), []).append(block)
    return {key: tuple(blocks) for key, blocks in result.items()}


def _invariant_lcs_pairs(sequence_a, sequence_b) -> tuple:
    """Return exact pairs present in every optimal LCS, without choosing a tie."""
    rows, cols = len(sequence_a), len(sequence_b)
    next_lengths = [0] * (cols + 1)
    next_forced = [frozenset() for _ in range(cols + 1)]
    for i in range(rows - 1, -1, -1):
        lengths = [0] * (cols + 1)
        forced = [frozenset() for _ in range(cols + 1)]
        for j in range(cols - 1, -1, -1):
            choices = []
            if sequence_a[i] == sequence_b[j]:
                choices.append((1 + next_lengths[j + 1], next_forced[j + 1] | {(i, j)}))
            choices.append((next_lengths[j], next_forced[j]))
            choices.append((lengths[j + 1], forced[j + 1]))
            best = max(length for length, _ in choices)
            winners = [pairs for length, pairs in choices if length == best]
            lengths[j] = best
            forced[j] = frozenset.intersection(*winners)
        next_lengths = lengths
        next_forced = forced
    return tuple(sorted(next_forced[0]))


def _align_slide_text(slide_a, slide_b) -> tuple:
    """Return one exact alignment result shared by all paragraph diff facets."""
    containers_a = _container_blocks(slide_a)
    containers_b = _container_blocks(slide_b)
    shapes_a = {shape.shape_id: shape for shape in _iter_shape_tree(slide_a.shapes)}
    shapes_b = {shape.shape_id: shape for shape in _iter_shape_tree(slide_b.shapes)}
    alignments = []
    for key in sorted(set(containers_a) | set(containers_b)):
        before = containers_a.get(key, ())
        after = containers_b.get(key, ())
        fingerprints_a = tuple(block.anchor.content_hash for block in before)
        fingerprints_b = tuple(block.anchor.content_hash for block in after)
        ordered = _invariant_lcs_pairs(fingerprints_a, fingerprints_b)
        counts_a = Counter(fingerprints_a)
        counts_b = Counter(fingerprints_b)
        unique_a = {
            value: i for i, value in enumerate(fingerprints_a) if counts_a[value] == 1
        }
        unique_b = {
            value: i for i, value in enumerate(fingerprints_b) if counts_b[value] == 1
        }
        global_unique = {
            (index_a, unique_b[value]) for value, index_a in unique_a.items() if value in unique_b
        }
        pairs = tuple(sorted(set(ordered) | global_unique))
        reference = shapes_b.get(key[1]) or shapes_a[key[1]]
        alignments.append(
            _ContainerAlignment(
                _shape_ref(reference),
                before,
                after,
                pairs,
                ordered,
                tuple(sorted(global_unique - set(ordered))),
            )
        )
    return tuple(alignments)


def _single_text_event(kind, alignment, block_a, block_b) -> dict:
    event = {
        "kind": kind,
        "shape": alignment.shape,
        "before_location": _block_location(block_a) if block_a is not None else None,
        "after_location": _block_location(block_b) if block_b is not None else None,
        "before": block_a.text if block_a is not None else None,
        "after": block_b.text if block_b is not None else None,
    }
    fields_a, fields_b = _field_markers(block_a), _field_markers(block_b)
    if fields_a != fields_b:
        event["field_types_before"] = [field_type for _, field_type in fields_a]
        event["field_types_after"] = [field_type for _, field_type in fields_b]
        event["fields_before"] = [
            {"offset": offset, "type": field_type} for offset, field_type in fields_a
        ]
        event["fields_after"] = [
            {"offset": offset, "type": field_type} for offset, field_type in fields_b
        ]
    return event


def _ambiguity_event(alignment, before, after) -> dict:
    def candidates(blocks):
        return [{"location": _block_location(block), "text": block.text} for block in blocks]

    return {
        "kind": "ambiguity",
        "shape": alignment.shape,
        "before_location": None,
        "after_location": None,
        "before_candidates": candidates(before),
        "after_candidates": candidates(after),
        "reason": "repeated exact paragraph fingerprints do not establish correspondence",
    }


def _observable_sequence(blocks, detail, slide) -> tuple:
    values = []
    for block in blocks:
        item = [block.text, _field_markers(block)]
        if detail == "full" and block.container != "table-cell":
            item.append(tuple((run.text, run.field_type, run.font.to_dict()) for run in block.runs))
            bullet = _bullet_payload(_paragraph_by_block(slide, block))
            item.append(bullet[1] if bullet is not None else None)
        values.append(tuple(item))
    return tuple(values)


def _text_events(alignments, detail, slide_a, slide_b) -> "List[dict]":
    """Classify exact pairs and conservative unmatched regions into v5 text events."""
    events = []
    for alignment in alignments:
        paired_a = {i for i, _ in alignment.pairs}
        paired_b = {j for _, j in alignment.pairs}
        for index_a, index_b in alignment.move_pairs:
            events.append(
                _single_text_event(
                    "move", alignment, alignment.before[index_a], alignment.after[index_b]
                )
            )
        anchors = (
            ((-1, -1),) + alignment.ordered_pairs + ((len(alignment.before), len(alignment.after)),)
        )
        for (left_a, left_b), (right_a, right_b) in zip(anchors, anchors[1:]):
            before = tuple(
                alignment.before[i] for i in range(left_a + 1, right_a) if i not in paired_a
            )
            after = tuple(
                alignment.after[j] for j in range(left_b + 1, right_b) if j not in paired_b
            )
            if not before and not after:
                continue
            if not after:
                events.extend(
                    _single_text_event("deletion", alignment, block, None) for block in before
                )
            elif not before:
                events.extend(
                    _single_text_event("insertion", alignment, None, block) for block in after
                )
            elif len(before) == len(after) == 1:
                events.append(_single_text_event("replacement", alignment, before[0], after[0]))
            elif _observable_sequence(before, detail, slide_a) != _observable_sequence(
                after, detail, slide_b
            ):
                events.append(_ambiguity_event(alignment, before, after))
        for index_a, index_b in alignment.pairs:
            if (index_a, index_b) in alignment.move_pairs:
                continue
            block_a, block_b = alignment.before[index_a], alignment.after[index_b]
            if block_a.text != block_b.text or _field_markers(block_a) != _field_markers(block_b):
                events.append(_single_text_event("replacement", alignment, block_a, block_b))
    return events


def _field_markers(block) -> tuple:
    """Return field types positioned in visible literal text, independent of run splitting."""
    if block is None:
        return ()
    offset = 0
    markers = []
    for run in block.runs:
        if run.field_type is not None:
            markers.append((offset, run.field_type))
        else:
            offset += len(run.text)
    return tuple(markers)


def _run_identity(run) -> tuple:
    return ("field", run.field_type) if run.field_type is not None else ("text", run.text)


def _font_values(font) -> tuple:
    return (
        font.size.value,
        font.name.value,
        font.color_rgb.value,
        font.bold.value if font.bold is not None else None,
        font.italic.value if font.italic is not None else None,
        font.underline.value if font.underline is not None else None,
    )


def _effective_shifts(alignments) -> "Tuple[EffectiveShift, ...]":
    """Formatting changes for uniquely identified runs in exactly aligned paragraphs."""
    shifts = []
    for alignment in alignments:
        for index_a, index_b in alignment.pairs:
            block_a, block_b = alignment.before[index_a], alignment.after[index_b]
            if block_a.container == "table-cell" or block_b.container == "table-cell":
                continue
            runs_a = [(i, run) for i, run in enumerate(block_a.runs) if run.text != "\v"]
            runs_b = [(i, run) for i, run in enumerate(block_b.runs) if run.text != "\v"]
            identities_a = [_run_identity(run) for _, run in runs_a]
            identities_b = [_run_identity(run) for _, run in runs_b]
            by_identity_a = {
                identity: item
                for identity, item in zip(identities_a, runs_a)
                if identity[1] and identities_a.count(identity) == 1
            }
            by_identity_b = {
                identity: item
                for identity, item in zip(identities_b, runs_b)
                if identity[1] and identities_b.count(identity) == 1
            }
            for identity in sorted(set(by_identity_a) & set(by_identity_b)):
                run_index_a, run_a = by_identity_a[identity]
                run_index_b, run_b = by_identity_b[identity]
                if _font_values(run_a.font) == _font_values(run_b.font):
                    continue
                shifts.append(
                    EffectiveShift(
                        part=block_a.anchor.part,
                        shape=alignment.shape,
                        before_location=_block_location(block_a),
                        after_location=_block_location(block_b),
                        before_run_index=run_index_a,
                        after_run_index=run_index_b,
                        text=run_b.text,
                        before=run_a.font.to_dict(),
                        after=run_b.font.to_dict(),
                    )
                )
    return tuple(shifts)


def _paragraph_by_block(slide, block):
    """Resolve an inspection-v3 ordinary/group block to its paragraph for bullet inspection."""
    locator = block.anchor.locator
    shape = next(
        shape for shape in _iter_shape_tree(slide.shapes) if shape.shape_id == locator["shape_id"]
    )
    return shape.text_frame.paragraphs[locator["paragraph_index"]]


def _bullet_payload(paragraph):
    from pptx.errors import UnsupportedStructureError
    from pptx.inspect import effective_paragraph_format

    try:
        fmt = effective_paragraph_format(paragraph)
    except UnsupportedStructureError:
        return None
    values = (
        fmt.bullet.type,
        fmt.bullet.char,
        fmt.bullet.number_scheme,
        fmt.bullet.start_at,
        fmt.bullet_font.value,
        fmt.bullet_size.value,
    )
    payload = {
        "bullet": fmt.bullet.to_dict(),
        "bullet_font": fmt.bullet_font.to_dict(),
        "bullet_size": fmt.bullet_size.to_dict(),
    }
    return values, payload


def _bullet_shifts(slide_a, slide_b, alignments) -> "Tuple[BulletShift, ...]":
    """Bullet changes driven by the same exact paragraph pairs as text and fonts."""
    shifts = []
    for alignment in alignments:
        for index_a, index_b in alignment.pairs:
            block_a, block_b = alignment.before[index_a], alignment.after[index_b]
            if block_a.container == "table-cell" or block_b.container == "table-cell":
                continue
            before = _bullet_payload(_paragraph_by_block(slide_a, block_a))
            after = _bullet_payload(_paragraph_by_block(slide_b, block_b))
            if before is None or after is None or before[0] == after[0]:
                continue
            shifts.append(
                BulletShift(
                    part=block_a.anchor.part,
                    shape=alignment.shape,
                    before_location=_block_location(block_a),
                    after_location=_block_location(block_b),
                    text=block_b.text,
                    before=before[1],
                    after=after[1],
                )
            )
    return tuple(shifts)


def _table_signature(table) -> tuple:
    """Exact cell paragraph fingerprints arranged as a stable row/column grid."""
    from pptx.inspect import _paragraph_fingerprint

    return tuple(
        tuple(
            tuple(_paragraph_fingerprint(paragraph._p) for paragraph in cell.text_frame.paragraphs)
            for cell in row.cells
        )
        for row in table.rows
    )


def _insertion_candidates(before, after, count) -> list:
    return [
        index
        for index in range(len(after) - count + 1)
        if after[:index] + after[index + count :] == before
    ]


def _dimension_change(before, after, axis) -> dict:
    delta = len(after) - len(before)
    kind = "insertion" if delta > 0 else "deletion"
    count = abs(delta)
    candidates = (
        _insertion_candidates(before, after, count)
        if delta > 0
        else _insertion_candidates(after, before, count)
    )
    result = {"kind": kind, "count": count}
    if len(candidates) == 1:
        result["index"] = candidates[0]
    else:
        result["ambiguous"] = True
        result["candidate_indexes"] = candidates
        result["axis"] = axis
    return result


def _table_structure_changes(slide_a, slide_b) -> tuple:
    """Return exact table-grid evidence for stable table-frame IDs."""
    shapes_a = {shape.shape_id: shape for shape in _iter_shape_tree(slide_a.shapes)}
    shapes_b = {shape.shape_id: shape for shape in _iter_shape_tree(slide_b.shapes)}
    changes = []
    for shape_id in sorted(set(shapes_a) & set(shapes_b)):
        shape_a, shape_b = shapes_a[shape_id], shapes_b[shape_id]
        if not getattr(shape_a, "has_table", False) or not getattr(shape_b, "has_table", False):
            continue
        table_a, table_b = shape_a.table, shape_b.table
        rows_a, cols_a = len(table_a.rows), len(table_a.columns)
        rows_b, cols_b = len(table_b.rows), len(table_b.columns)
        if (rows_a, cols_a) == (rows_b, cols_b):
            continue
        change = {
            "table": _shape_ref(shape_b),
            "before": {"rows": rows_a, "columns": cols_a},
            "after": {"rows": rows_b, "columns": cols_b},
        }
        grid_a, grid_b = _table_signature(table_a), _table_signature(table_b)
        if cols_a == cols_b and rows_a != rows_b:
            change["row_change"] = _dimension_change(grid_a, grid_b, "row")
        elif rows_a == rows_b and cols_a != cols_b:
            columns_a = tuple(tuple(row[column] for row in grid_a) for column in range(cols_a))
            columns_b = tuple(tuple(row[column] for row in grid_b) for column in range(cols_b))
            change["column_change"] = _dimension_change(columns_a, columns_b, "column")
        changes.append(change)
    return tuple(changes)


def _notes_text(slide) -> "Optional[str]":
    """Notes text of `slide`, or None when it carries no notes slide."""
    if not slide.has_notes_slide:
        return None
    return slide.read_notes_text()
