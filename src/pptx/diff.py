"""Deck diff - the verification mirror (paper-pptx addition).

PPTX carries no revision markup: a deck is amnesiac about its own history, so "what
changed between v3 and v4?" has no programmable answer anywhere in the ecosystem. This
provides it as a typed report beside the file - the deck-format analogue of a
redline - assembled from the permanent slide ids, the visibility-complete text layer,
the effective-value resolver, and the kernel's semantic XML comparison.

The matching contract, declared honestly: slides match by their PERMANENT slide id,
which serves lineage-derived decks (v4 saved from v3 - the actual use case). Independently
built decks can reuse the same numeric ids, so unrelated decks are outside this matching
contract. One documented hazard: deleting the max-id slide then adding a new one recycles
the id (upstream allocates max+1), which id-based matching reads as one edited slide -
order add-before-delete when building lineage.

Report-only: no annotated-copy rendering, no visual diffing, no similarity scoring -
those are harness products built ON this report.
"""

from __future__ import annotations

import hashlib
import io
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

SCHEMA_NAME = "paper-deck-diff"
SCHEMA_VERSION = 3  # -- was 2: bullet_shifts added

_DETAIL_LEVELS = ("structure", "text", "full")


@dataclass(frozen=True)
class BulletShift:
    """One paragraph whose resolved bullet changed between the two decks.

    Bullets live only in formatting, so a list losing its bullets changes no text and no
    field marker: `text_changes` stays empty while the slide visibly loses every glyph.
    This is the facet that reports it.

    Paragraphs are identified by (shape_id, paragraph text) rather than by index, matching
    `pptx.rebind._resolution_state(align_by_content=True)`: inserting a paragraph shifts
    every later index and would otherwise pair unrelated paragraphs. Paragraphs whose text
    repeats within their shape are skipped rather than guessed at.

    Fields:

    * ``part`` -- partname of the slide the paragraph lives on.
    * ``shape_id`` -- id of the shape containing the paragraph.
    * ``paragraph_index`` -- 0-based paragraph index within that shape, for legibility.
    * ``text`` -- the paragraph's text, which is also half its identity.
    * ``before`` / ``after`` -- the resolved ``bullet``, ``bullet_font`` and ``bullet_size``
      payloads from :func:`pptx.inspect.effective_paragraph_format` on each side.
    """

    part: str
    shape_id: int
    paragraph_index: int
    text: str
    before: dict
    after: dict

    def to_dict(self) -> dict:
        return {
            "part": self.part,
            "shape_id": self.shape_id,
            "paragraph_index": self.paragraph_index,
            "text": self.text,
            "before": self.before,
            "after": self.after,
        }


@dataclass(frozen=True)
class SlideRef:
    slide_id: int
    position: int
    title: Optional[str]

    def to_dict(self) -> dict:
        return {"slide_id": self.slide_id, "position": self.position, "title": self.title}


@dataclass(frozen=True)
class MovedSlide:
    slide_id: int
    from_position: int
    to_position: int

    def to_dict(self) -> dict:
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
    shapes_added: Tuple[str, ...] = ()
    shapes_removed: Tuple[str, ...] = ()
    geometry_changes: Tuple[dict, ...] = ()
    images_replaced: Tuple[str, ...] = ()
    chart_data_changes: Tuple[dict, ...] = ()
    text_changes: Tuple[dict, ...] = ()
    notes_change: Optional[dict] = None
    effective_shifts: tuple = ()
    bullet_shifts: tuple = ()

    @property
    def is_empty(self) -> bool:
        return not (
            self.shapes_added
            or self.shapes_removed
            or self.geometry_changes
            or self.images_replaced
            or self.chart_data_changes
            or self.text_changes
            or self.notes_change
            or self.effective_shifts
            or self.bullet_shifts
        )

    def to_dict(self) -> dict:
        payload: dict = {"slide_id": self.slide_id}
        if self.shapes_added:
            payload["shapes_added"] = list(self.shapes_added)
        if self.shapes_removed:
            payload["shapes_removed"] = list(self.shapes_removed)
        if self.geometry_changes:
            payload["geometry_changes"] = list(self.geometry_changes)
        if self.images_replaced:
            payload["images_replaced"] = list(self.images_replaced)
        if self.chart_data_changes:
            payload["chart_data_changes"] = list(self.chart_data_changes)
        if self.text_changes:
            payload["text_changes"] = list(self.text_changes)
        if self.notes_change is not None:
            payload["notes_change"] = self.notes_change
        if self.effective_shifts:
            payload["effective_shifts"] = [s.to_dict() for s in self.effective_shifts]
        if self.bullet_shifts:
            payload["bullet_shifts"] = [s.to_dict() for s in self.bullet_shifts]
        return payload


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
        return not (
            self.slides_added
            or self.slides_removed
            or self.slides_moved
            or self.slide_changes
            or self.package_changes
        )

    def to_dict(self) -> dict:
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
    geometry, image replacement), "text" (+ text-block deltas, chart data per
    series/category, notes), "full" (+ per-run effective-value shifts via the resolver
    - expensive on large decks, deliberately opt-in).

    When either side is an open |Presentation|, both sides are normalized by serializing
    before the package-level comparison, because a live presentation has no on-disk package
    to read; when both sides are a path or a file-like object, the packages are compared
    exactly.

    Matching is by permanent slide id and is intended for lineage-derived decks.
    Independently built decks can reuse ids and are outside this contract. One declared
    hazard: slide ids allocate as max+1, so deleting the highest-id slide and then adding
    a new one RECYCLES the id, and this diff will read that delete-plus-add as one edited
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
    title = slide.shapes.title
    if title is None or not title.has_text_frame:
        return None
    return title.text_frame.text or None


# ---------------------------------------------------------------------- within one slide


def _shape_keys(slide) -> dict:
    """Deterministic shape key map: unique name, else `<kind>#<ordinal>` (declared
    fallback for unnamed shapes; duplicate names get the synthetic key too - honest
    ambiguity handling rather than a guess)."""
    from lxml import etree

    shapes = list(slide.shapes)
    names = [shape.name for shape in shapes]
    keyed = {}
    kind_counters: "Dict[str, int]" = {}
    for shape, name in zip(shapes, names):
        kind = etree.QName(shape._element.tag).localname
        ordinal = kind_counters.get(kind, 0)
        kind_counters[kind] = ordinal + 1
        if name and names.count(name) == 1:
            keyed[("name", name)] = shape
        else:
            keyed[("fallback", kind, ordinal)] = shape
    return keyed


def _shape_key_text(key) -> str:
    return key[1] if key[0] == "name" else "%s#%d" % (key[1], key[2])


def _diff_slide(slide_id, slide_a, slide_b, detail) -> SlideChange:
    shapes_a = _shape_keys(slide_a)
    shapes_b = _shape_keys(slide_b)
    shapes_added = tuple(_shape_key_text(key) for key in sorted(set(shapes_b) - set(shapes_a)))
    shapes_removed = tuple(_shape_key_text(key) for key in sorted(set(shapes_a) - set(shapes_b)))

    geometry_changes: "List[dict]" = []
    images_replaced: "List[str]" = []
    chart_changes: "List[dict]" = []
    for key in sorted(set(shapes_a) & set(shapes_b)):
        shape_a, shape_b = shapes_a[key], shapes_b[key]
        key_text = _shape_key_text(key)
        for facet in ("left", "top", "width", "height", "rotation"):
            value_a = getattr(shape_a, facet, None)
            value_b = getattr(shape_b, facet, None)
            if value_a != value_b:
                geometry_changes.append(
                    {"shape": key_text, "facet": facet, "before": value_a, "after": value_b}
                )
        if _image_hash(shape_a) is not None and _image_hash(shape_a) != _image_hash(shape_b):
            images_replaced.append(key_text)
        if detail in ("text", "full") and getattr(shape_a, "has_chart", False) and getattr(
            shape_b, "has_chart", False
        ):
            chart_changes.extend(_diff_chart(key_text, shape_a.chart, shape_b.chart))

    text_changes: "List[dict]" = []
    notes_change = None
    if detail in ("text", "full"):
        text_changes = _diff_text(slide_a, slide_b)
        notes_a = _notes_text(slide_a)
        notes_b = _notes_text(slide_b)
        if notes_a != notes_b:
            notes_change = {"before": notes_a, "after": notes_b}

    effective_shifts: tuple = ()
    bullet_shifts: tuple = ()
    if detail == "full":
        from pptx.rebind import _resolution_state, _shifts_between

        effective_shifts = _shifts_between(
            _resolution_state(slide_a, align_by_content=True),
            _resolution_state(slide_b, align_by_content=True),
        )
        bullet_shifts = _bullet_shifts_between(
            _bullet_state(slide_a), _bullet_state(slide_b)
        )

    return SlideChange(
        slide_id=slide_id,
        shapes_added=shapes_added,
        shapes_removed=shapes_removed,
        geometry_changes=tuple(geometry_changes),
        images_replaced=tuple(images_replaced),
        chart_data_changes=tuple(chart_changes),
        text_changes=tuple(text_changes),
        notes_change=notes_change,
        effective_shifts=effective_shifts,
        bullet_shifts=bullet_shifts,
    )


def _image_hash(shape) -> "Optional[str]":
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
    categories = [str(category) for category in chart.plots[0].categories]
    series = []
    for plot_index, plot in enumerate(chart.plots):
        for series_index, one_series in enumerate(plot.series):
            series.append(
                (plot_index, series_index, one_series.name, tuple(one_series.values))
            )
    return categories, series


def _text_blocks_by_stable_key(slide) -> dict:
    """Blocks keyed (shape_id, block_ordinal-within-shape) - stable across shape
    adds/removes elsewhere on the slide, unlike the slide-global block index (which
    would misattribute every edit below an added or removed shape)."""
    from pptx.inspect import inspect_text

    keyed = {}
    block_ordinals: "Dict[int, int]" = {}
    for block in inspect_text(slide).blocks:
        ordinal = block_ordinals.get(block.shape_id, 0)
        block_ordinals[block.shape_id] = ordinal + 1
        keyed[(block.shape_id, ordinal)] = block
    return keyed


def _diff_text(slide_a, slide_b) -> "List[dict]":
    blocks_a = _text_blocks_by_stable_key(slide_a)
    blocks_b = _text_blocks_by_stable_key(slide_b)
    changes = []
    for key in sorted(set(blocks_a) | set(blocks_b)):
        block_a = blocks_a.get(key)
        block_b = blocks_b.get(key)
        text_a = block_a.text if block_a is not None else None
        text_b = block_b.text if block_b is not None else None
        fields_a = _field_markers(block_a)
        fields_b = _field_markers(block_b)
        if text_a != text_b or fields_a != fields_b:
            reference = block_b if block_b is not None else block_a
            change = {
                "shape_id": key[0],
                "shape_name": reference.shape_name,
                "block_ordinal": key[1],
                "before": text_a,
                "after": text_b,
            }
            if fields_a != fields_b:
                change["field_types_before"] = [field_type for _, field_type in fields_a]
                change["field_types_after"] = [field_type for _, field_type in fields_b]
                change["fields_before"] = [
                    {"offset": offset, "type": field_type} for offset, field_type in fields_a
                ]
                change["fields_after"] = [
                    {"offset": offset, "type": field_type} for offset, field_type in fields_b
                ]
            changes.append(change)
    return changes


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


def _iter_text_shapes(shapes):
    """Yield every shape carrying a text frame, descending into groups."""
    for shape in shapes:
        if getattr(shape, "has_text_frame", False):
            yield shape
        nested = getattr(shape, "shapes", None)
        if nested is not None:
            for inner in _iter_text_shapes(nested):
                yield inner


def _bullet_state(slide) -> dict:
    """(comparable bullet values, payload) per paragraph, keyed (shape_id, paragraph text).

    Content-keyed for the same reason `pptx.rebind._resolution_state(align_by_content=True)`
    is: a paragraph inserted above shifts every later index, and index keys would then pair
    unrelated paragraphs and report bullet changes that never happened. Paragraphs whose text
    repeats within their shape carry no unique identity and are skipped rather than guessed at.

    Paragraphs the resolver refuses -- table cells, which are not `p:sp` -- are skipped too. A
    refusal here is a scope boundary, not a diff failure.
    """
    from pptx.errors import UnsupportedStructureError
    from pptx.inspect import effective_paragraph_format

    state: dict = {}
    partname = str(slide.part.partname)
    for shape in _iter_text_shapes(slide.shapes):
        paragraphs = list(shape.text_frame.paragraphs)
        texts = [paragraph.text for paragraph in paragraphs]
        for index, paragraph in enumerate(paragraphs):
            text = texts[index]
            if not text or texts.count(text) != 1:
                continue
            try:
                fmt = effective_paragraph_format(paragraph)
            except UnsupportedStructureError:
                continue
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
            state[(shape.shape_id, text)] = (values, text, partname, payload, index)
    return state


def _bullet_shifts_between(before_state, after_state) -> "Tuple[BulletShift, ...]":
    shifts = []
    for key in sorted(set(before_state) & set(after_state), key=lambda k: (k[0], k[1])):
        before_values, text, partname, before_payload, index = before_state[key]
        after_values, _, _, after_payload, _ = after_state[key]
        if before_values != after_values:
            shifts.append(
                BulletShift(
                    part=partname,
                    shape_id=key[0],
                    paragraph_index=index,
                    text=text,
                    before=before_payload,
                    after=after_payload,
                )
            )
    return tuple(shifts)


def _notes_text(slide) -> "Optional[str]":
    if not slide.has_notes_slide:
        return None
    return slide.read_notes_text()
