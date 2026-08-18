"""Layout rebind - the template-migration primitive (paper-pptx addition).

Moving a slide between layouts without corrupting placeholder inheritance is the
package-level piece of template migration; deciding what maps where is the caller's job
(the workflow stays in the harness). `slide.rebind_layout(...)` reconciles the slide's
placeholders against the target layout (auto-matching by type then type-family, explicit
map overrides), retargets the slide->layout relationship, and returns the REQUIRED
`RebindReport`: the effective-value resolver runs before and after, and every text run
whose *resolved values* changed appears in the report - appearance shifts are reported,
never silent.

Orphans - source placeholders with no destination match - follow `orphan_policy`:
"refuse" (default; typed, atomic) or "bake": the placeholder becomes a free shape with
its effective formatting written explicitly (geometry materialized from inheritance, each
run's resolved size/name/color/emphasis made local) so the text keeps its look without a
binding. Bake writes only values the resolver can localize exactly. Color-transform
recipes are preserved; unsupported script-specific formatting refuses before mutation.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import TYPE_CHECKING, Dict, List, Optional, Tuple

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.errors import UnsupportedStructureError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu

if TYPE_CHECKING:
    from pptx.slide import Slide, SlideLayout

__all__ = [
    "RebindReport",
    "RunShift",
    "SCHEMA_NAME",
    "SCHEMA_VERSION",
    "rebind_layout",
]

SCHEMA_NAME = "paper-rebind-report"
SCHEMA_VERSION = 1

# -- auto-match falls back from exact type to these interchangeable families (what
# -- PowerPoint itself does when switching a slide between title and content layouts)
_TYPE_FAMILIES = (
    frozenset({PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}),
    frozenset({PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE}),
)

_MC_ALTERNATE_CONTENT = (
    "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
)


@dataclass(frozen=True)
class RunShift:
    """One run whose resolved effective values changed across the rebind.

    Runs are identified by (shape_id, block_ordinal-within-shape, run_index) - a STABLE
    key that survives shapes being added or removed elsewhere on the slide. Keying by
    the slide-global block index would pair unrelated runs the moment any earlier shape
    disappears.

    Fields:

    * ``part`` -- partname of the slide the run lives on.
    * ``shape_id`` -- id of the shape containing the run.
    * ``block_ordinal`` -- 0-based paragraph-block index within that shape.
    * ``run_index`` -- 0-based run index within the block.
    * ``text`` -- the run's text (for human legibility of the shift).
    * ``before`` / ``after`` -- the run's resolved effective-font payload
      (``EffectiveFont.to_dict()``) before and after the rebind.
    """

    part: str
    shape_id: int
    block_ordinal: int
    run_index: int
    text: str
    before: dict
    after: dict

    def to_dict(self) -> dict:
        return {
            "part": self.part,
            "shape_id": self.shape_id,
            "block_ordinal": self.block_ordinal,
            "run_index": self.run_index,
            "text": self.text,
            "before": self.before,
            "after": self.after,
        }


@dataclass(frozen=True)
class RebindReport:
    """What one rebind did: the mapping used, orphan handling, and every resolution shift.

    Fields:

    * ``source_layout`` / ``source_layout_name`` -- partname and display name of the
      layout the slide was bound to before.
    * ``target_layout`` / ``target_layout_name`` -- partname and display name of the
      layout the slide is bound to after.
    * ``placeholder_map_used`` -- the resolved ``(source_idx, target_idx)`` pairs; a
      ``target_idx`` of ``None`` marks a source placeholder that was orphaned.
    * ``baked_orphans`` -- names of orphan placeholders converted to free shapes (only
      populated under ``orphan_policy="bake"``).
    * ``run_shifts`` -- one `RunShift` for every run whose resolved effective
      values changed; empty when the rebind preserved appearance exactly.
    """

    source_layout: str
    source_layout_name: str
    target_layout: str
    target_layout_name: str
    placeholder_map_used: Tuple[Tuple[int, Optional[int]], ...]
    baked_orphans: Tuple[str, ...]
    run_shifts: Tuple[RunShift, ...]

    def to_dict(self) -> dict:
        return {
            "schema": SCHEMA_NAME,
            "version": SCHEMA_VERSION,
            "source_layout": self.source_layout,
            "source_layout_name": self.source_layout_name,
            "target_layout": self.target_layout,
            "target_layout_name": self.target_layout_name,
            "placeholder_map_used": [
                {"source_idx": src, "target_idx": tgt} for src, tgt in self.placeholder_map_used
            ],
            "baked_orphans": list(self.baked_orphans),
            "run_shifts": [shift.to_dict() for shift in self.run_shifts],
        }


def rebind_layout(
    slide: "Slide",
    target_layout: "SlideLayout",
    *,
    placeholder_map="auto",
    orphan_policy: str = "refuse",
) -> RebindReport:
    """Rebind `slide` to `target_layout`; return the required shift report."""
    from pptx._transaction import PackageTransaction
    from pptx.slide import (
        SlideLayout,
        _require_layout_enrolled,
        _require_slide_enrolled,
    )

    if not isinstance(target_layout, SlideLayout):
        raise ValueError("target_layout must be a SlideLayout, got %r" % (target_layout,))
    if target_layout.part.package is not slide.part.package:
        raise ValueError(
            "target_layout belongs to a different presentation; rebind is same-package "
            "only (cross-package composition is import_slide's job)"
        )
    if orphan_policy not in ("refuse", "bake"):
        raise ValueError("orphan_policy must be 'refuse' or 'bake', got %r" % (orphan_policy,))
    _require_slide_enrolled(slide)
    _require_layout_enrolled(target_layout)
    with PackageTransaction(slide.part.package, slide, target_layout):
        return _rebind_layout_impl(
            slide,
            target_layout,
            placeholder_map=placeholder_map,
            orphan_policy=orphan_policy,
        )


def _rebind_layout_impl(slide, target_layout, *, placeholder_map, orphan_policy) -> RebindReport:
    """Perform a rebind inside the caller's package transaction."""
    source_layout = slide.slide_layout
    from pptx.slide import _require_layout_enrolled

    _require_layout_enrolled(source_layout, argument="source layout")
    if target_layout.part is source_layout.part:
        raise ValueError("target_layout is already this slide's layout")
    if any(True for _ in slide._element.spTree.iterchildren(_MC_ALTERNATE_CONTENT)):
        raise UnsupportedStructureError(
            "slide contains mc:AlternateContent; shapes inside it are invisible to "
            "placeholder reconciliation, so rebinding could silently orphan them"
        )

    slide_phs = [shape for shape in slide.shapes if shape.is_placeholder]
    source_idx_of = {id(shape): shape.element.ph_idx for shape in slide_phs}
    mapping = _compute_mapping(slide_phs, target_layout, placeholder_map)

    orphan_shapes = [shape for shape in slide_phs if mapping[source_idx_of[id(shape)]] is None]
    if orphan_shapes and orphan_policy == "refuse":
        raise UnsupportedStructureError(
            "no placeholder in layout %r matches: %s; pass an explicit placeholder_map "
            "or orphan_policy='bake'"
            % (
                target_layout.name or str(target_layout.part.partname),
                ", ".join(
                    "%r (type %s, idx %d)"
                    % (shape.name, shape.element.ph_type.name, shape.element.ph_idx)
                    for shape in orphan_shapes
                ),
            )
        )

    # -- bake preconditions, validated for EVERY orphan before any write ------------------
    for shape in orphan_shapes:
        if shape.element.findall(".//" + qn("a:fld")):
            raise UnsupportedStructureError(
                "placeholder %r contains a field (a:fld); a baked field would freeze "
                "volatile content - remove the placeholder or map it explicitly" % shape.name
            )
        if any(
            getattr(shape, attribute) is None for attribute in ("left", "top", "width", "height")
        ):
            raise UnsupportedStructureError(
                "placeholder %r has no resolvable geometry (no a:xfrm anywhere in its "
                "inheritance chain); baking would place it unpredictably" % shape.name
            )
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    effective = run.effective_font()
                    unsafe_name = any(
                        "non-Latin" in step.detail for step in effective.name.provenance
                    )
                    unsafe_color = (
                        any(
                            "unapplied transforms" in step.detail
                            for step in effective.color_rgb.provenance
                        )
                        and effective.color_rgb.bake_color_xml is None
                    )
                    if unsafe_name or unsafe_color:
                        raise UnsupportedStructureError(
                            "placeholder %r has unresolved text formatting; baking would "
                            "change its appearance" % shape.name
                        )

    before_state = _resolution_state(slide)
    slide_part = slide.part
    element_snapshot = copy.deepcopy(slide._element)
    rels_snapshot = dict(slide_part.rels._rels)

    try:
        # -- mutate: bake orphans first (they must resolve under the SOURCE layout) --------
        baked_names = []
        for shape in orphan_shapes:
            _bake_placeholder(shape)
            baked_names.append(shape.name)

        for shape in slide_phs:
            target = mapping[source_idx_of[id(shape)]]
            if target is None:
                continue  # -- already baked to a free shape above
            target_type, target_idx = target
            ph = shape.element.ph
            ph.type = target_type
            ph.idx = target_idx

        for rId, rel in list(slide_part.rels.items()):
            if not rel.is_external and rel.reltype == RT.SLIDE_LAYOUT:
                slide_part.drop_rel(rId)
        slide_part.relate_to(target_layout.part, RT.SLIDE_LAYOUT)

        after_state = _resolution_state(slide)
    except BaseException:
        _restore_element(slide._element, element_snapshot)
        slide_part.rels._rels.clear()
        slide_part.rels._rels.update(rels_snapshot)
        raise

    return RebindReport(
        source_layout=str(source_layout.part.partname),
        source_layout_name=source_layout.name,
        target_layout=str(target_layout.part.partname),
        target_layout_name=target_layout.name,
        placeholder_map_used=tuple(
            (source_idx, target[1] if target else None)
            for source_idx, target in sorted(mapping.items())
        ),
        baked_orphans=tuple(baked_names),
        run_shifts=_shifts_between(before_state, after_state),
    )


# ------------------------------------------------------------------------------- matching


def _layout_placeholder_slots(target_layout) -> "List[Tuple[PP_PLACEHOLDER, int]]":
    slots = []
    for placeholder in target_layout.placeholders:
        slots.append((placeholder.element.ph_type, placeholder.element.ph_idx))
    return sorted(slots, key=lambda slot: slot[1])


def _compute_mapping(slide_phs, target_layout, placeholder_map):
    """Return `{source_idx: (target_type, target_idx) | None}` for every slide placeholder."""
    source_idxs = [shape.element.ph_idx for shape in slide_phs]
    if len(set(source_idxs)) != len(source_idxs):
        raise UnsupportedStructureError(
            "slide has duplicate placeholder idx values %r; reconciliation would be "
            "ambiguous" % source_idxs
        )
    slots = _layout_placeholder_slots(target_layout)
    slot_by_idx = {idx: (ph_type, idx) for ph_type, idx in slots}
    claimed = set()
    mapping: "Dict[int, Optional[Tuple[PP_PLACEHOLDER, int]]]" = {}

    explicit: "Dict[int, Optional[int]]" = {}
    if placeholder_map != "auto":
        if not isinstance(placeholder_map, dict):
            raise ValueError(
                "placeholder_map must be 'auto' or a {source_idx: target_idx | None} dict"
            )
        for source_idx, target_idx in placeholder_map.items():
            if source_idx not in source_idxs:
                raise ValueError(
                    "placeholder_map source idx %r is not a placeholder on this slide"
                    % (source_idx,)
                )
            if target_idx is not None and target_idx not in slot_by_idx:
                raise ValueError(
                    "placeholder_map target idx %r is not a placeholder on the target "
                    "layout" % (target_idx,)
                )
            explicit[source_idx] = target_idx
        targets = [t for t in explicit.values() if t is not None]
        if len(set(targets)) != len(targets):
            raise ValueError("placeholder_map maps two source placeholders to one target")

    for source_idx, target_idx in explicit.items():
        if target_idx is None:
            mapping[source_idx] = None
        else:
            mapping[source_idx] = slot_by_idx[target_idx]
            claimed.add(target_idx)

    # -- auto matching in three GLOBAL passes: every exact type+idx match settles before
    # -- any type-matching, and every type match before any family fallback. Interleaving
    # -- the tiers per-placeholder would let a lower-idx placeholder steal a higher-idx
    # -- placeholder's exact slot.
    unmatched = sorted(
        (shape for shape in slide_phs if shape.element.ph_idx not in mapping),
        key=lambda s: s.element.ph_idx,
    )

    still_unmatched = []
    for shape in unmatched:
        source_idx = shape.element.ph_idx
        exact = slot_by_idx.get(source_idx)
        if exact is not None and exact[0] == shape.element.ph_type and (source_idx not in claimed):
            mapping[source_idx] = exact
            claimed.add(source_idx)
        else:
            still_unmatched.append(shape)

    unmatched, still_unmatched = still_unmatched, []
    for shape in unmatched:
        source_type = shape.element.ph_type
        same_type = [slot for slot in slots if slot[0] == source_type and slot[1] not in claimed]
        if same_type:
            mapping[shape.element.ph_idx] = same_type[0]
            claimed.add(same_type[0][1])
        else:
            still_unmatched.append(shape)

    for shape in still_unmatched:
        source_type = shape.element.ph_type
        family = next((f for f in _TYPE_FAMILIES if source_type in f), None)
        familial = (
            [slot for slot in slots if slot[0] in family and slot[1] not in claimed]
            if family is not None
            else []
        )
        if familial:
            mapping[shape.element.ph_idx] = familial[0]
            claimed.add(familial[0][1])
        else:
            mapping[shape.element.ph_idx] = None
    return mapping


# ------------------------------------------------------------------------ resolve and bake


def _resolution_state(slide, *, align_by_content: bool = False):
    """(comparable values, payload) per run, keyed (shape_id, block_ordinal, run_index).

    The key is shape-scoped so shapes appearing or disappearing elsewhere on the slide
    (a baked-away furniture placeholder, an added shape) can never pair unrelated runs.
    """
    from pptx.inspect import inspect_text

    state = {}
    block_ordinals: dict = {}
    for block in inspect_text(slide).blocks:
        ordinal = block_ordinals.get(block.shape_id, 0)
        block_ordinals[block.shape_id] = ordinal + 1
        indexed_runs = [
            (run_index, run)
            for run_index, run in enumerate(block.runs)
            if run.text != "\v"
        ]
        identities = [
            ("field", run.field_type) if run.field_type is not None else ("text", run.text)
            for _, run in indexed_runs
        ]
        for (run_index, run), identity in zip(indexed_runs, identities):
            if align_by_content and (not identity[1] or identities.count(identity) != 1):
                continue
            font = run.font
            key = (
                (block.shape_id, ordinal, identity)
                if align_by_content
                else (block.shape_id, ordinal, run_index)
            )
            values = (
                font.size.value,
                font.name.value,
                font.color_rgb.value,
                font.bold.value if font.bold is not None else None,
                font.italic.value if font.italic is not None else None,
                font.underline.value if font.underline is not None else None,
            )
            state[key] = (values, run.text, block.anchor.part, font.to_dict(), run_index)
    return state


def _shifts_between(before_state, after_state) -> "Tuple[RunShift, ...]":
    shifts = []
    for key in sorted(set(before_state) & set(after_state)):
        before_values, text, part, before_payload, run_index = before_state[key]
        after_values, _, _, after_payload, _ = after_state[key]
        if before_values != after_values:
            shifts.append(
                RunShift(
                    part=part,
                    shape_id=key[0],
                    block_ordinal=key[1],
                    run_index=run_index,
                    text=text,
                    before=before_payload,
                    after=after_payload,
                )
            )
    return tuple(shifts)


def _bake_placeholder(shape) -> None:
    """Turn placeholder `shape` into a free shape with its effective look made local."""
    # -- geometry first: materialize the inherited position and size. Read ALL values
    # -- before writing ANY: writing `left` creates an a:off whose y defaults to 0, which
    # -- a subsequent `top` read would see as a local value, poisoning the inheritance.
    geometry = {
        attribute: getattr(shape, attribute) for attribute in ("left", "top", "width", "height")
    }
    for attribute, value in geometry.items():
        setattr(shape, attribute, Emu(value))

    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                effective = run.effective_font()
                if effective.size.resolved and effective.size.value is not None:
                    run.font.size = Emu(effective.size.value)
                if effective.name.resolved and effective.name.value is not None:
                    run.font.name = effective.name.value
                if effective.color_rgb.resolved and effective.color_rgb.value is not None:
                    from pptx.dml.color import RGBColor

                    run.font.color.rgb = RGBColor.from_string(effective.color_rgb.value)
                elif effective.color_rgb.bake_color_xml is not None:
                    from pptx.dml.color import RGBColor
                    from pptx.oxml import parse_xml

                    color = parse_xml(effective.color_rgb.bake_color_xml)
                    run.font.color.rgb = RGBColor.from_string(color.get("val"))
                    target = run._r.get_or_add_rPr().find(qn("a:solidFill"))[0]
                    for transform in color:
                        target.append(copy.deepcopy(transform))
                for emphasis in ("bold", "italic"):
                    effective_value = getattr(effective, emphasis)
                    if (
                        effective_value is not None
                        and effective_value.resolved
                        and effective_value.value is not None
                    ):
                        setattr(run.font, emphasis, effective_value.value)
                if (
                    effective.underline is not None
                    and effective.underline.resolved
                    and (effective.underline.value is not None)
                ):
                    from pptx.enum.text import MSO_TEXT_UNDERLINE_TYPE

                    # -- the resolved value is the raw `u` token ("none", "sng", "dbl",
                    # -- ...); bake the exact token so the style (including "explicitly
                    # -- not underlined") cannot re-resolve under the new layout
                    run.font.underline = MSO_TEXT_UNDERLINE_TYPE.from_xml(effective.underline.value)

    ph = shape.element.ph
    ph.getparent().remove(ph)


def _restore_element(target, snapshot) -> None:
    """Restore `target` in place so existing slide proxies remain usable."""
    target.attrib.clear()
    target.attrib.update(snapshot.attrib)
    target.text = snapshot.text
    target.tail = snapshot.tail
    for child in list(target):
        target.remove(child)
    for child in snapshot:
        target.append(copy.deepcopy(child))
