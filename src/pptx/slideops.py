"""Internal machinery for slide clone/delete/reorder (paper-pptx). Not public API.

The public surface is `Slides.clone/delete/reorder/move` and `SlideClonePolicy` in
`pptx.slide`. Everything here operates on the in-memory opc package — parts, relationships,
content types — never on unpacked files; `[Content_Types].xml` regenerates from live parts at
save, and a part no longer reachable through the relationship graph is simply never
serialized, so orphans structurally cannot reach disk.

Clone is validate-fully-then-mutate: the complete relationship plan for the source slide (and
every deep-copied part's own relationships) is validated against the policy BEFORE any part is
created, so a `RelationshipPolicyError` provably leaves the package untouched.
"""

from __future__ import annotations

import copy
import re
import uuid
from typing import TYPE_CHECKING, Dict, List, Tuple

from pptx.errors import RelationshipPolicyError
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI

if TYPE_CHECKING:
    from pptx.parts.slide import SlidePart

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
#: Microsoft chart-style extension parts (LibreOffice and recent PowerPoint emit these)
_CHART_COLOR_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
_CHART_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle"

#: media-like relationship types, shared between clone and source by default
_MEDIA_RELTYPES = frozenset([RT.IMAGE, RT.MEDIA, RT.VIDEO, RT.AUDIO])
#: relationship types allowed FROM a chart part, all deep-copied with it
_CHART_CHILD_RELTYPES = frozenset([RT.PACKAGE, _CHART_COLOR_STYLE, _CHART_STYLE])
#: relationship types allowed FROM a notes-slide part
_NOTES_CHILD_RELTYPES = frozenset([RT.NOTES_MASTER, RT.SLIDE])
_A16_CREATION_ID = "{http://schemas.microsoft.com/office/drawing/2014/main}creationId"
_A16_CXN_DE_REFS = "{http://schemas.microsoft.com/office/drawing/2014/main}cxnDERefs"
_A16_PRED_DE_REF = "{http://schemas.microsoft.com/office/drawing/2014/main}predDERef"
_A_FIELD = "{http://schemas.openxmlformats.org/drawingml/2006/main}fld"


class _CopySession(set):
    """Partname reservations and document-wide identity remaps for one copy."""

    def __init__(self, package, source_parts):
        super().__init__()
        self._reserved = _document_identity_values(package)
        self._a16_mapping = {}
        seen = set()
        for part in source_parts:
            root = _xml_root(part)
            if root is None:
                continue
            for element in root.iter(_A16_CREATION_ID):
                old = element.get("id")
                if not old:
                    continue
                key = old.lower()
                if key in seen:
                    raise RelationshipPolicyError(
                        "copied parts contain duplicate a16 creation identity %r" % old
                    )
                seen.add(key)
                self._a16_mapping[key] = self._fresh_guid()

    def remap(self, source_part, copied_part) -> None:
        root = _xml_root(copied_part)
        if root is None:
            return
        self.remap_element(root)
        if not isinstance(copied_part, XmlPart):
            from lxml import etree

            copied_part.blob = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

    def remap_element(self, root) -> None:
        """Apply this copy session's document-wide identity plan to an XML subtree."""
        for element in root.iter():
            if element.tag == _A16_CREATION_ID:
                _replace_identity_attribute(element, "id", self._a16_mapping)
            elif element.tag == _A16_CXN_DE_REFS:
                _replace_identity_attribute(element, "st", self._a16_mapping)
                _replace_identity_attribute(element, "end", self._a16_mapping)
            elif element.tag == _A16_PRED_DE_REF:
                _replace_identity_attribute(element, "pred", self._a16_mapping)
            elif element.tag == _A_FIELD and element.get("id"):
                element.set("id", self._fresh_guid())

    def _fresh_guid(self) -> str:
        while True:
            value = "{%s}" % str(uuid.uuid4()).upper()
            if value.lower() not in self._reserved:
                self._reserved.add(value.lower())
                return value


def _replace_identity_attribute(element, name: str, mapping: dict) -> None:
    value = element.get(name)
    if value is not None and value.lower() in mapping:
        element.set(name, mapping[value.lower()])


def _xml_root(part):
    if isinstance(part, XmlPart):
        return part._element
    content_type = part.content_type.partition(";")[0].strip().lower()
    if content_type not in ("application/xml", "text/xml") and not content_type.endswith("+xml"):
        return None
    from pptx.oxml import parse_xml

    try:
        return parse_xml(part.blob)
    except Exception:
        return None


def _document_identity_values(package) -> set:
    values = set()
    for part in package.iter_parts():
        root = _xml_root(part)
        if root is None:
            continue
        for element in root.iter():
            if element.tag in (_A16_CREATION_ID, _A_FIELD) and element.get("id"):
                values.add(element.get("id").lower())
    return values


def clone_slide_part(source_part: "SlidePart", policy) -> "SlidePart":
    """Return a new `SlidePart` that is a policy-governed deep copy of `source_part`.

    The new part is fully related (layout, media, charts+workbooks, notes per `policy`) but
    NOT yet added to the presentation's slide list — the caller owns `p:sldIdLst`.
    """
    from pptx.parts.slide import SlidePart

    package = source_part.package
    plan = _validated_plan(source_part, policy)

    # -- Parts created here are unreachable from the package rels graph until the caller
    # -- relates the new slide, so `package.next_partname` cannot see them. `allocated`
    # -- tracks every partname assigned during THIS clone so two deep copies sharing a
    # -- template (e.g. two charts) can never collide.
    copied_sources = [source_part]
    for _, action, rel in plan:
        if action == "copy":
            copied_sources.append(rel.target_part)
        elif action == "chart":
            copied_sources.append(rel.target_part)
            copied_sources.extend(
                child.target_part
                for child in rel.target_part.rels.values()
                if not child.is_external
            )
        elif action == "notes":
            copied_sources.append(rel.target_part)
    allocated = _CopySession(package, copied_sources)
    new_part = SlidePart(
        _allocate_partname(package, "/ppt/slides/slide%d.xml", allocated),
        CT.PML_SLIDE,
        package,
        copy.deepcopy(source_part._element),
    )
    allocated.remap(source_part, new_part)

    rId_mapping: "Dict[str, str]" = {}
    for old_rId, action, rel in plan:
        if action == "external":
            rId_mapping[old_rId] = new_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif action == "share":
            rId_mapping[old_rId] = new_part.relate_to(rel.target_part, rel.reltype)
        elif action == "copy":
            rId_mapping[old_rId] = new_part.relate_to(
                _copy_leaf_part(rel.target_part, allocated), rel.reltype
            )
        elif action == "chart":
            rId_mapping[old_rId] = new_part.relate_to(
                _copy_chart_part(rel.target_part, allocated), rel.reltype
            )
        elif action == "notes":
            rId_mapping[old_rId] = new_part.relate_to(
                _copy_notes_part(rel.target_part, new_part, allocated), rel.reltype
            )
        # -- action == "drop": no relationship on the clone (notes policy) --

    _rewrite_r_references(new_part._element, rId_mapping)
    return new_part


def remove_slide_from_id_lists(presentation_elm, slide_id: int, rId: str) -> None:
    """Purge `slide_id`/`rId` bookkeeping for a deleted slide from auxiliary ID lists.

    Sections (`p14:sectionLst`, in the presentation's extension list) reference slides by
    slide id; custom shows (`p:custShowLst`) reference them by relationship id. Neither is
    reachable through the relationship graph, so without this step a delete leaves dangling
    entries behind — the corruption class this step closes. Empty sections and
    empty custom-show slide lists are schema-valid and left in place (matching PowerPoint,
    which keeps an emptied section).
    """
    slide_id_str = str(slide_id)
    for section_sldId in presentation_elm.findall(
        ".//{%s}sectionLst//{%s}sldId" % (_P14_NS, _P14_NS)
    ):
        if section_sldId.get("id") == slide_id_str:
            section_sldId.getparent().remove(section_sldId)
    for show_sld in presentation_elm.findall(".//{%s}custShowLst//{%s}sld" % (_P_NS, _P_NS)):
        if show_sld.get("{%s}id" % _R_NS) == rId:
            show_sld.getparent().remove(show_sld)


def enroll_clone_in_section(presentation_elm, source_slide_id: int, clone_slide_id: int) -> None:
    """Add `clone_slide_id` to the section holding `source_slide_id`, directly after it.

    No-op when the deck has no sections or the source slide is not enrolled in one. Custom
    shows are deliberately NOT extended: a copy is not part of a curated show.
    """
    source_id_str = str(source_slide_id)
    for section_sldId in presentation_elm.findall(
        ".//{%s}sectionLst//{%s}sldId" % (_P14_NS, _P14_NS)
    ):
        if section_sldId.get("id") == source_id_str:
            clone_entry = section_sldId.makeelement(
                "{%s}sldId" % _P14_NS, {"id": str(clone_slide_id)}
            )
            section_sldId.addnext(clone_entry)
            return


def _allocate_partname(package, template: str, allocated: "set") -> PackURI:
    """Return the lowest-numbered partname from `template` unused in `package` OR `allocated`.

    Records the returned name in `allocated` so subsequent allocations within the same
    (not-yet-related) clone operation cannot reuse it.
    """
    used = {str(part.partname) for part in package.iter_parts()} | allocated
    index = 1
    while template % index in used:
        index += 1
    partname = template % index
    allocated.add(partname)
    return PackURI(partname)


def _validated_plan(source_part, policy) -> "List[Tuple[str, str, object]]":
    """Return [(rId, action, rel)] for every source relationship, or raise before mutating.

    Also pre-validates the relationship graphs of parts that will be deep-copied (charts and
    notes), so no failure can occur after part creation begins.
    """
    plan = []
    unsupported = []
    for rId in sorted(source_part.rels, key=_rId_sort_key):
        rel = source_part.rels[rId]
        if rel.is_external:
            plan.append((rId, "external", rel))
            continue
        target = _owned_target(source_part, rel, "slide clone")
        if rel.reltype == RT.SLIDE_LAYOUT:
            plan.append((rId, "share", rel))
        elif rel.reltype in _MEDIA_RELTYPES:
            plan.append((rId, "share" if policy.share_media else "copy", rel))
        elif rel.reltype == RT.CHART:
            if not policy.deep_copy_charts:
                raise RelationshipPolicyError(
                    "cloning a slide with a chart requires deep_copy_charts=True: sharing an"
                    " editable chart part between slides is exactly the cross-contamination"
                    " this API exists to prevent, and is not offered in v0"
                )
            _validate_chart_rels(target)
            plan.append((rId, "chart", rel))
        elif rel.reltype == RT.NOTES_SLIDE:
            if policy.deep_copy_notes:
                _validate_notes_rels(target)
                plan.append((rId, "notes", rel))
            else:
                plan.append((rId, "drop", rel))
        else:
            unsupported.append(rel.reltype)
    if unsupported:
        raise RelationshipPolicyError(
            "slide has relationship types clone does not support in v0: %s"
            % ", ".join(sorted(unsupported))
        )
    return plan


def _validate_chart_rels(chart_part) -> None:
    for rId in chart_part.rels:
        rel = chart_part.rels[rId]
        if rel.is_external:
            continue
        child_part = _owned_target(chart_part, rel, "chart clone")
        if rel.reltype not in _CHART_CHILD_RELTYPES:
            raise RelationshipPolicyError(
                "chart part %s has relationship type clone does not support in v0: %s"
                % (chart_part.partname, rel.reltype)
            )
        if any(not child.is_external for child in child_part.rels.values()):
            raise RelationshipPolicyError(
                "chart child part %s has internal relationships of its own; clone supports"
                " only leaf chart children in v0" % rel.target_part.partname
            )


def _validate_notes_rels(notes_part) -> None:
    for rId in notes_part.rels:
        rel = notes_part.rels[rId]
        if not rel.is_external:
            _owned_target(notes_part, rel, "notes clone")
        if not rel.is_external and rel.reltype not in _NOTES_CHILD_RELTYPES:
            raise RelationshipPolicyError(
                "notes slide %s has relationship type clone does not support in v0: %s"
                % (notes_part.partname, rel.reltype)
            )


def _owned_target(owner_part, rel, context: str):
    """Return an internal target only when it belongs to the owner's package."""
    try:
        target = rel.target_part
    except (AssertionError, AttributeError, TypeError, ValueError) as exc:
        raise RelationshipPolicyError(
            "%s relationship %s has an invalid internal target" % (context, rel.rId)
        ) from exc
    if target.package is not owner_part.package:
        raise RelationshipPolicyError(
            "%s relationship %s targets a part owned by another package" % (context, rel.rId)
        )
    return target


def _copy_chart_part(chart_part, allocated):
    """Return a deep copy of `chart_part` including its embedded workbook and style parts."""
    new_chart = _copy_leaf_part(chart_part, allocated)
    rId_mapping = {}
    for rId in sorted(chart_part.rels, key=_rId_sort_key):
        rel = chart_part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_chart.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rId_mapping[rId] = new_chart.relate_to(
                _copy_leaf_part(rel.target_part, allocated), rel.reltype
            )
    _rewrite_r_references(new_chart._element, rId_mapping)
    return new_chart


def _copy_notes_part(notes_part, new_slide_part, allocated):
    """Return a deep copy of `notes_part`, related to the notes master and the CLONE slide."""
    new_notes = _copy_leaf_part(notes_part, allocated)
    rId_mapping = {}
    for rId in sorted(notes_part.rels, key=_rId_sort_key):
        rel = notes_part.rels[rId]
        if rel.is_external:
            rId_mapping[rId] = new_notes.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        elif rel.reltype == RT.SLIDE:
            rId_mapping[rId] = new_notes.relate_to(new_slide_part, RT.SLIDE)
        else:  # -- RT.NOTES_MASTER: shared singleton
            rId_mapping[rId] = new_notes.relate_to(rel.target_part, rel.reltype)
    _rewrite_r_references(new_notes._element, rId_mapping)
    return new_notes


def _copy_leaf_part(part, allocated):
    """Return a new part of `part`'s class with copied content and a fresh partname."""
    package = part.package
    partname = _allocate_partname(package, _partname_template(str(part.partname)), allocated)
    if isinstance(part, XmlPart):
        copied = type(part)(partname, part.content_type, package, copy.deepcopy(part._element))
    else:
        copied = type(part)(partname, part.content_type, package, part.blob)
    if hasattr(allocated, "remap"):
        allocated.remap(part, copied)
    return copied


def _partname_template(partname: str) -> str:
    """Return a next_partname template for `partname`, e.g. "/ppt/charts/chart%d.xml"."""
    template, substitutions = re.subn(r"[0-9]+(?=\.[^.]+$)", "%d", partname, count=1)
    if substitutions == 0:
        stem, dot, ext = partname.rpartition(".")
        template = "%s%%d%s%s" % (stem, dot, ext)
    return template


def _rewrite_r_references(root, rId_mapping: "Dict[str, str]") -> None:
    """Rewrite every r-namespace attribute in `root` per `rId_mapping`, in place."""
    if not rId_mapping:
        return
    prefix = "{%s}" % _R_NS
    for element in root.iter():
        for attr_name, attr_value in element.attrib.items():
            if attr_name.startswith(prefix) and attr_value in rId_mapping:
                element.set(attr_name, rId_mapping[attr_value])


def _rId_sort_key(rId: str):
    match = re.fullmatch(r"rId([0-9]+)", rId)
    return (0, int(match.group(1))) if match else (1, rId)
