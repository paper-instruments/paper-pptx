"""Slide-related objects, including masters, layouts, and notes."""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator, Sequence, cast

from pptx.dml.fill import FillFormat
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.errors import TargetNotFoundError, UnsupportedStructureError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.shapes.shapetree import (
    LayoutPlaceholders,
    LayoutShapes,
    MasterPlaceholders,
    MasterShapes,
    NotesSlidePlaceholders,
    NotesSlideShapes,
    SlidePlaceholders,
    SlideShapes,
)
from pptx.shared import ElementProxy, ParentedElementProxy, PartElementProxy
from pptx.util import lazyproperty


def _relationship_references(root, rId: str) -> bool:
    """Return whether any relationship-qualified XML attribute contains `rId`."""
    return any(
        value == rId
        and name.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}")
        for element in root.iter()
        for name, value in element.attrib.items()
    )


def _require_slide_enrolled(slide: "Slide", *, argument: str = "slide") -> None:
    """Refuse a detached slide proxy before it can mutate unreachable XML."""
    package = slide.part.package
    presentation_part = package.presentation_part
    matches = []
    for sldId in presentation_part._element.sldIdLst.sldId_lst:
        try:
            rel = presentation_part.rels[sldId.rId]
            if not rel.is_external and rel.reltype == RT.SLIDE and rel.target_part is slide.part:
                matches.append(sldId)
        except (AssertionError, KeyError, ValueError):
            continue
    if len(matches) != 1 or slide._element is not slide.part._element:
        raise TargetNotFoundError("%s is stale or no longer enrolled" % argument)


def _require_layout_enrolled(layout: "SlideLayout", *, argument: str = "slide_layout") -> None:
    """Refuse a detached layout proxy before it can be selected as a target."""
    from pptx.parts.slide import SlideMasterPart

    package = layout.part.package
    matches = []
    for part in package.iter_parts():
        if not isinstance(part, SlideMasterPart):
            continue
        id_list = part._element.sldLayoutIdLst
        if id_list is None:
            continue
        for entry in id_list.sldLayoutId_lst:
            try:
                rel = part.rels[entry.rId]
                if (
                    not rel.is_external
                    and rel.reltype == RT.SLIDE_LAYOUT
                    and rel.target_part is layout.part
                ):
                    matches.append((part, entry))
            except (AssertionError, KeyError, ValueError):
                continue
    if len(matches) != 1 or layout._element is not layout.part._element:
        raise TargetNotFoundError("%s is stale or no longer enrolled" % argument)


def _inbound_relationships(package, target_part):
    """Return every reachable internal relationship targeting `target_part`."""
    inbound = []
    owners = [package] + list(package.iter_parts())
    for owner in owners:
        relationships = package._rels if owner is package else owner.rels
        for rId, rel in relationships.items():
            if rel.is_external:
                continue
            try:
                if rel.target_part is target_part:
                    inbound.append((owner, rId, rel))
            except (AssertionError, ValueError):
                continue
    return tuple(inbound)


if TYPE_CHECKING:
    from datetime import datetime

    from pptx.oxml.presentation import CT_SlideIdList, CT_SlideMasterIdList
    from pptx.oxml.slide import (
        CT_CommonSlideData,
        CT_NotesSlide,
        CT_Slide,
        CT_SlideLayoutIdList,
        CT_SlideMaster,
    )
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import SlideLayoutPart, SlideMasterPart, SlidePart
    from pptx.presentation import Presentation
    from pptx.rebind import RebindReport
    from pptx.shapes.placeholder import LayoutPlaceholder, MasterPlaceholder
    from pptx.shapes.shapetree import NotesSlidePlaceholder
    from pptx.text.text import TextFrame


class _BaseSlide(PartElementProxy):
    """Base class for slide objects, including masters, layouts and notes."""

    _element: CT_Slide

    @lazyproperty
    def background(self) -> _Background:
        """|_Background| object providing slide background properties.

        This property returns a |_Background| object whether or not the
        slide, master, or layout has an explicitly defined background.

        The same |_Background| object is returned on every call for the same
        slide object.
        """
        return _Background(self._element.cSld)

    @property
    def name(self) -> str:
        """String representing the internal name of this slide.

        Returns an empty string (`''`) if no name is assigned. Assigning an empty string or |None|
        to this property causes any name to be removed.
        """
        return self._element.cSld.name

    @name.setter
    def name(self, value: str | None):
        new_value = "" if value is None else value
        self._element.cSld.name = new_value


class _BaseMaster(_BaseSlide):
    """Base class for master objects such as |SlideMaster| and |NotesMaster|.

    Provides access to placeholders and regular shapes.
    """

    @lazyproperty
    def placeholders(self) -> MasterPlaceholders:
        """|MasterPlaceholders| collection of placeholder shapes in this master.

        Sequence sorted in `idx` order.
        """
        return MasterPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |MasterShapes| containing sequence of shape objects
        appearing on this slide.
        """
        return MasterShapes(self._element.spTree, self)


class NotesMaster(_BaseMaster):
    """Proxy for the notes master XML document.

    Provides access to shapes, the most commonly used of which are placeholders.
    """


class NotesSlide(_BaseSlide):
    """Notes slide object.

    Provides access to slide notes placeholder and other shapes on the notes handout
    page.
    """

    element: CT_NotesSlide  # pyright: ignore[reportIncompatibleMethodOverride]

    def clone_master_placeholders(self, notes_master: NotesMaster) -> None:
        """Selectively add placeholder shape elements from `notes_master`.

        Selected placeholder shape elements from `notes_master` are added to the shapes
        collection of this notes slide. Z-order of placeholders is preserved. Certain
        placeholders (header, date, footer) are not cloned.
        """

        def iter_cloneable_placeholders() -> Iterator[MasterPlaceholder]:
            """Generate a reference to each cloneable placeholder in `notes_master`.

            These are the placeholders that should be cloned to a notes slide when the a new notes
            slide is created.
            """
            cloneable = (
                PP_PLACEHOLDER.SLIDE_IMAGE,
                PP_PLACEHOLDER.BODY,
                PP_PLACEHOLDER.SLIDE_NUMBER,
            )
            for placeholder in notes_master.placeholders:
                if placeholder.element.ph_type in cloneable:
                    yield placeholder

        shapes = self.shapes
        for placeholder in iter_cloneable_placeholders():
            shapes.clone_placeholder(cast("LayoutPlaceholder", placeholder))

    @property
    def notes_placeholder(self) -> NotesSlidePlaceholder | None:
        """the notes placeholder on this notes slide, the shape that contains the actual notes text.

        Return |None| if no notes placeholder is present; while this is probably uncommon, it can
        happen if the notes master does not have a body placeholder, or if the notes placeholder
        has been deleted from the notes slide.
        """
        for placeholder in self.placeholders:
            if placeholder.placeholder_format.type == PP_PLACEHOLDER.BODY:
                return placeholder
        return None

    @property
    def notes_text_frame(self) -> TextFrame | None:
        """The text frame of the notes placeholder on this notes slide.

        |None| if there is no notes placeholder. This is a shortcut to accommodate the common case
        of simply adding "notes" text to the notes "page".
        """
        notes_placeholder = self.notes_placeholder
        if notes_placeholder is None:
            return None
        return notes_placeholder.text_frame

    @lazyproperty
    def placeholders(self) -> NotesSlidePlaceholders:
        """Instance of |NotesSlidePlaceholders| for this notes-slide.

        Contains the sequence of placeholder shapes in this notes slide.
        """
        return NotesSlidePlaceholders(self.element.spTree, self)

    @lazyproperty
    def shapes(self) -> NotesSlideShapes:
        """Sequence of shape objects appearing on this notes slide."""
        return NotesSlideShapes(self._element.spTree, self)


class Slide(_BaseSlide):
    """Slide object. Provides access to shapes and slide-level properties."""

    part: SlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    def apply_footers(
        self,
        *,
        footer: str | None = None,
        slide_number: bool = False,
        date_format: str | None = None,
        fixed_date: str | None = None,
        now: "datetime | None" = None,
    ) -> None:
        """Apply the complete footer state to this slide only (the dialog's "Apply").

        paper-pptx addition. Same parameters, mechanism, and refusals as
        :meth:`.Presentation.apply_footers`, restricted to this slide — the per-slide
        override path (e.g. removing just this slide's footer while the rest of the deck
        keeps it). Each call sets this slide's full three-element state.
        """
        from pptx.hf import apply_slide_footers

        apply_slide_footers(
            self,
            footer=footer,
            slide_number=slide_number,
            date_format=date_format,
            fixed_date=fixed_date,
            now=now,
        )

    def rebind_layout(
        self,
        target_layout: "SlideLayout",
        *,
        placeholder_map="auto",
        orphan_policy: str = "refuse",
    ) -> "RebindReport":
        """Move this slide to `target_layout`; return the required |RebindReport|.

        paper-pptx addition — the template-migration *primitive* (bulk-migration
        workflows are left to the caller). Placeholders reconcile against the
        target layout: auto-matching binds by exact type+idx, then same type, then
        interchangeable type family (title/ctrTitle; body/object/subTitle). The two
        fallback tiers bind only when exactly one unclaimed target exists; ambiguity
        refuses and requires an explicit map. Pass
        `placeholder_map={source_idx: target_idx | None}` to override any of it (None
        force-orphans a source). Source placeholders with no destination follow
        `orphan_policy`: "refuse" (default; typed, atomic) or "bake" — convert to a free
        shape with inherited geometry materialized and each run's *resolved* effective
        formatting written locally, so the text keeps its look.

        The report is not optional: the effective-value resolver runs before and after,
        and every run whose resolved values changed appears with its before/after payloads
        — a rebind never shifts appearance silently. Same-package only (cross-package
        composition is `import_slide`'s job); slides carrying `mc:AlternateContent`
        refuse (shapes inside are invisible to reconciliation).
        """
        from pptx.rebind import rebind_layout

        return rebind_layout(
            self,
            target_layout,
            placeholder_map=placeholder_map,
            orphan_policy=orphan_policy,
        )

    @property
    def follow_master_background(self):
        """|True| if this slide inherits the slide master background.

        Assigning |False| causes background inheritance from the master to be
        interrupted; if there is no custom background for this slide,
        a default background is added. If a custom background already exists
        for this slide, assigning |False| has no effect.

        Assigning |True| causes any custom background for this slide to be
        deleted and inheritance from the master restored.
        """
        return self._element.bg is None

    @property
    def has_notes_slide(self) -> bool:
        """`True` if this slide has a notes slide, `False` otherwise.

        A notes slide is created by :attr:`.notes_slide` when one doesn't exist; use this property
        to test for a notes slide without the possible side effect of creating one.
        """
        return self.part.has_notes_slide

    @property
    def notes_slide(self) -> NotesSlide:
        """The |NotesSlide| instance for this slide.

        If the slide does not have a notes slide, one is created. The same single instance is
        returned on each call.
        """
        return self.part.notes_slide

    @lazyproperty
    def placeholders(self) -> SlidePlaceholders:
        """Sequence of placeholder shapes in this slide."""
        return SlidePlaceholders(self._element.spTree, self)

    def read_notes_text(self) -> str:
        """Return the text of this slide's existing speaker notes.

        paper-pptx addition. Unlike :attr:`notes_slide`, this NEVER creates a notes slide:
        a slide with no notes part raises |UnsupportedStructureError| (as does a notes slide
        with no body placeholder). Returns "" for an empty existing notes body.
        """
        return self._existing_notes_text_frame().text

    def replace_notes_text(self, text: str) -> None:
        """Replace the text of this slide's existing speaker notes with `text`.

        paper-pptx addition. Only the notes *body* placeholder is touched — slide-number and
        other notes placeholders are preserved untouched. The first paragraph's properties
        and its first run's character formatting are kept and applied to the replacement
        text; `"\\n"` in `text` starts a new paragraph. Never creates a notes slide: a slide
        with no notes part raises |UnsupportedStructureError| before anything changes
        (creating the notes part graph is intentionally not supported).
        """
        if not isinstance(text, str):
            raise ValueError("text must be a str, got %r" % type(text).__name__)
        try:
            text.encode("utf-8")  # -- lone surrogates would explode mid-mutation otherwise
        except UnicodeEncodeError:
            raise ValueError("text contains characters not encodable in XML: %r" % (text,))
        text_frame = self._existing_notes_text_frame()  # -- full validation before mutation

        txBody = text_frame._txBody
        paragraphs = txBody.p_lst
        first_p = paragraphs[0]
        first_r = first_p.find(qn("a:r"))
        rPr_template = None
        if first_r is not None:
            rPr = first_r.find(qn("a:rPr"))
            if rPr is not None:
                rPr_template = copy.deepcopy(rPr)

        # -- keep the first a:p element (preserving its a:pPr); drop the rest --
        for surplus_p in paragraphs[1:]:
            txBody.remove(surplus_p)
        for content in first_p.content_children:
            first_p.remove(content)
        pPr_template = first_p.find(qn("a:pPr"))

        lines = text.split("\n")
        for index, line in enumerate(lines):
            if index == 0:
                p = first_p
            else:
                p = txBody.add_p()
                if pPr_template is not None:
                    p.insert(0, copy.deepcopy(pPr_template))
            if line == "":
                continue  # -- an empty line is an empty paragraph
            r = p.add_r()
            if rPr_template is not None:
                r.insert(0, copy.deepcopy(rPr_template))
            r.text = line

    def _existing_notes_text_frame(self) -> TextFrame:
        """Return the body-placeholder text frame of this slide's EXISTING notes slide.

        Raises |UnsupportedStructureError| (never creates anything) when the slide has no
        notes part or its notes slide has no body placeholder.
        """
        if not self.has_notes_slide:
            raise UnsupportedStructureError(
                "slide %d has no notes slide; creating one is out of scope for this API"
                " (use notes_slide if you explicitly want creation)" % self.slide_id
            )
        notes_slide = self.part.part_related_by(RT.NOTES_SLIDE).notes_slide
        text_frame = notes_slide.notes_text_frame
        if text_frame is None:
            raise UnsupportedStructureError(
                "notes slide of slide %d has no body placeholder to hold notes text" % self.slide_id
            )
        return text_frame

    @lazyproperty
    def shapes(self) -> SlideShapes:
        """Sequence of shape objects appearing on this slide."""
        return SlideShapes(self._element.spTree, self)

    @property
    def slide_id(self) -> int:
        """Integer value that uniquely identifies this slide within this presentation.

        The slide id does not change if the position of this slide in the slide sequence is changed
        by adding, rearranging, or deleting slides.
        """
        return self.part.slide_id

    @property
    def slide_layout(self) -> SlideLayout:
        """|SlideLayout| object this slide inherits appearance from."""
        return self.part.slide_layout


@dataclass(frozen=True)
class SlideClonePolicy:
    """Relationship policy for `Slides.clone` (paper-pptx addition).

    Defaults encode the production-proven policy: charts (with their embedded workbooks and
    style parts) and speaker notes are deep-copied so clone and original can never
    cross-contaminate; image/media parts are shared deliberately.

    - `deep_copy_charts`: must be True to clone a slide bearing charts; False refuses
      (`RelationshipPolicyError`) rather than share an editable chart part between slides.
    - `deep_copy_notes`: False drops the notes slide from the clone (original unaffected).
    - `share_media`: False deep-copies image/media parts instead of sharing them.
    """

    deep_copy_charts: bool = True
    deep_copy_notes: bool = True
    share_media: bool = True


class Slides(ParentedElementProxy):
    """Sequence of slides belonging to an instance of |Presentation|.

    Has list semantics for access to individual slides. Supports indexed access, len(), and
    iteration.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldIdLst: CT_SlideIdList, prs: Presentation):
        super(Slides, self).__init__(sldIdLst, prs)
        self._sldIdLst = sldIdLst

    def __getitem__(self, idx: int) -> Slide:
        """Provide indexed access, (e.g. 'slides[0]')."""
        try:
            sldId = self._sldIdLst.sldId_lst[idx]
        except IndexError:
            raise IndexError("slide index out of range")
        return self.part.related_slide(sldId.rId)

    def __iter__(self) -> Iterator[Slide]:
        """Support iteration, e.g. `for slide in slides:`."""
        for sldId in self._sldIdLst.sldId_lst:
            yield self.part.related_slide(sldId.rId)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(slides) == 4`."""
        return len(self._sldIdLst)

    def add_slide(self, slide_layout: SlideLayout) -> Slide:
        """Return a newly added slide that inherits layout from `slide_layout`."""
        rId, slide = self.part.add_slide(slide_layout)
        slide.shapes.clone_layout_placeholders(slide_layout)
        self._sldIdLst.add_sldId(rId)
        return slide

    def clone(
        self,
        source: Slide | int,
        *,
        after: Slide | int | None = None,
        policy: SlideClonePolicy | None = None,
    ) -> Slide:
        """Return a new slide that is a policy-governed deep copy of `source`.

        paper-pptx addition. The clone's relationship graph follows `policy` (default
        |SlideClonePolicy|): layout shared; charts deep-copied WITH their embedded workbooks
        and style parts; notes deep-copied and re-linked to the clone; image/media shared;
        external (hyperlink) relationships copied. A slide bearing any other relationship
        type (OLE objects, controls, SmartArt, comments, …) refuses with
        |RelationshipPolicyError| before anything changes.

        The clone is inserted directly after `source`, or after the slide given by `after`.
        `source`/`after` accept a |Slide| or a 0-based index; a |Slide| from another
        presentation raises |TargetNotFoundError|.
        """
        from pptx._transaction import PackageTransaction
        from pptx.slideops import clone_slide_part, enroll_clone_in_section

        if policy is None:
            policy = SlideClonePolicy()
        if not isinstance(policy, SlideClonePolicy):
            raise ValueError("policy must be a SlideClonePolicy, got %r" % (policy,))
        source_slide = self._resolve_slide(source)
        anchor_slide = source_slide if after is None else self._resolve_slide(after)
        anchor_index = self.index(anchor_slide)

        source_slide_id = source_slide.slide_id
        with PackageTransaction(self.part.package, self, source_slide, anchor_slide):
            new_part = clone_slide_part(source_slide.part, policy)
            rId = self.part.relate_to(new_part, RT.SLIDE)
            self._sldIdLst.add_sldId(rId)
            sldId = self._sldIdLst[-1]
            self._sldIdLst.remove(sldId)
            self._sldIdLst.insert(anchor_index + 1, sldId)
            # -- enroll the copy in the source's section, right after it (custom shows are
            # -- deliberately not extended: a copy is not part of a curated show)
            enroll_clone_in_section(self._sldIdLst.getparent(), source_slide_id, sldId.id)
            cloned_slide = new_part.slide
        return cloned_slide

    def delete(self, slide: Slide | int) -> None:
        """Remove `slide` from this presentation.

        paper-pptx addition. Removes the slide's `p:sldId` entry and the presentation's
        relationship to the slide part; parts then unreachable through the relationship
        graph (the slide, and e.g. its charts and notes if unshared) are never serialized
        again — orphans structurally cannot reach disk. Deleting the last slide is allowed.
        """
        from pptx._transaction import PackageTransaction
        from pptx.slideops import remove_slide_from_id_lists

        target = self._resolve_slide(slide)
        for sldId in self._sldIdLst.sldId_lst:
            if sldId.id == target.slide_id:
                slide_id, rId = sldId.id, sldId.rId
                notes_owners = {
                    rel.target_part
                    for rel in target.part.rels.values()
                    if not rel.is_external and rel.reltype == RT.NOTES_SLIDE
                }
                for notes_part in notes_owners:
                    shared_notes = [
                        (owner, notes_rId)
                        for owner, notes_rId, _ in _inbound_relationships(
                            self.part.package, notes_part
                        )
                        if owner is not target.part
                    ]
                    if shared_notes:
                        raise UnsupportedStructureError(
                            "slide deletion refused: its notes part is shared by another "
                            "reachable package part"
                        )
                aliases = [
                    (owner, alias_rId)
                    for owner, alias_rId, _ in _inbound_relationships(
                        self.part.package, target.part
                    )
                    if not ((owner is self.part and alias_rId == rId) or owner in notes_owners)
                ]
                if aliases:
                    raise UnsupportedStructureError(
                        "slide deletion refused: slide part has additional inbound "
                        "relationship aliases"
                    )
                with PackageTransaction(self.part.package, self, target):
                    self._sldIdLst.remove(sldId)
                    # -- sections (by slide id) and custom shows (by rId) reference slides
                    # -- outside the rels graph; purge those entries too
                    remove_slide_from_id_lists(self._sldIdLst.getparent(), slide_id, rId)
                    if _relationship_references(self.part._element, rId):
                        raise UnsupportedStructureError(
                            "slide deletion refused: relationship %s remains referenced" % rId
                        )
                    self.part.drop_rel(rId)
                return

    def move(self, slide: Slide | int, to_index: int) -> None:
        """Move `slide` so it sits at 0-based `to_index` in the slide sequence.

        paper-pptx addition. `to_index` outside `range(len(slides))` raises |ValueError|.
        """
        target = self._resolve_slide(slide)
        if (
            not isinstance(to_index, int)
            or isinstance(to_index, bool)
            or not (0 <= to_index < len(self))
        ):
            raise ValueError(
                "to_index must be an int in range 0..%d, got %r" % (len(self) - 1, to_index)
            )
        current_index = self.index(target)
        sldId = self._sldIdLst.sldId_lst[current_index]
        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self, target):
            self._sldIdLst.remove(sldId)
            self._sldIdLst.insert(to_index, sldId)

    def reorder(self, new_order: Sequence[int]) -> None:
        """Permute the slide sequence: new position i shows the slide now at `new_order[i]`.

        paper-pptx addition. `new_order` must be an exact permutation of
        `range(len(slides))`; anything else raises |ValueError| before any change.
        """
        if sorted(new_order) != list(range(len(self))):
            raise ValueError(
                "new_order must be a permutation of range(%d), got %r"
                % (len(self), list(new_order))
            )
        current = list(self._sldIdLst.sldId_lst)
        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            for index in new_order:
                self._sldIdLst.append(current[index])  # -- re-appending moves the element

    def _resolve_slide(self, value: Slide | int) -> Slide:
        """Return the |Slide| in this collection for `value` (a Slide or 0-based index).

        An int resolves with normal indexed-access semantics (|IndexError| when out of
        range); a |Slide| not belonging to this presentation raises |TargetNotFoundError|.
        """
        if isinstance(value, int) and not isinstance(value, bool):
            slide = self[value]
            _require_slide_enrolled(slide)
            return slide
        if isinstance(value, Slide):
            for slide in self:
                if slide == value:
                    _require_slide_enrolled(value)
                    return slide
            raise TargetNotFoundError(
                "slide with id %d is not in this presentation's slide collection" % value.slide_id
            )
        raise ValueError("expected a Slide or int index, got %r" % (value,))

    def get(self, slide_id: int, default: Slide | None = None) -> Slide | None:
        """Return the slide identified by int `slide_id` in this presentation.

        Returns `default` if not found.
        """
        slide = self.part.get_slide(slide_id)
        if slide is None:
            return default
        return slide

    def index(self, slide: Slide) -> int:
        """Map `slide` to its zero-based position in this slide sequence.

        Raises |ValueError| on *slide* not present.
        """
        for idx, this_slide in enumerate(self):
            if this_slide == slide:
                return idx
        raise ValueError("%s is not in slide collection" % slide)


class HeaderFooters(object):
    """Header/footer placeholder visibility flags of a layout or master (paper-pptx addition).

    Wraps the `p:hf` element. Each property is tri-state: |True|/|False| when the attribute
    is explicit, |None| when it is absent — meaning "inherit" (a layout inherits from its
    master; the schema default is visible). Assigning |None| removes the attribute.
    """

    def __init__(self, owner):
        """Bind to the layout or master that owns these flags."""
        super(HeaderFooters, self).__init__()
        self._owner = owner
        self._element = owner._element

    @property
    def slide_number_visible(self) -> bool | None:
        """Visibility of the slide-number placeholder (`p:hf/@sldNum`)."""
        hf = self._element.hf
        return hf.sldNum if hf is not None else None

    @slide_number_visible.setter
    def slide_number_visible(self, value: bool | None):
        self._set_flag("sldNum", value)

    @property
    def footer_visible(self) -> bool | None:
        """Visibility of the footer placeholder (`p:hf/@ftr`)."""
        hf = self._element.hf
        return hf.ftr if hf is not None else None

    @footer_visible.setter
    def footer_visible(self, value: bool | None):
        self._set_flag("ftr", value)

    @property
    def date_visible(self) -> bool | None:
        """Visibility of the date placeholder (`p:hf/@dt`)."""
        hf = self._element.hf
        return hf.dt if hf is not None else None

    @date_visible.setter
    def date_visible(self, value: bool | None):
        self._set_flag("dt", value)

    def _set_flag(self, attr_name: str, value: "bool | None") -> None:
        """Set one `p:hf` visibility flag inside a transaction.

        Refuses a detached element, and any value that is not True, False, or None.
        """
        from pptx._ownership import require_element_attached
        from pptx._transaction import PackageTransaction

        require_element_attached(
            self._element, self._owner.part, argument="header/footer flags"
        )
        if not isinstance(value, bool) and value is not None:
            raise ValueError("visibility must be True, False, or None, got %r" % (value,))
        with PackageTransaction(self._owner.part.package, self, self._owner):
            if value is None:
                hf = self._element.hf
                if hf is not None:
                    setattr(hf, attr_name, None)
                return
            setattr(self._element.get_or_add_hf(), attr_name, value)


class SlideLayout(_BaseSlide):
    """Slide layout object.

    Provides access to placeholders, regular shapes, and slide layout-level properties.
    """

    part: SlideLayoutPart  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def header_footers(self) -> HeaderFooters:
        """|HeaderFooters| flags for this layout (paper-pptx addition)."""
        return HeaderFooters(self)

    def iter_cloneable_placeholders(self) -> Iterator[LayoutPlaceholder]:
        """Generate layout-placeholders on this slide-layout that should be cloned to a new slide.

        Used when creating a new slide from this slide-layout.
        """
        latent_ph_types = (
            PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.SLIDE_NUMBER,
        )
        for ph in self.placeholders:
            if ph.element.ph_type not in latent_ph_types:
                yield ph

    @lazyproperty
    def placeholders(self) -> LayoutPlaceholders:
        """Sequence of placeholder shapes in this slide layout.

        Placeholders appear in `idx` order.
        """
        return LayoutPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self) -> LayoutShapes:
        """Sequence of shapes appearing on this slide layout."""
        return LayoutShapes(self._element.spTree, self)

    @property
    def slide_master(self) -> SlideMaster:
        """Slide master from which this slide-layout inherits properties."""
        return self.part.slide_master

    @property
    def used_by_slides(self):
        """Tuple of slide objects based on this slide layout."""
        # ---getting Slides collection requires going around the horn a bit---
        slides = self.part.package.presentation_part.presentation.slides
        return tuple(s for s in slides if s.slide_layout == self)


class SlideLayouts(ParentedElementProxy):
    """Sequence of slide layouts belonging to a slide-master.

    Supports indexed access, len(), iteration, index() and remove().
    """

    part: SlideMasterPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldLayoutIdLst: CT_SlideLayoutIdList, parent: SlideMaster):
        super(SlideLayouts, self).__init__(sldLayoutIdLst, parent)
        self._sldLayoutIdLst = sldLayoutIdLst

    def __getitem__(self, idx: int) -> SlideLayout:
        """Provides indexed access, e.g. `slide_layouts[2]`."""
        try:
            sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[idx]
        except IndexError:
            raise IndexError("slide layout index out of range")
        return self.part.related_slide_layout(sldLayoutId.rId)

    def __iter__(self) -> Iterator[SlideLayout]:
        """Generate each |SlideLayout| in the collection, in sequence."""
        for sldLayoutId in self._sldLayoutIdLst.sldLayoutId_lst:
            yield self.part.related_slide_layout(sldLayoutId.rId)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(slides) == 4`."""
        return len(self._sldLayoutIdLst)

    def get_by_name(self, name: str, default: SlideLayout | None = None) -> SlideLayout | None:
        """Return the first layout named `name`, or `default` when none is found.

        This preserves the upstream first-match contract. Use :meth:`get_unique_by_name` when the
        result will drive an edit and duplicate layout names must be surfaced rather than guessed.
        """
        for slide_layout in self:
            if slide_layout.name == name:
                return slide_layout
        return default

    def get_unique_by_name(self, name: str) -> SlideLayout:
        """Return the only layout named `name`; refuse when zero or multiple layouts match.

        Use this for mutation targeting because layout names are not required to be unique. A
        missing name raises |TargetNotFoundError| and duplicates raise |AmbiguousTargetError|;
        resolve ambiguity by selecting a layout explicitly by index or partname.
        """
        from pptx.errors import AmbiguousTargetError, TargetNotFoundError

        matches = [
            (idx, slide_layout)
            for idx, slide_layout in enumerate(self)
            if slide_layout.name == name
        ]
        if not matches:
            raise TargetNotFoundError("no slide layout in this master is named %r" % name)
        if len(matches) > 1:
            candidates = ", ".join(
                "%d (%s)" % (idx, slide_layout.part.partname) for idx, slide_layout in matches
            )
            raise AmbiguousTargetError(
                "%d slide layouts in this master are named %r; candidates: %s; select the "
                "intended layout explicitly by index or partname" % (len(matches), name, candidates)
            )
        return matches[0][1]

    def index(self, slide_layout: SlideLayout) -> int:
        """Return zero-based index of `slide_layout` in this collection.

        Raises `ValueError` if `slide_layout` is not present in this collection.
        """
        for idx, this_layout in enumerate(self):
            if slide_layout == this_layout:
                return idx
        raise ValueError("layout not in this SlideLayouts collection")

    def remove(self, slide_layout: SlideLayout) -> None:
        """Remove `slide_layout` from the collection.

        Raises ValueError when `slide_layout` is in use; a slide layout which is the basis for one
        or more slides cannot be removed.

        Refuses with TargetNotFoundError when the layout belongs to another presentation, and with
        UnsupportedStructureError when its part carries inbound relationships beyond this
        collection's own, where dropping it would strand whatever else points at it.
        """
        from pptx._transaction import PackageTransaction

        # Preserve the established error contract before attachment checks.
        if slide_layout.used_by_slides:
            raise ValueError("cannot remove slide-layout in use by one or more slides")

        # Upstream supports isolated collection proxies in its unit-level API contract.
        # A real presentation always supplies a parent and takes the hardened path below.
        if self._parent is None:
            target_idx = self.index(slide_layout)
            target_sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[target_idx]
            self._sldLayoutIdLst.remove(target_sldLayoutId)
            slide_layout.slide_master.part.drop_rel(target_sldLayoutId.rId)
            return

        if not isinstance(slide_layout, SlideLayout):
            raise ValueError("slide_layout must be a SlideLayout, got %r" % (slide_layout,))
        if slide_layout.part.package is not self.part.package:
            raise TargetNotFoundError("slide_layout belongs to a different presentation")
        _require_layout_enrolled(slide_layout)

        # ---target layout is identified by its index in this collection---
        target_idx = self.index(slide_layout)

        # --remove layout from p:sldLayoutIds of its master
        # --this stops layout from showing up, but doesn't remove it from package
        target_sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[target_idx]
        rId = target_sldLayoutId.rId
        aliases = [
            (owner, alias_rId)
            for owner, alias_rId, _ in _inbound_relationships(self.part.package, slide_layout.part)
            if owner is not self.part or alias_rId != rId
        ]
        if aliases:
            raise UnsupportedStructureError(
                "slide-layout removal refused: layout part has additional inbound "
                "relationship aliases"
            )

        with PackageTransaction(self.part.package, self, slide_layout):
            self._sldLayoutIdLst.remove(target_sldLayoutId)
            if _relationship_references(self.part._element, rId):
                raise UnsupportedStructureError(
                    "slide-layout removal refused: relationship %s remains referenced" % rId
                )
            # --drop relationship from master to layout
            # --this removes layout from package, along with everything (only) it refers to
            self.part.drop_rel(rId)


class SlideMaster(_BaseMaster):
    """Slide master object.

    Provides access to slide layouts. Access to placeholders, regular shapes, and slide master-level
    properties is inherited from |_BaseMaster|.
    """

    _element: CT_SlideMaster  # pyright: ignore[reportIncompatibleVariableOverride]

    @property
    def header_footers(self) -> HeaderFooters:
        """|HeaderFooters| flags for this master (paper-pptx addition)."""
        return HeaderFooters(self)

    @lazyproperty
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| object providing access to this slide-master's layouts."""
        return SlideLayouts(self._element.get_or_add_sldLayoutIdLst(), self)


class SlideMasters(ParentedElementProxy):
    """Sequence of |SlideMaster| objects belonging to a presentation.

    Has list access semantics, supporting indexed access, len(), and iteration.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldMasterIdLst: CT_SlideMasterIdList, parent: Presentation):
        super(SlideMasters, self).__init__(sldMasterIdLst, parent)
        self._sldMasterIdLst = sldMasterIdLst

    def __getitem__(self, idx: int) -> SlideMaster:
        """Provides indexed access, e.g. `slide_masters[2]`."""
        try:
            sldMasterId = self._sldMasterIdLst.sldMasterId_lst[idx]
        except IndexError:
            raise IndexError("slide master index out of range")
        return self.part.related_slide_master(sldMasterId.rId)

    def __iter__(self):
        """Generate each |SlideMaster| instance in the collection, in sequence."""
        for smi in self._sldMasterIdLst.sldMasterId_lst:
            yield self.part.related_slide_master(smi.rId)

    def __len__(self):
        """Support len() built-in function, e.g. `len(slide_masters) == 4`."""
        return len(self._sldMasterIdLst)


class _Background(ElementProxy):
    """Provides access to slide background properties.

    Note that the presence of this object does not by itself imply an
    explicitly-defined background; a slide with an inherited background still
    has a |_Background| object.
    """

    def __init__(self, cSld: CT_CommonSlideData):
        super(_Background, self).__init__(cSld)
        self._cSld = cSld

    @lazyproperty
    def fill(self):
        """|FillFormat| instance for this background.

        This |FillFormat| object is used to interrogate or specify the fill
        of the slide background.

        Note that accessing this property is potentially destructive. A slide
        background can also be specified by a background style reference and
        accessing this property will remove that reference, if present, and
        replace it with NoFill. This is frequently the case for a slide
        master background.

        This is also the case when there is no explicitly defined background
        (background is inherited); merely accessing this property will cause
        the background to be set to NoFill and the inheritance link will be
        interrupted. This is frequently the case for a slide background.

        Of course, if you are accessing this property in order to set the
        fill, then these changes are of no consequence, but the existing
        background cannot be reliably interrogated using this property unless
        you have already established it is an explicit fill.

        If the background is already a fill, then accessing this property
        makes no changes to the current background.
        """
        bgPr = self._cSld.get_or_add_bgPr()
        return FillFormat.from_fill_parent(bgPr)
