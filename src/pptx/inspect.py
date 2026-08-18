"""Effective-style inspection: resolved font values with provenance (paper-pptx).

Upstream python-pptx reports `None` for any character property that is inherited, leaving
callers blind to what a deck actually renders. This module implements the documented
inheritance walk for **font size, font name, and font color** and reports, for every resolved
value, the ordered chain of sources consulted (read-only inspection is provenance-bearing).

The pinned walk, per run (levels consulted in order until one supplies the value):

1. the run's own `a:rPr`
2. the containing paragraph's `a:pPr/a:defRPr`
3. the shape's own `p:txBody/a:lstStyle` at the paragraph's indent level
4. placeholder shapes: the layout placeholder's `lstStyle` (matched by `idx`, falling back to
   `type`), then the master placeholder's `lstStyle` (title-family → master title placeholder,
   otherwise master body placeholder), then the master's `p:txStyles` family style
   (title/ctrTitle → `titleStyle`; body/subTitle/obj and vertical variants → `bodyStyle`;
   anything else → `otherStyle`)
5. non-placeholder shapes: the shape's `p:style/a:fontRef`, then the presentation's
   `p:defaultTextStyle`

Font-name theme references (`+mj-lt`/`+mn-lt`) resolve through the master's theme part
(`a:fontScheme`). Scheme colors resolve through the slide's `p:clrMapOvr` override when
present, else the master's `p:clrMap`, then the theme's `a:clrScheme` (`a:sysClr` uses its
`lastClr`). Anything genuinely outside this walk reports `resolved=False` with the consulted
chain — never a guessed default.

Everything here is strictly read-only: no method mutates any element, and raw-XML reads are
used throughout (several upstream proxy accessors are get-or-add and would dirty the tree).
"""

from __future__ import annotations

import copy
import hashlib
import unicodedata
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterator, Optional, Tuple

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.errors import UnsupportedStructureError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_TextBulletSizePercent
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Centipoints, Length

if TYPE_CHECKING:
    from pptx.slide import Slide
    from pptx.text.text import _Run

__all__ = [
    "BULLET_FOLLOWS_TEXT",
    "BlockAnchor",
    "DECK_MANIFEST_SCHEMA",
    "DECK_MANIFEST_VERSION",
    "DeckManifest",
    "EffectiveBullet",
    "EffectiveFont",
    "EffectiveParagraphFormat",
    "EffectiveShapeFormat",
    "EffectiveValue",
    "InspectedRun",
    "ProvenanceStep",
    "SCHEMA_NAME",
    "SCHEMA_VERSION",
    "ShapeManifest",
    "SlideManifest",
    "TextBlock",
    "TextInspection",
    "content_hash",
    "effective_font",
    "effective_paragraph_format",
    "effective_shape_format",
    "inspect_deck",
    "inspect_text",
    "iter_text_bodies",
]

SCHEMA_NAME = "paper-text-inspection"
SCHEMA_VERSION = 2  # -- visibility-complete traversal, container/blind fields

#: Resolved value of a bullet typeface or size that defers to the run's own text, whether an
#: `a:buFontTx` / `a:buSzTx` said so explicitly or the whole chain went silent. A sentinel
#: rather than |None|, because "follow the text" is an answer: it renders predictably, and
#: |None| is reserved for values the resolver could not determine.
BULLET_FOLLOWS_TEXT = "follow text"

_TITLE_FAMILY = frozenset([PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE])
_BODY_FAMILY = frozenset(
    [
        PP_PLACEHOLDER.BODY,
        PP_PLACEHOLDER.SUBTITLE,
        PP_PLACEHOLDER.OBJECT,
        PP_PLACEHOLDER.VERTICAL_BODY,
        PP_PLACEHOLDER.VERTICAL_OBJECT,
    ]
)
#: schemeClr tokens addressing theme slots directly, bypassing the clrMap indirection
_DIRECT_THEME_SLOTS = frozenset(
    ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
     "accent6", "hlink", "folHlink"]
)


def content_hash(text: str) -> str:
    """First 8 hex chars of SHA-256 over the NFC-normalized text (the pinned anchor hash).

    Unicode normalization only — whitespace is content and is never trimmed.
    """
    normalized = unicodedata.normalize("NFC", text)
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:8]


@dataclass(frozen=True)
class ProvenanceStep:
    """One consulted level of an inheritance walk."""

    level: str  #: e.g. "layout placeholder lstStyle lvl1"
    part: Optional[str]  #: partname consulted, |None| for in-part levels
    detail: str  #: what was found (or not) at this level
    supplied: bool  #: True on the step that supplied the resolved value

    def to_dict(self) -> dict:
        return {
            "level": self.level,
            "part": self.part,
            "detail": self.detail,
            "supplied": self.supplied,
        }


@dataclass(frozen=True)
class EffectiveValue:
    """A resolved (or honestly-unresolved) effective value with its provenance chain."""

    value: object  #: EMU int for sizes, str for names, "RRGGBB" str for colors; None if unresolved
    value_pt: Optional[float]  #: convenience for sizes only, never instead of the EMU value
    resolved: bool
    provenance: Tuple[ProvenanceStep, ...]
    bake_color_xml: Optional[bytes] = field(default=None, repr=False, compare=False)

    def to_dict(self) -> dict:
        return {
            "value": (
                int(self.value)
                if isinstance(self.value, int) and not isinstance(self.value, bool)
                else self.value
            ),
            "value_pt": self.value_pt,
            "resolved": self.resolved,
            "provenance": [step.to_dict() for step in self.provenance],
        }


@dataclass(frozen=True)
class EffectiveFont:
    """Effective size, name, color, and emphasis of one run."""

    size: EffectiveValue
    name: EffectiveValue
    color_rgb: EffectiveValue
    bold: EffectiveValue = None  # pyright: ignore[reportAssignmentType]
    italic: EffectiveValue = None  # pyright: ignore[reportAssignmentType]
    underline: EffectiveValue = None  # pyright: ignore[reportAssignmentType]

    def to_dict(self) -> dict:
        return {
            "schema": "paper-effective-font",
            "version": 2,  # -- bold/italic/underline added
            "size": self.size.to_dict(),
            "name": self.name.to_dict(),
            "color_rgb": self.color_rgb.to_dict(),
            "bold": self.bold.to_dict() if self.bold is not None else None,
            "italic": self.italic.to_dict() if self.italic is not None else None,
            "underline": self.underline.to_dict() if self.underline is not None else None,
        }


@dataclass(frozen=True)
class EffectiveBullet:
    """The bullet that actually renders on one paragraph, with its provenance chain."""

    type: Optional[str]  #: "none", "character", "numbered", "picture"; |None| if unresolved
    char: Optional[str]  #: the glyph, character bullets only
    number_scheme: Optional[str]  #: ECMA autonumber token, e.g. "arabicPeriod"; numbered only
    start_at: Optional[int]  #: first number of the sequence, numbered only
    resolved: bool
    provenance: Tuple[ProvenanceStep, ...]

    def to_dict(self) -> dict:
        return {
            "type": self.type,
            "char": self.char,
            "number_scheme": self.number_scheme,
            "start_at": self.start_at,
            "resolved": self.resolved,
            "provenance": [step.to_dict() for step in self.provenance],
        }


@dataclass(frozen=True)
class BlockAnchor:
    """The pinned anchor shape: part + block index + content hash (detects staleness).

    Fields:

    * ``part`` -- partname of the story the block lives in (a slide or notes slide).
    * ``block_index`` -- 0-based index of the paragraph block within that part, in the
      visibility-complete traversal order.
    * ``content_hash`` -- first 8 hex chars of the SHA-256 of the block's NFC-normalized
      text; a mismatch on re-inspection is how a stale anchor is detected.
    """

    part: str
    block_index: int
    content_hash: str

    def to_dict(self) -> dict:
        return {
            "part": self.part,
            "block_index": self.block_index,
            "content_hash": self.content_hash,
        }


@dataclass(frozen=True)
class InspectedRun:
    text: str
    font: EffectiveFont
    field_type: Optional[str] = None

    def to_dict(self) -> dict:
        payload = {"text": self.text, "font": self.font.to_dict()}
        if self.field_type is not None:
            payload["field_type"] = self.field_type
        return payload


@dataclass(frozen=True)
class TextBlock:
    """One paragraph of one text body, with anchor and per-run effective fonts.

    `container` says where the text lives: "shape" (a top-level `p:sp`), "group" (a `p:sp`
    inside `p:grpSp` nesting; `container_detail` is the slash-joined group-name path), or
    "table-cell" (`container_detail` is `"<frame-name>!r{row}c{col}"`). `blind` is True when
    the block's *text* is visible but its effective values are unresolvable by design in
    this version (table-cell runs inherit through table styles, a chain not yet walked);
    a blind block's runs carry `resolved=False` values, never guesses.
    """

    anchor: BlockAnchor
    shape_id: int
    shape_name: str
    placeholder_type: Optional[str]
    level: int
    text: str
    runs: Tuple[InspectedRun, ...]
    container: str = "shape"
    container_detail: Optional[str] = None
    blind: bool = False
    fields: Tuple[str, ...] = ()

    def to_dict(self) -> dict:
        return {
            "anchor": self.anchor.to_dict(),
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "placeholder_type": self.placeholder_type,
            "container": self.container,
            "container_detail": self.container_detail,
            "blind": self.blind,
            "fields": list(self.fields),
            "level": self.level,
            "text": self.text,
            "runs": [run.to_dict() for run in self.runs],
        }


@dataclass(frozen=True)
class TextInspection:
    """Inspection payload for one slide. `.to_dict()` is deterministic (golden-tested)."""

    part: str
    blocks: Tuple[TextBlock, ...] = field(default_factory=tuple)

    @property
    def blind_region_count(self) -> int:
        """Number of blocks whose effective values are unresolvable by design (see TextBlock)."""
        return sum(1 for block in self.blocks if block.blind)

    def to_dict(self) -> dict:
        return {
            "schema": SCHEMA_NAME,
            "version": SCHEMA_VERSION,
            "part": self.part,
            "blind_region_count": self.blind_region_count,
            "blocks": [block.to_dict() for block in self.blocks],
        }


def effective_font(run: "_Run") -> EffectiveFont:
    """Return the `EffectiveFont` for `run`, a `_Run` on a slide shape.

    Raises `UnsupportedStructureError` for runs outside a `p:sp` shape on a slide part
    (table-cell and chart text are not resolved).
    """
    r = run._r
    sp = _ancestor_sp(r)
    if sp is None:
        raise UnsupportedStructureError(
            "effective_font resolves runs inside p:sp shapes only in v0 (got a run outside"
            " any p:sp, e.g. table-cell or chart text)"
        )
    part = run.part
    if not hasattr(part, "slide_layout"):
        raise UnsupportedStructureError(
            "effective_font resolves runs on slide parts only in v0, got %r"
            % type(part).__name__
        )
    return _FontResolver(part).effective_font(r, sp)


@dataclass(frozen=True)
class EffectiveParagraphFormat:
    """Effective alignment, line spacing, and bullet of one paragraph (paper-pptx addition)."""

    alignment: EffectiveValue  #: ECMA algn token, e.g. "l", "ctr", "r", "just"
    line_spacing: EffectiveValue  #: float lines (1.0 = single) or EMU |Length| for points
    bullet: EffectiveBullet  #: the bullet that renders, resolved through the same chain
    bullet_font: EffectiveValue  #: bullet typeface, or |BULLET_FOLLOWS_TEXT|
    bullet_size: EffectiveValue  #: fraction of text size, |Length| for points, or follows text

    def to_dict(self) -> dict:
        return {
            "schema": "paper-effective-paragraph-format",
            "version": 3,  # -- was 2: bullet_font/bullet_size added
            "alignment": self.alignment.to_dict(),
            "line_spacing": self.line_spacing.to_dict(),
            "bullet": self.bullet.to_dict(),
            "bullet_font": self.bullet_font.to_dict(),
            "bullet_size": self.bullet_size.to_dict(),
        }


def effective_paragraph_format(paragraph) -> EffectiveParagraphFormat:
    """Return effective alignment, line spacing, and bullet for `paragraph`, with provenance.

    Same inheritance walk as `effective_font`, over paragraph-level properties: the
    paragraph's own `a:pPr`, the shape's `lstStyle` level entry, the placeholder chain (or
    the presentation's `defaultTextStyle`). Exhaustion resolves rather than reporting
    unresolved: to the ECMA-376 schema default for alignment (left), and to the rendering
    conventions for line spacing (single), bullet (no bullet), and bullet typeface/size
    (follow the text), none of which the schema defines.

    The bullet's kind, typeface, and size are three separate XSD choice groups that inherit
    independently, so each is walked on its own and carries its own provenance. A paragraph
    can take its `buChar` from the master and its `buSzPct` from the layout.
    """
    p = paragraph._p
    txBody = p.getparent()
    sp = txBody.getparent() if txBody is not None else None
    if sp is None or sp.tag != qn("p:sp"):
        raise UnsupportedStructureError(
            "effective_paragraph_format resolves paragraphs inside p:sp shapes only in v0.1"
        )
    part = paragraph.part
    from pptx.parts.slide import SlidePart

    if not isinstance(part, SlidePart):
        raise UnsupportedStructureError(
            "effective_paragraph_format resolves paragraphs on slide parts only in v0.1,"
            " got %r" % type(part).__name__
        )
    resolver = _FontResolver(part)
    pPr = p.find(qn("a:pPr"))
    level = int(pPr.get("lvl", "0")) if pPr is not None else 0
    if not 0 <= level <= 8:
        raise UnsupportedStructureError(
            "paragraph carries out-of-schema indent level lvl=%d (valid range 0..8)" % level
        )
    chain = list(resolver.paragraph_chain(p, sp, level))
    return EffectiveParagraphFormat(
        alignment=resolver.resolve_alignment(chain),
        line_spacing=resolver.resolve_line_spacing(chain),
        bullet=resolver.resolve_bullet(chain),
        bullet_font=resolver.resolve_bullet_font(chain),
        bullet_size=resolver.resolve_bullet_size(chain),
    )


DECK_MANIFEST_SCHEMA = "paper-deck-manifest"
DECK_MANIFEST_VERSION = 1


@dataclass(frozen=True)
class ShapeManifest:
    """Structural facts of one shape (paper-pptx addition). Group children nest."""

    shape_id: int
    name: str
    kind: str  #: e.g. "PICTURE", "TEXT_BOX", "PLACEHOLDER", "GROUP", "GraphicFrame"
    z_index: int  #: position within the containing collection (0 = backmost)
    placeholder_type: Optional[str]
    x: Optional[int]  #: EMU; placeholder geometry is upstream-inherited; None = unresolvable
    y: Optional[int]
    cx: Optional[int]
    cy: Optional[int]
    rotation: Optional[float]
    text_block_count: int
    autofit: Optional[dict]  #: {"mode", "font_scale", "line_space_reduction"} or None
    table: Optional[dict]  #: {"rows", "cols"} or None
    chart: Optional[dict]  #: {"chart_type", "series_names"} or None
    image: Optional[dict]  #: {"ext", "natural_size_px", "displayed_size_emu"} or None
    children: Tuple["ShapeManifest", ...] = ()

    def to_dict(self) -> dict:
        return {
            "shape_id": self.shape_id,
            "name": self.name,
            "kind": self.kind,
            "z_index": self.z_index,
            "placeholder_type": self.placeholder_type,
            "x": self.x,
            "y": self.y,
            "cx": self.cx,
            "cy": self.cy,
            "rotation": self.rotation,
            "text_block_count": self.text_block_count,
            "autofit": self.autofit,
            "table": self.table,
            "chart": self.chart,
            "image": self.image,
            "children": [child.to_dict() for child in self.children],
        }


@dataclass(frozen=True)
class SlideManifest:
    """Structural facts of one slide."""

    part: str
    slide_id: int
    layout_name: str
    has_notes: bool
    shapes: Tuple[ShapeManifest, ...]
    alternate_content_count: int = 0  #: mc:AlternateContent subtrees (not surveyable)

    def to_dict(self) -> dict:
        return {
            "part": self.part,
            "slide_id": self.slide_id,
            "layout_name": self.layout_name,
            "has_notes": self.has_notes,
            "alternate_content_count": self.alternate_content_count,
            "shapes": [shape.to_dict() for shape in self.shapes],
        }


@dataclass(frozen=True)
class DeckManifest:
    """Structural survey of a whole deck. `.to_dict()` is deterministic (golden-tested)."""

    slide_width: Optional[int]  #: EMU
    slide_height: Optional[int]
    slides: Tuple[SlideManifest, ...]
    masters: Tuple[dict, ...]  #: [{"part", "layouts": [names]}]

    @property
    def slide_count(self) -> int:
        return len(self.slides)

    def to_dict(self) -> dict:
        return {
            "schema": DECK_MANIFEST_SCHEMA,
            "version": DECK_MANIFEST_VERSION,
            "slide_width": self.slide_width,
            "slide_height": self.slide_height,
            "slide_count": self.slide_count,
            "slides": [slide.to_dict() for slide in self.slides],
            "masters": [dict(master) for master in self.masters],
        }


def inspect_deck(prs) -> DeckManifest:
    """Return a structural `DeckManifest` of `prs` (paper-pptx addition).

    The survey every brownfield edit starts with, as one deterministic typed payload:
    per-slide shape inventory (identity, kind, z-order, geometry where explicit, placeholder
    role, table/chart/image/autofit facts), group children nested, layout and master
    inventory. Values that are inherited rather than explicit report None — never a guess.
    Read-only.
    """
    from pptx.errors import materialize_slides

    slides = []
    for slide in materialize_slides(prs, "inspect_deck"):
        shapes = tuple(
            _shape_manifest(shape, z_index) for z_index, shape in enumerate(slide.shapes)
        )
        slides.append(
            SlideManifest(
                part=str(slide.part.partname),
                slide_id=slide.slide_id,
                layout_name=slide.slide_layout.name,
                has_notes=slide.has_notes_slide,
                shapes=shapes,
                alternate_content_count=len(
                    slide._element.spTree.findall(".//" + _MC_ALTERNATE_CONTENT)
                ),
            )
        )
    masters = tuple(
        {
            "part": str(master.part.partname),
            "layouts": [layout.name for layout in master.slide_layouts],
        }
        for master in prs.slide_masters
    )
    return DeckManifest(
        slide_width=int(prs.slide_width) if prs.slide_width is not None else None,
        slide_height=int(prs.slide_height) if prs.slide_height is not None else None,
        slides=tuple(slides),
        masters=masters,
    )


def _shape_manifest(shape, z_index: int) -> ShapeManifest:
    from pptx.shapes.group import GroupShape
    from pptx.shapes.picture import Picture
    from pptx.shapes.shapetree import _shape_kind

    placeholder_type = None
    if shape.is_placeholder:
        ph_type = shape.placeholder_format.type
        placeholder_type = ph_type.name if ph_type is not None else None

    autofit = None
    text_block_count = 0
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        text_frame = shape.text_frame
        text_block_count = len(text_frame.paragraphs)
        mode = text_frame.auto_size
        if mode is not None or text_frame.font_scale is not None:
            autofit = {
                "mode": mode.name if mode is not None else None,
                "font_scale": text_frame.font_scale,
                "line_space_reduction": text_frame.line_space_reduction,
            }

    table = None
    if getattr(shape, "has_table", False) and shape.has_table:
        tbl = shape.table
        table = {"rows": len(tbl.rows), "cols": len(tbl.columns)}

    chart = None
    if getattr(shape, "has_chart", False) and shape.has_chart:
        chart_obj = shape.chart
        chart = {
            "chart_type": chart_obj.chart_type.name,
            "series_names": [series.name for series in chart_obj.series],
        }

    image = None
    if isinstance(shape, Picture):
        try:
            img = shape.image
            image = {
                "ext": img.ext,
                "natural_size_px": list(img.size),
                "displayed_size_emu": [
                    int(shape.width) if shape.width is not None else None,
                    int(shape.height) if shape.height is not None else None,
                ],
            }
        except (KeyError, ValueError):
            image = {"ext": None, "natural_size_px": None, "displayed_size_emu": None}

    children = ()
    if isinstance(shape, GroupShape):
        children = tuple(
            _shape_manifest(child, child_index)
            for child_index, child in enumerate(shape.shapes)
        )

    def emu_or_none(value):
        return int(value) if value is not None else None

    try:
        rotation = float(shape.rotation)
    except (AttributeError, NotImplementedError):
        rotation = None

    return ShapeManifest(
        shape_id=shape.shape_id,
        name=shape.name,
        kind=_shape_kind(shape),
        z_index=z_index,
        placeholder_type=placeholder_type,
        x=emu_or_none(shape.left),
        y=emu_or_none(shape.top),
        cx=emu_or_none(shape.width),
        cy=emu_or_none(shape.height),
        rotation=rotation,
        text_block_count=text_block_count,
        autofit=autofit,
        table=table,
        chart=chart,
        image=image,
        children=children,
    )


@dataclass(frozen=True)
class EffectiveShapeFormat:
    """Effective solid fill and line color of one shape (paper-pptx addition).

    It resolves EXPLICIT `p:spPr` fills fully (solid colors through the scheme/clrMap/
    theme walk; "none" for noFill). A shape whose fill comes only from its `p:style`
    fill/line reference reports unresolved — the provenance carries the reference index and
    its resolved phClr color, but theme format-scheme modulation is not applied (guessing it
    would violate fail-loudly).
    """

    fill_rgb: EffectiveValue
    line_rgb: EffectiveValue

    def to_dict(self) -> dict:
        return {
            "schema": "paper-effective-shape-format",
            "version": 1,
            "fill_rgb": self.fill_rgb.to_dict(),
            "line_rgb": self.line_rgb.to_dict(),
        }


def effective_shape_format(shape) -> EffectiveShapeFormat:
    """Return the effective solid fill and line color of `shape`, with provenance."""
    from pptx.parts.slide import SlidePart

    part = shape.part
    if not isinstance(part, SlidePart):
        raise UnsupportedStructureError(
            "effective_shape_format resolves shapes on slide parts only in v0.1, got %r"
            % type(part).__name__
        )
    resolver = _FontResolver(part)
    sp = shape._element
    partname = str(part.partname)
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        spPr = sp.find(qn("p:grpSpPr"))
    style = sp.find(qn("p:style"))

    fill_rgb = resolver.resolve_container_fill(
        spPr,
        "shape spPr",
        partname,
        style.find(qn("a:fillRef")) if style is not None else None,
        "shape style fillRef",
    )
    ln = spPr.find(qn("a:ln")) if spPr is not None else None
    line_rgb = resolver.resolve_container_fill(
        ln,
        "shape spPr a:ln",
        partname,
        style.find(qn("a:lnRef")) if style is not None else None,
        "shape style lnRef",
    )
    return EffectiveShapeFormat(fill_rgb=fill_rgb, line_rgb=line_rgb)


def inspect_text(slide: "Slide") -> TextInspection:
    """Return a `TextInspection` of every text block on `slide`, visibility-complete.

    Traversal is depth-first document order over the shape tree: top-level `p:sp` shapes,
    `p:sp` shapes inside groups (recursively, to any depth), and table cells
    (row-major within each table graphic-frame). `block_index` numbers blocks consecutively
    in that pinned order. Table-cell blocks report their text but are *blind regions* for
    effective values (see `TextBlock`); chart text lives in the chart part, not the slide
    part, and is out of scope here.
    """
    part = slide.part
    partname = str(part.partname)
    resolver = _FontResolver(part)
    blocks: list = []
    _walk_container(
        slide._element.spTree, (), partname, resolver, blocks
    )
    return TextInspection(part=partname, blocks=tuple(blocks))


def iter_text_bodies(spTree):
    """Yield `(kind, owner_elm, txBody, group_path, cell_detail)` depth-first, document order.

    The single source of traversal truth shared by `inspect_text` and `pptx.edit`
    (visibility-complete: top-level `p:sp`, grouped `p:sp` recursively to any depth,
    table cells row-major). `kind` is "shape" | "group" | "table-cell"; `owner_elm` is the
    `p:sp` or `p:graphicFrame`; `cell_detail` is `"<frame-name>!r{row}c{col}"` for table
    cells, None otherwise.
    """
    return _iter_container(spTree, ())


_MC_ALTERNATE_CONTENT = (
    "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
)


def _iter_container(container_elm, group_path):
    for child in container_elm:
        if child.tag == _MC_ALTERNATE_CONTENT:
            # -- markup-compatibility content renders one of several branches depending on
            # -- the consumer; reporting any single branch as "the" text would be a guess.
            # -- Yield a marker so consumers can report a typed blind region (§1.5), never
            # -- silence. One marker = one block index.
            yield "alternate-content", child, None, group_path, None
        elif child.tag == qn("p:sp"):
            txBody = child.find(qn("p:txBody"))
            if txBody is not None:
                yield ("group" if group_path else "shape"), child, txBody, group_path, None
        elif child.tag == qn("p:grpSp"):
            for item in _iter_container(child, group_path + (_cNvPr_name(child),)):
                yield item
        elif child.tag == qn("p:graphicFrame"):
            tbl = child.find(".//%s" % qn("a:tbl"))
            if tbl is None:
                kind = (
                    "chart"
                    if child.find(".//%s" % qn("c:chart")) is not None
                    else "graphic-frame"
                )
                yield kind, child, None, group_path, _cNvPr_name(child)
                continue
            frame_name = _cNvPr_name(child)
            for row_index, tr in enumerate(tbl.findall(qn("a:tr"))):
                for col_index, tc in enumerate(tr.findall(qn("a:tc"))):
                    txBody = tc.find(qn("a:txBody"))
                    if txBody is not None:
                        detail = "%s!r%dc%d" % (frame_name, row_index, col_index)
                        yield "table-cell", child, txBody, group_path, detail


def _walk_container(spTree, group_path, partname, resolver, blocks) -> None:
    """Append TextBlocks for every text body under `spTree` (see iter_text_bodies)."""
    for kind, owner, txBody, path, cell_detail in iter_text_bodies(spTree):
        if kind in ("alternate-content", "chart", "graphic-frame"):
            cNvPr = owner.find(".//%s" % qn("p:cNvPr"))
            blocks.append(
                TextBlock(
                    anchor=BlockAnchor(partname, len(blocks), content_hash("")),
                    shape_id=int(cNvPr.get("id")) if cNvPr is not None else 0,
                    shape_name=(cNvPr.get("name") or "") if cNvPr is not None else "",
                    placeholder_type=None,
                    level=0,
                    text="",
                    runs=(),
                    container=kind,
                    container_detail=cell_detail or ("/".join(path) if path else None),
                    blind=True,
                )
            )
            continue
        if kind == "table-cell":
            for p in txBody.findall(qn("a:p")):
                runs = tuple(
                    InspectedRun(
                        (
                            ""
                            if item.tag == qn("a:fld")
                            else "\v"
                            if item.tag == qn("a:br")
                            else _run_text(item)
                        ),
                        _blind_font("table-cell"),
                        (item.get("type") or "") if item.tag == qn("a:fld") else None,
                    )
                    for item in p
                    if item.tag in (qn("a:r"), qn("a:fld"), qn("a:br"))
                )
                _append_block(
                    blocks, partname, owner, p, runs, None, "table-cell", cell_detail, True
                )
        else:
            placeholder_type = None
            if kind == "shape" and owner.has_ph_elm:
                ph_type = owner.ph_type
                placeholder_type = ph_type.name if ph_type is not None else None
            container_detail = "/".join(path) if path else None
            for p in txBody.findall(qn("a:p")):
                runs = tuple(
                    InspectedRun(
                        (
                            ""
                            if item.tag == qn("a:fld")
                            else "\v"
                            if item.tag == qn("a:br")
                            else _run_text(item)
                        ),
                        resolver.effective_font(item, owner),
                        (item.get("type") or "") if item.tag == qn("a:fld") else None,
                    )
                    for item in p
                    if item.tag in (qn("a:r"), qn("a:fld"), qn("a:br"))
                )
                _append_block(
                    blocks, partname, owner, p, runs, placeholder_type, kind,
                    container_detail, False,
                )


def _append_block(
    blocks, partname, shape_elm, p, runs, placeholder_type, container, container_detail, blind
) -> None:
    text = "".join(inspected.text for inspected in runs if inspected.field_type is None)
    pPr = p.find(qn("a:pPr"))
    level = int(pPr.get("lvl", "0")) if pPr is not None else 0
    cNvPr = shape_elm.find(".//%s" % qn("p:cNvPr"))
    # -- fields (a:fld) are recognized by type but excluded from text/hash: their display
    # -- text is volatile (PowerPoint re-renders it), so hashing it would rot anchors
    fields = tuple(run.field_type for run in runs if run.field_type is not None)
    blocks.append(
        TextBlock(
            anchor=BlockAnchor(partname, len(blocks), content_hash(text)),
            shape_id=int(cNvPr.get("id")),
            shape_name=cNvPr.get("name") or "",
            placeholder_type=placeholder_type,
            level=level,
            text=text,
            runs=runs,
            container=container,
            container_detail=container_detail,
            blind=blind,
            fields=fields,
        )
    )


def _cNvPr_name(shape_elm) -> str:
    cNvPr = shape_elm.find(".//%s" % qn("p:cNvPr"))
    return (cNvPr.get("name") or "") if cNvPr is not None else ""


def _run_text(r) -> str:
    t = r.find(qn("a:t"))
    return t.text or "" if t is not None else ""


def _blind_font(region_kind: str) -> EffectiveFont:
    """Return an EffectiveFont whose values are honestly unresolved for a blind region."""
    step = ProvenanceStep(
        level="%s text" % region_kind,
        part=None,
        detail="%s inheritance (e.g. table styles) is not resolved in v0.1; text is"
        " reported, values are not guessed" % region_kind,
        supplied=False,
    )
    unresolved = EffectiveValue(value=None, value_pt=None, resolved=False, provenance=(step,))
    return EffectiveFont(
        size=unresolved,
        name=unresolved,
        color_rgb=unresolved,
        bold=unresolved,
        italic=unresolved,
        underline=unresolved,
    )


def _ancestor_sp(element):
    """Return the nearest `p:sp` ancestor of `element`, or None."""
    parent = element.getparent()
    sp_tag = qn("p:sp")
    while parent is not None and parent.tag != sp_tag:
        parent = parent.getparent()
    return parent


class _FontResolver(object):
    """Executes the pinned inheritance walk for runs on one slide part. Strictly read-only."""

    def __init__(self, slide_part):
        self._slide_part = slide_part
        self._layout = slide_part.slide_layout
        self._master = self._layout.slide_master
        self._theme_root, self._theme_partname = self._load_theme()

    # ------------------------------------------------------------------------- public

    def effective_font(self, r, sp) -> EffectiveFont:
        p = r.getparent()
        pPr = p.find(qn("a:pPr"))
        level = int(pPr.get("lvl", "0")) if pPr is not None else 0
        if not 0 <= level <= 8:
            raise UnsupportedStructureError(
                "paragraph carries out-of-schema indent level lvl=%d (valid range 0..8);"
                " refusing to resolve styles for malformed structure" % level
            )
        chain = list(self._chain(r, p, sp, level))
        return EffectiveFont(
            size=self._resolve_size(chain),
            name=self._resolve_name(chain, "" if r.tag == qn("a:fld") else _run_text(r)),
            color_rgb=self._resolve_color(chain, sp),
            bold=self._resolve_flag(chain, "b", "bold"),
            italic=self._resolve_flag(chain, "i", "italic"),
            underline=self._resolve_underline(chain),
        )

    # ------------------------------------------------------------------- chain assembly

    def _chain(self, r, p, sp, level) -> Iterator[Tuple[str, str, object]]:
        """Generate (level-label, partname, rPr-like element or None) in consultation order."""
        slide_partname = str(self._slide_part.partname)
        yield "run rPr", slide_partname, r.find(qn("a:rPr"))

        pPr = p.find(qn("a:pPr"))
        yield (
            "paragraph defRPr",
            slide_partname,
            pPr.find(qn("a:defRPr")) if pPr is not None else None,
        )

        txBody = p.getparent()
        yield (
            "shape lstStyle lvl%d" % (level + 1),
            slide_partname,
            self._defRPr_from_lstStyle(txBody.find(qn("a:lstStyle")), level),
        )

        ph = sp.find(qn("p:nvSpPr") + "/" + qn("p:nvPr") + "/" + qn("p:ph"))
        if ph is not None:
            for step in self._placeholder_chain(sp, level):
                yield step
        else:
            yield "shape fontRef", slide_partname, self._rPr_from_fontRef(sp)
            yield (
                "presentation defaultTextStyle lvl%d" % (level + 1),
                str(self._presentation_part.partname),
                self._defRPr_from_lstStyle(
                    self._presentation_part._element.find(qn("p:defaultTextStyle")), level
                ),
            )

    def _placeholder_chain(self, sp, level):
        idx, ph_type = sp.ph_idx, sp.ph_type
        layout_ph = self._layout_placeholder(idx, ph_type)
        yield (
            "layout placeholder lstStyle lvl%d" % (level + 1),
            str(self._layout.part.partname),
            self._ph_lstStyle_defRPr(layout_ph, level),
        )

        master_ph_type = self._master_placeholder_type(ph_type)
        master_ph = self._master_placeholder(master_ph_type)
        yield (
            "master placeholder lstStyle lvl%d" % (level + 1),
            str(self._master.part.partname),
            self._ph_lstStyle_defRPr(master_ph, level),
        )

        family, style = self._master_family_style(ph_type)
        yield (
            "master txStyles %s lvl%d" % (family, level + 1),
            str(self._master.part.partname),
            self._defRPr_from_lstStyle(style, level),
        )

    def paragraph_chain(self, p, sp, level):
        """Generate (label, partname, pPr-like element) for paragraph-level properties."""
        slide_partname = str(self._slide_part.partname)
        yield "paragraph pPr", slide_partname, p.find(qn("a:pPr"))
        txBody = p.getparent()
        yield (
            "shape lstStyle lvl%d" % (level + 1),
            slide_partname,
            self._pPr_from_lstStyle(txBody.find(qn("a:lstStyle")), level),
        )
        ph = sp.find(qn("p:nvSpPr") + "/" + qn("p:nvPr") + "/" + qn("p:ph"))
        if ph is not None:
            idx, ph_type = sp.ph_idx, sp.ph_type
            layout_ph = self._layout_placeholder(idx, ph_type)
            yield (
                "layout placeholder lstStyle lvl%d" % (level + 1),
                str(self._layout.part.partname),
                self._ph_lstStyle_pPr(layout_ph, level),
            )
            master_ph_type = self._master_placeholder_type(ph_type)
            master_ph = self._master_placeholder(master_ph_type)
            yield (
                "master placeholder lstStyle lvl%d" % (level + 1),
                str(self._master.part.partname),
                self._ph_lstStyle_pPr(master_ph, level),
            )
            family, style = self._master_family_style(ph_type)
            yield (
                "master txStyles %s lvl%d" % (family, level + 1),
                str(self._master.part.partname),
                style.pPr_for_lvl(level) if style is not None else None,
            )
        else:
            defaultTextStyle = self._presentation_part._element.find(
                qn("p:defaultTextStyle")
            )
            yield (
                "presentation defaultTextStyle lvl%d" % (level + 1),
                str(self._presentation_part.partname),
                self._pPr_from_lstStyle(defaultTextStyle, level),
            )

    def resolve_alignment(self, chain) -> EffectiveValue:
        """Resolve `algn` over a paragraph chain; ECMA default is "l" (left)."""
        steps = []
        for label, partname, pPr in chain:
            algn = pPr.get("algn") if pPr is not None else None
            if algn is not None:
                steps.append(ProvenanceStep(label, partname, 'algn="%s"' % algn, True))
                return EffectiveValue(algn, None, True, tuple(steps))
            steps.append(
                ProvenanceStep(label, partname, self._absence(pPr, "no alignment"), False)
            )
        steps.append(
            ProvenanceStep("schema default", None, 'alignment defaults to "l"', True)
        )
        return EffectiveValue("l", None, True, tuple(steps))

    def resolve_line_spacing(self, chain) -> EffectiveValue:
        """Resolve `a:lnSpc` over a paragraph chain: float lines, or `Length` for points.

        The rendering default is single spacing (100%), so exhaustion resolves to 1.0.
        """
        steps = []
        for label, partname, pPr in chain:
            lnSpc = pPr.find(qn("a:lnSpc")) if pPr is not None else None
            if lnSpc is None:
                steps.append(
                    ProvenanceStep(label, partname, self._absence(pPr, "no lnSpc"), False)
                )
                continue
            spcPct = lnSpc.find(qn("a:spcPct"))
            if spcPct is not None:
                raw = spcPct.get("val")
                lines = (
                    float(raw[:-1]) / 100.0 if raw.endswith("%") else int(raw) / 100000.0
                )
                steps.append(
                    ProvenanceStep(label, partname, 'lnSpc spcPct val="%s"' % raw, True)
                )
                return EffectiveValue(lines, None, True, tuple(steps))
            spcPts = lnSpc.find(qn("a:spcPts"))
            if spcPts is not None:
                length = Length(Centipoints(int(spcPts.get("val"))))
                steps.append(
                    ProvenanceStep(
                        label, partname, 'lnSpc spcPts val="%s"' % spcPts.get("val"), True
                    )
                )
                return EffectiveValue(length, length.pt, True, tuple(steps))
            steps.append(
                ProvenanceStep(label, partname, "lnSpc with no spcPct/spcPts", False)
            )
            return EffectiveValue(None, None, False, tuple(steps))
        steps.append(
            ProvenanceStep("rendering default", None, "line spacing defaults to 100%", True)
        )
        return EffectiveValue(1.0, None, True, tuple(steps))

    def resolve_bullet(self, chain) -> EffectiveBullet:
        """Resolve the `EG_TextBullet` choice over a paragraph chain.

        The first rung supplying any member wins outright — `a:buNone` is a positive answer
        ("explicitly no bullet"), not an absence. There is no ECMA default bullet, so
        exhaustion resolves to "none" through a rendering-default step. Attributes are read
        raw: the generated descriptors are `RequiredAttribute` and would raise on exactly the
        malformed input this reports honestly.
        """
        steps = []

        def malformed(detail):
            """Report a present-but-unusable member and stop, per `resolve_line_spacing`."""
            steps.append(ProvenanceStep(label, partname, detail, False))
            return EffectiveBullet(None, None, None, None, False, tuple(steps))

        def supplied(detail, type_, char=None, scheme=None, start_at=None):
            steps.append(ProvenanceStep(label, partname, detail, True))
            return EffectiveBullet(type_, char, scheme, start_at, True, tuple(steps))

        for label, partname, pPr in chain:
            bullet = pPr.eg_bullet if pPr is not None else None
            if bullet is None and pPr is not None:
                bullet = pPr.find(qn("a:buBlip"))
            if bullet is None:
                steps.append(
                    ProvenanceStep(label, partname, self._absence(pPr, "no bullet"), False)
                )
                continue
            if bullet.tag == qn("a:buNone"):
                return supplied("buNone", "none")
            if bullet.tag == qn("a:buChar"):
                char = bullet.get("char")
                if char is None:
                    return malformed("buChar with no char")
                return supplied('buChar char="%s"' % char, "character", char=char)
            if bullet.tag == qn("a:buAutoNum"):
                scheme = bullet.get("type")
                if scheme is None:
                    return malformed("buAutoNum with no type")
                raw_start = bullet.get("startAt")
                try:
                    start_at = 1 if raw_start is None else int(raw_start)
                except ValueError:
                    return malformed('buAutoNum with non-integer startAt="%s"' % raw_start)
                return supplied(
                    'buAutoNum type="%s"' % scheme, "numbered", scheme=scheme, start_at=start_at
                )
            return supplied("buBlip", "picture")
        steps.append(
            ProvenanceStep("rendering default", None, "no bullet renders by default", True)
        )
        return EffectiveBullet("none", None, None, None, True, tuple(steps))

    def resolve_bullet_font(self, chain) -> EffectiveValue:
        """Resolve the `EG_TextBulletTypeface` choice over a paragraph chain.

        Inherits independently of the bullet kind, so it gets its own walk: one rung can supply
        the typeface while another supplied the bullet. `a:buFontTx` is a positive answer
        meaning "draw the bullet in the text's own typeface", not a miss. No ECMA default
        exists, and with neither member present PowerPoint follows the text, so exhaustion
        resolves to `BULLET_FOLLOWS_TEXT` rather than reporting unresolved.
        """
        steps = []
        for label, partname, pPr in chain:
            if pPr is not None and pPr.find(qn("a:buFontTx")) is not None:
                steps.append(ProvenanceStep(label, partname, "buFontTx", True))
                return EffectiveValue(BULLET_FOLLOWS_TEXT, None, True, tuple(steps))
            buFont = pPr.find(qn("a:buFont")) if pPr is not None else None
            if buFont is None:
                steps.append(
                    ProvenanceStep(label, partname, self._absence(pPr, "no bullet font"), False)
                )
                continue
            typeface = buFont.get("typeface")  # -- raw: the descriptor raises when absent
            if not typeface:
                steps.append(ProvenanceStep(label, partname, "buFont with no typeface", False))
                return EffectiveValue(None, None, False, tuple(steps))
            if typeface.startswith("+"):
                steps.append(
                    ProvenanceStep(
                        label,
                        partname,
                        'buFont typeface="%s" (theme reference)' % typeface,
                        False,
                    )
                )
                return self._resolve_theme_font(typeface, steps)
            steps.append(
                ProvenanceStep(label, partname, 'buFont typeface="%s"' % typeface, True)
            )
            return EffectiveValue(typeface, None, True, tuple(steps))
        steps.append(
            ProvenanceStep("rendering default", None, "bullet follows the text typeface", True)
        )
        return EffectiveValue(BULLET_FOLLOWS_TEXT, None, True, tuple(steps))

    def resolve_bullet_size(self, chain) -> EffectiveValue:
        """Resolve the `EG_TextBulletSize` choice over a paragraph chain.

        Three members, tested in schema order: `a:buSzTx` (follow the text), `a:buSzPct` (a
        fraction of the text size), `a:buSzPts` (absolute, reported as a `Length` with
        `value_pt` populated, as `resolve_line_spacing` does for `spcPts`). Inherits
        independently of the bullet kind and exhausts to `BULLET_FOLLOWS_TEXT`.
        """
        steps = []
        for label, partname, pPr in chain:
            if pPr is None:
                steps.append(
                    ProvenanceStep(label, partname, self._absence(pPr, "no bullet size"), False)
                )
                continue
            if pPr.find(qn("a:buSzTx")) is not None:
                steps.append(ProvenanceStep(label, partname, "buSzTx", True))
                return EffectiveValue(BULLET_FOLLOWS_TEXT, None, True, tuple(steps))
            buSzPct = pPr.find(qn("a:buSzPct"))
            if buSzPct is not None:
                raw = buSzPct.get("val")
                try:
                    fraction = ST_TextBulletSizePercent.convert_from_xml(raw)
                except (TypeError, ValueError):
                    steps.append(
                        ProvenanceStep(
                            label, partname, 'buSzPct with unreadable val="%s"' % raw, False
                        )
                    )
                    return EffectiveValue(None, None, False, tuple(steps))
                steps.append(ProvenanceStep(label, partname, 'buSzPct val="%s"' % raw, True))
                return EffectiveValue(fraction, None, True, tuple(steps))
            buSzPts = pPr.find(qn("a:buSzPts"))
            if buSzPts is not None:
                raw = buSzPts.get("val")
                try:
                    length = Length(Centipoints(int(raw)))
                except (TypeError, ValueError):
                    steps.append(
                        ProvenanceStep(
                            label, partname, 'buSzPts with unreadable val="%s"' % raw, False
                        )
                    )
                    return EffectiveValue(None, None, False, tuple(steps))
                steps.append(ProvenanceStep(label, partname, 'buSzPts val="%s"' % raw, True))
                return EffectiveValue(length, length.pt, True, tuple(steps))
            steps.append(
                ProvenanceStep(label, partname, self._absence(pPr, "no bullet size"), False)
            )
        steps.append(
            ProvenanceStep("rendering default", None, "bullet follows the text size", True)
        )
        return EffectiveValue(BULLET_FOLLOWS_TEXT, None, True, tuple(steps))

    def resolve_container_fill(
        self, container, label, partname, style_ref, ref_label
    ) -> EffectiveValue:
        """Resolve the solid fill of `container` (spPr or a:ln), else report the style ref.

        Explicit solid colors resolve fully (srgb direct, scheme through clrMap/theme);
        `a:noFill` resolves to "none". A style fill/line reference reports unresolved with
        the reference index and its resolved phClr color in provenance — theme format-scheme
        modulation is not applied.
        """
        steps = []
        if container is not None:
            solidFill = container.find(qn("a:solidFill"))
            if solidFill is not None:
                return self._resolve_solid_fill(solidFill, label, partname, steps)
            if container.find(qn("a:noFill")) is not None:
                steps.append(ProvenanceStep(label, partname, "a:noFill", True))
                return EffectiveValue("none", None, True, tuple(steps))
            other = [
                el.tag.rsplit("}", 1)[-1]
                for el in container
                if el.tag.rsplit("}", 1)[-1].endswith("Fill")
                or el.tag.rsplit("}", 1)[-1] in ("gradFill", "blipFill", "pattFill", "grpFill")
            ]
            if other:
                steps.append(
                    ProvenanceStep(
                        label, partname, "%s (not a single solid color)" % other[0], False
                    )
                )
                return EffectiveValue(None, None, False, tuple(steps))
            steps.append(ProvenanceStep(label, partname, "no explicit fill", False))
        else:
            steps.append(ProvenanceStep(label, partname, "element not present", False))

        if style_ref is None:
            steps.append(ProvenanceStep(ref_label, partname, "no style reference", False))
            return EffectiveValue(None, None, False, tuple(steps))
        idx = style_ref.get("idx")
        ph_color = self._style_ref_color(style_ref, ref_label, partname)
        steps.extend(ph_color[1])
        steps.append(
            ProvenanceStep(
                ref_label,
                partname,
                'idx="%s": theme format-scheme modulation of the reference color%s is not'
                " resolved in v0.1"
                % (idx, " (%s)" % ph_color[0] if ph_color[0] else ""),
                False,
            )
        )
        return EffectiveValue(None, None, False, tuple(steps))

    def _style_ref_color(self, style_ref, ref_label, partname):
        """Return (rgb-or-None, steps) for the phClr color child of a style reference."""
        srgbClr = style_ref.find(qn("a:srgbClr"))
        if srgbClr is not None:
            return srgbClr.get("val").upper(), [
                ProvenanceStep(
                    ref_label, partname, 'reference color srgbClr val="%s"' % srgbClr.get("val"),
                    False,
                )
            ]
        schemeClr = style_ref.find(qn("a:schemeClr"))
        if schemeClr is not None:
            resolved = self._resolve_scheme_color(schemeClr.get("val"), [], partname)
            rgb = resolved.value if resolved.resolved else None
            return rgb, [
                ProvenanceStep(
                    ref_label,
                    partname,
                    'reference color schemeClr val="%s" -> %s'
                    % (schemeClr.get("val"), rgb if rgb else "unresolved"),
                    False,
                )
            ]
        return None, []

    def _resolve_solid_fill(self, solidFill, label, partname, steps) -> EffectiveValue:
        srgbClr = solidFill.find(qn("a:srgbClr"))
        if srgbClr is not None:
            detail = 'srgbClr val="%s"%s' % (
                srgbClr.get("val"),
                self._transforms_note(srgbClr),
            )
            if len(srgbClr):
                steps.append(ProvenanceStep(label, partname, detail, False))
                return self._transformed_color_value(srgbClr, srgbClr.get("val"), steps)
            steps.append(ProvenanceStep(label, partname, detail, True))
            return EffectiveValue(srgbClr.get("val").upper(), None, True, tuple(steps))
        schemeClr = solidFill.find(qn("a:schemeClr"))
        if schemeClr is not None:
            steps.append(
                ProvenanceStep(
                    label,
                    partname,
                    'schemeClr val="%s"%s'
                    % (schemeClr.get("val"), self._transforms_note(schemeClr)),
                    False,
                )
            )
            if len(schemeClr):
                base = self._resolve_scheme_color(schemeClr.get("val"), steps, partname)
                if not base.resolved or base.value is None:
                    return base
                return self._transformed_color_value(
                    schemeClr, str(base.value), list(base.provenance)
                )
            return self._resolve_scheme_color(schemeClr.get("val"), steps, partname)
        steps.append(
            ProvenanceStep(
                label,
                partname,
                "solidFill with unsupported color element (%s)"
                % ", ".join(el.tag.rsplit("}", 1)[-1] for el in solidFill),
                False,
            )
        )
        return EffectiveValue(None, None, False, tuple(steps))

    @staticmethod
    def _pPr_from_lstStyle(lstStyle, level):
        if lstStyle is None:
            return None
        return lstStyle.pPr_for_lvl(level)

    @staticmethod
    def _ph_lstStyle_pPr(placeholder_proxy, level):
        if placeholder_proxy is None:
            return None
        txBody = placeholder_proxy._element.find(qn("p:txBody"))
        if txBody is None:
            return None
        return _FontResolver._pPr_from_lstStyle(txBody.find(qn("a:lstStyle")), level)

    def _master_family_style(self, ph_type):
        txStyles = self._master.element.txStyles
        if ph_type in _TITLE_FAMILY:
            return "titleStyle", txStyles.titleStyle if txStyles is not None else None
        if ph_type in _BODY_FAMILY:
            return "bodyStyle", txStyles.bodyStyle if txStyles is not None else None
        return "otherStyle", txStyles.otherStyle if txStyles is not None else None

    def _layout_placeholder(self, idx, ph_type):
        """Return only the exact-idx layout placeholder used by PowerPoint inheritance."""
        if idx == 4294967295:
            return None
        matches = [
            placeholder
            for placeholder in self._layout.placeholders
            if placeholder.element.ph_idx == idx
        ]
        if len(matches) > 1:
            def inheritance_type(placeholder_type):
                if placeholder_type in _TITLE_FAMILY:
                    return "title"
                if placeholder_type in _BODY_FAMILY:
                    return "body"
                return placeholder_type

            typed_matches = [
                placeholder
                for placeholder in matches
                if inheritance_type(placeholder.element.ph_type)
                == inheritance_type(ph_type)
            ]
            if len(typed_matches) == 1:
                matches = typed_matches
        if len(matches) > 1:
            raise UnsupportedStructureError(
                "layout %s has %d placeholders with idx %d; effective formatting "
                "inheritance is ambiguous"
                % (self._layout.part.partname, len(matches), idx)
            )
        if matches:
            return matches[0]
        raise UnsupportedStructureError(
            "slide placeholder idx %d (type %s) has no layout placeholder with the same "
            "idx; PowerPoint placeholder inheritance requires an exact idx match"
            % (idx, ph_type.name)
        )

    def _master_placeholder(self, ph_type):
        matches = [
            placeholder
            for placeholder in self._master.placeholders
            if placeholder.element.ph_type == ph_type
        ]
        if len(matches) > 1:
            raise UnsupportedStructureError(
                "slide master %s has %d placeholders of type %s; effective formatting "
                "inheritance is ambiguous"
                % (self._master.part.partname, len(matches), ph_type.name)
            )
        return matches[0] if matches else None

    @staticmethod
    def _master_placeholder_type(ph_type):
        if ph_type in _TITLE_FAMILY:
            return PP_PLACEHOLDER.TITLE
        if ph_type in (
            PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.HEADER,
            PP_PLACEHOLDER.SLIDE_NUMBER,
        ):
            return ph_type
        return PP_PLACEHOLDER.BODY

    @staticmethod
    def _ph_lstStyle_defRPr(placeholder_proxy, level):
        if placeholder_proxy is None:
            return None
        txBody = placeholder_proxy._element.find(qn("p:txBody"))
        if txBody is None:
            return None
        return _FontResolver._defRPr_from_lstStyle(txBody.find(qn("a:lstStyle")), level)

    @staticmethod
    def _defRPr_from_lstStyle(lstStyle, level):
        if lstStyle is None:
            return None
        lvl_pPr = lstStyle.pPr_for_lvl(level)
        if lvl_pPr is None:
            return None
        return lvl_pPr.find(qn("a:defRPr"))

    # ------------------------------------------------------------------------ resolvers

    def _resolve_size(self, chain) -> EffectiveValue:
        steps = []
        for label, partname, rPr in chain:
            sz = rPr.get("sz") if rPr is not None else None
            if sz is not None:
                steps.append(ProvenanceStep(label, partname, 'sz="%s"' % sz, True))
                size = Length(Centipoints(int(sz)))
                return EffectiveValue(size, size.pt, True, tuple(steps))
            steps.append(ProvenanceStep(label, partname, self._absence(rPr, "no size"), False))
        return EffectiveValue(None, None, False, tuple(steps))

    def _resolve_flag(self, chain, attr: str, what: str) -> EffectiveValue:
        """Resolve boolean rPr attribute `attr` (b/i); its schema default (false) is explicit,
        so exhaustion resolves to False rather than reporting unresolved."""
        steps = []
        for label, partname, rPr in chain:
            raw = rPr.get(attr) if rPr is not None else None
            if raw is not None:
                value = raw in ("1", "true")
                steps.append(ProvenanceStep(label, partname, '%s="%s"' % (attr, raw), True))
                return EffectiveValue(value, None, True, tuple(steps))
            steps.append(
                ProvenanceStep(label, partname, self._absence(rPr, "no %s" % what), False)
            )
        steps.append(
            ProvenanceStep("schema default", None, "%s defaults to false" % what, True)
        )
        return EffectiveValue(False, None, True, tuple(steps))

    def _resolve_underline(self, chain) -> EffectiveValue:
        """Resolve the `u` token ("sng", "dbl", ...); schema default is "none"."""
        steps = []
        for label, partname, rPr in chain:
            raw = rPr.get("u") if rPr is not None else None
            if raw is not None:
                steps.append(ProvenanceStep(label, partname, 'u="%s"' % raw, True))
                return EffectiveValue(raw, None, True, tuple(steps))
            steps.append(
                ProvenanceStep(label, partname, self._absence(rPr, "no underline"), False)
            )
        steps.append(
            ProvenanceStep("schema default", None, 'underline defaults to "none"', True)
        )
        return EffectiveValue("none", None, True, tuple(steps))

    def _resolve_name(self, chain, text="") -> EffectiveValue:
        if self._has_non_latin_letters(text):
            return EffectiveValue(
                None,
                None,
                False,
                (
                    ProvenanceStep(
                        "script-specific font",
                        str(self._slide_part.partname),
                        "run contains non-Latin text; East Asian and complex-script font "
                        "selection is not resolved",
                        False,
                    ),
                ),
            )
        steps = []
        for label, partname, rPr in chain:
            latin = rPr.find(qn("a:latin")) if rPr is not None else None
            typeface = latin.get("typeface") if latin is not None else None
            if typeface is None:
                steps.append(
                    ProvenanceStep(label, partname, self._absence(rPr, "no latin typeface"), False)
                )
                continue
            if typeface.startswith("+"):
                steps.append(
                    ProvenanceStep(
                        label, partname, 'latin typeface="%s" (theme reference)' % typeface, False
                    )
                )
                return self._resolve_theme_font(typeface, steps)
            steps.append(
                ProvenanceStep(label, partname, 'latin typeface="%s"' % typeface, True)
            )
            return EffectiveValue(typeface, None, True, tuple(steps))
        return EffectiveValue(None, None, False, tuple(steps))

    def _resolve_theme_font(self, token, steps) -> EffectiveValue:
        scheme_element = {"mj": "a:majorFont", "mn": "a:minorFont"}.get(token[1:3])
        if self._theme_root is None or scheme_element is None or not token.endswith("-lt"):
            steps.append(
                ProvenanceStep(
                    "theme fontScheme",
                    self._theme_partname,
                    "unresolvable theme font reference %r" % token,
                    False,
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        latin = self._theme_root.find(
            qn("a:themeElements") + "/" + qn("a:fontScheme") + "/" + qn(scheme_element)
            + "/" + qn("a:latin")
        )
        typeface = latin.get("typeface") if latin is not None else None
        if not typeface:
            steps.append(
                ProvenanceStep(
                    "theme fontScheme",
                    self._theme_partname,
                    "%s has no latin typeface" % scheme_element,
                    False,
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        steps.append(
            ProvenanceStep(
                "theme fontScheme %s" % scheme_element[2:],
                self._theme_partname,
                'latin typeface="%s"' % typeface,
                True,
            )
        )
        return EffectiveValue(typeface, None, True, tuple(steps))

    def _resolve_color(self, chain, sp) -> EffectiveValue:
        steps = []
        for label, partname, rPr in chain:
            solidFill = rPr.find(qn("a:solidFill")) if rPr is not None else None
            if solidFill is None:
                fill_kind = self._non_solid_fill_kind(rPr)
                if fill_kind is not None:
                    steps.append(
                        ProvenanceStep(
                            label, partname, "%s (not a single solid color)" % fill_kind, False
                        )
                    )
                    return EffectiveValue(None, None, False, tuple(steps))
                steps.append(
                    ProvenanceStep(label, partname, self._absence(rPr, "no solidFill"), False)
                )
                continue
            srgbClr = solidFill.find(qn("a:srgbClr"))
            if srgbClr is not None:
                detail = 'srgbClr val="%s"%s' % (
                    srgbClr.get("val"),
                    self._transforms_note(srgbClr),
                )
                if len(srgbClr):
                    steps.append(ProvenanceStep(label, partname, detail, False))
                    return self._transformed_color_value(
                        srgbClr, srgbClr.get("val"), steps
                    )
                steps.append(ProvenanceStep(label, partname, detail, True))
                return EffectiveValue(srgbClr.get("val").upper(), None, True, tuple(steps))
            schemeClr = solidFill.find(qn("a:schemeClr"))
            if schemeClr is not None:
                steps.append(
                    ProvenanceStep(
                        label,
                        partname,
                        'schemeClr val="%s"%s'
                        % (schemeClr.get("val"), self._transforms_note(schemeClr)),
                        False,
                    )
                )
                if len(schemeClr):
                    base = self._resolve_scheme_color(schemeClr.get("val"), steps, partname)
                    if not base.resolved or base.value is None:
                        return base
                    return self._transformed_color_value(
                        schemeClr, str(base.value), list(base.provenance)
                    )
                return self._resolve_scheme_color(schemeClr.get("val"), steps, partname)
            steps.append(
                ProvenanceStep(
                    label,
                    partname,
                    "solidFill with unsupported color element (%s)"
                    % ", ".join(
                        el.tag.rsplit("}", 1)[-1] for el in solidFill
                    ),
                    False,
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        return EffectiveValue(None, None, False, tuple(steps))

    def _resolve_scheme_color(self, token, steps, source_partname) -> EffectiveValue:
        slot, map_step = self._map_scheme_token(token, source_partname)
        steps.append(map_step)
        if slot is None:
            return EffectiveValue(None, None, False, tuple(steps))
        if self._theme_root is None:
            steps.append(
                ProvenanceStep(
                    "theme clrScheme", self._theme_partname, "theme part unavailable", False
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        slot_element = self._theme_root.find(
            qn("a:themeElements") + "/" + qn("a:clrScheme") + "/" + qn("a:%s" % slot)
        )
        color_child = slot_element[0] if slot_element is not None and len(slot_element) else None
        if color_child is None:
            steps.append(
                ProvenanceStep(
                    "theme clrScheme %s" % slot,
                    self._theme_partname,
                    "slot missing from clrScheme",
                    False,
                )
            )
            return EffectiveValue(None, None, False, tuple(steps))
        local = color_child.tag.rsplit("}", 1)[-1]
        if local == "srgbClr":
            rgb = color_child.get("val").upper()
            detail = 'srgbClr val="%s"' % color_child.get("val")
        elif local == "sysClr":
            rgb = (color_child.get("lastClr") or "").upper() or None
            detail = 'sysClr val="%s" lastClr="%s"' % (
                color_child.get("val"),
                color_child.get("lastClr"),
            )
        else:
            rgb, detail = None, "unsupported clrScheme child %s" % local
        steps.append(
            ProvenanceStep(
                "theme clrScheme %s" % slot, self._theme_partname, detail, rgb is not None
            )
        )
        return EffectiveValue(rgb, None, rgb is not None, tuple(steps))

    def _map_scheme_token(self, token, source_partname):
        """Return (theme-slot, ProvenanceStep) for schemeClr `token`."""
        override = self._clrMapOvr_mapping(source_partname)
        if override is not None:
            mapping, source_label, source_part = override
        else:
            clrMap = self._master.element.find(qn("p:clrMap"))
            mapping = dict(clrMap.attrib) if clrMap is not None else {}
            source_label, source_part = "master clrMap", str(self._master.part.partname)
        if token in mapping:
            slot = mapping[token]
            return slot, ProvenanceStep(
                source_label, source_part, '%s="%s"' % (token, slot), False
            )
        if token in _DIRECT_THEME_SLOTS:
            return token, ProvenanceStep(
                source_label,
                source_part,
                '"%s" addresses the theme slot directly (no mapping entry)' % token,
                False,
            )
        return None, ProvenanceStep(
            source_label, source_part, "unmappable scheme color token %r" % token, False
        )

    def _clrMapOvr_mapping(self, source_partname):
        """Return the color-map override belonging to the property source part."""
        layout_partname = str(self._layout.part.partname)
        master_partname = str(self._master.part.partname)
        if source_partname == master_partname:
            return None
        if source_partname == layout_partname:
            root = self._layout._element
            label = "layout clrMapOvr"
            partname = layout_partname
        else:
            root = self._slide_part._element
            label = "slide clrMapOvr"
            partname = str(self._slide_part.partname)
        clrMapOvr = root.find(qn("p:clrMapOvr"))
        if clrMapOvr is None:
            return None
        overrideClrMapping = clrMapOvr.find(qn("a:overrideClrMapping"))
        if overrideClrMapping is None:
            return None  # -- a:masterClrMapping (or empty): defer to the master's clrMap
        return (
            dict(overrideClrMapping.attrib),
            label,
            partname,
        )

    # ------------------------------------------------------------------------- helpers

    @staticmethod
    def _transforms_note(color_element) -> str:
        """Return a note naming any color-transform children (lumMod, tint, …), or ""."""
        transforms = [
            "%s=%s" % (child.tag.rsplit("}", 1)[-1], child.get("val"))
            for child in color_element
        ]
        if not transforms:
            return ""
        return " with unapplied transforms [%s]" % ", ".join(transforms)

    @staticmethod
    def _transformed_color_value(color_element, base_rgb, steps) -> EffectiveValue:
        localized = OxmlElement("a:srgbClr")
        localized.set("val", str(base_rgb).upper())
        for transform in color_element:
            localized.append(copy.deepcopy(transform))
        from lxml import etree

        return EffectiveValue(
            None,
            None,
            False,
            tuple(steps),
            bake_color_xml=etree.tostring(localized),
        )

    @staticmethod
    def _has_non_latin_letters(text: str) -> bool:
        for char in text:
            if not unicodedata.category(char).startswith("L"):
                continue
            if not unicodedata.name(char, "").startswith("LATIN"):
                return True
        return False

    @staticmethod
    def _rPr_from_fontRef(sp):
        """Return an rPr-like projection of this shape's explicit fontRef."""
        fontRef = sp.find(qn("p:style") + "/" + qn("a:fontRef"))
        if fontRef is None or fontRef.get("idx") not in ("major", "minor"):
            return None
        rPr = OxmlElement("a:defRPr")
        latin = OxmlElement("a:latin")
        latin.set("typeface", "+%s-lt" % ("mj" if fontRef.get("idx") == "major" else "mn"))
        rPr.append(latin)
        if len(fontRef):
            solidFill = OxmlElement("a:solidFill")
            solidFill.append(copy.deepcopy(fontRef[0]))
            rPr.append(solidFill)
        return rPr

    @staticmethod
    def _non_solid_fill_kind(rPr):
        if rPr is None:
            return None
        for tag in ("a:gradFill", "a:blipFill", "a:pattFill", "a:noFill", "a:grpFill"):
            if rPr.find(qn(tag)) is not None:
                return tag
        return None

    @staticmethod
    def _absence(rPr, what) -> str:
        return "level not present" if rPr is None else what + " here"

    @property
    def _presentation_part(self):
        return self._slide_part.package.presentation_part

    def _load_theme(self):
        try:
            theme_part = self._master.part.part_related_by(RT.THEME)
        except KeyError:
            return None, None
        try:
            root = parse_xml(theme_part.blob)
        except Exception:
            raise UnsupportedStructureError(
                "theme part %s is not parseable XML" % theme_part.partname
            )
        return root, str(theme_part.partname)
