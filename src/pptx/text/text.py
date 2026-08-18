"""Text-related objects such as TextFrame and Paragraph."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, cast

from pptx.dml.fill import FillFormat
from pptx.enum.dml import MSO_FILL
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import MSO_AUTO_SIZE, MSO_UNDERLINE, MSO_VERTICAL_ANCHOR
from pptx.errors import UnsupportedStructureError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import ST_TextWrappingType
from pptx.shapes import Subshape
from pptx.text.bullet import BulletFormat
from pptx.text.fonts import FontFiles
from pptx.text.layout import TextFitter
from pptx.util import Centipoints, Emu, Length, Pt, lazyproperty

if TYPE_CHECKING:
    from pptx.dml.color import ColorFormat
    from pptx.enum.text import (
        MSO_TEXT_UNDERLINE_TYPE,
        MSO_VERTICAL_ANCHOR,
        PP_PARAGRAPH_ALIGNMENT,
    )
    from pptx.oxml.action import CT_Hyperlink
    from pptx.oxml.text import (
        CT_RegularTextRun,
        CT_TextBody,
        CT_TextCharacterProperties,
        CT_TextParagraph,
        CT_TextParagraphProperties,
    )
    from pptx.types import ProvidesExtents, ProvidesPart


class TextFrame(Subshape):
    """The part of a shape that contains its text.

    Not all shapes have a text frame. Corresponds to the `p:txBody` element that can
    appear as a child element of `p:sp`. Not intended to be constructed directly.
    """

    def __init__(self, txBody: CT_TextBody, parent: ProvidesPart):
        super(TextFrame, self).__init__(parent)
        self._element = self._txBody = txBody
        self._parent = parent

    def add_paragraph(self):
        """
        Return new `_Paragraph` instance appended to the sequence of
        paragraphs contained in this text frame.
        """
        p = self._txBody.add_p()
        return _Paragraph(p, self)

    @property
    def auto_size(self) -> MSO_AUTO_SIZE | None:
        """Resizing strategy used to fit text within this shape.

        Determins the type of automatic resizing used to fit the text of this shape within its
        bounding box when the text would otherwise extend beyond the shape boundaries. May be
        `None`, `MSO_AUTO_SIZE.NONE`, `MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT`, or
        `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`.
        """
        return self._bodyPr.autofit

    @auto_size.setter
    def auto_size(self, value: MSO_AUTO_SIZE | None):
        self._bodyPr.autofit = value

    def clear(self):
        """Remove all paragraphs except one empty one."""
        for p in self._txBody.p_lst[1:]:
            self._txBody.remove(p)
        p = self.paragraphs[0]
        p.clear()

    def fit_text(
        self,
        font_family: str = "Calibri",
        max_size: int = 18,
        bold: bool = False,
        italic: bool = False,
        font_file: str | None = None,
    ):
        """Fit text-frame text entirely within bounds of its shape.

        Make the text in this text frame fit entirely within the bounds of its shape by setting
        word wrap on and applying the "best-fit" font size to all the text it contains.

        `TextFrame.auto_size` is set to `MSO_AUTO_SIZE.NONE`. The font size will not
        be set larger than `max_size` points. If the path to a matching TrueType font is provided
        as `font_file`, that font file will be used for the font metrics. If `font_file` is `None`,
        best efforts are made to locate a font file with matchhing `font_family`, `bold`, and
        `italic` installed on the current system (usually succeeds if the font is installed).
        """
        # ---no-op when empty as fit behavior not defined for that case---
        if self.text == "":
            return  # pragma: no cover

        font_size = self._best_fit_font_size(font_family, max_size, bold, italic, font_file)
        self._apply_fit(font_family, font_size, bold, italic)

    @property
    def font_scale(self) -> float | None:
        """Font scale percent of this frame's `a:normAutofit`, e.g. 62.5 (paper-pptx addition).

        100.0 when the frame has `a:normAutofit` with no explicit scale; `None` when the
        frame's autofit setting is anything other than `a:normAutofit`. Read-only: PowerPoint
        owns this value (it records the shrink-to-fit reduction last applied); use
        `normalize_autofit` to freeze it into explicit sizes.
        """
        normAutofit = self._bodyPr.normAutofit
        return normAutofit.fontScale if normAutofit is not None else None

    @property
    def line_space_reduction(self) -> float | None:
        """Line-spacing reduction percent of `a:normAutofit`, e.g. 20.0 (paper-pptx addition).

        0.0 when the frame has `a:normAutofit` with no explicit reduction; `None` when the
        frame's autofit setting is anything other than `a:normAutofit`. Read-only.
        """
        normAutofit = self._bodyPr.normAutofit
        return normAutofit.lnSpcReduction if normAutofit is not None else None

    @property
    def margin_bottom(self) -> Length:
        """`Length` value representing the inset of text from the bottom text frame border.

        `pptx.util.Inches` provides a convenient way of setting the value, e.g.
        `text_frame.margin_bottom = Inches(0.05)`.
        """
        return self._bodyPr.bIns

    @margin_bottom.setter
    def margin_bottom(self, emu: Length):
        self._bodyPr.bIns = emu

    @property
    def margin_left(self) -> Length:
        """Inset of text from left text frame border as `Length` value."""
        return self._bodyPr.lIns

    @margin_left.setter
    def margin_left(self, emu: Length):
        self._bodyPr.lIns = emu

    @property
    def margin_right(self) -> Length:
        """Inset of text from right text frame border as `Length` value."""
        return self._bodyPr.rIns

    @margin_right.setter
    def margin_right(self, emu: Length):
        self._bodyPr.rIns = emu

    @property
    def margin_top(self) -> Length:
        """Inset of text from top text frame border as `Length` value."""
        return self._bodyPr.tIns

    @margin_top.setter
    def margin_top(self, emu: Length):
        self._bodyPr.tIns = emu

    def normalize_autofit(
        self, *, min_font_size: Length | None = None, resolve: bool = False
    ) -> None:
        """Freeze this frame's rendered text metrics and set explicit no-autofit.

        paper-pptx addition. What the reader currently sees is made explicit, then the frame's
        autofit is set to `a:noAutofit`:

        - `a:normAutofit` with a font scale: every explicit font size in the frame (run,
          paragraph-default, and end-paragraph properties) is multiplied by the scale. If any
          run's size is not locally resolvable (neither its own `sz` nor its paragraph's
          default), `UnsupportedStructureError` is raised — unless `resolve=True`, in which
          case the size is resolved through the effective-style walk (placeholder → layout →
          master → theme) and frozen from the resolved value; what the walk cannot resolve
          still refuses. This API never silently guesses.
        - `a:normAutofit` with a line-spacing reduction: every paragraph's explicit line
          spacing is reduced accordingly; any paragraph without explicit line spacing raises
          `UnsupportedStructureError` (`resolve` covers font sizes only in this version).
        - `a:spAutoFit`, `a:noAutofit`, or no autofit element: no text metrics change.

        `min_font_size` (a `Length`, e.g. `Pt(11)`) is applied after freezing: explicit sizes
        below the floor are raised to it. Validation completes fully before the first write
        (a refusal leaves the frame byte-identical).
        """
        from pptx._ownership import require_element_attached

        require_element_attached(self._txBody, self.part, argument="text frame")
        scale = self.font_scale
        reduction = self.line_space_reduction
        scale = 100.0 if scale is None else scale
        reduction = 0.0 if reduction is None else reduction
        if not isinstance(resolve, bool):
            raise ValueError("resolve must be a bool, got %r" % (resolve,))
        if min_font_size is not None and (
            isinstance(min_font_size, bool)
            or not isinstance(min_font_size, int)
            or min_font_size <= 0
        ):
            raise ValueError(
                "min_font_size must be a positive Length (EMU int) or None, got %r"
                % (min_font_size,)
            )
        resolved_run_sizes = {}  # -- run element -> resolved centipoints (resolve=True plan)

        # -- validation pass: complete before any mutation (refusal atomicity). Raw-XML reads
        # -- only: the font/paragraph proxy accessors are get-or-add and would themselves
        # -- dirty the tree with empty rPr/pPr elements.
        shape_name = getattr(self._parent, "name", "<unknown shape>")
        for para_idx, p in enumerate(self._txBody.p_lst):
            pPr = p.find(qn("a:pPr"))
            if scale != 100.0:
                defRPr = pPr.find(qn("a:defRPr")) if pPr is not None else None
                para_has_size = defRPr is not None and defRPr.get("sz") is not None
                # -- a:fld (fields) render text exactly like runs and must be sizable too
                for r in p.findall(qn("a:r")) + p.findall(qn("a:fld")):
                    rPr = r.find(qn("a:rPr"))
                    run_has_size = rPr is not None and rPr.get("sz") is not None
                    if not run_has_size and not para_has_size:
                        if resolve:
                            resolved = self._resolve_run_size(r)
                            if resolved is not None:
                                resolved_run_sizes[r] = resolved
                                continue
                        raise UnsupportedStructureError(
                            "cannot freeze font scale %.1f%%: %s in paragraph %d of shape %r"
                            " has no locally resolvable font size (set explicit sizes first,"
                            " or pass resolve=True to resolve effective values before"
                            " normalizing)"
                            % (
                                scale,
                                "field" if r.tag == qn("a:fld") else "run",
                                para_idx,
                                shape_name,
                            )
                        )
            if reduction != 0.0:
                lnSpc = pPr.find(qn("a:lnSpc")) if pPr is not None else None
                if lnSpc is None:
                    raise UnsupportedStructureError(
                        "cannot freeze line-spacing reduction %.1f%%: paragraph %d of shape"
                        " %r has no explicit line spacing (set paragraph.line_spacing on"
                        " every paragraph first; resolve=True resolves font sizes only —"
                        " spacing resolution is not supported in this version)"
                        % (reduction, para_idx, shape_name)
                    )

        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            if scale != 100.0:
                for rPr in self._iter_explicitly_sized_rPrs():
                    rPr.sz = max(100, int(round(rPr.sz * scale / 100.0)))
                for r, resolved_centipoints in resolved_run_sizes.items():
                    r.get_or_add_rPr().sz = max(
                        100, int(round(resolved_centipoints * scale / 100.0))
                    )
            if reduction != 0.0:
                for paragraph in self.paragraphs:
                    spacing = paragraph.line_spacing
                    factor = (100.0 - reduction) / 100.0
                    if isinstance(spacing, Length):
                        paragraph.line_spacing = Centipoints(
                            int(round(spacing.centipoints * factor))
                        )
                    else:
                        paragraph.line_spacing = spacing * factor
            if min_font_size is not None:
                floor_centipoints = Emu(min_font_size).centipoints
                for rPr in self._iter_explicitly_sized_rPrs():
                    if rPr.sz < floor_centipoints:
                        rPr.sz = floor_centipoints
            self.auto_size = MSO_AUTO_SIZE.NONE

    def _resolve_run_size(self, r):
        """Return `r`'s effective font size in centipoints via the inheritance walk, or None.

        Used by `normalize_autofit(resolve=True)`. Import is deferred: `pptx.inspect` is a
        leaf module and importing it at module scope would be a cycle.
        """
        from pptx.inspect import _FontResolver
        from pptx.parts.slide import SlidePart

        sp = self._txBody.getparent()
        if sp is None or sp.tag != qn("p:sp"):
            return None  # -- only p:sp text bodies resolve through the placeholder chain
        if not isinstance(self.part, SlidePart):
            return None  # -- layout/master frames have no slide chain; typed refusal follows
        effective = _FontResolver(self.part).effective_font(r, sp)
        if not effective.size.resolved:
            return None
        return Emu(effective.size.value).centipoints

    def _iter_explicitly_sized_rPrs(self):
        """Generate every rPr-family element in this frame carrying an explicit `sz`.

        Covers run, field (`a:fld`), and line-break (`a:br`) properties, paragraph default
        run properties, and end-paragraph run properties. `sz` values are centipoints ints.
        """
        for p in self._txBody.p_lst:
            candidates = [p.find(qn("a:pPr")), p.find(qn("a:endParaRPr"))]
            for content_tag in ("a:r", "a:fld", "a:br"):
                candidates.extend(
                    child.find(qn("a:rPr")) for child in p.findall(qn(content_tag))
                )
            for element in candidates:
                if element is None:
                    continue
                if element.tag == qn("a:pPr"):
                    element = element.find(qn("a:defRPr"))
                if element is not None and element.get("sz") is not None:
                    yield element

    @property
    def paragraphs(self) -> tuple[_Paragraph, ...]:
        """Sequence of paragraphs in this text frame.

        A text frame always contains at least one paragraph.
        """
        return tuple([_Paragraph(p, self) for p in self._txBody.p_lst])

    @property
    def text(self) -> str:
        """All text in this text-frame as a single string.

        Read/write. The return value contains all text in this text-frame. A line-feed character
        (`"\\n"`) separates the text for each paragraph. A vertical-tab character (`"\\v"`) appears
        for each line break (aka. soft carriage-return) encountered.

        The vertical-tab character is how PowerPoint represents a soft carriage return in clipboard
        text, which is why that encoding was chosen.

        Assignment replaces all text in the text frame. A new paragraph is added for each line-feed
        character (`"\\n"`) encountered. A line-break (soft carriage-return) is inserted for each
        vertical-tab character (`"\\v"`) encountered.

        Any control character other than newline, tab, or vertical-tab are escaped as plain-text
        like "_x001B_" (for ESC (ASCII 32) in this example).
        """
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @text.setter
    def text(self, text: str):
        txBody = self._txBody
        txBody.clear_content()
        for p_text in text.split("\n"):
            p = txBody.add_p()
            p.append_text(p_text)

    @property
    def vertical_anchor(self) -> MSO_VERTICAL_ANCHOR | None:
        """Represents the vertical alignment of text in this text frame.

        `None` indicates the effective value should be inherited from this object's style hierarchy.
        """
        return self._txBody.bodyPr.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, value: MSO_VERTICAL_ANCHOR | None):
        bodyPr = self._txBody.bodyPr
        bodyPr.anchor = value

    @property
    def word_wrap(self) -> bool | None:
        """`True` when lines of text in this shape are wrapped to fit within the shape's width.

        Read-write. Valid values are True, False, or None. True and False turn word wrap on and
        off, respectively. Assigning None to word wrap causes any word wrap setting to be removed
        from the text frame, causing it to inherit this setting from its style hierarchy.
        """
        return {
            ST_TextWrappingType.SQUARE: True,
            ST_TextWrappingType.NONE: False,
            None: None,
        }[self._txBody.bodyPr.wrap]

    @word_wrap.setter
    def word_wrap(self, value: bool | None):
        if value not in (True, False, None):
            raise ValueError(  # pragma: no cover
                "assigned value must be True, False, or None, got %s" % value
            )
        self._txBody.bodyPr.wrap = {
            True: ST_TextWrappingType.SQUARE,
            False: ST_TextWrappingType.NONE,
            None: None,
        }[value]

    def _apply_fit(self, font_family: str, font_size: int, is_bold: bool, is_italic: bool):
        """Arrange text in this text frame to fit inside its extents.

        This is accomplished by setting auto size off, wrap on, and setting the font of
        all its text to `font_family`, `font_size`, `is_bold`, and `is_italic`.
        """
        self.auto_size = MSO_AUTO_SIZE.NONE
        self.word_wrap = True
        self._set_font(font_family, font_size, is_bold, is_italic)

    def _best_fit_font_size(
        self, family: str, max_size: int, bold: bool, italic: bool, font_file: str | None
    ) -> int:
        """Return font-size in points that best fits text in this text-frame.

        The best-fit font size is the largest integer point size not greater than `max_size` that
        allows all the text in this text frame to fit inside its extents when rendered using the
        font described by `family`, `bold`, and `italic`. If `font_file` is specified, it is used
        to calculate the fit, whether or not it matches `family`, `bold`, and `italic`.
        """
        if font_file is None:
            font_file = FontFiles.find(family, bold, italic)
        return TextFitter.best_fit_font_size(self.text, self._extents, max_size, font_file)

    @property
    def _bodyPr(self):
        return self._txBody.bodyPr

    @property
    def _extents(self) -> tuple[Length, Length]:
        """(cx, cy) 2-tuple representing the effective rendering area of this text-frame.

        Margins are taken into account.
        """
        parent = cast("ProvidesExtents", self._parent)
        return (
            Length(parent.width - self.margin_left - self.margin_right),
            Length(parent.height - self.margin_top - self.margin_bottom),
        )

    def _set_font(self, family: str, size: int, bold: bool, italic: bool):
        """Set the font properties of all the text in this text frame."""

        def iter_rPrs(txBody: CT_TextBody) -> Iterator[CT_TextCharacterProperties]:
            for p in txBody.p_lst:
                for elm in p.content_children:
                    yield elm.get_or_add_rPr()
                # generate a:endParaRPr for each <a:p> element
                yield p.get_or_add_endParaRPr()

        def set_rPr_font(
            rPr: CT_TextCharacterProperties, name: str, size: int, bold: bool, italic: bool
        ):
            f = Font(rPr)
            f.name, f.size, f.bold, f.italic = family, Pt(size), bold, italic

        txBody = self._element
        for rPr in iter_rPrs(txBody):
            set_rPr_font(rPr, family, size, bold, italic)


class Font(object):
    """Character properties object, providing font size, font name, bold, italic, etc.

    Corresponds to `a:rPr` child element of a run. Also appears as `a:defRPr` and
    `a:endParaRPr` in paragraph and `a:defRPr` in list style elements.
    """

    def __init__(self, rPr: CT_TextCharacterProperties):
        super(Font, self).__init__()
        self._element = self._rPr = rPr

    @property
    def bold(self) -> bool | None:
        """Get or set boolean bold value of `Font`, e.g. `paragraph.font.bold = True`.

        If set to `None`, the bold setting is cleared and is inherited from an enclosing shape's
        setting, or a setting in a style or master. Returns None if no bold attribute is present,
        meaning the effective bold value is inherited from a master or the theme.
        """
        return self._rPr.b

    @bold.setter
    def bold(self, value: bool | None):
        self._rPr.b = value

    @lazyproperty
    def color(self) -> ColorFormat:
        """The `ColorFormat` instance that provides access to the color settings for this font."""
        if self.fill.type != MSO_FILL.SOLID:
            self.fill.solid()
        return self.fill.fore_color

    @lazyproperty
    def fill(self) -> FillFormat:
        """`FillFormat` instance for this font.

        Provides access to fill properties such as fill color.
        """
        return FillFormat.from_fill_parent(self._rPr)

    @property
    def italic(self) -> bool | None:
        """Get or set boolean italic value of `Font` instance.

        Has the same behaviors as bold with respect to None values.
        """
        return self._rPr.i

    @italic.setter
    def italic(self, value: bool | None):
        self._rPr.i = value

    @property
    def language_id(self) -> MSO_LANGUAGE_ID | None:
        """Get or set the language id of this `Font` instance.

        The language id is a member of the `MsoLanguageId` enumeration. Assigning `None`
        removes any language setting, the same behavior as assigning `MSO_LANGUAGE_ID.NONE`.
        """
        lang = self._rPr.lang
        if lang is None:
            return MSO_LANGUAGE_ID.NONE
        return self._rPr.lang

    @language_id.setter
    def language_id(self, value: MSO_LANGUAGE_ID | None):
        if value == MSO_LANGUAGE_ID.NONE:
            value = None
        self._rPr.lang = value

    @property
    def name(self) -> str | None:
        """Get or set the typeface name for this `Font` instance.

        Causes the text it controls to appear in the named font, if a matching font is found.
        Returns `None` if the typeface is currently inherited from the theme. Setting it to `None`
        removes any override of the theme typeface.
        """
        latin = self._rPr.latin
        if latin is None:
            return None
        return latin.typeface

    @name.setter
    def name(self, value: str | None):
        if value is None:
            self._rPr._remove_latin()  # pyright: ignore[reportPrivateUsage]
        else:
            latin = self._rPr.get_or_add_latin()
            latin.typeface = value

    @property
    def size(self) -> Length | None:
        """Indicates the font height in English Metric Units (EMU).

        Read/write. `None` indicates the font size should be inherited from its style hierarchy,
        such as a placeholder or document defaults (usually 18pt). `Length` is a subclass of `int`
        having properties for convenient conversion into points or other length units. Likewise,
        the `pptx.util.Pt` class allows convenient specification of point values::

            >>> font.size = Pt(24)
            >>> font.size
            304800
            >>> font.size.pt
            24.0
        """
        sz = self._rPr.sz
        if sz is None:
            return None
        return Centipoints(sz)

    @size.setter
    def size(self, emu: Length | None):
        if emu is None:
            self._rPr.sz = None
        else:
            sz = Emu(emu).centipoints
            self._rPr.sz = sz

    @property
    def underline(self) -> bool | MSO_TEXT_UNDERLINE_TYPE | None:
        """Indicaties the underline setting for this font.

        Value is `True`, `False`, `None`, or a member of the `MsoTextUnderlineType`
        enumeration. `None` is the default and indicates the underline setting should be inherited
        from the style hierarchy, such as from a placeholder. `True` indicates single underline.
        `False` indicates no underline. Other settings such as double and wavy underlining are
        indicated with members of the `MsoTextUnderlineType` enumeration.
        """
        u = self._rPr.u
        if u is MSO_UNDERLINE.NONE:
            return False
        if u is MSO_UNDERLINE.SINGLE_LINE:
            return True
        return u

    @underline.setter
    def underline(self, value: bool | MSO_TEXT_UNDERLINE_TYPE | None):
        if value is True:
            value = MSO_UNDERLINE.SINGLE_LINE
        elif value is False:
            value = MSO_UNDERLINE.NONE
        self._element.u = value


class _Hyperlink(Subshape):
    """Text run hyperlink object.

    Corresponds to `a:hlinkClick` child element of the run's properties element (`a:rPr`).
    """

    def __init__(self, rPr: CT_TextCharacterProperties, parent: ProvidesPart):
        super(_Hyperlink, self).__init__(parent)
        self._rPr = rPr

    @property
    def address(self) -> str | None:
        """The URL of the hyperlink.

        Read/write. URL can be on http, https, mailto, or file scheme; others may work.
        """
        if self._hlinkClick is None:
            return None
        return self.part.target_ref(self._hlinkClick.rId)

    @address.setter
    def address(self, url: str | None):
        # implements all three of add, change, and remove hyperlink
        if self._hlinkClick is not None:
            self._remove_hlinkClick()
        if url:
            self._add_hlinkClick(url)

    def _add_hlinkClick(self, url: str):
        rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)
        self._rPr.add_hlinkClick(rId)

    @property
    def _hlinkClick(self) -> CT_Hyperlink | None:
        return self._rPr.hlinkClick

    def _remove_hlinkClick(self):
        assert self._hlinkClick is not None
        self.part.drop_rel(self._hlinkClick.rId)
        self._rPr._remove_hlinkClick()  # pyright: ignore[reportPrivateUsage]


class _Paragraph(Subshape):
    """Paragraph object. Not intended to be constructed directly."""

    def __init__(self, p: CT_TextParagraph, parent: ProvidesPart):
        super(_Paragraph, self).__init__(parent)
        self._element = self._p = p

    def add_line_break(self):
        """Add line break at end of this paragraph."""
        self._p.add_br()

    def add_datetime_field(self, format_code: str = "datetime") -> None:
        """Append a real date/time field (`a:fld type="datetime…"`) to this paragraph.

        paper-pptx addition. `format_code` is "datetime" or "datetime1" through
        "datetime13" (the ECMA-376 date/time field formats); anything else raises
        `ValueError`. The field's cached text is left empty; PowerPoint renders the live
        value. Recognized by `pptx.inspect.inspect_text` via `TextBlock.fields`.
        """
        valid = {"datetime"} | {"datetime%d" % n for n in range(1, 14)}
        if format_code not in valid:
            raise ValueError(
                "format_code must be 'datetime' or 'datetime1'..'datetime13', got %r"
                % (format_code,)
            )
        self._add_field(format_code, "")

    def add_slide_number_field(self) -> None:
        """Append a real slide-number field (`a:fld type="slidenum"`) to this paragraph.

        paper-pptx addition. This is the honest version of typing a literal page
        number into a footer: PowerPoint renders the actual slide number and keeps it right
        across reordering. The cached text is "1". Recognized by
        `pptx.inspect.inspect_text` via `TextBlock.fields`.
        """
        self._add_field("slidenum", "1")

    def _add_field(self, field_type: str, cached_text: str) -> None:
        import uuid

        from pptx._ownership import require_element_attached
        from pptx._transaction import PackageTransaction

        require_element_attached(self._p, self.part, argument="paragraph")
        with PackageTransaction(self.part.package, self):
            self._p.add_fld("{%s}" % str(uuid.uuid4()).upper(), field_type, cached_text)

    def add_run(self) -> _Run:
        """Return a new run appended to the runs in this paragraph."""
        r = self._p.add_r()
        return _Run(r, self)

    @property
    def alignment(self) -> PP_PARAGRAPH_ALIGNMENT | None:
        """Horizontal alignment of this paragraph.

        The value `None` indicates the paragraph should 'inherit' its effective value from its
        style hierarchy. Assigning `None` removes any explicit setting, causing its inherited
        value to be used.
        """
        return self._pPr.algn

    @alignment.setter
    def alignment(self, value: PP_PARAGRAPH_ALIGNMENT | None):
        self._pPr.algn = value

    @lazyproperty
    def bullet(self) -> BulletFormat:
        """`BulletFormat` object providing bullet and numbering control for this paragraph.

        paper-pptx addition. Read properties report this paragraph's local `a:pPr` state only
        (`None` means inherited); setters write real `a:buChar`/`a:buAutoNum`/`a:buNone`
        bullets with hanging-indent geometry.
        """
        return BulletFormat(self._p, self.part)

    def clear(self):
        """Remove all content from this paragraph.

        Paragraph properties are preserved. Content includes runs, line breaks, and fields.
        """
        for elm in self._element.content_children:
            self._element.remove(elm)
        return self

    @property
    def font(self) -> Font:
        """`Font` object containing default character properties for the runs in this paragraph.

        These character properties override default properties inherited from parent objects such
        as the text frame the paragraph is contained in and they may be overridden by character
        properties set at the run level.
        """
        return Font(self._defRPr)

    @property
    def level(self) -> int:
        """Indentation level of this paragraph.

        Read-write. Integer in range 0..8 inclusive. 0 represents a top-level paragraph and is the
        default value. Indentation level is most commonly encountered in a bulleted list, as is
        found on a word bullet slide.
        """
        return self._pPr.lvl

    @level.setter
    def level(self, level: int):
        self._pPr.lvl = level

    @property
    def line_spacing(self) -> int | float | Length | None:
        """The space between baselines in successive lines of this paragraph.

        A value of `None` indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. A numeric value, e.g. `2` or `1.5`,
        indicates spacing is applied in multiples of line heights. A `Length` value such as
        `Pt(12)` indicates spacing is a fixed height. The `Pt` value class is a convenient way to
        apply line spacing in units of points.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.line_spacing

    @line_spacing.setter
    def line_spacing(self, value: int | float | Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.line_spacing = value

    @property
    def runs(self) -> tuple[_Run, ...]:
        """Sequence of runs in this paragraph."""
        return tuple(_Run(r, self) for r in self._element.r_lst)

    @property
    def space_after(self) -> Length | None:
        """The spacing to appear between this paragraph and the subsequent paragraph.

        A value of `None` indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. `Length` objects provide convenience
        properties, such as `.pt` and `.inches`, that allow easy conversion to various length
        units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_after

    @space_after.setter
    def space_after(self, value: Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.space_after = value

    @property
    def space_before(self) -> Length | None:
        """The spacing to appear between this paragraph and the prior paragraph.

        A value of `None` indicates no explicit value is assigned and its effective value is
        inherited from the paragraph's style hierarchy. `Length` objects provide convenience
        properties, such as `.pt` and `.cm`, that allow easy conversion to various length units.
        """
        pPr = self._p.pPr
        if pPr is None:
            return None
        return pPr.space_before

    @space_before.setter
    def space_before(self, value: Length | None):
        pPr = self._p.get_or_add_pPr()
        pPr.space_before = value

    @property
    def text(self) -> str:
        """Text of paragraph as a single string.

        Read/write. This value is formed by concatenating the text in each run and field making up
        the paragraph, adding a vertical-tab character (`"\\v"`) for each line-break element
        (`<a:br>`, soft carriage-return) encountered.

        While the encoding of line-breaks as a vertical tab might be surprising at first, doing so
        is consistent with PowerPoint's clipboard copy behavior and allows a line-break to be
        distinguished from a paragraph boundary within the str return value.

        Assignment causes all content in the paragraph to be replaced. Each vertical-tab character
        (`"\\v"`) in the assigned str is translated to a line-break, as is each line-feed
        character (`"\\n"`). Contrast behavior of line-feed character in `TextFrame.text` setter.
        If line-feed characters are intended to produce new paragraphs, use `TextFrame.text`
        instead. Any other control characters in the assigned string are escaped as a hex
        representation like "_x001B_" (for ESC (ASCII 27) in this example).
        """
        return "".join(elm.text for elm in self._element.content_children)

    @text.setter
    def text(self, text: str):
        self.clear()
        self._element.append_text(text)

    @property
    def _defRPr(self) -> CT_TextCharacterProperties:
        """The element that defines the default run properties for runs in this paragraph.

        Causes the element to be added if not present.
        """
        return self._pPr.get_or_add_defRPr()

    @property
    def _pPr(self) -> CT_TextParagraphProperties:
        """Contains the properties for this paragraph.

        Causes the element to be added if not present.
        """
        return self._p.get_or_add_pPr()


class _Run(Subshape):
    """Text run object. Corresponds to `a:r` child element in a paragraph."""

    def __init__(self, r: CT_RegularTextRun, parent: ProvidesPart):
        super(_Run, self).__init__(parent)
        self._r = r

    def effective_font(self):
        """Return an `EffectiveFont` with this run's resolved size, name, and color.

        paper-pptx addition: executes the documented inheritance walk (run -> paragraph ->
        shape list style -> placeholder/layout/master styles -> presentation defaults ->
        theme) and reports the provenance chain for each value. Read-only. See
        `pptx.inspect.effective_font`.
        """
        from pptx.inspect import effective_font

        return effective_font(self)

    @property
    def font(self):
        """`Font` instance containing run-level character properties for the text in this run.

        Character properties can be and perhaps most often are inherited from parent objects such
        as the paragraph and slide layout the run is contained in. Only those specifically
        overridden at the run level are contained in the font object.
        """
        rPr = self._r.get_or_add_rPr()
        return Font(rPr)

    @lazyproperty
    def hyperlink(self) -> _Hyperlink:
        """Proxy for any `a:hlinkClick` element under the run properties element.

        Created on demand, the hyperlink object is available whether an `a:hlinkClick` element is
        present or not, and creates or deletes that element as appropriate in response to actions
        on its methods and attributes.
        """
        rPr = self._r.get_or_add_rPr()
        return _Hyperlink(rPr, self)

    @property
    def text(self):
        """Read/write. A unicode string containing the text in this run.

        Assignment replaces all text in the run. The assigned value can be a 7-bit ASCII
        string, a UTF-8 encoded 8-bit string, or unicode. String values are converted to
        unicode assuming UTF-8 encoding.

        Any other control characters in the assigned string other than tab or newline
        are escaped as a hex representation. For example, ESC (ASCII 27) is escaped as
        "_x001B_". Contrast the behavior of `TextFrame.text` and `_Paragraph.text` with
        respect to line-feed and vertical-tab characters.
        """
        return self._r.text

    @text.setter
    def text(self, text: str):
        self._r.text = text
