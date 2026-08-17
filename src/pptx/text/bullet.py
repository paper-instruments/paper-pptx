"""BulletFormat proxy, providing bullet and numbering control for a paragraph (paper-pptx).

Upstream python-pptx exposes paragraph indent levels but no way to make a real PowerPoint
bullet in an arbitrary text box, which pushes callers into fake glyph bullets ("• " or "- " as
literal text). `BulletFormat` writes the real `a:pPr` bullet vocabulary through the oxml
descriptor layer.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.text import PP_BULLET_TYPE
from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import (
    ST_TextAutonumberScheme,
    ST_TextBulletSizePercent,
    ST_TextBulletStartAtNum,
    ST_TextIndent,
    ST_TextMargin,
)
from pptx.util import Emu, Length

if TYPE_CHECKING:
    from pptx.oxml.text import CT_TextParagraph, CT_TextParagraphProperties

#: Default bullet geometry, from production-proven values: 0.375" left margin with a
#: 0.1875" hanging indent renders a conventionally-indented bullet in PowerPoint.
DEFAULT_BULLET_LEFT_MARGIN = Emu(342900)
DEFAULT_BULLET_HANGING_INDENT = Emu(171450)


def _require_xml_encodable(value: str, name: str) -> None:
    """Raise `ValueError` when `value` cannot be represented in XML 1.0.

    Part of validate-fully-then-mutate: a string that passes isinstance checks but explodes
    at serialization time would otherwise leave a half-mutated tree behind.
    """
    try:
        value.encode("utf-8")
    except UnicodeEncodeError:
        raise ValueError("%s contains characters not encodable in XML: %r" % (name, value))
    if any(
        not (
            ch in "\t\n\r"
            or "\x20" <= ch <= "\ud7ff"
            or "\ue000" <= ch <= "\ufffd"
            or "\U00010000" <= ch <= "\U0010ffff"
        )
        for ch in value
    ):
        raise ValueError("%s contains characters not permitted in XML 1.0: %r" % (name, value))


class BulletFormat(object):
    """Bullet and numbering state of a single paragraph.

    Read properties report this paragraph's **local** `a:pPr` state only: `None` means "nothing
    set here", in which case rendering inherits from the placeholder/list-style chain (use the
    effective-style inspection API to see what actually renders).
    """

    def __init__(self, p: CT_TextParagraph, part=None):
        super(BulletFormat, self).__init__()
        self._p = p
        self._part = part

    @property
    def type(self) -> PP_BULLET_TYPE | None:
        """Kind of bullet explicitly set on this paragraph, `None` when nothing local is set."""
        pPr = self._p.pPr
        if pPr is None:
            return None
        if pPr.buNone is not None:
            return PP_BULLET_TYPE.NONE
        if pPr.buAutoNum is not None:
            return PP_BULLET_TYPE.NUMBERED
        if pPr.buChar is not None:
            return PP_BULLET_TYPE.CHARACTER
        if pPr.find(qn("a:buBlip")) is not None:
            return PP_BULLET_TYPE.PICTURE
        return None

    @property
    def char(self) -> str | None:
        """The bullet character (`a:buChar/@char`), `None` unless a character bullet is set."""
        pPr = self._p.pPr
        buChar = pPr.buChar if pPr is not None else None
        return buChar.char if buChar is not None else None

    @property
    def number_scheme(self) -> str | None:
        """Numbering scheme token (`a:buAutoNum/@type`), e.g. "arabicPeriod", or `None`."""
        pPr = self._p.pPr
        buAutoNum = pPr.buAutoNum if pPr is not None else None
        return buAutoNum.type if buAutoNum is not None else None

    @property
    def start_at(self) -> int | None:
        """First number of an auto-numbered sequence, `None` unless numbering is set."""
        pPr = self._p.pPr
        buAutoNum = pPr.buAutoNum if pPr is not None else None
        return buAutoNum.startAt if buAutoNum is not None else None

    @property
    def font_name(self) -> str | None:
        """Bullet-specific typeface (`a:buFont/@typeface`), or `None` when not set locally."""
        pPr = self._p.pPr
        buFont = pPr.buFont if pPr is not None else None
        return buFont.typeface if buFont is not None else None

    @property
    def size_percent(self) -> float | None:
        """Bullet size as a fraction of the text size (`a:buSzPct`), e.g. 0.75, or `None`."""
        pPr = self._p.pPr
        buSzPct = pPr.buSzPct if pPr is not None else None
        return buSzPct.val if buSzPct is not None else None

    def set_character(
        self,
        char: str = "•",
        *,
        font_name: str | None = None,
        size_percent: float | None = None,
        left_margin: Length | None = DEFAULT_BULLET_LEFT_MARGIN,
        hanging_indent: Length | None = DEFAULT_BULLET_HANGING_INDENT,
    ) -> None:
        """Give this paragraph a real character bullet (`a:buChar`).

        `left_margin`/`hanging_indent` write `marL`/`indent` so the bullet hangs correctly;
        pass `None` for either to leave the existing paragraph attribute untouched.
        `font_name` sets a bullet-specific typeface (`a:buFont`); `size_percent` scales the
        bullet relative to the text size (fraction, 0.25–4.0).
        """
        self._require_attached()
        if not isinstance(char, str) or len(char) == 0:
            raise ValueError("char must be a non-empty str, got %r" % (char,))
        _require_xml_encodable(char, "char")
        self._validate_common(font_name, size_percent, left_margin, hanging_indent)
        with self._transaction():
            pPr = self._p.get_or_add_pPr()
            self._remove_buBlip(pPr)
            pPr.get_or_change_to_buChar().char = char
            self._apply_common(pPr, font_name, size_percent, left_margin, hanging_indent)

    def set_numbered(
        self,
        scheme: str = "arabicPeriod",
        *,
        start_at: int = 1,
        font_name: str | None = None,
        size_percent: float | None = None,
        left_margin: Length | None = DEFAULT_BULLET_LEFT_MARGIN,
        hanging_indent: Length | None = DEFAULT_BULLET_HANGING_INDENT,
    ) -> None:
        """Give this paragraph PowerPoint automatic numbering (`a:buAutoNum`).

        `scheme` is an ECMA-376 auto-number token like "arabicPeriod" or "romanUcParenR";
        an unknown token raises `ValueError`. Numbering restarts at `start_at`.
        """
        self._require_attached()
        if isinstance(start_at, bool):
            raise ValueError("start_at must be a positive int, got %r" % (start_at,))
        ST_TextAutonumberScheme.validate(scheme)
        ST_TextBulletStartAtNum.validate(start_at)
        self._validate_common(font_name, size_percent, left_margin, hanging_indent)
        with self._transaction():
            pPr = self._p.get_or_add_pPr()
            self._remove_buBlip(pPr)
            buAutoNum = pPr.get_or_change_to_buAutoNum()
            buAutoNum.type = scheme
            buAutoNum.startAt = start_at
            self._apply_common(pPr, font_name, size_percent, left_margin, hanging_indent)

    def set_none(self) -> None:
        """Set an explicit "no bullet" (`a:buNone`), overriding any inherited bullet.

        Margins, bullet font, and bullet size attributes are left untouched.
        """
        self._require_attached()
        with self._transaction():
            pPr = self._p.get_or_add_pPr()
            self._remove_buBlip(pPr)
            pPr.get_or_change_to_buNone()

    def _require_attached(self) -> None:
        if self._part is None:
            return
        from pptx._ownership import require_element_attached

        require_element_attached(self._p, self._part, argument="paragraph")

    def _transaction(self):
        if self._part is None:
            from contextlib import nullcontext

            return nullcontext()
        from pptx._transaction import PackageTransaction

        return PackageTransaction(self._part.package, self)

    @staticmethod
    def _remove_buBlip(pPr) -> None:
        """Remove any `a:buBlip` picture bullet from `pPr`.

        The bullet vocabulary is a 4-way schema choice including `a:buBlip`; the descriptor
        choice group covers only the three writable kinds, so an existing picture bullet
        must be cleared here or the written pPr would carry two bullet-type elements
        (schema-invalid).
        """
        for buBlip in pPr.findall(qn("a:buBlip")):
            pPr.remove(buBlip)

    @staticmethod
    def _validate_common(font_name, size_percent, left_margin, hanging_indent) -> None:
        """Validate shared setter arguments fully, BEFORE any mutation (§1.3 discipline).

        Routes through the same simpletype validators that guard the eventual XML writes, so
        validation and serialization can never disagree.
        """
        if font_name is not None and not isinstance(font_name, str):
            raise ValueError("font_name must be a str or None, got %r" % (font_name,))
        if font_name is not None:
            _require_xml_encodable(font_name, "font_name")
        if size_percent is not None:
            ST_TextBulletSizePercent.validate(size_percent)
        for name, value in (("left_margin", left_margin), ("hanging_indent", hanging_indent)):
            if value is not None and not isinstance(value, int):
                raise ValueError("%s must be a Length (EMU int) or None, got %r" % (name, value))
        if left_margin is not None:
            ST_TextMargin.validate(int(left_margin))
        if hanging_indent is not None:
            ST_TextIndent.validate(-abs(int(hanging_indent)))

    def _apply_common(
        self,
        pPr: CT_TextParagraphProperties,
        font_name: str | None,
        size_percent: float | None,
        left_margin: Length | None,
        hanging_indent: Length | None,
    ) -> None:
        if font_name is not None:
            pPr.get_or_add_buFont().typeface = font_name
        if size_percent is not None:
            pPr.get_or_add_buSzPct().val = size_percent
        if left_margin is not None:
            pPr.marL = Emu(left_margin)
        if hanging_indent is not None:
            pPr.indent = Emu(-abs(int(hanging_indent)))
