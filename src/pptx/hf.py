"""Footer, date, and slide-number application machinery (paper-pptx addition).

Reproduces what PowerPoint's Insert > Header & Footer dialog actually persists: the
dialog materializes minimal placeholder
`p:sp` shapes on each slide - `dt`/`ftr`/`sldNum`, with `idx` matching the slide layout's
furniture so geometry and formatting inherit; slide-number and auto-date content are real
`a:fld` elements; footer text and fixed dates are literal runs; `p:hf` flags are left
untouched (the dialog leaves them absent, and absent means all-visible).

This package AUTHORS fields; it never computes their live values. An `a:fld`'s cached
`a:t` text is a consumer-refreshed hint - PowerPoint and LibreOffice rewrite it on open
(proven by the `lo_footers_applied` fixture round-trip) - so the cache seeded here (slide
position for `slidenum`, an injectable-clock rendering for `datetime*`) is never something
the package vouches for at render time.
"""

from __future__ import annotations

import uuid
from copy import deepcopy
from datetime import datetime
from typing import TYPE_CHECKING, Dict, List, Optional, Tuple

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.errors import UnsupportedStructureError
from pptx.oxml.ns import qn

if TYPE_CHECKING:
    from pptx.presentation import Presentation
    from pptx.slide import Slide

# -- ISO 29500-1 21.1.2.2.4 reserved datetime field tokens, with the strftime rendering
# -- used ONLY to seed the consumer-refreshed cached text (spec example formats)
DATETIME_FIELD_FORMATS: "Dict[str, str]" = {
    "datetime": "%m/%d/%Y",
    "datetime1": "%m/%d/%Y",
    "datetime2": "%A, %B %d, %Y",
    "datetime3": "%d %B %Y",
    "datetime4": "%B %d, %Y",
    "datetime5": "%d-%b-%y",
    "datetime6": "%B %y",
    "datetime7": "%b-%y",
    "datetime8": "%m/%d/%Y %I:%M %p",
    "datetime9": "%m/%d/%Y %I:%M:%S %p",
    "datetime10": "%H:%M",
    "datetime11": "%H:%M:%S",
    "datetime12": "%I:%M %p",
    "datetime13": "%I:%M:%S %p",
}

_KINDS = ("dt", "ftr", "sldNum")
_PH_TYPE_FOR_KIND = {
    "dt": PP_PLACEHOLDER.DATE,
    "ftr": PP_PLACEHOLDER.FOOTER,
    "sldNum": PP_PLACEHOLDER.SLIDE_NUMBER,
}
_KIND_LABEL = {"dt": "date", "ftr": "footer", "sldNum": "slide-number"}


def apply_presentation_footers(
    prs: "Presentation",
    *,
    footer: "Optional[str]" = None,
    slide_number: bool = False,
    date_format: "Optional[str]" = None,
    fixed_date: "Optional[str]" = None,
    skip_title_slides: bool = False,
    now: "Optional[datetime]" = None,
) -> None:
    """Apply the complete footer state to every slide (the dialog's "Apply to All")."""
    from pptx._ownership import require_element_attached

    require_element_attached(prs._element, prs.part, argument="presentation")
    _validate_arguments(footer, slide_number, date_format, fixed_date, now)
    if not isinstance(skip_title_slides, bool):
        raise ValueError("skip_title_slides must be True or False")

    from pptx.errors import materialize_slides

    first_number = _first_slide_number(prs)
    plans: "List[Tuple[Slide, int, bool]]" = []
    for index, slide in enumerate(materialize_slides(prs, "apply_footers")):
        is_title = slide.slide_layout._element.get("type") == "title"
        all_off = skip_title_slides and is_title
        plans.append((slide, first_number + index, all_off))

    # -- validate the COMPLETE deck before the first write (refusal atomicity across
    # -- slides: a refusal on slide 4 must not leave slides 1-3 already rewritten)
    for slide, _, all_off in plans:
        if not all_off:
            _validate_slide_furniture(slide, footer, slide_number, date_format, fixed_date)

    from pptx._transaction import PackageTransaction

    with PackageTransaction(prs.part.package, prs, *(slide for slide, _, _ in plans)):
        for slide, number, all_off in plans:
            if all_off:
                _apply_to_slide(slide, number, None, False, None, None, now)
            else:
                _apply_to_slide(
                    slide, number, footer, slide_number, date_format, fixed_date, now
                )


def apply_slide_footers(
    slide: "Slide",
    *,
    footer: "Optional[str]" = None,
    slide_number: bool = False,
    date_format: "Optional[str]" = None,
    fixed_date: "Optional[str]" = None,
    now: "Optional[datetime]" = None,
) -> None:
    """Apply the complete footer state to one slide (the dialog's per-slide "Apply")."""
    from pptx._ownership import require_element_attached

    require_element_attached(slide._element, slide.part, argument="slide")
    _validate_arguments(footer, slide_number, date_format, fixed_date, now)
    prs = slide.part.package.presentation_part.presentation
    number = _first_slide_number(prs) + prs.slides.index(slide)
    _validate_slide_furniture(slide, footer, slide_number, date_format, fixed_date)
    from pptx._transaction import PackageTransaction

    with PackageTransaction(slide.part.package, slide):
        _apply_to_slide(slide, number, footer, slide_number, date_format, fixed_date, now)


# ------------------------------------------------------------------------------ validation


def _validate_arguments(footer, slide_number, date_format, fixed_date, now) -> None:
    """Check the header and footer arguments before any slide is touched."""
    if footer is not None:
        _validate_text("footer", footer)
    if not isinstance(slide_number, bool):
        raise ValueError("slide_number must be True or False")
    if date_format is not None and fixed_date is not None:
        raise ValueError(
            "date_format (automatic date field) and fixed_date (literal text) are the "
            "dialog's two exclusive date modes; pass only one"
        )
    if date_format is not None and date_format not in DATETIME_FIELD_FORMATS:
        raise ValueError(
            "date_format must be one of %s, got %r"
            % (", ".join(sorted(DATETIME_FIELD_FORMATS)), date_format)
        )
    if fixed_date is not None:
        _validate_text("fixed_date", fixed_date)
    if now is not None and not isinstance(now, datetime):
        raise ValueError("now must be a datetime.datetime or None, got %r" % (now,))


def _validate_text(name: str, value) -> None:
    """Refuse text that is not a non-empty str, holds a lone surrogate, or holds a C0 control.

    Tab is the one control allowed through, so a footer cannot span lines: a newline is refused.
    """
    if not isinstance(value, str) or not value:
        raise ValueError("%s must be a non-empty str, or None to remove" % name)
    try:
        value.encode("utf-8")
    except UnicodeEncodeError:
        raise ValueError("%s is not XML-encodable (lone surrogate)" % name)
    if any(ch < " " and ch != "\t" for ch in value):
        raise ValueError("%s must not contain C0 control characters" % name)


def _validate_slide_furniture(slide, footer, slide_number, date_format, fixed_date) -> None:
    """Refuse before any write when a wanted element cannot bind on `slide`'s layout."""
    wanted = _wanted_kinds(footer, slide_number, date_format, fixed_date)
    if not wanted:
        return
    layout = slide.slide_layout
    furniture = _layout_furniture(layout)
    for kind in wanted:
        if kind not in furniture:
            raise UnsupportedStructureError(
                "layout %r has no %s placeholder to inherit from; a materialized slide "
                "placeholder would not bind - add the placeholder to the layout first"
                % (layout.name or layout.part.partname, _KIND_LABEL[kind])
            )
        # -- nearest explicit p:hf declaration wins (layout over master; absent = inherit),
        # -- mirroring the HeaderFooters proxy's tri-state semantics
        master = layout.slide_master
        for level_name, owner_name, hf in (
            ("layout", layout.name or str(layout.part.partname), layout._element.hf),
            ("master", str(master.part.partname), master.element.hf),
        ):
            flag = getattr(hf, kind) if hf is not None else None
            if flag is False:
                raise UnsupportedStructureError(
                    "the %s %r carries p:hf flags disabling the %s placeholder, so "
                    "applied content would not render; clear the flag first via that "
                    "%s's header_footers"
                    % (level_name, owner_name, _KIND_LABEL[kind], level_name)
                )
            if flag is not None:
                break


def _wanted_kinds(footer, slide_number, date_format, fixed_date) -> "List[str]":
    """Which of date, footer, and slide number the caller asked to be present."""
    wanted = []
    if date_format is not None or fixed_date is not None:
        wanted.append("dt")
    if footer is not None:
        wanted.append("ftr")
    if slide_number:
        wanted.append("sldNum")
    return wanted


# -------------------------------------------------------------------------------- mechanics


def _first_slide_number(prs: "Presentation") -> int:
    """The number this deck starts counting slides from."""
    return int(prs._element.get("firstSlideNum", "1"))


def _layout_furniture(layout) -> dict:
    """Map kind -> layout placeholder proxy for the hf furniture present on `layout`."""
    found = {}
    for placeholder in layout.placeholders:
        for kind, ph_type in _PH_TYPE_FOR_KIND.items():
            if placeholder.element.ph_type == ph_type and kind not in found:
                found[kind] = placeholder
    return found


def _slide_phs(slide, kind):
    """ALL of the slide's own placeholder shapes of `kind`, in document order.

    Duplicates are schema-legal; each apply sets the dialog's one-per-kind state, so the
    first is updated and any extras are removed (deterministic, documented).
    """
    return [
        shape
        for shape in slide.shapes
        if shape.is_placeholder and shape.element.ph_type == _PH_TYPE_FOR_KIND[kind]
    ]


def _apply_to_slide(slide, number, footer, slide_number, date_format, fixed_date, now):
    """Write the wanted date, footer, and slide-number placeholders onto one slide.

    Clones the placeholder from the layout when the slide has none, and deletes the shapes for
    unwanted kinds along with any duplicates past the first.
    """
    wanted = _wanted_kinds(footer, slide_number, date_format, fixed_date)
    furniture = _layout_furniture(slide.slide_layout) if wanted else {}

    for kind in _KINDS:
        existing_shapes = _slide_phs(slide, kind)
        if kind not in wanted:
            for shape in existing_shapes:
                # -- the dialog's uncheck removes the shape; delete() keeps rel hygiene
                slide.shapes.delete(shape)
            continue

        for extra in existing_shapes[1:]:  # -- the dialog state is one shape per kind
            slide.shapes.delete(extra)
        existing = existing_shapes[0] if existing_shapes else None
        if existing is None:
            slide.shapes.clone_placeholder(furniture[kind])
            existing = _slide_phs(slide, kind)[0]
        txBody = existing._element.get_or_add_txBody()

        if kind == "sldNum":
            _write_field(txBody, "slidenum", str(number), _preserved_fld_id(txBody, "slidenum"))
        elif kind == "dt":
            if date_format is not None:
                cached = (now or datetime.now()).strftime(DATETIME_FIELD_FORMATS[date_format])
                _write_field(txBody, date_format, cached, _preserved_fld_id(txBody, "datetime"))
            else:
                _write_literal(txBody, fixed_date)
        else:  # -- ftr
            _write_literal(txBody, footer)


def _preserved_fld_id(txBody, type_prefix: str) -> str:
    """The existing field id for this family, or a fresh GUID.

    Per ISO 29500-1 21.1.2.2.4, a field's id token "persists in the file as the same
    token until the text field is removed" - reusing it makes re-application a no-op
    for unchanged content instead of a spurious rewrite.
    """
    for fld in txBody.iter(qn("a:fld")):
        if (fld.get("type") or "").startswith(type_prefix):
            existing = fld.get("id")
            if existing:
                return existing
    return "{%s}" % str(uuid.uuid4()).upper()


def _preserved_bits(txBody):
    """Deep-copies of the first paragraph's pPr, first content rPr, and endParaRPr."""
    if not txBody.p_lst:
        return None, None, None
    first_p = txBody.p_lst[0]
    pPr = first_p.find(qn("a:pPr"))
    rPr = None
    for child in first_p:
        if child.tag in (qn("a:r"), qn("a:fld")):
            rPr = child.find(qn("a:rPr"))
            break
    endParaRPr = first_p.find(qn("a:endParaRPr"))
    return (
        deepcopy(pPr) if pPr is not None else None,
        deepcopy(rPr) if rPr is not None else None,
        deepcopy(endParaRPr) if endParaRPr is not None else None,
    )


def _fresh_paragraph(txBody):
    """Clear `txBody` to a single empty paragraph, returning (p, preserved bits applied)."""
    pPr, rPr, endParaRPr = _preserved_bits(txBody)
    txBody.clear_content()
    txBody.unclear_content()  # -- guarantees exactly one empty a:p
    p = txBody.p_lst[0]
    if pPr is not None:
        p.append(pPr)  # -- p is empty; pPr lands first per the schema sequence
    return p, rPr, endParaRPr


def _write_field(txBody, field_type: str, cached_text: str, field_id: str) -> None:
    """Replace a placeholder's text with a field, so PowerPoint re-renders the value on open."""
    p, rPr, endParaRPr = _fresh_paragraph(txBody)
    fld = p.add_fld(field_id, field_type, cached_text)
    if rPr is not None:
        fld.insert(0, rPr)
    if endParaRPr is not None:
        p.append(endParaRPr)


def _write_literal(txBody, text: str) -> None:
    """Replace a placeholder's text with a literal run."""
    p, rPr, endParaRPr = _fresh_paragraph(txBody)
    r = p.add_r()
    r.text = text
    if rPr is not None:
        r.insert(0, rPr)
    if endParaRPr is not None:
        p.append(endParaRPr)
