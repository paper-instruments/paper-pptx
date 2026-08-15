"""Contracts for slide partname allocation after a delete.

Upstream python-pptx could only append slides, so naming the next one `slide{count+1}`
was always free. `Slides.delete` is a paper-pptx addition that leaves gaps in the
sequence, after which the count no longer implies a free name.
"""

from __future__ import annotations

import pytest

from pptx import Presentation
from pptx.errors import PackageLimitError
from pptx.util import Inches

from .contract import save_reopen


def _deck(slide_count=5):
    prs = Presentation()
    layout = prs.slide_layouts[6]
    for idx in range(slide_count):
        slide = prs.slides.add_slide(layout)
        shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        shape.text_frame.text = "SLIDE-%d" % (idx + 1)
    return prs


def _slide_partnames(prs):
    return sorted(
        str(part.partname)
        for part in prs.part.package.iter_parts()
        if "/slides/slide" in str(part.partname)
    )


def _texts(prs):
    return [
        next(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame) for slide in prs.slides
    ]


def test_add_after_delete_reuses_the_freed_partname_not_a_live_one():
    """The gap left by the delete is what the next add must take."""
    prs = _deck()
    prs.slides.delete(prs.slides[1])  # frees /ppt/slides/slide2.xml
    assert "/ppt/slides/slide2.xml" not in _slide_partnames(prs)

    prs.slides.add_slide(prs.slide_layouts[6])

    partnames = _slide_partnames(prs)
    assert "/ppt/slides/slide2.xml" in partnames
    assert len(partnames) == len(set(partnames))


def test_deck_survives_save_reopen_after_delete_then_add():
    """The whole point: the written package must be readable."""
    prs = _deck()
    prs.slides.delete(prs.slides[1])
    new_slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = new_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    shape.text_frame.text = "NEW"

    assert _texts(save_reopen(prs)) == ["SLIDE-1", "SLIDE-3", "SLIDE-4", "SLIDE-5", "NEW"]


def test_repeated_delete_and_add_never_collides():
    """Churn is where a count-based allocator drifts furthest from the free set."""
    prs = _deck(slide_count=6)

    for _ in range(4):
        prs.slides.delete(prs.slides[0])
        prs.slides.add_slide(prs.slide_layouts[6])
        partnames = _slide_partnames(prs)
        assert len(partnames) == len(set(partnames)), partnames

    assert len(save_reopen(prs).slides) == 6


def test_writer_refuses_to_serialize_two_parts_sharing_a_partname(tmp_path):
    """Defence in depth: duplicate detection existed only on the reading side.

    A package whose parts collide serializes into a zip with duplicate member names;
    readers keep the last copy and the other part vanishes silently.
    """
    prs = _deck(slide_count=3)
    slide_parts = [
        part for part in prs.part.package.iter_parts() if "/slides/slide" in str(part.partname)
    ]
    slide_parts[1].partname = slide_parts[0].partname
    destination = tmp_path / "collision.pptx"

    with pytest.raises(PackageLimitError, match="sharing a partname"):
        prs.save(str(destination))

    assert not destination.exists()
