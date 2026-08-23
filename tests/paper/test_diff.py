"""Contract tests: diff_decks, the verification mirror.

Required by the plan: diff(A, A) empty across the ENTIRE corpus; the reorder-only
fixture reads as moves (never delete-plus-add); determinism goldens; the declared
id-based matching contract. The release-level self-consistency checks (operation report
vs diff of input/output) live with the standing eval jobs in test_walkthrough_*.
"""

from __future__ import annotations

import io
import json
import zipfile

import pytest
from lxml import etree

from pptx import Presentation
from pptx.diff import ShapeRef, _common_boundary_ranges, diff_decks
from pptx.errors import UnsupportedStructureError
from pptx.oxml.ns import qn
from pptx.util import Inches

from . import corpus

V1 = "self_generated/lineage_v1.pptx"
V2 = "self_generated/lineage_v2.pptx"
REORDER = "self_generated/lineage_reorder.pptx"
WALNUT_CHART_NOTES = "other_producers/walnut_chart_notes_absolute_rels.pptx"
WALNUT_SHARED_MEDIA = "other_producers/walnut_shared_media_absolute_rels.pptx"
ALIGN_V1 = "self_generated/diff_alignment_v1.pptx"
ALIGN_V2 = "self_generated/diff_alignment_v2.pptx"

NONCORRUPT_RELPATHS = [
    r for r in corpus.iter_fixture_relpaths() if not corpus.is_corrupt_fixture(r)
]


def _path(relpath):
    return str(corpus.fixture_path(relpath))


# ------------------------------------------------------------------- the core invariants


@pytest.mark.parametrize("relpath", NONCORRUPT_RELPATHS)
def test_self_diff_is_empty_across_entire_corpus(relpath):
    report = diff_decks(_path(relpath), _path(relpath), detail="text")
    assert report.is_empty, report.to_dict()


@pytest.mark.parametrize("relpath", [WALNUT_CHART_NOTES, WALNUT_SHARED_MEDIA])
def test_deck_diff_ignores_walnut_serialization_only_package_rewrites(relpath, tmp_path):
    serialized = tmp_path / "ordinary-save.pptx"
    Presentation(_path(relpath)).save(str(serialized))

    report = diff_decks(_path(relpath), str(serialized), detail="full")
    assert report.is_empty, report.to_dict()


def test_reorder_only_fixture_reads_as_moves_never_delete_plus_add():
    report = diff_decks(_path(V1), _path(REORDER))
    assert report.slides_added == ()
    assert report.slides_removed == ()
    assert [(m.slide_id, m.from_position, m.to_position) for m in report.slides_moved] == [
        (260, 4, 0)
    ]
    assert report.slide_changes == ()


def test_lineage_pair_reports_the_exact_edit_list():
    """Every edit the v2 sidecar documents, attributed to the right slide id."""
    report = diff_decks(_path(V1), _path(V2), detail="text")

    assert [(r.slide_id, r.title) for r in report.slides_added] == [
        (261, "Lineage slide six, new")
    ]
    assert [(r.slide_id, r.title) for r in report.slides_removed] == [
        (260, "Lineage slide five")
    ]
    assert len(report.slides_moved) == 1  # -- one displacement (256<->257 tie, declared)

    changes = {change.slide_id: change for change in report.slide_changes}
    event = changes[256].text_changes[0]
    assert event["kind"] == "replacement"
    assert event["shape"] == ShapeRef(2, "Title 1")
    assert event["before"][0]["location"] == {
        "container": "shape",
        "paragraph_index": 0,
    }
    assert event["before_range"] == {"start": 0, "end": 1}
    assert event["after_range"] == {"start": 0, "end": 1}
    assert event["after"][0]["text"] == "Lineage slide one, retitled"
    assert changes[257].notes_change == {
        "before": "Original notes for slide two.",
        "after": "Updated notes for slide two.",
    }
    assert changes[258].chart_data_changes == (
        {
            "chart": ShapeRef(shape_id=3, name="lineage_chart"),
            "series": "FY",
            "category": "South",
            "before": 20.0,
            "after": 25.0,
        },
    )
    assert changes[259].images_replaced == (ShapeRef(shape_id=2, name="lineage_pic"),)
    assert changes[259].geometry_changes == (
        {
            "shape": ShapeRef(shape_id=3, name="lineage_box"),
            "facet": "left",
            "before": 3657600,
            "after": 4572000,
        },
    )
    assert report.to_dict()["version"] == 5


def test_diff_matches_frozen_golden():
    report = diff_decks(_path(V1), _path(V2), detail="text")
    actual = (
        json.dumps(report.to_dict(), indent=2, ensure_ascii=False) + "\n"
    ).encode("utf-8")
    golden_path = corpus.FIXTURES_DIR.parent / "goldens" / "lineage_v1_v2.diff.json"
    assert actual == golden_path.read_bytes()


# ------------------------------------------------------------------------ bullet shifts

BULLETED = "self_generated/branded_template.pptx"


def _saved(prs, tmp_path, name):
    path = tmp_path / name
    prs.save(str(path))
    return str(path)


def _bullet_shifts(before_path, after_path):
    report = diff_decks(before_path, after_path, detail="full")
    return [
        shift
        for change in report.slide_changes
        for shift in change.bullet_shifts
    ]


def test_suppressing_inherited_bullets_reports_a_bullet_shift(tmp_path):
    """The gap this facet closes.

    `set_none()` changes no text and no field marker, so `text_changes` stays empty while
    every visible bullet on the slide disappears. Before `bullet_shifts` the whole
    within-slide report was empty and only `package_changes` noticed the part had moved.
    """
    before = _saved(Presentation(_path(BULLETED)), tmp_path, "before.pptx")
    prs = Presentation(_path(BULLETED))
    for paragraph in prs.slides[0].placeholders[1].text_frame.paragraphs:
        paragraph.bullet.set_none()
    after = _saved(prs, tmp_path, "after.pptx")

    report = diff_decks(before, after, detail="full")
    change = report.slide_changes[0]
    assert change.text_changes == ()  # -- the text is untouched
    assert [(s.text, s.before["bullet"]["type"], s.after["bullet"]["type"]) for s in
            change.bullet_shifts] == [
        ("Body level one", "character", "none"),
        ("Body level two", "character", "none"),
    ]


def test_bullet_typeface_change_alone_reports_a_shift(tmp_path):
    """A symbol bullet losing its typeface keeps the same glyph and stops rendering as one.

    This is why the typeface had to join the resolver before this facet existed: comparing
    the kind and char alone would call these two decks identical.
    """
    wingding = "\uf0a7"  # -- private use: a filled square in Wingdings, else a box
    paths = []
    for name, font in (("wing.pptx", "Wingdings"), ("plain.pptx", None)):
        prs = Presentation(_path(BULLETED))
        for paragraph in prs.slides[0].placeholders[1].text_frame.paragraphs:
            paragraph.bullet.set_character(wingding, font_name=font)
        paths.append(_saved(prs, tmp_path, name))

    shifts = _bullet_shifts(*paths)
    assert shifts
    for shift in shifts:
        assert shift.before["bullet"]["char"] == shift.after["bullet"]["char"]  # -- same glyph
        assert shift.before["bullet_font"]["value"] == "Wingdings"
        assert shift.after["bullet_font"]["value"] != "Wingdings"


def test_bullet_shifts_are_absent_when_bullets_are_unchanged(tmp_path):
    """An omitted facet keeps every existing payload byte-identical."""
    before = _saved(Presentation(_path(BULLETED)), tmp_path, "a.pptx")
    prs = Presentation(_path(BULLETED))
    prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0].text = "Retitled"
    after = _saved(prs, tmp_path, "b.pptx")

    report = diff_decks(before, after, detail="full")
    assert report.slide_changes[0].text_changes  # -- the text change is reported
    assert report.slide_changes[0].bullet_shifts == ()
    assert "bullet_shifts" not in report.slide_changes[0].to_dict()


def test_bullet_shifts_only_populate_at_full_detail(tmp_path):
    before = _saved(Presentation(_path(BULLETED)), tmp_path, "c.pptx")
    prs = Presentation(_path(BULLETED))
    for paragraph in prs.slides[0].placeholders[1].text_frame.paragraphs:
        paragraph.bullet.set_none()
    after = _saved(prs, tmp_path, "d.pptx")

    assert _bullet_shifts(before, after)  # -- detail="full"
    for detail in ("structure", "text"):
        report = diff_decks(before, after, detail=detail)
        assert all(change.bullet_shifts == () for change in report.slide_changes)


def test_bullet_shift_payload_carries_provenance_for_both_sides(tmp_path):
    before = _saved(Presentation(_path(BULLETED)), tmp_path, "e.pptx")
    prs = Presentation(_path(BULLETED))
    for paragraph in prs.slides[0].placeholders[1].text_frame.paragraphs:
        paragraph.bullet.set_none()
    after = _saved(prs, tmp_path, "f.pptx")

    shift = _bullet_shifts(before, after)[0]
    payload = shift.to_dict()
    assert sorted(payload) == [
        "after",
        "after_location",
        "before",
        "before_location",
        "part",
        "shape",
        "text",
    ]
    assert payload["part"] == "/ppt/slides/slide1.xml"
    for side in ("before", "after"):
        assert sorted(payload[side]) == ["bullet", "bullet_font", "bullet_size"]
        assert payload[side]["bullet"]["provenance"]
    # -- the inherited side names the master; the suppressed side names the paragraph
    assert payload["before"]["bullet"]["provenance"][-1]["level"].startswith("master txStyles")
    assert payload["after"]["bullet"]["provenance"][0]["level"] == "paragraph pPr"


def test_diff_is_deterministic_across_runs():
    first = diff_decks(_path(V1), _path(V2), detail="full").to_dict()
    second = diff_decks(_path(V1), _path(V2), detail="full").to_dict()
    assert first == second


def test_package_changes_make_metadata_only_edit_nonempty():
    changed = Presentation(_path("self_generated/minimal_clean.pptx"))
    changed.core_properties.author = "Different author"
    report = diff_decks(_path("self_generated/minimal_clean.pptx"), changed, detail="full")
    assert not report.is_empty
    assert [delta.partname for delta in report.package_changes] == [
        "/docProps/core.xml"
    ]


def test_package_changes_make_z_order_only_edit_nonempty():
    changed = Presentation(_path("self_generated/gauntlet.pptx"))
    slide = changed.slides[2]
    slide.shapes.move(slide.shapes[0], len(slide.shapes) - 1)
    report = diff_decks(_path("self_generated/gauntlet.pptx"), changed, detail="full")
    assert not report.is_empty
    assert "/ppt/slides/slide3.xml" in {
        delta.partname for delta in report.package_changes
    }


@pytest.mark.parametrize("source_kind", ["path", "stream"])
def test_package_changes_use_original_package_members(tmp_path, source_kind):
    source = corpus.fixture_path("self_generated/minimal_clean.pptx")
    modified = io.BytesIO()
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(modified, "w") as after:
        for info in before.infolist():
            data = before.read(info.filename)
            if info.filename == "[Content_Types].xml":
                root = etree.fromstring(data)
                namespace = root.tag.partition("}")[0] + "}"
                etree.SubElement(
                    root,
                    namespace + "Default",
                    Extension="paper-unused",
                    ContentType="application/x-paper-unused",
                )
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
            after.writestr(info, data)
    modified.seek(0)

    if source_kind == "path":
        changed_source = tmp_path / "unused-content-type.pptx"
        changed_source.write_bytes(modified.getvalue())
    else:
        changed_source = modified

    report = diff_decks(source, changed_source)
    assert [delta.partname for delta in report.package_changes] == ["/[Content_Types].xml"]


def _deck_with_directory_records(tmp_path):
    """Copy a fixture, prefixing ZIP folder records that `save()` cannot reproduce.

    Three records, one of them nested, so a predicate keying on a top-level prefix
    cannot satisfy the assertions by accident.
    """
    target = tmp_path / "directory-records.pptx"
    source = corpus.fixture_path("self_generated/minimal_clean.pptx")
    with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(target, "w") as outgoing:
        for name in ("docProps/", "ppt/", "ppt/slides/"):
            record = zipfile.ZipInfo(name)
            record.external_attr = (0o040755 << 16) | 0x10
            outgoing.writestr(record, b"")
        for info in incoming.infolist():
            outgoing.writestr(info, incoming.read(info.filename))
    return str(target)


def test_directory_record_deck_gets_one_verdict_however_it_is_named(tmp_path):
    """A folder-record deck read exactly and read serialized must agree, and be empty.

    Folder records exist on disk and can never appear in a serialization, so comparing
    the disk side against the other side's rendering reported three removals that
    described the mismatched comparison rather than any difference in the document.
    """
    deck = _deck_with_directory_records(tmp_path)

    from_disk = diff_decks(deck, deck)
    from_proxy = diff_decks(deck, Presentation(deck))

    assert from_disk.is_empty == from_proxy.is_empty
    assert from_disk.is_empty, from_disk.to_dict()
    assert from_proxy.is_empty, from_proxy.to_dict()


def test_control_deck_gets_one_verdict_however_it_is_named():
    """The same agreement on a deck with no folder records, which held before the fix."""
    deck = _path("self_generated/minimal_clean.pptx")

    from_disk = diff_decks(deck, deck)
    from_proxy = diff_decks(deck, Presentation(deck))

    assert from_disk.is_empty == from_proxy.is_empty
    assert from_disk.is_empty, from_disk.to_dict()
    assert from_proxy.is_empty, from_proxy.to_dict()


def test_seekable_stream_positions_survive_success_and_failure():
    data = corpus.fixture_path("self_generated/minimal_clean.pptx").read_bytes()
    before = io.BytesIO(data)
    after = io.BytesIO(data)
    before.seek(11)
    after.seek(23)

    assert diff_decks(before, after).is_empty
    assert before.tell() == 11
    assert after.tell() == 23

    unreadable = io.BytesIO(b"not a presentation")
    before.seek(17)
    unreadable.seek(3)
    with pytest.raises(UnsupportedStructureError, match="not a readable presentation"):
        diff_decks(before, unreadable)
    assert before.tell() == 17
    assert unreadable.tell() == 3


def test_exact_package_stream_read_failure_is_typed_and_restores_position():
    class FailingExactRead(io.BytesIO):
        def read(self, size=-1):
            if self.tell() == 0 and size == 1024 * 1024:
                raise OSError("forced exact-map read failure")
            return super().read(size)

    data = corpus.fixture_path("self_generated/minimal_clean.pptx").read_bytes()
    source = FailingExactRead(data)
    source.seek(17)

    with pytest.raises(UnsupportedStructureError, match="cannot read before input stream"):
        diff_decks(source, _path("self_generated/minimal_clean.pptx"))

    assert source.tell() == 17


# ------------------------------------------------------------------------------ detail levels


def test_structure_detail_excludes_text_chart_and_notes():
    report = diff_decks(_path(V1), _path(V2), detail="structure")
    changed_ids = {change.slide_id for change in report.slide_changes}
    assert 259 in changed_ids  # -- geometry + image replacement are structural
    assert 256 not in changed_ids  # -- pure text edit invisible at structure level
    assert 257 not in changed_ids  # -- notes edit invisible at structure level
    assert 258 not in changed_ids  # -- chart data invisible at structure level


def test_full_detail_reports_effective_shifts():
    """Rebind a slide in a saved copy; diff(original, rebound) at full detail must
    carry the same run-level shifts the rebind report declared."""
    import tempfile
    from pathlib import Path

    source = Presentation(_path("self_generated/gauntlet.pptx"))
    rebind_report = source.slides[0].rebind_layout(
        source.slide_layouts[3]
    )  # -- Two Content: three genuine size shifts
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "rebound.pptx"
        source.save(str(out))
        diff = diff_decks(_path("self_generated/gauntlet.pptx"), str(out), detail="full")

    changed = {change.slide_id: change for change in diff.slide_changes}
    slide_id = Presentation(_path("self_generated/gauntlet.pptx")).slides[0].slide_id
    shifts = changed[slide_id].effective_shifts
    assert {(s.text, s.before["size"]["value"], s.after["size"]["value"]) for s in shifts} == {
        (s.text, s.before["size"]["value"], s.after["size"]["value"])
        for s in rebind_report.run_shifts
    }


def test_diff_reports_geometry_changed_by_layout_rebind():
    before = Presentation(_path("self_generated/minimal_clean.pptx"))
    after = Presentation(_path("self_generated/minimal_clean.pptx"))
    after.slides[0].rebind_layout(after.slide_layouts[1])

    report = diff_decks(before, after, detail="structure")

    geometry = report.slide_changes[0].geometry_changes
    assert geometry
    assert {change["facet"] for change in geometry} == {
        "left", "top", "width", "height"
    }


def test_bad_detail_raises_valueerror():
    with pytest.raises(ValueError, match="detail"):
        diff_decks(_path(V1), _path(V2), detail="everything")


# ---------------------------------------------------------------------- matching contract


def test_id_matching_contract_on_rebuilt_and_re_idd_decks():
    """The declared contract, asserted for real (the previous
    version was vacuous): colliding ids on a rebuilt deck MATCH and surface the content
    difference honestly; distinct ids read as full replacement."""
    rebuilt = Presentation()
    slide = rebuilt.slides.add_slide(rebuilt.slide_layouts[0])
    slide.shapes.title.text_frame.paragraphs[0].add_run().text = "Rebuilt title"

    # -- id collision (both decks start at 256): matched slide, text delta surfaced
    report = diff_decks(_path("self_generated/minimal_clean.pptx"), rebuilt)
    report_text = diff_decks(
        _path("self_generated/minimal_clean.pptx"), rebuilt, detail="text"
    )
    assert report.slides_added == ()
    assert report.slides_removed == ()
    changed_texts = [
        ([block["text"] for block in c["before"]], [block["text"] for block in c["after"]])
        for change in report_text.slide_changes
        for c in change.text_changes
    ]
    assert (["Minimal clean fixture"], ["Rebuilt title"]) in changed_texts

    # -- distinct ids: full replacement, never a spurious match
    rebuilt._element.sldIdLst.sldId_lst[0].set("id", "300")
    report2 = diff_decks(_path("self_generated/minimal_clean.pptx"), rebuilt)
    assert [ref.slide_id for ref in report2.slides_added] == [300]
    assert [ref.slide_id for ref in report2.slides_removed] == [256]


def test_duplicate_named_shape_additions_use_stable_references():
    prs_a = Presentation(_path("self_generated/minimal_clean.pptx"))
    slide = prs_a.slides[0]
    box_one = slide.shapes.add_textbox(0, 0, 914400, 914400)
    box_two = slide.shapes.add_textbox(914400, 0, 914400, 914400)
    box_one.name = "Duplicate"
    box_two.name = "Duplicate"
    box_two.text_frame.paragraphs[0].add_run().text = "changed"
    import tempfile
    from pathlib import Path

    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "dupes.pptx"
        prs_a.save(str(out))
        report = diff_decks(_path("self_generated/minimal_clean.pptx"), str(out))
    change = report.slide_changes[0]
    assert change.shapes_added == (
        ShapeRef(shape_id=4, name="Duplicate"),
        ShapeRef(shape_id=5, name="Duplicate"),
    )
    assert change.to_dict()["shapes_added"] == [
        {"shape_id": 4, "name": "Duplicate"},
        {"shape_id": 5, "name": "Duplicate"},
    ]


def test_user_shape_name_is_display_only_in_geometry_reference():
    base = Presentation(_path("self_generated/minimal_clean.pptx"))
    slide = base.slides[0]
    named = slide.shapes.add_textbox(0, 0, 914400, 914400)
    named.name = "sp#1"
    for left in (914400, 1828800):
        duplicate = slide.shapes.add_textbox(left, 0, 914400, 914400)
        duplicate.name = "Duplicate"
    stream = io.BytesIO()
    base.save(stream)
    changed = Presentation(io.BytesIO(stream.getvalue()))
    changed.slides[0].shapes.shape_by_name("sp#1").left += 100

    report = diff_decks(base, changed)
    geometry = [
        delta
        for change in report.slide_changes
        for delta in change.geometry_changes
    ]
    assert any(delta["shape"] == ShapeRef(shape_id=4, name="sp#1") for delta in geometry)


def _serialized(prs) -> bytes:
    stream = io.BytesIO()
    prs.save(stream)
    return stream.getvalue()


def test_duplicate_name_z_order_change_has_no_specialized_shape_facets():
    before = Presentation(_path("self_generated/minimal_clean.pptx"))
    slide = before.slides[0]
    first = slide.shapes.add_textbox(0, 0, 914400, 914400)
    second = slide.shapes.add_textbox(1828800, 0, 914400, 914400)
    first.name = second.name = "Duplicate"
    first.text = "First"
    second.text = "Second"
    before_bytes = _serialized(before)

    after = Presentation(io.BytesIO(before_bytes))
    after_slide = after.slides[0]
    after_slide.shapes.move(after_slide.shapes[-2], len(after_slide.shapes) - 1)
    report = diff_decks(io.BytesIO(before_bytes), after, detail="full")

    assert report.slide_changes == ()
    assert "/ppt/slides/slide1.xml" in {
        change.partname for change in report.package_changes
    }


def test_empty_and_renamed_shape_names_do_not_change_identity():
    before = Presentation(_path("self_generated/minimal_clean.pptx"))
    box = before.slides[0].shapes.add_textbox(0, 0, 914400, 914400)
    box.name = ""
    before_bytes = _serialized(before)

    after = Presentation(io.BytesIO(before_bytes))
    after.slides[0].shapes[-1].name = "Current display name"
    report = diff_decks(io.BytesIO(before_bytes), after)

    assert report.slide_changes == ()
    assert report.package_changes


def test_matched_geometry_uses_after_side_name_and_stable_id():
    before = Presentation(_path("self_generated/minimal_clean.pptx"))
    box = before.slides[0].shapes.add_textbox(0, 0, 914400, 914400)
    box.name = "Before name"
    shape_id = box.shape_id
    before_bytes = _serialized(before)

    after = Presentation(io.BytesIO(before_bytes))
    changed = after.slides[0].shapes[-1]
    changed.name = "After name"
    changed.left += 100
    report = diff_decks(io.BytesIO(before_bytes), after)
    change = report.slide_changes[0]

    assert change.shapes_added == ()
    assert change.shapes_removed == ()
    assert change.geometry_changes == (
        {
            "shape": ShapeRef(shape_id=shape_id, name="After name"),
            "facet": "left",
            "before": 0,
            "after": 100,
        },
    )
    assert change.to_dict()["geometry_changes"][0]["shape"] == {
        "shape_id": shape_id,
        "name": "After name",
    }


def test_addition_and_removal_use_their_own_side_names():
    before = Presentation(_path("self_generated/minimal_clean.pptx"))
    before_bytes = _serialized(before)
    removed = before.slides[0].shapes.title

    after = Presentation(io.BytesIO(before_bytes))
    after.slides[0].shapes.delete(after.slides[0].shapes.title)
    added = after.slides[0].shapes.add_textbox(0, 0, 914400, 914400)
    added.name = "After only"
    report = diff_decks(before, after)
    change = report.slide_changes[0]

    assert change.shapes_removed == (
        ShapeRef(shape_id=removed.shape_id, name=removed.name),
    )
    assert change.shapes_added == (
        ShapeRef(shape_id=added.shape_id, name="After only"),
    )


def test_incompatible_shape_kind_reuse_is_removal_plus_addition():
    from PIL import Image as PILImage

    before = Presentation()
    slide = before.slides.add_slide(before.slide_layouts[6])
    box = slide.shapes.add_textbox(0, 0, 914400, 914400)
    box.name = "Reused"
    shape_id = box.shape_id
    before_bytes = _serialized(before)

    after = Presentation(io.BytesIO(before_bytes))
    after_slide = after.slides[0]
    after_slide.shapes.delete(after_slide.shapes[0])
    image = io.BytesIO()
    PILImage.new("RGB", (2, 2), (10, 20, 30)).save(image, format="PNG")
    picture = after_slide.shapes.add_picture(io.BytesIO(image.getvalue()), 0, 0)
    picture.name = "Reused"
    assert picture.shape_id == shape_id

    report = diff_decks(io.BytesIO(before_bytes), after, detail="text")
    change = report.slide_changes[0]
    expected = (ShapeRef(shape_id=shape_id, name="Reused"),)
    assert change.shapes_removed == expected
    assert change.shapes_added == expected
    assert change.geometry_changes == ()
    assert change.images_replaced == ()
    assert change.chart_data_changes == ()
    assert len(change.text_changes) == 1
    assert change.text_changes[0]["kind"] == "deletion"
    assert _event_texts(change.text_changes[0], "before") == [""]
    assert change.text_changes[0]["after"] == []


def test_incompatible_graphic_frame_payload_reuse_is_removal_plus_addition():
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    before = Presentation()
    slide = before.slides.add_slide(before.slide_layouts[6])
    table = slide.shapes.add_table(2, 2, 0, 0, 914400, 914400)
    table.name = "Reused"
    shape_id = table.shape_id
    before_bytes = _serialized(before)

    after = Presentation(io.BytesIO(before_bytes))
    after_slide = after.slides[0]
    after_slide.shapes.delete(after_slide.shapes[0])
    chart_data = CategoryChartData()
    chart_data.categories = ["A"]
    chart_data.add_series("Series", (1,))
    chart = after_slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 0, 914400, 914400, chart_data
    )
    chart.name = "Reused"
    assert chart.shape_id == shape_id

    report = diff_decks(io.BytesIO(before_bytes), after)
    change = report.slide_changes[0]
    expected = (ShapeRef(shape_id=shape_id, name="Reused"),)
    assert change.shapes_removed == expected
    assert change.shapes_added == expected
    assert change.geometry_changes == ()
    assert change.chart_data_changes == ()


def test_duplicate_shape_ids_refuse_with_deterministic_candidate_details():
    before = Presentation(_path("self_generated/minimal_clean.pptx"))
    after = Presentation(_path("self_generated/minimal_clean.pptx"))
    first = after.slides[0].shapes.add_textbox(0, 0, 914400, 914400)
    group = after.slides[0].shapes.add_group_shape()
    second = group.shapes.add_textbox(914400, 0, 914400, 914400)
    first.name = "First duplicate"
    second.name = "Second duplicate"
    second._element._nvXxPr.cNvPr.set("id", str(first.shape_id))

    messages = []
    for _ in range(2):
        with pytest.raises(UnsupportedStructureError) as exc_info:
            diff_decks(before, after)
        messages.append(str(exc_info.value))
    assert messages[0] == messages[1]
    assert messages[0] == (
        "diff_decks refused: after side of slide 256 has duplicate shape IDs "
        "(id 4: 2 candidates [sp name='First duplicate', sp name='Second duplicate']); "
        "shape IDs must be unique across the slide"
    )


def test_moving_leaf_into_group_remains_a_top_level_removal():
    before = Presentation()
    slide = before.slides.add_slide(before.slide_layouts[6])
    leaf = slide.shapes.add_textbox(0, 0, 914400, 914400)
    leaf.name = "Leaf"
    group = slide.shapes.add_group_shape()
    before_bytes = _serialized(before)

    after = Presentation(io.BytesIO(before_bytes))
    after_slide = after.slides[0]
    after_leaf = next(shape for shape in after_slide.shapes if shape.shape_id == leaf.shape_id)
    after_group = next(shape for shape in after_slide.shapes if shape.shape_id == group.shape_id)
    after_group._element.append(after_leaf._element)

    report = diff_decks(io.BytesIO(before_bytes), after)
    change = report.slide_changes[0]
    assert change.shapes_removed == (ShapeRef(shape_id=leaf.shape_id, name="Leaf"),)
    assert change.shapes_added == ()
    assert change.geometry_changes == ()


# ------------------------------------------------------ exact paragraph/table alignment


def _alignment_deck(paragraphs, rows=None):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    box.name = "alignment_text"
    box.text_frame.paragraphs[0].text = paragraphs[0]
    for text in paragraphs[1:]:
        box.text_frame.add_paragraph().text = text
    if rows is not None:
        frame = slide.shapes.add_table(
            len(rows), len(rows[0]), Inches(1), Inches(3.5), Inches(6), Inches(2)
        )
        frame.name = "alignment_table"
        for row_index, values in enumerate(rows):
            for column_index, text in enumerate(values):
                frame.table.cell(row_index, column_index).text = text
    return prs


def _event_texts(event, side):
    return [block["text"] for block in event[side]]


def _set_cell_paragraphs(cell, paragraphs):
    text_frame = cell.text_frame
    text_frame.paragraphs[0].text = paragraphs[0]
    for text in paragraphs[1:]:
        text_frame.add_paragraph().text = text


def _group_alignment_table(prs):
    slide = prs.slides[0]
    table = next(shape for shape in slide.shapes if shape.has_table)
    group = slide.shapes.add_group_shape()
    group.name = "table_group"
    group._element.append(table._element)


def _nest_alignment_text(prs, group_names):
    slide = prs.slides[0]
    text_box = slide.shapes[0]
    outer = slide.shapes.add_group_shape()
    outer.name = group_names[0]
    inner = outer.shapes.add_group_shape()
    inner.name = group_names[1]
    inner._element.append(text_box._element)


@pytest.mark.parametrize(
    ("before", "after", "location"),
    [
        (("A", "B"), ("X", "A", "B"), 0),
        (("A", "B"), ("A", "X", "B"), 1),
        (("A", "B"), ("A", "B", "X"), 2),
    ],
)
def test_paragraph_insertion_reports_only_the_inserted_block(before, after, location):
    report = diff_decks(_alignment_deck(before), _alignment_deck(after), detail="text")
    events = report.slide_changes[0].text_changes
    assert [(event["kind"], _event_texts(event, "after")) for event in events] == [
        ("insertion", ["X"])
    ]
    assert events[0]["before"] == []
    assert events[0]["before_range"] == {"start": location, "end": location}
    assert events[0]["after_range"] == {"start": location, "end": location + 1}
    assert events[0]["after"][0]["location"]["paragraph_index"] == location
    assert events[0]["after"][0]["fields"] == []


@pytest.mark.parametrize(
    ("before", "after", "location"),
    [
        (("X", "A", "B"), ("A", "B"), 0),
        (("A", "X", "B"), ("A", "B"), 1),
        (("A", "B", "X"), ("A", "B"), 2),
    ],
)
def test_paragraph_deletion_reports_only_the_deleted_block(before, after, location):
    event = diff_decks(
        _alignment_deck(before), _alignment_deck(after), detail="text"
    ).slide_changes[0].text_changes[0]
    assert event["kind"] == "deletion"
    assert _event_texts(event, "before") == ["X"]
    assert event["after"] == []
    assert event["before_range"] == {"start": location, "end": location + 1}
    assert event["after_range"] == {"start": location, "end": location}


def test_paragraph_deletion_and_bounded_replacement_do_not_cascade():
    deletion = (
        diff_decks(_alignment_deck(("A", "drop", "B")), _alignment_deck(("A", "B")), detail="text")
        .slide_changes[0]
        .text_changes
    )
    replacement = (
        diff_decks(
            _alignment_deck(("A", "old", "B")),
            _alignment_deck(("A", "new", "B")),
            detail="text",
        )
        .slide_changes[0]
        .text_changes
    )
    assert [(event["kind"], _event_texts(event, "before")) for event in deletion] == [
        ("deletion", ["drop"])
    ]
    assert [
        (event["kind"], _event_texts(event, "before"), _event_texts(event, "after"))
        for event in replacement
    ] == [
        ("replacement", ["old"], ["new"])
    ]


def test_paragraph_reorder_is_one_changed_region_without_move_claims():
    events = (
        diff_decks(
            _alignment_deck(("A", "B", "C")),
            _alignment_deck(("B", "A", "C")),
            detail="text",
        )
        .slide_changes[0]
        .text_changes
    )
    assert len(events) == 1
    assert events[0]["kind"] == "changed_region"
    assert _event_texts(events[0], "before") == ["A", "B"]
    assert _event_texts(events[0], "after") == ["B", "A"]


def test_multiple_disjoint_changes_are_one_canonical_changed_region():
    events = (
        diff_decks(
            _alignment_deck(("edge", "old one", "same", "old two", "tail")),
            _alignment_deck(("edge", "new one", "same", "new two", "tail")),
            detail="text",
        )
        .slide_changes[0]
        .text_changes
    )
    assert len(events) == 1
    assert events[0]["kind"] == "changed_region"
    assert events[0]["before_range"] == {"start": 1, "end": 4}
    assert events[0]["after_range"] == {"start": 1, "end": 4}
    assert _event_texts(events[0], "before") == ["old one", "same", "old two"]
    assert _event_texts(events[0], "after") == ["new one", "same", "new two"]


def test_nested_group_text_alignment_uses_leaf_id_not_group_path():
    before = _alignment_deck(("A", "B"))
    after = _alignment_deck(("X", "A", "B"))
    _nest_alignment_text(before, ("old outer", "old inner"))
    _nest_alignment_text(after, ("new outer", "new inner"))

    events = diff_decks(before, after, detail="text").slide_changes[0].text_changes
    assert len(events) == 1
    assert events[0]["kind"] == "insertion"
    assert events[0]["shape"] == ShapeRef(2, "alignment_text")
    assert events[0]["after"][0]["location"] == {
        "container": "group",
        "paragraph_index": 0,
    }


def test_repeated_region_uses_prefix_first_canonical_replacement():
    before = _alignment_deck(("A", "same", "same", "B"))
    unchanged = _alignment_deck(("A", "same", "same", "B"))
    changed = _alignment_deck(("A", "same", "changed", "B"))
    assert diff_decks(before, unchanged, detail="text").slide_changes == ()
    events = diff_decks(before, changed, detail="text").slide_changes[0].text_changes
    assert len(events) == 1
    assert events[0]["kind"] == "replacement"
    assert events[0]["before_range"] == {"start": 2, "end": 3}
    assert _event_texts(events[0], "before") == ["same"]
    assert _event_texts(events[0], "after") == ["changed"]


def test_repeated_insertion_uses_longest_prefix_before_suffix():
    events = diff_decks(
        _alignment_deck(("same", "same")),
        _alignment_deck(("same", "same", "same")),
        detail="text",
    ).slide_changes[0].text_changes
    assert len(events) == 1
    assert events[0]["kind"] == "insertion"
    assert events[0]["after_range"] == {"start": 2, "end": 3}


def test_repeated_exact_boundaries_can_establish_a_replacement():
    events = (
        diff_decks(
            _alignment_deck(("same", "old", "same")),
            _alignment_deck(("same", "new", "same")),
            detail="text",
        )
        .slide_changes[0]
        .text_changes
    )
    assert [
        (event["kind"], _event_texts(event, "before"), _event_texts(event, "after"))
        for event in events
    ] == [
        ("replacement", ["old"], ["new"])
    ]


def test_nfc_representation_and_whitespace_remain_content():
    nfc = diff_decks(
        _alignment_deck(("caf\N{LATIN SMALL LETTER E WITH ACUTE}",)),
        _alignment_deck(("cafe\N{COMBINING ACUTE ACCENT}",)),
        detail="text",
    )
    nfc_event = nfc.slide_changes[0].text_changes[0]
    assert nfc_event["kind"] == "replacement"
    assert nfc_event["before"][0]["location"] == {
        "container": "shape",
        "paragraph_index": 0,
    }
    assert nfc_event["after"][0]["location"] == {
        "container": "shape",
        "paragraph_index": 0,
    }
    whitespace = diff_decks(_alignment_deck(("A  B",)), _alignment_deck(("A B",)), detail="text")
    assert whitespace.slide_changes[0].text_changes[0]["kind"] == "replacement"


def test_table_row_insertion_reports_grid_and_only_real_text_insertions():
    before = _alignment_deck(("stable",), (("Name", "Value"), ("Alpha", "10")))
    after = _alignment_deck(("stable",), (("Name", "Value"), ("New", "5"), ("Alpha", "10")))
    structure = diff_decks(before, after, detail="structure").slide_changes[0]
    assert structure.table_structure_changes[0] == {
        "table": ShapeRef(3, "alignment_table"),
        "before": {"rows": 2, "columns": 2},
        "after": {"rows": 3, "columns": 2},
    }
    assert structure.text_changes == ()
    text = diff_decks(before, after, detail="text").slide_changes[0]
    assert len(text.text_changes) == 1
    assert text.text_changes[0]["kind"] == "insertion"
    assert _event_texts(text.text_changes[0], "after") == ["New", "5"]
    assert [block["location"] for block in text.text_changes[0]["after"]] == [
        {"container": "table-cell", "paragraph_index": 0, "row": 1, "column": 0},
        {"container": "table-cell", "paragraph_index": 0, "row": 1, "column": 1},
    ]

    deletion = diff_decks(after, before, detail="text").slide_changes[0]
    assert deletion.table_structure_changes[0]["before"]["rows"] == 3
    assert deletion.table_structure_changes[0]["after"]["rows"] == 2
    assert len(deletion.text_changes) == 1
    assert deletion.text_changes[0]["kind"] == "deletion"
    assert _event_texts(deletion.text_changes[0], "before") == ["New", "5"]


def test_table_column_insertion_and_deletion_keep_exact_cell_coordinates():
    narrow = _alignment_deck(("stable",), (("A", "B"), ("C", "D")))
    wide = _alignment_deck(("stable",), (("A", "X", "B"), ("C", "Y", "D")))

    insertion = diff_decks(narrow, wide, detail="text").slide_changes[0]
    assert insertion.table_structure_changes[0] == {
        "table": ShapeRef(3, "alignment_table"),
        "before": {"rows": 2, "columns": 2},
        "after": {"rows": 2, "columns": 3},
    }
    assert len(insertion.text_changes) == 1
    assert insertion.text_changes[0]["kind"] == "changed_region"
    assert _event_texts(insertion.text_changes[0], "before") == ["B", "C"]
    assert _event_texts(insertion.text_changes[0], "after") == ["X", "B", "C", "Y"]
    assert insertion.text_changes[0]["after"][0]["location"] == {
        "container": "table-cell",
        "paragraph_index": 0,
        "row": 0,
        "column": 1,
    }

    deletion = diff_decks(wide, narrow, detail="text").slide_changes[0]
    assert deletion.table_structure_changes[0]["before"]["columns"] == 3
    assert deletion.table_structure_changes[0]["after"]["columns"] == 2
    assert len(deletion.text_changes) == 1
    assert deletion.text_changes[0]["kind"] == "changed_region"


def test_multiple_paragraphs_in_one_table_cell_do_not_collide_with_other_cells_or_tables():
    before = _alignment_deck(("stable",), (("left", "right"),))
    after = _alignment_deck(("stable",), (("left", "right"),))
    _set_cell_paragraphs(before.slides[0].shapes[1].table.cell(0, 0), ("A", "B"))
    _set_cell_paragraphs(after.slides[0].shapes[1].table.cell(0, 0), ("A", "new", "B"))
    for prs in (before, after):
        second = prs.slides[0].shapes.add_table(
            1, 1, Inches(7), Inches(3.5), Inches(2), Inches(1)
        )
        _set_cell_paragraphs(second.table.cell(0, 0), ("A", "B"))

    events = diff_decks(before, after, detail="text").slide_changes[0].text_changes
    assert len(events) == 1
    assert events[0]["kind"] == "insertion"
    assert _event_texts(events[0], "after") == ["new"]
    assert events[0]["after"][0]["location"] == {
        "container": "table-cell",
        "paragraph_index": 1,
        "row": 0,
        "column": 0,
    }


def test_grouped_table_uses_leaf_shape_identity_for_grid_and_text():
    before = _alignment_deck(("stable",), (("A", "B"),))
    after = _alignment_deck(("stable",), (("A", "X", "B"),))
    _group_alignment_table(before)
    _group_alignment_table(after)

    change = diff_decks(before, after, detail="text").slide_changes[0]
    grid = change.table_structure_changes[0]
    assert grid["table"] == ShapeRef(3, "alignment_table")
    assert grid["before"] == {"rows": 1, "columns": 2}
    assert grid["after"] == {"rows": 1, "columns": 3}
    assert [(event["kind"], _event_texts(event, "after")) for event in change.text_changes] == [
        ("insertion", ["X"])
    ]
    assert change.text_changes[0]["shape"] == ShapeRef(3, "alignment_table")


def test_added_or_removed_table_has_no_matched_grid_change():
    without_table = _alignment_deck(("stable",))
    with_table = _alignment_deck(("stable",), (("cell",),))

    added = diff_decks(without_table, with_table, detail="structure").slide_changes[0]
    removed = diff_decks(with_table, without_table, detail="structure").slide_changes[0]
    assert added.shapes_added == (ShapeRef(3, "alignment_table"),)
    assert removed.shapes_removed == (ShapeRef(3, "alignment_table"),)
    assert added.table_structure_changes == ()
    assert removed.table_structure_changes == ()


def test_duplicate_blank_table_row_reports_only_factual_dimensions():
    before = _alignment_deck(("stable",), (("", ""), ("", "")))
    after = _alignment_deck(("stable",), (("", ""), ("", ""), ("", "")))
    change = diff_decks(before, after, detail="text").slide_changes[0]
    assert change.table_structure_changes[0] == {
        "table": ShapeRef(3, "alignment_table"),
        "before": {"rows": 2, "columns": 2},
        "after": {"rows": 3, "columns": 2},
    }
    assert len(change.text_changes) == 1
    assert change.text_changes[0]["kind"] == "insertion"
    assert change.text_changes[0]["after_range"] == {"start": 4, "end": 6}


def test_effective_and_bullet_shifts_follow_exact_alignment_after_insertion():
    before = _alignment_deck(("A", "target"))
    before_target = before.slides[0].shapes[0].text_frame.paragraphs[1]
    before_target.bullet.set_character("•")
    after = _alignment_deck(("inserted", "A", "target"))
    after_target = after.slides[0].shapes[0].text_frame.paragraphs[2]
    after_target.runs[0].font.bold = True
    after_target.bullet.set_none()

    change = diff_decks(before, after, detail="full").slide_changes[0]
    shift = next(shift for shift in change.effective_shifts if shift.text == "target")
    assert shift.before_location["paragraph_index"] == 1
    assert shift.after_location["paragraph_index"] == 2
    bullet = next(shift for shift in change.bullet_shifts if shift.text == "target")
    assert bullet.before_location["paragraph_index"] == 1
    assert bullet.after_location["paragraph_index"] == 2


def test_repeated_paragraphs_cannot_create_false_full_detail_shifts():
    before = _alignment_deck(("same", "same"))
    before_paragraph = before.slides[0].shapes[0].text_frame.paragraphs[1]
    before_paragraph.runs[0].font.bold = True
    before_paragraph.bullet.set_character("•")

    after = _alignment_deck(("same", "same", "same"))
    after_paragraph = after.slides[0].shapes[0].text_frame.paragraphs[2]
    after_paragraph.runs[0].font.bold = True
    after_paragraph.bullet.set_character("•")

    change = diff_decks(before, after, detail="full").slide_changes[0]
    assert change.text_changes[0]["kind"] == "insertion"
    assert change.effective_shifts == ()
    assert change.bullet_shifts == ()


def test_changed_region_has_no_specialized_formatting_correspondence():
    before = _alignment_deck(("old",))
    after = _alignment_deck(("new",))
    after.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.bold = True

    change = diff_decks(before, after, detail="full").slide_changes[0]
    assert change.text_changes[0]["kind"] == "replacement"
    assert change.effective_shifts == ()
    assert change.bullet_shifts == ()


def test_run_splitting_does_not_create_a_text_event():
    before = _alignment_deck(("AB",))
    after = _alignment_deck(("",))
    paragraph = after.slides[0].shapes[0].text_frame.paragraphs[0]
    paragraph.add_run().text = "A"
    paragraph.add_run().text = "B"

    report = diff_decks(before, after, detail="text")
    assert report.slide_changes == ()


def test_boundary_scan_uses_linear_equality_operations():
    class ComparableText:
        comparisons = 0

        def __init__(self, value):
            self.value = value

        def __eq__(self, other):
            type(self).comparisons += 1
            return self.value == other.value

    class Block:
        def __init__(self, value):
            self.text = ComparableText(value)
            self.runs = ()

    for size in (8, 64, 512):
        identical_a = tuple(Block(index) for index in range(size))
        identical_b = tuple(Block(index) for index in range(size))
        ComparableText.comparisons = 0
        assert _common_boundary_ranges(identical_a, identical_b) == ((size, size), (size, size))
        assert ComparableText.comparisons <= size + 1

        different_a = tuple(Block(("a", index)) for index in range(size))
        different_b = tuple(Block(("b", index)) for index in range(size))
        ComparableText.comparisons = 0
        assert _common_boundary_ranges(different_a, different_b) == ((0, size), (0, size))
        assert ComparableText.comparisons <= 2 * size + 2


def test_table_effective_and_bullet_formatting_remain_out_of_scope():
    before = _alignment_deck(("stable",), (("cell",),))
    after = _alignment_deck(("stable",), (("cell",),))
    paragraph = after.slides[0].shapes[1].table.cell(0, 0).text_frame.paragraphs[0]
    paragraph.runs[0].font.bold = True
    paragraph.bullet.set_character("•")

    report = diff_decks(before, after, detail="full")
    assert report.slide_changes == ()
    assert report.package_changes


def test_frozen_alignment_pair_preserves_real_changes_without_cascade():
    report = diff_decks(_path(ALIGN_V1), _path(ALIGN_V2), detail="text")
    change = report.slide_changes[0]
    assert change.table_structure_changes[0]["before"] == {"rows": 2, "columns": 2}
    assert change.table_structure_changes[0]["after"] == {"rows": 3, "columns": 2}
    assert [(event["kind"], _event_texts(event, "after")) for event in change.text_changes] == [
        ("insertion", ["Inserted paragraph"]),
        ("insertion", ["New", "5"]),
    ]


# ------------------------------------------------------------- regressions


def test_trailing_whitespace_edit_is_a_reported_text_change():
    """Regression: whitespace is content. The frozen
    trailing-space pair must diff as a text change, never as 'identical'."""
    report = diff_decks(
        _path("self_generated/whitespace_trailing_a.pptx"),
        _path("self_generated/whitespace_trailing_b.pptx"),
        detail="text",
    )
    deltas = [
        (_event_texts(c, "before"), _event_texts(c, "after"))
        for change in report.slide_changes
        for c in change.text_changes
    ]
    assert (["Trailing space "], ["Trailing space"]) in deltas


def test_field_type_change_is_reported_when_visible_text_is_unchanged():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    paragraph = slide.shapes.add_textbox(
        0, 0, 914400, 914400
    ).text_frame.paragraphs[0]
    paragraph.add_slide_number_field()
    before = io.BytesIO()
    prs.save(before)

    changed = Presentation(io.BytesIO(before.getvalue()))
    field = changed.slides[0].shapes[0].text_frame.paragraphs[0]._p.find(qn("a:fld"))
    field.set("type", "datetime1")

    report = diff_decks(io.BytesIO(before.getvalue()), changed, detail="text")
    change = report.slide_changes[0].text_changes[0]
    assert _event_texts(change, "before") == _event_texts(change, "after") == [""]
    assert change["before"][0]["fields"] == [{"offset": 0, "type": "slidenum"}]
    assert change["after"][0]["fields"] == [{"offset": 0, "type": "datetime1"}]


def test_field_movement_is_reported_without_false_formatting_shifts():
    before = Presentation()
    slide = before.slides.add_slide(before.slide_layouts[6])
    paragraph = slide.shapes.add_textbox(0, 0, 914400, 914400).text_frame.paragraphs[0]
    first = paragraph.add_run()
    first.text = "Page "
    first.font.bold = True
    paragraph.add_slide_number_field()
    paragraph.add_run().text = " of total"
    stream = io.BytesIO()
    before.save(stream)
    after = Presentation(io.BytesIO(stream.getvalue()))
    changed_paragraph = after.slides[0].shapes[0].text_frame.paragraphs[0]._p
    field = changed_paragraph.find(qn("a:fld"))
    changed_paragraph.remove(field)
    pPr = changed_paragraph.find(qn("a:pPr"))
    changed_paragraph.insert(1 if pPr is not None else 0, field)

    report = diff_decks(before, after, detail="full")
    change = report.slide_changes[0].text_changes[0]

    assert change["before"][0]["fields"] == [{"offset": 5, "type": "slidenum"}]
    assert change["after"][0]["fields"] == [{"offset": 0, "type": "slidenum"}]
    assert report.slide_changes[0].effective_shifts == ()


def test_inserted_literal_run_does_not_create_false_formatting_shifts():
    before = Presentation()
    slide = before.slides.add_slide(before.slide_layouts[6])
    paragraph = slide.shapes.add_textbox(0, 0, 914400, 914400).text_frame.paragraphs[0]
    first = paragraph.add_run()
    first.text = "A"
    first.font.bold = True
    paragraph.add_run().text = "B"
    stream = io.BytesIO()
    before.save(stream)
    after = Presentation(io.BytesIO(stream.getvalue()))
    changed = after.slides[0].shapes[0].text_frame.paragraphs[0]
    inserted = changed.add_run()
    inserted.text = "X"
    inserted.font.italic = True
    changed._p.remove(inserted._r)
    first_run = changed._p.find(qn("a:r"))
    first_run.addprevious(inserted._r)

    report = diff_decks(before, after, detail="full")

    assert report.slide_changes[0].effective_shifts == ()


def test_full_detail_sees_emphasis_shifts():
    """Regression: bold/italic/underline participate in resolution-state
    comparison - a mutant dropping them must fail here."""
    prs_b = Presentation(_path("self_generated/minimal_clean.pptx"))
    run = prs_b.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
    run.font.bold = True
    run.font.italic = True
    diff = diff_decks(_path("self_generated/minimal_clean.pptx"), prs_b, detail="full")
    shifts = [s for change in diff.slide_changes for s in change.effective_shifts]
    assert len(shifts) == 1
    shift = shifts[0]
    assert shift.before["bold"]["value"] is False
    assert shift.after["bold"]["value"] is True
    assert shift.before["italic"]["value"] is False
    assert shift.after["italic"]["value"] is True


def test_shape_removal_does_not_misattribute_later_blocks():
    """Regression: text/effective comparison keys are shape-scoped, so
    removing an early shape must read as that shape's blocks disappearing - never as
    edits to the shapes below it."""
    prs_b = Presentation(_path("self_generated/gauntlet.pptx"))
    slide = prs_b.slides[0]
    title = slide.shapes.title
    slide.shapes.delete(title)
    diff = diff_decks(_path("self_generated/gauntlet.pptx"), prs_b, detail="full")
    change = next(c for c in diff.slide_changes)
    assert ShapeRef(shape_id=2, name="Title 1") in change.shapes_removed
    # -- the title's block reads as removed (after=[]); the body blocks are untouched
    for delta in change.text_changes:
        assert delta["after"] == [], delta
    assert change.effective_shifts == ()  # -- no phantom shifts from index slippage


def test_diff_accepts_presentation_objects():
    """Regression: the natural first attempt - passing open decks."""
    prs_a = Presentation(_path(V1))
    prs_b = Presentation(_path(V2))
    report = diff_decks(prs_a, prs_b, detail="structure")
    assert [ref.slide_id for ref in report.slides_added] == [261]


def test_unreadable_package_refuses_typed(tmp_path):
    from pptx.errors import UnsupportedStructureError

    garbage = tmp_path / "not-a-deck.pptx"
    garbage.write_bytes(b"this is not a zip archive")
    with pytest.raises(UnsupportedStructureError, match="not a readable presentation"):
        diff_decks(str(garbage), _path(V1))


def test_text_diff_refuses_notes_without_a_body_placeholder():
    from pptx.enum.shapes import PP_PLACEHOLDER

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    notes_slide = slide.notes_slide
    body = next(
        shape
        for shape in notes_slide.placeholders
        if shape.placeholder_format.type == PP_PLACEHOLDER.BODY
    )
    body._element.getparent().remove(body._element)

    with pytest.raises(UnsupportedStructureError, match="no body placeholder"):
        diff_decks(prs, prs, detail="text")


# ------------------------------------------------------------------ typed refusals on bad input


def test_corrupt_input_speaks_as_typed_refusal():
    """Bad input produces typed, specific refusals from the
    organs - never raw tracebacks. (Upstream loader behavior is unchanged, additively.)"""
    from pptx.errors import UnsupportedStructureError

    corrupt = _path("self_generated/corrupt_dangling_sldid.pptx")
    with pytest.raises(UnsupportedStructureError, match="relationship graph is broken"):
        diff_decks(corrupt, corrupt)
    with pytest.raises(UnsupportedStructureError, match="relationship graph is broken"):
        Presentation(corrupt).apply_footers(slide_number=True)
    dest = Presentation(_path("self_generated/template_alpha.pptx"))
    with pytest.raises(UnsupportedStructureError, match="relationship graph is broken"):
        dest.import_slide(Presentation(corrupt), 0, mode="bake")
