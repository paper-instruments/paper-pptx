"""Contract tests: table structure operations.

insert_row / delete_row / insert_column / delete_column with grid-width bookkeeping and
cell-wise merged-region guards. The grid invariant (every `a:tr` holds exactly one `a:tc`
per `a:gridCol`) is asserted directly after every mutation, per the plan.
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.errors import (
    PaperRefusal,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches

from . import corpus
from .contract import (
    assert_changed_parts,
    assert_refusal_atomic,
    save_reopen,
    save_to_bytes,
    zip_member_map,
)
from .fragval import assert_tbl_fragment_valid
from .lo import lo_load_smoke
from .relint import dangling_relationship_targets, missing_relationship_references

MERGED = "self_generated/merged_tables.pptx"
LO_MERGED = "libreoffice_export/lo_merged_tables.pptx"
GAUNTLET = "self_generated/gauntlet.pptx"


def _open(relpath):
    return Presentation(str(corpus.fixture_path(relpath)))


def _merged_table(prs):
    return prs.slides[0].shapes.shape_by_name("merged_table").table


def _gauntlet_table(prs):
    return prs.slides[2].shapes.table_by_name("gauntlet_table")


def assert_grid_consistent(table):
    """The invariant: one a:tc per a:gridCol in every row, continuations included."""
    col_count = len(table._tbl.tblGrid.gridCol_lst)
    for tr in table._tbl.tr_lst:
        assert len(tr.tc_lst) == col_count
    assert_tbl_fragment_valid(table)


def _cell_texts(table, row_idx):
    return [table.cell(row_idx, c).text_frame.text for c in range(len(table.columns))]


def _assert_canonical_merge(table, top, left, bottom, right):
    """Assert exact python-pptx rectangular merge topology."""
    row_span = bottom - top + 1
    grid_span = right - left + 1
    for row_idx in range(top, bottom + 1):
        for col_idx in range(left, right + 1):
            tc = table.cell(row_idx, col_idx)._tc
            assert tc.rowSpan == (row_span if row_idx == top else 1)
            assert tc.gridSpan == (grid_span if col_idx == left else 1)
            assert tc.hMerge is (col_idx != left)
            assert tc.vMerge is (row_idx != top)


# ------------------------------------------------------------------------------ insert_row


def test_insert_row_after_last_survives_save_reopen():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    new_row = table.insert_row(2)
    new_row.cells[0].text_frame.paragraphs[0].add_run().text = "appended"
    assert_grid_consistent(table)

    reopened_table = _gauntlet_table(save_reopen(prs))
    assert len(reopened_table.rows) == 4
    assert reopened_table.cell(3, 0).text_frame.text == "appended"
    assert_grid_consistent(reopened_table)


def test_insert_row_at_top_with_after_minus_one():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.insert_row(-1)
    assert_grid_consistent(table)
    reopened_table = _gauntlet_table(save_reopen(prs))
    assert reopened_table.cell(0, 0).text_frame.text == ""
    assert reopened_table.cell(1, 0).text_frame.text == "r0c0"  # -- old first row shifted down


def test_insert_row_has_exact_part_budget():
    prs = _open(GAUNTLET)
    before = save_to_bytes(prs)
    _gauntlet_table(prs).insert_row(2)
    assert_changed_parts(before, save_to_bytes(prs), expect_changed=["ppt/slides/slide3.xml"])


def test_insert_row_default_height_copies_neighbor():
    """The neighbor's height must be DISTINCT from every other row's, or a wrong-source
    mutant passes; asserted through save->reopen."""
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.rows[1].height = Emu(777240)  # -- make the neighbor unmistakable
    table.insert_row(1)
    reopened_table = _gauntlet_table(save_reopen(prs))
    assert reopened_table.rows[2].height == Emu(777240)  # -- the new row, at index 2
    assert reopened_table.rows[1].height == Emu(777240)  # -- the neighbor it copied
    assert reopened_table.rows[0].height != Emu(777240)


def test_insert_column_default_width_copies_the_distinct_neighbor():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.columns[1].width = Emu(555120)  # -- distinct from columns 0 and 2
    table.insert_column(1)
    reopened_table = _gauntlet_table(save_reopen(prs))
    assert reopened_table.columns[2].width == Emu(555120)  # -- the new column
    assert reopened_table.columns[0].width != Emu(555120)


def test_insert_row_copy_format_from_copies_tcPr_but_never_merges_or_text():
    prs = _open(MERGED)
    table = _merged_table(prs)
    from pptx.dml.color import RGBColor

    for col_idx in range(4):
        cell = table.cell(1, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x11, 0x22, 0x33)
    template_height = table.rows[1].height

    new_row = table.insert_row(4, copy_format_from=1)
    assert new_row.height == template_height
    assert_grid_consistent(table)

    reopened_table = _merged_table(save_reopen(prs))
    for col_idx in range(4):
        tc = reopened_table.cell(5, col_idx)._tc
        assert tc.tcPr is not None
        fill = tc.tcPr.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill"
        )
        assert fill is not None, "cell (5, %d) did not inherit the template fill" % col_idx
        assert tc.gridSpan == 1
        assert tc.rowSpan == 1
        assert not tc.hMerge
        assert not tc.vMerge
        assert reopened_table.cell(5, col_idx).text_frame.text == ""


def test_insert_row_boundary_inside_vertical_merge_refuses_atomically():
    prs = _open(MERGED)
    raised = assert_refusal_atomic(
        prs, lambda p: _merged_table(p).insert_row(2), UnsupportedStructureError
    )
    assert "(2, 0)" in str(raised)
    assert isinstance(raised, PaperRefusal)


def test_insert_row_below_merged_header_is_allowed():
    """The cell-wise rule: a horizontally-merged header must not poison row operations
    whose boundary only touches the region's edge."""
    prs = _open(MERGED)
    table = _merged_table(prs)
    table.insert_row(0)
    assert_grid_consistent(table)
    reopened_table = _merged_table(save_reopen(prs))
    assert reopened_table.cell(0, 0)._tc.gridSpan == 4  # -- header intact above the new row
    assert reopened_table.cell(1, 0).text_frame.text == ""


# ------------------------------------------------------------------------------ delete_row


def test_delete_row_survives_save_reopen_with_exact_budget():
    prs = _open(GAUNTLET)
    before = save_to_bytes(prs)
    table = _gauntlet_table(prs)
    deleted_texts = _cell_texts(table, 1)
    table.delete_row(1)
    assert_grid_consistent(table)
    after = save_to_bytes(prs)
    assert_changed_parts(before, after, expect_changed=["ppt/slides/slide3.xml"])

    reopened_table = _gauntlet_table(Presentation(io.BytesIO(after)))
    assert len(reopened_table.rows) == 2
    assert deleted_texts[0] not in [c.text_frame.text for c in reopened_table.iter_cells()]


def test_delete_row_containing_only_horizontal_merge_is_allowed():
    """Deleting the merged header row removes the whole merge with it - allowed."""
    prs = _open(MERGED)
    table = _merged_table(prs)
    table.delete_row(0)
    assert_grid_consistent(table)
    reopened_table = _merged_table(save_reopen(prs))
    assert len(reopened_table.rows) == 4
    assert reopened_table.cell(0, 0).text_frame.text == "r1c0"
    # -- the vertical merge (previously rows 2..3) moved up intact
    assert reopened_table.cell(1, 0)._tc.rowSpan == 2


@pytest.mark.parametrize("row_idx", [2, 3])
def test_delete_row_intersecting_vertical_merge_refuses_atomically(row_idx):
    prs = _open(MERGED)
    raised = assert_refusal_atomic(
        prs, lambda p: _merged_table(p).delete_row(row_idx), UnsupportedStructureError
    )
    assert "spanning rows 2..3" in str(raised)


def test_delete_last_remaining_row_raises_valueerror_atomically():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.delete_row(2)
    table.delete_row(1)
    before = save_to_bytes(prs)
    with pytest.raises(ValueError, match="last remaining row"):
        _gauntlet_table(prs).delete_row(0)
    assert_changed_parts(before, save_to_bytes(prs))  # -- empty budget


# --------------------------------------------------------------------------- insert_column


def test_insert_column_survives_save_reopen_with_exact_budget():
    prs = _open(GAUNTLET)
    before = save_to_bytes(prs)
    table = _gauntlet_table(prs)
    new_column = table.insert_column(0, width=Inches(1))
    assert new_column.width == Inches(1)
    assert_grid_consistent(table)
    after = save_to_bytes(prs)
    assert_changed_parts(before, after, expect_changed=["ppt/slides/slide3.xml"])

    reopened_table = _gauntlet_table(Presentation(io.BytesIO(after)))
    assert len(reopened_table.columns) == 4
    assert reopened_table.cell(0, 1).text_frame.text == ""
    assert reopened_table.cell(0, 0).text_frame.text == "r0c0"
    assert reopened_table.cell(0, 2).text_frame.text == "r0c1"


def test_insert_column_default_width_copies_neighbor_and_updates_frame():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    neighbor_width = table.columns[1].width
    table.insert_column(1)
    assert table.columns[2].width == neighbor_width
    shape = save_reopen(prs).slides[2].shapes.shape_by_name("gauntlet_table")
    assert shape.width == Emu(sum(c.width for c in shape.table.columns))


@pytest.mark.parametrize(("after", "copy_format_from"), [(-1, 2), (2, 0)])
def test_insert_column_copies_preinsertion_direct_format_by_row(after, copy_format_from):
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    template_width = Emu(654321)
    table.columns[copy_format_from].width = template_width

    for row_idx in range(3):
        cell = table.cell(row_idx, copy_format_from)
        cell.margin_left = Emu(100000 + row_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x10 + row_idx, 0x20, 0x30)
    template_tcPr_xml = [
        etree.tostring(table.cell(row_idx, copy_format_from)._tc.tcPr)
        for row_idx in range(3)
    ]

    table.insert_column(after, copy_format_from=copy_format_from)
    reopened_table = _gauntlet_table(save_reopen(prs))
    inserted_idx = after + 1
    assert reopened_table.columns[inserted_idx].width == template_width
    for row_idx in range(3):
        inserted_cell = reopened_table.cell(row_idx, inserted_idx)
        assert etree.tostring(inserted_cell._tc.tcPr) == template_tcPr_xml[row_idx]
        assert inserted_cell.text == ""
        assert inserted_cell._tc.rowSpan == 1
        assert inserted_cell._tc.gridSpan == 1
        assert not inserted_cell._tc.hMerge
        assert not inserted_cell._tc.vMerge


def test_insert_column_preserves_minimal_properties_when_template_has_no_tcPr():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    template_tc = table.cell(1, 1)._tc
    assert template_tc.tcPr is not None
    template_tc.remove(template_tc.tcPr)

    table.insert_column(2, copy_format_from=1)
    reopened_table = _gauntlet_table(save_reopen(prs))
    inserted_tcPr = reopened_table.cell(1, 3)._tc.tcPr
    assert inserted_tcPr is not None
    assert len(inserted_tcPr) == 0
    assert dict(inserted_tcPr.attrib) == {}


def test_insert_column_explicit_width_wins_over_template_width():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.columns[1].width = Emu(456789)

    table.insert_column(2, width=Emu(987654), copy_format_from=1)

    reopened_table = _gauntlet_table(save_reopen(prs))
    assert reopened_table.columns[3].width == Emu(987654)


def test_insert_column_copies_opaque_tcPr_extension_verbatim():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    template_tcPr = table.cell(1, 1)._tc.get_or_add_tcPr()
    extLst = OxmlElement("a:extLst")
    ext = OxmlElement("a:ext")
    ext.set("uri", "urn:paper:test:opaque-cell-properties")
    opaque = etree.SubElement(ext, "{urn:paper:test}opaque")
    opaque.set("producer-token", "keep-exactly")
    extLst.append(ext)
    template_tcPr.append(extLst)
    expected_extension = etree.tostring(extLst)

    table.insert_column(2, copy_format_from=1)

    reopened_table = _gauntlet_table(save_reopen(prs))
    copied_extLst = reopened_table.cell(1, 3)._tc.tcPr.find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}extLst"
    )
    assert copied_extLst is not None
    assert etree.tostring(copied_extLst) == expected_extension


def test_insert_column_template_never_copies_text_or_merge_state():
    prs = _open(MERGED)
    table = _merged_table(prs)

    table.insert_column(3, copy_format_from=0)

    reopened_table = _merged_table(save_reopen(prs))
    for row_idx in range(len(reopened_table.rows)):
        cell = reopened_table.cell(row_idx, 4)
        assert cell.text == ""
        assert cell._tc.rowSpan == 1
        assert cell._tc.gridSpan == 1
        assert not cell._tc.hMerge
        assert not cell._tc.vMerge


@pytest.mark.parametrize("after", [0, 1, 2])
def test_insert_column_boundary_inside_horizontal_merge_refuses_atomically(after):
    """Every interior boundary of the header merge (cols 0..3) must refuse - including
    the left edge (after=0), where an off-by-one in the intersection test would
    silently split the merge."""
    prs = _open(MERGED)
    raised = assert_refusal_atomic(
        prs, lambda p: _merged_table(p).insert_column(after), UnsupportedStructureError
    )
    assert "columns 0..3" in str(raised)


def test_insert_column_at_right_edge_of_merge_is_allowed():
    prs = _open(MERGED)
    table = _merged_table(prs)
    table.insert_column(3)
    assert_grid_consistent(table)
    reopened_table = _merged_table(save_reopen(prs))
    assert len(reopened_table.columns) == 5
    assert reopened_table.cell(0, 0)._tc.gridSpan == 4  # -- header merge untouched


# --------------------------------------------------------------------------- delete_column


def test_delete_column_survives_save_reopen_with_exact_budget():
    prs = _open(GAUNTLET)
    before = save_to_bytes(prs)
    table = _gauntlet_table(prs)
    table.delete_column(1)
    assert_grid_consistent(table)
    after = save_to_bytes(prs)
    assert_changed_parts(before, after, expect_changed=["ppt/slides/slide3.xml"])

    reopened_table = _gauntlet_table(Presentation(io.BytesIO(after)))
    assert len(reopened_table.columns) == 2
    assert [c.text_frame.text for c in reopened_table.rows[0].cells] == ["r0c0", "r0c2"]
    shape = Presentation(io.BytesIO(after)).slides[2].shapes.shape_by_name("gauntlet_table")
    assert shape.width == Emu(sum(c.width for c in shape.table.columns))


def test_delete_column_containing_whole_vertical_merge_is_allowed():
    """A vertical merge wholly inside the deleted column goes with it - cell-wise rule."""
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.cell(0, 1).merge(table.cell(1, 1))  # -- rowSpan=2 merge contained in column 1
    table.delete_column(1)
    assert_grid_consistent(table)
    reopened_table = _gauntlet_table(save_reopen(prs))
    assert len(reopened_table.columns) == 2
    for cell in reopened_table.iter_cells():
        assert not cell._tc.is_spanned
        assert not cell._tc.is_merge_origin


@pytest.mark.parametrize("col_idx", [0, 1, 3])
def test_delete_column_intersecting_horizontal_merge_refuses_atomically(col_idx):
    prs = _open(MERGED)
    raised = assert_refusal_atomic(
        prs, lambda p: _merged_table(p).delete_column(col_idx), UnsupportedStructureError
    )
    assert "columns 0..3" in str(raised)


def test_delete_last_remaining_column_raises_valueerror_atomically():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.delete_column(2)
    table.delete_column(1)
    before = save_to_bytes(prs)
    with pytest.raises(ValueError, match="last remaining column"):
        _gauntlet_table(prs).delete_column(0)
    assert_changed_parts(before, save_to_bytes(prs))  # -- empty budget


# ----------------------------------------------------------------- arguments and invariants


@pytest.mark.parametrize("bad", [True, False, "1", 1.0, None, 99, -2])
def test_bad_indices_raise_valueerror_before_any_change(bad):
    prs = _open(MERGED)
    before = save_to_bytes(prs)
    table = _merged_table(prs)
    operations = [
        lambda: table.insert_row(bad),
        lambda: table.delete_row(bad),
        lambda: table.insert_column(bad),
        lambda: table.delete_column(bad),
    ]
    if bad is not None:  # -- None is copy_format_from's legitimate default
        operations.append(lambda: table.insert_row(0, copy_format_from=bad))
        operations.append(lambda: table.insert_column(3, copy_format_from=bad))
    for operation in operations:
        with pytest.raises(ValueError):
            operation()
    assert_changed_parts(before, save_to_bytes(prs))  # -- empty budget


def test_insert_column_rejects_bad_width():
    prs = _open(MERGED)
    table = _merged_table(prs)
    for bad in (True, 0, -914400, 1.5, "wide"):
        with pytest.raises(ValueError):
            table.insert_column(3, width=bad)


def test_insert_then_delete_row_is_a_complete_noop():
    prs = _open(MERGED)
    before = save_to_bytes(prs)
    table = _merged_table(prs)
    table.insert_row(4)
    table.delete_row(5)
    assert_changed_parts(before, save_to_bytes(prs))  # -- empty budget


def test_insert_then_delete_column_is_a_complete_noop():
    prs = _open(MERGED)
    before = save_to_bytes(prs)
    table = _merged_table(prs)
    table.insert_column(3)
    table.delete_column(4)
    assert_changed_parts(before, save_to_bytes(prs))  # -- empty budget


def test_operations_on_libreoffice_authored_merged_table():
    """Producer diversity: the same guards and surgery on LibreOffice-written bytes."""
    prs = Presentation(str(corpus.fixture_path(LO_MERGED)))
    table = prs.slides[0].shapes.shape_by_name("merged_table").table
    raised = assert_refusal_atomic(
        prs,
        lambda p: p.slides[0].shapes.shape_by_name("merged_table").table.delete_row(2),
        UnsupportedStructureError,
    )
    assert "spanning rows 2..3" in str(raised)

    table.insert_row(4)
    table.insert_column(3)
    assert_grid_consistent(table)
    reopened = save_reopen(prs)
    reopened_table = reopened.slides[0].shapes.shape_by_name("merged_table").table
    assert len(reopened_table.rows) == 6
    assert len(reopened_table.columns) == 5
    assert reopened_table.cell(0, 0)._tc.gridSpan == 4


# --------------------------------------------------------------------------- extend_merge


@pytest.mark.parametrize(
    ("initial_corner", "requested_corner"),
    [((1, 0), (1, 1)), ((0, 1), (1, 1)), ((1, 1), (2, 2))],
)
def test_extend_merge_grows_right_down_or_both_and_survives_reopen(
    initial_corner, requested_corner
):
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(*initial_corner))

    origin.extend_merge(table.cell(*requested_corner))

    reopened_table = _gauntlet_table(save_reopen(prs))
    _assert_canonical_merge(reopened_table, 0, 0, *requested_corner)


def test_extend_merge_moves_only_new_cells_in_row_major_order_and_preserves_origin():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    frame = prs.slides[2].shapes.shape_by_name("gauntlet_table")
    frame_geometry = (frame.left, frame.top, frame.width, frame.height)
    table_style_xml = etree.tostring(table._tbl.tblPr)
    origin = table.cell(0, 0)
    origin.fill.solid()
    origin.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)
    origin.merge(table.cell(1, 1))
    origin_tc = origin._tc
    origin_tcPr_xml = etree.tostring(origin_tc.tcPr)
    expected_text = "\n".join(
        ["r0c0", "r0c1", "r1c0", "r1c1", "r0c2", "r1c2", "r2c0", "r2c1", "r2c2"]
    )
    before = save_to_bytes(prs)

    target = table.cell(2, 2)
    origin.extend_merge(target)

    assert origin._tc is origin_tc
    assert etree.tostring(origin._tc.tcPr) == origin_tcPr_xml
    assert target._tc.getparent() is not None
    assert (frame.left, frame.top, frame.width, frame.height) == frame_geometry
    after = save_to_bytes(prs)
    assert_changed_parts(before, after, expect_changed=["ppt/slides/slide3.xml"])
    zip_map = zip_member_map(after)
    assert dangling_relationship_targets(zip_map) == []
    assert missing_relationship_references(zip_map) == []

    reopened = Presentation(io.BytesIO(after))
    reopened_frame = reopened.slides[2].shapes.shape_by_name("gauntlet_table")
    reopened_table = reopened_frame.table
    assert reopened_table.cell(0, 0).text == expected_text
    assert etree.tostring(reopened_table._tbl.tblPr) == table_style_xml
    reopened_geometry = (
        reopened_frame.left,
        reopened_frame.top,
        reopened_frame.width,
        reopened_frame.height,
    )
    assert reopened_geometry == frame_geometry
    _assert_canonical_merge(reopened_table, 0, 0, 2, 2)


def test_extend_merge_preserves_an_unrelated_merge():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(0, 1))
    unrelated = table.cell(1, 2)
    unrelated.merge(table.cell(2, 2))

    origin.extend_merge(table.cell(1, 1))

    reopened_table = _gauntlet_table(save_reopen(prs))
    _assert_canonical_merge(reopened_table, 0, 0, 1, 1)
    _assert_canonical_merge(reopened_table, 1, 2, 2, 2)


def test_extend_merge_clears_explicit_default_attributes_from_canonical_topology():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(0, 1))
    for col_idx in (0, 1):
        tc = table.cell(1, col_idx)._tc
        tc.set("rowSpan", "1")
        tc.set("gridSpan", "1")
        tc.set("hMerge", "false")
        tc.set("vMerge", "false")

    origin.extend_merge(table.cell(1, 1))

    reopened_table = _gauntlet_table(save_reopen(prs))
    for row_idx in range(2):
        for col_idx in range(2):
            tc = reopened_table.cell(row_idx, col_idx)._tc
            if row_idx != 0:
                assert "rowSpan" not in tc.attrib
            if col_idx != 0:
                assert "gridSpan" not in tc.attrib
            if col_idx == 0:
                assert "hMerge" not in tc.attrib
            if row_idx == 0:
                assert "vMerge" not in tc.attrib
    _assert_canonical_merge(reopened_table, 0, 0, 1, 1)


def test_extend_merge_at_current_corner_is_an_exact_noop():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(1, 1))
    before = save_to_bytes(prs)

    origin.extend_merge(table.cell(1, 1))

    assert_changed_parts(before, save_to_bytes(prs))


def test_extend_merge_rejects_wrong_type_nonorigin_and_other_table_atomically():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(0, 1))

    for operation in (
        lambda: origin.extend_merge("not a cell"),
        lambda: table.cell(2, 0).extend_merge(table.cell(2, 1)),
    ):
        before = save_to_bytes(prs)
        with pytest.raises(ValueError):
            operation()
        assert_changed_parts(before, save_to_bytes(prs))

    other_shape = prs.slides[2].shapes.add_table(
        2, 2, Inches(0), Inches(0), Inches(1), Inches(1)
    )
    before = save_to_bytes(prs)
    with pytest.raises(ValueError, match="different table"):
        origin.extend_merge(other_shape.table.cell(1, 1))
    assert_changed_parts(before, save_to_bytes(prs))


@pytest.mark.parametrize("requested_corner", [(0, 2), (2, 0), (0, 0)])
def test_extend_merge_rejects_shrink_or_up_left_request_atomically(requested_corner):
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(1, 1))
    before = save_to_bytes(prs)

    with pytest.raises(ValueError, match="current lower-right"):
        origin.extend_merge(table.cell(*requested_corner))

    assert_changed_parts(before, save_to_bytes(prs))


def test_extend_merge_rejects_stale_origin_before_stale_target():
    prs = _open(GAUNTLET)
    slide = prs.slides[2]
    shape = slide.shapes.shape_by_name("gauntlet_table")
    table = shape.table
    origin = table.cell(0, 0)
    origin.merge(table.cell(0, 1))
    target = table.cell(1, 1)
    slide.shapes.delete(shape)

    with pytest.raises(TargetNotFoundError, match="merge origin is stale"):
        origin.extend_merge(target)


def test_extend_merge_rejects_a_stale_target_before_table_ownership():
    prs = _open(GAUNTLET)
    slide = prs.slides[2]
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(0, 1))
    target_shape = slide.shapes.add_table(2, 2, Inches(0), Inches(0), Inches(1), Inches(1))
    target = target_shape.table.cell(1, 1)
    slide.shapes.delete(target_shape)

    with pytest.raises(TargetNotFoundError, match="merge target is stale"):
        origin.extend_merge(target)


def test_extend_merge_rejects_an_intersecting_merge_but_not_an_unrelated_one():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(0, 1))
    conflict = table.cell(1, 1)
    conflict.merge(table.cell(1, 2))

    raised = assert_refusal_atomic(
        prs,
        lambda p: _gauntlet_table(p).cell(0, 0).extend_merge(
            _gauntlet_table(p).cell(1, 2)
        ),
        UnsupportedStructureError,
    )
    assert "spanning rows 1..1, columns 1..2" in str(raised)


@pytest.mark.parametrize(
    ("corruption", "expected"),
    [
        (lambda table: setattr(table.cell(0, 1)._tc, "hMerge", False), "non-canonical"),
        (lambda table: setattr(table.cell(1, 0)._tc, "hMerge", True), "non-canonical"),
        (lambda table: setattr(table.cell(0, 1)._tc, "rowSpan", 3), "non-canonical"),
        (lambda table: setattr(table.cell(0, 0)._tc, "rowSpan", 99), "out-of-bounds"),
    ],
)
def test_extend_merge_refuses_malformed_current_topology_atomically(corruption, expected):
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.cell(0, 0).merge(table.cell(1, 1))
    corruption(table)
    before = save_to_bytes(prs)

    with pytest.raises(UnsupportedStructureError, match=expected):
        table.cell(0, 0).extend_merge(table.cell(2, 2))

    assert_changed_parts(before, save_to_bytes(prs))


def test_extend_merge_refuses_orphan_merge_state_in_new_area_atomically():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(0, 1))
    table.cell(1, 1)._tc.hMerge = True
    before = save_to_bytes(prs)

    with pytest.raises(UnsupportedStructureError, match=r"cell \(1, 1\) carrying merge state"):
        origin.extend_merge(table.cell(1, 1))

    assert_changed_parts(before, save_to_bytes(prs))


@pytest.mark.parametrize("failure_stage", ["after-content", "after-topology"])
def test_extend_merge_rolls_back_late_failure_and_retained_proxies(monkeypatch, failure_stage):
    import pptx.table as table_module

    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    origin = table.cell(0, 0)
    origin.merge(table.cell(1, 1))
    target = table.cell(2, 2)
    origin_tc = origin._tc
    target_tc = target._tc
    before = save_to_bytes(prs)
    apply_topology = table_module._apply_merge_topology

    def fail_topology(tc_range):
        if failure_stage == "after-topology":
            apply_topology(tc_range)
        raise RuntimeError("forced late merge-extension failure")

    monkeypatch.setattr(table_module, "_apply_merge_topology", fail_topology)
    with pytest.raises(RuntimeError, match="forced late merge-extension failure"):
        origin.extend_merge(target)

    assert_changed_parts(before, save_to_bytes(prs))
    assert origin._tc is origin_tc
    assert target._tc is target_tc
    assert origin.text == "r0c0\nr0c1\nr1c0\nr1c1"
    _assert_canonical_merge(table, 0, 0, 1, 1)


# ------------------------------------------------------- anchored writes reach table cells


def test_anchored_replace_reaches_table_cells_including_merge_origin():
    """Cell text edits route through the anchored-write path -
    prove that path reaches table cells, merged origins included."""
    from pptx.edit import replace_text, replace_text_at
    from pptx.inspect import inspect_text

    prs = _open(MERGED)
    result = replace_text(prs, "Merged header", "Merged HEADING")
    assert result.replacements == 1

    reopened = save_reopen(prs)
    reopened_table = _merged_table(reopened)
    assert reopened_table.cell(0, 0).text_frame.text == "Merged HEADING"

    # -- and the single-block anchored path, via an inspect_text table-cell anchor
    blocks = inspect_text(reopened.slides[0])
    cell_block = next(
        b for b in blocks.blocks if b.container == "table-cell" and b.text == "r1c2"
    )
    assert cell_block.anchor.locator == {
        "kind": "table-cell",
        "shape_id": cell_block.shape_id,
        "row": 1,
        "column": 2,
        "paragraph_index": 0,
    }
    result = replace_text_at(reopened, cell_block.anchor, "r1c2", "R1C2 EDITED")
    assert result.replacements == 1
    final = save_reopen(reopened)
    assert _merged_table(final).cell(1, 2).text_frame.text == "R1C2 EDITED"


@pytest.mark.parametrize(
    "locator_change",
    [
        {"row": 99},
        {"column": 99},
        {"paragraph_index": 99},
        {"kind": "shape"},
    ],
)
def test_table_anchor_refuses_changed_coordinates_or_container_atomically(locator_change):
    from pptx.edit import replace_text_at
    from pptx.errors import PaperRefusal
    from pptx.inspect import BlockAnchor, inspect_text

    from .contract import assert_refusal_atomic

    prs = _open(MERGED)
    block = next(
        b for b in inspect_text(prs.slides[0]).blocks
        if b.container == "table-cell" and b.text == "r1c2"
    )
    locator = dict(block.anchor.locator)
    locator.update(locator_change)
    anchor = BlockAnchor(
        block.anchor.part,
        block.anchor.block_index,
        block.anchor.content_hash,
        block.anchor.version,
        locator,
    )

    assert_refusal_atomic(
        prs,
        lambda p: replace_text_at(p, anchor, "r1c2", "changed"),
        PaperRefusal,
    )


def test_table_anchor_uses_paragraph_index_within_its_exact_cell():
    from pptx.edit import replace_text_at
    from pptx.inspect import inspect_text

    prs = _open(MERGED)
    cell = _merged_table(prs).cell(1, 2)
    cell.text_frame.add_paragraph().text = "second paragraph"
    block = next(
        b for b in inspect_text(prs.slides[0]).blocks if b.text == "second paragraph"
    )
    assert block.anchor.locator["paragraph_index"] == 1

    replace_text_at(prs, block.anchor, "second", "SECOND")

    final = save_reopen(prs)
    assert _merged_table(final).cell(1, 2).text_frame.paragraphs[1].text == "SECOND paragraph"


def test_column_surgery_refuses_a_ragged_table_before_mutation():
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    malformed_row = table._tbl.tr_lst[1]
    malformed_row.remove(malformed_row.tc_lst[-1])
    before = table._tbl.xml

    with pytest.raises(UnsupportedStructureError, match="cells for"):
        table.delete_column(0)

    assert table._tbl.xml == before


@pytest.mark.parametrize(
    ("operation", "notification"),
    [
        (lambda table: table.insert_row(2), "notify_height_changed"),
        (lambda table: table.delete_row(1), "notify_height_changed"),
        (lambda table: table.insert_column(2), "notify_width_changed"),
        (
            lambda table: table.insert_column(2, copy_format_from=1),
            "notify_width_changed",
        ),
        (lambda table: table.delete_column(1), "notify_width_changed"),
    ],
)
def test_table_surgery_rolls_back_late_extent_failure(monkeypatch, operation, notification):
    from pptx.table import Table

    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    rows = table.rows
    columns = table.columns
    before = save_to_bytes(prs)

    def fail_extent_update(self):
        raise RuntimeError("forced late table failure")

    monkeypatch.setattr(Table, notification, fail_extent_update)
    with pytest.raises(RuntimeError, match="forced late table failure"):
        operation(table)

    assert_changed_parts(before, save_to_bytes(prs))
    assert table.rows is rows
    assert table.columns is columns


def test_table_surgery_refuses_a_table_on_a_deleted_shape():
    from pptx.errors import TargetNotFoundError

    prs = _open(GAUNTLET)
    slide = prs.slides[2]
    shape = slide.shapes.shape_by_name("gauntlet_table")
    table = shape.table
    slide.shapes.delete(shape)

    with pytest.raises(TargetNotFoundError, match="table shape is stale"):
        table.insert_row(0)


# --------------------------------------------------------------------------------- lo_smoke


@pytest.mark.lo_smoke
def test_surgered_table_loads_in_libreoffice(tmp_path):
    prs = _open(GAUNTLET)
    table = _gauntlet_table(prs)
    table.insert_column(2, copy_format_from=1)
    origin = table.cell(0, 0)
    origin.merge(table.cell(1, 1))
    origin.extend_merge(table.cell(2, 2))
    assert_grid_consistent(table)
    out = tmp_path / "surgered.pptx"
    prs.save(str(out))
    lo_load_smoke(out, tmp_path)
