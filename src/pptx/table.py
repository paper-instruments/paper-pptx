"""Table-related objects such as Table and Cell."""

from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING, Iterator

from pptx.dml.fill import FillFormat
from pptx.oxml.table import TcRange
from pptx.shapes import Subshape
from pptx.text.text import TextFrame
from pptx.util import Emu, lazyproperty

if TYPE_CHECKING:
    from pptx.enum.text import MSO_VERTICAL_ANCHOR
    from pptx.oxml.table import (
        CT_Table,
        CT_TableCell,
        CT_TableCellProperties,
        CT_TableCol,
        CT_TableRow,
    )
    from pptx.parts.slide import BaseSlidePart
    from pptx.shapes.graphfrm import GraphicFrame
    from pptx.types import ProvidesPart
    from pptx.util import Length


def _replace_cell_properties(
    tc: CT_TableCell, tcPr_snapshot: CT_TableCellProperties | None
) -> None:
    """Replace `tc` direct properties with a detached snapshot when one exists."""
    if tcPr_snapshot is None:
        return
    existing_tcPr = tc.tcPr
    if existing_tcPr is not None:
        tc.remove(existing_tcPr)
    tc.append(tcPr_snapshot)


def _apply_merge_topology(tc_range: TcRange) -> None:
    """Write canonical rectangular merge attributes for `tc_range`."""
    row_count, col_count = tc_range.dimensions
    for tc in tc_range.iter_tcs():
        tc.rowSpan = 1
        tc.gridSpan = 1
        tc.hMerge = False
        tc.vMerge = False
    for tc in tc_range.iter_top_row_tcs():
        tc.rowSpan = row_count
    for tc in tc_range.iter_left_col_tcs():
        tc.gridSpan = col_count
    for tc in tc_range.iter_except_left_col_tcs():
        tc.hMerge = True
    for tc in tc_range.iter_except_top_row_tcs():
        tc.vMerge = True


class Table(object):
    """A DrawingML table object.

    Not intended to be constructed directly, use
    :meth:`.Slide.shapes.add_table` to add a table to a slide.
    """

    def __init__(self, tbl: CT_Table, graphic_frame: GraphicFrame):
        super(Table, self).__init__()
        self._tbl = tbl
        self._graphic_frame = graphic_frame

    def cell(self, row_idx: int, col_idx: int) -> _Cell:
        """Return cell at `row_idx`, `col_idx`.

        Return value is an instance of |_Cell|. `row_idx` and `col_idx` are zero-based, e.g.
        cell(0, 0) is the top, left cell in the table.
        """
        return _Cell(self._tbl.tc(row_idx, col_idx), self)

    @lazyproperty
    def columns(self) -> _ColumnCollection:
        """|_ColumnCollection| instance for this table.

        Provides access to |_Column| objects representing the table's columns. |_Column| objects
        are accessed using list notation, e.g. `col = tbl.columns[0]`.
        """
        return _ColumnCollection(self._tbl, self)

    def delete_column(self, col_idx: int) -> None:
        """Remove the column at 0-based `col_idx`, cells included.

        paper-pptx addition. The `a:gridCol` and every row's cell at that
        grid position are removed together, so each row keeps exactly one `a:tc` per grid
        column, and the graphic frame's width is recalculated from the remaining columns.

        Merged-cell guard is cell-wise: the operation refuses
        (|UnsupportedStructureError|, tree untouched) only when a *horizontal* merge extends
        beyond the deleted column - a vertical merge lying wholly inside the column is
        deleted with it. Deleting the last remaining column raises |ValueError| (delete the
        table's shape instead).
        """
        self._validate_structure()
        col_count = len(self._tbl.tblGrid.gridCol_lst)
        self._validate_grid_index(col_idx, "col_idx", col_count)
        if col_count == 1:
            raise ValueError(
                "cannot delete the last remaining column; delete the table shape instead"
            )
        conflicts = [
            region
            for region in self._merge_regions()
            if region.gridSpan > 1 and region.left <= col_idx <= region.right
        ]
        if conflicts:
            self._refuse_merge_conflict("deleting column %d" % col_idx, conflicts)

        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            self._tbl.tblGrid.remove(self._tbl.tblGrid.gridCol_lst[col_idx])
            for tr in self._tbl.tr_lst:
                tr.remove(tr.tc_lst[col_idx])
            self.notify_width_changed()

    def delete_row(self, row_idx: int) -> None:
        """Remove the row at 0-based `row_idx`, cells included.

        paper-pptx addition. The graphic frame's height is recalculated
        from the remaining rows.

        Merged-cell guard is cell-wise: the operation refuses
        (|UnsupportedStructureError|, tree untouched) only when a *vertical* merge extends
        beyond the deleted row - a horizontal merge lying wholly inside the row (e.g. a
        merged header) is deleted with it and must not poison other rows' operations.
        Deleting the last remaining row raises |ValueError|.
        """
        self._validate_structure()
        row_count = len(self._tbl.tr_lst)
        self._validate_grid_index(row_idx, "row_idx", row_count)
        if row_count == 1:
            raise ValueError(
                "cannot delete the last remaining row; delete the table shape instead"
            )
        conflicts = [
            region
            for region in self._merge_regions()
            if region.rowSpan > 1 and region.top <= row_idx <= region.bottom
        ]
        if conflicts:
            self._refuse_merge_conflict("deleting row %d" % row_idx, conflicts)

        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            self._tbl.remove(self._tbl.tr_lst[row_idx])
            self.notify_height_changed()

    @property
    def first_col(self) -> bool:
        """When `True`, indicates first column should have distinct formatting.

        Read/write. Distinct formatting is used, for example, when the first column contains row
        headings (is a side-heading column).
        """
        return self._tbl.firstCol

    @first_col.setter
    def first_col(self, value: bool):
        self._tbl.firstCol = value

    @property
    def first_row(self) -> bool:
        """When `True`, indicates first row should have distinct formatting.

        Read/write. Distinct formatting is used, for example, when the first row contains column
        headings.
        """
        return self._tbl.firstRow

    @first_row.setter
    def first_row(self, value: bool):
        self._tbl.firstRow = value

    @property
    def horz_banding(self) -> bool:
        """When `True`, indicates rows should have alternating shading.

        Read/write. Used to allow rows to be traversed more easily without losing track of which
        row is being read.
        """
        return self._tbl.bandRow

    @horz_banding.setter
    def horz_banding(self, value: bool):
        self._tbl.bandRow = value

    def insert_column(
        self,
        after: int,
        *,
        width: Length | None = None,
        copy_format_from: int | None = None,
    ) -> _Column:
        """Insert a new empty column immediately after 0-based column `after`; return it.

        paper-pptx addition. `after=-1` inserts before the first column.
        `copy_format_from` is the zero-based index of a column in the table before insertion.
        When provided, each new cell receives a deep copy of that row's template-cell `a:tcPr`
        (direct formatting only); text and merge state are never copied. `width` (EMU int) wins
        when provided, otherwise the template column supplies the width. With no template, width
        defaults to the neighboring column at `after` (the first column when `after=-1`). A new
        minimal empty cell is inserted at the same grid position in every row - the grid stays
        consistent by construction - and the graphic frame's width is recalculated.

        Merged-cell guard is cell-wise: refuses (|UnsupportedStructureError|, tree
        untouched) only when the insertion boundary would split a horizontal merge; vertical
        merges elsewhere in the table never block the operation.
        """
        self._validate_structure()
        col_count = len(self._tbl.tblGrid.gridCol_lst)
        self._validate_grid_index(after, "after", col_count, lower=-1)
        if copy_format_from is not None:
            self._validate_grid_index(copy_format_from, "copy_format_from", col_count)
        if width is not None and (
            isinstance(width, bool) or not isinstance(width, int) or width <= 0
        ):
            raise ValueError("width must be a positive int (EMU), got %r" % (width,))
        # -- a boundary at either table edge can never split a merge --
        if -1 < after < col_count - 1:
            conflicts = [
                region
                for region in self._merge_regions()
                if region.left <= after < region.right
            ]
            if conflicts:
                self._refuse_merge_conflict(
                    "inserting a column between columns %d and %d" % (after, after + 1),
                    conflicts,
                )

        template_tcPrs = (
            tuple(
                deepcopy(tr.tc_lst[copy_format_from].tcPr)
                if tr.tc_lst[copy_format_from].tcPr is not None
                else None
                for tr in self._tbl.tr_lst
            )
            if copy_format_from is not None
            else None
        )
        default_width_idx = (
            copy_format_from
            if copy_format_from is not None
            else (0 if after == -1 else after)
        )
        new_width = Emu(width) if width is not None else Emu(
            self._tbl.tblGrid.gridCol_lst[default_width_idx].w
        )
        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            gridCol = self._tbl.tblGrid.insert_gridCol_at(after + 1, new_width)
            for row_idx, tr in enumerate(self._tbl.tr_lst):
                tc = tr.insert_tc_at(after + 1)
                if template_tcPrs is not None:
                    _replace_cell_properties(tc, template_tcPrs[row_idx])
            self.notify_width_changed()
            column = _Column(gridCol, self.columns)
        return column

    def insert_row(self, after: int, *, copy_format_from: int | None = None) -> _Row:
        """Insert a new empty row immediately after 0-based row `after`; return it.

        paper-pptx addition. `after=-1` inserts before the first row. The
        new row holds one minimal empty cell per grid column. Row height and per-cell
        formatting (each cell's `a:tcPr`: fill, margins, anchor) are copied from the row at
        `copy_format_from` when given, otherwise the height of the neighboring row at
        `after` (the first row when `after=-1`) is used and cells carry default formatting.
        Merge attributes are never copied - the new row is always unmerged. Text is never
        copied. The graphic frame's height is recalculated.

        Merged-cell guard is cell-wise: refuses (|UnsupportedStructureError|, tree
        untouched) only when the insertion boundary would split a vertical merge; a merged
        header row never blocks body-row insertion.
        """
        self._validate_structure()
        row_count = len(self._tbl.tr_lst)
        self._validate_grid_index(after, "after", row_count, lower=-1)
        if copy_format_from is not None:
            self._validate_grid_index(copy_format_from, "copy_format_from", row_count)
        # -- a boundary at either table edge can never split a merge --
        if -1 < after < row_count - 1:
            conflicts = [
                region
                for region in self._merge_regions()
                if region.top <= after < region.bottom
            ]
            if conflicts:
                self._refuse_merge_conflict(
                    "inserting a row between rows %d and %d" % (after, after + 1), conflicts
                )

        template_idx = copy_format_from if copy_format_from is not None else (
            0 if after == -1 else after
        )
        template_tr = self._tbl.tr_lst[template_idx]
        template_tcPrs = (
            tuple(
                deepcopy(tc.tcPr) if tc.tcPr is not None else None
                for tc in template_tr.tc_lst
            )
            if copy_format_from is not None
            else None
        )
        from pptx._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            tr = self._tbl.insert_tr_at(after + 1, Emu(template_tr.h))
            for col_idx in range(len(self._tbl.tblGrid.gridCol_lst)):
                tc = tr.add_tc()
                if template_tcPrs is not None:
                    _replace_cell_properties(tc, template_tcPrs[col_idx])
            self.notify_height_changed()
            row = _Row(tr, self.rows)
        return row

    def iter_cells(self) -> Iterator[_Cell]:
        """Generate _Cell object for each cell in this table.

        Each grid cell is generated in left-to-right, top-to-bottom order.
        """
        return (_Cell(tc, self) for tc in self._tbl.iter_tcs())

    @property
    def last_col(self) -> bool:
        """When `True`, indicates the rightmost column should have distinct formatting.

        Read/write. Used, for example, when a row totals column appears at the far right of the
        table.
        """
        return self._tbl.lastCol

    @last_col.setter
    def last_col(self, value: bool):
        self._tbl.lastCol = value

    @property
    def last_row(self) -> bool:
        """When `True`, indicates the bottom row should have distinct formatting.

        Read/write. Used, for example, when a totals row appears as the bottom row.
        """
        return self._tbl.lastRow

    @last_row.setter
    def last_row(self, value: bool):
        self._tbl.lastRow = value

    def notify_height_changed(self) -> None:
        """Called by a row when its height changes.

        Triggers the graphic frame to recalculate its total height (as the sum of the row
        heights).
        """
        new_table_height = Emu(sum([row.height for row in self.rows]))
        self._graphic_frame.height = new_table_height

    def notify_width_changed(self) -> None:
        """Called by a column when its width changes.

        Triggers the graphic frame to recalculate its total width (as the sum of the column
        widths).
        """
        new_table_width = Emu(sum([col.width for col in self.columns]))
        self._graphic_frame.width = new_table_width

    @property
    def part(self) -> BaseSlidePart:
        """The package part containing this table."""
        return self._graphic_frame.part

    @lazyproperty
    def rows(self):
        """|_RowCollection| instance for this table.

        Provides access to |_Row| objects representing the table's rows. |_Row| objects are
        accessed using list notation, e.g. `col = tbl.rows[0]`.
        """
        return _RowCollection(self._tbl, self)

    @property
    def vert_banding(self) -> bool:
        """When `True`, indicates columns should have alternating shading.

        Read/write. Used to allow columns to be traversed more easily without losing track of
        which column is being read.
        """
        return self._tbl.bandCol

    @vert_banding.setter
    def vert_banding(self, value: bool):
        self._tbl.bandCol = value

    def _merge_regions(self) -> list[_MergeRegion]:
        """All merged regions in this table, one entry per merge-origin cell.

        paper-pptx helper for the row/column surgery guards. Read-only scan.
        """
        regions = []
        for row_idx, tr in enumerate(self._tbl.tr_lst):
            for col_idx, tc in enumerate(tr.tc_lst):
                if tc.is_merge_origin:
                    regions.append(_MergeRegion(row_idx, col_idx, tc.rowSpan, tc.gridSpan))
        return regions

    def _validate_structure(self) -> None:
        """Refuse malformed table grids before a surgery operation mutates XML."""
        from pptx._ownership import require_shape_attached
        from pptx.errors import UnsupportedStructureError

        require_shape_attached(self._graphic_frame, argument="table shape")
        col_count = len(self._tbl.tblGrid.gridCol_lst)
        row_count = len(self._tbl.tr_lst)
        if col_count == 0 or row_count == 0:
            raise UnsupportedStructureError(
                "table has no rows or columns; repair the table before editing it"
            )
        for row_idx, tr in enumerate(self._tbl.tr_lst):
            if len(tr.tc_lst) != col_count:
                raise UnsupportedStructureError(
                    "table row %d has %d cells for %d grid columns; repair the table before "
                    "editing it" % (row_idx, len(tr.tc_lst), col_count)
                )
        for region in self._merge_regions():
            if (
                region.rowSpan < 1
                or region.gridSpan < 1
                or region.bottom >= row_count
                or region.right >= col_count
            ):
                raise UnsupportedStructureError(
                    "table contains an out-of-bounds merged region %s; repair the table before "
                    "editing it" % region
                )

    def _refuse_merge_conflict(self, operation: str, conflicts: list[_MergeRegion]) -> None:
        """Raise |UnsupportedStructureError| naming every merged region in `conflicts`."""
        from pptx.errors import UnsupportedStructureError

        raise UnsupportedStructureError(
            "%s would cut through merged cell%s %s; split the merge first (cell.split()) or "
            "operate outside the merged region"
            % (
                operation,
                "" if len(conflicts) == 1 else "s",
                ", ".join(str(region) for region in conflicts),
            )
        )

    @staticmethod
    def _validate_grid_index(value: int, name: str, count: int, *, lower: int = 0) -> None:
        """Raise |ValueError| unless `value` is an int in `range(lower, count)`."""
        if isinstance(value, bool) or not isinstance(value, int) or not (
            lower <= value < count
        ):
            raise ValueError(
                "%s must be an int in range %d..%d, got %r" % (name, lower, count - 1, value)
            )


class _MergeRegion(object):
    """Extent of one merged-cell region: origin (top, left) plus row/grid spans.

    paper-pptx addition, used by the surgery guards and their messages.
    """

    def __init__(self, top: int, left: int, rowSpan: int, gridSpan: int):
        """A merged region, given its origin `(top, left)` and its row and column spans."""
        self.top = top
        self.left = left
        self.rowSpan = rowSpan
        self.gridSpan = gridSpan

    @property
    def bottom(self) -> int:
        """0-based index of the last row the region covers."""
        return self.top + self.rowSpan - 1

    @property
    def right(self) -> int:
        """0-based index of the last grid column the region covers."""
        return self.left + self.gridSpan - 1

    def __repr__(self) -> str:
        """Origin and span bounds, for debugging."""
        return "<_MergeRegion origin=(%d, %d) rows %d..%d cols %d..%d>" % (
            self.top, self.left, self.top, self.bottom, self.left, self.right,
        )

    def __str__(self) -> str:
        """Origin and span bounds in prose, as the surgery refusal messages quote them."""
        return "(%d, %d) spanning rows %d..%d, columns %d..%d" % (
            self.top, self.left, self.top, self.bottom, self.left, self.right,
        )

    def intersects(self, other: _MergeRegion) -> bool:
        """True when this region and `other` share at least one grid cell."""
        return not (
            self.bottom < other.top
            or other.bottom < self.top
            or self.right < other.left
            or other.right < self.left
        )


def _validate_merge_topology(tbl: CT_Table, region: _MergeRegion) -> None:
    """Refuse unless `region` has the exact continuation topology for its spans."""
    from pptx.errors import UnsupportedStructureError

    row_count = len(tbl.tr_lst)
    col_count = len(tbl.tblGrid.gridCol_lst)
    if (
        region.top < 0
        or region.left < 0
        or region.rowSpan < 1
        or region.gridSpan < 1
        or region.bottom >= row_count
        or region.right >= col_count
    ):
        raise UnsupportedStructureError(
            "merge origin declares an out-of-bounds region %s; repair the table before "
            "extending the merge" % region
        )

    for row_idx in range(region.top, region.bottom + 1):
        tr = tbl.tr_lst[row_idx]
        if len(tr.tc_lst) <= region.right:
            raise UnsupportedStructureError(
                "merge region %s crosses ragged row %d; repair the table before extending "
                "the merge" % (region, row_idx)
            )
        for col_idx in range(region.left, region.right + 1):
            tc = tr.tc_lst[col_idx]
            expected_row_span = region.rowSpan if row_idx == region.top else 1
            expected_grid_span = region.gridSpan if col_idx == region.left else 1
            expected_h_merge = col_idx != region.left
            expected_v_merge = row_idx != region.top
            if (
                tc.rowSpan != expected_row_span
                or tc.gridSpan != expected_grid_span
                or tc.hMerge is not expected_h_merge
                or tc.vMerge is not expected_v_merge
            ):
                raise UnsupportedStructureError(
                    "merge region %s has non-canonical continuation state at cell (%d, %d); "
                    "repair or split the merge before extending it"
                    % (region, row_idx, col_idx)
                )


def _validate_requested_merge_extension(
    tbl: CT_Table,
    origin_tc: CT_TableCell,
    current_region: _MergeRegion,
    requested_region: _MergeRegion,
) -> None:
    """Refuse unless newly absorbed cells are rectangular, live-grid, and unmerged."""
    from pptx.errors import UnsupportedStructureError

    row_count = len(tbl.tr_lst)
    col_count = len(tbl.tblGrid.gridCol_lst)
    if requested_region.bottom >= row_count or requested_region.right >= col_count:
        raise UnsupportedStructureError(
            "requested merge region %s is outside the table grid" % requested_region
        )
    for row_idx in range(requested_region.top, requested_region.bottom + 1):
        if len(tbl.tr_lst[row_idx].tc_lst) <= requested_region.right:
            raise UnsupportedStructureError(
                "requested merge region %s crosses ragged row %d; repair the table before "
                "extending the merge" % (requested_region, row_idx)
            )

    conflicts = []
    for row_idx, tr in enumerate(tbl.tr_lst):
        for col_idx, tc in enumerate(tr.tc_lst):
            if tc is origin_tc or not tc.is_merge_origin:
                continue
            region = _MergeRegion(row_idx, col_idx, tc.rowSpan, tc.gridSpan)
            if region.intersects(requested_region):
                conflicts.append(region)
    if conflicts:
        raise UnsupportedStructureError(
            "extending merge %s to %s would overlap merged cell%s %s"
            % (
                current_region,
                requested_region,
                "" if len(conflicts) == 1 else "s",
                ", ".join(str(region) for region in conflicts),
            )
        )

    for row_idx in range(requested_region.top, requested_region.bottom + 1):
        for col_idx in range(requested_region.left, requested_region.right + 1):
            if row_idx <= current_region.bottom and col_idx <= current_region.right:
                continue
            tc = tbl.tr_lst[row_idx].tc_lst[col_idx]
            if tc.rowSpan != 1 or tc.gridSpan != 1 or tc.hMerge or tc.vMerge:
                raise UnsupportedStructureError(
                    "extending merge %s to %s would absorb cell (%d, %d) carrying merge "
                    "state; split or repair that merge first"
                    % (current_region, requested_region, row_idx, col_idx)
                )


class _Cell(Subshape):
    """Table cell"""

    def __init__(self, tc: CT_TableCell, parent: ProvidesPart):
        super(_Cell, self).__init__(parent)
        self._tc = tc

    def __eq__(self, other: object) -> bool:
        """|True| if this object proxies the same element as `other`.

        Equality for proxy objects is defined as referring to the same XML element, whether or not
        they are the same proxy object instance.
        """
        if not isinstance(other, type(self)):
            return False
        return self._tc is other._tc

    def __ne__(self, other: object) -> bool:
        if not isinstance(other, type(self)):
            return True
        return self._tc is not other._tc

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this cell.

        Provides access to fill properties such as foreground color.
        """
        tcPr = self._tc.get_or_add_tcPr()
        return FillFormat.from_fill_parent(tcPr)

    @property
    def is_merge_origin(self) -> bool:
        """True if this cell is the top-left grid cell in a merged cell."""
        return self._tc.is_merge_origin

    @property
    def is_spanned(self) -> bool:
        """True if this cell is spanned by a merge-origin cell.

        A merge-origin cell "spans" the other grid cells in its merge range, consuming their area
        and "shadowing" the spanned grid cells.

        Note this value is |False| for a merge-origin cell. A merge-origin cell spans other grid
        cells, but is not itself a spanned cell.
        """
        return self._tc.is_spanned

    @property
    def margin_left(self) -> Length:
        """Left margin of cells.

        Read/write. If assigned |None|, the default value is used, 0.1 inches for left and right
        margins and 0.05 inches for top and bottom.
        """
        return self._tc.marL

    @margin_left.setter
    def margin_left(self, margin_left: Length | None):
        self._validate_margin_value(margin_left)
        self._tc.marL = margin_left

    @property
    def margin_right(self) -> Length:
        """Right margin of cell."""
        return self._tc.marR

    @margin_right.setter
    def margin_right(self, margin_right: Length | None):
        self._validate_margin_value(margin_right)
        self._tc.marR = margin_right

    @property
    def margin_top(self) -> Length:
        """Top margin of cell."""
        return self._tc.marT

    @margin_top.setter
    def margin_top(self, margin_top: Length | None):
        self._validate_margin_value(margin_top)
        self._tc.marT = margin_top

    @property
    def margin_bottom(self) -> Length:
        """Bottom margin of cell."""
        return self._tc.marB

    @margin_bottom.setter
    def margin_bottom(self, margin_bottom: Length | None):
        self._validate_margin_value(margin_bottom)
        self._tc.marB = margin_bottom

    def merge(self, other_cell: _Cell) -> None:
        """Create merged cell from this cell to `other_cell`.

        This cell and `other_cell` specify opposite corners of the merged cell range. Either
        diagonal of the cell region may be specified in either order, e.g. self=bottom-right,
        other_cell=top-left, etc.

        Raises |ValueError| if the specified range already contains merged cells anywhere within
        its extents or if `other_cell` is not in the same table as `self`.
        """
        tc_range = TcRange(self._tc, other_cell._tc)

        if not tc_range.in_same_table:
            raise ValueError("other_cell from different table")
        if tc_range.contains_merged_cell:
            raise ValueError("range contains one or more merged cells")

        tc_range.move_content_to_origin()
        _apply_merge_topology(tc_range)

    def extend_merge(self, other_cell: _Cell) -> None:
        """Extend this merge-origin cell through `other_cell` atomically.

        paper-pptx addition. `other_cell` is the requested new lower-right corner. The existing
        rectangular merge may grow rightward, downward, or in both directions. Only paragraphs
        from newly absorbed cells move to this origin, in row-major order; existing merged content
        is not moved again.

        The existing merge must have canonical continuation topology and every newly absorbed cell
        must be unmerged. Malformed or conflicting merge structure raises
        |UnsupportedStructureError| without mutation. A stale origin or target raises
        |TargetNotFoundError|. Passing a non-cell, a cell from another table, a non-origin receiver,
        or a corner that would shrink or move the merge raises |ValueError|.
        """
        if not isinstance(other_cell, _Cell):
            raise ValueError("other_cell must be a table cell")

        from pptx._ownership import require_element_attached

        require_element_attached(self._tc, self.part, argument="merge origin")
        require_element_attached(other_cell._tc, other_cell.part, argument="merge target")

        tbl = self._tc.tbl
        if tbl is not other_cell._tc.tbl:
            raise ValueError("other_cell from different table")
        if not self.is_merge_origin:
            raise ValueError(
                "cell is not a merge-origin cell; use merge() to create a new merge"
            )

        current_region = _MergeRegion(
            self._tc.row_idx,
            self._tc.col_idx,
            self._tc.rowSpan,
            self._tc.gridSpan,
        )
        _validate_merge_topology(tbl, current_region)

        requested_region = _MergeRegion(
            current_region.top,
            current_region.left,
            other_cell._tc.row_idx - current_region.top + 1,
            other_cell._tc.col_idx - current_region.left + 1,
        )
        if (
            requested_region.bottom < current_region.bottom
            or requested_region.right < current_region.right
        ):
            raise ValueError(
                "other_cell must be at or beyond the current lower-right merge corner; "
                "merge shrinking and upward/leftward extension are not supported"
            )
        if (
            requested_region.bottom == current_region.bottom
            and requested_region.right == current_region.right
        ):
            return

        _validate_requested_merge_extension(tbl, self._tc, current_region, requested_region)

        from pptx._transaction import PackageTransaction

        requested_range = TcRange(self._tc, other_cell._tc)
        with PackageTransaction(self.part.package, self, other_cell):
            for tc in requested_range.iter_tcs():
                if tc.row_idx <= current_region.bottom and tc.col_idx <= current_region.right:
                    continue
                self._tc.append_ps_from(tc)
            _apply_merge_topology(requested_range)

    @property
    def span_height(self) -> int:
        """int count of rows spanned by this cell.

        The value of this property may be misleading (often 1) on cells where `.is_merge_origin`
        is not |True|, since only a merge-origin cell contains complete span information. This
        property is only intended for use on cells known to be a merge origin by testing
        `.is_merge_origin`.
        """
        return self._tc.rowSpan

    @property
    def span_width(self) -> int:
        """int count of columns spanned by this cell.

        The value of this property may be misleading (often 1) on cells where `.is_merge_origin`
        is not |True|, since only a merge-origin cell contains complete span information. This
        property is only intended for use on cells known to be a merge origin by testing
        `.is_merge_origin`.
        """
        return self._tc.gridSpan

    def split(self) -> None:
        """Remove merge from this (merge-origin) cell.

        The merged cell represented by this object will be "unmerged", yielding a separate
        unmerged cell for each grid cell previously spanned by this merge.

        Raises |ValueError| when this cell is not a merge-origin cell. Test with
        `.is_merge_origin` before calling.
        """
        if not self.is_merge_origin:
            raise ValueError("not a merge-origin cell; only a merge-origin cell can be sp" "lit")

        tc_range = TcRange.from_merge_origin(self._tc)

        for tc in tc_range.iter_tcs():
            tc.rowSpan = tc.gridSpan = 1
            tc.hMerge = tc.vMerge = False

    @property
    def text(self) -> str:
        """Textual content of cell as a single string.

        The returned string will contain a newline character (`"\\n"`) separating each paragraph
        and a vertical-tab (`"\\v"`) character for each line break (soft carriage return) in the
        cell's text.

        Assignment to `text` replaces all text currently contained in the cell. A newline
        character (`"\\n"`) in the assigned text causes a new paragraph to be started. A
        vertical-tab (`"\\v"`) character in the assigned text causes a line-break (soft
        carriage-return) to be inserted. (The vertical-tab character appears in clipboard text
        copied from PowerPoint as its encoding of line-breaks.)
        """
        return self.text_frame.text

    @text.setter
    def text(self, text: str):
        self.text_frame.text = text

    @property
    def text_frame(self) -> TextFrame:
        """|TextFrame| containing the text that appears in the cell."""
        txBody = self._tc.get_or_add_txBody()
        return TextFrame(txBody, self)

    @property
    def vertical_anchor(self) -> MSO_VERTICAL_ANCHOR | None:
        """Vertical alignment of this cell.

        This value is a member of the :ref:`MsoVerticalAnchor` enumeration or |None|. A value of
        |None| indicates the cell has no explicitly applied vertical anchor setting and its
        effective value is inherited from its style-hierarchy ancestors.

        Assigning |None| to this property causes any explicitly applied vertical anchor setting to
        be cleared and inheritance of its effective value to be restored.
        """
        return self._tc.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, mso_anchor_idx: MSO_VERTICAL_ANCHOR | None):
        self._tc.anchor = mso_anchor_idx

    @staticmethod
    def _validate_margin_value(margin_value: Length | None) -> None:
        """Raise ValueError if `margin_value` is not a positive integer value or |None|."""
        if not isinstance(margin_value, int) and margin_value is not None:
            tmpl = "margin value must be integer or None, got '%s'"
            raise TypeError(tmpl % margin_value)


class _Column(Subshape):
    """Table column"""

    def __init__(self, gridCol: CT_TableCol, parent: _ColumnCollection):
        super(_Column, self).__init__(parent)
        self._parent = parent
        self._gridCol = gridCol

    @property
    def width(self) -> Length:
        """Width of column in EMU."""
        return self._gridCol.w

    @width.setter
    def width(self, width: Length):
        self._gridCol.w = width
        self._parent.notify_width_changed()


class _Row(Subshape):
    """Table row"""

    def __init__(self, tr: CT_TableRow, parent: _RowCollection):
        super(_Row, self).__init__(parent)
        self._parent = parent
        self._tr = tr

    @property
    def cells(self):
        """Read-only reference to collection of cells in row.

        An individual cell is referenced using list notation, e.g. `cell = row.cells[0]`.
        """
        return _CellCollection(self._tr, self)

    @property
    def height(self) -> Length:
        """Height of row in EMU."""
        return self._tr.h

    @height.setter
    def height(self, height: Length):
        self._tr.h = height
        self._parent.notify_height_changed()


class _CellCollection(Subshape):
    """Horizontal sequence of row cells"""

    def __init__(self, tr: CT_TableRow, parent: _Row):
        super(_CellCollection, self).__init__(parent)
        self._parent = parent
        self._tr = tr

    def __getitem__(self, idx: int) -> _Cell:
        """Provides indexed access, (e.g. 'cells[0]')."""
        if idx < 0 or idx >= len(self._tr.tc_lst):
            msg = "cell index [%d] out of range" % idx
            raise IndexError(msg)
        return _Cell(self._tr.tc_lst[idx], self)

    def __iter__(self) -> Iterator[_Cell]:
        """Provides iterability."""
        return (_Cell(tc, self) for tc in self._tr.tc_lst)

    def __len__(self) -> int:
        """Supports len() function (e.g. 'len(cells) == 1')."""
        return len(self._tr.tc_lst)


class _ColumnCollection(Subshape):
    """Sequence of table columns."""

    def __init__(self, tbl: CT_Table, parent: Table):
        super(_ColumnCollection, self).__init__(parent)
        self._parent = parent
        self._tbl = tbl

    def __getitem__(self, idx: int):
        """Provides indexed access, (e.g. 'columns[0]')."""
        if idx < 0 or idx >= len(self._tbl.tblGrid.gridCol_lst):
            msg = "column index [%d] out of range" % idx
            raise IndexError(msg)
        return _Column(self._tbl.tblGrid.gridCol_lst[idx], self)

    def __len__(self):
        """Supports len() function (e.g. 'len(columns) == 1')."""
        return len(self._tbl.tblGrid.gridCol_lst)

    def notify_width_changed(self):
        """Called by a column when its width changes. Pass along to parent."""
        self._parent.notify_width_changed()


class _RowCollection(Subshape):
    """Sequence of table rows"""

    def __init__(self, tbl: CT_Table, parent: Table):
        super(_RowCollection, self).__init__(parent)
        self._parent = parent
        self._tbl = tbl

    def __getitem__(self, idx: int) -> _Row:
        """Provides indexed access, (e.g. 'rows[0]')."""
        if idx < 0 or idx >= len(self):
            msg = "row index [%d] out of range" % idx
            raise IndexError(msg)
        return _Row(self._tbl.tr_lst[idx], self)

    def __len__(self):
        """Supports len() function (e.g. 'len(rows) == 1')."""
        return len(self._tbl.tr_lst)

    def notify_height_changed(self):
        """Called by a row when its height changes. Pass along to parent."""
        self._parent.notify_height_changed()
