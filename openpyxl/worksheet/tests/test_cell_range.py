# Copyright (c) 2010-2025 openpyxl
import copy

import pytest

from openpyxl.utils.exceptions import CellNotMergedException


@pytest.fixture
def cell_range():
    from openpyxl.worksheet.cell_range import CellRange

    return CellRange


@pytest.fixture
def multi_cell_range():
    from openpyxl.worksheet.cell_range import MultiCellRange

    return MultiCellRange


class TestCellRange:
    def test_ctor(self, cell_range):
        cr = cell_range(min_col=1, min_row=1, max_col=5, max_row=7)
        assert (cr.min_col, cr.min_row, cr.max_col, cr.max_row) == (1, 1, 5, 7)
        assert cr.coord == "A1:E7"

    def test_dict(self, cell_range):
        cr = cell_range("Sheet1!A1:E7")
        assert cr.coord == "A1:E7"
        assert dict(cr) == {"max_col": 5, "max_row": 7, "min_col": 1, "min_row": 1}

    def test_max_row_too_small(self, cell_range):
        with pytest.raises(ValueError):
            cell_range("A4:B1")

    def test_max_col_too_small(self, cell_range):
        with pytest.raises(ValueError):
            cell_range("F1:B5")

    @pytest.mark.parametrize(
        "range_string, title, coord",
        [("Sheet1!$A$1:B4", "Sheet1", "A1:B4"), ("A1:B4", None, "A1:B4")],
    )
    def test_from_string(self, cell_range, range_string, title, coord):
        cr = cell_range(range_string)
        assert cr.coord == coord
        assert cr.title == title

    def test_repr(self, cell_range):
        cr = cell_range("Sheet1!$A$1:B4")
        assert repr(cr) == "<CellRange 'Sheet1'!A1:B4>"

    def test_str(self, cell_range):
        cr = cell_range("'Sheet 1'!$A$1:B4")
        assert str(cr) == "'Sheet 1'!A1:B4"
        cr = cell_range("A1")
        assert str(cr) == "A1"

    def test_eq(self, cell_range):
        cr1 = cell_range("'Sheet 1'!$A$1:B4")
        cr2 = cell_range("'Sheet 1'!$A$1:B4")
        assert cr1 == cr2

    def test_ne(self, cell_range):
        cr1 = cell_range("'Sheet 1'!$A$1:B4")
        cr2 = cell_range("Sheet1!$A$1:B4")
        assert cr1 != cr2

    def test_copy(self, cell_range):
        cr1 = cell_range("Sheet1!$A$1:B4")
        cr2 = copy.copy(cr1)
        assert cr2 is not cr1

    def test_shift(self, cell_range):
        cr = cell_range("A1:B4")
        cr.shift(1, 2)
        assert cr.coord == "B3:C6"

    def test_shift_negative(self, cell_range):
        cr = cell_range("A1:B4")
        with pytest.raises(ValueError):
            cr.shift(-1, 2)

    def test_union(self, cell_range):
        cr1 = cell_range("A1:D4")
        cr2 = cell_range("E5:K10")
        cr3 = cr1.union(cr2)
        assert cr3.bounds == (1, 1, 11, 10)

    def test_no_union(self, cell_range):
        cr1 = cell_range("Sheet1!A1:D4")
        cr2 = cell_range("Sheet2!E5:K10")
        with pytest.raises(ValueError):
            cr1.union(cr2)

    def test_expand(self, cell_range):
        cr = cell_range("E5:K10")
        cr.expand(right=2, down=2, left=1, up=2)
        assert cr.coord == "D3:M12"

    def test_shrink(self, cell_range):
        cr = cell_range("E5:K10")
        cr.shrink(right=2, bottom=2, left=1, top=2)
        assert cr.coord == "F7:I8"

    def test_size(self, cell_range):
        cr = cell_range("E5:K10")
        assert cr.size == {"columns": 7, "rows": 6}

    def test_intersection(self, cell_range):
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("D2:F7")
        cr3 = cr1.intersection(cr2)
        assert cr3.coord == "E5:F7"

    def test_no_intersection(self, cell_range):
        cr1 = cell_range("A1:F5")
        cr2 = cell_range("M5:P17")
        with pytest.raises(ValueError):
            assert cr1 & cr2 == cell_range("A1")

    def test_isdisjoint_order(self, cell_range):
        """
        Order of the test does not matter
        """
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("A1:C12")
        assert cr1.isdisjoint(cr2) is cr2.isdisjoint(cr1)

    def test_isdisjoint_by_col(self, cell_range):
        """
        Tested ranges differ only by columns
        """
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("A5:C10")
        assert cr1.isdisjoint(cr2) is True

    def test_isdisjoint_by_row(self, cell_range):
        """
        Tested ranges differ only by rows
        """
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("E12:K12")
        assert cr1.isdisjoint(cr2) is True

    def test_isdisjoint_in_both(self, cell_range):
        """
        Tested ranges differ in both rows and columns
        """
        cr1 = cell_range("A1:B2")
        cr2 = cell_range("D4")
        assert cr1.isdisjoint(cr2) is True

    def test_is_not_disjoint(self, cell_range):
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("D2:F7")
        assert cr1.isdisjoint(cr2) is False

    def test_is_not_disjoint_in_both(self, cell_range):
        """
        Tested ranges overlap in both rows and columns
        """
        cr1 = cell_range("A1:D4")
        cr2 = cell_range("B2:C3")
        assert cr1.isdisjoint(cr2) is False

    def test_issubset(self, cell_range):
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("F6:J8")
        assert cr2.issubset(cr1) is True

    def test_is_not_subset(self, cell_range):
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("D4:M8")
        assert cr2.issubset(cr1) is False

    def test_issuperset(self, cell_range):
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("F6:J8")
        assert cr1.issuperset(cr2) is True

    def test_is_not_superset(self, cell_range):
        cr1 = cell_range("E5:K10")
        cr2 = cell_range("A1:D4")
        assert cr1.issuperset(cr2) is False

    def test_contains(self, cell_range):
        cr = cell_range("A1:F10")
        assert "B3" in cr

    def test_doesnt_contain(self, cell_range):
        cr = cell_range("A1:F10")
        assert not "M1" in cr

    @pytest.mark.parametrize(
        "r1, r2, expected",
        [("Sheet1!A1:B4", "Sheet1!D5:E5", None), ("Sheet1!A1:B4", "D5:E5", None)],
    )
    def test_check_title(self, cell_range, r1, r2, expected):
        cr1 = cell_range(r1)
        cr2 = cell_range(r2)
        assert cr1._check_title(cr2) is expected

    @pytest.mark.parametrize(
        "r1, r2",
        [("A1:B4", "Sheet1!D5:E5"), ("Sheet1!A1:B4", "Sheet2!D5:E5")],
    )
    def test_different_worksheets(self, cell_range, r1, r2):
        cr1 = cell_range(r1)
        cr2 = cell_range(r2)
        with pytest.raises(ValueError):
            cr1._check_title(cr2)

    def test_lt(self, cell_range):
        cr1 = cell_range("A1:F5")
        cr2 = cell_range("A2:F4")
        assert cr2 < cr1

    def test_gt(self, cell_range):
        cr1 = cell_range("A1:F5")
        cr2 = cell_range("A2:F4")
        assert cr1 > cr2

    def test_edge_cells(self, cell_range):
        cr = cell_range("A1:C3")
        assert cr.top == [(1, 1), (1, 2), (1, 3)]
        assert cr.bottom == [(3, 1), (3, 2), (3, 3)]
        assert cr.left == [(1, 1), (2, 1), (3, 1)]
        assert cr.right == [(1, 3), (2, 3), (3, 3)]

    def test_rows(self, cell_range):
        cr = cell_range("A1:B3")
        assert list(cr.rows) == [[(1, 1), (1, 2)], [(2, 1), (2, 2)], [(3, 1), (3, 2)]]

    def test_cols(self, cell_range):
        cr = cell_range("A1:B3")
        assert list(cr.cols) == [[(1, 1), (2, 1), (3, 1)], [(1, 2), (2, 2), (3, 2)]]

    def test_cells(self, cell_range):
        cr = cell_range("A1:B3")
        cells = list(cr.cells)
        assert cells == [(1, 1), (1, 2), (2, 1), (2, 2), (3, 1), (3, 2)]


class TestMultiCellRange:
    def test_ctor(self, multi_cell_range, cell_range):
        cr = cell_range("A1")
        cells = multi_cell_range(ranges=[cr])
        assert cells.ranges == {cr}

    def test_from_string(self, multi_cell_range, cell_range):
        cells = multi_cell_range("A1 B2:B5")
        assert cells.ranges == {cell_range("A1"), cell_range("B2:B5")}

    def test_add_coord(self, multi_cell_range, cell_range):
        cr = cell_range("A1")
        cells = multi_cell_range(ranges=[cr])
        cells.add("B2")
        assert cells.ranges == {cr, cell_range("B2")}

    def test_add_cell_range(self, multi_cell_range, cell_range):
        cr1 = cell_range("A1")
        cr2 = cell_range("B2")
        cells = multi_cell_range(ranges=[cr1])
        cells.add(cr2)
        assert cells.ranges == {cr1, cr2}

    def test_iadd(self, multi_cell_range):
        cells = multi_cell_range()
        cells.add("A1")
        assert cells == "A1"

    def test_avoid_duplicates(self, multi_cell_range):
        cells = multi_cell_range("A1:D4")
        cells.add("A3")
        assert cells == "A1:D4"

    def test_repr(self, multi_cell_range, cell_range):
        cr1 = cell_range("a1")
        cr2 = cell_range("B2")
        cells = multi_cell_range(ranges=[cr1, cr2])
        assert repr(cells) == "<MultiCellRange [A1 B2]>"

    def test_contains(self, multi_cell_range, cell_range):
        cr = cell_range("A1:E4")
        cells = multi_cell_range([cr])
        assert "C3" in cells

    def test_doesnt_contain(self, multi_cell_range):
        cells = multi_cell_range("A1:D5")
        assert "F6" not in cells

    def test_eq(self, multi_cell_range):
        cells = multi_cell_range("A1:D4 E5")
        assert cells == "A1:D4 E5"

    def test_ne(self, multi_cell_range):
        cells = multi_cell_range("A1")
        assert cells != "B4"

    def test_empty(self, multi_cell_range):
        cells = multi_cell_range()
        assert bool(cells) is False

    def test_not_empty(self, multi_cell_range):
        cells = multi_cell_range("A1")
        assert bool(cells) is True

    def test_remove(self, multi_cell_range):
        cells = multi_cell_range("A1:D4")
        cells.remove("A1:D4")

    def test_remove_invalid(self, multi_cell_range):
        cells = multi_cell_range("A1:D4")
        with pytest.raises(KeyError):
            cells.remove("A1")

    def test_iter(self, multi_cell_range, cell_range):
        cells = multi_cell_range("A1")
        assert list(cells) == [cell_range("A1")]

    def test_copy(self, multi_cell_range, cell_range):
        r1 = multi_cell_range("A1")
        r2 = copy.copy(r1)
        assert list(r1)[0] is not list(r2)[0]

    def test_getitem(self, multi_cell_range, cell_range):
        mcr = multi_cell_range("A1:B2")
        assert mcr["A2"] == "A1:B2"
        with pytest.raises(CellNotMergedException) as e:
            mcr["C3"]
