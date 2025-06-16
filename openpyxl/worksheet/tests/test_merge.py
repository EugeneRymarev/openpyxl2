# Copyright (c) 2010-2025 openpyxl
import copy

import pytest
from openpyxl.styles.borders import Border
from openpyxl.styles.borders import Side
from openpyxl.styles.protection import Protection
from openpyxl.tests.helper import compare_xml
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def merge_cell():
    from openpyxl.worksheet.merge import MergeCell

    return MergeCell


@pytest.fixture
def merged_cell_range():
    from openpyxl.worksheet.merge import MergedCellRange

    return MergedCellRange


class TestMergeCell:
    def test_ctor(self, merge_cell):
        cell = merge_cell("A1")
        node = cell.to_tree()
        xml = tostring(node)
        expected = '<mergeCell ref="A1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, merge_cell):
        xml = '<mergeCell ref="A1"/>'
        node = fromstring(xml)
        cell = merge_cell.from_tree(node)
        assert cell == CellRange("A1")

    def test_copy(self, merge_cell):
        cell = merge_cell("A1")
        cp = copy.copy(cell)
        assert cp == cell


def default_border():
    return Side(border_style=None, color=None)


def thin_border():
    return Side(border_style="thin", color="000000")


def double_border():
    return Side(border_style="double", color="000000")


def thick_border():
    return Side(border_style="thick", color="000000")


def start_border():
    return Border(
        top=thick_border(),
        left=thick_border(),
        right=thin_border(),
        bottom=double_border(),
    )


class TestMergedCellRange:
    def test_ctor(self, merged_cell_range):
        ws = Workbook().active
        cells = merged_cell_range(ws, "A1:E4")
        assert cells.start_cell == ws["A1"]

    @pytest.mark.parametrize("end", [("C1"), ("A3"), ("C3")])
    def test_get_borders(self, merged_cell_range, end):
        ws = Workbook().active
        ws["A1"].border = Border(top=thick_border(), left=thick_border())
        ws[end].border = Border(right=thin_border(), bottom=double_border())
        mcr = merged_cell_range(ws, "A1:" + end)
        assert mcr.start_cell.coordinate == "A1"
        assert mcr.start_cell.border == start_border()

    def test_format_1x3(self, merged_cell_range):
        ws = Workbook().active
        mcr = merged_cell_range(ws, "A1:C1")
        mcr.start_cell.border = start_border()
        mcr.format()
        b1_border = Border(
            top=thick_border(),
            left=default_border(),
            right=default_border(),
            bottom=double_border(),
            diagonal=default_border(),
        )
        assert ws["B1"].border == b1_border
        c1_border = Border(
            top=thick_border(),
            left=default_border(),
            right=thin_border(),
            bottom=double_border(),
            diagonal=default_border(),
        )
        assert ws["C1"].border == c1_border

    def test_format_3x1(self, merged_cell_range):
        ws = Workbook().active
        mcr = merged_cell_range(ws, "A1:A3")
        mcr.start_cell.border = start_border()
        mcr.format()
        a2_border = Border(
            top=default_border(),
            left=thick_border(),
            right=thin_border(),
            bottom=default_border(),
            diagonal=default_border(),
        )
        assert ws["A2"].border == a2_border
        a3_border = Border(
            top=default_border(),
            left=thick_border(),
            right=thin_border(),
            bottom=double_border(),
            diagonal=default_border(),
        )
        assert ws["A3"].border == a3_border

    def test_format_3x3(self, merged_cell_range):
        ws = Workbook().active
        mcr = merged_cell_range(ws, "A1:C3")
        mcr.start_cell.border = start_border()
        mcr.format()
        for coord in mcr.top:
            cell = ws._cells.get(coord)
            assert cell.border.top == thick_border()
        for coord in mcr.bottom:
            cell = ws._cells.get(coord)
            assert cell.border.bottom == double_border()
        for coord in mcr.left:
            cell = ws._cells.get(coord)
            assert cell.border.left == thick_border()
        for coord in mcr.right:
            cell = ws._cells.get(coord)
            assert cell.border.right == thin_border()
        b2_border = Border(
            top=default_border(),
            left=default_border(),
            right=default_border(),
            bottom=default_border(),
            diagonal=default_border(),
        )
        assert ws["B2"].border == b2_border

    def test_format_protection(self, merged_cell_range):
        ws = Workbook().active
        mcr = merged_cell_range(ws, "A1:C1")
        mcr.start_cell.protection = Protection(locked=False, hidden=False)
        mcr.format()
        assert ws["B1"].protection == Protection(locked=False, hidden=False)
        assert ws["C1"].protection == Protection(locked=False, hidden=False)
        mcr = merged_cell_range(ws, "D1:F1")
        mcr.start_cell.protection = Protection(locked=True, hidden=True)
        mcr.format()
        assert ws["E1"].protection == Protection(locked=True, hidden=True)
        assert ws["F1"].protection == Protection(locked=True, hidden=True)

    def test_copy(self, merged_cell_range):
        ws = Workbook().active
        mcr1 = merged_cell_range(ws, "A1:J6")
        mcr2 = copy.copy(mcr1)
        assert mcr2 == mcr1

    def test_contains(seld, merged_cell_range):
        ws = Workbook().active
        mcr = merged_cell_range(ws, "B2:M20")
        assert "D4" in mcr

    def test_not_contained(self, merged_cell_range):
        ws = Workbook().active
        mcr = merged_cell_range(ws, "B2:M20")
        assert "A1" not in mcr

    def test_empty_side(seld, merged_cell_range):
        ws = Workbook().active
        ws["A1"].border = Border(bottom=Side(style="thin"))
        mcr = merged_cell_range(ws, "A1:C3")
        mcr.format()
        assert ws["C3"].border.bottom == Side(style="thin")
