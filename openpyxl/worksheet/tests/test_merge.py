# Copyright (c) 2010-2024 openpyxl
from copy import copy

import pytest

from ..cell_range import CellRange
from openpyxl import Workbook
from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.styles.protection import Protection
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def MergeCell():
    from ..merge import MergeCell

    return MergeCell


class TestMergeCell:

    def test_ctor(self, MergeCell):
        cell = MergeCell("A1")
        node = cell.to_tree()
        xml = tostring(node)
        expected = "<mergeCell ref='A1' />"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, MergeCell):
        xml = "<mergeCell ref='A1' />"
        node = fromstring(xml)
        cell = MergeCell.from_tree(node)
        assert cell == CellRange("A1")

    def test_copy(self, MergeCell):
        cell = MergeCell("A1")
        cp = copy(cell)
        assert cp == cell


@pytest.fixture
def MergedCellRange():
    from ..merge import MergedCellRange

    return MergedCellRange


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

    def test_ctor(self, MergedCellRange):
        ws = Workbook().active
        cells = MergedCellRange(ws, "A1:E4")
        assert cells.start_cell == ws["A1"]

    @pytest.mark.parametrize("end", [("C1"), ("A3"), ("C3")])
    def test_get_borders(self, MergedCellRange, end):
        ws = Workbook().active
        ws["A1"].border = Border(top=thick_border(), left=thick_border())
        ws[end].border = Border(right=thin_border(), bottom=double_border())

        mcr = MergedCellRange(ws, "A1:" + end)
        assert mcr.start_cell.coordinate == "A1"
        assert mcr.start_cell.border == start_border()

    def test_format_1x3(self, MergedCellRange):
        ws = Workbook().active
        mcr = MergedCellRange(ws, "A1:C1")
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

    def test_format_3x1(self, MergedCellRange):
        ws = Workbook().active
        mcr = MergedCellRange(ws, "A1:A3")
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

    def test_format_3x3(self, MergedCellRange):
        ws = Workbook().active
        mcr = MergedCellRange(ws, "A1:C3")
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

    def test_format_protection(self, MergedCellRange):
        ws = Workbook().active
        mcr = MergedCellRange(ws, "A1:C1")
        mcr.start_cell.protection = Protection(locked=False, hidden=False)
        mcr.format()
        assert ws["B1"].protection == Protection(locked=False, hidden=False)
        assert ws["C1"].protection == Protection(locked=False, hidden=False)

        mcr = MergedCellRange(ws, "D1:F1")
        mcr.start_cell.protection = Protection(locked=True, hidden=True)
        mcr.format()
        assert ws["E1"].protection == Protection(locked=True, hidden=True)
        assert ws["F1"].protection == Protection(locked=True, hidden=True)

    def test_copy(self, MergedCellRange):
        ws = Workbook().active
        mcr1 = MergedCellRange(ws, "A1:J6")
        mcr2 = copy(mcr1)
        assert mcr2 == mcr1

    def test_contains(seld, MergedCellRange):
        ws = Workbook().active
        mcr = MergedCellRange(ws, "B2:M20")
        assert "D4" in mcr

    def test_not_contained(self, MergedCellRange):
        ws = Workbook().active
        mcr = MergedCellRange(ws, "B2:M20")
        assert "A1" not in mcr

    def test_empty_side(seld, MergedCellRange):

        ws = Workbook().active
        ws["A1"].border = Border(bottom=Side(style="thin"))
        mcr = MergedCellRange(ws, "A1:C3")

        mcr.format()
        assert ws["C3"].border.bottom == Side(style="thin")
