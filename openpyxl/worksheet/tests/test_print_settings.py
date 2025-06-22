# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.worksheet.cell_range import CellRange


@pytest.fixture
def col_range():
    from openpyxl.worksheet.print_settings import ColRange

    return ColRange


@pytest.fixture
def row_range():
    from openpyxl.worksheet.print_settings import RowRange

    return RowRange


@pytest.fixture
def print_titles():
    from openpyxl.worksheet.print_settings import PrintTitles

    return PrintTitles


@pytest.fixture
def print_area():
    from openpyxl.worksheet.print_settings import PrintArea

    return PrintArea


class TestColRange:
    def test_from_string(self, col_range):
        cols = col_range("$B:$E")
        assert cols.min_col == "B"
        assert cols.max_col == "E"

    def test_str(self, col_range):
        cols = col_range(min_col="A", max_col="D")
        assert str(cols) == "$A:$D"

    def test_repr(self, col_range):
        cols = col_range(min_col="A", max_col="D")
        assert repr(cols) == "Range of columns from 'A' to 'D'"

    @pytest.mark.parametrize("expected", ["$B:$E", "B:E"])
    def test_eq(self, col_range, expected):
        cols = col_range(min_col="B", max_col="E")
        assert cols == expected


class TestRowRange:
    def test_from_string(self, row_range):
        rows = row_range("$2:$6")
        assert rows.min_row == 2
        assert rows.max_row == 6

    def test_str(self, row_range):
        cols = row_range(min_row=1, max_row=4)
        assert str(cols) == "$1:$4"

    def test_repr(self, row_range):
        cols = row_range(min_row=2, max_row=6)
        assert repr(cols) == "Range of rows from '2' to '6'"

    @pytest.mark.parametrize("expected", ["$2:$7", "2:7"])
    def test_eq(self, row_range, expected):
        rows = row_range(min_row=2, max_row=7)
        assert rows == expected


class TestPrintTitles:
    @pytest.mark.parametrize(
        "value, expected",
        [
            ["'Sheet1'!$1:$2,$A:$A", "'Sheet1'!$1:$2,'Sheet1'!$A:$A"],
            ["'Sheet 1'!$A:$A", "'Sheet 1'!$A:$A"],
            ["Sheet1!$5:$17", "'Sheet1'!$5:$17"],
            ["Table1!$J:$J,Table1!$10:$10", "'Table1'!$10:$10,'Table1'!$J:$J"],
        ],
    )
    def test_from_string(self, print_titles, value, expected):
        titles = print_titles.from_string(value)
        assert str(titles) == expected

    def test_eq(self, print_titles):
        assert print_titles.from_string("'Sheet 1'!$A:$A") == "'Sheet 1'!$A:$A"


class TestPrintArea:
    @pytest.mark.parametrize(
        "value, expected",
        [
            ("Sheet1!$A$1:$E$15", {CellRange("A1:E15")}),
            ("$A$1:$E$15", {CellRange("A1:E15")}),
            (
                "'Blatt1'!$A$1:$F$14,'Blatt1'!$H$10:$I$17,Blatt1!$I$16:$K$25",
                {CellRange("A1:F14"), CellRange("H10:I17"), CellRange("I16:K25")},
            ),
            ("MySheet!#REF!", set()),
            ("'C,D'!$A$1:$B$3", {CellRange("A1:B3")}),
            (
                "Sheet!$A$1:$D$5,Sheet!$B$9:$F$14",
                {CellRange("A1:D5"), CellRange("B9:F14")},
            ),
        ],
    )
    def test_from_string(self, print_area, value, expected):
        area = print_area.from_string(value)
        assert area.ranges == expected

    def test_empty(self, print_area):
        area = print_area()
        assert area.ranges == set()

    def test_str(self, print_area):
        area = print_area.from_string("Sheet!$A$1:$D$5,Sheet!$B$9:$F$14")
        assert str(area) == "''!$A$1:$D$5,''!$B$9:$F$14"

    def test_eq(self, print_area):
        area = print_area.from_string("Sheet1!$A$1:$E$15")
        assert area == "''!$A$1:$E$15"
