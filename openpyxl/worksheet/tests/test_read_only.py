# Copyright (c) 2010-2025 openpyxl
import datetime
import io
import zipfile

import pytest

from openpyxl.cell.read_only import EMPTY_CELL
from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.reader.excel import load_workbook
from openpyxl.styles.styleable import StyleArray


@pytest.fixture
def dummy_workbook():
    class Workbook:
        epoch = None
        _cell_styles = [StyleArray([0, 0, 0, 0, 0, 0, 0, 0, 0])]
        data_only = False

        def __init__(self):
            self.sheetnames = []
            self._archive = zipfile.ZipFile(io.BytesIO(), "w")
            self._date_formats = set()
            self._timedelta_formats = set()

    return Workbook()


@pytest.fixture
def read_only_worksheet(dummy_workbook, datadir):
    from openpyxl.worksheet._read_only import ReadOnlyWorksheet

    datadir.chdir()
    wb = dummy_workbook
    wb._archive.write("sheet_inline_strings.xml", "sheet1.xml")
    ws = ReadOnlyWorksheet(wb, "Sheet", "sheet1.xml", [])
    return ws


class TestReadOnlyWorksheet:
    def test_from_xml(self, read_only_worksheet):
        ws = read_only_worksheet
        cells = tuple(ws.iter_rows(min_row=1, min_col=1, max_row=1, max_col=1))
        assert len(cells) == 1
        assert cells[0][0].value == "col1"

    @pytest.mark.parametrize("row, column", [(2, 1), (3, 1), (5, 1)])
    def test_read_cell_from_empty_row(
        self,
        dummy_workbook,
        read_only_worksheet,
        row,
        column,
    ):
        src = b"""
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="2"/>
                <row r="4"/>
            </sheetData>
        </worksheet>
        """
        wb = dummy_workbook
        wb._archive.writestr("sheet1.xml", src)
        ws = read_only_worksheet
        ws._xml = io.BytesIO(src)
        cell = ws._get_cell(row, column)
        assert cell is EMPTY_CELL

    def test_empty_cell(self, read_only_worksheet):
        row = [{"column": 4, "value": None, "row": 1}]
        ws = read_only_worksheet
        cells = ws._get_row(row, max_col=4, values_only=True)
        assert cells == (None, None, None, None)

    def test_pad_row_left(self, read_only_worksheet):
        row = [{"column": 4, "value": 4}, {"column": 8, "value": 8}]
        ws = read_only_worksheet
        cells = ws._get_row(row, max_col=4, values_only=True)
        assert cells == (None, None, None, 4)

    def test_pad_row(self, read_only_worksheet):
        row = [{"column": 4, "value": 4}, {"column": 8, "value": 8}]
        ws = read_only_worksheet
        cells = ws._get_row(row, min_col=4, max_col=8, values_only=True)
        assert cells == (4, None, None, None, 8)

    def test_pad_row_right(self, read_only_worksheet):
        row = [{"column": 4, "value": 4}, {"column": 8, "value": 8}]
        ws = read_only_worksheet
        cells = ws._get_row(row, min_col=6, max_col=10, values_only=True)
        assert cells == (None, None, 8, None, None)

    def test_pad_row_cells(self, read_only_worksheet):
        row = [{"column": 4, "value": 4, "row": 2}, {"column": 8, "value": 8, "row": 2}]
        ws = read_only_worksheet
        cells = ws._get_row(row, min_col=6, max_col=10)
        expected = (
            EMPTY_CELL,
            EMPTY_CELL,
            ReadOnlyCell(ws, 2, 8, 8, "n", 0),
            EMPTY_CELL,
            EMPTY_CELL,
        )
        assert cells == expected

    def test_read_rows(self, read_only_worksheet):
        ws = read_only_worksheet
        rows = ws._cells_by_row(
            min_row=1, max_row=None, min_col=1, max_col=3, values_only=True
        )
        rows = list(ws.rows)
        assert len(rows) == 10

    def test_pad_rows_before(self, read_only_worksheet):
        ws = read_only_worksheet
        rows = ws._cells_by_row(
            min_row=8,
            max_row=10,
            min_col=1,
            max_col=3,
            values_only=True,
        )
        assert list(rows) == [(None, None, None), (None, None, None), (7, 8, 9)]

    def test_pad_rows_after(self, read_only_worksheet):
        ws = read_only_worksheet
        rows = ws._cells_by_row(
            min_row=4,
            max_row=6,
            min_col=1,
            max_col=3,
            values_only=True,
        )
        assert list(rows) == [(7, 8, 9), (None, None, None), (None, None, None)]

    def test_pad_rows_between(self, read_only_worksheet):
        ws = read_only_worksheet
        rows = ws._cells_by_row(
            min_row=4,
            max_row=None,
            min_col=1,
            max_col=3,
            values_only=True,
        )
        expected = [
            (7, 8, 9),
            (None, None, None),
            (None, None, None),
            (None, None, None),
            (None, None, None),
            (None, None, None),
            (7, 8, 9),
        ]
        assert list(rows) == expected

    def test_pad_rows_bounded(self, read_only_worksheet):
        ws = read_only_worksheet
        rows = ws._cells_by_row(
            min_row=8,
            max_row=15,
            min_col=1,
            max_col=3,
            values_only=True,
        )
        assert list(rows) == [(None, None, None), (None, None, None), (7, 8, 9)]

    def test_calculate_dimension(self, read_only_worksheet):
        ws = read_only_worksheet
        assert ws.calculate_dimension(True) == "A1:C10"

    def test_reset_dimensions(self, read_only_worksheet):
        ws = read_only_worksheet
        ws._max_row = 5
        ws._max_column = 10
        ws.reset_dimensions()
        assert ws.max_row is ws.max_column is None

    def test_cell(self, read_only_worksheet):
        ws = read_only_worksheet
        c = ws.cell(row=1, column=1)
        assert c.value == "col1"

    def test_iter(self, read_only_worksheet):
        ws = read_only_worksheet
        row = None
        for row in ws:
            pass
        c = row[-1]
        assert c.value == 9

    def test_cleanup_on_break(self, read_only_worksheet):
        xml = b"""
        <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        </sheet>
        """
        src = io.BytesIO(xml)

        def mock_source():
            return src

        ws = read_only_worksheet
        ws._get_source = mock_source
        for row in ws:
            break
        assert src.closed


def test_implementation_compatibility(read_only_worksheet, dummy_workbook):
    from openpyxl.worksheet.worksheet import Worksheet

    std = Worksheet(dummy_workbook)
    std_attrs = set(std.__dict__)
    std_only = {
        "HeaderFooter",
        "_WorkbookChild__title",
        "_cells",
        "_charts",
        "_comments",
        "_current_row",
        "_drawing",
        "_hyperlinks",
        "_images",
        "_parent",
        "_pivots",
        "_print_area",
        "_print_cols",
        "_print_rows",
        "_rels",
        "_shapes",
        "_tables",
        "auto_filter",
        "col_breaks",
        "column_dimensions",
        "controls",
        "conditional_formatting",
        "data_validations",
        "legacy_drawing",
        "merged_cells",
        "page_margins",
        "page_setup",
        "print_options",
        "protection",
        "row_breaks",
        "row_dimensions",
        "scenarios",
        "sheet_format",
        "sheet_properties",
        "sheet_state",
        "views",
    }
    ro = read_only_worksheet
    ro_attrs = set(ro.__dict__)
    ro_only = {"_worksheet_path", "parent", "title", "_shared_strings"}
    assert std_attrs > std_only
    assert ro_attrs > ro_only
    assert not ro_attrs - ro_only - std_attrs
    extra = std_attrs - std_only - ro_attrs
    assert not extra, f"Missing attributes {extra}"


def test_read_datetime(datadir):
    # Check read only sheets correctly parse datetime and timedelta cells where appropriate
    datadir.chdir()
    wb = load_workbook("test_datetime.xlsx", read_only=True)
    ws = wb.active
    assert type(ws["A1"].value) == datetime.timedelta
    assert type(ws["A2"].value) == datetime.datetime
