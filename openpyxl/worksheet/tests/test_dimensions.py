# Copyright (c) 2010-2025 openpyxl
import copy

import pytest

from openpyxl.styles.styleable import StyleArray
from openpyxl.tests.helper import compare_xml
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.xml.functions import tostring


@pytest.fixture
def dimension():
    from openpyxl.worksheet.dimensions import Dimension

    return Dimension


@pytest.fixture
def row_dimension():
    from openpyxl.worksheet.dimensions import RowDimension

    return RowDimension


@pytest.fixture
def column_dimension():
    from openpyxl.worksheet.dimensions import ColumnDimension

    return ColumnDimension


class DummyWorkbook:
    def __init__(self):
        self.shared_styles = IndexedList()
        self._cell_styles = IndexedList()
        self._cell_styles.add(StyleArray())
        self._cell_styles.add(StyleArray([10, 0, 0, 0, 0, 0, 0, 0, 0, 0]))
        self.sheetnames = []


class DummyWorksheet:
    def __init__(self):
        self.parent = DummyWorkbook()


class TestDimension:
    def test_dimension_interface(self, dimension):
        d = dimension(1, True, 1, False, DummyWorksheet())
        assert isinstance(d.parent, DummyWorksheet)
        assert dict(d) == {"hidden": "1", "outlineLevel": "1"}

    def test_invalid_dimension_ctor(self, dimension):
        with pytest.raises(TypeError):
            dimension()

    def test_repr(self, dimension):
        d = dimension(
            worksheet="Sheet1",
            index=1,
            hidden=False,
            outlineLevel=None,
            collapsed=True,
        )
        assert repr(d) == "<Dimension Instance, Attributes={'collapsed': '1'}>"


class TestRowDimension:
    @pytest.mark.parametrize(
        "key, value, expected",
        [
            ("ht", 1, {"ht": "1", "customHeight": "1"}),
            ("thickBot", True, {"thickBot": "1"}),
            ("thickTop", True, {"thickTop": "1"}),
        ],
    )
    def test_row_dimension(self, row_dimension, key, value, expected):
        rd = row_dimension(worksheet=DummyWorksheet())
        setattr(rd, key, value)
        assert dict(rd) == expected

    def test_row_auto_assign(self, row_dimension):
        from openpyxl.worksheet.worksheet import Worksheet

        ws = Worksheet(DummyWorkbook())
        row_info = ws.row_dimensions
        assert isinstance(row_info[1], row_dimension)

    def test_copy(self, row_dimension):
        rd1 = row_dimension(worksheet=DummyWorksheet(), s=[])
        rd2 = copy.copy(rd1)
        assert rd1._style is not rd2._style
        assert dict(rd1) == dict(rd2)


class TestColDimension:
    @pytest.mark.parametrize(
        "key, value, expected",
        [
            ("width", 1, {"width": "1", "customWidth": "1"}),
            ("bestFit", True, {"bestFit": "1", "width": "13", "customWidth": "1"}),
        ],
    )
    def test_col_dimensions(self, column_dimension, key, value, expected):
        cd = column_dimension(worksheet=DummyWorksheet())
        setattr(cd, key, value)
        assert dict(cd) == expected

    def test_column_dimension(self, column_dimension):
        from openpyxl.worksheet.worksheet import Worksheet

        ws = Worksheet(DummyWorkbook())
        cols = ws.column_dimensions
        assert isinstance(cols["A"], column_dimension)

    def test_col_reindex(self, column_dimension):
        cd = column_dimension(DummyWorksheet(), index="D")
        assert dict(cd) == {"customWidth": "1", "width": "13"}
        cd.reindex()
        assert dict(cd) == {"max": "4", "min": "4", "width": "13", "customWidth": "1"}

    def test_col_width(self, column_dimension):
        cd = column_dimension(DummyWorksheet(), index="A", width=4)
        cd.reindex()
        col = cd.to_tree()
        xml = tostring(col)
        expected = '<col width="4" min="1" max="1" customWidth="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_col_style(self, column_dimension):
        from openpyxl.styles.fonts import Font
        from openpyxl.workbook.workbook import Workbook
        from openpyxl.worksheet.worksheet import Worksheet

        ws = Worksheet(Workbook())
        cd = column_dimension(ws, index="A")
        cd.font = Font(color="FF0000")
        cd.reindex()
        col = cd.to_tree()
        xml = tostring(col)
        expected = '<col max="1" min="1" style="1" customWidth="1" width="13"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_outline_cols(self, column_dimension):
        ws = DummyWorksheet()
        cd = column_dimension(ws, index="B", outline_level=1)
        cd.reindex()
        col = cd.to_tree()
        xml = tostring(col)
        expected = '<col max="2" min="2" outlineLevel="1" customWidth="1" width="13"/>'
        diff = compare_xml(expected, xml)
        assert diff is None, diff

    def test_copy(self, column_dimension):
        cd1 = column_dimension(worksheet=DummyWorksheet(), style=[])
        cd2 = copy.copy(cd1)
        assert cd1._style is not cd2._style
        assert dict(cd1) == dict(cd2)

    def test_no_named_style(self, column_dimension):
        cd = column_dimension(worksheet=DummyWorksheet())
        with pytest.raises(AttributeError):
            cd.style = "Normal"

    def test_empty_col(self, column_dimension):
        ws = DummyWorksheet()
        cd = column_dimension(ws, index="C")
        cd.width = 0
        cd.reindex()
        assert cd.to_tree() is None

    def test_range(self, column_dimension):
        ws = DummyWorksheet()
        cd = column_dimension(ws, index="C")
        cd.reindex()
        assert cd.range == "C:C"


class TestGrouping:
    def test_group_columns_simple(self):
        from openpyxl.worksheet.worksheet import Worksheet

        ws = Worksheet(DummyWorkbook())
        dims = ws.column_dimensions
        dims.group("A", "C", 1)
        assert len(dims) == 1
        group = list(dims.values())[0]
        assert group.outline_level == 1
        assert group.range == "A:C"

    def test_group_columns_collapse(self):
        from openpyxl.worksheet.worksheet import Worksheet

        ws = Worksheet(DummyWorkbook())
        dims = ws.column_dimensions
        dims.group("A", "C", 1, hidden=True)
        group = list(dims.values())[0]
        assert group.hidden

    def test_no_cols(self):
        from openpyxl.worksheet.dimensions import DimensionHolder

        dh = DimensionHolder(None)
        node = dh.to_tree()
        assert node is None

    def test_group_rows_simple(self):
        from openpyxl.worksheet.worksheet import Worksheet

        ws = Worksheet(DummyWorkbook())
        dims = ws.row_dimensions
        dims.group(1, 5, 1)
        assert len(dims) == 5
        group = list(dims.values())[0]
        assert group.outline_level == 1

    def test_group_rows_collapse(self):
        from openpyxl.worksheet.worksheet import Worksheet

        ws = Worksheet(DummyWorkbook())
        dims = ws.row_dimensions
        dims.group(1, 10, 1, hidden=True)
        group = list(dims.values())[5]
        assert group.hidden

    def test_no_rows(self):
        from openpyxl.worksheet.dimensions import DimensionHolder

        dh = DimensionHolder(None)
        node = dh.to_tree()
        assert node is None

    def test_to_tree(self):
        from openpyxl.worksheet.worksheet import Worksheet

        ws = Worksheet(DummyWorkbook())
        dims = ws.column_dimensions
        dims["A"].width = 5
        _ = dims["D"]
        assert dims.to_tree() is not None
