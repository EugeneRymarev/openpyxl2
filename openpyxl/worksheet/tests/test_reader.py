# Copyright (c) 2010-2024 openpyxl
import datetime
from io import BytesIO

import pytest
from lxml.etree import fromstring
from lxml.etree import iterparse

from ..formula import ArrayFormula
from ..formula import DataTableFormula
from ..pagebreak import Break
from ..pagebreak import ColBreak
from ..pagebreak import RowBreak
from ..scenario import InputCells
from ..scenario import Scenario
from ..scenario import ScenarioList
from ..worksheet import Worksheet
from openpyxl.cell.rich_text import CellRichText
from openpyxl.cell.rich_text import TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.formula.translate import Translator
from openpyxl.packaging.relationship import Relationship
from openpyxl.packaging.relationship import RelationshipList
from openpyxl.styles.borders import DEFAULT_BORDER
from openpyxl.styles.colors import Color
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.named_styles import NamedStyleList
from openpyxl.styles.styleable import StyleArray
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.xml.constants import SHEET_MAIN_NS


@pytest.mark.parametrize(
    "value, expected",
    [
        ("4.2", 4.2),
        ("-42.000", -42),
        ("0", 0),
        ("0.9999", 0.9999),
        ("99E-02", 0.99),
        ("4", 4),
        ("-1E3", -1000),
        ("1E-3", 0.001),
        ("2e+2", 200.0),
    ],
)
def test_number_convesion(value, expected):
    from .._reader import _cast_number

    assert _cast_number(value) == expected


@pytest.fixture
def Workbook():
    class DummyWorkbook:

        data_only = False
        _colors = []
        encoding = "utf8"
        epoch = CALENDAR_WINDOWS_1900

        def __init__(self):
            self._differential_styles = [DifferentialStyle()] * 5
            self.shared_strings = ["a"] * 30
            self._fonts = IndexedList()
            self._fills = IndexedList()
            self._number_formats = IndexedList()
            self._borders = IndexedList([DEFAULT_BORDER] * 30)
            self._alignments = IndexedList()
            from openpyxl.styles import Protection

            prot_1 = Protection(locked=False)
            prot_2 = Protection(locked=True)
            self._protections = IndexedList()
            self._protections.add(prot_1)
            self._protections.add(prot_2)
            self._cell_styles = IndexedList()
            self._named_styles = NamedStyleList()
            self.vba_archive = None
            for i in range(23):
                self._cell_styles.add((StyleArray([i] * 9)))
            self._cell_styles.add(StyleArray([0, 4, 6, 0, 1, 1, 0, 0, 0]))
            self._cell_styles.add(StyleArray([0, 4, 6, 0, 0, 0, 0, 0, 0]))
            for i in range(25, 29):
                self._cell_styles.add((StyleArray([i] * 9)))
            self._cell_styles.add(StyleArray([0, 4, 6, 0, 0, 1, 0, 0, 0]))
            self.sheetnames = []
            self._date_formats = set()
            self._timedelta_formats = set()

        def create_sheet(self, title):
            return Worksheet(self)

    return DummyWorkbook()


@pytest.fixture
def WorkSheetParser():
    """Setup a parser instance with an empty source"""
    from .._reader import WorkSheetParser

    styles = IndexedList()
    for i in range(29):
        styles.add((StyleArray([i] * 9)))
    styles.add(StyleArray([0, 4, 6, 14, 0, 1, 0, 0, 0]))
    date_formats = set([1, 29, 30])
    timedelta_formats = set([30])
    return WorkSheetParser(
        None,
        {0: "a"},
        date_formats=date_formats,
        timedelta_formats=timedelta_formats,
    )


from warnings import simplefilter

simplefilter("always")


class TestWorksheetParser:

    @pytest.mark.parametrize(
        "filename, expected",
        [
            ("dimension.xml", (4, 1, 27, 30)),
            ("no_dimension.xml", None),
            ("invalid_dimension.xml", (None, 1, None, 113)),
        ],
    )
    def test_read_dimension(self, WorkSheetParser, datadir, filename, expected):
        datadir.chdir()
        parser = WorkSheetParser
        parser.source = filename

        dimension = parser.parse_dimensions()
        assert dimension == expected

    def test_col_width(self, datadir, WorkSheetParser):
        datadir.chdir()
        parser = WorkSheetParser
        with open("complex-styles-worksheet.xml", "rb") as src:
            cols = iterparse(src, tag=f"{{{SHEET_MAIN_NS}}}col")
            for _, col in cols:
                parser.parse_column_dimensions(col)
        assert set(parser.column_dimensions) == set(["A", "C", "E", "I", "G"])
        assert parser.column_dimensions["A"] == {
            "max": "1",
            "min": "1",
            "index": "A",
            "customWidth": "1",
            "width": "31.1640625",
        }

    def test_hidden_col(self, WorkSheetParser):
        parser = WorkSheetParser
        src = '<col min="4" max="4" width="0" hidden="1" customWidth="1"/>'
        element = fromstring(src)

        parser.parse_column_dimensions(element)
        assert "D" in parser.column_dimensions
        assert parser.column_dimensions["D"] == {
            "customWidth": "1",
            "width": "0",
            "hidden": "1",
            "max": "4",
            "min": "4",
            "index": "D",
        }

    def test_styled_col(self, datadir, WorkSheetParser):
        datadir.chdir()
        parser = WorkSheetParser

        with open("complex-styles-worksheet.xml", "rb") as src:
            cols = iterparse(src, tag=f"{{{SHEET_MAIN_NS}}}col")
            for _, col in cols:
                parser.parse_column_dimensions(col)
        assert "I" in parser.column_dimensions
        cd = parser.column_dimensions["I"]
        assert cd == {
            "customWidth": "1",
            "max": "9",
            "min": "9",
            "width": "25",
            "index": "I",
            "style": "28",
        }

    def test_row_dimensions(self, WorkSheetParser):
        src = '<row r="2" spans="1:6" />'
        element = fromstring(src)

        parser = WorkSheetParser
        parser.parse_row(element)

        assert 2 not in parser.row_dimensions

    def test_hidden_row(self, WorkSheetParser):
        parser = WorkSheetParser
        src = '<row r="2" spans="1:4" hidden="1" />'
        element = fromstring(src)
        parser.parse_row(element)
        assert parser.row_dimensions["2"] == {"r": "2", "spans": "1:4", "hidden": "1"}

    def test_styled_row(self, datadir, WorkSheetParser):
        parser = WorkSheetParser
        src = '<row r="23" s="28" spans="1:8" />'
        element = fromstring(src)

        parser.parse_row(element)
        rd = parser.row_dimensions["23"]
        assert rd == {"r": "23", "s": "28", "spans": "1:8"}

    def test_read_row_with_exponent(self, WorkSheetParser):
        parser = WorkSheetParser
        src = '<row r="1.048573e6" spans="1:8" /> '
        element = fromstring(src)
        parser.parse_row(element)
        assert parser.row_counter == 1048573

    def test_invalid_row_number(self, WorkSheetParser):
        parser = WorkSheetParser
        src = '<row r="1.5" spans="1:8" /> '
        element = fromstring(src)
        with pytest.raises(ValueError):
            parser.parse_row(element)

    def test_sheet_protection(self, datadir, WorkSheetParser):
        parser = WorkSheetParser
        src = '<sheetProtection xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" password="DAA7" sheet="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" sort="0" autoFilter="0" pivotTables="0"/>'
        element = fromstring(src)
        parser.parse_sheet_protection(element)

        assert dict(parser.protection) == {
            "autoFilter": "0",
            "deleteColumns": "0",
            "deleteRows": "0",
            "formatCells": "0",
            "formatColumns": "0",
            "formatRows": "0",
            "insertColumns": "0",
            "insertHyperlinks": "0",
            "insertRows": "0",
            "objects": "0",
            "password": "DAA7",
            "pivotTables": "0",
            "scenarios": "0",
            "selectLockedCells": "0",
            "selectUnlockedCells": "0",
            "sheet": "1",
            "sort": "0",
        }

    def test_formula(self, WorkSheetParser):
        parser = WorkSheetParser

        src = """
        <c r="A1" t="str" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <f>IF(TRUE, "y", "n")</f>
            <v>y</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "f",
            "row": 1,
            "style_id": 0,
            "value": '=IF(TRUE, "y", "n")',
        }

    def test_formula_data_only(self, WorkSheetParser):
        parser = WorkSheetParser
        parser.data_only = True

        src = """
        <c r="A1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <f>1+2</f>
            <v>3</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "n",
            "row": 1,
            "style_id": 0,
            "value": 3,
        }

    def test_string_formula_data_only(self, WorkSheetParser):
        parser = WorkSheetParser
        parser.data_only = True

        src = """
        <c r="A1" t="str" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <f>IF(TRUE, "y", "n")</f>
            <v>y</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "s",
            "row": 1,
            "style_id": 0,
            "value": "y",
        }

    def test_number(self, WorkSheetParser):
        parser = WorkSheetParser

        src = """
        <c r="A1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <v>1</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "n",
            "row": 1,
            "style_id": 0,
            "value": 1,
        }

    def test_datetime(self, WorkSheetParser):
        parser = WorkSheetParser

        src = """
        <c r="A1" t="d" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <v>2011-12-25T14:23:55</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "d",
            "row": 1,
            "style_id": 0,
            "value": datetime.datetime(2011, 12, 25, 14, 23, 55),
        }

    def test_timedelta(self, WorkSheetParser):
        parser = WorkSheetParser

        src = """
        <c r="A1" t="n" s="30" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <v>1.25</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "d",
            "row": 1,
            "style_id": 30,
            "value": datetime.timedelta(days=1, hours=6),
        }

    def test_mac_date(self, WorkSheetParser):
        parser = WorkSheetParser
        parser.epoch = CALENDAR_MAC_1904

        src = """
        <c r="A1" t="n" s="29" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <v>41184</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "d",
            "row": 1,
            "style_id": 29,
            "value": datetime.datetime(2016, 10, 3, 0, 0),
        }

    @pytest.mark.parametrize("value", [-693595, 2958466])
    def test_out_of_range_datetime(self, WorkSheetParser, recwarn, value):
        parser = WorkSheetParser
        src = f"""
        <c r="A1" t="n" s="29" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <v>{value}</v>
        </c>
        """
        element = fromstring(src)

        parser.parse_cell(element)
        w = recwarn.pop()
        assert issubclass(w.category, UserWarning)

    def test_string(self, WorkSheetParser):
        parser = WorkSheetParser

        src = """
        <c r="A1" t="s" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <v>0</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "s",
            "row": 1,
            "style_id": 0,
            "value": "a",
        }

    def test_boolean(self, WorkSheetParser):
        parser = WorkSheetParser

        src = """
        <c r="A1" t="b" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <v>1</v>
        </c>
        """
        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "b",
            "row": 1,
            "style_id": 0,
            "value": True,
        }

    def test_inline_string(self, WorkSheetParser):
        parser = WorkSheetParser
        src = """
        <c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="A1" s="0" t="inlineStr">
          <is>
          <t>ID</t>
          </is>
        </c>
        """

        element = fromstring(src)

        cell = parser.parse_cell(element)
        assert cell == {
            "column": 1,
            "data_type": "s",
            "row": 1,
            "style_id": 0,
            "value": "ID",
        }

    def test_inline_richtext(self, WorkSheetParser):
        parser = WorkSheetParser
        parser.rich_text = True
        src = """
        <c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="R2" s="4" t="inlineStr">
        <is>
          <r>
            <rPr>
              <sz val="8.0" />
            </rPr>
            <t xml:space="preserve">11 de September de 2014</t>
          </r>
          </is>
        </c>
        """

        element = fromstring(src)
        cell = parser.parse_cell(element)
        expected = CellRichText(
            TextBlock(
                font=InlineFont(sz="8.0"),
                text="11 de September de 2014",
            )
        )
        assert cell == {
            "column": 18,
            "data_type": "s",
            "row": 2,
            "style_id": 4,
            "value": expected,
        }

    def test_parse_richtext(self):
        from .._reader import parse_richtext_string

        src = """
        <is>
          <r>
            <rPr>
              <sz val="8.0" />
            </rPr>
            <t xml:space="preserve">11 de September de 2014</t>
          </r>
          </is>
        """
        element = fromstring(src)
        value = parse_richtext_string(element)
        assert value == CellRichText(
            TextBlock(font=InlineFont(sz="8.0"), text="11 de September de 2014")
        )

    def test_sheet_views(self, WorkSheetParser):
        parser = WorkSheetParser

        src = b"""
        <sheetViews xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetView tabSelected="1" zoomScale="200" zoomScaleNormal="200" zoomScalePageLayoutView="200" workbookViewId="0">
            <pane xSplit="5" ySplit="19" topLeftCell="F20" activePane="bottomRight" state="frozenSplit"/>
            <selection pane="topRight" activeCell="F1" sqref="F1"/>
            <selection pane="bottomLeft" activeCell="A20" sqref="A20"/>
            <selection pane="bottomRight" activeCell="E22" sqref="E22"/>
          </sheetView>
        </sheetViews>
        """

        parser.source = BytesIO(src)
        for _ in parser.parse():
            pass

        view = parser.views.sheetView[0]

        assert view.zoomScale == 200
        assert len(view.selection) == 3

    def test_shared_formula(self, WorkSheetParser):
        parser = WorkSheetParser
        src = """
        <c r="A9" t="str" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <f t="shared" si="0"/>
          <v>9</v>
        </c>
        """
        element = fromstring(src)
        parser.shared_formulae["0"] = Translator("=A4*B4", "A1")
        formula = parser.parse_formula(element)
        assert formula == "=A12*B12"

    def test_array_formula(self, WorkSheetParser, datadir):
        parser = WorkSheetParser

        src = """
        <c r="C10" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <f t="array" ref="C10:C14">SUM(A10:A14*B10:B14)</f>
        </c>"""
        element = fromstring(src)

        formula = parser.parse_formula(element)
        assert isinstance(formula, ArrayFormula)
        assert formula.ref == "C10:C14"
        assert formula.text == "=SUM(A10:A14*B10:B14)"

    def test_table_formula(self, WorkSheetParser):
        parser = WorkSheetParser
        src = """
        <c r="C9" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <f t="dataTable" ref="C9:C24" dt2D="0" dtr="0" r1="C4"/>
          <v>1</v>
       </c>"""
        element = fromstring(src)
        formula = parser.parse_formula(element)
        assert isinstance(formula, DataTableFormula)

    def test_extended_conditional_formatting(self, WorkSheetParser, recwarn):
        parser = WorkSheetParser

        src = """
        <extLst>
            <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                <x14:conditionalFormattings>
                    <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                        <x14:cfRule type="containsText" priority="1" operator="containsText" id="{5F3BF262-1B4E-49BC-812D-8F5B1802CE0A}">
                            <xm:f>NOT(ISERROR(SEARCH($AA$7,A1)))</xm:f>
                            <xm:f>$AA$7</xm:f>
                            <x14:dxf>
                                <fill>
                                    <patternFill>
                                        <bgColor rgb="FF6699FF"/>
                                    </patternFill>
                                </fill>
                            </x14:dxf>
                        </x14:cfRule>
                    </x14:conditionalFormatting>
                </x14:conditionalFormattings>
            </ext>
        </extLst>
        """
        element = fromstring(src)

        parser.parse_extensions(element)
        w = recwarn.pop()
        assert issubclass(w.category, UserWarning)

    def test_bad_conditional_format_rule(self, WorkSheetParser, recwarn):
        parser = WorkSheetParser

        src = """
                <conditionalFormatting sqref="G124:G144">
                    <cfRule type="colorScale" priority="1">
                        <colorScale>
                            <cfvo type="num" val="&quot;&lt;&gt;100%&quot;"/>
                            <cfvo type="max"/>
                            <color rgb="FFFF7128"/>
                            <color rgb="FFFFEF9C"/>
                        </colorScale>
                    </cfRule>
                </conditionalFormatting>
                """
        element = fromstring(src)

        parser.parse_formatting(element)
        w = recwarn.pop()
        assert issubclass(w.category, UserWarning)

    def test_cell_without_coordinates(self, WorkSheetParser):
        parser = WorkSheetParser

        src = """
        <row r="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c t="s">
            <v>2</v>
          </c>
          <c t="s">
            <v>4</v>
          </c>
          <c t="s">
            <v>3</v>
          </c>
          <c t="s">
            <v>6</v>
          </c>
          <c t="s">
            <v>9</v>
          </c>
        </row>
        """
        element = fromstring(src)

        parser.shared_strings = ["Whatever"] * 10
        parser.parse_row(element)

        assert parser.row_counter == 1
        assert parser.col_counter == 5

    def test_row_and_cell_without_coordinates(self, WorkSheetParser):
        parser = WorkSheetParser
        src = """
        <row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c>
            <v>2</v>
          </c>
          <c>
            <v>4</v>
          </c>
          <c>
            <v>3</v>
          </c>
        </row>
        """
        element = fromstring(src)
        max_row, cells = parser.parse_row(element)
        expected = [
            {"column": 1, "row": 1, "data_type": "n", "value": 2, "style_id": 0},
            {"column": 2, "row": 1, "data_type": "n", "value": 4, "style_id": 0},
            {"column": 3, "row": 1, "data_type": "n", "value": 3, "style_id": 0},
        ]
        for expected_cell, cell in zip(expected, cells):
            assert expected_cell == cell

    def test_row_and_cell_skipping_coordinates(self, WorkSheetParser):
        parser = WorkSheetParser
        src = """
        <row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c>
            <v>1</v>
          </c>
          <c r="D1">
            <v>2</v>
          </c>
          <c>
            <v>3</v>
          </c>
          <c r="G1">
            <v>4</v>
          </c>
        </row>
        """
        element = fromstring(src)
        _, cells = parser.parse_row(element)
        expected = [
            {"column": 1, "row": 1, "data_type": "n", "value": 1, "style_id": 0},
            {"column": 4, "row": 1, "data_type": "n", "value": 2, "style_id": 0},
            {"column": 5, "row": 1, "data_type": "n", "value": 3, "style_id": 0},
            {"column": 7, "row": 1, "data_type": "n", "value": 4, "style_id": 0},
        ]
        assert len(cells) == len(expected)
        for expected_cell, cell in zip(expected, cells):
            assert expected_cell == cell

    def test_second_row_cell_index_without_coordinates(self, WorkSheetParser):
        parser = WorkSheetParser
        src = """
        <row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c>
            <v>2</v>
          </c>
        </row>
        """
        element = fromstring(src)
        parser.parse_row(element)
        max_row, cells = parser.parse_row(element)
        expected = [
            {"column": 1, "row": 2, "data_type": "n", "value": 2, "style_id": 0},
        ]
        for expected_cell, cell in zip(expected, cells):
            assert expected_cell == cell

    def test_external_hyperlinks(self, WorkSheetParser):
        src = b"""
        <hyperlinks xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           display="http://test.com" r:id="rId1" ref="A1"/>
        </hyperlinks>
        """
        parser = WorkSheetParser
        parser.source = BytesIO(src)

        for _ in parser.parse():
            pass

        parser.parse()
        assert len(parser.hyperlinks.hyperlink) == 1

    def test_local_hyperlinks(self, WorkSheetParser):
        src = b"""
         <hyperlinks xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <hyperlink ref="B4:B7" location="'STP nn000TL-10, PKG 2.52'!A1" display="STP 10000TL-10"/>
        </hyperlinks>
        """
        parser = WorkSheetParser
        parser.source = BytesIO(src)

        for _ in parser.parse():
            pass

        assert len(parser.hyperlinks.hyperlink) == 1

    def test_merge_cells(self, WorkSheetParser):
        src = b"""
        <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <mergeCells>
            <mergeCell ref="C2:F2"/>
            <mergeCell ref="B19:C20"/>
            <mergeCell ref="E19:G19"/>
          </mergeCells>
        </sheet>
        """

        parser = WorkSheetParser
        parser.source = BytesIO(src)

        for _ in parser.parse():
            pass

        assert len(parser.merged_cells.mergeCell) == 3

    def test_conditonal_formatting(self, WorkSheetParser):
        src = b"""
        <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <conditionalFormatting sqref="S1:S10">
            <cfRule type="top10" dxfId="25" priority="12" percent="1" rank="10"/>
        </conditionalFormatting>
        <conditionalFormatting sqref="T1:T10">
          <cfRule type="top10" dxfId="24" priority="11" bottom="1" rank="4"/>
        </conditionalFormatting>
        </sheet>
        """
        parser = WorkSheetParser
        parser.source = BytesIO(src)

        for _ in parser.parse():
            pass
        assert len(parser.formatting) == 2

    def test_sheet_properties(self, WorkSheetParser):
        src = b"""
        <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetPr codeName="Sheet3">
          <tabColor rgb="FF92D050"/>
          <outlinePr summaryBelow="1" summaryRight="1"/>
          <pageSetUpPr/>
        </sheetPr>
        </sheet>
        """
        parser = WorkSheetParser
        parser.source = BytesIO(src)

        for _ in parser.parse():
            pass

        assert parser.sheet_properties.tabColor.rgb == "FF92D050"
        assert parser.sheet_properties.codeName == "Sheet3"

    def test_sheet_format(self, WorkSheetParser):

        src = b"""
        <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetFormatPr defaultRowHeight="14.25" baseColWidth="15"/>
        </sheet>
        """
        parser = WorkSheetParser
        parser.source = BytesIO(src)

        for _ in parser.parse():
            pass

        assert parser.sheet_format.defaultRowHeight == 14.25
        assert parser.sheet_format.baseColWidth == 15

    def test_tables(self, WorkSheetParser):
        src = b"""
        <tableParts count="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <tablePart r:id="rId1"/>
        </tableParts>
        """
        parser = WorkSheetParser
        parser.source = BytesIO(src)

        for _ in parser.parse():
            pass

        assert parser.tables.tablePart[0].id == "rId1"

    def test_auto_filter(self, WorkSheetParser):
        src = b"""
        <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <autoFilter ref="A1:AK3237">
            <sortState ref="A2:AM3269">
                <sortCondition ref="B1:B3269"/>
            </sortState>
          </autoFilter>
        </sheet>
        """

        parser = WorkSheetParser
        parser.source = BytesIO(src)
        for _ in parser.parse():
            pass

        assert parser.auto_filter.ref == "A1:AK3237"
        assert parser.auto_filter.sortState.ref == "A2:AM3269"

    def test_page_row_break(self, WorkSheetParser):
        src = b"""
        <rowBreaks count="1" manualBreakCount="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
           <brk id="15" man="1" max="16383" min="0"/>
        </rowBreaks>
        """
        expected_pagebreak = RowBreak()
        expected_pagebreak.append(Break(id=15))

        parser = WorkSheetParser
        el = fromstring(src)
        parser.parse_row_breaks(el)

        assert parser.row_breaks == expected_pagebreak

    def test_col_break(self, WorkSheetParser):
        src = b"""
        <colBreaks count="1" manualBreakCount="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <brk id="15" man="1" max="16383" min="0"/>
        </colBreaks>
        """
        expected_pagebreak = ColBreak()
        expected_pagebreak.append(Break(id=15))

        parser = WorkSheetParser
        el = fromstring(src)
        parser.parse_col_breaks(el)

        assert parser.col_breaks == expected_pagebreak

    def test_scenarios(self, WorkSheetParser):
        src = b"""
        <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <scenarios current="0" show="0">
        <scenario name="Worst case" locked="1" user="User" comment="comment">
          <inputCells r="B2" val="50000" />
        </scenario>
        </scenarios>
        </sheet>
        """

        parser = WorkSheetParser
        parser.source = BytesIO(src)
        for _ in parser.parse():
            pass

        c = InputCells(r="B2", val="50000")
        s = Scenario(
            name="Worst case",
            inputCells=[c],
            locked=True,
            user="User",
            comment="comment",
        )
        scenarios = ScenarioList(scenario=[s], current="0", show="0")

        assert parser.scenarios == scenarios

    def test_legacy(self, WorkSheetParser):
        src = """
        <legacyDrawing r:id="rId3" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        parser = WorkSheetParser
        element = fromstring(src)

        parser.parse_legacy(element)
        assert parser.legacy_drawing == "rId3"

    def test_custom_views_breaks(self, WorkSheetParser):
        src = b"""
        <sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <customSheetViews >
            <customSheetView scale="70">
            <selection activeCell="F14" sqref="F14"/>
            <rowBreaks count="1" manualBreakCount="1">
              <brk id="1" max="16383" man="1"/>
            </rowBreaks>
            <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
            <pageSetup paperSize="9" scale="50" firstPageNumber="99" orientation="landscape" useFirstPageNumber="1"/>
            <headerFooter alignWithMargins="0"/>
            </customSheetView>
        </customSheetViews>
        </sheet>
        """
        parser = WorkSheetParser
        parser.source = BytesIO(src)
        for _ in parser.parse():
            pass
        assert parser.row_breaks == RowBreak()


@pytest.fixture
def WorksheetReader():
    from .._reader import WorksheetReader

    return WorksheetReader


@pytest.fixture
def PrimedWorksheetReader(Workbook, WorksheetReader, datadir):
    wb = Workbook
    ws = wb.create_sheet("Sheet")
    datadir.chdir()
    src = "complex-styles-worksheet.xml"
    reader = WorksheetReader(
        ws,
        src,
        wb.shared_strings,
        data_only=False,
        rich_text=False,
    )
    return reader


class TestWorksheetReader:

    def test_cell(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        assert ws["C1"].value == "a"

    def test_array_formula(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        assert ws["E2"].value.text == "=C2:C11*D2:D11"

    def test_formatting(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        reader.bind_formatting()

        fmts = ws.conditional_formatting
        assert fmts["T1:T10"][-1].dxf == DifferentialStyle()
        assert len(fmts["A1"]) == 2

    def test_merged(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        reader.bind_merged_cells()

        assert ws.merged_cells == "G18:H18 G23:H24 A18:B18"

    @pytest.mark.parametrize(
        argnames="input_coordinate,expected_coordinate",
        argvalues=[("H18", "G18"), ("G18", "G18"), ("I18", None), ("H23", "G23")],
    )
    def test_normalize_merged_cell_link(
        self,
        PrimedWorksheetReader,
        input_coordinate,
        expected_coordinate,
    ):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        reader.bind_merged_cells()

        normalized = reader.normalize_merged_cell_link(input_coordinate)

        if expected_coordinate is None:
            assert normalized is None
        else:
            assert normalized.coordinate == expected_coordinate

    def test_external_hyperlinks(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        r = Relationship(type="hyperlink", Id="rId1", Target="../")
        rels = RelationshipList()
        rels.append(r)
        ws._rels = rels

        reader.bind_hyperlinks()

        assert ws["A1"].hyperlink.target == "../"

    def test_internal_hyperlinks(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        r = Relationship(type="hyperlink", Id="rId1", Target="../")
        rels = RelationshipList()
        rels.append(r)
        ws._rels = rels

        reader.bind_hyperlinks()

        assert ws["B4"].hyperlink.location == "'STP nn000TL-10, PKG 2.52'!A1"
        assert ws["B4"].hyperlink.ref == "B4"

    def test_merged_hyperlinks(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        r = Relationship(type="hyperlink", Id="rId1", Target="../")
        rels = RelationshipList()
        rels.append(r)
        ws._rels = rels

        reader.bind_merged_cells()
        reader.bind_hyperlinks()

        assert ws.merged_cells == "G18:H18 G23:H24 A18:B18"
        assert ws["A18"].hyperlink.display == "http://test.com"
        assert ws["B18"].hyperlink is None

        # Link referencing H24 should be placed on G23 because H24 is a merged cell
        # and G23 is the top-left cell in the merged range
        assert ws["G23"].hyperlink.tooltip == "openpyxl"
        assert ws["H24"].hyperlink is None

    def test_tables(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        r = Relationship(type="table", Id="rId1", Target="../tables/table1.xml")
        rels = RelationshipList()
        rels.append(r)
        ws._rels = rels

        reader.bind_tables()

        assert reader.tables == ["../tables/table1.xml"]

    def test_cols(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        reader.bind_col_dimensions()

        assert len(ws.column_dimensions) == 5
        assert ws.column_dimensions["I"].parent == ws

    def test_rows(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        reader.bind_row_dimensions()

        assert len(ws.row_dimensions) == 7
        assert ws.row_dimensions[4].parent == ws

    def test_properties(self, PrimedWorksheetReader):
        reader = PrimedWorksheetReader
        reader.bind_cells()
        ws = reader.ws

        reader.bind_properties()

        properties = (
            "page_margins",
            "page_setup",
            "views",
            "sheet_format",
            "legacy_drawing",
            "protection",
        )
        for k in properties:
            assert getattr(ws, k) == getattr(reader.parser, k)

    def test_more_rows_than_cells(self, Workbook, WorksheetReader, datadir):
        ws = Workbook.create_sheet("Sheet")
        datadir.chdir()
        reader = WorksheetReader(ws, "more_rows_than_cells.xml", None, None, False)
        reader.bind_cells()
        assert ws._current_row == 3
