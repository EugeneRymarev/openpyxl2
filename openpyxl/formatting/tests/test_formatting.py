# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.reader.excel import load_workbook
from openpyxl.styles.borders import Border
from openpyxl.styles.borders import Side
from openpyxl.styles.colors import Color
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.differential import DifferentialStyleList
from openpyxl.styles.fills import PatternFill
from openpyxl.styles.fonts import Font
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def conditional_formatting():
    from ..formatting import ConditionalFormatting

    return ConditionalFormatting


class DummyWorkbook:
    def __init__(self):
        self._differential_styles = DifferentialStyleList()
        self.worksheets = []


class DummyWorksheet:
    def __init__(self):
        self.conditional_formatting = ConditionalFormattingList()
        self.parent = DummyWorkbook()


def test_conditional_formatting_read(datadir):
    datadir.chdir()
    reference_file = "conditional-formatting.xlsx"
    wb = load_workbook(reference_file)
    ws = wb.active
    rules = ws.conditional_formatting
    assert len(rules) == 30
    # First test the conditional formatting rules read
    rule = rules["A1:A1048576"][0]
    assert dict(rule) == {"priority": "30", "type": "colorScale"}
    rule = rules["B1:B10"][0]
    assert dict(rule) == {"priority": "29", "type": "colorScale"}
    rule = rules["C1:C10"][0]
    assert dict(rule) == {"priority": "28", "type": "colorScale"}
    rule = rules["D1:D10"][0]
    assert dict(rule) == {"priority": "27", "type": "colorScale"}
    rule = rules["E1:E10"][0]
    assert dict(rule) == {"priority": "26", "type": "colorScale"}
    rule = rules["F1:F10"][0]
    assert dict(rule) == {"priority": "25", "type": "colorScale"}
    rule = rules["G1:G10"][0]
    assert dict(rule) == {"priority": "24", "type": "colorScale"}
    rule = rules["H1:H10"][0]
    assert dict(rule) == {"priority": "23", "type": "colorScale"}
    rule = rules["I1:I10"][0]
    assert dict(rule) == {"priority": "22", "type": "colorScale"}
    rule = rules["J1:J10"][0]
    assert dict(rule) == {"priority": "21", "type": "colorScale"}
    rule = rules["K1:K10"][0]
    assert dict(rule) == {"priority": "20", "type": "dataBar"}
    rule = rules["L1:L10"][0]
    assert dict(rule) == {"priority": "19", "type": "dataBar"}
    rule = rules["M1:M10"][0]
    assert dict(rule) == {"priority": "18", "type": "dataBar"}
    rule = rules["N1:N10"][0]
    assert dict(rule) == {"priority": "17", "type": "iconSet"}
    rule = rules["O1:O10"][0]
    assert dict(rule) == {"priority": "16", "type": "iconSet"}
    rule = rules["P1:P10"][0]
    assert dict(rule) == {"priority": "15", "type": "iconSet"}
    rule = rules["Q1:Q10"][0]
    expected = {
        "text": "3",
        "priority": "14",
        "dxfId": "27",
        "operator": "containsText",
        "type": "containsText",
    }
    assert dict(rule) == expected
    expected = DifferentialStyle(
        font=Font(color="FF9C0006"),
        fill=PatternFill(bgColor="FFFFC7CE"),
    )
    assert rule.dxf == expected
    rule = rules["R1:R10"][0]
    expected = {
        "operator": "between",
        "dxfId": "26",
        "type": "cellIs",
        "priority": "13",
    }
    assert dict(rule) == expected
    expected = DifferentialStyle(
        font=Font(color="FF9C6500"),
        fill=PatternFill(bgColor="FFFFEB9C"),
    )
    assert rule.dxf == expected
    rule = rules["S1:S10"][0]
    expected = {
        "priority": "12",
        "dxfId": "25",
        "percent": "1",
        "type": "top10",
        "rank": "10",
    }
    assert dict(rule) == expected
    rule = rules["T1:T10"][0]
    expected = {
        "priority": "11",
        "dxfId": "24",
        "type": "top10",
        "rank": "4",
        "bottom": "1",
    }
    assert dict(rule) == expected
    rule = rules["U1:U10"][0]
    assert dict(rule) == {"priority": "10", "dxfId": "23", "type": "aboveAverage"}
    rule = rules["V1:V10"][0]
    expected = {
        "aboveAverage": "0",
        "dxfId": "22",
        "type": "aboveAverage",
        "priority": "9",
    }
    assert dict(rule) == expected
    rule = rules["W1:W10"][0]
    expected = {
        "priority": "8",
        "dxfId": "21",
        "type": "aboveAverage",
        "equalAverage": "1",
    }
    assert dict(rule) == expected
    rule = rules["X1:X10"][0]
    expected = {
        "aboveAverage": "0",
        "dxfId": "20",
        "priority": "7",
        "type": "aboveAverage",
        "equalAverage": "1",
    }
    assert dict(rule) == expected
    rule = rules["Y1:Y10"][0]
    expected = {"priority": "6", "dxfId": "19", "type": "aboveAverage", "stdDev": "1"}
    assert dict(rule) == expected
    rule = rules["Z1:Z10"][0]
    expected = {
        "aboveAverage": "0",
        "dxfId": "18",
        "type": "aboveAverage",
        "stdDev": "1",
        "priority": "5",
    }
    assert dict(rule) == expected
    expected = DifferentialStyle(
        font=Font(b=True, i=True, color="FF9C0006"),
        fill=PatternFill(bgColor="FFFFC7CE"),
        border=Border(
            left=Side(style="thin", color=Color(theme=5)),
            right=Side(style="thin", color=Color(theme=5)),
            top=Side(style="thin", color=Color(theme=5)),
            bottom=Side(style="thin", color=Color(theme=5)),
        ),
    )
    assert rule.dxf == expected
    rule = rules["AA1:AA10"][0]
    expected = {"priority": "4", "dxfId": "17", "type": "aboveAverage", "stdDev": "2"}
    assert dict(rule) == expected
    rule = rules["AB1:AB10"][0]
    assert dict(rule) == {"priority": "3", "dxfId": "16", "type": "duplicateValues"}
    rule = rules["AC1:AC10"][0]
    assert dict(rule) == {"priority": "2", "dxfId": "15", "type": "uniqueValues"}
    rule = rules["AD1:AD10"][0]
    expected = {"priority": "1", "dxfId": "14", "type": "expression"}
    assert dict(rule) == expected


class TestConditionalFormatting:
    def test_ctor(self, conditional_formatting):
        cf = conditional_formatting(sqref="A1:B5")
        xml = tostring(cf.to_tree())
        expected = '<conditionalFormatting sqref="A1:B5"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_tree(self, conditional_formatting):
        src = '<conditionalFormatting sqref="A1:B5"/>'
        tree = fromstring(src)
        cf = conditional_formatting.from_tree(tree)
        assert cf.sqref == "A1:B5"

    def test_eq(self, conditional_formatting):
        c1 = conditional_formatting("A1:B5")
        c2 = conditional_formatting("A1:B5", pivot=True)
        assert c1 == c2

    def test_hash(self, conditional_formatting):
        c1 = conditional_formatting("A1:B5")
        assert hash(c1) == hash("A1:B5")

    def test_repr(self, conditional_formatting):
        c1 = conditional_formatting("A1:B5")
        assert repr(c1) == "<ConditionalFormatting A1:B5>"

    def test_contains(self, conditional_formatting):
        c2 = conditional_formatting("A1:A5 B1:B5")
        assert "B2" in c2
