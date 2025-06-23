# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.styles.colors import Color
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def format_object():
    from openpyxl.formatting.rule import FormatObject

    return FormatObject


@pytest.fixture
def color_scale():
    from openpyxl.formatting.rule import ColorScale

    return ColorScale


@pytest.fixture
def color_scale_rule():
    from openpyxl.formatting.rule import ColorScaleRule

    return ColorScaleRule


@pytest.fixture
def data_bar():
    from openpyxl.formatting.rule import DataBar

    return DataBar


@pytest.fixture
def icon_set():
    from openpyxl.formatting.rule import IconSet

    return IconSet


@pytest.fixture
def rule():
    from openpyxl.formatting.rule import Rule

    return Rule


class TestFormatObject:
    def test_create(self, format_object):
        xml = fromstring('<cfvo type="num" val="3"/>')
        cfvo = format_object.from_tree(xml)
        assert cfvo.type == "num"
        assert cfvo.val == 3
        assert cfvo.gte is None

    def test_serialise(self, format_object):
        cfvo = format_object(type="percent", val=4)
        xml = tostring(cfvo.to_tree())
        expected = '<cfvo type="percent" val="4"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.parametrize(
        "typ, value, expected",
        [
            ("num", "5", 5.0),
            ("percent", "70", 70),
            ("max", 10, 10),
            ("min", "4.2", 4.2),
            ("formula", "=A2*4", "=A2*4"),
            ("percentile", 10, 10),
            ("formula", None, None),
        ],
    )
    def test_value_types(self, format_object, typ, value, expected):
        cfvo = format_object(type=typ, val=value)
        assert cfvo.val == expected

    def test_cell_reference(self, format_object):
        cfvo = format_object(type="num")
        cfvo.val = "$K$6"
        assert cfvo.val == "$K$6"


class TestColorScale:
    def test_create(self, color_scale):
        src = """
        <colorScale
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="FFFF7128"/>
            <color rgb="FFFFEF9C"/>
        </colorScale>
        """
        xml = fromstring(src)
        cs = color_scale.from_tree(xml)
        assert len(cs.cfvo) == 2
        assert len(cs.color) == 2

    def test_serialise(self, color_scale, format_object):
        fo1 = format_object(type="min", val="0")
        fo2 = format_object(type="percent", val="50")
        fo3 = format_object(type="max", val="0")
        col1 = Color(rgb="FFFF0000")
        col2 = Color(rgb="FFFFFF00")
        col3 = Color(rgb="FF00B050")
        cs = color_scale(cfvo=[fo1, fo2, fo3], color=[col1, col2, col3])
        xml = tostring(cs.to_tree())
        expected = """
        <colorScale>
            <cfvo type="min" val="0"/>
            <cfvo type="percent" val="50"/>
            <cfvo type="max" val="0"/>
            <color rgb="FFFF0000"/>
            <color rgb="FFFFFF00"/>
            <color rgb="FF00B050"/>
        </colorScale>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_two_colors(self, color_scale_rule):
        cfRule = color_scale_rule(
            start_type="min",
            start_value=None,
            start_color="FFAA0000",
            end_type="max",
            end_value=None,
            end_color="FF00AA00",
        )
        xml = tostring(cfRule.to_tree())
        expected = """
        <cfRule priority="0" type="colorScale">
            <colorScale>
                <cfvo type="min"/>
                <cfvo type="max"/>
                <color rgb="FFAA0000"/>
                <color rgb="FF00AA00"/>
            </colorScale>
        </cfRule>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_three_colors(self, color_scale_rule):
        cfRule = color_scale_rule(
            start_type="percentile",
            start_value=10,
            start_color="FFAA0000",
            mid_type="percentile",
            mid_value=50,
            mid_color="FF0000AA",
            end_type="percentile",
            end_value=90,
            end_color="FF00AA00",
        )
        xml = tostring(cfRule.to_tree())
        expected = """
        <cfRule priority="0" type="colorScale">
            <colorScale>
                <cfvo type="percentile" val="10"></cfvo>
                <cfvo type="percentile" val="50"></cfvo>
                <cfvo type="percentile" val="90"></cfvo>
                <color rgb="FFAA0000"></color>
                <color rgb="FF0000AA"></color>
                <color rgb="FF00AA00"></color>
            </colorScale>
        </cfRule>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestDataBar:
    def test_create(self, data_bar):
        src = """
        <dataBar xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="FF638EC6"/>
        </dataBar>
        """
        xml = fromstring(src)
        db = data_bar.from_tree(xml)
        assert len(db.cfvo) == 2
        assert db.color.value == "FF638EC6"

    def test_serialise(self, data_bar, format_object):
        fo1 = format_object(type="min", val="0")
        fo2 = format_object(type="percent", val="50")
        db = data_bar(
            minLength=4,
            maxLength=10,
            cfvo=[fo1, fo2],
            color="FF2266",
            showValue=True,
        )
        xml = tostring(db.to_tree())
        expected = """
        <dataBar maxLength="10" minLength="4" showValue="1">
            <cfvo type="min" val="0"></cfvo>
            <cfvo type="percent" val="50"></cfvo>
            <color rgb="00FF2266"></color>
        </dataBar>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestIconSet:
    def test_create(self, icon_set):
        src = """
        <iconSet
                iconSet="5Rating"
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cfvo type="percent" val="0"/>
            <cfvo type="percentile" val="20"/>
            <cfvo type="percentile" val="40"/>
            <cfvo type="percentile" val="60"/>
            <cfvo type="percentile" val="80"/>
        </iconSet>
        """
        xml = fromstring(src)
        icon = icon_set.from_tree(xml)
        assert icon.iconSet == "5Rating"
        assert len(icon.cfvo) == 5

    def test_serialise(self, icon_set, format_object):
        fo1 = format_object(type="num", val="2")
        fo2 = format_object(type="num", val="4")
        fo3 = format_object(type="num", val="6")
        fo4 = format_object(type="percent", val="0")
        icon = icon_set(
            cfvo=[fo1, fo2, fo3, fo4],
            iconSet="4ArrowsGray",
            reverse=True,
            showValue=False,
        )
        xml = tostring(icon.to_tree())
        expected = """
        <iconSet iconSet="4ArrowsGray" showValue="0" reverse="1">
            <cfvo type="num" val="2"/>
            <cfvo type="num" val="4"/>
            <cfvo type="num" val="6"/>
            <cfvo type="percent" val="0"/>
        </iconSet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestRule:
    def test_create(self, rule, datadir):
        datadir.chdir()
        with open("worksheet.xml") as src:
            xml = fromstring(src.read())
        rules = []
        s = f"{{{SHEET_MAIN_NS}}}conditionalFormatting/{{{SHEET_MAIN_NS}}}cfRule"
        for el in xml.findall(s):
            rules.append(rule.from_tree(el))
        assert len(rules) == 30
        assert rules[17].formula == ["2", "7"]
        assert rules[-1].formula == ["AD1>3"]

    def test_serialise(self, rule):
        r = rule(type="cellIs", dxfId="26", priority="13", operator="between")
        r.formula = ["2", "7"]
        xml = tostring(r.to_tree())
        expected = """
        <cfRule type="cellIs" dxfId="26" priority="13" operator="between">
            <formula>2</formula>
            <formula>7</formula>
        </cfRule>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_non_ascii_formula(self, rule):
        r = rule(
            type="cellIs",
            priority=10,
            formula=[b"D\xc3\xbcsseldorf".decode("utf-8")],
        )
        xml = tostring(r.to_tree())
        expected = b"""
        <cfRule priority="10" type="cellIs">
            <formula>D\xc3\xbcsseldorf</formula>
        </cfRule>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


def test_formula_rule():
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles.differential import DifferentialStyle

    cf = FormulaRule(formula=["ISBLANK(C1)"], stopIfTrue=True)
    assert dict(cf) == {"priority": "0", "stopIfTrue": "1", "type": "expression"}
    assert cf.formula == ["ISBLANK(C1)"]
    assert cf.dxf == DifferentialStyle()


def test_cell_is_rule():
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles.fills import PatternFill

    red_fill = PatternFill(
        start_color="FFEE1111",
        end_color="FFEE1111",
        fill_type="solid",
    )
    r = CellIsRule(operator="<", formula=["C$1"], stopIfTrue=True, fill=red_fill)
    expected = {
        "operator": "lessThan",
        "priority": "0",
        "type": "cellIs",
        "stopIfTrue": "1",
    }
    assert dict(r) == expected
    assert r.formula == ["C$1"]
    assert r.dxf.fill == red_fill


@pytest.mark.parametrize(
    "value, expansion",
    [
        ("<=", "lessThanOrEqual"),
        (">", "greaterThan"),
        ("!=", "notEqual"),
        ("=", "equal"),
        (">=", "greaterThanOrEqual"),
        ("==", "equal"),
        ("<", "lessThan"),
    ],
)
def test_operator_expansion(value, expansion):
    from openpyxl.formatting.rule import CellIsRule

    cf1 = CellIsRule(operator=value, formula=[])
    cf2 = CellIsRule(operator=expansion, formula=[])
    assert cf1.operator == expansion
    assert cf2.operator == expansion


def test_iconset_rule():
    from openpyxl.formatting.rule import IconSetRule

    r = IconSetRule("5Arrows", "percent", [10, 20, 30, 40, 50])
    xml = tostring(r.to_tree())
    expected = """
    <cfRule priority="0" type="iconSet">
        <iconSet iconSet="5Arrows">
            <cfvo type="percent" val="10"/>
            <cfvo type="percent" val="20"/>
            <cfvo type="percent" val="30"/>
            <cfvo type="percent" val="40"/>
            <cfvo type="percent" val="50"/>
        </iconSet>
    </cfRule>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_databar_rule():
    from openpyxl.formatting.rule import DataBarRule

    r = DataBarRule(
        start_type="percentile",
        start_value=10,
        end_type="percentile",
        end_value="90",
        color="FF638EC6",
    )
    xml = tostring(r.to_tree())
    expected = """
    <cfRule type="dataBar" priority="0">
        <dataBar>
            <cfvo type="percentile" val="10"/>
            <cfvo type="percentile" val="90"/>
            <color rgb="FF638EC6"/>
        </dataBar>
    </cfRule>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
