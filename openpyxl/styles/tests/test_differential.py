# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.styles.borders import Border
from openpyxl.styles.borders import Side
from openpyxl.styles.colors import Color
from openpyxl.styles.fills import PatternFill
from openpyxl.styles.fonts import Font
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def differential_style():
    from openpyxl.styles.differential import DifferentialStyle

    return DifferentialStyle


@pytest.fixture
def differential_style_list():
    from openpyxl.styles.differential import DifferentialStyleList

    return DifferentialStyleList


def test_parse(differential_style, datadir):
    datadir.chdir()
    with open("dxf_style.xml") as content:
        src = content.read()
    xml = fromstring(src)
    formats = []
    nodes = f"{{{SHEET_MAIN_NS}}}dxfs/{{{SHEET_MAIN_NS}}}dxf"
    for node in xml.findall(nodes):
        formats.append(differential_style.from_tree(node))
    assert len(formats) == 164
    cond = formats[1]
    expected = Font(
        underline="double",
        b=False,
        color=Color(auto=1),
        strikethrough=True,
        italic=True,
    )
    assert cond.font == expected
    assert cond.fill == PatternFill(end_color="FFFFC7CE")
    expected = Border(
        left=Side(),
        right=Side(),
        top=Side(style="thin", color=Color(theme=4)),
        bottom=Side(style="thin", color=Color(theme=4)),
        diagonal=None,
    )
    assert cond.border == expected


def test_serialise(differential_style):
    cond = differential_style()
    cond.font = Font(name="Calibri", family=2, sz=11)
    cond.fill = PatternFill()
    cond.border = Border(left=Side(), top=Side(style="thin", color=Color(auto=1)))
    xml = tostring(cond.to_tree())
    expected = """
    <dxf>
        <font>
            <name val="Calibri"></name>
            <family val="2"></family>
            <sz val="11"></sz>
        </font>
        <fill>
            <patternFill/>
        </fill>
        <border>
            <left/>
            <top style="thin">
                <color auto="1"/>
            </top>
        </border>
    </dxf>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


class TestDifferentialStyleList:
    def test_ctor(self, differential_style_list, differential_style):
        cond = differential_style()
        cond.font = Font(name="Calibri", family=2, sz=11)
        differential = differential_style_list(dxf=[cond])
        xml = tostring(differential.to_tree())
        expected = """
        <dxfs count="1">
            <dxf>
                <font>
                    <name val="Calibri"></name>
                    <family val="2"></family>
                    <sz val="11"></sz>
                </font>
            </dxf>
        </dxfs>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, differential_style_list):
        src = """
        <dxfs count="1">
            <dxf>
                <font>
                    <name val="Calibri"></name>
                    <family val="2"></family>
                    <sz val="11"></sz>
                </font>
            </dxf>
        </dxfs>
        """
        node = fromstring(src)
        differential = differential_style_list.from_tree(node)
        assert differential.dxf[0].font == Font(name="Calibri", family=2, sz=11)
