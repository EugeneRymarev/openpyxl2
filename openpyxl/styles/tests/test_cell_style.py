# Copyright (c) 2010-2025 openpyxl
import copy

import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def style_array():
    from openpyxl.styles.cell_style import StyleArray

    return StyleArray


@pytest.fixture
def cell_style():
    from openpyxl.styles.cell_style import CellStyle

    return CellStyle


@pytest.fixture
def cell_style_list():
    from openpyxl.styles.cell_style import CellStyleList

    return CellStyleList


class TestStyleArray:
    def test_ctor(self, style_array):
        style = style_array(range(9))
        assert style.fontId == 0
        assert style.numFmtId == 3
        assert style.xfId == 8

    def test_hash(self, style_array):
        s1 = style_array((range(9)))
        s2 = style_array((range(9)))
        assert hash(s1) == hash(s2)

    def test_copy(self, style_array):
        s1 = style_array((range(9)))
        s2 = copy.copy(s1)
        assert type(s1) == type(s2)
        assert s1 == s2


class TestCellStyle:
    def test_ctor(self, cell_style):
        cell_style = cell_style(xfId=0)
        xml = tostring(cell_style.to_tree())
        expected = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, cell_style):
        from openpyxl.styles.alignment import Alignment

        src = """
        <xf numFmtId="0"
            fontId="0"
            fillId="0"
            borderId="0"
            xfId="0"
            applyAlignment="1">
            <alignment horizontal="center"/>
        </xf>
        """
        node = fromstring(src)
        cs = cell_style.from_tree(node)
        expected = cell_style(
            alignment=Alignment(horizontal="center"),
            applyAlignment=True,
            xfId=0,
        )
        assert cs == expected

    def test_to_array(self, cell_style):
        from openpyxl.styles.cell_style import StyleArray

        xf = cell_style(
            numFmtId=43,
            fontId=1,
            fillId=2,
            borderId=4,
            xfId=None,
            quotePrefix=True,
            pivotButton=True,
            applyNumberFormat=None,
            applyFont=None,
            applyFill=None,
            applyBorder=None,
            applyAlignment=None,
            applyProtection=None,
            alignment=None,
            protection=None,
        )
        style = xf.to_array()
        assert style == StyleArray([1, 2, 4, 43, 0, 0, 1, 1, 0])

    def test_from_array(self, cell_style):
        from openpyxl.styles.cell_style import StyleArray

        style = StyleArray([5, 10, 15, 0, 0, 0, 1, 1, 15])
        xf = cell_style.from_array(style)
        expected = {
            "borderId": "15",
            "fillId": "10",
            "fontId": "5",
            "numFmtId": "0",
            "pivotButton": "1",
            "quotePrefix": "1",
            "xfId": "15",
        }
        assert dict(xf) == expected


class TestCellStyleList:
    def test_ctor(self, cell_style_list):
        cs = cell_style_list()
        xml = tostring(cs.to_tree())
        expected = '<cellXfs count="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, cell_style_list):
        src = '<cellXfs count="0"/>'
        node = fromstring(src)
        cs = cell_style_list.from_tree(node)
        assert cs == cell_style_list()

    def test_to_array(self, cell_style_list):
        src = """
        <cellXfs count="29">
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"/>
            <xf numFmtId="0"
                fontId="3"
                fillId="0"
                borderId="0"
                xfId="0"
                applyFont="1"/>
            <xf numFmtId="0"
                fontId="4"
                fillId="0"
                borderId="0"
                xfId="0"
                applyFont="1"/>
            <xf numFmtId="0"
                fontId="5"
                fillId="0"
                borderId="0"
                xfId="0"
                applyFont="1"/>
            <xf numFmtId="0"
                fontId="6"
                fillId="0"
                borderId="0"
                xfId="0"
                applyFont="1"/>
            <xf numFmtId="0"
                fontId="7"
                fillId="0"
                borderId="0"
                xfId="0"
                applyFont="1"/>
            <xf numFmtId="0"
                fontId="0"
                fillId="2"
                borderId="0"
                xfId="0"
                applyFill="1"/>
            <xf numFmtId="0"
                fontId="0"
                fillId="3"
                borderId="0"
                xfId="0"
                applyFill="1"/>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment horizontal="left"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment horizontal="right"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment vertical="top"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment vertical="center"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1"/>
            <xf numFmtId="2"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyNumberFormat="1"/>
            <xf numFmtId="14"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyNumberFormat="1"/>
            <xf numFmtId="10"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyNumberFormat="1"/>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="1"
                xfId="0"
                applyBorder="1"/>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="2"
                xfId="0"
                applyBorder="1"/>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment wrapText="1"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment shrinkToFit="1"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyFill="1"
                applyBorder="1"/>
            <xf numFmtId="0"
                fontId="0"
                fillId="0"
                borderId="0"
                xfId="0"
                applyAlignment="1">
                <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="4"
                borderId="3"
                xfId="0"
                applyFill="1"
                applyBorder="1"
                applyAlignment="1">
                <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="4"
                borderId="4"
                xfId="0"
                applyFill="1"
                applyBorder="1"
                applyAlignment="1">
                <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="4"
                borderId="5"
                xfId="0"
                applyFill="1"
                applyBorder="1"
                applyAlignment="1">
                <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0"
                fontId="0"
                fillId="4"
                borderId="6"
                xfId="0"
                applyFill="1"
                applyBorder="1"
                applyAlignment="1">
                <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0"
                fontId="6"
                fillId="5"
                borderId="0"
                xfId="0"
                applyFont="1"
                applyFill="1"/>
        </cellXfs>
        """
        node = fromstring(src)
        xfs = cell_style_list.from_tree(node)
        styles = xfs._to_array()
        assert len(styles) == 29
        assert len(xfs.alignments) == 9
        assert len(xfs.prots) == 1
