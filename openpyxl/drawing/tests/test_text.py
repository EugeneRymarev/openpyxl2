# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def paragraph():
    from openpyxl.drawing.text import Paragraph

    return Paragraph


@pytest.fixture
def paragraph_properties():
    from openpyxl.drawing.text import ParagraphProperties

    return ParagraphProperties


@pytest.fixture
def character_properties():
    from openpyxl.drawing.text import CharacterProperties

    return CharacterProperties


@pytest.fixture
def font():
    from openpyxl.drawing.text import Font

    return Font


@pytest.fixture
def hyperlink():
    from openpyxl.drawing.text import Hyperlink

    return Hyperlink


@pytest.fixture
def line_break():
    from openpyxl.drawing.text import LineBreak

    return LineBreak


class TestParagraph:
    def test_ctor(self, paragraph):
        text = paragraph()
        xml = tostring(text.to_tree())
        expected = """
        <p xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <r>
                <t/>
            </r>
        </p>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, paragraph):
        src = "<p/>"
        node = fromstring(src)
        text = paragraph.from_tree(node)
        assert text == paragraph()

    def test_multiline(self, paragraph):
        src = """
        <p>
            <r>
                <t>Adjusted Absorbance vs.</t>
            </r>
            <r>
                <t> Concentration</t>
            </r>
        </p>
        """
        node = fromstring(src)
        para = paragraph.from_tree(node)
        assert len(para.text) == 2


class TestParagraphProperties:
    def test_ctor(self, paragraph_properties):
        text = paragraph_properties(defTabSz=91400)
        xml = tostring(text.to_tree())
        expected = """
        <pPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
             defTabSz="91400"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, paragraph_properties):
        src = '<pPr defTabSz="91400"/>'
        node = fromstring(src)
        text = paragraph_properties.from_tree(node)
        assert text == paragraph_properties(defTabSz=91400)


class TestTextBox:
    def test_from_xml(self, datadir):
        datadir.chdir()
        with open("text_box_drawing.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        drawing = SpreadsheetDrawing.from_tree(node)
        anchor = drawing.twoCellAnchor[0]
        box = anchor.sp
        meta = box.nvSpPr
        graphic = box.graphicalProperties
        text = box.txBody
        assert len(text.p) == 2


class TestCharacterProperties:
    def test_ctor(self, character_properties):
        from openpyxl.drawing.text import Font

        normal_font = Font(typeface="Arial")
        text = character_properties(
            latin=normal_font,
            sz=900,
            b=False,
            solidFill="FFC000",
        )
        xml = tostring(text.to_tree())
        expected = """
        <a:defRPr
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                b="0"
                sz="900">
            <a:solidFill>
                <a:srgbClr val="FFC000"/>
            </a:solidFill>
            <a:latin typeface="Arial"/>
        </a:defRPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, character_properties):
        src = '<defRPr sz="110"/>'
        node = fromstring(src)
        text = character_properties.from_tree(node)
        assert text == character_properties(sz=110)


class TestFont:
    def test_ctor(self, font):
        fut = font("Arial")
        xml = tostring(fut.to_tree())
        expected = """
        <latin typeface="Arial"
               xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, font):
        src = """
        <latin typeface="Arial"
               pitchFamily="40"
               xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        """
        node = fromstring(src)
        fut = font.from_tree(node)
        assert fut == font(typeface="Arial", pitchFamily=40)


class TestHyperlink:
    def test_ctor(self, hyperlink):
        link = hyperlink()
        xml = tostring(link.to_tree())
        expected = """
        <hlinkClick
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, hyperlink):
        src = '<hlinkClick tooltip="Select/de-select all"/>'
        node = fromstring(src)
        link = hyperlink.from_tree(node)
        assert link == hyperlink(tooltip="Select/de-select all")


class TestLineBreak:
    def test_ctor(self, line_break):
        fut = line_break()
        xml = tostring(fut.to_tree())
        expected = '<br xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, line_break):
        src = "<br/>"
        node = fromstring(src)
        fut = line_break.from_tree(node)
        assert fut == line_break()
