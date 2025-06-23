# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def inline_font():
    from openpyxl.cell.text import InlineFont

    return InlineFont


@pytest.fixture
def rich_text():
    from openpyxl.cell.text import RichText

    return RichText


@pytest.fixture
def text_():
    from openpyxl.cell.text import Text

    return Text


@pytest.fixture
def phonetic_text():
    from openpyxl.cell.text import PhoneticText

    return PhoneticText


@pytest.fixture
def phonetic_properties():
    from openpyxl.cell.text import PhoneticProperties

    return PhoneticProperties


class TestInlineFont:
    def test_ctor(self, inline_font):
        font = inline_font()
        xml = tostring(font.to_tree())
        expected = "<RPrElt/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, inline_font):
        src = "<RPrElt/>"
        node = fromstring(src)
        font = inline_font.from_tree(node)
        assert font == inline_font()


class TestRichText:
    def test_ctor(self, rich_text):
        text = rich_text()
        xml = tostring(text.to_tree())
        expected = "<RElt/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, rich_text):
        src = "<RElt/>"
        node = fromstring(src)
        text = rich_text.from_tree(node)
        assert text == rich_text()


class TestText:
    def test_ctor(self, text_):
        text = text_()
        text.plain = "comment"
        xml = tostring(text.to_tree())
        expected = """
        <text>
            <t>comment</t>
        </text>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.parametrize(
        "src, expected",
        [
            (
                """
                <is>
                    <t>ID</t>
                </is>
                """,
                "ID",
            ),
            (
                """
                <is>
                    <r>
                        <rPr/>
                        <t xml:space="preserve">11 de September de 2014</t>
                    </r>
                </is>
                """,
                "11 de September de 2014",
            ),
        ],
    )
    def test_from_xml(self, text_, src, expected):
        node = fromstring(src)
        text = text_.from_tree(node)
        assert text.content == expected

    def test_empty_element(self, text_):
        src = """
        <si>
            <r>
                <t>Replaced Data</t>
            </r>
            <r>
                <rPr>
                    <sz val="11"/>
                    <color rgb="FF008080"/>
                    <rFont val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </rPr>
                <t/>
            </r>
        </si>
        """
        node = fromstring(src)
        text = text_.from_tree(node)
        assert text.content == "Replaced Data"


class TestPhoneticText:
    def test_ctor(self, phonetic_text):
        text = phonetic_text(sb=9, eb=10, t="\u3088")
        xml = tostring(text.to_tree())
        expected = b"""
        <rPh sb="9" eb="10">
            <t>&#12424;</t>
        </rPh>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, phonetic_text):
        src = b"""
        <rPh sb="9" eb="10">
            <t>&#12424;</t>
        </rPh>
        """
        node = fromstring(src)
        text = phonetic_text.from_tree(node)
        assert text == phonetic_text(sb=9, eb=10, t="\u3088")


class TestPhoneticProperties:
    def test_ctor(self, phonetic_properties):
        props = phonetic_properties(fontId=0, type="Hiragana")
        xml = tostring(props.to_tree())
        expected = '<phoneticPr fontId="0" type="Hiragana"></phoneticPr>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, phonetic_properties):
        src = '<phoneticPr fontId="0" type="noConversion"/>'
        node = fromstring(src)
        props = phonetic_properties.from_tree(node)
        assert props == phonetic_properties(fontId=0, type="noConversion")
