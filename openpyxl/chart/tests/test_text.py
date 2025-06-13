# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.title import title_maker
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def rich_text():
    from openpyxl.chart.text import RichText

    return RichText


@pytest.fixture
def text():
    from openpyxl.chart.text import Text

    return Text


class TestRichText:
    def test_ctor(self, rich_text):
        text = rich_text()
        xml = tostring(text.to_tree())
        expected = """
        <rich xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:bodyPr/>
            <a:p>
                <a:r>
                    <a:t/>
                </a:r>
            </a:p>
        </rich>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, rich_text):
        src = "<rich/>"
        node = fromstring(src)
        tx = rich_text.from_tree(node)
        assert tx == rich_text()


class TestText:
    def test_ctor(self, text):
        tx = text()
        tx.strRef = StrRef(f="Sheet1!$A$1")
        xml = tostring(tx.to_tree())
        expected = """
        <tx>
            <strRef>
                <f>Sheet1!$A$1</f>
            </strRef>
        </tx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, text):
        src = """
        <tx>
            <strRef>
            <f>Sheet1!$A$1</f>
            </strRef>
        </tx>
        """
        node = fromstring(src)
        tx = text.from_tree(node)
        assert tx == text(strRef=StrRef(f="Sheet1!$A$1"))

    def test_only_one(self, text):
        title = title_maker("Chart title")
        tx = text()
        tx.strRef = StrRef(f="Sheet1!$A$1")
        tx.rich = title.tx.rich
        expected = """
        <tx>
            <strRef>
            <f>Sheet1!$A$1</f>
            </strRef>
        </tx>
        """
        xml = tostring(tx.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
