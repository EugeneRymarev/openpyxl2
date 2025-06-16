# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def hyperlink():
    from openpyxl.worksheet.hyperlink import Hyperlink

    return Hyperlink


@pytest.fixture
def hyperlink_list():
    from openpyxl.worksheet.hyperlink import HyperlinkList

    return HyperlinkList


class TestHyperlink:
    def test_ctor(self, hyperlink):
        link = hyperlink(
            target="http://test.com",
            ref="A1",
            id="rId1",
            display="Link elsewhere",
        )
        xml = tostring(link.to_tree())
        expected = """
        <hyperlink
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                display="Link elsewhere"
                r:id="rId1"
                ref="A1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, hyperlink):
        src = """
        <hyperlink
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                display="http://test.com"
                r:id="rId1"
                ref="A1"/>
        """
        node = fromstring(src)
        link = hyperlink.from_tree(node)
        assert link == hyperlink(display="http://test.com", ref="A1", id="rId1")


class TestHyperlinkList:
    def test_ctor(self, hyperlink_list):
        fut = hyperlink_list()
        xml = tostring(fut.to_tree())
        expected = "<hyperlinks/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, hyperlink_list):
        src = "<hyperlinks/>"
        node = fromstring(src)
        fut = hyperlink_list.from_tree(node)
        assert fut == hyperlink_list()
