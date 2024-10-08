# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def Hyperlink():
    from ..hyperlink import Hyperlink

    return Hyperlink


class TestHyperlink:

    def test_ctor(self, Hyperlink):
        hyperlink = Hyperlink(
            target="http://test.com",
            ref="A1",
            id="rId1",
            display="Link elsewhere",
        )
        xml = tostring(hyperlink.to_tree())
        expected = """
        <hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           display="Link elsewhere" r:id="rId1" ref="A1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, Hyperlink):
        src = """
        <hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            display="http://test.com" r:id="rId1" ref="A1"/>
        """
        node = fromstring(src)
        hyperlink = Hyperlink.from_tree(node)
        assert hyperlink == Hyperlink(display="http://test.com", ref="A1", id="rId1")


@pytest.fixture
def HyperlinkList():
    from ..hyperlink import HyperlinkList

    return HyperlinkList


class TestHyperlinkList:

    def test_ctor(self, HyperlinkList):
        fut = HyperlinkList()
        xml = tostring(fut.to_tree())
        expected = "<hyperlinks />"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, HyperlinkList):
        src = "<hyperlinks />"
        node = fromstring(src)
        fut = HyperlinkList.from_tree(node)
        assert fut == HyperlinkList()
