# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def TrendlineLabel():
    from ..trendline import TrendlineLabel

    return TrendlineLabel


class TestTrendlineLabel:

    def test_ctor(self, TrendlineLabel):
        trendline = TrendlineLabel()
        xml = tostring(trendline.to_tree())
        expected = """
        <trendlineLbl></trendlineLbl>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, TrendlineLabel):
        src = """
        <trendlineLbl></trendlineLbl>
        """
        node = fromstring(src)
        trendline = TrendlineLabel.from_tree(node)
        assert trendline == TrendlineLabel()


@pytest.fixture
def Trendline():
    from ..trendline import Trendline

    return Trendline


class TestTrendline:

    def test_ctor(self, Trendline):
        trendline = Trendline(name="Bob")
        xml = tostring(trendline.to_tree())
        expected = """
        <trendline name="Bob">
          <trendlineType val="linear" />
        </trendline>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, Trendline):
        src = """
        <trendline name="Bob">
          <trendlineType val="log" />
        </trendline>
        """
        node = fromstring(src)
        trendline = Trendline.from_tree(node)
        assert trendline == Trendline(trendlineType="log", name="Bob")
