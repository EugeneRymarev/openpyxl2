# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def trendline_label():
    from openpyxl.chart.trendline import TrendlineLabel

    return TrendlineLabel


@pytest.fixture
def trendline():
    from openpyxl.chart.trendline import Trendline

    return Trendline


class TestTrendlineLabel:
    def test_ctor(self, trendline_label):
        tl = trendline_label()
        xml = tostring(tl.to_tree())
        expected = "<trendlineLbl></trendlineLbl>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, trendline_label):
        src = "<trendlineLbl></trendlineLbl>"
        node = fromstring(src)
        trendline = trendline_label.from_tree(node)
        assert trendline == trendline_label()


class TestTrendline:

    def test_ctor(self, trendline):
        tl = trendline(name="Bob")
        xml = tostring(tl.to_tree())
        expected = """
        <trendline name="Bob">
            <trendlineType val="linear"/>
        </trendline>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, trendline):
        src = """
        <trendline name="Bob">
            <trendlineType val="log"/>
        </trendline>
        """
        node = fromstring(src)
        tl = trendline.from_tree(node)
        assert tl == trendline(trendlineType="log", name="Bob")
