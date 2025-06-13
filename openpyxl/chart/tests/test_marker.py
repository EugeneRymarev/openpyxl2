# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def marker_():
    from openpyxl.chart.marker import Marker

    return Marker


@pytest.fixture
def data_point():
    from openpyxl.chart.marker import DataPoint

    return DataPoint


class TestMarker:
    def test_ctor(self, marker_):
        marker = marker_(symbol=None, size=5)
        xml = tostring(marker.to_tree())
        expected = """
        <marker>
            <symbol val="none"/>
            <size val="5"/>
            <spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:ln>
                    <a:prstDash val="solid"/>
                </a:ln>
            </spPr>
        </marker>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, marker_):
        src = """
        <marker>
            <symbol val="square"/>
            <size val="5"/>
        </marker>
        """
        node = fromstring(src)
        marker = marker_.from_tree(node)
        assert marker == marker_(symbol="square", size=5)


class TestDataPoint:
    def test_ctor(self, data_point):
        dp = data_point(idx=9)
        xml = tostring(dp.to_tree())
        expected = """
        <dPt>
            <idx val="9"/>
            <spPr>
                <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:prstDash val="solid"/>
                </a:ln>
            </spPr>
        </dPt>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, data_point):
        src = """
        <dPt>
            <idx val="9"/>
            <marker>
                <symbol val="triangle"/>
                <size val="5"/>
            </marker>
            <bubble3D val="0"/>
        </dPt>
        """
        node = fromstring(src)
        dp = data_point.from_tree(node)
        assert dp.idx == 9
        assert dp.bubble3D is False
