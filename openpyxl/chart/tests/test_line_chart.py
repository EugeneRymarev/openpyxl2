# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def line_chart():
    from openpyxl.chart.line_chart import LineChart

    return LineChart


@pytest.fixture
def line_chart_3d():
    from openpyxl.chart.line_chart import LineChart3D

    return LineChart3D


class TestLineChart:
    def test_ctor(self, line_chart):
        chart = line_chart()
        xml = tostring(chart.to_tree())
        expected = """
        <lineChart>
            <grouping val="standard"></grouping>
            <axId val="10"></axId>
            <axId val="100"></axId>
        </lineChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, line_chart):
        src = """
        <lineChart>
            <grouping val="stacked"></grouping>
            <axId val="10"></axId>
            <axId val="100"></axId>
        </lineChart>
        """
        node = fromstring(src)
        chart = line_chart.from_tree(node)
        assert chart.axId == [10, 100]
        assert chart.grouping == "stacked"

    def test_axes(self, line_chart):
        chart = line_chart()
        assert set(chart._axes) == {10, 100}


class TestLineChart3D:
    def test_ctor(self, line_chart_3d):
        chart = line_chart_3d()
        xml = tostring(chart.to_tree())
        expected = """
        <line3DChart>
            <grouping val="standard"></grouping>
            <axId val="10"></axId>
            <axId val="100"></axId>
            <axId val="1000"></axId>
        </line3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, line_chart_3d):
        src = """
        <line3DChart>
            <grouping val="standard"></grouping>
        </line3DChart>
        """
        node = fromstring(src)
        chart = line_chart_3d.from_tree(node)
        assert chart == line_chart_3d()
