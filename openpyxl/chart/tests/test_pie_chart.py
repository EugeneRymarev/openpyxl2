# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def pie_chart():
    from openpyxl.chart.pie_chart import PieChart

    return PieChart


@pytest.fixture
def pie_chart_3d():
    from openpyxl.chart.pie_chart import PieChart3D

    return PieChart3D


@pytest.fixture
def doughnut_chart():
    from openpyxl.chart.pie_chart import DoughnutChart

    return DoughnutChart


@pytest.fixture
def projected_pie_chart():
    from openpyxl.chart.pie_chart import ProjectedPieChart

    return ProjectedPieChart


@pytest.fixture
def custom_split():
    from openpyxl.chart.pie_chart import CustomSplit

    return CustomSplit


class TestPieChart:
    def test_ctor(self, pie_chart):
        chart = pie_chart()
        xml = tostring(chart.to_tree())
        expected = """
        <pieChart>
            <varyColors val="1"/>
            <firstSliceAng val="0"/>
        </pieChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pie_chart):
        src = """
        <pieChart>
            <varyColors val="1"/>
            <explosion val="5"/>
            <firstSliceAng val="60"/>
        </pieChart>
        """
        node = fromstring(src)
        chart = pie_chart.from_tree(node)
        assert dict(chart) == {}
        assert chart.varyColors is True
        assert chart.firstSliceAng == 60


class TestPieChart3D:
    def test_ctor(self, pie_chart_3d):
        chart = pie_chart_3d()
        xml = tostring(chart.to_tree())
        expected = """
        <pie3DChart>
            <varyColors val="1"/>
        </pie3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestDoughnutChart:
    def test_ctor(self, doughnut_chart):
        chart = doughnut_chart()
        xml = tostring(chart.to_tree())
        expected = """
        <doughnutChart>
            <varyColors val="1"/>
            <firstSliceAng val="0"/>
            <holeSize val="10"/>
        </doughnutChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, doughnut_chart):
        src = """
        <doughnutChart>
            <firstSliceAng val="0"/>
            <holeSize val="50"/>
        </doughnutChart>
        """
        node = fromstring(src)
        chart = doughnut_chart.from_tree(node)
        assert dict(chart) == {}
        assert chart.firstSliceAng == 0
        assert chart.holeSize == 50


class TestProjectedPieChart:
    def test_ctor(self, projected_pie_chart):
        chart = projected_pie_chart()
        xml = tostring(chart.to_tree())
        expected = """
        <ofPieChart>
            <varyColors val="1"/>
            <ofPieType val="pie"/>
            <splitType val="auto"/>
            <secondPieSize val="75"/>
            <serLines/>
        </ofPieChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, projected_pie_chart):
        src = """
        <ofPieChart>
            <varyColors val="1"/>
            <ofPieType val="pie"/>
            <splitType val="auto"/>
            <dLbls>
                <showLegendKey val="0"/>
                <showVal val="0"/>
                <showCatName val="0"/>
                <showSerName val="0"/>
                <showPercent val="0"/>
                <showBubbleSize val="0"/>
                <showLeaderLines val="1"/>
            </dLbls>
            <gapWidth val="150"/>
            <secondPieSize val="75"/>
            <serLines/>
        </ofPieChart>
        """
        node = fromstring(src)
        chart = projected_pie_chart.from_tree(node)
        assert dict(chart) == {}
        assert chart.gapWidth == 150
        assert chart.secondPieSize == 75


class TestCustomSplit:
    def test_ctor(self, custom_split):
        pie_chart = custom_split([1, 2, 3])
        xml = tostring(pie_chart.to_tree())
        expected = """
        <custSplit>
            <secondPiePt val="1"/>
            <secondPiePt val="2"/>
            <secondPiePt val="3"/>
        </custSplit>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, custom_split):
        src = """
        <custSplit>
            <secondPiePt val="1"/>
            <secondPiePt val="2"/>
        </custSplit>
        """
        node = fromstring(src)
        chart = custom_split.from_tree(node)
        assert chart == custom_split([1, 2])
