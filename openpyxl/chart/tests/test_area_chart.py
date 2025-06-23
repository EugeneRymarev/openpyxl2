# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.chart.series_factory import SeriesFactory
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def area_chart():
    from openpyxl.chart.area_chart import AreaChart

    return AreaChart


@pytest.fixture
def area_chart_3d():
    from openpyxl.chart.area_chart import AreaChart3D

    return AreaChart3D


class TestAreaChart:
    def test_ctor(self, area_chart):
        chart = area_chart()
        xml = tostring(chart.to_tree())
        expected = """
        <areaChart>
            <grouping val="standard"></grouping>
            <axId val="10"></axId>
            <axId val="100"></axId>
        </areaChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, area_chart):
        src = """
        <areaChart>
            <grouping val="percentStacked"/>
            <varyColors val="1"/>
            <axId val="10"></axId>
            <axId val="100"></axId>
        </areaChart>
        """
        node = fromstring(src)
        chart = area_chart.from_tree(node)
        assert chart == area_chart(grouping="percentStacked", varyColors=True)

    def test_write(self, area_chart):
        s1 = SeriesFactory(values="Sheet1!$A$1:$A$12")
        s2 = SeriesFactory(values="Sheet1!$B$1:$B$12")
        chart = area_chart(ser=[s1, s2])
        xml = tostring(chart._write())
        expected = """
        <chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <chart>
                <plotArea>
                    <areaChart>
                        <grouping val="standard"></grouping>
                        <ser>
                            <idx val="0"></idx>
                            <order val="0"></order>
                            <spPr>
                                <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                    <a:prstDash val="solid"/>
                                </a:ln>
                            </spPr>
                            <val>
                                <numRef>
                                    <f>'Sheet1'!$A$1:$A$12</f>
                                </numRef>
                            </val>
                        </ser>
                        <ser>
                            <idx val="1"></idx>
                            <order val="1"></order>
                            <spPr>
                                <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                    <a:prstDash val="solid"/>
                                </a:ln>
                            </spPr>
                            <val>
                                <numRef>
                                    <f>'Sheet1'!$B$1:$B$12</f>
                                </numRef>
                            </val>
                        </ser>
                        <axId val="10"></axId>
                        <axId val="100"></axId>
                    </areaChart>
                    <catAx>
                        <axId val="10"></axId>
                        <scaling>
                            <orientation val="minMax"></orientation>
                        </scaling>
                        <axPos val="l"/>
                        <majorTickMark val="none"/>
                        <minorTickMark val="none"/>
                        <crossAx val="100"></crossAx>
                        <lblOffset val="100"></lblOffset>
                    </catAx>
                    <valAx>
                        <axId val="100"></axId>
                        <scaling>
                            <orientation val="minMax"></orientation>
                        </scaling>
                        <axPos val="l"/>
                        <majorGridlines/>
                        <majorTickMark val="none"/>
                        <minorTickMark val="none"/>
                        <crossAx val="10"></crossAx>
                    </valAx>
                </plotArea>
                <legend>
                    <legendPos val="r"></legendPos>
                    <overlay val="0"/>
                </legend>
                <plotVisOnly val="1"/>
                <dispBlanksAs val="gap"></dispBlanksAs>
            </chart>
        </chartSpace>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestAreaChart3D:
    def test_ctor(self, area_chart_3d):
        chart = area_chart_3d()
        xml = tostring(chart.to_tree())
        expected = """
        <area3DChart>
            <grouping val="standard"></grouping>
            <axId val="10"></axId>
            <axId val="100"></axId>
            <axId val="1000"></axId>
        </area3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, area_chart_3d):
        src = """
        <area3DChart>
            <grouping val="standard"></grouping>
            <axId val="10"></axId>
            <axId val="100"></axId>
            <gapDepth val="150"/>
        </area3DChart>
        """
        node = fromstring(src)
        chart = area_chart_3d.from_tree(node)
        assert chart == area_chart_3d(gapDepth=150)
