# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def surface_chart():
    from openpyxl.chart.surface_chart import SurfaceChart

    return SurfaceChart


@pytest.fixture
def surface_chart_3d():
    from openpyxl.chart.surface_chart import SurfaceChart3D

    return SurfaceChart3D


@pytest.fixture
def band_format():
    from openpyxl.chart.surface_chart import BandFormat

    return BandFormat


@pytest.fixture
def band_format_list():
    from openpyxl.chart.surface_chart import BandFormatList

    return BandFormatList


class TestSurfaceChart:
    def test_ctor(self, surface_chart):
        chart = surface_chart()
        xml = tostring(chart.to_tree())
        expected = """
        <surfaceChart>
            <axId val="10"></axId>
            <axId val="100"></axId>
            <axId val="1000"></axId>
        </surfaceChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, surface_chart):
        src = """
        <surfaceChart>
            <wireframe val="0"/>
            <ser>
                <idx val="0"/>
                <order val="0"/>
            </ser>
            <ser>
                <idx val="1"/>
                <order val="1"/>
            </ser>
            <bandFmts/>
            <axId val="2086876920"/>
            <axId val="2078923400"/>
            <axId val="2079274408"/>
        </surfaceChart>
        """
        node = fromstring(src)
        chart = surface_chart.from_tree(node)
        assert chart.axId == [2086876920, 2078923400, 2079274408]


class TestSurfaceChart3D:
    def test_ctor(self, surface_chart_3d):
        chart = surface_chart_3d()
        xml = tostring(chart.to_tree())
        expected = """
        <surface3DChart>
            <axId val="10"></axId>
            <axId val="100"></axId>
            <axId val="1000"></axId>
        </surface3DChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, surface_chart_3d):
        src = """
        <surface3DChart>
            <wireframe val="0"/>
            <ser>
                <idx val="0"/>
                <order val="0"/>
                <val>
                    <numRef>
                        <f>Blatt1!$A$1:$A$12</f>
                    </numRef>
                </val>
            </ser>
            <ser>
                <idx val="1"/>
                <order val="1"/>
                <val>
                    <numRef>
                        <f>Blatt1!$B$1:$B$12</f>
                    </numRef>
                </val>
            </ser>
            <bandFmts/>
            <axId val="2082935272"/>
            <axId val="2082938248"/>
            <axId val="2082941288"/>
        </surface3DChart>
        """
        node = fromstring(src)
        chart = surface_chart_3d.from_tree(node)
        assert len(chart.ser) == 2
        assert chart.axId == [2082935272, 2082938248, 2082941288]


class TestBandFormat:
    def test_ctor(self, band_format):
        fmt = band_format()
        xml = tostring(fmt.to_tree())
        expected = """
        <bandFmt>
            <idx val="0"/>
        </bandFmt>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, band_format):
        src = """
        <bandFmt>
            <idx val="4"></idx>
        </bandFmt>
        """
        node = fromstring(src)
        fmt = band_format.from_tree(node)
        assert fmt == band_format(idx=4)


class TestBandFormatList:
    def test_ctor(self, band_format_list):
        fmt = band_format_list()
        xml = tostring(fmt.to_tree())
        expected = "<bandFmts/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, band_format_list):
        src = "<bandFmts/>"
        node = fromstring(src)
        fmt = band_format_list.from_tree(node)
        assert fmt == band_format_list()
