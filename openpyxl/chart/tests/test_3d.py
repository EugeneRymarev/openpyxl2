# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def view_3d():
    from openpyxl.chart._3d import View3D

    return View3D


@pytest.fixture
def surface():
    from openpyxl.chart._3d import Surface

    return Surface


@pytest.fixture
def _3d_base():
    from openpyxl.chart._3d import _3DBase

    return _3DBase


class TestView3D:
    def test_ctor(self, view_3d):
        view = view_3d()
        xml = tostring(view.to_tree())
        expected = """
        <view3D>
            <rotX val="15"></rotX>
            <rotY val="20"></rotY>
            <rAngAx val="1"></rAngAx>
        </view3D>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, view_3d):
        src = """
        <view3D>
            <rotX val="15"/>
            <rotY val="20"/>
            <rAngAx val="0"/>
            <perspective val="30"/>
        </view3D>
        """
        node = fromstring(src)
        view = view_3d.from_tree(node)
        assert view == view_3d(rotX=15, rotY=20, rAngAx=False, perspective=30)


class TestSurface:
    def test_ctor(self, surface):
        surface = surface(thickness=0)
        xml = tostring(surface.to_tree())
        expected = """
        <surface>
            <thickness val="0"/>
        </surface>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, surface):
        src = """
        <floor xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <thickness val="0"/>
        </floor>
        """
        node = fromstring(src)
        surface = surface.from_tree(node)
        assert surface == surface(thickness=0)


class TestSurface:
    def test_ctor(self, _3d_base):
        base = _3d_base()
        xml = tostring(base.to_tree())
        expected = """
        <ChartBase>
            <backWall/>
            <floor/>
            <sideWall/>
            <view3D>
                <rotX val="15"/>
                <rotY val="20"/>
                <rAngAx val="1"/>
            </view3D>
        </ChartBase>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, _3d_base):
        src = """
        <ChartBase>
            <backWall/>
            <floor/>
            <sideWall/>
            <view3D>
                <rotX val="15"/>
                <rotY val="20"/>
                <rAngAx val="1"/>
            </view3D>
        </ChartBase>
        """
        node = fromstring(src)
        base = _3d_base.from_tree(node)
        assert base == _3d_base()
