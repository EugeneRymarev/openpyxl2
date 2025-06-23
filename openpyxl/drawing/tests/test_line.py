# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def line_properties():
    from openpyxl.drawing.line import LineProperties

    return LineProperties


@pytest.fixture
def line_end_properties():
    from openpyxl.drawing.line import LineEndProperties

    return LineEndProperties


@pytest.fixture
def dash_stop():
    from openpyxl.drawing.line import DashStop

    return DashStop


class TestLineProperties:
    def test_ctor(self, line_properties):
        line = line_properties(w=10, miter=4)
        xml = tostring(line.to_tree())
        expected = """
        <ln w="10"
            xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <prstDash val="solid"/>
            <miter lim="4"/>
        </ln>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_color(self, line_properties):
        line = line_properties(w=10)
        line.solidFill = "FF0000"
        xml = tostring(line.to_tree())
        expected = """
        <ln w="10"
            xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <solidFill>
                <srgbClr val="FF0000"/>
            </solidFill>
            <prstDash val="solid"/>
        </ln>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, line_properties):
        src = """
        <ln w="38100" cmpd="sng">
            <prstDash val="solid"/>
            <miter lim="5"/>
        </ln>
        """
        node = fromstring(src)
        line = line_properties.from_tree(node)
        assert line == line_properties(w=38100, cmpd="sng", miter=5)


class TestLineEndProperties:
    def test_ctor(self, line_end_properties):
        line = line_end_properties()
        xml = tostring(line.to_tree())
        expected = (
            '<end xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        )
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, line_end_properties):
        src = "<end/>"
        node = fromstring(src)
        line = line_end_properties.from_tree(node)
        assert line == line_end_properties()


class TestDashStop:
    def test_ctor(self, dash_stop):
        line = dash_stop()
        xml = tostring(line.to_tree())
        expected = """
        <ds xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
            d="0"
            sp="0">
        </ds>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, dash_stop):
        src = '<ds d="10" sp="15"></ds>'
        node = fromstring(src)
        line = dash_stop.from_tree(node)
        assert line == dash_stop(d=10, sp=15)
