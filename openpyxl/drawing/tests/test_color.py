# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def color_choice():
    from openpyxl.drawing.colors import ColorChoice

    return ColorChoice


@pytest.fixture
def system_color():
    from openpyxl.drawing.colors import SystemColor

    return SystemColor


@pytest.fixture
def hsl_color():
    from openpyxl.drawing.colors import HSLColor

    return HSLColor


@pytest.fixture
def rgb_percent():
    from openpyxl.drawing.colors import RGBPercent

    return RGBPercent


@pytest.fixture
def color_mapping():
    from openpyxl.drawing.colors import ColorMapping

    return ColorMapping


@pytest.fixture
def scheme_color():
    from openpyxl.drawing.colors import SchemeColor

    return SchemeColor


class TestColorChoice:
    def test_ctor(self, color_choice):
        color = color_choice()
        color.RGB = "000000"
        xml = tostring(color.to_tree())
        expected = """
        <colorChoice
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <srgbClr val="000000"/>
        </colorChoice>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, color_choice):
        src = "<colorChoice/>"
        node = fromstring(src)
        color = color_choice.from_tree(node)
        assert color == color_choice()


class TestSystemColor:
    def test_ctor(self, system_color):
        colors = system_color()
        xml = tostring(colors.to_tree())
        expected = """
        <sysClr val="windowText"
                 xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
        </sysClr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, system_color):
        src = '<sysClr val="windowText"></sysClr>'
        node = fromstring(src)
        colors = system_color.from_tree(node)
        assert colors == system_color(val="windowText")


class TestHSLColor:
    def test_ctor(self, hsl_color):
        colors = hsl_color(hue=50, sat=10, lum=90)
        xml = tostring(colors.to_tree())
        expected = '<hslClr hue="50" lum="90" sat="10"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, hsl_color):
        src = '<hslClr hue="0" lum="70" sat="20"/>'
        node = fromstring(src)
        colors = hsl_color.from_tree(node)
        assert colors == hsl_color(hue=0, sat=20, lum=70)


class TestRGBPercent:
    def test_ctor(self, rgb_percent):
        colors = rgb_percent(r=30, g=40, b=20)
        xml = tostring(colors.to_tree())
        expected = """
        <rgbClr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
                b="20"
                g="40"
                r="30"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, rgb_percent):
        src = '<rgbClr b="20" g="40" r="30"/>'
        node = fromstring(src)
        colors = rgb_percent.from_tree(node)
        assert colors == rgb_percent(r=30, g=40, b=20)


class TestColorMapping:
    def test_ctor(self, color_mapping):
        colors = color_mapping()
        xml = tostring(colors.to_tree())
        expected = """
        <clrMapOvr
                accent1="accent1"
                accent2="accent2"
                accent3="accent3"
                accent4="accent4"
                accent5="accent5"
                accent6="accent6"
                bg1="lt1"
                bg2="lt2"
                folHlink="folHlink"
                hlink="hlink"
                tx1="dk1"
                tx2="dk2"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, color_mapping):
        src = """
        <clrMapOvr
                accent1="accent1"
                accent2="accent2"
                accent3="accent3"
                accent4="accent4"
                accent5="accent5"
                accent6="accent6"
                bg1="lt1"
                bg2="lt2"
                folHlink="folHlink"
                hlink="hlink"
                tx1="dk1"
                tx2="dk2"/>
        """
        node = fromstring(src)
        colors = color_mapping.from_tree(node)
        assert colors == color_mapping()


class TestSchemeColor:
    def test_ctor(self, scheme_color):
        sclr = scheme_color(val="tx1")
        xml = tostring(sclr.to_tree())
        expected = """
        <schemeClr
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
                val="tx1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, scheme_color):
        src = '<schemeClr val="tx1"/>'
        node = fromstring(src)
        sclr = scheme_color.from_tree(node)
        assert sclr == scheme_color(val="tx1")
