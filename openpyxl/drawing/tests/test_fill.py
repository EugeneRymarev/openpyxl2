# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def pattern_fill_properties():
    from openpyxl.drawing.fill import PatternFillProperties

    return PatternFillProperties


@pytest.fixture
def relative_rect():
    from openpyxl.drawing.fill import RelativeRect

    return RelativeRect


@pytest.fixture
def stretch_info_properties():
    from openpyxl.drawing.fill import StretchInfoProperties

    return StretchInfoProperties


@pytest.fixture
def gradient_stop():
    from openpyxl.drawing.fill import GradientStop

    return GradientStop


@pytest.fixture
def linear_shade_properties():
    from openpyxl.drawing.fill import LinearShadeProperties

    return LinearShadeProperties


@pytest.fixture
def path_shade_properties():
    from openpyxl.drawing.fill import PathShadeProperties

    return PathShadeProperties


@pytest.fixture
def gradient_fill_properties():
    from openpyxl.drawing.fill import GradientFillProperties

    return GradientFillProperties


@pytest.fixture
def blip():
    from openpyxl.drawing.fill import Blip

    return Blip


@pytest.fixture
def blip_fill_properties():
    from openpyxl.drawing.fill import BlipFillProperties

    return BlipFillProperties


class TestPatternFillProperties:
    def test_ctor(self, pattern_fill_properties):
        fill = pattern_fill_properties(prst="cross")
        xml = tostring(fill.to_tree())
        expected = """
        <pattFill
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
                prst="cross">
            </pattFill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pattern_fill_properties):
        src = '<pattFill prst="dashHorz"/>'
        node = fromstring(src)
        fill = pattern_fill_properties.from_tree(node)
        assert fill == pattern_fill_properties("dashHorz")


class TestRelativeRect:
    def test_ctor(self, relative_rect):
        fill = relative_rect(10, 15, 20, 25)
        xml = tostring(fill.to_tree())
        expected = """
        <rect xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
              b="25"
              l="10"
              r="20"
              t="15"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, relative_rect):
        src = '<rect b="25" l="10" r="20" t="15"/>'
        node = fromstring(src)
        fill = relative_rect.from_tree(node)
        assert fill == relative_rect(10, 15, 20, 25)

    def test_from_src_rect(self, relative_rect):
        src = '<srcRect l="71321" t="10170" r="4935" b="80270"/>'
        node = fromstring(src)
        fill = relative_rect.from_tree(node)
        assert fill == relative_rect(l=71321, t=10170, r=4935, b=80270)


class TestStretchInfoProperties:
    def test_ctor(self, stretch_info_properties):
        fill = stretch_info_properties()
        xml = tostring(fill.to_tree())
        expected = """
        <stretch xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <fillRect/>
        </stretch>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, stretch_info_properties):
        src = "<stretch/>"
        node = fromstring(src)
        fill = stretch_info_properties.from_tree(node)
        assert fill == stretch_info_properties()


class TestGradientStop:
    def test_ctor(self, gradient_stop):
        fill = gradient_stop(pos=0, prstClr="blue")
        xml = tostring(fill.to_tree())
        expected = """
        <a:gs xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              pos="0">
            <a:prstClr val="blue"/>
        </a:gs>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, gradient_stop):
        src = """
        <a:gs xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              pos="0">
            <a:prstClr val="blue"/>
        </a:gs>
        """
        node = fromstring(src)
        fill = gradient_stop.from_tree(node)
        assert fill == gradient_stop(pos=0, prstClr="blue")


class TestLinearShadeProperties:
    def test_ctor(self, linear_shade_properties):
        fill = linear_shade_properties(ang=0, scaled=True)
        xml = tostring(fill.to_tree())
        expected = """
        <a:lin xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               ang="0"
               scaled="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, linear_shade_properties):
        src = """
        <a:lin xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               ang="0"
               scaled="1"/>
        """
        node = fromstring(src)
        fill = linear_shade_properties.from_tree(node)
        assert fill == linear_shade_properties(ang=0, scaled=True)


class TestPathShadeProperties:
    def test_ctor(self, path_shade_properties):
        fill = path_shade_properties(path="circle")
        xml = tostring(fill.to_tree())
        expected = """
        <a:path xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                path="circle"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, path_shade_properties):
        src = """
        <a:path xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                path="circle"/>
        """
        node = fromstring(src)
        fill = path_shade_properties.from_tree(node)
        assert fill == path_shade_properties(path="circle")


class TestGradientFillProperties:
    def test_ctor(self, gradient_fill_properties):
        fill = gradient_fill_properties(flip="xy")
        xml = tostring(fill.to_tree())
        expected = """
        <a:gradFill
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                flip="xy"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, gradient_fill_properties):
        src = """
        <a:gradFill
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                flip="xy"/>
        """
        node = fromstring(src)
        fill = gradient_fill_properties.from_tree(node)
        assert fill == gradient_fill_properties(flip="xy")


class TestBlip:
    def test_ctor(self, blip):
        fill = blip(embed="rId1")
        xml = tostring(fill.to_tree())
        expected = """
        <blip xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
              r:embed="rId1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, blip):
        src = "<blip/>"
        node = fromstring(src)
        fill = blip.from_tree(node)
        assert fill == blip()


class TestBlipFillProperties:
    def test_ctor(self, blip_fill_properties):
        fill = blip_fill_properties()
        xml = tostring(fill.to_tree())
        expected = """
        <blipFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:stretch >
                <a:fillRect/>
            </a:stretch>
        </blipFill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, blip_fill_properties):
        src = "<blipFill/>"
        node = fromstring(src)
        fill = blip_fill_properties.from_tree(node)
        assert fill == blip_fill_properties()
