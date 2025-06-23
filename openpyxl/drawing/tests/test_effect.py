# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def outer_shadow():
    from openpyxl.drawing.effect import OuterShadow

    return OuterShadow


@pytest.fixture
def tint_effect():
    from openpyxl.drawing.effect import TintEffect

    return TintEffect


@pytest.fixture
def luminance_effect():
    from openpyxl.drawing.effect import LuminanceEffect

    return LuminanceEffect


class TestOuterShadow:
    def test_ctor(self, outer_shadow):
        shadow = outer_shadow(algn="tl", srgbClr="000000")
        xml = tostring(shadow.to_tree())
        expected = """
        <outerShdw
                algn="tl"
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <srgbClr val="000000"/>
        </outerShdw>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, outer_shadow):
        src = """
        <outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl">
            <srgbClr val="000000">
            </srgbClr>
        </outerShdw>
        """
        node = fromstring(src)
        shadow = outer_shadow.from_tree(node)
        expected = outer_shadow(
            algn="tl",
            blurRad=38100,
            dist=38100,
            dir=2700000,
            srgbClr="000000",
        )
        assert shadow == expected


class TestTintEffect:
    def test_ctor(self, tint_effect):
        tint = tint_effect()
        xml = tostring(tint.to_tree())
        expected = '<tint hue="0" amt="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, tint_effect):
        src = '<tint hue="56" amt="85"/>'
        node = fromstring(src)
        tint = tint_effect.from_tree(node)
        assert tint == tint_effect(hue=56, amt=85)


class TestLuminanceEffect:
    def test_ctor(self, luminance_effect):
        lum = luminance_effect()
        xml = tostring(lum.to_tree())
        expected = '<lum bright="0" contrast="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, luminance_effect):
        src = '<lum bright="45" contrast="80"/>'
        node = fromstring(src)
        lum = luminance_effect.from_tree(node)
        assert lum == luminance_effect(bright=45, contrast=80)
