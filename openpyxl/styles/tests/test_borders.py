# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.styles.colors import BLUE
from openpyxl.styles.colors import Color
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def side():
    from openpyxl.styles.borders import Side

    return Side


@pytest.fixture
def border():
    from openpyxl.styles.borders import Border

    return Border


class TestBorder:
    def test_create(self, border):
        src = """
        <border xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <left style="thin">
                <color rgb="FF006600"/>
            </left>
            <right style="thin">
                <color rgb="FF006600"/>
            </right>
            <top style="thin">
                <color rgb="FF006600"/>
            </top>
            <bottom/>
        </border>
        """
        xml = fromstring(src)
        bd = border.from_tree(xml)
        assert bd.left.style == "thin"
        assert bd.right.color.value == "FF006600"
        assert bd.bottom.style is None
        assert bd.diagonal is None

    def test_serialise(self, border, side):
        medium_blue = side(border_style="medium", color=Color(BLUE))
        bd = border(
            left=medium_blue,
            right=medium_blue,
            top=medium_blue,
            bottom=medium_blue,
            outline=False,
            diagonalDown=True,
        )
        xml = tostring(bd.to_tree())
        expected = """
        <border diagonalDown="1" outline="0">
            <left style="medium">
                <color rgb="000000FF"></color>
            </left>
            <right style="medium">
                <color rgb="000000FF"></color>
            </right>
            <top style="medium">
                <color rgb="000000FF"></color>
            </top>
            <bottom style="medium">
                <color rgb="000000FF"></color>
            </bottom>
        </border>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
