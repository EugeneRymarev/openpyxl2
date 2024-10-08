# Copyright (c) 2010-2024 openpyxl
import pytest

from .. import colors
from ..colors import Color
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def Side():
    from ..borders import Side

    return Side


@pytest.fixture
def Border():
    from ..borders import Border

    return Border


class TestBorder:

    def test_create(self, Border):
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
        bd = Border.from_tree(xml)
        assert bd.left.style == "thin"
        assert bd.right.color.value == "FF006600"
        assert bd.bottom.style == None
        assert bd.diagonal == None

    def test_serialise(self, Border, Side):
        medium_blue = Side(border_style="medium", color=Color(colors.BLUE))
        bd = Border(
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
