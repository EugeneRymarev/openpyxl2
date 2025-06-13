# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def sheet_background_picture():
    from openpyxl.chartsheet.chartsheet import SheetBackgroundPicture

    return SheetBackgroundPicture


@pytest.fixture
def drawing_hf():
    from openpyxl.chartsheet.chartsheet import DrawingHF

    return DrawingHF


class TestSheetBackgroundPicture:
    def test_read(self, sheet_background_picture):
        src = """
        <picture
                r:id="rId5"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        xml = fromstring(src)
        picture = sheet_background_picture.from_tree(xml)
        assert picture.id == "rId5"

    def test_write(self, sheet_background_picture):
        picture = sheet_background_picture(id="rId5")
        expected = """
        <picture
                r:id="rId5"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        xml = tostring(picture.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestDrawingHF:
    def test_read(self, drawing_hf):
        src = """
        <drawingHF
                lho="7"
                lhf="6"
                r:id="rId3"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        xml = fromstring(src)
        drawing = drawing_hf.from_tree(xml)
        assert drawing.lho == 7

    def test_write(self, drawing_hf):
        drawing = drawing_hf(lho=7, lhf=6, id="rId3")
        expected = """
        <drawingHF
                lho="7"
                lhf="6"
                r:id="rId3"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        xml = tostring(drawing.to_tree("drawingHF"))
        diff = compare_xml(xml, expected)
        assert diff is None, diff
