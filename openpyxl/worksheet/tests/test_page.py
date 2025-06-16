# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import tostring


@pytest.fixture
def page_margins():
    from openpyxl.worksheet.page import PageMargins

    return PageMargins


@pytest.fixture
def print_page_setup():
    from openpyxl.worksheet.page import PrintPageSetup

    return PrintPageSetup


@pytest.fixture
def dummy_worksheet():
    from openpyxl.workbook.workbook import Workbook

    wb = Workbook()
    return wb.active


@pytest.fixture
def print_options():
    from openpyxl.worksheet.page import PrintOptions

    return PrintOptions


class TestPageMargins:
    def test_ctor(self, page_margins):
        pm = page_margins()
        expected = {
            "bottom": "1",
            "footer": "0.5",
            "header": "0.5",
            "left": "0.75",
            "right": "0.75",
            "top": "1",
        }
        assert dict(pm) == expected

    def test_write(self, page_margins):
        pm = page_margins()
        pm.left = 2.0
        pm.right = 2.0
        pm.top = 2.0
        pm.bottom = 2.0
        pm.header = 1.5
        pm.footer = 1.5
        xml = tostring(pm.to_tree())
        expected = """
        <pageMargins
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                left="2"
                right="2"
                top="2"
                bottom="2"
                header="1.5"
                footer="1.5"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestPageSetup:
    def test_ctor(self, print_page_setup):
        p = print_page_setup()
        assert dict(p) == {}
        p.scale = 1
        assert p.scale == 1
        p.paperHeight = "24.73mm"
        assert p.paperHeight == "24.73mm"
        assert p.cellComments is None
        p.orientation = "default"
        assert p.orientation == "default"
        p.id = "a12"
        expected = {
            "scale": "1",
            "paperHeight": "24.73mm",
            "orientation": "default",
            "id": "a12",
        }
        assert dict(p) == expected

    def test_fit_to_page(self, dummy_worksheet):
        ws = dummy_worksheet
        p = ws.page_setup
        assert p.fitToPage is None
        p.fitToPage = 1
        assert p.fitToPage == True

    def test_auto_page_breaks(self, dummy_worksheet):
        ws = dummy_worksheet
        p = ws.page_setup
        assert p.autoPageBreaks is None
        p.autoPageBreaks = 1
        assert p.autoPageBreaks == True

    def test_write(self, print_page_setup):
        page_setup = print_page_setup()
        page_setup.orientation = "landscape"
        page_setup.paperSize = 3
        page_setup.fitToHeight = False
        page_setup.fitToWidth = True
        xml = tostring(page_setup.to_tree())
        expected = """
        <pageSetup
                orientation="landscape"
                paperSize="3"
                fitToHeight="0"
                fitToWidth="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestPrintOptions:
    def test_ctor(self, print_options):
        p = print_options()
        assert dict(p) == {}
        p.horizontalCentered = True
        p.verticalCentered = True
        assert dict(p) == {"verticalCentered": "1", "horizontalCentered": "1"}

    def test_write(self, print_options):
        po = print_options()
        po.horizontalCentered = True
        po.verticalCentered = True
        xml = tostring(po.to_tree())
        expected = '<printOptions horizontalCentered="1" verticalCentered="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff
