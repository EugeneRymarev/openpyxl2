# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def book_view():
    from openpyxl.workbook.views import BookView

    return BookView


@pytest.fixture
def custom_workbook_view():
    from openpyxl.workbook.views import CustomWorkbookView

    return CustomWorkbookView


class TestBookView:
    def test_ctor(self, book_view):
        view = book_view()
        xml = tostring(view.to_tree())
        expected = """
        <workbookView
                activeTab="0"
                autoFilterDateGrouping="1"
                firstSheet="0"
                minimized="0"
                showHorizontalScroll="1"
                showSheetTabs="1"
                showVerticalScroll="1"
                tabRatio="600"
                visibility="visible"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, book_view):
        src = "<workbookView/>"
        node = fromstring(src)
        view = book_view.from_tree(node)
        assert view == book_view()


class TestCustomWorkbookView:
    def test_ctor(self, custom_workbook_view):
        view = custom_workbook_view(
            name="custom view",
            guid="{00000000-5BD2-4BC8-9F70-7020E1357FB2}",
            windowWidth=800,
            windowHeight=600,
            activeSheetId=1,
        )
        xml = tostring(view.to_tree())
        expected = """
        <customWorkbookView
                activeSheetId="1"
                guid="{00000000-5BD2-4BC8-9F70-7020E1357FB2}"
                name="custom view"
                showComments="commIndicator"
                showObjects="all"
                windowHeight="600"
                windowWidth="800"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, custom_workbook_view):
        src = """
        <customWorkbookView
                activeSheetId="1"
                guid="{00000000-5BD2-4BC8-9F70-7020E1357FB2}"
                name="custom view"
                showComments="commIndicator"
                showObjects="all"
                windowHeight="600"
                windowWidth="800"/>
        """
        node = fromstring(src)
        view = custom_workbook_view.from_tree(node)
        expected = custom_workbook_view(
            name="custom view",
            guid="{00000000-5BD2-4BC8-9F70-7020E1357FB2}",
            windowWidth=800,
            windowHeight=600,
            activeSheetId=1,
        )
        assert view == expected
