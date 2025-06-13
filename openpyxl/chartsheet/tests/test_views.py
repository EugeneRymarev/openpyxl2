# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def chartsheet_view():
    from openpyxl.chartsheet.views import ChartsheetView

    return ChartsheetView


@pytest.fixture
def chartsheet_view_list():
    from openpyxl.chartsheet.views import ChartsheetViewList

    return ChartsheetViewList


class TestChartsheetView:
    def test_read(self, chartsheet_view):
        src = """
        <sheetView
                tabSelected="1"
                zoomScale="80"
                workbookViewId="0"
                zoomToFit="1"/>
        """
        xml = fromstring(src)
        chart = chartsheet_view.from_tree(xml)
        assert chart.tabSelected == True

    def test_write(self, chartsheet_view):
        view = chartsheet_view(
            tabSelected=True, zoomScale=80, workbookViewId=0, zoomToFit=True,
        )
        expected = """
        <sheetView
                tabSelected="1"
                zoomScale="80"
                workbookViewId="0"
                zoomToFit="1"/>
        """
        xml = tostring(view.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestChartsheetViewList:
    def test_read(self, chartsheet_view_list):
        src = """
        <sheetViews>
            <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
        </sheetViews>
        """
        xml = fromstring(src)
        views = chartsheet_view_list.from_tree(xml)
        assert views.sheetView[0].tabSelected == 1

    def test_write(self, chartsheet_view_list):
        views = chartsheet_view_list()
        expected = """
        <sheetViews>
            <sheetView workbookViewId="0" zoomToFit="1"/>
        </sheetViews>
        """
        xml = tostring(views.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
