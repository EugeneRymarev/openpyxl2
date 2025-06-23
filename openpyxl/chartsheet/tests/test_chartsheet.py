# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.chartsheet.views import ChartsheetView
from openpyxl.chartsheet.views import ChartsheetViewList
from openpyxl.tests.helper import compare_xml
from openpyxl.worksheet.drawing import Drawing
from openpyxl.worksheet.page import PageMargins
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def chartsheet():
    from openpyxl.chartsheet.chartsheet import Chartsheet

    return Chartsheet


class DummyWorkbook:
    def __init__(self):
        self.sheetnames = []
        self._charts = []


class TestChartsheet:
    def test_ctor(self, chartsheet):
        cs = chartsheet(parent=DummyWorkbook())
        assert cs.title == "Chart"

    def test_read(self, chartsheet):
        src = """
        <chartsheet
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheetPr/>
            <sheetViews>
                <sheetView
                        tabSelected="1"
                        zoomScale="80"
                        workbookViewId="0"
                        zoomToFit="1"/>
            </sheetViews>
            <pageMargins
                    left="0.7"
                    right="0.7"
                    top="0.75"
                    bottom="0.75"
                    header="0.3"
                    footer="0.3"/>
            <drawing r:id="rId1"/>
        </chartsheet>
        """
        xml = fromstring(src)
        chart = chartsheet.from_tree(xml)
        assert chart.pageMargins.left == 0.7
        assert chart.sheetViews.sheetView[0].tabSelected == True

    def test_write(self, chartsheet):
        sheetview = ChartsheetView(
            tabSelected=True,
            zoomScale=80,
            workbookViewId=0,
            zoomToFit=True,
        )
        chartsheetViews = ChartsheetViewList(sheetView=[sheetview])
        pageMargins = PageMargins(
            left=0.7,
            right=0.7,
            top=0.75,
            bottom=0.75,
            header=0.3,
            footer=0.3,
        )
        drawing = Drawing("rId1")
        item = chartsheet(
            sheetViews=chartsheetViews,
            pageMargins=pageMargins,
            drawing=drawing,
        )
        expected = """
        <chartsheet
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheetViews>
                <sheetView tabSelected="1" zoomScale="80" workbookViewId="0" zoomToFit="1"/>
            </sheetViews>
            <pageMargins
                    left="0.7"
                    right="0.7"
                    top="0.75"
                    bottom="0.75"
                    header="0.3"
                    footer="0.3"/>
            <drawing r:id="rId1"/>
        </chartsheet>
        """
        xml = tostring(item.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_charts(self, chartsheet):
        class DummyChart:
            pass

        cs = chartsheet(parent=DummyWorkbook())
        cs.add_chart(DummyChart())
        expected = """
        <chartsheet
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheetViews>
                <sheetView workbookViewId="0" zoomToFit="1"></sheetView>
            </sheetViews>
            <drawing r:id="rId1"/>
        </chartsheet>
        """
        xml = tostring(cs.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
