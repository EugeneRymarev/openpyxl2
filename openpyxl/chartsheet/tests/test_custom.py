# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.worksheet.page import PageMargins
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def custom_chartsheet_view():
    from openpyxl.chartsheet.custom import CustomChartsheetView

    return CustomChartsheetView


@pytest.fixture
def custom_chartsheet_views():
    from openpyxl.chartsheet.custom import CustomChartsheetViews

    return CustomChartsheetViews


class TestCustomChartsheetView:
    def test_read(self, custom_chartsheet_view):
        src = """
        <customSheetView
                guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}"
                scale="88"
                zoomToFit="1">
            <pageMargins
                    left="0.23622047244094491"
                    right="0.23622047244094491"
                    top="0.74803149606299213"
                    bottom="0.74803149606299213"
                    header="0.31496062992125984"
                    footer="0.31496062992125984"/>
            <pageSetup
                    paperSize="7"
                    orientation="landscape"
                    r:id="rId1"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
            <headerFooter/>
        </customSheetView>
        """
        xml = fromstring(src)
        view = custom_chartsheet_view.from_tree(xml)
        assert view.state == "visible"
        assert view.scale == 88
        assert view.pageMargins.left == 0.23622047244094491

    def test_write(self, custom_chartsheet_view):
        page_margins = PageMargins(
            left=0.2362204724409449,
            right=0.2362204724409449,
            top=0.7480314960629921,
            bottom=0.7480314960629921,
            header=0.3149606299212598,
            footer=0.3149606299212598,
        )
        view = custom_chartsheet_view(
            guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}",
            scale=88,
            zoomToFit=1,
            pageMargins=page_margins,
        )
        expected = """
        <customSheetView
                guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}"
                scale="88"
                state="visible"
                zoomToFit="1">
            <pageMargins
                    left="0.2362204724409449"
                    right="0.2362204724409449"
                    top="0.7480314960629921"
                    bottom="0.7480314960629921"
                    header="0.3149606299212598"
                    footer="0.3149606299212598"/>
        </customSheetView>
        """
        xml = tostring(view.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestCustomChartsheetViews:
    def test_read(self, custom_chartsheet_views):
        src = """
        <customSheetViews>
            <customSheetView
                    guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}"
                    scale="88"
                    zoomToFit="1">
                <pageMargins
                        left="0.23622047244094491"
                        right="0.23622047244094491"
                        top="0.74803149606299213"
                        bottom="0.74803149606299213"
                        header="0.31496062992125984"
                        footer="0.31496062992125984"/>
                <pageSetup
                        paperSize="7"
                        orientation="landscape"
                        r:id="rId1"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                <headerFooter/>
            </customSheetView>
        </customSheetViews>
        """
        xml = fromstring(src)
        views = custom_chartsheet_views.from_tree(xml)
        assert views.customSheetView[0].state == "visible"
        assert views.customSheetView[0].scale == 88
        assert views.customSheetView[0].pageMargins.left == 0.23622047244094491

    def test_write(self, custom_chartsheet_views):
        from openpyxl.chartsheet.custom import CustomChartsheetView

        page_margins = PageMargins(
            left=0.2362204724409449,
            right=0.2362204724409449,
            top=0.7480314960629921,
            bottom=0.7480314960629921,
            header=0.3149606299212598,
            footer=0.3149606299212598,
        )
        view = CustomChartsheetView(
            guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}",
            scale=88,
            zoomToFit=1,
            pageMargins=page_margins,
        )
        views = custom_chartsheet_views(customSheetView=[view])
        expected = """
        <customSheetViews>
            <customSheetView
                    guid="{C43F44F8-8CE9-4A07-A9A9-0646C7C6B826}"
                    scale="88"
                    state="visible"
                    zoomToFit="1">
                <pageMargins
                        left="0.2362204724409449"
                        right="0.2362204724409449"
                        top="0.7480314960629921"
                        bottom="0.7480314960629921"
                        header="0.3149606299212598"
                        footer="0.3149606299212598"/>
            </customSheetView>
        </customSheetViews>
        """
        xml = tostring(views.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
