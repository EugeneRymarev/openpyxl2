# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def web_publish_item():
    from openpyxl.chartsheet.publish import WebPublishItem

    return WebPublishItem


@pytest.fixture
def web_publish_items():
    from openpyxl.chartsheet.publish import WebPublishItems

    return WebPublishItems


class TestWebPublishItem:
    def test_read(self, web_publish_item):
        src = r"""
        <webPublishItem
                id="6433"
                divId="Views_6433"
                sourceType="chart"
                sourceRef=""
                sourceObject="Chart 1"
                destinationFile="D:\Publish.mht"
                autoRepublish="0"/>
        """
        xml = fromstring(src)
        item = web_publish_item.from_tree(xml)
        assert item.id == 6433
        assert item.sourceObject == "Chart 1"

    def test_write(self, web_publish_item):
        item = web_publish_item(
            id=6433,
            divId="Views_6433",
            sourceType="chart",
            sourceRef="",
            sourceObject="Chart 1",
            destinationFile=r"D:\Publish.mht",
            title="First Chart",
            autoRepublish=False,
        )
        expected = r"""
        <webPublishItem
                id="6433"
                divId="Views_6433"
                sourceType="chart"
                sourceRef=""
                sourceObject="Chart 1"
                destinationFile="D:\Publish.mht"
                title="First Chart"
                autoRepublish="0"/>
        """
        xml = tostring(item.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestWebPublishItems:
    def test_read(self, web_publish_items):
        src = r"""
        <webPublishItems count="1">
            <webPublishItem
                    id="6433"
                    divId="Views_6433"
                    sourceType="chart"
                    sourceRef=""
                    sourceObject="Chart 1"
                    destinationFile="D:\Publish.mht"
                    autoRepublish="0"/>
        </webPublishItems>
        """
        xml = fromstring(src)
        items = web_publish_items.from_tree(xml)
        assert items.count == 1
        assert items.webPublishItem[0].sourceObject == "Chart 1"

    def test_write(self, web_publish_items):
        from openpyxl.chartsheet.publish import WebPublishItem

        item = WebPublishItem(
            id=6433,
            divId="Views_6433",
            sourceType="chart",
            sourceRef="",
            sourceObject="Chart 1",
            destinationFile=r"D:\Publish.mht",
            title="First Chart",
            autoRepublish=False,
        )
        item2 = WebPublishItem(
            id=64487,
            divId="Views_64487",
            sourceType="chart",
            sourceRef="Ref_545421",
            sourceObject="Chart 15",
            destinationFile=r"D:\Publish_12.mht",
            title="Second Chart",
            autoRepublish=True,
        )
        items = web_publish_items(webPublishItem=[item, item2])
        expected = r"""
        <WebPublishItems count="2">
            <webPublishItem
                    id="6433"
                    divId="Views_6433"
                    sourceType="chart"
                    sourceRef=""
                    sourceObject="Chart 1"
                    destinationFile="D:\Publish.mht"
                    title="First Chart"
                    autoRepublish="0"/>
            <webPublishItem
                    id="64487"
                    divId="Views_64487"
                    sourceType="chart"
                    sourceRef="Ref_545421"
                    sourceObject="Chart 15"
                    destinationFile="D:\Publish_12.mht"
                    title="Second Chart"
                    autoRepublish="1"/>
        </WebPublishItems>
        """
        xml = tostring(items.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
