# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def bubble_chart():
    from openpyxl.chart.bubble_chart import BubbleChart

    return BubbleChart


class TestBubbleChart:
    def test_ctor(self, bubble_chart):
        bc = bubble_chart()
        xml = tostring(bc.to_tree())
        expected = """
        <bubbleChart>
            <axId val="10"/>
            <axId val="20"/>
        </bubbleChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, bubble_chart):
        src = """
        <bubbleChart>
            <axId val="10"/>
            <axId val="20"/>
        </bubbleChart>
        """
        node = fromstring(src)
        bc = bubble_chart.from_tree(node)
        assert bc.axId == [10, 20]
