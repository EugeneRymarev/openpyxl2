# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def BubbleChart():
    from ..bubble_chart import BubbleChart

    return BubbleChart


class TestBubbleChart:

    def test_ctor(self, BubbleChart):
        bubble_chart = BubbleChart()
        xml = tostring(bubble_chart.to_tree())
        expected = """
        <bubbleChart>
          <axId val="10" />
          <axId val="20" />
        </bubbleChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, BubbleChart):
        src = """
        <bubbleChart>
          <axId val="10" />
          <axId val="20" />
        </bubbleChart>
        """
        node = fromstring(src)
        bubble_chart = BubbleChart.from_tree(node)
        assert bubble_chart.axId == [10, 20]
