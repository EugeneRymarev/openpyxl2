# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def scatter_chart():
    from openpyxl.chart.scatter_chart import ScatterChart

    return ScatterChart


class TestScatterChart:
    def test_ctor(self, scatter_chart):
        chart = scatter_chart()
        xml = tostring(chart.to_tree())
        expected = """
        <scatterChart>
            <axId val="10"></axId>
            <axId val="20"></axId>
        </scatterChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, scatter_chart):
        src = """
        <scatterChart>
            <axId val="10"></axId>
            <axId val="20"></axId>
        </scatterChart>
        """
        node = fromstring(src)
        chart = scatter_chart.from_tree(node)
        assert chart.axId == [10, 20]
