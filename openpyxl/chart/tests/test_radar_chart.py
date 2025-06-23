# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def radar_chart():
    from openpyxl.chart.radar_chart import RadarChart

    return RadarChart


class TestRadarChart:
    def test_ctor(self, radar_chart):
        chart = radar_chart()
        xml = tostring(chart.to_tree())
        expected = """
        <radarChart>
            <radarStyle val="standard"/>
            <axId val="10"></axId>
            <axId val="100"></axId>
        </radarChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, radar_chart):
        src = """
        <radarChart>
            <radarStyle val="marker"/>
            <varyColors val="0"/>
            <axId val="2107159976"/>
            <axId val="2107207992"/>
        </radarChart>
        """
        node = fromstring(src)
        chart = radar_chart.from_tree(node)
        assert dict(chart) == {}
        assert chart.type == "marker"
        assert chart.axId == [2107159976, 2107207992]
