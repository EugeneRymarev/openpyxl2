# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def chartsheet_properties():
    from openpyxl.chartsheet.properties import ChartsheetProperties

    return ChartsheetProperties


class TestChartsheetPr:
    def test_read(self, chartsheet_properties):
        src = """
        <sheetPr codeName="Chart1">
            <tabColor rgb="FFDCD8F4"/>
        </sheetPr>
        """
        xml = fromstring(src)
        properties = chartsheet_properties.from_tree(xml)
        assert properties.codeName == "Chart1"
        assert properties.tabColor.rgb == "FFDCD8F4"

    def test_write(self, chartsheet_properties):
        from openpyxl.styles.colors import Color

        properties = chartsheet_properties()
        properties.codeName = "Chart Openpyxl"
        properties.tabColor = Color(rgb="FFFFFFF4")
        expected = """
        <sheetPr codeName="Chart Openpyxl">
            <tabColor rgb="FFFFFFF4"/>
        </sheetPr>
        """
        xml = tostring(properties.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff
