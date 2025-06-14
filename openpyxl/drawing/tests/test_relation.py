# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def chart_relation():
    from openpyxl.drawing.graphic import ChartRelation

    return ChartRelation


class TestChartRelation:
    def test_ctor(self, chart_relation):
        rel = chart_relation("rId1")
        xml = tostring(rel.to_tree())
        expected = """
        <c:chart
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                r:id="rId1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, chart_relation):
        src = """
        <c:chart
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                r:id="rId1"/>
        """
        node = fromstring(src)
        rel = chart_relation.from_tree(node)
        assert rel == chart_relation("rId1")
