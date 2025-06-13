# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def chart_container():
    from openpyxl.chart.chartspace import ChartContainer

    return ChartContainer


@pytest.fixture
def protection():
    from openpyxl.chart.chartspace import Protection

    return Protection


@pytest.fixture
def external_data():
    from openpyxl.chart.chartspace import ExternalData

    return ExternalData


@pytest.fixture
def chart_space():
    from openpyxl.chart.chartspace import ChartSpace

    return ChartSpace


class TestChartContainer:
    def test_ctor(self, chart_container):
        container = chart_container()
        xml = tostring(container.to_tree())
        expected = """
        <chart>
            <plotArea></plotArea>
            <plotVisOnly val="1"/>
            <dispBlanksAs val="gap"></dispBlanksAs>
        </chart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, chart_container):
        src = """
        <chart>
            <plotArea></plotArea>
            <dispBlanksAs val="gap"></dispBlanksAs>
        </chart>
        """
        node = fromstring(src)
        container = chart_container.from_tree(node)
        assert container == chart_container()


class TestProtection:
    def test_ctor(self, protection):
        prot = protection()
        xml = tostring(prot.to_tree())
        expected = "<protection/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, protection):
        src = """
        <protection>
            <chartObject val="1"/>
        </protection>
        """
        node = fromstring(src)
        prot = protection.from_tree(node)
        assert prot == protection(chartObject=True)


class TestExternalData:
    def test_ctor(self, external_data):
        data = external_data(id="rId1")
        xml = tostring(data.to_tree())
        expected = '<externalData id="rId1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, external_data):
        src = '<externalData id="rId1"/>'
        node = fromstring(src)
        data = external_data.from_tree(node)
        assert data == external_data(id="rId1")


class TestChartSpace:
    def test_ctor(self, chart_space, chart_container):
        cs = chart_space(chart=chart_container())
        xml = tostring(cs.to_tree())
        expected = """
        <chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <chart>
                <plotArea></plotArea>
                <plotVisOnly val="1"/>
                <dispBlanksAs val="gap"></dispBlanksAs>
            </chart>
        </chartSpace>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, chart_space, chart_container):
        src = """
        <chartSpace>
            <chart/>
        </chartSpace>
        """
        node = fromstring(src)
        cs = chart_space.from_tree(node)
        assert cs == chart_space(chart=chart_container())
