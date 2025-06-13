# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def legend_():
    from openpyxl.chart.legend import Legend

    return Legend


@pytest.fixture
def legend_entry():
    from openpyxl.chart.legend import LegendEntry

    return LegendEntry


class TestLegend:

    def test_ctor(self, legend_):
        legend = legend_()
        xml = tostring(legend.to_tree())
        expected = """
        <legend>
            <legendPos val="r"/>
            <overlay val="0"/>
        </legend>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, legend_):
        src = """
        <legend>
            <legendPos val="r"/>
        </legend>
        """
        node = fromstring(src)
        legend = legend_.from_tree(node)
        assert legend == legend_()


class TestLegendEntry:

    def test_ctor(self, legend_entry):
        legend = legend_entry(idx=0, delete=True)
        xml = tostring(legend.to_tree())
        expected = """
        <legendEntry>
            <idx val="0"/>
            <delete val="1"/>
        </legendEntry>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, legend_entry):
        src = """
        <legendEntry>
            <idx val="0"></idx>
        </legendEntry>
        """
        node = fromstring(src)
        legend = legend_entry.from_tree(node)
        assert legend == legend_entry()
