# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def Legend():
    from ..legend import Legend

    return Legend


class TestLegend:

    def test_ctor(self, Legend):
        legend = Legend()
        xml = tostring(legend.to_tree())
        expected = """
        <legend>
          <legendPos val="r" />
        </legend>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, Legend):
        src = """
        <legend>
          <legendPos val="r" />
        </legend>
        """
        node = fromstring(src)
        legend = Legend.from_tree(node)
        assert legend == Legend()


@pytest.fixture
def LegendEntry():
    from ..legend import LegendEntry

    return LegendEntry


class TestLegendEntry:

    def test_ctor(self, LegendEntry):
        legend = LegendEntry(idx=0, delete=True)
        xml = tostring(legend.to_tree())
        expected = """
        <legendEntry>
          <idx val="0" />
          <delete val= "1" />
        </legendEntry>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, LegendEntry):
        src = """
        <legendEntry>
          <idx val="0"></idx>
        </legendEntry>
        """
        node = fromstring(src)
        legend = LegendEntry.from_tree(node)
        assert legend == LegendEntry()
