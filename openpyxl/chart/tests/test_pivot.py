# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def pivot_source():
    from openpyxl.chart.pivot import PivotSource

    return PivotSource


@pytest.fixture
def pivot_format():
    from openpyxl.chart.pivot import PivotFormat

    return PivotFormat


class TestPivotSource:
    def test_ctor(self, pivot_source):
        fut = pivot_source(name="[template.xlsx]PIVOT!PivotTable6", fmtId=0)
        xml = tostring(fut.to_tree())
        expected = """
        <pivotSource>
            <name>
                [template.xlsx]PIVOT!PivotTable6
            </name>
            <fmtId val="0"/>
        </pivotSource>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pivot_source):
        src = """
        <pivotSource>
            <name>[template.xlsx]PIVOT!PivotTable6</name>
            <fmtId val="0"/>
        </pivotSource>
        """
        node = fromstring(src)
        fut = pivot_source.from_tree(node)
        assert fut == pivot_source(name="[template.xlsx]PIVOT!PivotTable6", fmtId=0)


class TestPivotFormat:
    def test_ctor(self, pivot_format):
        fmt = pivot_format()
        xml = tostring(fmt.to_tree())
        expected = """
        <pivotFmt>
            <idx val="0"/>
        </pivotFmt>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pivot_format):
        src = """
        <pivotFmt>
            <idx val="0"/>
        </pivotFmt>
        """
        node = fromstring(src)
        fmt = pivot_format.from_tree(node)
        assert fmt == pivot_format()
