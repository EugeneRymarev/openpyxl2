# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def table_style():
    from openpyxl.styles.table import TableStyle

    return TableStyle


@pytest.fixture
def table_style_list():
    from openpyxl.styles.table import TableStyleList

    return TableStyleList


@pytest.fixture
def table_style_element():
    from openpyxl.styles.table import TableStyleElement

    return TableStyleElement


class TestTableStyle:
    def test_ctor(self, table_style):
        table = table_style(name="medium")
        xml = tostring(table.to_tree())
        expected = '<tableStyle name="medium"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, table_style):
        src = '<tableStyle name="medium"/>'
        node = fromstring(src)
        table = table_style.from_tree(node)
        assert table == table_style(name="medium")


class TestTableStyleList:
    def test_ctor(self, table_style_list):
        table = table_style_list()
        xml = tostring(table.to_tree())
        expected = """
        <tableStyles
                count="0"
                defaultTableStyle="TableStyleMedium9"
                defaultPivotStyle="PivotStyleLight16"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, table_style_list):
        src = "<tableStyles/>"
        node = fromstring(src)
        table = table_style_list.from_tree(node)
        assert table == table_style_list()


class TestTableStyleElement:
    def test_ctor(self, table_style_element):
        table = table_style_element(type="wholeTable", dxfId=4)
        xml = tostring(table.to_tree())
        expected = '<tableStyleElement type="wholeTable" dxfId="4"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, table_style_element):
        src = '<tableStyleElement type="secondRowStripe" size="2"/>'
        node = fromstring(src)
        table = table_style_element.from_tree(node)
        assert table == table_style_element(type="secondRowStripe", size=2)
