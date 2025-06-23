# Copyright (c) 2010-2025 openpyxl
import io
import zipfile

import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.filters import FilterColumn
from openpyxl.worksheet.filters import Filters
from openpyxl.worksheet.related import Related
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def table_column():
    from openpyxl.worksheet.table import TableColumn

    return TableColumn


@pytest.fixture
def table():
    from openpyxl.worksheet.table import Table

    return Table


@pytest.fixture
def table_formula():
    from openpyxl.worksheet.table import TableFormula

    return TableFormula


@pytest.fixture
def table_style_info():
    from openpyxl.worksheet.table import TableStyleInfo

    return TableStyleInfo


@pytest.fixture
def xml_column_props():
    from openpyxl.worksheet.table import XMLColumnProps

    return XMLColumnProps


@pytest.fixture
def table_part_list():
    from openpyxl.worksheet.table import TablePartList

    return TablePartList


@pytest.fixture
def table_list():
    from openpyxl.worksheet.table import TableList

    return TableList


class TestTableColumn:
    def test_ctor(self, table_column):
        col = table_column(id=1, name="Column1")
        xml = tostring(col.to_tree())
        expected = '<tableColumn id="1" name="Column1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, table_column):
        src = '<tableColumn id="1" name="Column1"/>'
        node = fromstring(src)
        col = table_column.from_tree(node)
        assert col == table_column(id=1, name="Column1")


class TestTable:
    def test_ctor(self, table, table_column):
        t = table(displayName="A_Sample_Table", ref="A1:D5")
        xml = tostring(t.to_tree())
        expected = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               displayName="A_Sample_Table"
               headerRowCount="1"
               name="A_Sample_Table"
               id="1"
               ref="A1:D5">
        </table>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_columns(self, table):
        t = table(displayName="A_Sample_Table", ref="A1:D5")
        t._initialise_columns()
        xml = tostring(t.to_tree())
        expected = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               displayName="A_Sample_Table"
               headerRowCount="1"
               name="A_Sample_Table"
               id="1"
               ref="A1:D5">
            <autoFilter ref="A1:D5"/>
            <tableColumns count="4">
                <tableColumn id="1" name="Column1"/>
                <tableColumn id="2" name="Column2"/>
                <tableColumn id="3" name="Column3"/>
                <tableColumn id="4" name="Column4"/>
            </tableColumns>
        </table>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_preserve_existing_filter(self, table):
        t = table(displayName="A_Sample_Table", ref="A1:D5")
        filters = Filters(blank=False, filter=["16"])
        col = FilterColumn(colId=0, filters=filters)
        t.autoFilter = AutoFilter(ref="A1:D5", filterColumn=[col])
        t._initialise_columns()
        assert t.autoFilter.filterColumn == [col]

    def test_column_names(self, table):
        t = table(displayName="A_Sample_Table", ref="A10:D14")
        t._initialise_columns()
        assert t.column_names == ["Column1", "Column2", "Column3", "Column4"]

    def test_from_xml(self, table):
        src = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               id="1"
               name="Table1"
               displayName="Table1"
               ref="A1:AA27">
        </table>
        """
        node = fromstring(src)
        t = table.from_tree(node)
        assert t == table(displayName="Table1", name="Table1", ref="A1:AA27")

    def test_path(self, table):
        t = table(displayName="Table1", ref="A1:M6")
        assert t.path == "/xl/tables/table1.xml"

    def test_write(self, table):
        out = io.BytesIO()
        archive = zipfile.ZipFile(out, "w")
        t = table(displayName="Table1", ref="B1:L10")
        t._write(archive)
        assert "xl/tables/table1.xml" in archive.namelist()


class TestTableFormula:
    def test_ctor(self, table_formula):
        formula = table_formula()
        formula.text = "=A1*4"
        xml = tostring(formula.to_tree())
        expected = "<tableFormula>=A1*4</tableFormula>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, table_formula):
        src = "<tableFormula>=A1*4</tableFormula>"
        node = fromstring(src)
        formula = table_formula.from_tree(node)
        assert formula.text == "=A1*4"


class TestTableInfo:
    def test_ctor(self, table_style_info):
        info = table_style_info(name="TableStyleMedium12")
        xml = tostring(info.to_tree())
        expected = '<tableStyleInfo name="TableStyleMedium12"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, table_style_info):
        src = '<tableStyleInfo name="TableStyleLight1" showRowStripes="1"/>'
        node = fromstring(src)
        info = table_style_info.from_tree(node)
        assert info == table_style_info(name="TableStyleLight1", showRowStripes=True)


class TestXMLColumnPr:
    def test_ctor(self, xml_column_props):
        col = xml_column_props(
            mapId="1", xpath="/xml/foo/element", xmlDataType="string"
        )
        xml = tostring(col.to_tree())
        expected = """
        <xmlColumnPr
                mapId="1"
                xpath="/xml/foo/element"
                xmlDataType="string"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, xml_column_props):
        src = """
        <xmlColumnPr
                mapId="1"
                xpath="/xml/foo/element"
                xmlDataType="string"/>
        """
        node = fromstring(src)
        col = xml_column_props.from_tree(node)
        expected = xml_column_props(
            mapId="1",
            xpath="/xml/foo/element",
            xmlDataType="string",
        )
        assert col == expected


class TestTablePartList:
    def test_ctor(self, table_part_list):
        tables = table_part_list()
        tables.append(Related(id="rId1"))
        xml = tostring(tables.to_tree())
        expected = """
        <tableParts
                count="1"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <tablePart r:id="rId1"/>
        </tableParts>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, table_part_list):
        src = """
        <tableParts
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <tablePart r:id="rId1"/>
              <tablePart r:id="rId2"/>
        </tableParts>
        """
        node = fromstring(src)
        tables = table_part_list.from_tree(node)
        assert len(tables.tablePart) == 2


class TestTableList:
    def test_append(self, table, table_list):
        tablelist = table_list()
        table1 = table(displayName="Table1", ref="A1:C10")
        tablelist.add(table1)
        assert len(tablelist) == 1

    def test_get(self, table, table_list):
        tablelist = table_list()
        table1 = table(displayName="Table1", ref="A1:C10")
        tablelist.add(table1)
        assert table1 == tablelist["Table1"]

    def test_get_by_range(self, table, table_list):
        tablelist = table_list()
        table1 = table(displayName="Table1", ref="A1:D10")
        tablelist.add(table1)
        assert True == isinstance(tablelist.get(table_range="A1:D10"), table)

    def test_add_type_error(self, table, table_list):
        tablelist = table_list()
        with pytest.raises(TypeError):
            tablelist.add("Not a Table")

    def test_get_table_does_not_exists(self, table, table_list):
        tablelist2 = table_list()
        with pytest.raises(KeyError):
            _ = tablelist2["NoTable"]

    def test_items(self, table, table_list):
        table1 = table(displayName="Table1", ref="A1:D10")
        table2 = table(displayName="Table2", ref="A1:D10")
        tablelist = table_list()
        tablelist.add(table1)
        tablelist.add(table2)
        assert tablelist.items() == [("Table1", "A1:D10"), ("Table2", "A1:D10")]
