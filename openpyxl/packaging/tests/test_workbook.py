# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def workbook_package():
    from openpyxl.packaging.workbook import WorkbookPackage

    return WorkbookPackage


class TestWorkbookPackage:
    def test_ctor(self, workbook_package):
        parser = workbook_package()
        xml = tostring(parser.to_tree())
        expected = """
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <workbookPr/>
        </workbook>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, workbook_package):
        src = "<workbook/>"
        node = fromstring(src)
        parser = workbook_package.from_tree(node)
        assert parser == workbook_package()


def test_read_workbook_code_name(datadir, workbook_package):
    datadir.chdir()
    with open("workbook_russian_code_name.xml", "rb") as src:
        xml = src.read()
    node = fromstring(xml)
    parser = workbook_package.from_tree(node)
    expected = "\u042d\u0442\u0430\u041a\u043d\u0438\u0433\u0430"
    assert parser.properties.codeName == expected
