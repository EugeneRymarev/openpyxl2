# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def workbook_properties():
    from openpyxl.workbook.properties import WorkbookProperties

    return WorkbookProperties


@pytest.fixture
def calc_properties():
    from openpyxl.workbook.properties import CalcProperties

    return CalcProperties


@pytest.fixture
def file_version():
    from openpyxl.workbook.properties import FileVersion

    return FileVersion


class TestWorkbookProperties:
    def test_ctor(self, workbook_properties):
        props = workbook_properties()
        xml = tostring(props.to_tree())
        expected = "<workbookPr/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, workbook_properties):
        src = "<workbookPr/>"
        node = fromstring(src)
        props = workbook_properties.from_tree(node)
        assert props == workbook_properties()


class TestCalcProperties:
    def test_ctor(self, calc_properties):
        calc = calc_properties()
        xml = tostring(calc.to_tree())
        expected = '<calcPr calcId="124519" fullCalcOnLoad="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, calc_properties):
        src = "<calcPr/>"
        node = fromstring(src)
        calc = calc_properties.from_tree(node)
        assert calc == calc_properties()


class TestFileVersion:
    def test_ctor(self, file_version):
        prop = file_version()
        xml = tostring(prop.to_tree())
        expected = "<fileVersion/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, file_version):
        src = "<fileVersion/>"
        node = fromstring(src)
        prop = file_version.from_tree(node)
        assert prop == file_version()
