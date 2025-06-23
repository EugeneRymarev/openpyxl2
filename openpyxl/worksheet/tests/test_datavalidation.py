# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def data_validation():
    from openpyxl.worksheet.datavalidation import DataValidation

    return DataValidation


@pytest.fixture
def data_validation_list():
    from openpyxl.worksheet.datavalidation import DataValidationList

    return DataValidationList


class TestDataValidation:
    def test_ctor(self, data_validation):
        dv = data_validation(allowBlank=True)
        xml = tostring(dv.to_tree())
        expected = """
        <dataValidation
                allowBlank="1"
                showDropDown="0"
                showErrorMessage="0"
                showInputMessage="0"
                sqref=""/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, data_validation):
        src = "<root/>"
        node = fromstring(src)
        dv = data_validation.from_tree(node)
        assert dv == data_validation()

    def test_list_validation(self, data_validation):
        dv = data_validation(
            type="list",
            formula1='"Dog,Cat,Fish"',
            allowBlank=False,
            showErrorMessage=True,
            showInputMessage=True,
        )
        assert dv.formula1, '"Dog,Cat == Fish"'
        dv_dict = dict(dv)
        assert dv_dict["type"] == "list"
        assert dv_dict["allowBlank"] == "0"
        assert dv_dict["showErrorMessage"] == "1"
        assert dv_dict["showInputMessage"] == "1"
        assert dv_dict["showDropDown"] == "0"

    def test_hide_drop_down(self, data_validation):
        dv = data_validation()
        assert not dv.hide_drop_down
        dv.hide_drop_down = True
        assert dv.showDropDown is True

    def test_writer_validation(self, data_validation):
        class DummyCell:
            coordinate = "A1"

        dv = data_validation(
            type="list",
            formula1='"Dog,Cat,Fish"',
            allowBlank=False,
            showErrorMessage=True,
            showInputMessage=True,
        )
        dv.add(DummyCell())
        xml = tostring(dv.to_tree())
        expected = """
        <dataValidation
                allowBlank="0"
                showDropDown="0"
                showErrorMessage="1"
                showInputMessage="1"
                sqref="A1"
                type="list">
            <formula1>&quot;Dog,Cat,Fish&quot;</formula1>
        </dataValidation>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_sqref(self, data_validation):
        dv = data_validation(sqref="A1")
        assert dv.sqref == MultiCellRange("A1")

    def test_add_after_sqref(self, data_validation):
        class DummyCell:
            coordinate = "A2"

        dv = data_validation()
        dv.sqref = "A1"
        dv.add(DummyCell())
        assert dv.cells == MultiCellRange("A1 A2")

    def test_read_formula(self, data_validation):
        xml = """
        <dataValidation
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                allowBlank="0"
                showDropDown="0"
                showErrorMessage="0"
                showInputMessage="1"
                sqref="A1"
                type="list">
            <formula1>&quot;Dog,Cat,Fish&quot;</formula1>
        </dataValidation>
        """
        xml = fromstring(xml)
        dv = data_validation.from_tree(xml)
        assert dv.type == "list"
        assert dv.formula1 == '"Dog,Cat,Fish"'

    def test_parser(self, data_validation):
        xml = """
        <dataValidation
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                type="list"
                errorStyle="warning"
                allowBlank="1"
                showInputMessage="1"
                showErrorMessage="1"
                error="Value must be between 1 and 3!"
                errorTitle="An Error Message"
                promptTitle="Multiplier"
                prompt="for monthly or quarterly reports"
                sqref="H6">
        </dataValidation>
        """
        xml = fromstring(xml)
        dv = data_validation.from_tree(xml)
        expected = data_validation(
            error="Value must be between 1 and 3!",
            errorStyle="warning",
            errorTitle="An Error Message",
            prompt="for monthly or quarterly reports",
            promptTitle="Multiplier",
            type="list",
            allowBlank="1",
            sqref="H6",
            showErrorMessage="1",
            showInputMessage="1",
        )
        assert dv == expected

    def test_contains(self, data_validation):
        dv = data_validation(sqref="A1:D4 E5")
        assert "C2" in dv


class TestDataValidationList:
    def test_ctor(self, data_validation_list):
        dvs = data_validation_list()
        xml = tostring(dvs.to_tree())
        expected = '<dataValidations count="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, data_validation_list):
        src = "<dataValidations/>"
        node = fromstring(src)
        dvs = data_validation_list.from_tree(node)
        assert dvs == data_validation_list()

    def test_empty_dv(self, data_validation_list, data_validation):
        dv = data_validation()
        dvs = data_validation_list(dataValidation=[dv])
        xml = tostring(dvs.to_tree())
        expected = '<dataValidations count="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.mark.parametrize(
    "cells, expected",
    [
        (["A1"], "A1"),
        (["A1", "B1"], "A1 B1"),
        (["A1", "A2", "A3", "A4", "B1", "B2", "B3", "B4"], "A1:A4 B1:B4"),
        (["A2", "A4", "A3", "A1", "A5"], "A1:A5"),
        (["AA1", "AA2", "B1", "B2", "B3", "AA4", "AA3"], "B1:B3 AA1:AA4"),
    ],
)
def test_collapse_cell_addresses(cells, expected):
    from openpyxl.worksheet.datavalidation import collapse_cell_addresses

    assert collapse_cell_addresses(cells) == expected


def test_expand_cell_ranges():
    from openpyxl.worksheet.datavalidation import expand_cell_ranges

    rs = "A1:A3 B1:B3"
    assert expand_cell_ranges(rs) == {"A1", "A2", "A3", "B1", "B2", "B3"}
