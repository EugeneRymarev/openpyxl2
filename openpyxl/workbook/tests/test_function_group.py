# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def function_group():
    from openpyxl.workbook.function_group import FunctionGroup

    return FunctionGroup


@pytest.fixture
def function_group_list():
    from openpyxl.workbook.function_group import FunctionGroupList

    return FunctionGroupList


class TestFunctionGroup:
    def test_ctor(self, function_group):
        fg = function_group(name="Statistics")
        xml = tostring(fg.to_tree())
        expected = '<functionGroup name="Statistics"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, function_group):
        src = '<functionGroup name="Database"/>'
        node = fromstring(src)
        fg = function_group.from_tree(node)
        assert fg == function_group(name="Database")


class TestFunctionGroupList:
    def test_ctor(self, function_group_list):
        fg = function_group_list()
        xml = tostring(fg.to_tree())
        expected = '<functionGroups builtInGroupCount="16"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, function_group_list):
        src = "<functionGroups/>"
        node = fromstring(src)
        fg = function_group_list.from_tree(node)
        assert fg == function_group_list()
