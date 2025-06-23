# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def smart_tag():
    from openpyxl.workbook.smart_tags import SmartTag

    return SmartTag


@pytest.fixture
def smart_tag_list():
    from openpyxl.workbook.smart_tags import SmartTagList

    return SmartTagList


@pytest.fixture
def smart_tag_properties():
    from openpyxl.workbook.smart_tags import SmartTagProperties

    return SmartTagProperties


class TestSmartTag:
    def test_ctor(self, smart_tag):
        smart_tags = smart_tag()
        xml = tostring(smart_tags.to_tree())
        expected = "<smartTagType/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, smart_tag):
        src = "<smartTagType/>"
        node = fromstring(src)
        smart_tags = smart_tag.from_tree(node)
        assert smart_tags == smart_tag()


class TestSmartTagList:
    def test_ctor(self, smart_tag_list):
        smart_tags = smart_tag_list()
        xml = tostring(smart_tags.to_tree())
        expected = "<smartTagTypes/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, smart_tag_list):
        src = "<smartTagTypes/>"
        node = fromstring(src)
        smart_tags = smart_tag_list.from_tree(node)
        assert smart_tags == smart_tag_list()


class TestSmartTagProperties:
    def test_ctor(self, smart_tag_properties):
        smart_tags = smart_tag_properties()
        xml = tostring(smart_tags.to_tree())
        expected = "<smartTagPr/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, smart_tag_properties):
        src = "<smartTagPr/>"
        node = fromstring(src)
        smart_tags = smart_tag_properties.from_tree(node)
        assert smart_tags == smart_tag_properties()
