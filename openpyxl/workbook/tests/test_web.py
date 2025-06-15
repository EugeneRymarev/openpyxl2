# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def web_publish_object():
    from openpyxl.workbook.web import WebPublishObject

    return WebPublishObject


@pytest.fixture
def web_publish_object_list():
    from openpyxl.workbook.web import WebPublishObjectList

    return WebPublishObjectList


@pytest.fixture
def web_publishing():
    from openpyxl.workbook.web import WebPublishing

    return WebPublishing


class TestWebPublishObject:
    def test_ctor(self, web_publish_object):
        obj = web_publish_object(id=1, divId="main", destinationFile="www")
        xml = tostring(obj.to_tree())
        expected = '<webPublishingObject destinationFile="www" divId="main" id="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, web_publish_object):
        src = '<webPublishingObject destinationFile="www" divId="main" id="1"/>'
        node = fromstring(src)
        obj = web_publish_object.from_tree(node)
        assert obj == web_publish_object(id=1, divId="main", destinationFile="www")


class TestWebPublishObjectList:
    def test_ctor(self, web_publish_object_list):
        objs = web_publish_object_list()
        xml = tostring(objs.to_tree())
        expected = "<webPublishingObjects/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, web_publish_object_list):
        src = "<webPublishingObjects/>"
        node = fromstring(src)
        objs = web_publish_object_list.from_tree(node)
        assert objs == web_publish_object_list()


class TestWebPublishing:
    def test_ctor(self, web_publishing):
        web = web_publishing()
        xml = tostring(web.to_tree())
        expected = '<webPublishing targetScreenSize="800x600"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, web_publishing):
        src = "<webPublishing/>"
        node = fromstring(src)
        web = web_publishing.from_tree(node)
        assert web == web_publishing()
