# Copyright (c) 2010-2025 openpyxl
import datetime

import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def error():
    from openpyxl.pivot.fields import Error

    return Error


@pytest.fixture
def boolean():
    from openpyxl.pivot.fields import Boolean

    return Boolean


@pytest.fixture
def missing():
    from openpyxl.pivot.fields import Missing

    return Missing


@pytest.fixture
def number():
    from openpyxl.pivot.fields import Number

    return Number


@pytest.fixture
def text():
    from openpyxl.pivot.fields import Text

    return Text


@pytest.fixture
def index():
    from openpyxl.pivot.fields import Index

    return Index


@pytest.fixture
def date_time_field():
    from openpyxl.pivot.fields import DateTimeField

    return DateTimeField


@pytest.fixture
def tuple_list():
    from openpyxl.pivot.fields import TupleList

    return TupleList


class TestError:
    def test_ctor(self, error):
        e = error(v="error")
        xml = tostring(e.to_tree())
        expected = '<e v="error"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, error):
        src = '<e v="error"/>'
        node = fromstring(src)
        e = error.from_tree(node)
        assert e == error(v="error")


class TestBoolean:
    def test_ctor(self, boolean):
        b = boolean()
        xml = tostring(b.to_tree())
        expected = '<b v="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, boolean):
        src = "<b/>"
        node = fromstring(src)
        b = boolean.from_tree(node)
        assert b == boolean()


class TestMissing:
    def test_ctor(self, missing):
        m = missing()
        xml = tostring(m.to_tree())
        expected = "<m/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, missing):
        src = "<m/>"
        node = fromstring(src)
        m = missing.from_tree(node)
        assert m == missing()


class TestNumber:
    def test_ctor(self, number):
        n = number(v=24)
        xml = tostring(n.to_tree())
        expected = '<n v="24"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, number):
        src = '<n v="15"/>'
        node = fromstring(src)
        n = number.from_tree(node)
        assert n == number(v=15)


class TestText:
    def test_ctor(self, text):
        t = text(v="UCLA")
        xml = tostring(t.to_tree())
        expected = '<s v="UCLA"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, text):
        src = '<s v="UCLA"/>'
        node = fromstring(src)
        t = text.from_tree(node)
        assert t == text(v="UCLA")


class TestIndex:
    def test_ctor(self, index):
        record = index()
        xml = tostring(record.to_tree())
        expected = '<x v="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, index):
        src = '<x v="1"/>'
        node = fromstring(src)
        record = index.from_tree(node)
        assert record == index(v=1)


class TestDateTimeField:
    def test_ctor(self, date_time_field):
        record = date_time_field(v=datetime.datetime(2016, 3, 24))
        xml = tostring(record.to_tree())
        expected = '<d v="2016-03-24T00:00:00"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, date_time_field):
        src = '<d v="2016-03-24T00:00:00"/>'
        node = fromstring(src)
        record = date_time_field.from_tree(node)
        assert record == date_time_field(v=datetime.datetime(2016, 3, 24))


class TestTupleList:
    def test_ctor(self, tuple_list):
        from openpyxl.pivot.fields import Tuple

        record = tuple_list(tpl=Tuple(item=1))
        xml = tostring(record.to_tree())
        expected = """
        <tpls>
            <tpl item="1"/>
        </tpls>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, tuple_list):
        from openpyxl.pivot.fields import Tuple

        src = """
        <tpls c="1">
            <tpl hier="1" item="4294967295"/>
        </tpls>
        """
        node = fromstring(src)
        record = tuple_list.from_tree(node)
        assert record == tuple_list(c=1, tpl=Tuple(hier=1, item=4294967295))
