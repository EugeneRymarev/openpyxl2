# Copyright (c) 2010-2025 openpyxl
import io
import zipfile

import pytest

from openpyxl.packaging.manifest import Manifest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def record():
    from openpyxl.pivot.record import Record

    return Record


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
def record_list():
    from openpyxl.pivot.record import RecordList

    return RecordList


class TestRecord:
    def test_ctor(self, record, number, text, index):
        n = [number(v=1), number(v=25)]
        s = [text(v="2014-03-24")]
        x = [index(), index(), index()]
        fields = n + s + x
        field = record(_fields=fields)
        xml = tostring(field.to_tree())
        expected = """
        <r>
            <n v="1"/>
            <n v="25"/>
            <s v="2014-03-24"/>
            <x v="0"/>
            <x v="0"/>
            <x v="0"/>
        </r>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, record, number, text, index):
        src = """
        <r>
            <n v="1"/>
            <x v="0"/>
            <s v="2014-03-24"/>
            <x v="0"/>
            <n v="25"/>
            <x v="0"/>
        </r>
        """
        node = fromstring(src)
        n = [number(v=1), number(v=25)]
        s = [text(v="2014-03-24")]
        x = [index(), index(), index()]
        fields = [
            number(v=1),
            index(),
            text(v="2014-03-24"),
            index(),
            number(v=25),
            index(),
        ]
        field = record.from_tree(node)
        assert field == record(_fields=fields)


class TestRecordList:
    def test_ctor(self, record_list):
        cache = record_list()
        xml = tostring(cache.to_tree())
        expected = """
        <pivotCacheRecords
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                count="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, record_list):
        src = '<pivotCacheRecords count="0"/>'
        node = fromstring(src)
        cache = record_list.from_tree(node)
        assert cache == record_list()

    def test_write(self, record_list):
        out = io.BytesIO()
        archive = zipfile.ZipFile(out, mode="w")
        manifest = Manifest()
        records = record_list()
        xml = tostring(records.to_tree())
        records._write(archive, manifest)
        manifest.append(records)
        assert archive.namelist() == [records.path[1:]]
        assert manifest.find(records.mime_type)
