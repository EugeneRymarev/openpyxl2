# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def external_reference():
    from openpyxl.workbook.external_reference import ExternalReference

    return ExternalReference


class TestExternalReference:
    def test_ctor(self, external_reference):
        er = external_reference(id="rId1")
        xml = tostring(er.to_tree())
        expected = """
        <externalReference
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                r:id="rId1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, external_reference):
        src = """
        <externalReference
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                r:id="rId2"/>
        """
        node = fromstring(src)
        er = external_reference.from_tree(node)
        assert er == external_reference(id="rId2")
