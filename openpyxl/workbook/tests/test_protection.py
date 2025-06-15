# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def workbook_protection():
    from openpyxl.workbook.protection import WorkbookProtection

    return WorkbookProtection


@pytest.fixture
def file_sharing():
    from openpyxl.workbook.protection import FileSharing

    return FileSharing


class TestWorkbookProtection:
    def test_ctor(self, workbook_protection):
        propt = workbook_protection()
        xml = tostring(propt.to_tree())
        expected = "<workbookPr/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_ctor_with_passwords(self, workbook_protection):
        prot = workbook_protection(
            workbookPassword="secret", revisionsPassword="secret"
        )
        assert prot.workbookPassword == "DAA7"
        assert prot.revisionsPassword == "DAA7"

    def test_from_xml(self, workbook_protection):
        src = """
        <workbookProtection
                workbookAlgorithmName="SHA-512"
                workbookHashValue="wDZaZrfM8uKpKghbfws7rY7pmVoOwHjy5qg5d2ABHdSMtH1y0IIkgwJT5Hl2lacSw1sNusImGBUQs/sHcql3hw=="
                workbookSaltValue="ah1OevWahpb3tQiJO3qrnQ=="
                workbookSpinCount="100000"
                lockStructure="1"
                workbookPassword="1234"
                revisionsPassword="ABCD"/>
        """
        node = fromstring(src)
        prot = workbook_protection.from_tree(node)
        expected = workbook_protection(
            workbookAlgorithmName="SHA-512",
            workbookHashValue="wDZaZrfM8uKpKghbfws7rY7pmVoOwHjy5qg5d2ABHdSMtH1y0IIkgwJT5Hl2lacSw1sNusImGBUQs/sHcql3hw==",
            workbookSaltValue="ah1OevWahpb3tQiJO3qrnQ==",
            workbookSpinCount=100000,
            lockStructure="1",
        )
        expected.set_workbook_password("1234", already_hashed=True)
        expected.set_revisions_password("ABCD", already_hashed=True)
        assert prot == expected


class TestFileSharing:
    def test_ctor(self, file_sharing):
        share = file_sharing(readOnlyRecommended=True)
        xml = tostring(share.to_tree())
        expected = '<fileSharing readOnlyRecommended="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, file_sharing):
        src = '<fileSharing userName="Alice"/>'
        node = fromstring(src)
        share = file_sharing.from_tree(node)
        assert share == file_sharing(userName="Alice")
