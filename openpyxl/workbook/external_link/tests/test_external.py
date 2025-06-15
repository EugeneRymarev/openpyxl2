# Copyright (c) 2010-2025 openpyxl
import zipfile

import pytest
from openpyxl.packaging.relationship import Relationship
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.constants import ARC_WORKBOOK_RELS
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def external_cell():
    from openpyxl.workbook.external_link.external import ExternalCell

    return ExternalCell


@pytest.fixture
def external_link():
    from openpyxl.workbook.external_link.external import ExternalLink

    return ExternalLink


@pytest.fixture
def external_book():
    from openpyxl.workbook.external_link.external import ExternalBook

    return ExternalBook


class TestExternalCell:
    def test_read(self, external_cell):
        src = """
        <cell r="B1" t="str">
            <v>D&#0252;sseldorf</v>
        </cell>
        """
        node = fromstring(src)
        cell = external_cell.from_tree(node)
        assert cell.v == "D\xfcsseldorf"


class TestExternalLink:
    def test_ctor(self, external_link):
        src = """
        <externalLink
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <externalBook  r:id="rId1"/>
        </externalLink>
        """
        node = fromstring(src)
        link = external_link.from_tree(node)
        assert link.externalBook.id == "rId1"

    def test_write(self, external_link):
        expected = """
        <externalLink
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        </externalLink>
        """
        link = external_link()
        link.file_link = Relationship(Target="somefile.xlsx", type="externalLink")
        xml = tostring(link.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_path(self, external_link):
        link = external_link()
        assert link.path == "/xl/externalLinks/externalLinkNone.xml"


class TestExternalBook:
    def test_ctor(self, external_book):
        from openpyxl.workbook.external_link.external import ExternalDefinedName
        from openpyxl.workbook.external_link.external import ExternalSheetNames

        book = external_book()
        book.sheetNames = ExternalSheetNames(sheetName=["Sheet1", "Sheet2", "Sheet3"])
        df = ExternalDefinedName(name="B2range", refersTo="='Sheet1'!$A$1:$A$10")
        book.definedNames = [df]
        xml = tostring(book.to_tree())
        expected = """
        <externalBook>
            <sheetNames>
                <sheetName val="Sheet1"/>
                <sheetName val="Sheet2"/>
                <sheetName val="Sheet3"/>
            </sheetNames>
            <definedNames>
                <definedName name="B2range" refersTo="='Sheet1'!$A$1:$A$10"/>
            </definedNames>
        </externalBook>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_read(self, external_book):
        src = """
        <externalBook>
            <sheetNames>
                <sheetName val="Sheet1"/>
                <sheetName val="Sheet2"/>
                <sheetName val="Sheet3"/>
            </sheetNames>
            <definedNames>
                <definedName name="B2range" refersTo="='Sheet1'!$A$1:$A$10"/>
            </definedNames>
        </externalBook>
        """
        node = fromstring(src)
        book = external_book.from_tree(node)
        assert book.definedNames[0].name == "B2range"
        assert book.definedNames[0].refersTo == "='Sheet1'!$A$1:$A$10"


def test_read_ole_link(datadir, external_link):
    datadir.chdir()
    with open("OLELink.xml") as src:
        node = fromstring(src.read())
    link = external_link.from_tree(node)
    assert link.externalBook is None


def test_read_external_link(datadir):
    from openpyxl.packaging.relationship import get_dependents
    from openpyxl.workbook.external_link.external import read_external_link

    datadir.chdir()
    archive = zipfile.ZipFile("book1.xlsx")
    rels = get_dependents(archive, ARC_WORKBOOK_RELS)
    rel = rels.get("rId4")
    book = read_external_link(archive, rel.Target)
    assert book.file_link.Target == "book2.xlsx"


def test_write_workbook(datadir, tmpdir):
    from openpyxl.reader.excel import load_workbook

    datadir.chdir()
    src = zipfile.ZipFile("book1.xlsx")
    orig_files = set(src.namelist())
    src.close()
    wb = load_workbook("book1.xlsx", keep_links=True)
    tmpdir.chdir()
    wb.save("book1.xlsx")
    src = zipfile.ZipFile("book1.xlsx")
    out_files = set(src.namelist())
    src.close()
    # remove files from archive that the other can't have
    out_files.discard("xl/sharedStrings.xml")
    orig_files.discard("xl/calcChain.xml")
    assert orig_files == out_files
