# Copyright (c) 2010-2025 openpyxl
import io
import mimetypes
import zipfile

import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.constants import WORKSHEET_TYPE
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def file_extension():
    from openpyxl.packaging.manifest import FileExtension

    return FileExtension


@pytest.fixture
def override():
    from openpyxl.packaging.manifest import Override

    return Override


@pytest.fixture
def manifest():
    from openpyxl.packaging.manifest import Manifest

    return Manifest


class TestFileExtension:
    def test_ctor(self, file_extension):
        ext = file_extension(ContentType="application/xml", Extension="xml")
        xml = tostring(ext.to_tree())
        expected = '<Default ContentType="application/xml" Extension="xml"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, file_extension):
        src = '<Default ContentType="application/xml" Extension="xml"/>'
        node = fromstring(src)
        ext = file_extension.from_tree(node)
        assert ext == file_extension(ContentType="application/xml", Extension="xml")


class TestOverride:
    def test_ctor(self, override):
        o = override(
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            PartName="/xl/workbook.xml",
        )
        xml = tostring(o.to_tree())
        expected = """
        <Override 
                ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                PartName="/xl/workbook.xml"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, override):
        src = """
        <Override
                ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                PartName="/xl/workbook.xml"/>
        """
        node = fromstring(src)
        o = override.from_tree(node)
        expected = override(
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            PartName="/xl/workbook.xml",
        )
        assert o == expected


class TestManifest:
    def test_ctor(self, manifest):
        m = manifest()
        xml = tostring(m.to_tree())
        expected = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default
                    ContentType="application/vnd.openxmlformats-package.relationships+xml"
                    Extension="rels"/>
            <Default ContentType="application/xml" Extension="xml"/>
            <Override
                    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
                    PartName="/xl/styles.xml"/>
            <Override
                    ContentType="application/vnd.openxmlformats-officedocument.theme+xml"
                    PartName="/xl/theme/theme1.xml"/>
            <Override
                    ContentType="application/vnd.openxmlformats-package.core-properties+xml"
                    PartName="/docProps/core.xml"/>
            <Override
                    ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"
                    PartName="/docProps/app.xml"/>
        </Types>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_mimetypes_init(self, manifest):
        mimetypes.init()
        m = manifest()
        # add some random xml file so manifest will update itself according
        # to the mime database entry for the extension .xml, which has been
        # changed to text/xml by the init call above
        m._register_mimetypes(["dummy.xml"])
        # reset to our correct type, so it won't interfere with unrelated tests
        mimetypes.add_type("application/xml", ".xml")
        xml = tostring(m.to_tree())
        expected = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default
                    ContentType="application/vnd.openxmlformats-package.relationships+xml"
                    Extension="rels"/>
            <Default ContentType="application/xml" Extension="xml"/>
            <Override
                    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
                    PartName="/xl/styles.xml"/>
            <Override
                    ContentType="application/vnd.openxmlformats-officedocument.theme+xml"
                    PartName="/xl/theme/theme1.xml"/>
            <Override
                    ContentType="application/vnd.openxmlformats-package.core-properties+xml"
                    PartName="/docProps/core.xml"/>
            <Override
                    ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"
                    PartName="/docProps/app.xml"/>
        </Types>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, datadir, manifest):
        datadir.chdir()
        with open("manifest.xml") as src:
            node = fromstring(src.read())
        m = manifest.from_tree(node)
        assert len(m.Default) == 2
        defaults = [
            ("application/xml", "xml"),
            ("application/vnd.openxmlformats-package.relationships+xml", "rels"),
        ]
        assert [(ct.ContentType, ct.Extension) for ct in m.Default] == defaults
        overrides = [
            (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "/xl/workbook.xml",
            ),
            (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                "/xl/worksheets/sheet1.xml",
            ),
            (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
                "/xl/chartsheets/sheet1.xml",
            ),
            (
                "application/vnd.openxmlformats-officedocument.theme+xml",
                "/xl/theme/theme1.xml",
            ),
            (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                "/xl/styles.xml",
            ),
            (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                "/xl/sharedStrings.xml",
            ),
            (
                "application/vnd.openxmlformats-officedocument.drawing+xml",
                "/xl/drawings/drawing1.xml",
            ),
            (
                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                "/xl/charts/chart1.xml",
            ),
            (
                "application/vnd.openxmlformats-package.core-properties+xml",
                "/docProps/core.xml",
            ),
            (
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                "/docProps/app.xml",
            ),
        ]
        assert [(ct.ContentType, ct.PartName) for ct in m.Override] == overrides

    def test_filenames(self, datadir, manifest):
        datadir.chdir()
        with open("manifest.xml") as src:
            node = fromstring(src.read())
        m = manifest.from_tree(node)
        expected = [
            "/xl/workbook.xml",
            "/xl/worksheets/sheet1.xml",
            "/xl/chartsheets/sheet1.xml",
            "/xl/theme/theme1.xml",
            "/xl/styles.xml",
            "/xl/sharedStrings.xml",
            "/xl/drawings/drawing1.xml",
            "/xl/charts/chart1.xml",
            "/docProps/core.xml",
            "/docProps/app.xml",
        ]
        assert m.filenames == expected

    def test_exts(self, datadir, manifest):
        datadir.chdir()
        with open("manifest.xml") as src:
            node = fromstring(src.read())
        m = manifest.from_tree(node)
        assert m.extensions == [("xml", "application/xml")]

    def test_no_dupe_overrides(self, manifest):
        m = manifest()
        assert len(m.Override) == 4
        m.Override.append("a")
        m.Override.append("a")
        assert len(m.Override) == 5

    def test_no_dupe_types(self, manifest):
        m = manifest()
        assert len(m.Default) == 2
        m.Default.append("a")
        m.Default.append("a")
        assert len(m.Default) == 3

    def test_append(self, manifest):
        from openpyxl.workbook.workbook import Workbook

        wb = Workbook()
        ws = wb.active
        m = manifest()
        m.append(ws)
        assert len(m.Override) == 5

    def test_write(self, manifest):
        from openpyxl.workbook.workbook import Workbook

        mf = manifest()
        wb = Workbook()
        archive = zipfile.ZipFile(io.BytesIO(), "w")
        mf._write(archive, wb)
        assert "/xl/workbook.xml" in mf.filenames

    @pytest.mark.parametrize(
        "file, registration",
        [
            (
                "xl/media/image1.png",
                '<Default ContentType="image/png" Extension="png"/>',
            ),
            (
                "xl/drawings/commentsDrawing.vml",
                '<Default ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"'
                ' Extension="vml"/>',
            ),
        ],
    )
    def test_media(self, manifest, file, registration):
        from openpyxl.workbook.workbook import Workbook

        wb = Workbook()
        m = manifest()
        m._register_mimetypes([file])
        xml = tostring(m.Default[-1].to_tree())
        diff = compare_xml(xml, registration)
        assert diff is None, diff

    def test_vba(self, datadir, manifest):
        from openpyxl.reader.excel import load_workbook

        datadir.chdir()
        wb = load_workbook("sample.xlsm", keep_vba=True)
        m = manifest()
        m._write_vba(wb)
        partnames = set([t.PartName for t in m.Override])
        expected = {
            "/xl/theme/theme1.xml",
            "/xl/styles.xml",
            "/docProps/core.xml",
            "/docProps/app.xml",
        }
        assert partnames == expected

    def test_no_defaults(self, manifest):
        """
        LibreOffice does not use the Default element
        """
        xml = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Override
                    PartName="/_rels/.rels"
                    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        </Types>
        """
        node = fromstring(xml)
        m = manifest.from_tree(node)
        exts = m.extensions
        assert exts == []

    def test_find(self, datadir, manifest):
        datadir.chdir()
        with open("manifest.xml", "rb") as src:
            xml = src.read()
        tree = fromstring(xml)
        m = manifest.from_tree(tree)
        ws = m.find(WORKSHEET_TYPE)
        assert ws.PartName == "/xl/worksheets/sheet1.xml"

    def test_find_none(self, manifest):
        m = manifest()
        assert m.find(WORKSHEET_TYPE) is None

    def test_findall(self, datadir, manifest):
        datadir.chdir()
        with open("manifest.xml", "rb") as src:
            xml = src.read()
        tree = fromstring(xml)
        m = manifest.from_tree(tree)
        sheets = m.findall(WORKSHEET_TYPE)
        assert len(list(sheets)) == 1
