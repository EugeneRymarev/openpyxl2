# Copyright (c) 2010-2025 openpyxl
import datetime
import io
import os
import string
import zipfile

import pytest

from openpyxl.chart.bar_chart import BarChart
from openpyxl.comments.comments import Comment
from openpyxl.connection.connections import Connection
from openpyxl.connection.connections import ConnectionList
from openpyxl.drawing.image import Image
from openpyxl.drawing.legacy import LegacyDrawing
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.packaging.relationship import Relationship
from openpyxl.packaging.relationship import RelationshipList
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.cache import CacheFieldList
from openpyxl.pivot.cache import CacheSource
from openpyxl.pivot.table import Location
from openpyxl.pivot.table import TableDefinition
from openpyxl.reader.excel import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.table import Table


@pytest.fixture
def excel_writer():
    from openpyxl.writer.excel import ExcelWriter

    return ExcelWriter


@pytest.fixture
def archive():
    out = io.BytesIO()
    return zipfile.ZipFile(out, "w")


@pytest.fixture
def emf(datadir):
    datadir.chdir()
    img = Image("checkbox.emf")
    return img


class TestExcelWriter:
    def test_worksheet(self, excel_writer, archive):
        wb = Workbook()
        ws = wb.active
        writer = excel_writer(wb, archive)
        writer.write_worksheets()
        assert ws.path[1:] in archive.namelist()
        assert ws.path in writer.manifest.filenames

    def test_worksheet_with_pivot_cache(self, excel_writer, archive):
        wb = Workbook()
        ws = wb.active
        ws._pivots = [
            TableDefinition(
                name="TestTable",
                cacheId=1,
                dataCaption="TestCap",
                location=Location(
                    ref="Test",
                    firstHeaderRow=1,
                    firstDataRow=1,
                    firstDataCol=1,
                ),
            )
        ]
        ws._pivots[0].cache = CacheDefinition(
            cacheSource=CacheSource(type="worksheet"),
            cacheFields=CacheFieldList(),
        )
        writer = excel_writer(wb, archive)
        writer.write_worksheets()
        expected = [
            "xl/worksheets/sheet1.xml",
            "xl/pivotCache/pivotCacheDefinition1.xml",
            "xl/pivotTables/_rels/pivotTable1.xml.rels",
            "xl/pivotTables/pivotTable1.xml",
            "xl/worksheets/_rels/sheet1.xml.rels",
        ]
        assert writer.archive.namelist() == expected

    def test_tables(self, excel_writer, archive):
        wb = Workbook()
        ws = wb.active
        ws.append(list(string.ascii_letters))
        ws._rels = []
        t = Table(displayName="Table1", ref="A1:D10")
        ws.add_table(t)
        writer = excel_writer(wb, archive)
        writer.write_worksheets()
        assert t.path[1:] in archive.namelist()
        assert t.path in writer.manifest.filenames

    def test_drawing(self, excel_writer, archive):
        wb = Workbook()
        drawing = SpreadsheetDrawing()
        writer = excel_writer(wb, archive)
        writer.write_drawing(drawing)
        assert drawing.path == "/xl/drawings/drawing1.xml"
        assert drawing.path[1:] in archive.namelist()
        assert drawing.path in writer.manifest.filenames

    def test_legacy(self, excel_writer, archive, emf):
        wb = Workbook()
        ws = wb.active
        drawing = LegacyDrawing("some vml")
        rels = RelationshipList()
        rel = Relationship(Type=emf.rel_type, Target="")
        rel.blob = emf
        rels.append(rel)
        drawing.children = rels
        ws.legacy_drawing = drawing
        writer = excel_writer(wb, archive)
        writer.write_legacy(ws)
        assert len(writer._images) == 1
        assert rel.Target == "/xl/media/image1.wmf"
        expected = [
            "xl/drawings/vmlDrawing1.vml",
            "xl/media/image1.wmf",
            "xl/drawings/_rels/vmlDrawing1.vml.rels",
        ]
        assert archive.namelist() == expected

    def test_write_chart(self, excel_writer, archive):
        wb = Workbook()
        ws = wb.active
        chart = BarChart()
        ws.add_chart(chart)
        writer = excel_writer(wb, archive)
        writer.write_worksheets()
        assert "xl/worksheets/sheet1.xml" in archive.namelist()
        assert ws.path in writer.manifest.filenames
        rel = ws._rels.get("rId1")
        expected = {
            "Id": "rId1",
            "Target": "/xl/drawings/drawing1.xml",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        }
        assert dict(rel) == expected

    def test_chartsheet(self, excel_writer, archive):
        wb = Workbook()
        cs = wb.create_chartsheet()
        writer = excel_writer(wb, archive)
        writer.write_chartsheets()
        assert cs.path in writer.manifest.filenames
        assert cs.path[1:] in writer.archive.namelist()

    def test_comment(self, excel_writer, archive):
        wb = Workbook()
        ws = wb.active
        ws["B5"].comment = Comment("A comment", "The Author")
        writer = excel_writer(None, archive)
        writer.write_comment(ws)
        assert archive.namelist() == ["xl/comments/comment1.xml"]
        assert "/xl/comments/comment1.xml" in writer.manifest.filenames
        # assert ws.legacy_drawing.vml[:15] == b'<xml><ns0:shape'
        # assert len(ws.legacy_drawing.vml) == 489

    def test_duplicate_comment(self, excel_writer, archive):
        wb = Workbook()
        ws = wb.active
        ws["B5"].comment = Comment("A comment", "The Author")
        writer = excel_writer(wb, archive)
        writer.write_comment(ws)
        writer.write_comment(ws)
        # assert len(ws.legacy_drawing.vml) == 489

    def test_merge_vba(self, excel_writer, archive, datadir):
        datadir.chdir()
        wb = load_workbook("vba+comments.xlsm")
        writer = excel_writer(wb, archive)
        writer._merge_vba()
        assert set(archive.namelist()) == {"xl/vbaProject.bin"}

    def test_duplicate_chart(self, excel_writer, archive):
        from openpyxl.chart import PieChart

        pc = PieChart()
        wb = Workbook()
        writer = excel_writer(wb, archive)
        writer._charts = [pc] * 2
        with pytest.raises(InvalidFileException):
            writer.write_charts()

    def test_controls(self, excel_writer, archive, emf):
        from openpyxl.packaging.relationship import Relationship
        from openpyxl.worksheet.controls import ActiveXControl
        from openpyxl.worksheet.controls import ControlList
        from openpyxl.xml.functions import fromstring

        src = """
        <controls
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <control shapeId="47129" r:id="rId8" name="MainSVCheckBox">
                <controlPr defaultSize="0" autoLine="0" r:id="rId9">
                    <anchor moveWithCells="1">
                        <from>
                            <xdr:col>12</xdr:col>
                            <xdr:colOff>219075</xdr:colOff>
                            <xdr:row>7</xdr:row>
                            <xdr:rowOff>95250</xdr:rowOff>
                        </from>
                        <to>
                            <xdr:col>12</xdr:col>
                            <xdr:colOff>400050</xdr:colOff>
                            <xdr:row>7</xdr:row>
                            <xdr:rowOff>276225</xdr:rowOff>
                        </to>
                    </anchor>
                </controlPr>
            </control>
        </controls>
        """
        tree = fromstring(src)
        controls = ControlList.from_tree(tree)
        ctrl = controls.control[0]
        ctrl.shape = ActiveXControl(persistence="persistStreamInit")
        ctrl.shape.bin = b"\001"
        prop = ctrl.controlPr
        prop.image = Relationship(type="image", Target="")
        prop.image.blob = emf
        wb = Workbook()
        ws = wb.active
        ws.controls = controls
        writer = excel_writer(wb, archive)
        writer.write_worksheet(ws)
        assert prop.image.target == "/xl/media/image1.wmf"
        assert len(writer._images) == 1
        expected = [
            "xl/activeX/activeX1.bin",
            "xl/activeX/_rels/activeX1.xml.rels",
            "xl/activeX/activeX1.xml",
            "xl/media/image1.wmf",
            "xl/worksheets/sheetNone.xml",
        ]
        assert archive.namelist() == expected

    def test_controls_with_shared_images(self, excel_writer, archive, emf):
        from openpyxl.packaging.relationship import Relationship
        from openpyxl.worksheet.controls import ActiveXControl
        from openpyxl.worksheet.controls import ControlList
        from openpyxl.xml.functions import fromstring

        src = """
        <controls
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
            <control shapeId="47129" r:id="rId8" name="MainSVCheckBox">
                <controlPr defaultSize="0" autoLine="0" r:id="rId9">
                    <anchor moveWithCells="1">
                        <from>
                            <xdr:col>12</xdr:col>
                            <xdr:colOff>219075</xdr:colOff>
                            <xdr:row>7</xdr:row>
                            <xdr:rowOff>95250</xdr:rowOff>
                        </from>
                        <to>
                            <xdr:col>12</xdr:col>
                            <xdr:colOff>400050</xdr:colOff>
                            <xdr:row>7</xdr:row>
                            <xdr:rowOff>276225</xdr:rowOff>
                        </to>
                    </anchor>
                </controlPr>
            </control>
            <control shapeId="47130" r:id="rId10" name="MainSVCheckBox">
                <controlPr defaultSize="0" autoLine="0" r:id="rId9">
                    <anchor moveWithCells="1">
                        <from>
                            <xdr:col>12</xdr:col>
                            <xdr:colOff>219075</xdr:colOff>
                            <xdr:row>7</xdr:row>
                            <xdr:rowOff>95250</xdr:rowOff>
                        </from>
                        <to>
                            <xdr:col>12</xdr:col>
                            <xdr:colOff>400050</xdr:colOff>
                            <xdr:row>7</xdr:row>
                            <xdr:rowOff>276225</xdr:rowOff>
                        </to>
                    </anchor>
                </controlPr>
            </control>
        </controls>
        """
        tree = fromstring(src)
        controls = ControlList.from_tree(tree)
        for ctrl in controls.control:
            ctrl.shape = ActiveXControl(persistence="persistStreamInit")
            ctrl.shape.bin = b"\001"
            prop = ctrl.controlPr
            prop.image = Relationship(type="image", Target="")
            prop.image.blob = emf
            prop.image.Target = "/xl/media/image1.emf"
        wb = Workbook()
        ws = wb.active
        ws.controls = controls
        writer = excel_writer(wb, archive)
        writer.write_worksheet(ws)
        expected = [
            "xl/activeX/activeX1.bin",
            "xl/activeX/_rels/activeX1.xml.rels",
            "xl/activeX/activeX1.xml",
            "xl/activeX/activeX2.bin",
            "xl/activeX/_rels/activeX2.xml.rels",
            "xl/activeX/activeX2.xml",
            "xl/media/image1.wmf",
            "xl/worksheets/sheetNone.xml",
        ]
        assert archive.namelist() == expected

    def test_add_image(self, excel_writer, emf):
        archive = zipfile.ZipFile(io.BytesIO(), "w")
        writer = excel_writer(None, archive)
        writer.add_image(emf)
        assert writer._images == [emf]
        assert writer.archive.namelist() == ["xl/media/image1.wmf"]

    def test_duplicate_image(self, excel_writer, emf):
        archive = zipfile.ZipFile(io.BytesIO(), "w")
        writer = excel_writer(None, archive)
        writer.add_image(emf)
        writer.add_image(emf)
        assert writer._images == [emf]
        assert writer.archive.namelist() == ["xl/media/image1.wmf"]

    def test_volatile_deps(self, excel_writer, archive):
        from openpyxl.volatile.volatile import VolMain
        from openpyxl.volatile.volatile import VolTopic
        from openpyxl.volatile.volatile import VolType
        from openpyxl.volatile.volatile import VolTypesList

        archive = zipfile.ZipFile(io.BytesIO(), "w")
        wb = Workbook()
        volmain = VolMain(first="teststring", tp=[VolTopic(t="s", v="aaa: 4447")])
        voltype = VolType(main=[volmain], type="realTimeData")
        wb._volatile_deps = VolTypesList(volType=[voltype])
        writer = excel_writer(wb, archive)
        writer.write_volatile_deps()
        assert writer.archive.namelist() == ["xl/volatileDependencies.xml"]

    def test_connections(self, excel_writer, archive):
        archive = zipfile.ZipFile(io.BytesIO(), "w")
        wb = Workbook()
        con = [Connection(id=1, refreshedVersion=8)]
        wb._connections = ConnectionList(connection=con)
        writer = excel_writer(wb, archive)
        writer.write_connections()
        assert writer.archive.namelist() == ["xl/connections.xml"]

    def test_connection_with_cache(self, excel_writer, archive):
        archive = zipfile.ZipFile(io.BytesIO(), "w")
        wb = Workbook()
        con = [Connection(id=1, refreshedVersion=8)]
        wb._connections = ConnectionList(connection=con)
        wb._connections[1]._cache = CacheDefinition(
            cacheSource=CacheSource(type="external"),
            cacheFields=CacheFieldList(),
        )
        writer = excel_writer(wb, archive)
        writer.write_connections()
        expected = ["xl/pivotCache/pivotCacheDefinition1.xml", "xl/connections.xml"]
        assert writer.archive.namelist() == expected


def test_write_empty_workbook(tmpdir):
    from openpyxl.writer.excel import save_workbook

    tmpdir.chdir()
    wb = Workbook()
    dest_filename = "empty_book.xlsx"
    save_workbook(wb, dest_filename)
    assert os.path.isfile(dest_filename)


def test_modified(tmpdir):
    from openpyxl.writer.excel import save_workbook

    tmpdir.chdir()
    wb = Workbook()
    modified = datetime.datetime(2011, 5, 19, 10, 23, 15)
    wb.properties.modified = modified
    dest_filename = "empty_book.xlsx"
    save_workbook(wb, dest_filename)
    assert wb.properties.modified > modified
