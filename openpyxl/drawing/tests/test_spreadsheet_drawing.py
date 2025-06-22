# Copyright (c) 2010-2025 openpyxl
import PIL
import pytest
from openpyxl.chart.bar_chart import BarChart
from openpyxl.drawing.image import Image
from openpyxl.packaging.relationship import Relationship
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def two_cell_anchor():
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor

    return TwoCellAnchor


@pytest.fixture
def one_cell_anchor():
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor

    return OneCellAnchor


@pytest.fixture
def absolute_anchor():
    from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor

    return AbsoluteAnchor


@pytest.fixture
def spreadsheet_drawing():
    from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing

    return SpreadsheetDrawing


class TestTwoCellAnchor:
    def test_ctor(self, two_cell_anchor):
        chart_drawing = two_cell_anchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <twoCellAnchor>
            <from>
                <col>0</col>
                <colOff>0</colOff>
                <row>0</row>
                <rowOff>0</rowOff>
            </from>
            <to>
                <col>0</col>
                <colOff>0</colOff>
                <row>0</row>
                <rowOff>0</rowOff>
            </to>
            <clientData/>
        </twoCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, two_cell_anchor):
        src = """
        <twoCellAnchor>
            <from>
                <col>0</col>
                <colOff>0</colOff>
                <row>0</row>
                <rowOff>0</rowOff>
            </from>
            <to>
                <col>0</col>
                <colOff>0</colOff>
                <row>0</row>
                <rowOff>0</rowOff>
            </to>
            <clientData></clientData>
        </twoCellAnchor>
        """
        node = fromstring(src)
        chart_drawing = two_cell_anchor.from_tree(node)
        assert chart_drawing == two_cell_anchor()


class TestOneCellAnchor:
    def test_ctor(self, one_cell_anchor):
        chart_drawing = one_cell_anchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <oneCellAnchor>
            <from>
                <col>0</col>
                <colOff>0</colOff>
                <row>0</row>
                <rowOff>0</rowOff>
            </from>
            <ext cx="0" cy="0"/>
            <clientData></clientData>
        </oneCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, one_cell_anchor):
        src = """
        <oneCellAnchor>
            <from>
                <col>0</col>
                <colOff>0</colOff>
                <row>0</row>
                <rowOff>0</rowOff>
            </from>
            <ext cx="0" cy="0"/>
            <clientData></clientData>
        </oneCellAnchor>
        """
        node = fromstring(src)
        chart_drawing = one_cell_anchor.from_tree(node)
        assert chart_drawing == one_cell_anchor()


class TestAbsoluteAnchor:
    def test_ctor(self, absolute_anchor):
        chart_drawing = absolute_anchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <absoluteAnchor>
            <pos x="0" y="0"/>
            <ext cx="0" cy="0"/>
            <clientData></clientData>
        </absoluteAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, absolute_anchor):
        src = """
        <absoluteAnchor>
            <pos x="0" y="0"/>
            <ext cx="0" cy="0"/>
            <clientData></clientData>
        </absoluteAnchor>
        """
        node = fromstring(src)
        chart_drawing = absolute_anchor.from_tree(node)
        assert chart_drawing == absolute_anchor()


class TestSpreadsheetDrawing:
    def test_ctor(self, spreadsheet_drawing):
        from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
        from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor
        from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor

        a = [AbsoluteAnchor(), AbsoluteAnchor()]
        o = [OneCellAnchor()]
        t = [TwoCellAnchor(), TwoCellAnchor()]
        chart_drawing = spreadsheet_drawing(
            absoluteAnchor=a,
            oneCellAnchor=o,
            twoCellAnchor=t,
        )
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <wsDr>
            <twoCellAnchor>
                <from>
                    <col>0</col>
                    <colOff>0</colOff>
                    <row>0</row>
                    <rowOff>0</rowOff>
                </from>
                <to>
                    <col>0</col>
                    <colOff>0</colOff>
                    <row>0</row>
                    <rowOff>0</rowOff>
                </to>
                <clientData></clientData>
            </twoCellAnchor>
            <twoCellAnchor>
                <from>
                    <col>0</col>
                    <colOff>0</colOff>
                    <row>0</row>
                    <rowOff>0</rowOff>
                </from>
                <to>
                    <col>0</col>
                    <colOff>0</colOff>
                    <row>0</row>
                    <rowOff>0</rowOff>
                </to>
                <clientData></clientData>
            </twoCellAnchor>
            <oneCellAnchor>
                <from>
                    <col>0</col>
                    <colOff>0</colOff>
                    <row>0</row>
                    <rowOff>0</rowOff>
                </from>
                <ext cx="0" cy="0"/>
                <clientData></clientData>
            </oneCellAnchor>
            <absoluteAnchor>
                <pos x="0" y="0"/>
                <ext cx="0" cy="0"/>
                <clientData></clientData>
            </absoluteAnchor>
            <absoluteAnchor>
                <pos x="0" y="0"/>
                <ext cx="0" cy="0"/>
                <clientData></clientData>
            </absoluteAnchor>
        </wsDr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_chart(self, spreadsheet_drawing):
        from openpyxl.chart._chart import ChartBase

        class Chart(ChartBase):
            anchor = "E15"
            width = 15
            height = 7.5

        drawing = spreadsheet_drawing()
        drawing.charts.append(Chart())
        xml = tostring(drawing._write())
        expected = """
        <wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <oneCellAnchor>
                <from>
                    <col>4</col>
                    <colOff>0</colOff>
                    <row>14</row>
                    <rowOff>0</rowOff>
                </from>
                <ext cx="5400000" cy="2700000"/>
                <graphicFrame>
                    <nvGraphicFramePr>
                        <cNvPr id="1" name="Chart 1"/>
                        <cNvGraphicFramePr/>
                    </nvGraphicFramePr>
                    <xfrm/>
                    <a:graphic>
                        <a:graphicData
                                uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                            <c:chart
                                    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                    r:id="rId1"/>
                        </a:graphicData>
                    </a:graphic>
                </graphicFrame>
                <clientData/>
            </oneCellAnchor>
        </wsDr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_hash_function(self, spreadsheet_drawing):
        drawing = spreadsheet_drawing()
        assert hash(drawing) == hash(id(drawing))

    def test_write_picture(self, spreadsheet_drawing):
        drawing = spreadsheet_drawing()
        pic = drawing._picture_frame(4)
        xml = tostring(pic.to_tree())
        expected = """
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <nvPicPr>
                <cNvPr id="4" name="Image 4"></cNvPr>
                <cNvPicPr/>
            </nvPicPr>
            <blipFill>
                <a:blip cstate="print" r:embed="rId4"/>
                <a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:fillRect/>
                </a:stretch>
            </blipFill>
            <spPr>
                <a:prstGeom prst="rect"/>
            </spPr>
        </pic>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_picture_with_alttext(self, spreadsheet_drawing):
        drawing = spreadsheet_drawing()
        pic = drawing._picture_frame(4, "I have set a desc")
        xml = tostring(pic.to_tree())
        expected = """
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <nvPicPr>
                <cNvPr descr="I have set a desc" id="4" name="Image 4"></cNvPr>
                <cNvPicPr/>
            </nvPicPr>
            <blipFill>
                <a:blip cstate="print" r:embed="rId4"/>
                <a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:fillRect/>
                </a:stretch>
            </blipFill>
            <spPr>
                <a:prstGeom prst="rect"/>
            </spPr>
        </pic>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_read_chart(self, spreadsheet_drawing, datadir):
        datadir.chdir()
        with open("spreadsheet_drawing_with_chart.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        drawing = spreadsheet_drawing.from_tree(node)
        chart_rels = drawing._chart_rels
        assert len(chart_rels) == 1
        assert chart_rels[0].anchor is not None

    @pytest.mark.parametrize(
        "path",
        ["spreadsheet_drawing_with_blip.xml", "two_cell_anchor_pic.xml"],
    )
    def test_read_blip(self, spreadsheet_drawing, datadir, path):
        datadir.chdir()
        with open(path, "rb") as src:
            xml = src.read()
        node = fromstring(xml)
        drawing = spreadsheet_drawing.from_tree(node)
        blip_rels = drawing._blip_rels
        assert len(blip_rels) == 1
        assert blip_rels[0].anchor is not None

    def test_ignore_external_blip(self, spreadsheet_drawing, datadir):
        with open("spreadsheet_drawing_external_image.xml") as src:
            xml = src.read()
        node = fromstring(xml)
        drawing = spreadsheet_drawing.from_tree(node)
        assert drawing._blip_rels == []

    def test_group_rels(self, spreadsheet_drawing, datadir):
        with open("multipic_group.xml", "rb") as src:
            xml = src.read()
        node = fromstring(xml)
        drawing = spreadsheet_drawing.from_tree(node)
        assert len(drawing._group_rels[0]) == 3
        anchor = drawing._group_rels[0][0]
        assert anchor.grpSp.sp is not None

    def test_write_rels(self, spreadsheet_drawing):
        from openpyxl.packaging.relationship import Relationship

        rel = Relationship(type="drawing", Target="../file.xml")
        drawing = spreadsheet_drawing()
        drawing._rels.append(rel)
        xml = tostring(drawing._write_rels())
        expected = """
        <Relationships
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship
                    Id="rId1"
                    Target="../file.xml"
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
        </Relationships>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_path(self, spreadsheet_drawing):
        drawing = spreadsheet_drawing()
        assert drawing.path == "/xl/drawings/drawingNone.xml"

    def test_empty(self, spreadsheet_drawing):
        drawing = spreadsheet_drawing()
        assert bool(drawing) is False

    @pytest.mark.parametrize("attr", ["charts", "images"])
    def test_bool(self, spreadsheet_drawing, attr):
        drawing = spreadsheet_drawing()
        getattr(drawing, attr).append(1)
        assert bool(drawing) is True

    def test_image_as_pic(self, spreadsheet_drawing):
        src = """
        <wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <twoCellAnchor>
                <from>
                    <col>0</col>
                    <colOff>0</colOff>
                    <row>0</row>
                    <rowOff>0</rowOff>
                </from>
                <to>
                    <col>8</col>
                    <colOff>158506</colOff>
                    <row>10</row>
                    <rowOff>64012</rowOff>
                </to>
                <pic>
                    <nvPicPr>
                        <cNvPr id="2" name="Picture 1"/>
                        <cNvPicPr/>
                    </nvPicPr>
                    <blipFill>
                        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                r:embed="rId1">
                        </a:blip>
                        <a:stretch>
                            <a:fillRect/>
                        </a:stretch>
                    </blipFill>
                    <spPr>
                        <a:ln>
                            <a:prstDash val="solid"/>
                        </a:ln>
                    </spPr>
                </pic>
                <clientData/>
            </twoCellAnchor>
        </wsDr>
        """
        node = fromstring(src)
        drawing = spreadsheet_drawing.from_tree(node)
        anchor = drawing.twoCellAnchor[0]
        drawing.twoCellAnchor = []
        img = Image(PIL.Image.new(mode="RGB", size=(1, 1)))
        img.format = "PNG"
        img.anchor = anchor
        drawing.images.append(img)
        xml = tostring(drawing._write())
        diff = compare_xml(xml, src)
        assert diff is None, diff

    @pytest.mark.xfail  # Group handling has changed
    def test_image_as_group(self, spreadsheet_drawing):
        src = """
        <wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <twoCellAnchor>
                <from>
                    <col>5</col>
                    <colOff>114300</colOff>
                    <row>0</row>
                    <rowOff>0</rowOff>
                </from>
                <to>
                    <col>8</col>
                    <colOff>317500</colOff>
                    <row>4</row>
                    <rowOff>165100</rowOff>
                </to>
                <grpSp>
                    <nvGrpSpPr>
                        <cNvPr id="2208" name="Group 1"/>
                        <cNvGrpSpPr>
                            <a:grpSpLocks/>
                        </cNvGrpSpPr>
                    </nvGrpSpPr>
                    <grpSpPr bwMode="auto">
                    </grpSpPr>
                    <pic>
                        <nvPicPr>
                            <cNvPr id="2209" name="Picture 2"/>
                            <cNvPicPr>
                                <a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
                            </cNvPicPr>
                        </nvPicPr>
                        <blipFill>
                            <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                    r:embed="rId1"
                                    cstate="print">
                            </a:blip>
                            <a:srcRect/>
                            <a:stretch>
                                <a:fillRect/>
                            </a:stretch>
                        </blipFill>
                        <spPr bwMode="auto">
                            <a:xfrm>
                                <a:off x="303" y="0"/>
                                <a:ext cx="321" cy="88"/>
                            </a:xfrm>
                            <a:prstGeom prst="rect"/>
                            <a:noFill/>
                            <a:ln>
                                <a:prstDash val="solid"/>
                            </a:ln>
                        </spPr>
                    </pic>
                </grpSp>
                <clientData/>
            </twoCellAnchor>
        </wsDr>
        """
        node = fromstring(src)
        drawing = spreadsheet_drawing.from_tree(node)
        anchor = drawing.twoCellAnchor[0]
        drawing.twoCellAnchor = []
        img = Image(PIL.Image.new(mode="RGB", size=(1, 1)))
        img.anchor = anchor
        img.format = "PNG"
        drawing.images.append(img)
        xml = tostring(drawing._write())
        diff = compare_xml(xml, src)
        assert diff is None, diff

    def test_shapes(self, spreadsheet_drawing, datadir):
        datadir.chdir()
        with open("commands.xml", "rb") as src:
            xml = src.read()
        tree = fromstring(xml)
        drawing = spreadsheet_drawing.from_tree(tree)
        assert len(drawing._shapes) == 4

    def test_hyperlink(self, spreadsheet_drawing, datadir):
        datadir.chdir()
        with open("hyperlink.xml", "rb") as src:
            xml = src.read()
        tree = fromstring(xml)
        drawing = spreadsheet_drawing.from_tree(tree)
        drawing.shapes = drawing._shapes
        drawing._write()
        expected = Relationship(Target="", Id="rId1", type="hyperlink", TargetMode="")
        assert drawing._rels.get("rId1") == expected


def test_check_anchor_chart():
    from openpyxl.drawing.spreadsheet_drawing import _check_anchor

    c = BarChart()
    anc = _check_anchor(c)
    assert anc._from.row == 14
    assert anc._from.col == 4
    assert anc.ext.width == 5400000
    assert anc.ext.height == 2700000


@pytest.mark.parametrize("anchor", ("E17", "e17"))
def test_check_chart_with_anchor(anchor):
    from openpyxl.drawing.spreadsheet_drawing import _check_anchor

    c = BarChart()
    c.anchor = anchor
    anc = _check_anchor(c)
    assert anc._from.row == 16
    assert anc._from.col == 4
    assert anc.ext.width == 5400000
    assert anc.ext.height == 2700000


@pytest.mark.pil_required
def test_check_anchor_image(datadir):
    from openpyxl.drawing.spreadsheet_drawing import _check_anchor
    from PIL.Image import Image as PILImage

    datadir.chdir()
    im = Image(PILImage())
    anc = _check_anchor(im)
    assert anc._from.row == 0
    assert anc._from.col == 0
    assert anc.ext.height == 0
    assert anc.ext.width == 0
