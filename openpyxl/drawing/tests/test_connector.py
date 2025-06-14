# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.shapes import Transform2D
from openpyxl.chart.text import RichText
from openpyxl.drawing.colors import SchemeColor
from openpyxl.drawing.geometry import FontReference
from openpyxl.drawing.geometry import GeomGuideList
from openpyxl.drawing.geometry import Point2D
from openpyxl.drawing.geometry import PositiveSize2D
from openpyxl.drawing.geometry import PresetGeometry2D
from openpyxl.drawing.geometry import ShapeStyle
from openpyxl.drawing.geometry import StyleMatrixReference
from openpyxl.drawing.properties import NonVisualDrawingProps
from openpyxl.drawing.properties import NonVisualDrawingShapeProps
from openpyxl.drawing.text import CharacterProperties
from openpyxl.drawing.text import ListStyle
from openpyxl.drawing.text import Paragraph
from openpyxl.drawing.text import ParagraphProperties
from openpyxl.drawing.text import RichTextProperties
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def connector_shape():
    from openpyxl.drawing.connector import ConnectorShape

    return ConnectorShape


@pytest.fixture
def shape_meta():
    from openpyxl.drawing.connector import ShapeMeta

    return ShapeMeta


@pytest.fixture
def shape():
    from openpyxl.drawing.connector import Shape

    return Shape


class TestConnectorShape:
    @pytest.mark.xfail
    def test_ctor(self, connector_shape):
        fut = connector_shape()
        xml = tostring(fut.to_tree())
        expected = "<root/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, connector_shape):
        src = """
        <cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               macro="">
            <nvCxnSpPr>
                <cNvPr id="3" name="Straight Arrow Connector 2">
                </cNvPr>
                <cNvCxnSpPr/>
            </nvCxnSpPr>
            <spPr>
                <a:xfrm flipH="1" flipV="1">
                    <a:off x="3321050" y="3829050"/>
                    <a:ext cx="165100" cy="368300"/>
                </a:xfrm>
                <a:prstGeom prst="straightConnector1">
                    <a:avLst/>
                </a:prstGeom>
                <a:ln>
                    <a:tailEnd type="triangle"/>
                </a:ln>
            </spPr>
        </cxnSp>
        """
        node = fromstring(src)
        cnx = connector_shape.from_tree(node)
        assert cnx.nvCxnSpPr.cNvPr.id == 3


class TestShapeMeta:
    @pytest.mark.xfail
    def test_ctor(self, shape_meta):
        meta = shape_meta(
            cNvPr=NonVisualDrawingProps(),
            cNvSpPr=NonVisualDrawingShapeProps(),
        )
        xml = tostring(meta.to_tree())
        expected = "<root/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.xfail
    def test_from_xml(self, shape_meta):
        src = "<root/>"
        node = fromstring(src)
        meta = shape_meta.from_tree(node)
        assert meta == shape_meta()


class TestShape:
    def test_ctor(self, shape):
        props = GraphicalProperties(
            xfrm=Transform2D(
                off=Point2D(x=1767840, y=1341120),
                ext=PositiveSize2D(cx=1539240, cy=281940),
            ),
            prstGeom=PresetGeometry2D(prst="roundRect", avLst=GeomGuideList()),
        )
        props.ln = None
        ln = StyleMatrixReference(
            idx=2,
            schemeClr=SchemeColor(val="accent1", shade=50000),
        )
        fill = StyleMatrixReference(idx=1, schemeClr=SchemeColor(val="accent1"))
        effect = StyleMatrixReference(idx=0, schemeClr=SchemeColor(val="accent1"))
        font = FontReference(idx="minor", schemeClr=SchemeColor(val="lt1"))
        style = ShapeStyle(lnRef=ln, fillRef=fill, effectRef=effect, fontRef=font)
        body = RichTextProperties(
            vertOverflow="clip",
            horzOverflow="clip",
            rtlCol=False,
            anchor="t",
        )
        p = Paragraph(
            endParaRPr=CharacterProperties(lang="en-US", sz="1100"),
            pPr=ParagraphProperties(algn="l"),
        )
        p.r = []
        text = RichText(bodyPr=body, lstStyle=ListStyle(), p=[p])
        s = shape(
            spPr=props,
            style=style,
            txBody=text,
            macro="[0]!RoundedRectangle1_Click",
        )
        xml = tostring(s.to_tree())
        expected = """
        <sp macro="[0]!RoundedRectangle1_Click"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <spPr>
                <a:xfrm>
                    <a:off x="1767840" y="1341120"/>
                    <a:ext cx="1539240" cy="281940"/>
                </a:xfrm>
                <a:prstGeom prst="roundRect">
                    <a:avLst/>
                </a:prstGeom>
            </spPr>
            <style>
                <a:lnRef idx="2">
                    <a:schemeClr val="accent1">
                        <a:shade val="50000"/>
                    </a:schemeClr>
                </a:lnRef>
                <a:fillRef idx="1">
                    <a:schemeClr val="accent1"/>
                </a:fillRef>
                <a:effectRef idx="0">
                    <a:schemeClr val="accent1"/>
                </a:effectRef>
                <a:fontRef idx="minor">
                    <a:schemeClr val="lt1"/>
                </a:fontRef>
            </style>
            <txBody>
                <a:bodyPr 
                        vertOverflow="clip"
                        horzOverflow="clip"
                        rtlCol="0"
                        anchor="t"/>
                <a:lstStyle/>
                <a:p>
                    <a:pPr algn="l"/>
                    <a:endParaRPr lang="en-US" sz="1100"/>
                </a:p>
            </txBody>
        </sp>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, shape, shape_meta):
        src = """
        <sp macro="[0]!RoundedRectangle1_Click"
            textlink=""
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <nvSpPr>
                <cNvPr id="2" name="Rounded Rectangle 1"/>
                <cNvSpPr/>
            </nvSpPr>
            <spPr>
                <a:xfrm>
                    <a:off x="1767840" y="1341120"/>
                    <a:ext cx="1539240" cy="281940"/>
                </a:xfrm>
                <a:prstGeom prst="roundRect">
                    <a:avLst/>
                </a:prstGeom>
            </spPr>
            <style>
                <a:lnRef idx="2">
                    <a:schemeClr val="accent1">
                        <a:shade val="50000"/>
                    </a:schemeClr>
                </a:lnRef>
                <a:fillRef idx="1">
                    <a:schemeClr val="accent1"/>
                </a:fillRef>
                <a:effectRef idx="0">
                    <a:schemeClr val="accent1"/>
                </a:effectRef>
                <a:fontRef idx="minor">
                    <a:schemeClr val="lt1"/>
                </a:fontRef>
            </style>
            <txBody>
                <a:bodyPr
                        vertOverflow="clip"
                        horzOverflow="clip"
                        rtlCol="0"
                        anchor="t"/>
                <a:lstStyle/>
                <a:p>
                    <a:pPr algn="l"/>
                    <a:endParaRPr lang="en-US" sz="1100"/>
                </a:p>
            </txBody>
        </sp>
        """
        node = fromstring(src)
        s = shape.from_tree(node)
        meta = shape_meta(
            cNvPr=NonVisualDrawingProps(id=2, name="Rounded Rectangle 1"),
            cNvSpPr=NonVisualDrawingShapeProps(),
        )
        props = GraphicalProperties(
            xfrm=Transform2D(
                off=Point2D(x=1767840, y=1341120),
                ext=PositiveSize2D(cx=1539240, cy=281940),
            ),
            prstGeom=PresetGeometry2D(prst="roundRect", avLst=GeomGuideList()),
        )
        ln = StyleMatrixReference(
            idx=2,
            schemeClr=SchemeColor(val="accent1", shade=50000),
        )
        fill = StyleMatrixReference(idx=1, schemeClr=SchemeColor(val="accent1"))
        effect = StyleMatrixReference(idx=0, schemeClr=SchemeColor(val="accent1"))
        font = FontReference(idx="minor", schemeClr=SchemeColor(val="lt1"))
        style = ShapeStyle(lnRef=ln, fillRef=fill, effectRef=effect, fontRef=font)
        body = RichTextProperties(
            vertOverflow="clip",
            horzOverflow="clip",
            rtlCol=False,
            anchor="t",
        )
        p = Paragraph(
            endParaRPr=CharacterProperties(lang="en-US", sz="1100"),
            pPr=ParagraphProperties(algn="l"),
        )
        text = RichText(bodyPr=body, lstStyle=ListStyle(), p=[p])
        s2 = shape(
            nvSpPr=meta,
            spPr=props,
            style=style,
            txBody=text,
            macro="[0]!RoundedRectangle1_Click",
        )
        assert s.meta == s2.meta
        assert s.spPr == s2.spPr
        assert s.style == s2.style
        assert s.txBody == s2.txBody
        assert s.macro == s2.macro
