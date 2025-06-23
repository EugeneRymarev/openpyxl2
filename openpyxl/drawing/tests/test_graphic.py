# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def non_visual_graphic_frame():
    from openpyxl.drawing.graphic import NonVisualGraphicFrame

    return NonVisualGraphicFrame


@pytest.fixture
def graphic_data():
    from openpyxl.drawing.graphic import GraphicData

    return GraphicData


@pytest.fixture
def graphic_object():
    from openpyxl.drawing.graphic import GraphicObject

    return GraphicObject


@pytest.fixture
def graphic_frame():
    from openpyxl.drawing.graphic import GraphicFrame

    return GraphicFrame


@pytest.fixture
def group_transform_2d():
    from openpyxl.drawing.geometry import GroupTransform2D

    return GroupTransform2D


@pytest.fixture
def group_shape():
    from openpyxl.drawing.graphic import GroupShape

    return GroupShape


@pytest.fixture
def non_visual_graphic_frame_properties():
    from openpyxl.drawing.graphic import NonVisualGraphicFrameProperties

    return NonVisualGraphicFrameProperties


class TestNonVisualGraphicFrame:
    def test_ctor(self, non_visual_graphic_frame):
        graphic = non_visual_graphic_frame()
        xml = tostring(graphic.to_tree())
        expected = """
        <nvGraphicFramePr>
            <cNvPr id="0" name="Chart 0"></cNvPr>
            <cNvGraphicFramePr></cNvGraphicFramePr>
        </nvGraphicFramePr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, non_visual_graphic_frame):
        src = """
        <nvGraphicFramePr>
            <cNvPr id="0" name="Chart 0"></cNvPr>
            <cNvGraphicFramePr></cNvGraphicFramePr>
        </nvGraphicFramePr>
        """
        node = fromstring(src)
        graphic = non_visual_graphic_frame.from_tree(node)
        assert graphic == non_visual_graphic_frame()


class TestGraphicData:
    def test_ctor(self, graphic_data):
        graphic = graphic_data()
        xml = tostring(graphic.to_tree())
        expected = """
        <graphicData
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
                uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, graphic_data):
        src = """
        <graphicData
                uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
        """
        node = fromstring(src)
        graphic = graphic_data.from_tree(node)
        assert graphic == graphic_data()

    def test_contains_chart(self, graphic_data):
        src = """
        <graphicData
                uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart
                    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    r:id="rId2"/>
        </graphicData>
        """
        node = fromstring(src)
        graphic = graphic_data.from_tree(node)
        assert graphic.chart is not None


class TestGraphicObject:
    def test_ctor(self, graphic_object):
        graphic = graphic_object()
        xml = tostring(graphic.to_tree())
        expected = """
        <graphic
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
            <graphicData
                    uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            </graphicData>
        </graphic>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, graphic_object):
        src = """
        <graphic>
            <graphicData
                    uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            </graphicData>
        </graphic>
        """
        node = fromstring(src)
        graphic = graphic_object.from_tree(node)
        assert graphic == graphic_object()


class TestGraphicFrame:
    def test_ctor(self, graphic_frame):
        graphic = graphic_frame()
        xml = tostring(graphic.to_tree())
        expected = """
        <graphicFrame
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <nvGraphicFramePr>
                <cNvPr id="0" name="Chart 0"></cNvPr>
                <cNvGraphicFramePr></cNvGraphicFramePr>
            </nvGraphicFramePr>
            <xfrm/>
            <a:graphic>
                <a:graphicData
                        uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
            </a:graphic>
        </graphicFrame>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, graphic_frame):
        src = """
        <graphicFrame>
            <nvGraphicFramePr>
                <cNvPr id="0" name="Chart 0"></cNvPr>
                <cNvGraphicFramePr></cNvGraphicFramePr>
            </nvGraphicFramePr>
            <xfrm></xfrm>
            <graphic>
                <graphicData
                    uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                </graphicData>
            </graphic>
        </graphicFrame>
        """
        node = fromstring(src)
        graphic = graphic_frame.from_tree(node)
        assert graphic == graphic_frame()


class TestGroupTransform2D:
    def test_ctor(self, group_transform_2d):
        xfrm = group_transform_2d(rot=0)
        xml = tostring(xfrm.to_tree())
        expected = """
        <xfrm rot="0"
              xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
        </xfrm>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, group_transform_2d):
        src = """
        <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:off x="0" y="394447"/>
            <a:ext cx="1944896" cy="707294"/>
            <a:chOff x="0" y="351692"/>
            <a:chExt cx="1918002" cy="670746"/>
        </a:xfrm>
        """
        node = fromstring(src)
        xfrm = group_transform_2d.from_tree(node)
        assert xfrm.off.y == 394447


class TestGroupShape:
    @pytest.mark.xfail
    def test_ctor(self, group_shape):
        grp = group_shape()
        xml = tostring(grp.to_tree())
        expected = "<root/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.xfail
    def test_from_xml(self, group_shape):
        src = """
        <xdr:grpSp
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <xdr:nvGrpSpPr>
                <xdr:cNvPr id="14" name="Group 13"/>
                <xdr:cNvGrpSpPr/>
            </xdr:nvGrpSpPr>
            <xdr:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="394447"/>
                    <a:ext cx="1944896" cy="707294"/>
                    <a:chOff x="0" y="351692"/>
                    <a:chExt cx="1918002" cy="670746"/>
                </a:xfrm>
            </xdr:grpSpPr>
            <xdr:sp macro="" textlink="">
                <xdr:nvSpPr>
                    <xdr:cNvPr id="15" name="Rectangle 14"/>
                    <xdr:cNvSpPr/>
                </xdr:nvSpPr>
                <xdr:spPr>
                    <a:xfrm>
                        <a:off x="562916" y="377825"/>
                        <a:ext cx="182880" cy="137982"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst/>
                    </a:prstGeom>
                    <a:solidFill>
                        <a:schemeClr val="accent3">
                            <a:lumMod val="60000"/>
                            <a:lumOff val="40000"/>
                        </a:schemeClr>
                    </a:solidFill>
                    <a:ln w="9525">
                        <a:solidFill>
                            <a:sysClr val="windowText" lastClr="000000"/>
                        </a:solidFill>
                    </a:ln>
                </xdr:spPr>
                <xdr:style>
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
                </xdr:style>
                <xdr:txBody>
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
                </xdr:txBody>
            </xdr:sp>
        </xdr:grpSp>
        """
        node = fromstring(src)
        grp = group_shape.from_tree(node)
        assert grp == group_shape()


class TestNonVisualGraphicFrameProperties:
    def test_ctor(self, non_visual_graphic_frame_properties):
        graphic = non_visual_graphic_frame_properties()
        xml = tostring(graphic.to_tree())
        expected = "<cNvGraphicFramePr></cNvGraphicFramePr>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, non_visual_graphic_frame_properties):
        src = "<cNvGraphicFramePr></cNvGraphicFramePr>"
        node = fromstring(src)
        graphic = non_visual_graphic_frame_properties.from_tree(node)
        assert graphic == non_visual_graphic_frame_properties()
