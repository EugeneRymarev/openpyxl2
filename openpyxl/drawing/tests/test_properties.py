# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.drawing.geometry import GroupTransform2D
from openpyxl.drawing.geometry import Point2D
from openpyxl.drawing.geometry import PositiveSize2D
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def non_visual_drawing_props():
    from openpyxl.drawing.properties import NonVisualDrawingProps

    return NonVisualDrawingProps


@pytest.fixture
def non_visual_group_drawing_shape_props():
    from openpyxl.drawing.properties import NonVisualGroupDrawingShapeProps

    return NonVisualGroupDrawingShapeProps


@pytest.fixture
def non_visual_group_shape():
    from openpyxl.drawing.properties import NonVisualGroupShape

    return NonVisualGroupShape


@pytest.fixture
def group_locking():
    from openpyxl.drawing.properties import GroupLocking

    return GroupLocking


@pytest.fixture
def group_shape_properties():
    from openpyxl.drawing.properties import GroupShapeProperties

    return GroupShapeProperties


@pytest.fixture
def non_visual_drawing_shape_props():
    from openpyxl.drawing.properties import NonVisualDrawingShapeProps

    return NonVisualDrawingShapeProps


class TestNonVisualDrawingProps:
    def test_ctor(self, non_visual_drawing_props):
        graphic = non_visual_drawing_props(id=2, name="Chart 1")
        xml = tostring(graphic.to_tree())
        expected = '<cNvPr id="2" name="Chart 1"></cNvPr>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, non_visual_drawing_props):
        src = '<cNvPr id="3" name="Chart 2"></cNvPr>'
        node = fromstring(src)
        graphic = non_visual_drawing_props.from_tree(node)
        assert graphic == non_visual_drawing_props(id=3, name="Chart 2")


class TestNonVisualGroupDrawingShapeProps:
    def test_ctor(self, non_visual_group_drawing_shape_props):
        props = non_visual_group_drawing_shape_props()
        xml = tostring(props.to_tree())
        expected = "<cNvGrpSpPr/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, non_visual_group_drawing_shape_props):
        src = "<cNvGrpSpPr/>"
        node = fromstring(src)
        props = non_visual_group_drawing_shape_props.from_tree(node)
        assert props == non_visual_group_drawing_shape_props()


class TestNonVisualGroupShape:
    def test_ctor(
        self,
        non_visual_group_shape,
        non_visual_drawing_props,
        non_visual_group_drawing_shape_props,
    ):
        props = non_visual_group_shape(
            cNvPr=non_visual_drawing_props(id=2208, name="Group 1"),
            cNvGrpSpPr=non_visual_group_drawing_shape_props(),
        )
        xml = tostring(props.to_tree())
        expected = """
        <nvGrpSpPr>
            <cNvPr id="2208" name="Group 1"/>
            <cNvGrpSpPr/>
        </nvGrpSpPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(
        self,
        non_visual_group_shape,
        non_visual_drawing_props,
        non_visual_group_drawing_shape_props,
    ):
        src = """
        <nvGrpSpPr>
            <cNvPr id="2208" name="Group 1"/>
            <cNvGrpSpPr/>
        </nvGrpSpPr>
        """
        node = fromstring(src)
        props = non_visual_group_shape.from_tree(node)
        expected = non_visual_group_shape(
            cNvPr=non_visual_drawing_props(id=2208, name="Group 1"),
            cNvGrpSpPr=non_visual_group_drawing_shape_props(),
        )
        assert props == expected


class TestGroupLocking:
    def test_ctor(self, group_locking):
        lock = group_locking()
        xml = tostring(lock.to_tree())
        expected = """
        <grpSpLocks
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, group_locking):
        src = "<grpSpLocks/>"
        node = fromstring(src)
        lock = group_locking.from_tree(node)
        assert lock == group_locking()


class TestGroupShapeProperties:
    def test_ctor(self, group_shape_properties):
        xfrm = GroupTransform2D(
            off=Point2D(x=2222500, y=0),
            ext=PositiveSize2D(cx=2806700, cy=825500),
            chOff=Point2D(x=303, y=0),
            chExt=PositiveSize2D(cx=321, cy=111),
        )
        props = group_shape_properties(bwMode="auto", xfrm=xfrm)
        xml = tostring(props.to_tree())
        expected = """
        <grpSpPr
                bwMode="auto"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:xfrm rot="0">
                <a:off x="2222500" y="0"/>
                <a:ext cx="2806700" cy="825500"/>
                <a:chOff x="303" y="0"/>
                <a:chExt cx="321" cy="111"/>
            </a:xfrm>
        </grpSpPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, group_shape_properties):
        src = "<grpSpPr/>"
        node = fromstring(src)
        fut = group_shape_properties.from_tree(node)
        assert fut == group_shape_properties()


class TestNonVisualDrawingShapeProps:
    def test_ctor(self, non_visual_drawing_shape_props):
        props = non_visual_drawing_shape_props(txBox=True)
        xml = tostring(props.to_tree())
        expected = '<cNvSpPr txBox="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, non_visual_drawing_shape_props):
        src = '<cNvSpPr txBox="1"/>'
        node = fromstring(src)
        props = non_visual_drawing_shape_props.from_tree(node)
        assert props == non_visual_drawing_shape_props(txBox=True)
