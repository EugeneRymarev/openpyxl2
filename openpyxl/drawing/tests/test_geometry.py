# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.drawing.colors import SchemeColor
from openpyxl.drawing.geometry import FontReference
from openpyxl.drawing.geometry import StyleMatrixReference
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def gradient_fill_properties():
    from openpyxl.drawing.fill import GradientFillProperties

    return GradientFillProperties


@pytest.fixture
def transform_2d():
    from openpyxl.drawing.geometry import Transform2D

    return Transform2D


@pytest.fixture
def camera():
    from openpyxl.drawing.geometry import Camera

    return Camera


@pytest.fixture
def light_rig():
    from openpyxl.drawing.geometry import LightRig

    return LightRig


@pytest.fixture
def bevel():
    from openpyxl.drawing.geometry import Bevel

    return Bevel


@pytest.fixture
def sphere_coords():
    from openpyxl.drawing.geometry import SphereCoords

    return SphereCoords


@pytest.fixture
def vector_3d():
    from openpyxl.drawing.geometry import Vector3D

    return Vector3D


@pytest.fixture
def point_3d():
    from openpyxl.drawing.geometry import Point3D

    return Point3D


@pytest.fixture
def shape_style():
    from openpyxl.drawing.geometry import ShapeStyle

    return ShapeStyle


class TestGradientFillProperties:
    def test_ctor(self, gradient_fill_properties):
        fill = gradient_fill_properties()
        xml = tostring(fill.to_tree())
        expected = """
        <a:gradFill
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, gradient_fill_properties):
        src = """
        <a:gradFill
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        """
        node = fromstring(src)
        fill = gradient_fill_properties.from_tree(node)
        assert fill == gradient_fill_properties()


class TestTransform2D:
    def test_ctor(self, transform_2d):
        shapes = transform_2d()
        xml = tostring(shapes.to_tree())
        expected = """
        <xfrm xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, transform_2d):
        src = "<root/>"
        node = fromstring(src)
        shapes = transform_2d.from_tree(node)
        assert shapes == transform_2d()


class TestCamera:
    def test_ctor(self, camera):
        cam = camera(prst="legacyObliqueFront")
        xml = tostring(cam.to_tree())
        expected = '<camera prst="legacyObliqueFront"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, camera):
        src = '<camera prst="orthographicFront"/>'
        node = fromstring(src)
        cam = camera.from_tree(node)
        assert cam == camera(prst="orthographicFront")


class TestLightRig:
    def test_ctor(self, light_rig):
        rig = light_rig(rig="threePt", dir="t")
        xml = tostring(rig.to_tree())
        expected = '<lightRig rig="threePt" dir="t"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, light_rig):
        src = '<lightRig rig="threePt" dir="t"/>'
        node = fromstring(src)
        rig = light_rig.from_tree(node)
        assert rig == light_rig(rig="threePt", dir="t")


class TestBevel:
    def test_ctor(self, bevel):
        b = bevel(w=10, h=20)
        xml = tostring(b.to_tree())
        expected = '<bevel w="10" h="20"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, bevel):
        src = '<bevel w="101600" h="101600"/>'
        node = fromstring(src)
        b = bevel.from_tree(node)
        assert b == bevel(w=101600, h=101600)


class TestSphereCoords:
    def test_ctor(self, sphere_coords):
        rot = sphere_coords(lat=90, lon=45, rev=60)
        xml = tostring(rot.to_tree())
        expected = '<sphereCoords lat="90" lon="45" rev="60"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, sphere_coords):
        src = '<sphereCoords lat="90" lon="45" rev="60"/>'
        node = fromstring(src)
        rot = sphere_coords.from_tree(node)
        assert rot == sphere_coords(lat=90, lon=45, rev=60)


class TestVector3D:
    def test_ctor(self, vector_3d):
        vector = vector_3d(dx=100000, dy=300000, dz=50000)
        xml = tostring(vector.to_tree())
        expected = '<vector dx="100000" dy="300000" dz="50000"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, vector_3d):
        src = '<vector dx="100000" dy="300000" dz="50000"/>'
        node = fromstring(src)
        vector = vector_3d.from_tree(node)
        assert vector == vector_3d(dx=100000, dy=300000, dz=50000)


class TestPoint3D:
    def test_ctor(self, point_3d):
        pt = point_3d(x=40000, y=60000, z=100000)
        xml = tostring(pt.to_tree())
        expected = '<anchor x="40000" y="60000" z="100000"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, point_3d):
        src = '<anchor x="40000" y="60000" z="100000"/>'
        node = fromstring(src)
        pt = point_3d.from_tree(node)
        assert pt == point_3d(x=40000, y=60000, z=100000)


class TestShapeStyle:
    def test_ctor(self, shape_style):
        ln = StyleMatrixReference(
            idx=2,
            schemeClr=SchemeColor(val="accent1", shade=50000),
        )
        fill = StyleMatrixReference(idx=1, schemeClr=SchemeColor(val="accent1"))
        effect = StyleMatrixReference(idx=0, schemeClr=SchemeColor(val="accent1"))
        font = FontReference(idx="minor", schemeClr=SchemeColor(val="lt1"))
        style = shape_style(lnRef=ln, fillRef=fill, effectRef=effect, fontRef=font)
        xml = tostring(style.to_tree())
        expected = """
        <style xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
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
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, shape_style):
        src = """
        <style xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
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
        """
        node = fromstring(src)
        style = shape_style.from_tree(node)
        ln = StyleMatrixReference(
            idx=2,
            schemeClr=SchemeColor(val="accent1", shade=50000),
        )
        fill = StyleMatrixReference(idx=1, schemeClr=SchemeColor(val="accent1"))
        effect = StyleMatrixReference(idx=0, schemeClr=SchemeColor(val="accent1"))
        font = FontReference(idx="minor", schemeClr=SchemeColor(val="lt1"))
        expected = shape_style(lnRef=ln, fillRef=fill, effectRef=effect, fontRef=font)
        assert style == expected
