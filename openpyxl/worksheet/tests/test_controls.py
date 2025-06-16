# Copyright (c) 2010-2025 openpyxl
import io
import zipfile

import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.worksheet.ole import AnchorMarker
from openpyxl.worksheet.ole import ObjectAnchor
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def control_property():
    from openpyxl.worksheet.controls import ControlProperty

    return ControlProperty


@pytest.fixture
def control():
    from openpyxl.worksheet.controls import Control

    return Control


@pytest.fixture
def control_list():
    from openpyxl.worksheet.controls import ControlList

    return ControlList


@pytest.fixture
def form_control():
    from openpyxl.worksheet.controls import FormControl

    return FormControl


@pytest.fixture
def active_x_control():
    from openpyxl.worksheet.controls import ActiveXControl

    return ActiveXControl


class TestControlProperty:
    def test_ctor(self, control_property):
        _from = AnchorMarker()
        to = AnchorMarker()
        anchor = ObjectAnchor(_from=_from, to=to)
        prop = control_property(anchor=anchor)
        xml = tostring(prop.to_tree())
        expected = """
        <controlPr
                autoFill="1"
                autoLine="1"
                autoPict="1"
                cf="pict"
                defaultSize="1"
                disabled="0"
                locked="1"
                print="1"
                recalcAlways="0"
                uiObject="0"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
            <anchor moveWithCells="0" sizeWithCells="0">
                <from>
                    <xdr:col>0</xdr:col>
                    <xdr:colOff>0</xdr:colOff>
                    <xdr:row>0</xdr:row>
                    <xdr:rowOff>0</xdr:rowOff>
                </from>
                <to>
                    <xdr:col>0</xdr:col>
                    <xdr:colOff>0</xdr:colOff>
                    <xdr:row>0</xdr:row>
                    <xdr:rowOff>0</xdr:rowOff>
                </to>
            </anchor>
        </controlPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, control_property):
        src = """
        <controlPr
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                autoLine="0">
            <anchor moveWithCells="1">
                <from>
                    <xdr:col>4</xdr:col>
                    <xdr:colOff>704850</xdr:colOff>
                    <xdr:row>59</xdr:row>
                    <xdr:rowOff>114300</xdr:rowOff>
                </from>
                <to>
                    <xdr:col>4</xdr:col>
                    <xdr:colOff>1190625</xdr:colOff>
                    <xdr:row>61</xdr:row>
                    <xdr:rowOff>47625</xdr:rowOff>
                </to>
            </anchor>
        </controlPr>
        """
        node = fromstring(src)
        prop = control_property.from_tree(node)
        _from = AnchorMarker(col=4, colOff=704850, row=59, rowOff=114300)
        to = AnchorMarker(col=4, colOff=1190625, row=61, rowOff=47625)
        anchor = ObjectAnchor(_from=_from, to=to, moveWithCells=True)
        assert prop == control_property(anchor=anchor, autoLine=False)


class TestControl:
    def test_ctor(self, control):
        ctrl = control(shapeId=1)
        xml = tostring(ctrl.to_tree())
        expected = '<control shapeId="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, control):
        src = """
        <control
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                shapeId="47120"
                r:id="rId8"
                name="RefProjectButton"/>
        """
        node = fromstring(src)
        ctrl = control.from_tree(node)
        assert ctrl == control(shapeId=47120, name="RefProjectButton", id="rId8")


class TestControlList:
    def test_ctor(self, control_list):
        ctrls = control_list()
        xml = tostring(ctrls.to_tree())
        expected = "<controls/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, control_list):
        src = """
        <controls
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <mc:AlternateContent
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                <mc:Choice Requires="x14">
                    <control shapeId="47120" r:id="rId8" name="RefWECProjectButton">
                        <controlPr defaultSize="0" autoLine="0" r:id="rId9">
                            <anchor moveWithCells="1">
                                <from>
                                    <xdr:col>3</xdr:col>
                                    <xdr:colOff>161925</xdr:colOff>
                                    <xdr:row>24</xdr:row>
                                    <xdr:rowOff>95250</xdr:rowOff>
                                </from>
                                <to>
                                    <xdr:col>6</xdr:col>
                                    <xdr:colOff>171450</xdr:colOff>
                                    <xdr:row>27</xdr:row>
                                    <xdr:rowOff>95250</xdr:rowOff>
                                </to>
                            </anchor>
                        </controlPr>
                    </control>
                </mc:Choice>
                <mc:Fallback>
                    <control shapeId="47120" r:id="rId8" name="RefWECProjectButton"/>
                </mc:Fallback>
            </mc:AlternateContent>
        </controls>
        """
        node = fromstring(src)
        ctrls = control_list.from_tree(node)
        assert len(ctrls) == 1


class TestFormControl:
    def test_ctor(self, form_control):
        ctrl = form_control(objectType="Button", lockText=True)
        xml = tostring(ctrl.to_tree())
        expected = """
        <formControlPr
                xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                colored="0"
                dropLines="8"
                dx="80"
                firstButton="0"
                horiz="0"
                justLastX="0"
                lockText="1"
                multiLine="0"
                noThreeD="0"
                noThreeD2="0"
                objectType="Button"
                passwordEdit="0"
                verticalBar="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, form_control):
        src = """
        <formControlPr
                xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                objectType="Button"
                lockText="1"/>
        """
        node = fromstring(src)
        ctrl = form_control.from_tree(node)
        assert ctrl == form_control(objectType="Button", lockText=True)

    def test_path(self, form_control):
        ctrl = form_control("button")
        ctrl.counter = 4
        assert ctrl.path == "/xl/ctrlProps/ctrlProp4.xml"

    def test_write(self, form_control):
        archive = zipfile.ZipFile(io.BytesIO(), "w")
        ctrl = form_control("Button")
        ctrl.counter = 1
        manifest = []
        ctrl._write(archive, manifest)
        assert archive.namelist() == ["xl/ctrlProps/ctrlProp1.xml"]


class TestActiveXControl:
    def test_ctor(self, active_x_control):
        ctrl = active_x_control(id="rId1", persistence="persistStreamInit")
        xml = tostring(ctrl.to_tree())
        expected = """
        <ax:ocx ax:classid="{8BD21D50-EC42-11CE-9E0D-00AA006002F3}"
                ax:persistence="persistStreamInit"
                r:id="rId1"
                xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, active_x_control):
        src = """
        <ocx classid="{8BD21D50-EC42-11CE-9E0D-00AA006002F3}"
             r:id="rId1"
             xmlns="http://schemas.microsoft.com/office/2006/activeX"
             persistence="persistStreamInit"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        node = fromstring(src)
        ctrl = active_x_control.from_tree(node)
        assert ctrl == active_x_control(id="rId1", persistence="persistStreamInit")

    def test_path(self, active_x_control):
        ctrl = active_x_control("rId4", persistence="persistStreamInit")
        ctrl.counter = 4
        assert ctrl.path == "/xl/activeX/activeX4.xml"

    def test_write(self, active_x_control):
        archive = zipfile.ZipFile(io.BytesIO(), "w")
        ctrl = active_x_control(persistence="persistStreamInit")
        ctrl.counter = 1
        ctrl.bin = b"\001"
        manifest = []
        ctrl._write(archive, manifest)
        expected = [
            "xl/activeX/activeX1.bin",
            "xl/activeX/_rels/activeX1.xml.rels",
            "xl/activeX/activeX1.xml",
        ]
        assert archive.namelist() == expected
