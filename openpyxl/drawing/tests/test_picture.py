# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def picture_locking():
    from openpyxl.drawing.picture import PictureLocking

    return PictureLocking


@pytest.fixture
def non_visual_picture_properties():
    from openpyxl.drawing.picture import NonVisualPictureProperties

    return NonVisualPictureProperties


@pytest.fixture
def picture_non_visual():
    from openpyxl.drawing.picture import PictureNonVisual

    return PictureNonVisual


@pytest.fixture
def picture_frame():
    from openpyxl.drawing.picture import PictureFrame

    return PictureFrame


class TestPictureLocking:
    def test_ctor(self, picture_locking):
        graphic = picture_locking(noChangeAspect=True)
        xml = tostring(graphic.to_tree())
        expected = """
        <picLocks
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"
                noChangeAspect="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, picture_locking):
        src = '<picLocks noRot="1"/>'
        node = fromstring(src)
        graphic = picture_locking.from_tree(node)
        assert graphic == picture_locking(noRot=1)


class TestNonVisualPictureProperties:
    def test_ctor(self, non_visual_picture_properties):
        graphic = non_visual_picture_properties()
        xml = tostring(graphic.to_tree())
        expected = "<cNvPicPr/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, non_visual_picture_properties):
        src = "<cNvPicPr/>"
        node = fromstring(src)
        graphic = non_visual_picture_properties.from_tree(node)
        assert graphic == non_visual_picture_properties()


class TestPictureNonVisual:
    def test_ctor(self, picture_non_visual):
        graphic = picture_non_visual()
        xml = tostring(graphic.to_tree())
        expected = """
        <nvPicPr>
            <cNvPr descr="Name of file" id="0" name="Image 1"/>
            <cNvPicPr/>
        </nvPicPr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, picture_non_visual):
        src = """
        <nvPicPr>
            <cNvPr descr="Name of file" id="0" name="Image 1"/>
            <cNvPicPr/>
        </nvPicPr>
        """
        node = fromstring(src)
        graphic = picture_non_visual.from_tree(node)
        assert graphic == picture_non_visual()


class TestPicture:
    def test_ctor(self, picture_frame):
        graphic = picture_frame()
        xml = tostring(graphic.to_tree())
        expected = """
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <nvPicPr>
                <cNvPr descr="Name of file" id="0" name="Image 1"/>
                <cNvPicPr/>
            </nvPicPr>
            <blipFill>
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
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, picture_frame):
        src = """
        <pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <nvPicPr>
                <cNvPr descr="Picture" id="1" name="Image 1"/>
                <cNvPicPr/>
            </nvPicPr>
            <blipFill>
                <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        cstate="print"
                        r:embed="rId1"/>
                <a:stretch>
                    <a:fillRect/>
                </a:stretch>
            </blipFill>
            <spPr>
                <a:xfrm>
                    <a:off x="303" y="0"/>
                    <a:ext cx="321" cy="88"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
                <a:ln>
                    <a:prstDash val="solid"/>
                </a:ln>
            </spPr>
        </pic>
        """
        node = fromstring(src)
        graphic = picture_frame.from_tree(node)
        xml = tostring(graphic.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff
