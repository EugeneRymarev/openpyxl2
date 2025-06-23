# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def picture_options():
    from openpyxl.chart.picture import PictureOptions

    return PictureOptions


class TestPictureOptions:
    def test_ctor(self, picture_options):
        picture = picture_options()
        xml = tostring(picture.to_tree())
        expected = "<pictureOptions/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, picture_options):
        src = "<pictureOptions/>"
        node = fromstring(src)
        picture = picture_options.from_tree(node)
        assert picture == picture_options()
