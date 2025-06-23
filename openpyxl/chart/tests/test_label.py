# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def data_label_list():
    from ..label import DataLabelList

    return DataLabelList


@pytest.fixture
def data_label():
    from ..label import DataLabel

    return DataLabel


class TestDataLabelList:
    def test_ctor(self, data_label_list):
        labels = data_label_list(numFmt="0.0%")
        xml = tostring(labels.to_tree())
        expected = """
        <dLbls>
            <numFmt formatCode="0.0%"/>
        </dLbls>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, data_label_list):
        src = """
        <dLbls>
            <showLegendKey val="0"/>
            <showVal val="0"/>
            <showCatName val="0"/>
            <showSerName val="0"/>
            <showPercent val="0"/>
            <showBubbleSize val="0"/>
        </dLbls>
        """
        node = fromstring(src)
        dl = data_label_list.from_tree(node)
        assert dl.showLegendKey is False
        assert dl.showVal is False
        assert dl.showCatName is False
        assert dl.showSerName is False
        assert dl.showPercent is False
        assert dl.showBubbleSize is False


class TestDataLabel:
    def test_ctor(self, data_label):
        label = data_label()
        xml = tostring(label.to_tree())
        expected = """
        <dLbl>
            <idx val="0"></idx>
        </dLbl>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, data_label):
        src = """
        <dLbl>
            <idx val="6"></idx>
        </dLbl>
        """
        node = fromstring(src)
        label = data_label.from_tree(node)
        assert label == data_label(idx=6)
