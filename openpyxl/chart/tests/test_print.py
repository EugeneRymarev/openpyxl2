# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def print_settings():
    from openpyxl.chart.print_settings import PrintSettings

    return PrintSettings


@pytest.fixture
def page_margins():
    from openpyxl.chart.print_settings import PageMargins

    return PageMargins


class TestPrintSettings:
    def test_ctor(self, print_settings):
        chartspace = print_settings()
        xml = tostring(chartspace.to_tree())
        expected = "<printSettings/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, print_settings):
        src = "<printSettings/>"
        node = fromstring(src)
        chartspace = print_settings.from_tree(node)
        assert chartspace == print_settings()


class TestPageMargins:
    def test_ctor(self, page_margins):
        pm = page_margins()
        xml = tostring(pm.to_tree())
        expected = """
        <pageMargins
                b="1"
                l="0.75"
                r="0.75"
                t="1"
                header="0.5"
                footer="0.5"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, page_margins):
        src = """
        <pageMargins
                b="1.0"
                l="0.75"
                r="0.75"
                t="1.0"
                header="0.5"
                footer="0.5"/>
        """
        node = fromstring(src)
        pm = page_margins.from_tree(node)
        assert pm == page_margins()
