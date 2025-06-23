# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def simple_test_props():
    from openpyxl.worksheet.properties import WorksheetProperties

    wsp = WorksheetProperties()
    wsp.filterMode = False
    wsp.tabColor = "FF123456"
    wsp.pageSetUpPr.fitToPage = False
    return wsp


def test_ctor():
    from openpyxl.worksheet.properties import Outline
    from openpyxl.worksheet.properties import WorksheetProperties

    color_test = "F0F0F0"
    outline_pr = Outline(summaryBelow=True, summaryRight=True)
    wsprops = WorksheetProperties(tabColor=color_test, outlinePr=outline_pr)
    assert dict(wsprops) == {}
    assert dict(wsprops.outlinePr) == {"summaryBelow": "1", "summaryRight": "1"}
    assert dict(wsprops.tabColor) == {"rgb": "00F0F0F0"}


def test_write_properties(simple_test_props):
    xml = tostring(simple_test_props.to_tree())
    expected = """
    <sheetPr filterMode="0">
        <tabColor rgb="FF123456"/>
        <outlinePr summaryBelow="1" summaryRight="1"></outlinePr>
        <pageSetUpPr fitToPage="0"/>
    </sheetPr>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_parse_properties(datadir, simple_test_props):
    from openpyxl.worksheet.properties import WorksheetProperties

    datadir.chdir()
    with open("sheetPr2.xml") as src:
        content = src.read()
    xml = fromstring(content)
    pi = WorksheetProperties.from_tree(xml)
    assert dict(pi) == dict(simple_test_props)
    assert pi.tabColor == simple_test_props.tabColor
    assert dict(pi.pageSetUpPr) == dict(simple_test_props.pageSetUpPr)
