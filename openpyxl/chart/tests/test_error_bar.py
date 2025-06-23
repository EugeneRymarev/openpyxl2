# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def error_bars():
    from openpyxl.chart.error_bar import ErrorBars

    return ErrorBars


class TestErrorBar:

    def test_ctor(self, error_bars):
        bar = error_bars()
        xml = tostring(bar.to_tree())
        expected = """
        <errBars>
            <errBarType val="both"></errBarType>
            <errValType val="fixedVal"></errValType>
        </errBars>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, error_bars):
        src = """
        <errBars>
            <errDir val="x"/>
            <errBarType val="both"/>
            <errValType val="fixedVal"/>
            <noEndCap val="1"/>
            <val val="10.0"/>
        </errBars>
        """
        node = fromstring(src)
        bar = error_bars.from_tree(node)
        assert bar == error_bars(noEndCap=True, errDir="x", val=10)
