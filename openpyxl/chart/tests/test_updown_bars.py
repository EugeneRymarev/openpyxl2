# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def up_down_bars():
    from openpyxl.chart.updown_bars import UpDownBars

    return UpDownBars


class TestUpDownBars:
    def test_ctor(self, up_down_bars):
        bars = up_down_bars(gapWidth=150)
        xml = tostring(bars.to_tree())
        expected = """
        <upbars>
            <gapWidth val="150"/>
        </upbars>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, up_down_bars):
        src = """
        <upDownBars>
            <gapWidth val="156"/>
        </upDownBars>
        """
        node = fromstring(src)
        bars = up_down_bars.from_tree(node)
        assert bars == up_down_bars(gapWidth=156)
