# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import tostring


@pytest.fixture
def break_():
    from openpyxl.worksheet.pagebreak import Break

    return Break


@pytest.fixture
def row_break():
    from openpyxl.worksheet.pagebreak import RowBreak

    return RowBreak


@pytest.fixture
def col_break():
    from openpyxl.worksheet.pagebreak import ColBreak

    return ColBreak


class TestBreak:
    def test_ctor(self, break_):
        brk = break_()
        assert dict(brk) == {"id": "0", "man": "1", "max": "16383", "min": "0"}
        xml = tostring(brk.to_tree())
        expected = '<brk id="0" man="1" max="16383" min="0"></brk>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestRowBreak:
    def test_no_breaks(self, row_break):
        pb = row_break()
        assert dict(pb) == {"count": "0", "manualBreakCount": "0"}

    def test_append(self, row_break):
        pb = row_break()
        pb.append()
        assert dict(pb) == {"count": "1", "manualBreakCount": "1"}

    def test_to_tree(self, row_break):
        pb = row_break()
        pb.append()
        xml = tostring(pb.to_tree())
        expected = """
        <rowBreaks count="1" manualBreakCount="1">
            <brk id="1" man="1" max="16383" min="0"></brk>
        </rowBreaks>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestColBreak:
    def test_to_tree(self, col_break):
        pb = col_break()
        pb.append()
        xml = tostring(pb.to_tree())
        expected = """
        <colBreaks count="1" manualBreakCount="1">
            <brk id="1" man="1" max="16383" min="0"></brk>
        </colBreaks>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
