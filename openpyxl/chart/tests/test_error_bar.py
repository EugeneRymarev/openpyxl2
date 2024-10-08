# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def ErrorBars():
    from ..error_bar import ErrorBars

    return ErrorBars


class TestErrorBar:

    def test_ctor(self, ErrorBars):
        bar = ErrorBars()
        xml = tostring(bar.to_tree())
        expected = """
        <errBars>
            <errBarType val="both"></errBarType>
            <errValType val="fixedVal"></errValType>
        </errBars>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, ErrorBars):
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
        bar = ErrorBars.from_tree(node)
        assert bar == ErrorBars(noEndCap=True, errDir="x", val=10)
