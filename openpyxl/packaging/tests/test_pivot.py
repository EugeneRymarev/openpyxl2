# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def pivot_cache():
    from openpyxl.packaging.workbook import PivotCache

    return PivotCache


class TestPivotCache:
    def test_ctor(self, pivot_cache):
        pivot = pivot_cache(cacheId=1, id="rId1")
        xml = tostring(pivot.to_tree())
        expected = """
        <pivotCache
                cacheId="1"
                r:id="rId1"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pivot_cache):
        src = '<pivotCache cacheId="2"/>'
        node = fromstring(src)
        pivot = pivot_cache.from_tree(node)
        assert pivot == pivot_cache(2)
