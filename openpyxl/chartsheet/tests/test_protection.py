# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def chartsheet_protection():
    from openpyxl.chartsheet.protection import ChartsheetProtection

    return ChartsheetProtection


class TestChartsheetProtection:
    def test_read(self, chartsheet_protection):
        src = """
        <sheetProtection
                algorithmName="SHA-512"
                hashValue="frzjS2RlYHFtCLJwGZod5i+414zeFhyLnVYY6A++RjBbtDfGng4+nU0Qpo1ZyIlXnfffImweadNwHNy5Bmm+zw=="
                saltValue="Bo89+SCcqbFEcOS/6LcjBw=="
                spinCount="100000" content="1"
                objects="1"
                xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
        """
        xml = fromstring(src)
        protection = chartsheet_protection.from_tree(xml)
        assert protection.algorithmName == "SHA-512"
        assert protection.saltValue == "Bo89+SCcqbFEcOS/6LcjBw=="

    def test_write(self, chartsheet_protection):
        protection = chartsheet_protection()
        protection.saltValue = "Bo89+SCcqbFEcOS/6LcjBw=="
        protection.content = "1"
        protection.objects = "1"
        protection.algorithmName = "SHA-512"
        protection.spinCount = "100000"
        expected = """
        <sheetProtection
                algorithmName="SHA-512"
                saltValue="Bo89+SCcqbFEcOS/6LcjBw=="
                spinCount="100000" content="1"
                objects="1"
                xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
        """
        xml = tostring(protection.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_password(self, chartsheet_protection):
        prot = chartsheet_protection()
        prot.password = "secret"
        assert prot.password == "DAA7"
