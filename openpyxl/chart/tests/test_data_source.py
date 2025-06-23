# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def num_ref():
    from openpyxl.chart.data_source import NumRef

    return NumRef


@pytest.fixture
def str_ref():
    from openpyxl.chart.data_source import StrRef

    return StrRef


@pytest.fixture
def str_val():
    from openpyxl.chart.data_source import StrVal

    return StrVal


@pytest.fixture
def str_data():
    from openpyxl.chart.data_source import StrData

    return StrData


@pytest.fixture
def level_():
    from openpyxl.chart.data_source import Level

    return Level


@pytest.fixture
def multi_level_str_data():
    from openpyxl.chart.data_source import MultiLevelStrData

    return MultiLevelStrData


@pytest.fixture
def multi_level_str_ref():
    from openpyxl.chart.data_source import MultiLevelStrRef

    return MultiLevelStrRef


@pytest.fixture
def ax_data_source():
    from openpyxl.chart.data_source import AxDataSource

    return AxDataSource


class TestNumRef:
    def test_from_xml(self, num_ref):
        src = """
        <numRef>
            <f>Blatt1!$A$1:$A$12</f>
        </numRef>
        """
        node = fromstring(src)
        num = num_ref.from_tree(node)
        assert num.ref == "Blatt1!$A$1:$A$12"

    def test_to_xml(self, num_ref):
        num = num_ref(f="Blatt1!$A$1:$A$12")
        xml = tostring(num.to_tree("numRef"))
        expected = """
        <numRef>
            <f>Blatt1!$A$1:$A$12</f>
        </numRef>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_tree_degree_sign(self, num_ref):
        src = b"""
        <numRef>
            <f>Hoja1!$A$2:$B$2</f>
            <numCache>
                <formatCode>0\xc2\xb0</formatCode>
                <ptCount val="2"/>
                <pt idx="0">
                    <v>3</v>
                </pt>
                <pt idx="1">
                    <v>14</v>
                </pt>
            </numCache>
        </numRef>
        """
        node = fromstring(src)
        numRef = num_ref.from_tree(node)
        assert numRef.numCache.formatCode == "0\xb0"


class TestStrRef:
    def test_ctor(self, str_ref):
        data_source = str_ref(f="Sheet1!A1")
        xml = tostring(data_source.to_tree())
        expected = """
        <strRef>
            <f>Sheet1!A1</f>
        </strRef>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, str_ref):
        src = """
        <strRef>
            <f>'Render Start'!$A$2</f>
        </strRef>
        """
        node = fromstring(src)
        data_source = str_ref.from_tree(node)
        assert data_source == str_ref(f="'Render Start'!$A$2")


class TestStrVal:
    def test_ctor(self, str_val):
        val = str_val(v="something")
        xml = tostring(val.to_tree())
        expected = """
        <strVal idx="0">
            <v>something</v>
        </strVal>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, str_val):
        src = """
        <pt idx="4">
            <v>else</v>
        </pt>
        """
        node = fromstring(src)
        val = str_val.from_tree(node)
        assert val == str_val(idx=4, v="else")


class TestStrData:
    def test_ctor(self, str_data):
        data_source = str_data(ptCount=1)
        xml = tostring(data_source.to_tree())
        expected = """
        <strData>
            <ptCount val="1"></ptCount>
        </strData>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, str_data):
        src = """
        <strData>
            <ptCount val="4"></ptCount>
        </strData>
        """
        node = fromstring(src)
        data_source = str_data.from_tree(node)
        assert data_source == str_data(ptCount=4)


class TestLevel:
    def test_ctor(self, level_):
        level = level_()
        xml = tostring(level.to_tree())
        expected = "<lvl/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, level_):
        src = "<root/>"
        node = fromstring(src)
        level = level_.from_tree(node)
        assert level == level_()


class TestMultiLevelStrData:
    def test_ctor(self, multi_level_str_data):
        multidata = multi_level_str_data()
        xml = tostring(multidata.to_tree())
        expected = "<multiLvlStrData/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, multi_level_str_data):
        src = "<multiLvlStrData/>"
        node = fromstring(src)
        multidata = multi_level_str_data.from_tree(node)
        assert multidata == multi_level_str_data()


class TestMultiLevelStrRef:
    def test_ctor(self, multi_level_str_ref):
        multi_ref = multi_level_str_ref(f="Sheet1!$A$1:$B$10")
        xml = tostring(multi_ref.to_tree())
        expected = """
        <multiLvlStrRef>
            <f>Sheet1!$A$1:$B$10</f>
        </multiLvlStrRef>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, multi_level_str_ref):
        src = """
        <multiLvlStrRef>
            <f>Sheet1!$A$1:$B$10</f>
        </multiLvlStrRef>
        """
        node = fromstring(src)
        multi_ref = multi_level_str_ref.from_tree(node)
        assert multi_ref == multi_level_str_ref(f="Sheet1!$A$1:$B$10")


class TestAxDataSource:
    def test_ctor(self, ax_data_source, str_ref):
        dummy = str_ref(f="")
        ax = ax_data_source(strRef=dummy)
        xml = tostring(ax.to_tree())
        expected = """
        <cat>
            <strRef>
                <f/>
            </strRef>
        </cat>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_no_source(self, ax_data_source):
        with pytest.raises(TypeError):
            ax = ax_data_source()

    def test_from_xml(self, ax_data_source, str_ref):
        src = """
        <cat>
            <strRef>
                <f/>
            </strRef>
        </cat>
        """
        node = fromstring(src)
        dummy = str_ref()
        ax = ax_data_source.from_tree(node)
        assert ax == ax_data_source(strRef=dummy)
