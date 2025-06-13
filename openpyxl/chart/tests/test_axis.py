# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def scaling():
    from openpyxl.chart.axis import Scaling

    return Scaling


@pytest.fixture
def _base_axis():
    from openpyxl.chart.axis import _BaseAxis

    return _BaseAxis


@pytest.fixture
def text_axis():
    from openpyxl.chart.axis import TextAxis

    return TextAxis


@pytest.fixture
def numeric_axis():
    from openpyxl.chart.axis import NumericAxis

    return NumericAxis


@pytest.fixture
def date_axis():
    from openpyxl.chart.axis import DateAxis

    return DateAxis


@pytest.fixture
def series_axis():
    from openpyxl.chart.axis import SeriesAxis

    return SeriesAxis


@pytest.fixture
def display_units_label():
    from openpyxl.chart.axis import DisplayUnitsLabel

    return DisplayUnitsLabel


@pytest.fixture
def display_units_label_list():
    from openpyxl.chart.axis import DisplayUnitsLabelList

    return DisplayUnitsLabelList


@pytest.fixture
def chart_lines():
    from openpyxl.chart.axis import ChartLines

    return ChartLines


class TestScale:
    def test_ctor(self, scaling):
        scale = scaling()
        xml = tostring(scale.to_tree())
        expected = """
        <scaling>
            <orientation val="minMax"></orientation>
        </scaling>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, scaling):
        xml = """
        <scaling>
            <logBase val="10"/>
            <orientation val="minMax"/>
        </scaling>
        """
        node = fromstring(xml)
        scale = scaling.from_tree(node)
        assert scale == scaling(logBase=10)


class TestAxis:
    def test_ctor(self, _base_axis, scaling):
        axis = _base_axis(axId=10, crossAx=100)
        xml = tostring(axis.to_tree(tagname="baseAxis"))
        expected = """
        <baseAxis>
            <axId val="10"></axId>
            <scaling>
                <orientation val="minMax"></orientation>
            </scaling>
            <axPos val="l"/>
            <majorTickMark val="none"/>
            <minorTickMark val="none"/>
            <crossAx val="100"/>
        </baseAxis>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestTextAxis:
    def test_ctor(self, text_axis):
        axis = text_axis(axId=10, crossAx=100)
        xml = tostring(axis.to_tree())
        expected = """
        <catAx>
            <axId val="10"></axId>
            <scaling>
                <orientation val="minMax"></orientation>
            </scaling>
            <axPos val="l"/>
            <majorTickMark val="none"/>
            <minorTickMark val="none"/>
            <crossAx val="100"/>
            <lblOffset val="100"/>
        </catAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, text_axis):
        src = """
        <catAx>
            <axId val="2065276984"/>
            <scaling>
                <orientation val="minMax"/>
            </scaling>
            <delete val="0"/>
            <axPos val="b"/>
            <majorTickMark val="out"/>
            <minorTickMark val="none"/>
            <tickLblPos val="nextTo"/>
            <crossAx val="2056619928"/>
            <crosses val="autoZero"/>
            <crossesAt val="30"/>
            <auto val="1"/>
            <lblAlgn val="ctr"/>
            <lblOffset val="100"/>
            <noMultiLvlLbl val="0"/>
        </catAx>
        """
        node = fromstring(src)
        axis = text_axis.from_tree(node)
        assert axis.scaling.orientation == "minMax"
        assert axis.auto is True
        assert axis.majorTickMark == "out"
        assert axis.minorTickMark is None
        assert axis.crossesAt == 30


class TestValAx:
    def test_ctor(self, numeric_axis):
        axis = numeric_axis(axId=100, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <valAx>
            <axId val="100"></axId>
            <scaling>
                <orientation val="minMax"></orientation>
            </scaling>
            <axPos val="l"/>
            <majorGridlines/>
            <majorTickMark val="none"/>
            <minorTickMark val="none"/>
            <crossAx val="10"/>
        </valAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, numeric_axis):
        src = """
        <valAx>
            <axId val="2056619928"/>
            <scaling>
                <logBase val="10"/>
                <orientation val="minMax"/>
            </scaling>
            <delete val="0"/>
            <axPos val="l"/>
            <majorGridlines/>
            <numFmt formatCode="General" sourceLinked="1"/>
            <majorTickMark val="out"/>
            <minorTickMark val="none"/>
            <tickLblPos val="nextTo"/>
            <crossAx val="2065276984"/>
            <crosses val="autoZero"/>
            <crossBetween val="between"/>
        </valAx>
        """
        node = fromstring(src)
        axis = numeric_axis.from_tree(node)
        assert axis.delete is False
        assert axis.crossAx == 2065276984
        assert axis.crossBetween == "between"
        assert axis.scaling.logBase == 10


class TestDateAx:
    def test_ctor(self, date_axis):
        axis = date_axis(axId=500, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <dateAx>
            <axId val="500"></axId>
            <scaling>
                <orientation val="minMax"></orientation>
            </scaling>
            <axPos val="l"/>
            <majorTickMark val="none"/>
            <minorTickMark val="none"/>
            <crossAx val="10"/>
        </dateAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, date_axis):
        from openpyxl.chart.data_source import NumFmt

        src = """
        <dateAx>
            <axId val="20"/>
            <scaling>
                <orientation val="minMax"/>
            </scaling>
            <delete val="0"/>
            <axPos val="b"/>
            <numFmt formatCode="d-mmm" sourceLinked="1"/>
            <majorTickMark val="out"/>
            <minorTickMark val="none"/>
            <tickLblPos val="nextTo"/>
            <crossAx val="10"/>
            <crosses val="autoZero"/>
            <auto val="1"/>
            <lblOffset val="100"/>
            <baseTimeUnit val="months"/>
        </dateAx>
        """
        node = fromstring(src)
        axis = date_axis.from_tree(node)
        expected = date_axis(
            axId=20,
            crossAx=10,
            axPos="b",
            delete=False,
            numFmt=NumFmt("d-mmm", True),
            majorTickMark="out",
            crosses="autoZero",
            tickLblPos="nextTo",
            auto=True,
            lblOffset=100,
            baseTimeUnit="months",
        )
        assert axis == expected


class TestSeriesAxis:
    def test_ctor(self, series_axis):
        axis = series_axis(axId=1000, crossAx=10)
        xml = tostring(axis.to_tree())
        expected = """
        <serAx>
            <axId val="1000"></axId>
            <scaling>
                <orientation val="minMax"></orientation>
            </scaling>
            <axPos val="l"/>
            <majorTickMark val="none"/>
            <minorTickMark val="none"/>
            <crossAx val="10"/>
        </serAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, series_axis):
        src = """
        <serAx>
            <axId val="1000"></axId>
            <scaling>
                <orientation val="minMax"></orientation>
            </scaling>
            <axPos val="l"/>
            <crossAx val="10"/>
        </serAx>
        """
        node = fromstring(src)
        axis = series_axis.from_tree(node)
        assert axis == series_axis()


class TestDispUnitsLabel:
    def test_ctor(self, display_units_label):
        axis = display_units_label()
        xml = tostring(axis.to_tree())
        expected = "<dispUnitsLbl/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, display_units_label):
        src = "<dispUnitsLbl/>"
        node = fromstring(src)
        axis = display_units_label.from_tree(node)
        assert axis == display_units_label()


class TestDisplayUnitList:
    def test_ctor(self, display_units_label_list):
        axis = display_units_label_list()
        xml = tostring(axis.to_tree())
        expected = "<dispUnits/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, display_units_label_list):
        src = "<dispUnits/>"
        node = fromstring(src)
        axis = display_units_label_list.from_tree(node)
        assert axis == display_units_label_list()


class TestChartLines:
    def test_ctor(self, chart_lines):
        axis = chart_lines()
        xml = tostring(axis.to_tree())
        expected = "<chartLines/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, chart_lines):
        src = "<chartLines/>"
        node = fromstring(src)
        axis = chart_lines.from_tree(node)
        assert axis == chart_lines()
