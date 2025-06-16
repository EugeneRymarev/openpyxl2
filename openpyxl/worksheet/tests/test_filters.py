# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def filter_column():
    from openpyxl.worksheet.filters import FilterColumn

    return FilterColumn


@pytest.fixture
def sort_condition():
    from openpyxl.worksheet.filters import SortCondition

    return SortCondition


@pytest.fixture
def auto_filter():
    from openpyxl.worksheet.filters import AutoFilter

    return AutoFilter


@pytest.fixture
def sort_state():
    from openpyxl.worksheet.filters import SortState

    return SortState


@pytest.fixture
def icon_filter():
    from openpyxl.worksheet.filters import IconFilter

    return IconFilter


@pytest.fixture
def color_filter():
    from openpyxl.worksheet.filters import ColorFilter

    return ColorFilter


@pytest.fixture
def dynamic_filter():
    from openpyxl.worksheet.filters import DynamicFilter

    return DynamicFilter


@pytest.fixture
def custom_filter():
    from openpyxl.worksheet.filters import CustomFilter

    return CustomFilter


@pytest.fixture
def number_filter():
    from openpyxl.worksheet.filters import NumberFilter

    return NumberFilter


@pytest.fixture
def blank_filter():
    from openpyxl.worksheet.filters import BlankFilter

    return BlankFilter


@pytest.fixture
def string_filter():
    from openpyxl.worksheet.filters import StringFilter

    return StringFilter


@pytest.fixture
def custom_filters():
    from openpyxl.worksheet.filters import CustomFilters

    return CustomFilters


@pytest.fixture
def top10():
    from openpyxl.worksheet.filters import Top10

    return Top10


@pytest.fixture
def date_group_item():
    from openpyxl.worksheet.filters import DateGroupItem

    return DateGroupItem


@pytest.fixture
def filters():
    from openpyxl.worksheet.filters import Filters

    return Filters


class TestFilterColumn:
    def test_ctor(self, filter_column, filters):
        fltrs = filters(blank=True, filter=["0"])
        col = filter_column(colId=5, filters=fltrs)
        expected = """
        <filterColumn colId="5" hiddenButton="0" showButton="1">
            <filters blank="1">
                <filter val="0"></filter>
            </filters>
        </filterColumn>
        """
        xml = tostring(col.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, filter_column, filters):
        xml = """
        <filterColumn colId="5">
            <filters blank="1">
                <filter val="0"></filter>
            </filters>
        </filterColumn>
        """
        node = fromstring(xml)
        col = filter_column.from_tree(node)
        fltrs = filters(blank=True, filter=["0"])
        assert col == filter_column(colId=5, filters=fltrs)


class TestSortCondition:
    def test_ctor(self, sort_condition):
        cond = sort_condition(ref="A2:A3", descending=True)
        expected = '<sortCondition descending="1" ref="A2:A3"></sortCondition>'
        xml = tostring(cond.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, sort_condition):
        xml = '<sortCondition descending="1" ref="B4:B8"/>'
        node = fromstring(xml)
        cond = sort_condition.from_tree(node)
        assert cond == sort_condition(ref="B4:B8", descending=True)


class TestAutoFilter:
    def test_ctor(self, auto_filter):
        af = auto_filter("A2:A3")
        expected = '<autoFilter ref="A2:A3"/>'
        xml = tostring(af.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, auto_filter):
        xml = '<autoFilter ref="A2:A3"/>'
        node = fromstring(xml)
        af = auto_filter.from_tree(node)
        assert af == auto_filter(ref="A2:A3")

    def test_add_filter_column(self, auto_filter):
        af = auto_filter("A1:F1")
        af.add_filter_column(5, ["0"], blank=True)
        expected = """
        <autoFilter ref="A1:F1">
            <filterColumn colId="5" hiddenButton="0" showButton="1">
                <filters blank="1">
                    <filter val="0"></filter>
                </filters>
            </filterColumn>
        </autoFilter>
        """
        xml = tostring(af.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_add_sort_condition(self, auto_filter):
        af = auto_filter("A2:B3")
        af.add_sort_condition("B2:B3", descending=True)
        expected = """
        <autoFilter ref="A2:B3">
            <sortState ref="A2:B3">
                <sortCondition descending="1" ref="B2:B3"/>
            </sortState>
        </autoFilter>
        """
        xml = tostring(af.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_bool(self, auto_filter):
        assert bool(auto_filter("A2:A3")) is True
        assert bool(auto_filter()) is False


class TestSortState:
    def test_ctor(self, sort_state):
        sort = sort_state(ref="A1:D5")
        xml = tostring(sort.to_tree())
        expected = '<sortState ref="A1:D5"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, sort_state):
        src = """
        <sortState ref="B1:B3">
            <sortCondition ref="B1"/>
        </sortState>
        """
        node = fromstring(src)
        sort = sort_state.from_tree(node)
        assert sort.ref == "B1:B3"

    def test_bool(self, sort_state):
        assert bool(sort_state()) is False
        assert bool(sort_state(ref="B4:B8")) is True


class TestIconFilter:
    def test_ctor(self, icon_filter):
        flt = icon_filter(iconSet="3Flags")
        xml = tostring(flt.to_tree())
        expected = '<iconFilter iconSet="3Flags"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, icon_filter):
        src = '<iconFilter iconSet="5Rating"/>'
        node = fromstring(src)
        flt = icon_filter.from_tree(node)
        assert flt == icon_filter(iconSet="5Rating")


class TestColorFilter:
    def test_ctor(self, color_filter):
        flt = color_filter()
        xml = tostring(flt.to_tree())
        expected = "<colorFilter/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, color_filter):
        src = "<colorFilter/>"
        node = fromstring(src)
        flt = color_filter.from_tree(node)
        assert flt == color_filter()


class TestDynamicFilter:
    def test_ctor(self, dynamic_filter):
        flt = dynamic_filter(type="aboveAverage")
        xml = tostring(flt.to_tree())
        expected = '<dynamicFilter type="aboveAverage"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, dynamic_filter):
        src = '<dynamicFilter type="today"/>'
        node = fromstring(src)
        flt = dynamic_filter.from_tree(node)
        assert flt == dynamic_filter(type="today")


class TestCustomFilter:
    def test_ctor(self, custom_filter):
        flt = custom_filter(operator="greaterThanOrEqual", val="0.2")
        xml = tostring(flt.to_tree())
        expected = '<customFilter operator="greaterThanOrEqual" val="0.2"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, custom_filter):
        src = '<customFilter operator="greaterThanOrEqual" val="0.2"/>'
        node = fromstring(src)
        flt = custom_filter.from_tree(node)
        assert flt == custom_filter(operator="greaterThanOrEqual", val="0.2")

    @pytest.mark.parametrize(
        "value, typ",
        ([" ", "BlankFilter"], ["2.5", "NumberFilter"], ["ab", "StringFilter"]),
    )
    def test_convert(self, custom_filter, value, typ):
        flt = custom_filter(val=value)
        flt = flt.convert()
        assert flt.__class__.__name__ == typ

    @pytest.mark.parametrize(
        "operator, value, attrs",
        (
            ["equal", "*ab", {"operator": "endswith", "val": "ab", "exclude": "0"}],
            ["notEqual", "*ab", {"operator": "endswith", "val": "ab", "exclude": "1"}],
            ["notEqual", "c?n", {"operator": "wildcard", "val": "c?n", "exclude": "1"}],
        ),
    )
    def test_convert_string(self, custom_filter, operator, value, attrs):
        flt = custom_filter(operator, value)
        flt = flt.convert()
        assert dict(flt) == attrs


class TestNumberFilter:
    def test_ctor(self, number_filter):
        flt = number_filter(operator="greaterThanOrEqual", val=0.2)
        xml = tostring(flt.to_tree(tagname="customFilter"))
        expected = '<customFilter operator="greaterThanOrEqual" val="0.2"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, number_filter):
        src = '<customFilter operator="greaterThanOrEqual" val="0.2"/>'
        node = fromstring(src)
        flt = number_filter.from_tree(node)
        assert flt == number_filter(operator="greaterThanOrEqual", val=0.2)


class TestBlankFilter:
    def test_ctor(self, blank_filter):
        flt = blank_filter()
        xml = tostring(flt.to_tree(tagname="customFilter"))
        expected = '<customFilter operator="notEqual" val=" "/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, blank_filter):
        src = '<customFilter operator="greaterThanOrEqual" val="0.2"/>'
        node = fromstring(src)
        flt = blank_filter.from_tree(node)
        assert flt == blank_filter()


class TestStringFilter:
    def test_startswith(self, string_filter):
        flt = string_filter(operator="startswith", val="baa")
        xml = tostring(flt.to_tree())
        expected = '<customFilter operator="equal" val="baa*"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_not_startswith(self, string_filter):
        flt = string_filter(operator="startswith", val="baa", exclude=True)
        xml = tostring(flt.to_tree())
        expected = '<customFilter operator="notEqual" val="baa*"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_contains(self, string_filter):
        flt = string_filter(operator="contains", val="baa")
        xml = tostring(flt.to_tree())
        expected = '<customFilter operator="equal" val="*baa*"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_not_contain(self, string_filter):
        flt = string_filter(operator="contains", val="baa", exclude=True)
        xml = tostring(flt.to_tree())
        expected = '<customFilter operator="notEqual" val="*baa*"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_endswith(self, string_filter):
        flt = string_filter(operator="endswith", val="baa")
        xml = tostring(flt.to_tree())
        expected = '<customFilter operator="equal" val="*baa"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_not_endswith(self, string_filter):
        flt = string_filter(operator="endswith", val="baa", exclude=True)
        xml = tostring(flt.to_tree())
        expected = '<customFilter operator="notEqual" val="*baa"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.parametrize(
        "value, expected",
        [("*n", "~*n"), ("n?", "n~?"), ("b~i", "b~~i"), ("foo~*ba*", "foo~~~*ba~*")],
    )
    def test_escape(self, string_filter, value, expected):
        flt = string_filter("contains", value)
        out = flt._escape()
        assert out == expected

    @pytest.mark.parametrize(
        "expected, value",
        [("*n", "~*n"), ("n?", "n~?"), ("b~i", "b~~i"), ("foo~*ba*", "foo~~~*ba~*")],
    )
    def test_unescape(self, string_filter, value, expected):
        out = string_filter._unescape(value)
        assert out == expected

    @pytest.mark.parametrize("value", ["c*n", "c?n", "foo~*ba*"])
    def test_dont_escape_wildcard(self, string_filter, value):
        flt = string_filter("wildcard", value)
        out = flt._escape()
        assert out == value

    @pytest.mark.parametrize(
        "value, operator, term",
        [
            ("*ffg", "endswith", "ffg"),
            ("foo*", "startswith", "foo"),
            ("*foo*", "contains", "foo"),
            ("c*n", "wildcard", "c*n"),
            ("c*n", "wildcard", "c*n"),
        ],
    )
    def test_guess_operator(self, string_filter, value, operator, term):
        op, val = string_filter._guess_operator(value)
        assert (op, val) == (operator, term)


class TestCustomFilters:
    def test_ctor(self, custom_filters):
        flt = custom_filters(_and=True)
        xml = tostring(flt.to_tree())
        expected = '<customFilters and="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_blank(self, custom_filters, blank_filter):
        flt = custom_filters(customFilter=[blank_filter()])
        xml = tostring(flt.to_tree())
        expected = """
        <customFilters>
            <customFilter operator="notEqual" val=" "></customFilter>
        </customFilters>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_number(self, custom_filters, number_filter):
        flt = custom_filters(customFilter=[number_filter("lessThan", 4)])
        xml = tostring(flt.to_tree())
        expected = """
        <customFilters>
            <customFilter operator="lessThan" val="4"/>
        </customFilters>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_string(self, custom_filters, string_filter):
        flt = custom_filters(customFilter=[string_filter("contains", "xml")])
        xml = tostring(flt.to_tree())
        expected = """
        <customFilters>
            <customFilter operator="equal" val="*xml*"/>
        </customFilters>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_escape_string(self, custom_filters, string_filter):
        flt = custom_filters(customFilter=[string_filter("contains", "*xml")])
        xml = tostring(flt.to_tree())
        expected = """
        <customFilters>
            <customFilter operator="equal" val="*~*xml*"/>
        </customFilters>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_wildcard(self, custom_filters, string_filter):
        flt = custom_filters(customFilter=[string_filter("wildcard", "c?n")])
        xml = tostring(flt.to_tree())
        expected = """
        <customFilters>
            <customFilter operator="equal" val="c?n"/>
        </customFilters>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, custom_filters, custom_filter):
        src = """
        <customFilters>
            <customFilter operator="greaterThanOrEqual" val="1"/>
        </customFilters>
        """
        node = fromstring(src)
        flt = custom_filters.from_tree(node)
        fltrs = [custom_filter("greaterThanOrEqual", "1")]
        assert flt == custom_filters(customFilter=fltrs)


class TestTop10:
    def test_ctor(self, top10):
        flt = top10(percent=1, val=5, filterVal=6)
        xml = tostring(flt.to_tree())
        expected = '<top10 percent="1" val="5" filterVal="6"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, top10):
        src = '<top10 percent="1" val="5" filterVal="6"/>'
        node = fromstring(src)
        flt = top10.from_tree(node)
        assert flt == top10(percent=1, val=5, filterVal=6)


class TestDateGroupItem:
    def test_ctor(self, date_group_item):
        flt = date_group_item(dateTimeGrouping="day", year=2006, month=1, day=2)
        xml = tostring(flt.to_tree())
        expected = """
        <dateGroupItem
                year="2006"
                month="1"
                day="2"
                dateTimeGrouping="day"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, date_group_item):
        src = '<dateGroupItem year="2005" dateTimeGrouping="year"/>'
        node = fromstring(src)
        flt = date_group_item.from_tree(node)
        assert flt == date_group_item(dateTimeGrouping="year", year=2005)


class TestFilters:
    def test_ctor(self, filters):
        flt = filters(calendarType="gregorian")
        xml = tostring(flt.to_tree())
        expected = '<filters calendarType="gregorian"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_filters(self, filters):
        flt = filters()
        flt.filter = [1, 2, 3]
        xml = tostring(flt.to_tree())
        expected = """
        <filters>
            <filter val="1"/>
            <filter val="2"/>
            <filter val="3"/>
        </filters>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, filters):
        src = """
        <filters>
            <filter val="0.316588716"/>
            <filter val="0.667439395"/>
            <filter val="0.823086999"/>
        </filters>
        """
        node = fromstring(src)
        flt = filters.from_tree(node)
        assert flt == filters(filter=[0.316588716, 0.667439395, 0.823086999])
