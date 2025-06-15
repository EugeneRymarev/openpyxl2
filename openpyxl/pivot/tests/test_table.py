# Copyright (c) 2010-2025 openpyxl
import io
import zipfile

import pytest
from openpyxl.packaging.manifest import Manifest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def pivot_field():
    from openpyxl.pivot.table import PivotField

    return PivotField


@pytest.fixture
def field_item():
    from openpyxl.pivot.table import FieldItem

    return FieldItem


@pytest.fixture
def row_col_item():
    from openpyxl.pivot.table import RowColItem

    return RowColItem


@pytest.fixture
def data_field():
    from openpyxl.pivot.table import DataField

    return DataField


@pytest.fixture
def location():
    from openpyxl.pivot.table import Location

    return Location


@pytest.fixture
def pivot_table_style():
    from openpyxl.pivot.table import PivotTableStyle

    return PivotTableStyle


@pytest.fixture
def table_definition():
    from openpyxl.pivot.table import TableDefinition

    return TableDefinition


@pytest.fixture
def page_field():
    from openpyxl.pivot.table import PageField

    return PageField


@pytest.fixture
def reference():
    from openpyxl.pivot.table import Reference

    return Reference


@pytest.fixture
def pivot_area():
    from openpyxl.pivot.table import PivotArea

    return PivotArea


@pytest.fixture
def chart_format():
    from openpyxl.pivot.table import ChartFormat

    return ChartFormat


@pytest.fixture
def pivot_filter():
    from openpyxl.pivot.table import PivotFilter

    return PivotFilter


@pytest.fixture
def format_():
    from openpyxl.pivot.table import Format

    return Format


@pytest.fixture
def conditional_format():
    from openpyxl.pivot.table import ConditionalFormat

    return ConditionalFormat


@pytest.fixture
def conditional_format_list():
    from openpyxl.pivot.table import ConditionalFormatList

    return ConditionalFormatList


@pytest.fixture
def auto_filter():
    from openpyxl.worksheet.filters import AutoFilter
    from openpyxl.worksheet.filters import CustomFilter
    from openpyxl.worksheet.filters import CustomFilters
    from openpyxl.worksheet.filters import FilterColumn

    cf1 = CustomFilter(operator="greaterThanOrEqual", val="1")
    cf2 = CustomFilter(operator="lessThanOrEqual", val="2")
    filters = CustomFilters(_and=True, customFilter=(cf1, cf2))
    col = FilterColumn(colId=0, customFilters=filters)
    af = AutoFilter(ref="A1", filterColumn=[col])
    return af


class TestPivotField:
    def test_ctor(self, pivot_field):
        field = pivot_field()
        xml = tostring(field.to_tree())
        expected = """
        <pivotField
                compact="1"
                defaultSubtotal="1"
                dragOff="1"
                dragToCol="1"
                dragToData="1"
                dragToPage="1"
                dragToRow="1"
                itemPageCount="10"
                outline="1"
                showAll="1"
                showDropDowns="1"
                sortType="manual"
                subtotalTop="1"
                topAutoShow="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pivot_field):
        src = "<pivotField/>"
        node = fromstring(src)
        field = pivot_field.from_tree(node)
        assert field == pivot_field()


class TestFieldItem:
    def test_ctor(self, field_item):
        item = field_item()
        xml = tostring(item.to_tree())
        expected = '<item sd="1" t="data"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, field_item):
        src = '<item m="1" x="2"/>'
        node = fromstring(src)
        item = field_item.from_tree(node)
        assert item == field_item(m=True, x=2)


class TestRowColItem:
    def test_ctor(self, row_col_item):
        fut = row_col_item(x=[4])
        xml = tostring(fut.to_tree())
        expected = """
        <i i="0" r="0" t="data">
            <x v="4"/>
        </i>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, row_col_item):
        src = """
        <i r="1">
            <x v="2"/>
        </i>
        """
        node = fromstring(src)
        fut = row_col_item.from_tree(node)
        assert fut == row_col_item(r=1, x=[2])


class TestDataField:
    def test_ctor(self, data_field):
        df = data_field(fld=1)
        xml = tostring(df.to_tree())
        expected = """
        <dataField
                baseField="-1"
                baseItem="1048832"
                fld="1"
                showDataAs="normal"
                subtotal="sum"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, data_field):
        src = """
        <dataField name="Sum of impressions" fld="4" baseField="0" baseItem="0"/>
        """
        node = fromstring(src)
        df = data_field.from_tree(node)
        expected = data_field(fld=4, name="Sum of impressions", baseField=0, baseItem=0)
        assert df == expected


class TestLocation:
    def test_ctor(self, location):
        loc = location(ref="A3:E14", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)
        xml = tostring(loc.to_tree())
        expected = """
        <location
                ref="A3:E14"
                firstHeaderRow="1"
                firstDataRow="2"
                firstDataCol="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, location):
        src = """
        <location
                ref="A3:E14"
                firstHeaderRow="1"
                firstDataRow="2"
                firstDataCol="1"/>
        """
        node = fromstring(src)
        loc = location.from_tree(node)
        expected = location(
            ref="A3:E14",
            firstHeaderRow=1,
            firstDataRow=2,
            firstDataCol=1,
        )
        assert loc == expected


class TestPivotTableStyle:

    def test_ctor(self, pivot_table_style):
        style = pivot_table_style(name="PivotStyleMedium4")
        xml = tostring(style.to_tree())
        expected = """
        <pivotTableStyleInfo
                name="PivotStyleMedium4"
                showRowHeaders="0"
                showColHeaders="0"
                showRowStripes="0"
                showColStripes="0"
                showLastColumn="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pivot_table_style):
        src = """
        <pivotTableStyleInfo
                name="PivotStyleMedium4"
                showRowHeaders="1"
                showColHeaders="1"
                showRowStripes="0"
                showColStripes="0"
                showLastColumn="1"/>
        """
        node = fromstring(src)
        style = pivot_table_style.from_tree(node)
        expected = pivot_table_style(
            name="PivotStyleMedium4",
            showRowHeaders=True,
            showColHeaders=True,
            showLastColumn=True,
        )
        assert style == expected

    def test_no_name(self, pivot_table_style):
        src = "<pivotTableStyleInfo/>"
        node = fromstring(src)
        style = pivot_table_style.from_tree(node)
        assert style == pivot_table_style()


@pytest.fixture
def dummy_pivot_table(table_definition, location):
    """
    Create a minimal pivot table
    """
    loc = location(ref="A3:E14", firstHeaderRow=1, firstDataRow=2, firstDataCol=1)
    defn = table_definition(
        name="PivotTable1",
        cacheId=68,
        applyWidthHeightFormats=True,
        dataCaption="Values",
        updatedVersion=4,
        createdVersion=4,
        gridDropZones=True,
        minRefreshableVersion=3,
        outlineData=True,
        useAutoFormatting=True,
        location=loc,
        indent=0,
        itemPrintTitles=True,
        outline=True,
    )
    return defn


class TestPivotTableDefinition:
    def test_ctor(self, dummy_pivot_table):
        defn = dummy_pivot_table
        xml = tostring(defn.to_tree())
        expected = """
        <pivotTableDefinition
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                name="PivotTable1"
                applyNumberFormats="0"
                applyBorderFormats="0"
                applyFontFormats="0"
                applyPatternFormats="0"
                applyAlignmentFormats="0"
                applyWidthHeightFormats="1"
                cacheId="68"
                asteriskTotals="0"
                chartFormat="0"
                colGrandTotals="1"
                compact="1"
                compactData="1"
                dataCaption="Values"
                dataOnRows="0"
                disableFieldList="0"
                editData="0"
                enableDrill="1"
                enableFieldProperties="1"
                enableWizard="1"
                fieldListSortAscending="0"
                fieldPrintTitles="0"
                updatedVersion="4"
                minRefreshableVersion="3"
                useAutoFormatting="1"
                itemPrintTitles="1"
                createdVersion="4"
                indent="0"
                outline="1"
                outlineData="1"
                gridDropZones="1"
                immersive="1"
                mdxSubqueries="0"
                mergeItem="0"
                multipleFieldFilters="0"
                pageOverThenDown="0"
                pageWrap="0"
                preserveFormatting="1"
                printDrill="0"
                published="0"
                rowGrandTotals="1"
                showCalcMbrs="1"
                showDataDropDown="1"
                showDataTips="1"
                showDrill="1"
                showDropZones="1"
                showEmptyCol="0"
                showEmptyRow="0"
                showError="0"
                showHeaders="1"
                showItems="1"
                showMemberPropertyTips="1"
                showMissing="1"
                showMultipleLabel="1"
                subtotalHiddenItems="0"
                visualTotals="1">
            <location
                    ref="A3:E14"
                    firstHeaderRow="1"
                    firstDataRow="2"
                    firstDataCol="1"/>
        </pivotTableDefinition>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, dummy_pivot_table, table_definition):
        src = """
        <pivotTableDefinition
                name="PivotTable1"
                applyNumberFormats="0"
                applyBorderFormats="0"
                applyFontFormats="0"
                applyPatternFormats="0"
                applyAlignmentFormats="0"
                applyWidthHeightFormats="1"
                cacheId="68"
                asteriskTotals="0"
                chartFormat="0"
                colGrandTotals="1"
                compact="1"
                compactData="1"
                dataCaption="Values"
                dataOnRows="0"
                disableFieldList="0"
                editData="0"
                enableDrill="1"
                enableFieldProperties="1"
                enableWizard="1"
                fieldListSortAscending="0"
                fieldPrintTitles="0"
                updatedVersion="4"
                minRefreshableVersion="3"
                useAutoFormatting="1"
                itemPrintTitles="1"
                createdVersion="4"
                indent="0"
                outline="1"
                outlineData="1"
                gridDropZones="1"
                immersive="1"
                mdxSubqueries="0"
                mergeItem="0"
                multipleFieldFilters="0"
                pageOverThenDown="0"
                pageWrap="0"
                preserveFormatting="1"
                printDrill="0"
                published="0"
                rowGrandTotals="1"
                showCalcMbrs="1"
                showDataDropDown="1"
                showDataTips="1"
                showDrill="1"
                showDropZones="1"
                showEmptyCol="0"
                showEmptyRow="0"
                showError="0"
                showHeaders="1"
                showItems="1"
                showMemberPropertyTips="1"
                showMissing="1"
                showMultipleLabel="1"
                subtotalHiddenItems="0"
                visualTotals="1">
           <location
                ref="A3:E14"
                firstHeaderRow="1"
                firstDataRow="2"
                firstDataCol="1"/>
        </pivotTableDefinition>
        """
        node = fromstring(src)
        defn = table_definition.from_tree(node)
        assert defn == dummy_pivot_table

    def test_write(self, dummy_pivot_table):
        out = io.BytesIO()
        archive = zipfile.ZipFile(out, "w")
        manifest = Manifest()
        defn = dummy_pivot_table
        defn._write(archive, manifest)
        assert archive.namelist() == [defn.path[1:]]
        assert manifest.find(defn.mime_type)

    @pytest.mark.xfail
    def test_formatted_fields(self, table_definition, datadir):
        datadir.chdir()
        with open("table_with_conditional.xml", "rb") as src:
            xml = src.read()
        tree = fromstring(xml)
        table = table_definition.from_tree(tree)
        assert table.formatted_fields() == {"Count": [2, 1], "Duration (minutes)": [3]}


class TestPageField:
    def test_ctor(self, page_field):
        pf = page_field(fld=64, hier=-1)
        xml = tostring(pf.to_tree())
        expected = '<pageField fld="64" hier="-1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, page_field):
        src = '<pageField fld="64" hier="-1"/>'
        node = fromstring(src)
        pf = page_field.from_tree(node)
        assert pf == page_field(fld=64, hier=-1)


class TestReference:
    def test_ctor(self, reference):
        ref = reference(field=4294967294, x=[0], selected=False)
        xml = tostring(ref.to_tree())
        expected = """
        <reference field="4294967294" selected="0">
            <x v="0"/>
        </reference>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, reference):
        src = """
        <reference field="4294967294" count="1" selected="0">
            <x v="0"/>
        </reference>
        """
        node = fromstring(src)
        ref = reference.from_tree(node)
        assert ref == reference(field=4294967294, x=[0], selected=False)


class TestPivotArea:
    def test_ctor(self, pivot_area):
        area = pivot_area(type="data", outline=False, fieldPosition=False)
        xml = tostring(area.to_tree())
        expected = """
        <pivotArea
                type="data"
                outline="0"
                fieldPosition="0"
                dataOnly="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pivot_area):
        src = """
        <pivotArea
                type="data"
                outline="0"
                fieldPosition="0"/>
        """
        node = fromstring(src)
        area = pivot_area.from_tree(node)
        assert area == pivot_area(type="data", outline=False, fieldPosition=False)


class TestChartFormat:
    def test_ctor(self, chart_format, pivot_area):
        area = pivot_area()
        fmt = chart_format(chart=0, format=12, series=1, pivotArea=area)
        xml = tostring(fmt.to_tree())
        expected = """
        <chartFormat chart="0" format="12" series="1">
            <pivotArea type="normal" outline="1" dataOnly="1"/>
        </chartFormat>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, chart_format, pivot_area):
        src = """
        <chartFormat chart="0" format="12" series="1">
            <pivotArea type="normal" outline="1" dataOnly="1"/>
        </chartFormat>
        """
        node = fromstring(src)
        fmt = chart_format.from_tree(node)
        area = pivot_area()
        assert fmt == chart_format(chart=0, format=12, series=1, pivotArea=area)


class TestPivotFilter:
    def test_ctor(self, pivot_filter, auto_filter):
        flt = pivot_filter(
            fld=0,
            id=6,
            evalOrder=-1,
            type="dateBetween",
            autoFilter=auto_filter,
        )
        xml = tostring(flt.to_tree())
        expected = """
        <filter fld="0" type="dateBetween" evalOrder="-1" id="6">
            <autoFilter ref="A1">
                <filterColumn colId="0" hiddenButton="0" showButton="1">
                    <customFilters and="1">
                        <customFilter operator="greaterThanOrEqual" val="1"/>
                        <customFilter operator="lessThanOrEqual" val="2"/>
                    </customFilters>
                </filterColumn>
            </autoFilter>
        </filter>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pivot_filter, auto_filter):
        src = """
        <filter fld="0" type="dateBetween" evalOrder="-1" id="6">
            <autoFilter ref="A1">
                <filterColumn colId="0">
                    <customFilters and="1">
                        <customFilter operator="greaterThanOrEqual" val="1"/>
                        <customFilter operator="lessThanOrEqual" val="2"/>
                    </customFilters>
                </filterColumn>
            </autoFilter>
        </filter>
        """
        node = fromstring(src)
        flt = pivot_filter.from_tree(node)
        expected = pivot_filter(
            fld=0,
            id=6,
            evalOrder=-1,
            type="dateBetween",
            autoFilter=auto_filter,
        )
        assert flt == expected


class TestFormat:
    def test_ctor(self, format_, pivot_area):
        area = pivot_area()
        fmt = format_(pivotArea=area)
        xml = tostring(fmt.to_tree())
        expected = """
        <format action="formatting">
            <pivotArea dataOnly="1" outline="1" type="normal"/>
        </format>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, format_, pivot_area):
        src = """
        <format action="blank">
            <pivotArea dataOnly="0" labelOnly="1" outline="0" fieldPosition="0"/>
        </format>
        """
        node = fromstring(src)
        fmt = format_.from_tree(node)
        area = pivot_area(
            outline=False,
            fieldPosition=False,
            labelOnly=True,
            dataOnly=False,
        )
        assert fmt == format_(action="blank", pivotArea=area)


class TestConditionalFormat:
    def test_ctor(self, conditional_format):
        fmt = conditional_format(priority=4)
        xml = tostring(fmt.to_tree())
        expected = """
        <conditionalFormat
                priority="4"
                scope="selection">
        </conditionalFormat>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, conditional_format):
        src = """
        <conditionalFormat
                priority="4"
                scope="selection">
        </conditionalFormat>
        """
        node = fromstring(src)
        fmt = conditional_format.from_tree(node)
        assert fmt == conditional_format(priority=4)


class TestConditionalFormatList:
    def test_ctor(self, conditional_format_list, conditional_format):
        fmt = conditional_format(priority=4)
        fmts = conditional_format_list(conditionalFormat=(fmt,))
        xml = tostring(fmts.to_tree())
        expected = """
        <conditionalFormats count="1">
            <conditionalFormat priority="4" scope="selection"/>
        </conditionalFormats>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, conditional_format_list, conditional_format):
        src = """
        <conditionalFormats>
            <conditionalFormat priority="4" scope="selection"/>
        </conditionalFormats>
        """
        node = fromstring(src)
        fmts = conditional_format_list.from_tree(node)
        assert fmts.conditionalFormat == [conditional_format(priority=4)]

    def test_by_priority(self, conditional_format_list):
        src = """
        <conditionalFormats count="3">
            <conditionalFormat scope="data" priority="3">
                <pivotAreas count="1">
                    <pivotArea outline="0" fieldPosition="0">
                        <references count="1">
                            <reference field="4294967294" count="1" selected="0">
                                <x v="0"/>
                            </reference>
                        </references>
                    </pivotArea>
                </pivotAreas>
            </conditionalFormat>
            <conditionalFormat scope="data" priority="2">
                <pivotAreas count="1">
                    <pivotArea outline="0" fieldPosition="0">
                        <references count="1">
                            <reference field="4294967294" count="1" selected="0">
                                <x v="1"/>
                            </reference>
                        </references>
                    </pivotArea>
                </pivotAreas>
            </conditionalFormat>
            <conditionalFormat scope="data" priority="1">
                <pivotAreas count="1">
                    <pivotArea outline="0" fieldPosition="0">
                        <references count="1">
                            <reference field="4294967294" count="1" selected="0">
                                <x v="1"/>
                            </reference>
                        </references>
                    </pivotArea>
                </pivotAreas>
            </conditionalFormat>
        </conditionalFormats>
        """
        node = fromstring(src)
        fmts = conditional_format_list.from_tree(node)
        prios = fmts.by_priority()
        assert list(prios.keys()) == [(0, 3), (1, 2), (1, 1)]

    @pytest.mark.xfail
    def test_dedupe(self, conditional_format_list):
        src = """
        <conditionalFormats count="3">
            <conditionalFormat scope="data" priority="3">
                <pivotAreas count="1">
                    <pivotArea outline="0" fieldPosition="0">
                        <references count="1">
                            <reference field="4294967294" count="1" selected="0">
                                <x v="0"/>
                            </reference>
                        </references>
                    </pivotArea>
                </pivotAreas>
            </conditionalFormat>
            <conditionalFormat scope="data" priority="2">
                <pivotAreas count="1">
                    <pivotArea outline="0" fieldPosition="0">
                        <references count="1">
                            <reference field="4294967294" count="1" selected="0">
                                <x v="1"/>
                            </reference>
                        </references>
                    </pivotArea>
                </pivotAreas>
            </conditionalFormat>
            <conditionalFormat scope="data" priority="1">
                <pivotAreas count="1">
                    <pivotArea outline="0" fieldPosition="0">
                        <references count="1">
                            <reference field="4294967294" count="1" selected="0">
                                <x v="1"/>
                            </reference>
                        </references>
                    </pivotArea>
                </pivotAreas>
            </conditionalFormat>
        </conditionalFormats>
        """
        node = fromstring(src)
        fmts = conditional_format_list.from_tree(node)
        fmts._dedupe()
        xml = tostring(fmts.to_tree())
        expected = """
        <conditionalFormats count="2">
            <conditionalFormat scope="data" priority="3">
                <pivotAreas>
                    <pivotArea
                            dataOnly="1"
                            outline="0"
                            fieldPosition="0"
                            type="normal">
                        <references count="1">
                            <reference field="4294967294" selected="0">
                                <x v="0"/>
                            </reference>
                        </references>
                    </pivotArea>
                </pivotAreas>
            </conditionalFormat>
            <conditionalFormat scope="data" priority="2">
                <pivotAreas>
                    <pivotArea
                            dataOnly="1"
                            outline="0"
                            fieldPosition="0"
                            type="normal">
                        <references count="1">
                            <reference field="4294967294" selected="0">
                                <x v="1"/>
                            </reference>
                        </references>
                    </pivotArea>
                </pivotAreas>
            </conditionalFormat>
        </conditionalFormats>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
