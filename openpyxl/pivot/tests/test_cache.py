# Copyright (c) 2010-2025 openpyxl
import datetime
import io
import zipfile

import pytest
from openpyxl.packaging.manifest import Manifest
from openpyxl.pivot.record import Text
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def cache_field():
    from openpyxl.pivot.cache import CacheField

    return CacheField


@pytest.fixture
def cache_field_list():
    from openpyxl.pivot.cache import CacheFieldList

    return CacheFieldList


@pytest.fixture
def shared_items():
    from openpyxl.pivot.cache import SharedItems

    return SharedItems


@pytest.fixture
def worksheet_source():
    from openpyxl.pivot.cache import WorksheetSource

    return WorksheetSource


@pytest.fixture
def cache_source():
    from openpyxl.pivot.cache import CacheSource

    return CacheSource


@pytest.fixture
def query():
    from openpyxl.pivot.cache import Query

    return Query


@pytest.fixture
def tuple_cache():
    from openpyxl.pivot.cache import TupleCache

    return TupleCache


@pytest.fixture
def pcdsdtc_entries():
    from openpyxl.pivot.cache import PCDSDTCEntries

    return PCDSDTCEntries


@pytest.fixture
def cache_definition():
    from openpyxl.pivot.cache import CacheDefinition

    return CacheDefinition


@pytest.fixture
def cache_hierarchy():
    from openpyxl.pivot.cache import CacheHierarchy

    return CacheHierarchy


@pytest.fixture
def measure_dimension_map():
    from openpyxl.pivot.cache import MeasureDimensionMap

    return MeasureDimensionMap


@pytest.fixture
def measure_group():
    from openpyxl.pivot.cache import MeasureGroup

    return MeasureGroup


@pytest.fixture
def pivot_dimension():
    from openpyxl.pivot.cache import PivotDimension

    return PivotDimension


@pytest.fixture
def calculated_member():
    from openpyxl.pivot.cache import CalculatedMember

    return CalculatedMember


@pytest.fixture
def calculated_item():
    from openpyxl.pivot.cache import CalculatedItem

    return CalculatedItem


@pytest.fixture
def server_format():
    from openpyxl.pivot.cache import ServerFormat

    return ServerFormat


@pytest.fixture
def olap_set():
    from openpyxl.pivot.cache import OLAPSet

    return OLAPSet


@pytest.fixture
def olapkpi():
    from openpyxl.pivot.cache import OLAPKPI

    return OLAPKPI


@pytest.fixture
def group_member():
    from openpyxl.pivot.cache import GroupMember

    return GroupMember


@pytest.fixture
def level_group():
    from openpyxl.pivot.cache import LevelGroup

    return LevelGroup


@pytest.fixture
def group_level():
    from openpyxl.pivot.cache import GroupLevel

    return GroupLevel


@pytest.fixture
def field_usage():
    from openpyxl.pivot.cache import FieldUsage

    return FieldUsage


@pytest.fixture
def group_items():
    from openpyxl.pivot.cache import GroupItems

    return GroupItems


@pytest.fixture
def range_pr():
    from openpyxl.pivot.cache import RangePr

    return RangePr


@pytest.fixture
def field_group():
    from openpyxl.pivot.cache import FieldGroup

    return FieldGroup


@pytest.fixture
def range_set():
    from openpyxl.pivot.cache import RangeSet

    return RangeSet


@pytest.fixture
def page_item():
    from openpyxl.pivot.cache import PageItem

    return PageItem


@pytest.fixture
def consolidation():
    from openpyxl.pivot.cache import Consolidation

    return Consolidation


@pytest.fixture
def cache_definition_collection():
    from openpyxl.pivot.cache import CacheDefinitionCollection

    return CacheDefinitionCollection


class TestCacheField:
    def test_ctor(self, cache_field):
        field = cache_field(name="ID")
        xml = tostring(field.to_tree())
        expected = """
        <cacheField
                databaseField="1"
                hierarchy="0"
                level="0"
                name="ID"
                sqlType="0"
                uniqueList="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, cache_field):
        src = '<cacheField name="ID"/>'
        node = fromstring(src)
        field = cache_field.from_tree(node)
        assert field == cache_field(name="ID")


class TestCacheFieldList:
    def test_from_xml(self, cache_field_list):
        src = """
        <cacheFields count="6">
            <cacheField name="ID" numFmtId="0">
            </cacheField>
        </cacheFields>
        """
        node = fromstring(src)
        fields = cache_field_list.from_tree(node)
        assert len(fields.cacheField) == 1


class TestSharedItems:
    def test_ctor(self, shared_items):
        s = [Text(v="Stanford"), Text(v="Cal"), Text(v="UCLA")]
        items = shared_items(_fields=s)
        xml = tostring(items.to_tree())
        expected = """
        <sharedItems count="3">
            <s v="Stanford"/>
            <s v="Cal"/>
            <s v="UCLA"/>
        </sharedItems>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, shared_items):
        src = """
        <sharedItems count="3">
            <s v="Stanford"></s>
            <s v="Cal"></s>
            <s v="UCLA"></s>
        </sharedItems>
        """
        node = fromstring(src)
        items = shared_items.from_tree(node)
        s = [Text(v="Stanford"), Text(v="Cal"), Text(v="UCLA")]
        assert items == shared_items(_fields=s)


class TestWorksheetSource:
    def test_ctor(self, worksheet_source):
        ws = worksheet_source(name="mydata")
        xml = tostring(ws.to_tree())
        expected = '<worksheetSource name="mydata"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, worksheet_source):
        src = '<worksheetSource name="mydata"/>'
        node = fromstring(src)
        ws = worksheet_source.from_tree(node)
        assert ws == worksheet_source(name="mydata")


class TestCacheSource:
    def test_ctor(self, cache_source, worksheet_source):
        ws = worksheet_source(name="mydata")
        source = cache_source(type="worksheet", worksheetSource=ws)
        xml = tostring(source.to_tree())
        expected = """
        <cacheSource type="worksheet">
            <worksheetSource name="mydata"/>
        </cacheSource>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, cache_source, worksheet_source):
        src = """
        <cacheSource type="worksheet">
            <worksheetSource name="mydata"/>
        </cacheSource>
        """
        node = fromstring(src)
        source = cache_source.from_tree(node)
        ws = worksheet_source(name="mydata")
        assert source == cache_source(type="worksheet", worksheetSource=ws)


class TestQuery:
    def test_ctor(self, query):
        q = query(mdx="[Description]")
        xml = tostring(q.to_tree())
        expected = '<query mdx="[Description]"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, query):
        src = """
        <query mdx="[Plant]">
            <tpls c="1">
                <tpl hier="1" item="4294967295"/>
            </tpls>
        </query>
        """
        node = fromstring(src)
        q = query.from_tree(node)
        assert q.mdx == "[Plant]"
        assert q.tpls.c == 1


class TestTupleCache:
    def test_ctor(self, tuple_cache, query):
        c = tuple_cache(queryCache=[query(mdx="[Plant]")])
        xml = tostring(c.to_tree())
        expected = """
        <tupleCache>
            <queryCache count="1">
                <query mdx="[Plant]"/>
            </queryCache>
        </tupleCache>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, tuple_cache):
        src = """
        <tupleCache>
            <queryCache count="1">
                <query mdx="[Description]">
                    <tpls c="1">
                        <tpl hier="1" item="4294967295"/>
                    </tpls>
                </query>
            </queryCache>
        </tupleCache>
        """
        node = fromstring(src)
        c = tuple_cache.from_tree(node)
        assert len(c.queryCache) == 1


class TestPCDSDTCEntries:
    def test_ctor(self, pcdsdtc_entries):
        from openpyxl.pivot.fields import Number

        entries = pcdsdtc_entries(n=Number(v=1))
        xml = tostring(entries.to_tree())
        expected = """
        <entries>
            <n v="1"/>
        </entries>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pcdsdtc_entries):
        from openpyxl.pivot.fields import Text

        src = """
        <entries>
            <s v="test"/>
        </entries>
        """
        node = fromstring(src)
        entries = pcdsdtc_entries.from_tree(node)
        assert entries == pcdsdtc_entries(s=Text(v="test"))


@pytest.fixture
def dummy_cache(
    cache_definition,
    worksheet_source,
    cache_source,
    cache_field,
    cache_field_list,
    tuple_cache,
):
    ws = worksheet_source(name="Sheet1")
    source = cache_source(type="worksheet", worksheetSource=ws)
    fields = cache_field_list(cacheField=[cache_field(name="field1")])
    tuples = tuple_cache()
    c = cache_definition(
        cacheSource=source,
        cacheFields=fields,
        tupleCache=tuples,
        saveData=True,
    )
    return c


class TestPivotCacheDefinition:
    def test_read(self, cache_definition, datadir):
        datadir.chdir()
        with open("pivotCacheDefinition.xml", "rb") as src:
            xml = fromstring(src.read())
        c = cache_definition.from_tree(xml)
        assert c.recordCount == 17
        assert c.cacheFields.count == 6

    def test_read_tuple_cache(self, cache_definition, datadir):
        # Different sample with use of tupleCache
        datadir.chdir()
        with open("pivotCacheDefinitionTupleCache.xml", "rb") as src:
            xml = fromstring(src.read())
        c = cache_definition.from_tree(xml)
        assert c.recordCount == 0
        assert c.tupleCache.entries.count == 1
        assert c.has_olap_cache == True

    def test_to_tree(self, dummy_cache):
        c = dummy_cache
        expected = """
        <pivotCacheDefinition
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                tupleCache="1"
                saveData="1">
            <cacheSource type="worksheet">
                <worksheetSource name="Sheet1"/>
            </cacheSource>
            <cacheFields count="1">
                <cacheField
                        databaseField="1"
                        hierarchy="0"
                        level="0"
                        name="field1"
                        sqlType="0"
                        uniqueList="1"/>
            </cacheFields>
            <tupleCache/>
        </pivotCacheDefinition>
        """
        xml = tostring(c.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_path(self, dummy_cache):
        assert dummy_cache.path == "/xl/pivotCache/pivotCacheDefinition1.xml"

    def test_write(self, dummy_cache):
        out = io.BytesIO()
        archive = zipfile.ZipFile(out, mode="w")
        manifest = Manifest()
        xml = tostring(dummy_cache.to_tree())
        dummy_cache._write(archive, manifest)
        assert archive.namelist() == [dummy_cache.path[1:]]
        assert manifest.find(dummy_cache.mime_type)


class TestCacheHierarchy:
    def test_ctor(self, cache_hierarchy):
        ch = cache_hierarchy(
            uniqueName="[Interval].[Date]",
            caption="Date",
            attribute=True,
            time=True,
            defaultMemberUniqueName="[Interval].[Date].[All]",
            allUniqueName="[Interval].[Date].[All]",
            dimensionUniqueName="[Interval]",
            memberValueDatatype=7,
            count=0,
        )
        xml = tostring(ch.to_tree())
        expected = """
        <cacheHierarchy
                uniqueName="[Interval].[Date]"
                caption="Date"
                attribute="1"
                time="1"
                defaultMemberUniqueName="[Interval].[Date].[All]"
                allUniqueName="[Interval].[Date].[All]"
                dimensionUniqueName="[Interval]"
                count="0"
                memberValueDatatype="7"
                hidden="0"
                iconSet="0"
                keyAttribute="0"
                measure="0"
                measures="0"
                oneField="0"
                set="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, cache_hierarchy):
        src = """
        <cacheHierarchy
                uniqueName="[Interval].[Date]"
                caption="Date"
                attribute="1"
                time="1"
                defaultMemberUniqueName="[Interval].[Date].[All]"
                allUniqueName="[Interval].[Date].[All]"
                dimensionUniqueName="[Interval]"
                displayFolder=""
                count="0"
                memberValueDatatype="7"
                unbalanced="0"/>
        """
        node = fromstring(src)
        ch = cache_hierarchy.from_tree(node)
        expected = cache_hierarchy(
            uniqueName="[Interval].[Date]",
            caption="Date",
            attribute=True,
            time=True,
            defaultMemberUniqueName="[Interval].[Date].[All]",
            allUniqueName="[Interval].[Date].[All]",
            dimensionUniqueName="[Interval]",
            memberValueDatatype=7,
            count=0,
            unbalanced=False,
            displayFolder="",
        )
        assert ch == expected


class TestMeasureDimensionMap:
    def test_ctor(self, measure_dimension_map):
        mdm = measure_dimension_map()
        xml = tostring(mdm.to_tree())
        expected = "<map/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, measure_dimension_map):
        src = "<map/>"
        node = fromstring(src)
        mdm = measure_dimension_map.from_tree(node)
        assert mdm == measure_dimension_map()


class TestMeasureGroup:
    def test_ctor(self, measure_group):
        mg = measure_group(name="a", caption="caption")
        xml = tostring(mg.to_tree())
        expected = '<measureGroup name="a" caption="caption"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, measure_group):
        src = '<measureGroup name="name" caption="caption"/>'
        node = fromstring(src)
        mg = measure_group.from_tree(node)
        assert mg == measure_group(name="name", caption="caption")


class TestPivotDimension:
    def test_ctor(self, pivot_dimension):
        pd = pivot_dimension(
            measure=True,
            name="name",
            uniqueName="name",
            caption="caption",
        )
        xml = tostring(pd.to_tree())
        expected = """
        <dimension
                caption="caption"
                measure="1"
                name="name"
                uniqueName="name"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, pivot_dimension):
        src = """
        <dimension
                caption="caption"
                measure="1"
                name="name"
                uniqueName="name"/>
        """
        node = fromstring(src)
        pd = pivot_dimension.from_tree(node)
        expected = pivot_dimension(
            measure=True,
            name="name",
            uniqueName="name",
            caption="caption",
        )
        assert pd == expected


class TestCalculatedMember:
    def test_ctor(self, calculated_member):
        cm = calculated_member(
            name="name",
            mdx="mdx",
            memberName="member",
            hierarchy="yes",
            parent="parent",
            solveOrder=1,
            set=True,
        )
        xml = tostring(cm.to_tree())
        expected = """
        <calculatedMember
                hierarchy="yes"
                mdx="mdx"
                memberName="member"
                name="name"
                parent="parent"
                set="1"
                solveOrder="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, calculated_member):
        src = '<calculatedMember mdx="mdx" name="name" set="1" solveOrder="1"/>'
        node = fromstring(src)
        cm = calculated_member.from_tree(node)
        assert cm == calculated_member(name="name", mdx="mdx", solveOrder=1, set=True)


class TestCalculatedItem:
    def test_ctor(self, calculated_item):
        from openpyxl.pivot.cache import PivotArea

        item = calculated_item(formula="SUM(15)", pivotArea=PivotArea(cacheIndex=1))
        xml = tostring(item.to_tree())
        expected = """
        <calculatedItem formula="SUM(15)">
            <pivotArea type="normal" dataOnly="1" cacheIndex="1" outline="1"/>
        </calculatedItem>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, calculated_item, datadir):
        datadir.chdir()
        with open("calculatedItem.xml", "rb") as src:
            xml = fromstring(src.read())
        item = calculated_item.from_tree(xml)
        assert item.formula == "SUM(15)"
        assert item.pivotArea.cacheIndex == 1


class TestServerFormat:
    def test_ctor(self, server_format):
        sf = server_format(culture="x", format="y")
        xml = tostring(sf.to_tree())
        expected = '<serverFormat culture="x" format="y"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, server_format):
        src = '<serverFormat  culture="x" format="y"/>'
        node = fromstring(src)
        sf = server_format.from_tree(node)
        assert sf == server_format(culture="x", format="y")


class TestOLAPSet:
    def test_ctor(self, olap_set):
        olap_set_ = olap_set(
            count=1,
            maxRank=2,
            setDefinition="TestSet",
            queryFailed=False,
        )
        xml = tostring(olap_set_.to_tree())
        expected = """
        <set count="1"
             maxRank="2"
             setDefinition="TestSet"
             queryFailed="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, olap_set):
        src = """
        <set count="3"
             maxRank="5"
             setDefinition="Other"
             queryFailed="1"/>
        """
        node = fromstring(src)
        olap_set_ = olap_set.from_tree(node)
        expected = olap_set(
            count=3,
            maxRank=5,
            setDefinition="Other",
            queryFailed=True,
        )
        assert olap_set_ == expected


class TestOLAPKPI:
    def test_ctor(self, olapkpi):
        kpi = olapkpi(
            uniqueName="TestKPI",
            caption="TestCaption",
            displayFolder="Folder\\Display",
            measureGroup="TestMeasure",
            parent="TestParent",
            value="TestValue",
            goal="[Measures].[Goals]",
            status="TestStatus",
            trend="TestTrend",
            weight="",
            time="TestTime",
        )
        xml = tostring(kpi.to_tree())
        expected = """
        <kpi uniqueName="TestKPI"
             caption="TestCaption"
             displayFolder="Folder\\Display"
             measureGroup="TestMeasure"
             parent="TestParent"
             value="TestValue"
             goal="[Measures].[Goals]"
             status="TestStatus"
             trend="TestTrend"
             weight=""
             time="TestTime"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, olapkpi):
        xml = """
        <kpi uniqueName="Growth in Customer Base"
             caption="Growth in Customer Base"
             displayFolder="Customer Perspective\\Expand Customer Base"
             measureGroup="Internet Sales"
             value="[Measures].[Growth in Customer Base]"
             goal="[Measures].[Growth in Customer Base Goal]"
             status="[Measures].[Growth in Customer Base Status]"
             trend="[Measures].[Growth in Customer Base Trend]"/>
        """
        node = fromstring(xml)
        kpi = olapkpi.from_tree(node)
        assert kpi.trend == "[Measures].[Growth in Customer Base Trend]"


class TestGroupMember:
    def test_ctor(self, group_member):
        member = group_member(uniqueName="[Product].[Product Categories].[Category]")
        xml = tostring(member.to_tree())
        expected = """
        <groupMember
                group="0"
                uniqueName="[Product].[Product Categories].[Category]"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, group_member):
        xml = """
        <groupMember
                uniqueName="[Product].[Product Categories]"
                group="1"/>
        """
        node = fromstring(xml)
        member = group_member.from_tree(node)
        assert member.group is True
        assert member.uniqueName == "[Product].[Product Categories]"


class TestLevelGroup:
    def test_ctor(self, level_group):
        level = level_group(
            name="CategoryXl_Grp_1",
            uniqueName="[Product].[Product Categories]",
            caption="Group1",
            uniqueParent="[Product].[Product Categories].[All Products]",
            id=1,
        )
        xml = tostring(level.to_tree())
        expected = """
        <group name="CategoryXl_Grp_1"
               uniqueName="[Product].[Product Categories]"
               caption="Group1"
               uniqueParent="[Product].[Product Categories].[All Products]"
               id="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, level_group):
        xml = """
        <group name="Cat1"
               uniqueName="[Product]"
               caption="CatGroup1"
               uniqueParent="[Product].[Product Categories].[All Products]"
               id="4"/>
        """
        node = fromstring(xml)
        level = level_group.from_tree(node)
        assert level.name == "Cat1"
        assert level.id == 4


class TestGroupLevel:
    def test_ctor(self, group_level):
        group = group_level(
            uniqueName="TestGroup",
            caption="TestCaption",
            user=True,
            customRollUp=True,
        )
        xml = tostring(group.to_tree())
        expected = """
        <groupLevel
                uniqueName="TestGroup"
                caption="TestCaption"
                user="1"
                customRollUp="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, group_level):
        xml = """
        <groupLevel
                uniqueName="[Product].[Product Categories].[Category]"
                caption="Category">
            <groups count="1">
                <group name="CategoryXl_Grp_1"
                       uniqueName="[Product].[ProductCategories].[Product Categories1].[GROUPMEMBER.[CategoryXl_Grp_1]].[Product]].[Product Categories]].[All Products]]]"
                       caption="Group1"
                       uniqueParent="[Product].[Product Categories].[All Products]"
                       id="1">
                </group>
            </groups>
        </groupLevel>
        """
        node = fromstring(xml)
        group = group_level.from_tree(node)
        assert group.caption == "Category"
        assert len(group.groups) == 1


class TestFieldUsage:
    def test_ctor(self, field_usage):
        field = field_usage(x=5)
        xml = tostring(field.to_tree())
        expected = '<fieldUsage x="5"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, field_usage):
        xml = '<fieldUsage x="-1"/>'
        node = fromstring(xml)
        field = field_usage.from_tree(node)
        assert field.x == -1


class TestGroupItems:
    def test_ctor(self, group_items):
        from openpyxl.pivot.record import Text

        group = group_items(s=[Text(v="1-2"), Text(v="3-4")])
        xml = tostring(group.to_tree())
        expected = """
        <groupItems count="2">
            <s v="1-2"/>
            <s v="3-4"/>
        </groupItems>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, group_items):
        xml = """
        <groupItems count="4">
            <s v="&lt;1"/>
            <s v="1-2"/>
            <s v="3-4"/>
            <s v="&gt;5"/>
        </groupItems>
        """
        node = fromstring(xml)
        group = group_items.from_tree(node)
        assert group.s[0].v == "<1"
        assert group.count == 4


class TestRangePr:
    def test_ctor(self, range_pr):
        r = range_pr(startNum=1, endNum=4, groupInterval=2)
        xml = tostring(r.to_tree())
        expected = """
        <rangePr
                autoStart="1"
                autoEnd="1"
                groupBy="range"
                startNum="1"
                endNum="4"
                groupInterval="2"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, range_pr):
        xml = """
        <rangePr
                groupBy="months"
                startDate="2002-01-01T00:00:00"
                endDate="2006-05-06T00:00:00"/>
        """
        node = fromstring(xml)
        r = range_pr.from_tree(node)
        assert r.groupBy == "months"
        assert r.startDate == datetime.datetime(year=2002, month=1, day=1)


class TestFieldGroup:
    def test_ctor(self, field_group):
        field = field_group(par=4, base=3)
        xml = tostring(field.to_tree())
        expected = '<fieldGroup par="4" base="3"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, field_group):
        xml = """
        <fieldGroup base="0">
            <rangePr startNum="1" endNum="4" groupInterval="2"/>
            <groupItems count="4">
                <s v="1-2"/>
                <s v="3-4"/>
            </groupItems>
        </fieldGroup>
        """
        node = fromstring(xml)
        field = field_group.from_tree(node)
        assert field.base == 0
        assert len(field.groupItems.s) == 2


class TestRangeSet:
    def test_ctor(self, range_set):
        rangeset = range_set(i1=1, i2=1, ref="A1:B3", sheet="Sheet2")
        xml = tostring(rangeset.to_tree())
        expected = '<rangeSet i1="1" i2="1" ref="A1:B3" sheet="Sheet2"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, range_set):
        xml = '<rangeSet i1="4" i2="4" ref="A1:B3" sheet="Sheet5"/>'
        node = fromstring(xml)
        rangeset = range_set.from_tree(node)
        assert rangeset.i1 == 4
        assert rangeset.ref == "A1:B3"


class TestPageItem:
    def test_ctor(self, page_item):
        page = page_item(name="TestPage")
        xml = tostring(page.to_tree())
        expected = '<pageItem name="TestPage"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, page_item):
        xml = '<pageItem name="NewPage"/>'
        node = fromstring(xml)
        p = page_item.from_tree(node)
        assert p.name == "NewPage"


class TestConsolidation:
    def test_ctor(self, consolidation):
        from openpyxl.pivot.cache import RangeSet

        cons = consolidation(autoPage=True, rangeSets=[RangeSet(i1=1, ref="A1:B3")])
        xml = tostring(cons.to_tree())
        expected = """
        <consolidation autoPage="1">
            <rangeSets count="1">
                <rangeSet i1="1" ref="A1:B3"/>
            </rangeSets>
        </consolidation>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, consolidation):
        xml = """
        <consolidation>
            <pages count="1">
                <pageItem name="TestName"/>
            </pages>
            <rangeSets count="1">
                <rangeSet i1="1" ref="A1:B3"/>
            </rangeSets>
        </consolidation>
        """
        node = fromstring(xml)
        cons = consolidation.from_tree(node)
        assert cons.autoPage is None
        assert len(cons.pages) == 1
        assert len(cons.rangeSets) == 1


class TestCacheDefinitionCollection:
    def test_sort(
        self,
        cache_source,
        cache_definition,
        cache_field_list,
        cache_definition_collection,
    ):
        caches = cache_definition_collection()
        for possible in ["worksheet", "external", "consolidation", "scenario"]:
            source = cache_source(type=possible)
            c = cache_definition(cacheSource=source, cacheFields=cache_field_list())
            caches.append(c)
        entries = caches.by_type()
        expected = ["consolidation", "external", "scenario", "worksheet"]
        assert [source for source, group in entries] == expected
