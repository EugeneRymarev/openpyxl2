# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def connection():
    from openpyxl.connection.connections import Connection

    return Connection


@pytest.fixture
def db_pr():
    from openpyxl.connection.connections import DbPr

    return DbPr


@pytest.fixture
def olap_pr():
    from openpyxl.connection.connections import OlapPr

    return OlapPr


@pytest.fixture
def text_field():
    from openpyxl.connection.connections import TextField

    return TextField


@pytest.fixture
def text_pr():
    from openpyxl.connection.connections import TextPr

    return TextPr


@pytest.fixture
def table_missing():
    from openpyxl.connection.connections import TableMissing

    return TableMissing


@pytest.fixture
def parameter():
    from openpyxl.connection.connections import Parameter

    return Parameter


@pytest.fixture
def web_pr():
    from openpyxl.connection.connections import WebPr

    return WebPr


@pytest.fixture
def tables():
    from openpyxl.connection.connections import Tables

    return Tables


@pytest.fixture
def connection_list():
    from openpyxl.connection.connections import ConnectionList

    return ConnectionList


class TestConnection:
    def test_ctor(self, connection):
        con = connection(id=3, refreshedVersion=8, background=True, keepAlive=True)
        xml = tostring(con.to_tree())
        expected = """
        <connection
                id="3"
                keepAlive="1"
                refreshedVersion="8"
                background="1"
                interval="0"
                reconnectionMethod="1"
                minRefreshableVersion="0"
                savePassword="0"
                new="0"
                deleted="0"
                onlyUseConnectionFile="0"
                refreshOnLoad="0"
                saveData="0"
                credentials="integrated"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, connection):
        src = """
        <connection
                id="2"
                keepAlive="1"
                name="Query - Table1"
                description="Connection to the 'Table2' query in the workbook."
                type="5"
                refreshedVersion="8"
                background="1"
                saveData="1"/>
        """
        node = fromstring(src)
        con = connection.from_tree(node)
        assert con.saveData == True
        assert con.name == "Query - Table1"
        assert con.type_descriptions[con.type] == "OLE DB-based source"

    @pytest.mark.parametrize("type, result", [(8, True), (102, False)])
    def test_known_type(self, connection, type, result):
        con = connection(id=3, refreshedVersion=8, background=True, type=type)
        assert con.is_known_connection is result


class TestDbPr:
    def test_ctor(self, db_pr):
        db_props = db_pr(
            connection="Data Model Connection",
            command="Model",
            commandType=True,
        )
        xml = tostring(db_props.to_tree())
        expected = """
        <dbPr connection="Data Model Connection"
              command="Model"
              commandType="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, db_pr):
        src = """
        <dbPr connection="Provider=Microsoft.Mashup.OleDb"
              command="SELECT * FROM [Table2]"/>
        """
        node = fromstring(src)
        db_props = db_pr.from_tree(node)
        assert db_props.command == "SELECT * FROM [Table2]"


class TestOlapPr:
    def test_ctor(self, olap_pr):
        olap_props = olap_pr(sendLocale=True, rowDrillCount=1000)
        xml = tostring(olap_props.to_tree())
        expected = """
        <olapPr local="0"
                sendLocale="1"
                rowDrillCount="1000"
                localRefresh="1"
                serverFill="1"
                serverNumberFormat="1"
                serverFont="1"
                serverFontColor="1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, olap_pr):
        src = '<olapPr serverFill="1" localRefresh="1" serverFontColor="1"/>'
        node = fromstring(src)
        olap_props = olap_pr.from_tree(node)
        expected = olap_pr(serverFill=True, localRefresh=True, serverFontColor=True)
        assert olap_props == expected


class TestTextField:
    def test_ctor(self, text_field):
        text = text_field(type="text")
        xml = tostring(text.to_tree())
        expected = '<textField type="text" position="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, text_field):
        from openpyxl.connection.connections import TextField

        src = '<textField type="DMY" position="2"/>'
        node = fromstring(src)
        text = TextField.from_tree(node)
        assert text == TextField(type="DMY", position=2)


class TestTextPr:
    def test_ctor(self, text_pr):
        text_props = text_pr(
            prompt=False,
            codePage=437,
            sourceFile="C:\\Desktop\\text data.txt",
            delimiter="|",
        )
        xml = tostring(text_props.to_tree())
        expected = """
        <textPr prompt="0"
                codePage="437"
                sourceFile="C:\\Desktop\\text data.txt"
                delimiter="|"
                fileType="win"
                firstRow="1"
                delimited="1"
                decimal="."
                thousands=","
                tab="1"
                qualifier="doubleQuote"
                space="0"
                comma="0"
                semicolon="0"
                consecutive="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, text_pr):
        from openpyxl.connection.connections import TextField

        src = """
        <textPr space="1"
                firstRow="13"
                sourceFile="C:\\Desktop\\text data.txt"
                delimiter="|">
            <textFields count="1">
                <textField/>
            </textFields>
        </textPr>
        """
        node = fromstring(src)
        text_props = text_pr.from_tree(node)
        expected = text_pr(
            space=True,
            firstRow=13,
            sourceFile="C:\\Desktop\\text data.txt",
            delimiter="|",
            textFields=[TextField()],
        )
        assert text_props == expected


class TestTableMissing:
    def test_ctor(self, table_missing):
        missing_table = table_missing()
        xml = tostring(missing_table.to_tree())
        expected = "<tableMissing/>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, table_missing):
        src = "<tableMissing/>"
        node = fromstring(src)
        missing_table = table_missing.from_tree(node)
        assert missing_table == table_missing()


class TestParameter:
    def test_ctor(self, parameter):
        param = parameter(name="TestName", boolean=True, sqlType=4)
        xml = tostring(param.to_tree())
        expected = """
        <parameter
                name="TestName"
                boolean="1"
                sqlType="4"
                parameterType="prompt"
                refreshOnChange="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, parameter):
        src = """
        <parameter
                name="user"
                refreshOnChange="1"
                parameterType="cell"
                cell="Sheet1!$C$1"/>
        """
        node = fromstring(src)
        param = parameter.from_tree(node)
        expected = parameter(
            name="user",
            refreshOnChange=True,
            parameterType="cell",
            cell="Sheet1!$C$1",
        )
        assert param == expected


class TestWebPr:
    def test_ctor(self, web_pr):
        web_props = web_pr(xml=True, firstRow=True, htmlTables=True)
        xml = tostring(web_props.to_tree())
        expected = """
        <webPr xml="1"
               firstRow="1"
               htmlTables="1"
               sourceData="0"
               parsePre="0"
               consecutive="0"
               xl97="0"
               textDates="0"
               xl2000="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, web_pr):
        src = """
        <webPr sourceData="1"
               parsePre="1"
               consecutive="1"
               url="http://ServerName/"
               htmlTables="1"/>
        """
        node = fromstring(src)
        web_props = web_pr.from_tree(node)
        expected = web_pr(
            sourceData=True,
            parsePre=True,
            consecutive=True,
            url="http://ServerName/",
            htmlTables=True,
        )
        assert web_props == expected


class TestTables:
    def test_ctor(self, tables):
        t = tables(x=3)
        xml = tostring(t.to_tree())
        expected = """
        <tables>
            <x v="3"/>
        </tables>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, tables):
        src = """
        <tables count="1">
            <s v="test"/>
        </tables>
        """
        node = fromstring(src)
        t = tables.from_tree(node)
        assert t == tables(s="test", count=1)


class TestConnectionList:
    def test_ctor(self, connection_list, connection):
        connections = connection_list(connection=[connection(id=1, refreshedVersion=4)])
        xml = tostring(connections.to_tree())
        expected = """
        <connections
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <connection
                    id="1"
                    refreshedVersion="4"
                    background="0"
                    credentials="integrated"
                    deleted="0"
                    interval="0"
                    keepAlive="0"
                    minRefreshableVersion="0"
                    new="0"
                    onlyUseConnectionFile="0"
                    reconnectionMethod="1"
                    refreshOnLoad="0"
                    saveData="0"
                    savePassword="0"/>
        </connections>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, connection_list, datadir, recwarn):
        datadir.chdir()
        with open("connections.xml", "rb") as src:
            node = fromstring(src.read())
        connections = connection_list.from_tree(node)
        assert len(connections.connection) == 2
        assert connections.connection[0].id == 2
        assert connections.connection[1].saveData == True
        w = recwarn.pop()
        assert issubclass(w.category, UserWarning)

    def test_access_by_id(self, connection_list, datadir):
        datadir.chdir()
        with open("connections.xml", "rb") as src:
            node = fromstring(src.read())
        connections = connection_list.from_tree(node)
        conn = connections[2]
        assert conn.name == "ThisWorkbookDataModel"

    def test_connection_not_found(self, connection_list, datadir):
        datadir.chdir()
        with open("connections.xml", "rb") as src:
            node = fromstring(src.read())
        connections = connection_list.from_tree(node)
        with pytest.raises(IndexError):
            connections[7]

    def test_caches(self, connection_list, connection):
        from openpyxl.pivot.cache import CacheDefinition
        from openpyxl.pivot.cache import CacheFieldList
        from openpyxl.pivot.cache import CacheSource

        con1 = connection(id=5, refreshedVersion=4)
        con2 = connection(id=7, refreshedVersion=4)
        con2._cache = CacheDefinition(
            cacheSource=CacheSource(type="external"),
            cacheFields=CacheFieldList(),
        )
        cons = connection_list(connection=[con1, con2])
        assert cons.caches == [con2._cache]
        assert con2.id == 2
        assert con2._cache.cacheSource.connectionId == 2
