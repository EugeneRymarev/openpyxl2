# Copyright (c) 2010-2025 openpyxl
import datetime

import pytest
from openpyxl.packaging.custom import BoolProperty
from openpyxl.packaging.custom import DateTimeProperty
from openpyxl.packaging.custom import FloatProperty
from openpyxl.packaging.custom import IntProperty
from openpyxl.packaging.custom import LinkProperty
from openpyxl.packaging.custom import StringProperty
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def custom_document_property():
    from openpyxl.packaging.custom import _CustomDocumentProperty

    return _CustomDocumentProperty


@pytest.fixture
def custom_document_property_list():
    from openpyxl.packaging.custom import _CustomDocumentPropertyList

    return _CustomDocumentPropertyList


@pytest.fixture
def custom_property_list():
    from openpyxl.packaging.custom import CustomPropertyList

    return CustomPropertyList


class TestCustomDocumentProperty:
    def test_ctor(self, custom_document_property):
        prop = custom_document_property(name="PropName9", bool=True)
        assert prop.type == "bool"
        assert prop.bool is True
        expected = """
        <property
                name="PropName9"
                pid="0"
                fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:bool
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                1
            </vt:bool>
        </property>
        """
        xml = tostring(prop.to_tree())
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, custom_document_property):
        src = """
        <property
                name="PropName1"
                pid="0"
                fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
            <vt:filetime
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                    >2020-08-24T20:19:22Z</vt:filetime>
        </property>
        """
        node = fromstring(src)
        prop = custom_document_property.from_tree(node)
        dt = datetime.datetime(2020, 8, 24, hour=20, minute=19, second=22)
        assert prop.filetime == dt and prop.name == "PropName1"
        src = """
        <property
                name="PropName4"
                pid="0"
                fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                linkTarget="ExampleName">
            <vt:lpwstr
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>
        </property>
        """
        node = fromstring(src)
        prop = custom_document_property.from_tree(node)
        assert prop.linkTarget == "ExampleName" and prop.name == "PropName4"

    def test_read_empty(self, custom_document_property):
        src = """
        <property
                fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                pid="7"
                name="TemplateUrl">
            <vt:lpwstr
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>
        </property>
        """
        node = fromstring(src)
        prop = custom_document_property.from_tree(node)
        assert prop.type == "lpwstr"

    def test_write_empty(self, custom_document_property):
        prop = custom_document_property(name="A name", lpwstr=None)
        node = prop.to_tree()
        expected = """
        <property
                fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                pid="0"
                name="A name">
            <vt:lpwstr
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>
        </property>
        """
        xml = tostring(node)
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestCustomDocumentPropertyList:
    def test_ctor(self, custom_document_property_list, custom_document_property):
        prop1 = custom_document_property(
            name="PropName1",
            filetime=datetime.datetime(2020, 8, 24, 20, 19, 22),
        )
        prop2 = custom_document_property(
            name="PropName2",
            linkTarget="ExampleName",
            lpwstr="",
        )
        prop3 = custom_document_property(name="PropName3", r8=2.5)
        props = custom_document_property_list(property=[prop1, prop2, prop3])
        xml = tostring(props.to_tree())
        expected = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:filetime>2020-08-24T20:19:22Z</vt:filetime>
            </property>
            <property
                    name="PropName2"
                    pid="3"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                    linkTarget="ExampleName">
                <vt:lpwstr/>
            </property>
            <property
                    name="PropName3"
                    pid="4"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:r8>2.5</vt:r8>
            </property>
        </Properties>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, custom_document_property_list, custom_document_property):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:filetime>2020-08-24T20:19:22Z</vt:filetime>
            </property>
            <property
                    name="PropName2"
                    pid="3"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:r8>2.5</vt:r8>
            </property>
            <property
                    name="PropName3"
                    pid="4"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:bool>true</vt:bool>
            </property>
        </Properties>
        """
        node = fromstring(src)
        props = custom_document_property_list.from_tree(node)
        expected = [
            custom_document_property(
                name="PropName1",
                filetime=datetime.datetime(2020, 8, 24, 20, 19, 22),
                pid=2,
            ),
            custom_document_property(name="PropName2", r8=2.5, pid=3),
            custom_document_property(name="PropName3", bool=True, pid=4),
        ]
        assert props.customProps == expected

    def test_len(self, custom_document_property_list):
        props = custom_document_property_list()
        assert len(props) == 0


class TestTypedPropertyList:
    def test_ctor(self, custom_property_list):
        prop_list = custom_property_list()
        assert prop_list.props == []

    def test_len(self, custom_property_list):
        prop_list = custom_property_list()
        assert len(prop_list) == 0

    def test_repr(self, custom_property_list):
        prop_list = custom_property_list()
        prop_list.append(StringProperty(name="PropName1", value="Something"))
        expected = (
            "CustomPropertyList containing [StringProperty,"
            " name=PropName1, value=Something]"
        )
        assert repr(prop_list) == expected

    def test_get_item(self, custom_property_list):
        prop_list = custom_property_list()
        prop = StringProperty(name="PropName1", value="Something")
        prop_list.append(prop)
        assert prop_list["PropName1"] == prop

    def test_get_item_missing(self, custom_property_list):
        prop_list = custom_property_list()
        with pytest.raises(KeyError):
            _ = prop_list["PropName1"]

    def test_delete(self, custom_property_list):
        prop_list = custom_property_list()
        prop_list.append(StringProperty(name="a prop", value="Something"))
        del prop_list["a prop"]
        assert prop_list.props == []

    def test_delete_missing(self, custom_property_list):
        prop_list = custom_property_list()
        with pytest.raises(KeyError):
            del prop_list["PropName1"]

    def test_string(self, custom_property_list):
        prop = StringProperty(name="PropName1", value="Something")
        prop_list = custom_property_list()
        prop_list.append(prop)
        tree = prop_list.to_tree()
        expected = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:lpwstr>Something</vt:lpwstr>
            </property>
        </Properties>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_int(self, custom_property_list):
        prop = IntProperty(name="PropName1", value=15)
        prop_list = custom_property_list()
        prop_list.append(prop)
        tree = prop_list.to_tree()
        expected = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:i4>15</vt:i4>
            </property>
        </Properties>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_float(self, custom_property_list):
        prop = IntProperty(name="PropName1", value=15)
        prop_list = custom_property_list()
        prop_list.append(prop)
        tree = prop_list.to_tree()
        expected = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:i4>15</vt:i4>
            </property>
        </Properties>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_bool(self, custom_property_list):
        prop = BoolProperty(name="PropName1", value=False)
        prop_list = custom_property_list()
        prop_list.append(prop)
        tree = prop_list.to_tree()
        expected = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:bool>0</vt:bool>
            </property>
        </Properties>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_datetime(self, custom_property_list):
        dt = datetime.datetime(2022, 5, 31, 12, 55, 13)
        prop = DateTimeProperty(name="PropName1", value=dt)
        prop_list = custom_property_list()
        prop_list.append(prop)
        tree = prop_list.to_tree()
        expected = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:filetime>2022-05-31T12:55:13Z</vt:filetime>
            </property>
        </Properties>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_link(self, custom_property_list):
        prop = LinkProperty(name="PropName1", value="A link")
        prop_list = custom_property_list()
        prop_list.append(prop)
        tree = prop_list.to_tree()
        expected = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                    linkTarget="A link">
                <vt:lpwstr/>
            </property>
        </Properties>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_names(self, custom_property_list):
        prop1 = StringProperty(name="PropName1", value="Something")
        prop2 = LinkProperty(name="PropName2", value="A link")
        prop_list = custom_property_list()
        prop_list.props = [prop1, prop2]
        assert prop_list.names == ["PropName1", "PropName2"]

    def test_duplicate(self, custom_property_list):
        prop1 = StringProperty(name="PropName1", value="Something")
        prop2 = LinkProperty(name="PropName1", value="A link")
        prop_list = custom_property_list()
        prop_list.props = [prop1]
        with pytest.raises(ValueError):
            prop_list.append(prop2)

    def test_from_link(self, custom_property_list):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                    linkTarget="A link">
                <vt:lpwstr/>
            </property>
        </Properties>
        """
        tree = fromstring(src)
        new_props = custom_property_list.from_tree(tree)
        assert new_props.props[0] == LinkProperty(name="PropName1", value="A link")

    def test_from_datetime(self, custom_property_list):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:filetime>2022-05-31T12:55:13Z</vt:filetime>
            </property>
        </Properties>
        """
        tree = fromstring(src)
        new_props = custom_property_list.from_tree(tree)
        expected = DateTimeProperty(
            name="PropName1",
            value=datetime.datetime(2022, 5, 31, 12, 55, 13),
        )
        assert new_props.props[0] == expected

    def test_from_string(self, custom_property_list):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:lpwstr>Something</vt:lpwstr>
            </property>
        </Properties>
        """
        tree = fromstring(src)
        new_props = custom_property_list.from_tree(tree)
        assert new_props.props[0] == StringProperty(name="PropName1", value="Something")

    def test_from_empty_string(self, custom_property_list):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:lpwstr/>
            </property>
        </Properties>
        """
        tree = fromstring(src)
        new_props = custom_property_list.from_tree(tree)
        assert new_props.props[0] == StringProperty(name="PropName1", value=None)

    def test_from_float(self, custom_property_list):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:r8>15</vt:r8>
            </property>
        </Properties>
        """
        tree = fromstring(src)
        new_props = custom_property_list.from_tree(tree)
        assert new_props.props[0] == FloatProperty(name="PropName1", value=15)

    def test_from_int(self, custom_property_list):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:i4>15</vt:i4>
            </property>
        </Properties>
        """
        tree = fromstring(src)
        new_props = custom_property_list.from_tree(tree)
        assert new_props.props[0] == IntProperty(name="PropName1", value=15)

    def test_from_bool(self, custom_property_list):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2" 
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:bool>0</vt:bool>
            </property>
        </Properties>
        """
        tree = fromstring(src)
        new_props = custom_property_list.from_tree(tree)
        assert new_props.props[0] == BoolProperty(name="PropName1", value=False)

    def test_unknown_type(self, custom_property_list, recwarn):
        src = """
        <Properties
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
            <property
                    name="PropName1"
                    pid="2"
                    fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}">
                <vt:cy>4.1256</vt:cy>
            </property>
        </Properties>
        """
        tree = fromstring(src)
        new_props = custom_property_list.from_tree(tree)
        assert recwarn.pop().category == UserWarning

    def test_cant_adapt(self, custom_property_list):
        from openpyxl.packaging.custom import _TypedProperty

        class DummyProperty(_TypedProperty):
            pass

        prop_list = custom_property_list()
        prop_list.append(DummyProperty(name="PropName1", value="Something"))
        with pytest.raises(TypeError):
            tree = prop_list.to_tree()
