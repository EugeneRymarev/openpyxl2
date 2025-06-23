# Copyright (c) 2010-2025 openpyxl
import copy

import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def serialisable():
    from openpyxl.descriptors.serialisable import Serialisable

    return Serialisable


@pytest.fixture
def relation(serialisable):
    from openpyxl.descriptors.excel import Relation

    class Dummy(serialisable):
        tagname = "dummy"
        rId = Relation()

        def __init__(self, rId=None):
            self.rId = rId

    return Dummy


@pytest.fixture
def keyword_attribute(serialisable):
    from openpyxl.descriptors.base import Bool

    class SomeElement(serialisable):
        tagname = "dummy"
        _from = Bool()

        def __init__(self, _from):
            self._from = _from

    return SomeElement


@pytest.fixture
def node(serialisable):
    from openpyxl.descriptors.base import Bool

    class SomeNode(serialisable):
        tagname = "from"
        val = Bool()

        def __init__(self, val):
            self.val = val

    return SomeNode


@pytest.fixture
def keyword_node(serialisable, node):
    from openpyxl.descriptors.base import Typed

    class SomeElement(serialisable):
        tagname = "dummy"
        _from = Typed(expected_type=node)

        def __init__(self, _from):
            self._from = _from

    return SomeElement


@pytest.fixture
def hyphenated_attribute(serialisable):
    from openpyxl.descriptors.base import Bool

    class SomeElement(serialisable):
        tagname = "dummy"
        z_order = Bool(hyphenated=True)
        a_order = Bool()

        def __init__(self, z_order, a_order):
            self.z_order = z_order
            self.a_order = a_order

    return SomeElement


@pytest.fixture
def expected_types(serialisable):
    from openpyxl.descriptors.base import Typed

    class Dummy(serialisable):
        tagname = "dummy"
        value = Typed(expected_type=(str, int))
        __attrs__ = ("value",)

        def __init__(self, value):
            self.value = value

    return Dummy


@pytest.fixture
def Immutable(serialisable):
    class Immutable(serialisable):
        __attrs__ = ("value",)

        def __init__(self, value=None):
            self.value = value

    return Immutable


class TestSerialisable:
    def test_hash(self, Immutable):
        d1 = Immutable()
        d2 = Immutable()
        assert hash(d1) == hash(d2)

    def test_add_attrs(self, Immutable):
        d1 = Immutable()
        d2 = Immutable(value=2)
        assert d1 + d2 == d2

    def test_str(self, Immutable):
        d = Immutable()
        expected = """<openpyxl.descriptors.tests.test_serialisable.Immutable object>
Parameters:
value=None"""
        assert str(d) == expected
        d2 = Immutable("hello")
        expected = """<openpyxl.descriptors.tests.test_serialisable.Immutable object>
Parameters:
value='hello'"""
        assert str(d2) == expected

    def test_eq(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(1)
        assert d1 is not d2
        assert d1 == d2

    def test_ne(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(2)
        assert d1 != d2

    def test_copy(self, Immutable):
        d1 = Immutable({})
        d2 = copy.copy(d1)
        assert d1.value is not d2.value


class TestRelation:
    def test_binding(self, relation):
        expected = (
            (
                "rId",
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}rId",
            ),
        )
        assert relation.__namespaced__ == expected

    def test_to_tree(self, relation):
        dummy = relation("rId1")
        xml = tostring(dummy.to_tree())
        expected = """
        <dummy r:rId="rId1"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_tree(self, relation):
        src = """
        <dummy r:rId="rId1"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
        """
        n = fromstring(src)
        obj = relation.from_tree(n)
        assert obj.rId == "rId1"


class TestKeywordAttribute:
    def test_to_tree(self, keyword_attribute):
        dummy = keyword_attribute(_from=True)
        xml = tostring(dummy.to_tree())
        expected = '<dummy from="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_tree(self, keyword_attribute):
        src = '<dummy from="1"/>'
        el = fromstring(src)
        dummy = keyword_attribute.from_tree(el)
        assert dummy._from is True


class TestKeywordNode:
    def test_to_tree(self, keyword_node, node):
        n = node(val=True)
        dummy = keyword_node(_from=n)
        xml = tostring(dummy.to_tree())
        expected = '<dummy><from val="1"/></dummy>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_tree(self, keyword_node):
        src = '<dummy><from val="1"/></dummy>'
        el = fromstring(src)
        dummy = keyword_node.from_tree(el)
        assert dummy._from.val is True


class TestHyphenatedAttribute:
    def test_to_tree(self, hyphenated_attribute):
        dummy = hyphenated_attribute(z_order=True, a_order=True)
        xml = tostring(dummy.to_tree())
        expected = '<dummy z-order="1" a_order="1"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_tree(self, hyphenated_attribute):
        src = '<dummy z-order="1" a_order="1"/>'
        el = fromstring(src)
        dummy = hyphenated_attribute.from_tree(el)
        assert dummy.z_order is True
        assert dummy.a_order is True


class TestExpectedTypes:
    @pytest.mark.parametrize("value", ["a", 2])
    def test_valid(self, expected_types, value):
        obj = expected_types(value)
        assert obj.value is value

    def test_to_tree(self, expected_types):
        obj = expected_types(value=2)
        assert dict(obj) == {"value": "2"}
        xml = tostring(obj.to_tree())
        diff = compare_xml(xml, """<dummy value="2"/>""")
        assert diff is None, diff

    def test_from_tree(self, expected_types):
        xml = "<dummy><value>1</value></dummy>"
        n = fromstring(xml)
        obj = expected_types.from_tree(n)
        assert obj.value == "1"
