# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.descriptors.base import Integer
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.styles.fonts import Font
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import Element
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def sequence():
    from openpyxl.descriptors.sequence import Sequence

    return Sequence


@pytest.fixture
def value_sequence():
    from openpyxl.descriptors.sequence import ValueSequence

    return ValueSequence


@pytest.fixture
def nested_sequence():
    from openpyxl.descriptors.sequence import NestedSequence

    return NestedSequence


@pytest.fixture
def multi_sequence():
    from openpyxl.descriptors.sequence import MultiSequence

    return MultiSequence


@pytest.fixture
def multi_sequence_part():
    from openpyxl.descriptors.sequence import MultiSequencePart

    return MultiSequencePart


@pytest.fixture
def dummy_(sequence):
    class Dummy(Serialisable):
        value = sequence(expected_type=int)

        def __init__(self, value=()):
            self.value = value

    return Dummy


class TestSequence:
    @pytest.mark.parametrize("value", [list(), tuple()])
    def test_valid_ctor(self, dummy_, value):
        dummy = dummy_()
        dummy.value = value
        assert dummy.value == list(value)

    @pytest.mark.parametrize("value", ["", b"", dict(), 1, None])
    def test_invalid_container(self, dummy_, value):
        dummy = dummy_()
        with pytest.raises(TypeError):
            dummy.value = value


class TestPrimitive:
    def test_to_tree(self, dummy_):
        dummy = dummy_([1, "2", 3])
        root = Element("root")
        nodes = dummy_.value.to_tree("el", dummy.value)
        for node in nodes:
            root.append(node)
        xml = tostring(root)
        expected = """
        <root>
            <el>1</el>
            <el>2</el>
            <el>3</el>
        </root>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, dummy_):
        src = """
        <root>
            <value>1</value>
            <value>2</value>
            <value>3</value>
        </root>
        """
        node = fromstring(src)
        dummy = dummy_.from_tree(node)
        assert dummy.value == [1, 2, 3]


class SomeType(Serialisable):
    value = Integer()

    def __init__(self, value):
        self.value = value


class TestComplex:
    def test_to_tree(self, sequence):
        class Dummy:
            vals = sequence(expected_type=SomeType, name="vals")

        dummy = Dummy()
        dummy.vals = [SomeType(1), SomeType(2), SomeType(3)]
        root = Element("root")
        for node in Dummy.vals.to_tree("el", dummy.vals):
            root.append(node)
        xml = tostring(root)
        expected = """
        <root>
            <el value="1"></el>
            <el value="2"></el>
            <el value="3"></el>
        </root>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, sequence):
        class Dummy(Serialisable):
            vals = sequence(expected_type=SomeType)

            def __init__(self, vals):
                self.vals = vals

        src = """
        <root>
            <vals value="1"></vals>
            <vals value="2"></vals>
            <vals value="3"></vals>
        </root>
        """
        node = fromstring(src)
        dummy = Dummy.from_tree(node)
        assert dummy.vals == [SomeType(1), SomeType(2), SomeType(3)]


class TestValueSequence:
    def test_to_tree(self, value_sequence):
        class Dummy(Serialisable):
            tagname = "el"
            size = value_sequence(expected_type=int)

        dummy = Dummy()
        dummy.size = [1, 2, 3]
        xml = tostring(dummy.to_tree())
        expected = """
        <el>
            <size val="1"></size>
            <size val="2"></size>
            <size val="3"></size>
        </el>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_tree(self, value_sequence):
        class Dummy(Serialisable):
            tagname = "el"
            __nested__ = ("size",)
            size = value_sequence(expected_type=int)

            def __init__(self, size):
                self.size = size

        src = """
        <el>
            <size val="1"></size>
            <size val="2"></size>
            <size val="3"></size>
        </el>
        """
        node = fromstring(src)
        desc = Dummy.size
        vals = desc.from_tree(node)
        dummy = Dummy.from_tree(node)
        assert dummy.size == [1, 2, 3]


@pytest.fixture
def complex_object(nested_sequence):
    class Complex(Serialisable):
        tagname = "style"
        fonts = nested_sequence(expected_type=Font, count=True)

        def __init__(self, fonts=()):
            self.fonts = fonts

    return Complex


class TestNestedSequence:
    def test_ctor(self, complex_object):
        style = complex_object()
        ft1 = Font(family=2, sz=11, name="Arial")
        ft2 = Font(bold=True)
        style.fonts = [ft1, ft2]
        expected = """
        <style>
            <fonts count="2">
                <font>
                    <name val="Arial"/>
                    <family val="2"></family>
                    <sz val="11"></sz>
                </font>
                <font>
                    <b val="1"></b>
                </font>
            </fonts>
        </style>
        """
        tree = style.__class__.fonts.to_tree("fonts", style.fonts)
        tree = style.to_tree()
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_tree(self, complex_object):
        xml = """
        <style>
            <fonts count="2">
                <font>
                    <name val="Calibri"></name>
                    <family val="2"></family>
                    <color rgb="00000000"></color>
                    <sz val="11"></sz>
                </font>
                <font>
                    <name val="Calibri"></name>
                    <family val="2"></family>
                    <b val="1"></b>
                    <color rgb="00000000"></color>
                    <sz val="11"></sz>
                </font>
            </fonts>
        </style>
        """
        node = fromstring(xml)
        style = complex_object.from_tree(node)
        assert len(style.fonts) == 2
        assert style.fonts[1].bold is True


class Larry(Serialisable):
    tagname = "l"
    value = Integer()

    def __init__(self, value):
        self.value = value


class Curly(Serialisable):
    tagname = "c"
    hair = Integer()

    def __init__(self, hair):
        self.hair = hair


class Mo(Serialisable):
    tagname = "m"
    cap = Integer()

    def __init__(self, cap):
        self.cap = cap


@pytest.fixture
def stooge(multi_sequence, multi_sequence_part):
    class Stooge(Serialisable):
        _stooges = multi_sequence(expected_type=SomeType)
        l = multi_sequence_part(expected_type=Larry, store="_stooges")
        c = multi_sequence_part(expected_type=Curly, store="_stooges")
        m = multi_sequence_part(expected_type=Mo, store="_stooges")

        def __init__(self, _stooges=()):
            self._stooges = _stooges

    return Stooge


class TestMultiSequence:
    def test_elements(self, stooge):
        assert stooge.__elements__ == ("_stooges",)

    def test_attrs(self, stooge):
        dummy = stooge()
        assert stooge.__attrs__ == ()

    def test_to_tree(self, stooge):
        dummy = stooge()
        dummy._stooges = [Larry(1), Curly(2), Larry(3), Mo(4)]
        root = Element("root")
        for node in stooge._stooges.to_tree("el", dummy._stooges):
            root.append(node)
        tree = dummy.to_tree("root")
        xml = tostring(root)
        expected = """
        <root>
            <l value="1"></l>
            <c hair="2"></c>
            <l value="3"></l>
            <m cap="4"></m>
        </root>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, stooge):
        src = """
        <root>
            <l value="1"></l>
            <c hair="2"></c>
            <l value="3"></l>
            <m cap="4"></m>
        </root>
        """
        node = fromstring(src)
        dummy = stooge.from_tree(node)
        assert dummy._stooges == [Larry(1), Curly(2), Larry(3), Mo(4)]
