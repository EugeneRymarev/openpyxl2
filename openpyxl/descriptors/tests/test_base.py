# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.descriptors import Strict


@pytest.fixture
def boolean():
    from openpyxl.descriptors.base import Bool

    class Dummy(Strict):
        value = Bool()

    return Dummy()


@pytest.fixture
def integer():
    from openpyxl.descriptors.base import Integer

    class Dummy(Strict):
        value = Integer()

    return Dummy()


@pytest.fixture
def float_():
    from openpyxl.descriptors.base import Float

    class Dummy(Strict):
        value = Float()

    return Dummy()


@pytest.fixture
def allow_none():
    from openpyxl.descriptors.base import Float

    class Dummy(Strict):
        value = Float(allow_none=True)

    return Dummy()


@pytest.fixture
def maximum():
    from openpyxl.descriptors.base import Max

    class Dummy(Strict):
        value = Max(max=5)

    return Dummy()


@pytest.fixture
def minimum():
    from openpyxl.descriptors.base import Min

    class Dummy(Strict):
        value = Min(min=0)

    return Dummy()


@pytest.fixture
def min_max():
    from openpyxl.descriptors.base import MinMax

    class Dummy(Strict):
        value = MinMax(min=-1, max=1)

    return Dummy()


@pytest.fixture
def set_():
    from openpyxl.descriptors.base import Set

    class Dummy(Strict):
        value = Set(values=[1, "a", None])

    return Dummy()


@pytest.fixture
def ascii_():
    from openpyxl.descriptors.base import ASCII

    class Dummy(Strict):
        value = ASCII()

    return Dummy()


@pytest.fixture
def string():
    from openpyxl.descriptors.base import String

    class Dummy(Strict):
        value = String()

    return Dummy()


@pytest.fixture
def tuple_():
    from openpyxl.descriptors.base import Tuple

    class Dummy(Strict):
        value = Tuple()

    return Dummy()


@pytest.fixture
def length():
    from openpyxl.descriptors.base import Length

    class Dummy(Strict):
        value = Length(length=4)

    return Dummy()


class TestDescriptor:
    from openpyxl.descriptors.base import Descriptor

    class Dummy:
        pass

    def test_ctor(self):
        d = self.Descriptor("key", size=1)
        assert d.name == "key"
        assert d.size == 1

    def test_setter(self):
        d = self.Descriptor("key")
        client = self.Dummy()
        d.__set__(client, 42)
        assert client.key == 42


class TestBool:
    def test_valid(self, boolean):
        boolean.value = True
        assert boolean.value

    @pytest.mark.parametrize(
        "value, expected",
        [
            (1, True),
            (0, False),
            ("true", True),
            ("false", False),
            ("0", False),
            ("f", False),
            ("", False),
            ([], False),
        ],
    )
    def test_cast(self, boolean, value, expected):
        boolean.value = value
        assert boolean.value == expected


def test_nested():
    from openpyxl.descriptors.base import Bool

    class DummyNested(Strict):
        value = Bool(nested=True)

    dummy = DummyNested()
    dummy.value = True
    assert dummy.__class__.value.nested == True


class TestInt:
    def test_valid(self, integer):
        integer.value = 4
        assert integer.value == 4

    @pytest.mark.parametrize("value", ["a", "4.5", None])
    def test_invalid(self, integer, value):
        with pytest.raises(TypeError):
            integer.value = value

    @pytest.mark.parametrize("value, expected", [("4", 4), (4.5, 4)])
    def test_cast(self, integer, value, expected):
        integer.value = value
        assert integer.value == expected


class TestFloat:
    def test_valid(self, float_):
        float_.value = 4
        assert float_.value == 4

    @pytest.mark.parametrize("value", ["a", None])
    def test_invalid(self, float_, value):
        with pytest.raises(TypeError):
            float_.value = value

    @pytest.mark.parametrize("value, expected", [("4.5", 4.5), (4.5, 4.5), (4, 4.0)])
    def test_cast(self, float_, value, expected):
        float_.value = value
        assert float_.value == expected


class TestAllowNone:
    def test_valid(self, allow_none):
        allow_none.value = None
        assert allow_none.value is None


class TestMax:
    def test_ctor(self):
        from openpyxl.descriptors.base import Max

        with pytest.raises(TypeError):

            class Dummy(Strict):
                value = Max()

    def test_valid(self, maximum):
        maximum.value = 4
        assert maximum.value == 4

    def test_invalid(self, maximum):
        with pytest.raises(ValueError):
            maximum.value = 6


class TestMin:
    def test_ctor(self):
        from openpyxl.descriptors.base import Min

        with pytest.raises(TypeError):

            class Dummy(Strict):
                value = Min()

    def test_valid(self, minimum):
        minimum.value = 2
        assert minimum.value == 2

    def test_invalid(self, minimum):
        with pytest.raises(ValueError):
            minimum.value = -1


class TestMinMax:
    def test_ctor(self):
        from openpyxl.descriptors.base import MinMax

        with pytest.raises(TypeError):

            class Dummy(Strict):
                value = MinMax(min=-10)

        with pytest.raises(TypeError):

            class Dummy(Strict):
                value = MinMax(max=10)

    def test_valid(self, min_max):
        min_max.value = 1
        assert min_max.value == 1

    def test_invalid(self, min_max):
        with pytest.raises(ValueError):
            min_max.value = 2


class TestValues:
    def test_ctor(self):
        from openpyxl.descriptors.base import Set

        with pytest.raises(TypeError):

            class Dummy(Strict):
                value = Set()

    def test_valid(self, set_):
        set_.value = 1
        assert set_.value == 1

    def test_invalid(self, set_):
        with pytest.raises(ValueError):
            set_.value = 2


def test_noneset():
    from openpyxl.descriptors.base import NoneSet

    class Dummy(Strict):
        value = NoneSet(values=[1, 2, 3])

    obj = Dummy()
    obj.value = "none"
    assert obj.value is None
    with pytest.raises(ValueError):
        obj.value = 5


class TestASCII:
    def test_valid(self, ascii_):
        ascii_.value = b"some text"
        assert ascii_.value == b"some text"

    @pytest.mark.parametrize("value", [b"\xc3\xbc".decode("utf-8"), 10, []])
    def test_invalid(self, ascii_, value):
        with pytest.raises(TypeError):
            ascii_.value = value


class TestString:
    def test_valid(self, string):
        value = b"\xc3\xbc".decode("utf-8")
        string.value = value
        assert string.value == value

    def test_invalid(self, string):
        with pytest.raises(TypeError):
            string.value = 5


class TestTuple:
    def test_valid(self, tuple_):
        tuple_.value = (1, 2)
        assert tuple_.value == (1, 2)

    def test_invalid(self, tuple_):
        with pytest.raises(TypeError):
            tuple_.value = [1, 2, 3]


class TestLength:
    def test_valid(self, length):
        length.value = "this"

    def test_invalid(self, length):
        with pytest.raises(ValueError):
            length.value = "2"
