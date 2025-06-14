# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.descriptors import Strict


@pytest.fixture
def universal_measure():
    from openpyxl.descriptors.excel import UniversalMeasure

    class Dummy(Strict):
        value = UniversalMeasure()

    return Dummy()


@pytest.fixture
def hex_binary():
    from openpyxl.descriptors.excel import HexBinary

    class Dummy(Strict):
        value = HexBinary()

    return Dummy()


@pytest.fixture
def text_point():
    from openpyxl.descriptors.excel import TextPoint

    class Dummy(Strict):
        value = TextPoint()

    return Dummy()


@pytest.fixture
def percentage():
    from openpyxl.descriptors.excel import Percentage

    class Dummy(Strict):
        value = Percentage()

    return Dummy()


@pytest.fixture
def guid():
    from openpyxl.descriptors.excel import Guid

    class Dummy(Strict):
        value = Guid()

    return Dummy()


@pytest.fixture
def base64_binary():
    from openpyxl.descriptors.excel import Base64Binary

    class Dummy(Strict):
        value = Base64Binary()

    return Dummy()


@pytest.fixture
def cell_range():
    from openpyxl.descriptors.excel import CellRange

    class Dummy(Strict):
        value = CellRange()

    return Dummy()


class TestUniversalMeasure:
    @pytest.mark.parametrize("value", ["24.73mm", "0cm", "24pt", "999pc", "50pi"])
    def test_valid(self, universal_measure, value):
        universal_measure.value = value
        assert universal_measure.value == value

    @pytest.mark.parametrize("value", [24.73, "24.73zz", "24.73 mm", None, "-24.73cm"])
    def test_invalid(self, universal_measure, value):
        with pytest.raises(ValueError):
            universal_measure.value = str(value)


class TestHexBinary:
    @pytest.mark.parametrize("value", ["aa35efd", "AABBCCDD"])
    def test_valid(self, hex_binary, value):
        hex_binary.value = value
        assert hex_binary.value == value

    @pytest.mark.parametrize("value", ["GGII", "35.5"])
    def test_invalid(self, hex_binary, value):
        with pytest.raises(ValueError):
            hex_binary.value = value


class TestTextPoint:
    @pytest.mark.parametrize("value", [-400000, "400000", 0])
    def test_valid(self, text_point, value):
        text_point.value = value
        assert text_point.value == int(value)

    def test_invalid_value(self, text_point):
        with pytest.raises(ValueError):
            text_point.value = -400001

    def test_invalid_type(self, text_point):
        with pytest.raises(TypeError):
            text_point.value = "40pt"


class TestPercentage:
    @pytest.mark.parametrize(
        "input, value",
        [("15%", 15000), (1500, 1500), ("15.5%", 15500)],
    )
    def test_valid(self, percentage, input, value):
        percentage.value = value
        assert percentage.value == value

    @pytest.mark.parametrize("value", ["2000000", "-1000001"])
    def test_invalid(self, percentage, value):
        with pytest.raises(ValueError):
            percentage.value = value


class TestGuid:
    @pytest.mark.parametrize("value", ["{00000000-5BD2-4BC8-9F70-7020E1357FB2}"])
    def test_valid(self, guid, value):
        guid.value = value
        assert guid.value == value

    @pytest.mark.parametrize("value", ["{00000000-5BD2-4BC8-9F70-7020E1357FB2"])
    def test_valid(self, guid, value):
        with pytest.raises(ValueError):
            guid.value = value


class TestBase64Binary:
    @pytest.mark.parametrize("value", ["9oN7nWkCAyEZib1RomSJTjmPpCY="])
    def test_valid(self, base64_binary, value):
        base64_binary.value = value
        assert base64_binary.value == value

    @pytest.mark.parametrize("value", ["==0F"])
    def test_valid(self, base64_binary, value):
        with pytest.raises(ValueError):
            base64_binary.value = value


class TestCellRange:
    @pytest.mark.parametrize("value", ["A1", "A1:H5", "A:B"])
    def test_valid(self, cell_range, value):
        cell_range.value = value
        assert cell_range.value == value

    @pytest.mark.parametrize("value", ["A1:", "A1:5", "A1:B4:C7"])
    def test_invalid(self, cell_range, value):
        with pytest.raises(ValueError):
            cell_range.value = value
