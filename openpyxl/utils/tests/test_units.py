import pytest

from openpyxl.utils.units import EMU_to_cm
from openpyxl.utils.units import EMU_to_inch
from openpyxl.utils.units import EMU_to_pixels
from openpyxl.utils.units import angle_to_degrees
from openpyxl.utils.units import cm_to_dxa
from openpyxl.utils.units import cm_to_EMU
from openpyxl.utils.units import degrees_to_angle
from openpyxl.utils.units import dxa_to_cm
from openpyxl.utils.units import dxa_to_inch
from openpyxl.utils.units import inch_to_dxa
from openpyxl.utils.units import inch_to_EMU
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils.units import pixels_to_points
from openpyxl.utils.units import points_to_pixels
from openpyxl.utils.units import short_color


@pytest.mark.parametrize(
    "value, expected",
    [
        (-120, -0.08333333333333334),
        (0, 0),
        (240, 0.16666666666666669),
        (1440, 1),
        (5000, 3.4722222222222223),
    ],
)
def test_dxa_to_inch(value, expected):
    fut = dxa_to_inch
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, -14400), (0, 0), (1, 1440), (2.37, 3412), (9, 12960)],
)
def test_inch_to_dxa(value, expected):
    fut = inch_to_dxa
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (-120, -0.2116666666666667),
        (0, 0),
        (240, 0.4233333333333334),
        (1440, 2.54),
        (5000, 8.819444444444445),
    ],
)
def test_dxa_to_cm(value, expected):
    fut = dxa_to_cm
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, -5669), (0, 0), (1, 566), (10.0, 5669), (1000, 566929)],
)
def test_cm_to_dxa(value, expected):
    fut = cm_to_dxa
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, -95250), (0, 0), (1, 9525), (10.0, 95250), (1000, 9525000)],
)
def test_pixels_to_emu(value, expected):
    fut = pixels_to_EMU
    assert fut(value) == expected


@pytest.mark.parametrize("value, expected", [(0, 0), (1000, 0), (5000, 1), (9525, 1)])
def test_emu_to_pixels(value, expected):
    fut = EMU_to_pixels
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-100000, -0.2778), (0, 0), (200000, 0.5556), (360000, 1), (500000, 1.3889)],
)
def test_emu_to_cm(value, expected):
    fut = EMU_to_cm
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, -3600000), (0, 0), (1, 360000), (3.23, 1162800)],
)
def test_cm_to_emu(value, expected):
    fut = cm_to_EMU
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-100000, -0.1094), (0, 0), (200000, 0.2187), (914400, 1), (500000, 0.5468)],
)
def test_emu_to_inch(value, expected):
    fut = EMU_to_inch
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, -9144000), (0, 0), (1, 914400), (3.23, 2953512)],
)
def test_inch_to_emu(value, expected):
    fut = inch_to_EMU
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, -7.5), (0, 0), (1, 0.75), (96, 72), (144, 108)],
)
def test_pixels_to_points(value, expected):
    fut = pixels_to_points
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, -13), (0, 0), (1, 2), (10.0, 14), (72, 96)],
)
def test_points_to_pixels(value, expected):
    fut = points_to_pixels
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, -600000), (0, 0), (1, 60000), (10.0, 600000), (1000, 60000000)],
)
def test_degrees_to_angle(value, expected):
    fut = degrees_to_angle
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [(-10, 0), (0, 0), (10, 0), (50000, 0.83), (60000, 1)],
)
def test_angle_to_degrees(value, expected):
    fut = angle_to_degrees
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        ("#FFFFF", "#FFFFF"),
        ("FF000000", "000000"),
        ("FFFF0000", "FF0000"),
        ("FF800000", "800000"),
        ("FFFFFF00", "FFFF00"),
        ("FF808000", "808000"),
    ],
)
def test_short_color(value, expected):
    fut = short_color
    assert fut(value) == expected
