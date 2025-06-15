# Copyright (c) 2010-2025 openpyxl
import datetime

import pytest


@pytest.mark.parametrize(
    "value, expected",
    [
        (datetime.datetime(2013, 7, 15, 6, 52, 33), "2013-07-15T06:52:33"),
        (datetime.datetime(2013, 7, 15, 6, 52, 33, 123456), "2013-07-15T06:52:33.123"),
        (datetime.date(2013, 7, 15), "2013-07-15"),
        (datetime.time(0, 1, 42), "00:01:42"),
        (datetime.time(0, 1, 42, 123456), "00:01:42.123"),
    ],
)
def test_to_iso(value, expected):
    from openpyxl.utils.datetime import to_ISO8601

    assert to_ISO8601(value) == expected


@pytest.mark.parametrize(
    "value, group, expected",
    [
        ("2011-06-30", "date", "2011-06-30"),
        ("12:19", "time", "12:19"),
        ("12:19:01", "time", "12:19:01"),
        ("12:19:01.123", "time", "12:19:01.123"),
        ("12:19:01.2", "time", "12:19:01.2"),
    ],
)
def test_iso_regex(value, group, expected):
    from openpyxl.utils.datetime import ISO_REGEX

    match = ISO_REGEX.match(value)
    assert match is not None
    assert match.groupdict()[group] == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        ("2011-06-30T13:35:26Z", datetime.datetime(2011, 6, 30, 13, 35, 26)),
        ("2013-03-04T12:19:01.00Z", datetime.datetime(2013, 3, 4, 12, 19, 1)),
        ("2011-06-30", datetime.date(2011, 6, 30)),
        ("12:19", datetime.time(12, 19)),
        ("12:19:01", datetime.time(12, 19, 1)),
        ("12:19:01.123", datetime.time(12, 19, 1, 123_000)),
        ("12:19:01.2", datetime.time(12, 19, 1, 200_000)),
        (
            "2020-12-03T12:19:01.300Z",
            datetime.datetime(2020, 12, 3, 12, 19, 1, 300_000),
        ),
        ("2020-12-03T12:19:01.030", datetime.datetime(2020, 12, 3, 12, 19, 1, 30_000)),
        ("2020-12-03T12:19:01.003Z", datetime.datetime(2020, 12, 3, 12, 19, 1, 3000)),
        ("2020-12-03T12:19:01.3Z", datetime.datetime(2020, 12, 3, 12, 19, 1, 300_000)),
        ("2020-12-03T12:19:01.03", datetime.datetime(2020, 12, 3, 12, 19, 1, 30_000)),
        ("PT0M", datetime.timedelta(0)),
        ("PT2H0M1S", datetime.timedelta(hours=2, seconds=1)),
        ("PT25H20M1.1S", datetime.timedelta(days=1, hours=1, minutes=20, seconds=1.1)),
        (
            "PT25H70M1.123S",
            datetime.timedelta(days=1, hours=2, minutes=10, seconds=1.123),
        ),
    ],
)
def test_from_iso(value, expected):
    from openpyxl.utils.datetime import from_ISO8601

    assert from_ISO8601(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (datetime.date(1899, 12, 31), 0),
        (datetime.date(1900, 1, 1), 1),
        (datetime.date(1900, 1, 15), 15),
        (datetime.date(1900, 2, 28), 59),
        (datetime.datetime(1900, 2, 28, 21, 0, 0), 59.875),
        (datetime.date(1900, 3, 1), 61),
        (datetime.datetime(2010, 1, 18, 14, 15, 20, 1600), 40196.5939815),
        (datetime.date(2009, 12, 20), 40167),
        (datetime.datetime(1506, 10, 15), -143617.0),
        (datetime.date(1, 1, 1), -693593),
        (datetime.time(0), 0),
        (datetime.time(6, 0), 0.25),
        (datetime.timedelta(hours=6), 0.25),
        (datetime.timedelta(hours=-6), -0.25),
    ],
)
def test_to_excel(value, expected):
    from openpyxl.utils.datetime import to_excel

    fut = to_excel
    assert fut(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (datetime.date(1904, 1, 1), 0),
        (datetime.date(2011, 10, 31), 39385),
        (datetime.datetime(2010, 1, 18, 14, 15, 20, 1600), 38734.5939815),
        (datetime.date(2009, 12, 20), 38705),
        (datetime.datetime(1506, 10, 15), -145079.0),
        (datetime.date(1, 1, 1), -695055),
        (datetime.time(0), 0),
        (datetime.time(6, 0), 0.25),
        (datetime.timedelta(hours=6), 0.25),
        (datetime.timedelta(hours=-6), -0.25),
    ],
)
def test_to_excel_mac(value, expected):
    from openpyxl.utils.datetime import CALENDAR_MAC_1904
    from openpyxl.utils.datetime import to_excel

    fut = to_excel
    assert fut(value, CALENDAR_MAC_1904) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (40167, datetime.datetime(2009, 12, 20)),
        (21980, datetime.datetime(1960, 3, 5)),
        (59, datetime.datetime(1900, 2, 28)),
        (-25063, datetime.datetime(1831, 5, 18, 0, 0)),
        (59.875, datetime.datetime(1900, 2, 28, 21, 0, 0)),
        (61, datetime.datetime(1900, 3, 1, 0, 0)),
        (40372.27616898148, datetime.datetime(2010, 7, 13, 6, 37, 41)),
        (40196.5939815, datetime.datetime(2010, 1, 18, 14, 15, 20, 2000)),
        (0.125, datetime.time(3, 0)),
        (42126.958333333219, datetime.datetime(2015, 5, 2, 23, 0, 0, 0)),
        (42126.999999999884, datetime.datetime(2015, 5, 3, 0, 0, 0)),
        (0, datetime.time(0)),
        (0.9999999995, datetime.datetime(1900, 1, 1)),
        (1, datetime.datetime(1900, 1, 1)),
        (-0.25, datetime.datetime(1899, 12, 29, 18, 0, 0)),
        (None, None),
    ],
)
def test_from_excel(value, expected):
    from openpyxl.utils.datetime import from_excel

    fut = from_excel
    assert fut(value) == expected


@pytest.mark.parametrize("value", [60, 60.5])
def test_invalid_date(value):
    from openpyxl.utils.datetime import from_excel

    with pytest.raises(ValueError):
        from_excel(value)


@pytest.mark.parametrize(
    "value, expected",
    [
        (0, datetime.timedelta(hours=0)),
        (0.5, datetime.timedelta(hours=12)),
        (-0.5, datetime.timedelta(hours=-12)),
        (1.25, datetime.timedelta(hours=30)),
        (-1.25, datetime.timedelta(hours=-30)),
        (0.0006944443, datetime.timedelta(minutes=1)),
        (-0.0006944443, datetime.timedelta(minutes=-1)),
        (0.0006944328, datetime.timedelta(minutes=1, microseconds=-1000)),
        (-0.0006944328, datetime.timedelta(minutes=-1, microseconds=1000)),
        (59.5, datetime.timedelta(days=59, hours=12)),
        (60.5, datetime.timedelta(days=60, hours=12)),
        (61.5, datetime.timedelta(days=61, hours=12)),
        (0.9999999995, datetime.timedelta(days=1)),
        (1.0000000005, datetime.timedelta(days=1)),
        (1.0000026378, datetime.timedelta(days=1, microseconds=228000)),
        (None, None),
    ],
)
def test_from_excel_timedelta(value, expected):
    from openpyxl.utils.datetime import from_excel

    fut = from_excel
    assert fut(value, timedelta=True) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (39385, datetime.datetime(2011, 10, 31)),
        (21980, datetime.datetime(1964, 3, 6)),
        (0, datetime.time(0, 0)),
        (-25063, datetime.datetime(1835, 5, 19)),
        (0.75, datetime.time(18, 0)),
        (-0.25, datetime.datetime(1903, 12, 31, 18, 0, 0)),
    ],
)
def test_from_excel_mac(value, expected):
    from openpyxl.utils.datetime import CALENDAR_MAC_1904
    from openpyxl.utils.datetime import from_excel

    fut = from_excel
    assert fut(value, CALENDAR_MAC_1904) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (datetime.time(13, 55, 12, 36), 0.5800000004166667),
        (datetime.time(3, 0, 0), 0.125),
        (datetime.datetime(2021, 3, 19, 13, 55, 12, 36), 0.5800000004166667),
        (datetime.datetime(1536, 12, 24, 3, 0, 0), 0.125),
    ],
)
def test_time_to_days(value, expected):
    from openpyxl.utils.datetime import time_to_days

    fut = time_to_days
    assert fut(value) == expected


def test_timedelta_to_days():
    from openpyxl.utils.datetime import timedelta_to_days

    fut = timedelta_to_days
    td = datetime.timedelta(days=1, hours=3)
    assert fut(td) == 1.125


def test_days_to_time():
    from openpyxl.utils.datetime import days_to_time

    td = datetime.timedelta(0, 51320, 1600)
    fut = days_to_time
    assert fut(td) == datetime.time(14, 15, 20, 1600)
