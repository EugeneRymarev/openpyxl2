# Copyright (c) 2010-2024 openpyxl
import pytest

from ..numbers import FORMAT_DATE_DATETIME
from ..numbers import FORMAT_DATE_DDMMYY
from ..numbers import FORMAT_DATE_DMMINUS
from ..numbers import FORMAT_DATE_DMYSLASH
from ..numbers import FORMAT_DATE_MYMINUS
from ..numbers import FORMAT_DATE_TIME1
from ..numbers import FORMAT_DATE_TIME2
from ..numbers import FORMAT_DATE_TIME3
from ..numbers import FORMAT_DATE_TIME4
from ..numbers import FORMAT_DATE_TIME5
from ..numbers import FORMAT_DATE_TIME6
from ..numbers import FORMAT_DATE_TIME7
from ..numbers import FORMAT_DATE_TIME8
from ..numbers import FORMAT_DATE_TIMEDELTA
from ..numbers import FORMAT_DATE_XLSX14
from ..numbers import FORMAT_DATE_XLSX15
from ..numbers import FORMAT_DATE_XLSX16
from ..numbers import FORMAT_DATE_XLSX17
from ..numbers import FORMAT_DATE_XLSX22
from ..numbers import FORMAT_DATE_YYMMDD
from ..numbers import FORMAT_DATE_YYMMDDSLASH
from ..numbers import FORMAT_DATE_YYYYMMDD2
from openpyxl.styles import numbers


def test_builtin_format():
    fmt = "0.00"
    assert numbers.builtin_format_code(2) == fmt


def test_number_descriptor():
    from openpyxl.descriptors import Strict

    from ..numbers import NumberFormatDescriptor

    class Dummy(Strict):
        value = NumberFormatDescriptor()

        def __init__(self, value=None):
            self.value = value

    dummy = Dummy()
    assert dummy.value == "General"


@pytest.mark.parametrize(
    "fmt, stripped",
    [
        ('"Y: "#.000"m"', "#.000"),
        ("[Red]", ""),
        ('[$-404]e"\xfc"m"\xfc"d"\xfc"', "emd"),
        ("#,##0\\ [$\u20bd-46D]", "#,##0\\ "),
    ],
)
def test_strip_quotes(fmt, stripped):
    from ..numbers import STRIP_RE

    assert STRIP_RE.sub("", fmt) == stripped


@pytest.mark.parametrize(
    "format, result",
    [
        ("DD/MM/YY", True),
        ("H:MM:SS;@", True),
        ("#,##0\\ [$\u20bd-46D]", False),
        ('m"M"d"D";@', True),
        ("[h]:mm:ss", True),
        ('"Y: "0.00"m";"Y: "-0.00"m";"Y: <num>m";@', False),
        ("#,##0\\ [$\u20bd-46D]", False),
        ('"$"#,##0_);[Red]("$"#,##0)', False),
        ('[$-404]e"\xfc"m"\xfc"d"\xfc"', True),
        (r"0_ ;[Red]\-0\ ", False),
        (r"\Y000000", False),
        (r'#,##0.0####" YMD"', False),
        ("[h]", True),
        ("[ss]", True),
        ("[s].000", True),
        ("[m]", True),
        ("[mm]", True),
        (r"[Blue]\+[h]:mm;[Red]\-[h]:mm;[Green][h]:mm", True),
        ("[>=100][Magenta][s].00", True),
        (r"[h]:mm;[=0]\-", True),
        ("[>=100][Magenta].00", False),
        ("[>=100][Magenta]General", False),
        (r"ha/p\\m", True),
        (r'#,##0.00\ _M"H"_);[Red]#,##0.00\ _M"S"_)', False),
    ],
)
def test_is_date_format(format, result):
    from ..numbers import is_date_format

    assert is_date_format(format) is result


@pytest.mark.parametrize(
    "format, result",
    [
        ("m:ss", False),
        ("[h]", True),
        ("[hh]", True),
        ("[h]:mm:ss", True),
        ("[hh]:mm:ss", True),
        ("[h]:mm:ss.000", True),
        ("[hh]:mm:ss.0", True),
        ("[h]:mm", True),
        ("[hh]:mm", True),
        ("[m]:ss", True),
        ("[mm]:ss", True),
        ("[m]:ss.000", True),
        ("[mm]:ss.0", True),
        ("[s]", True),
        ("[ss]", True),
        ("[s].000", True),
        ("[ss].0", True),
        ("[m]", True),
        ("[mm]", True),
        ("h:mm", False),
        (r"[Blue]\+[h]:mm;[Red]\-[h]:mm;[h]:mm", True),
        (r"[Blue]\+[h]:mm;[Red]\-[h]:mm;[Green][h]:mm", True),
        ("[>=100][Magenta][s].00", True),
        (r"[h]:mm;[=0]\-", True),
        ("[>=100][Magenta].00", False),
        ("[>=100][Magenta]General", False),
    ],
)
def test_is_timedelta_format(format, result):
    from ..numbers import is_timedelta_format

    assert is_timedelta_format(format) is result


@pytest.mark.parametrize(
    "fmt, typ",
    [
        (FORMAT_DATE_DATETIME, "datetime"),
        (FORMAT_DATE_DDMMYY, "date"),
        (FORMAT_DATE_DMMINUS, "date"),
        (FORMAT_DATE_DMMINUS, "date"),
        (FORMAT_DATE_DMYSLASH, "date"),
        (FORMAT_DATE_MYMINUS, "date"),
        (FORMAT_DATE_TIME1, "time"),
        (FORMAT_DATE_TIME2, "time"),
        (FORMAT_DATE_TIME3, "time"),
        (FORMAT_DATE_TIME4, "time"),
        (FORMAT_DATE_TIME5, "time"),
        (FORMAT_DATE_TIME6, "time"),
        (FORMAT_DATE_TIME7, "time"),
        (FORMAT_DATE_TIME8, "time"),
        (FORMAT_DATE_TIMEDELTA, "time"),
        (FORMAT_DATE_XLSX14, "date"),
        (FORMAT_DATE_XLSX15, "date"),
        (FORMAT_DATE_XLSX16, "date"),
        (FORMAT_DATE_XLSX17, "date"),
        (FORMAT_DATE_XLSX22, "datetime"),
        (FORMAT_DATE_YYMMDD, "date"),
        (FORMAT_DATE_YYMMDDSLASH, "date"),
        (FORMAT_DATE_YYYYMMDD2, "date"),
    ],
)
def test_datetime(fmt, typ):
    from ..numbers import is_datetime

    assert is_datetime(fmt) == typ
