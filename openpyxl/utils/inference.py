# Copyright (c) 2010-2024 openpyxl
"""
Type inference functions
"""
import datetime
import re

from openpyxl.styles import numbers

PERCENT_REGEX = re.compile(r"^(?P<number>\-?[0-9]*\.?[0-9]*\s?)\%$")
pattern1 = r"""
^(?: # HH:MM and HH:MM:SS
(?P<hour>[0-1]{0,1}[0-9]{2}):
(?P<minute>[0-5][0-9]):?
(?P<second>[0-5][0-9])?$)
|
^(?: # MM:SS.
([0-5][0-9]):
([0-5][0-9])?\.
(?P<microsecond>\d{1,6}))
"""
TIME_REGEX = re.compile(pattern1, re.VERBOSE)
pattern2 = r"^-?([\d]|[\d]+\.[\d]*|\.[\d]+|[1-9][\d]+\.?[\d]*)((E|e)[-+]?[\d]+)?$"
NUMBER_REGEX = re.compile(pattern2)


def cast_numeric(value):
    """Explicitly convert a string to a numeric value"""
    if NUMBER_REGEX.match(value):
        try:
            return int(value)
        except ValueError:
            return float(value)


def cast_percentage(value):
    """Explicitly convert a string to numeric value and format as a
    percentage"""
    match = PERCENT_REGEX.match(value)
    if match:
        return float(match.group("number")) / 100


def cast_time(value):
    """Explicitly convert a string to a number and format as datetime or
    time"""
    match = TIME_REGEX.match(value)
    if match:
        if match.group("microsecond") is not None:
            value = value[:12]
            pattern = "%M:%S.%f"
        elif match.group("second") is None:
            pattern = "%H:%M"
        else:
            pattern = "%H:%M:%S"
        value = datetime.datetime.strptime(value, pattern)
        return value.time()
