# Copyright (c) 2010-2025 openpyxl
import datetime
import math
import sys

from openpyxl.compat.numbers import NUMERIC_TYPES

VER = sys.version_info


def safe_string(value):
    """Safely and consistently format numeric values"""
    if isinstance(value, NUMERIC_TYPES):
        if math.isnan(value) or math.isinf(value):
            value = ""
        else:
            value = f"{value:.16g}"
    elif value is None:
        value = "none"
    elif isinstance(value, datetime.datetime):
        value = value.isoformat()
    elif not isinstance(value, str):
        value = str(value)
    return value
