# Copyright (c) 2010-2024 openpyxl
import sys
from datetime import datetime
from math import isinf
from math import isnan

VER = sys.version_info

from .numbers import NUMERIC_TYPES


def safe_string(value):
    """Safely and consistently format numeric values"""
    if isinstance(value, NUMERIC_TYPES):
        if isnan(value) or isinf(value):
            value = ""
        else:
            value = f"{value:.16g}"
    elif value is None:
        value = "none"
    elif isinstance(value, datetime):
        value = value.isoformat()
    elif not isinstance(value, str):
        value = str(value)
    return value
