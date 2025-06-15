# Copyright (c) 2010-2025 openpyxl
"""
OOXML has non-standard escaping for characters < \031
"""
import re


def escape(value):
    r"""
    Convert ASCII < 31 to OOXML: \n == _x + hex(ord(\n)) + _
    """
    char_regex = re.compile(r"[\001-\031]")

    def _sub(match):
        """
        Callback to escape chars
        """
        return f"_x{ord(match.group(0)):0>4x}_"

    return char_regex.sub(_sub, value)


def unescape(value):
    r"""
    Convert escaped strings to ASCIII: _x000a_ == \n
    """
    escaped_regex = re.compile("_x([0-9A-Fa-f]{4})_")

    def _sub(match):
        """
        Callback to unescape chars
        """
        return chr(int(match.group(1), 16))

    if "_x" in value:
        value = escaped_regex.sub(_sub, value)
    return value
