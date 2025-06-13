# Copyright (c) 2010-2025 openpyxl
"""
Utility descriptors for the chart module.
For convenience but also clarity.
"""
from openpyxl.chart.data_source import NumFmt
from openpyxl.descriptors.base import Typed
from openpyxl.descriptors.nested import NestedMinMax


class NestedGapAmount(NestedMinMax):
    allow_none = True
    min = 0
    max = 500


class NestedOverlap(NestedMinMax):
    allow_none = True
    min = -100
    max = 100


class NumberFormatDescriptor(Typed):
    """
    Allow direct assignment of format code
    """

    expected_type = NumFmt
    allow_none = True

    def __set__(self, instance, value):
        if isinstance(value, str):
            value = NumFmt(value)
        super().__set__(instance, value)
