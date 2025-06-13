# Copyright (c) 2010-2025 openpyxl
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.descriptors import NestedGapAmount
from openpyxl.descriptors.base import Typed
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable


class UpDownBars(Serialisable):
    tagname = "upbars"
    gapWidth = NestedGapAmount()
    upBars = Typed(expected_type=ChartLines, allow_none=True)
    downBars = Typed(expected_type=ChartLines, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    __elements__ = ("gapWidth", "upBars", "downBars")

    def __init__(self, gapWidth=150, upBars=None, downBars=None, extLst=None):
        self.gapWidth = gapWidth
        self.upBars = upBars
        self.downBars = downBars
