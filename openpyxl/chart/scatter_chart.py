# Copyright (c) 2010-2024 openpyxl
from ._chart import ChartBase
from .axis import NumericAxis
from .axis import TextAxis
from .label import DataLabelList
from .series import XYSeries
from openpyxl.descriptors import Alias
from openpyxl.descriptors import Sequence
from openpyxl.descriptors import Typed
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import NestedBool
from openpyxl.descriptors.nested import NestedNoneSet


class ScatterChart(ChartBase):
    tagname = "scatterChart"

    scatterStyle = NestedNoneSet(
        values=(["line", "lineMarker", "marker", "smooth", "smoothMarker"])
    )
    varyColors = NestedBool(allow_none=True)
    ser = Sequence(expected_type=XYSeries, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=(NumericAxis, TextAxis))
    y_axis = Typed(expected_type=NumericAxis)

    _series_type = "scatter"

    __elements__ = ("scatterStyle", "varyColors", "ser", "dLbls", "axId")

    def __init__(
        self,
        scatterStyle=None,
        varyColors=None,
        ser=(),
        dLbls=None,
        extLst=None,
        **kw,
    ):
        self.scatterStyle = scatterStyle
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.x_axis = NumericAxis(axId=10, crossAx=20)
        self.y_axis = NumericAxis(axId=20, crossAx=10)
        super().__init__(**kw)
