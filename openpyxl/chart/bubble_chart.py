from ._chart import ChartBase
from .axis import NumericAxis
from .label import DataLabelList
from .series import XYSeries
from openpyxl.descriptors import Alias
from openpyxl.descriptors import Sequence
from openpyxl.descriptors import Typed
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import NestedBool
from openpyxl.descriptors.nested import NestedMinMax
from openpyxl.descriptors.nested import NestedNoneSet


class BubbleChart(ChartBase):
    tagname = "bubbleChart"

    varyColors = NestedBool(allow_none=True)
    ser = Sequence(expected_type=XYSeries, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")
    bubble3D = NestedBool(allow_none=True)
    bubbleScale = NestedMinMax(min=0, max=300, allow_none=True)
    showNegBubbles = NestedBool(allow_none=True)
    sizeRepresents = NestedNoneSet(values=(["area", "w"]))
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=NumericAxis)
    y_axis = Typed(expected_type=NumericAxis)

    _series_type = "bubble"

    __elements__ = (
        "varyColors",
        "ser",
        "dLbls",
        "bubble3D",
        "bubbleScale",
        "showNegBubbles",
        "sizeRepresents",
        "axId",
    )

    def __init__(
        self,
        varyColors=None,
        ser=(),
        dLbls=None,
        bubble3D=None,
        bubbleScale=None,
        showNegBubbles=None,
        sizeRepresents=None,
        extLst=None,
        **kw
    ):
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.bubble3D = bubble3D
        self.bubbleScale = bubbleScale
        self.showNegBubbles = showNegBubbles
        self.sizeRepresents = sizeRepresents
        self.x_axis = NumericAxis(axId=10, crossAx=20)
        self.y_axis = NumericAxis(axId=20, crossAx=10)
        super().__init__(**kw)
