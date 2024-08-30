# Copyright (c) 2010-2024 openpyxl
from .data_source import NumDataSource
from .shapes import GraphicalProperties
from openpyxl.descriptors import Alias
from openpyxl.descriptors import Typed
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import NestedBool
from openpyxl.descriptors.nested import NestedFloat
from openpyxl.descriptors.nested import NestedNoneSet
from openpyxl.descriptors.nested import NestedSet
from openpyxl.descriptors.serialisable import Serialisable


class ErrorBars(Serialisable):
    tagname = "errBars"

    errDir = NestedNoneSet(values=(["x", "y"]))
    direction = Alias("errDir")
    errBarType = NestedSet(values=(["both", "minus", "plus"]))
    style = Alias("errBarType")
    errValType = NestedSet(
        values=(["cust", "fixedVal", "percentage", "stdDev", "stdErr"])
    )
    size = Alias("errValType")
    noEndCap = NestedBool(nested=True, allow_none=True)
    plus = Typed(expected_type=NumDataSource, allow_none=True)
    minus = Typed(expected_type=NumDataSource, allow_none=True)
    val = NestedFloat(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias("spPr")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = (
        "errDir",
        "errBarType",
        "errValType",
        "noEndCap",
        "minus",
        "plus",
        "val",
        "spPr",
    )

    def __init__(
        self,
        errDir=None,
        errBarType="both",
        errValType="fixedVal",
        noEndCap=None,
        plus=None,
        minus=None,
        val=None,
        spPr=None,
        extLst=None,
    ):
        self.errDir = errDir
        self.errBarType = errBarType
        self.errValType = errValType
        self.noEndCap = noEndCap
        self.plus = plus
        self.minus = minus
        self.val = val
        self.spPr = spPr
