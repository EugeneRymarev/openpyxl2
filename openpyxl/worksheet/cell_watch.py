from openpyxl.descriptors import Sequence
from openpyxl.descriptors.base import String
from openpyxl.descriptors.serialisable import Serialisable

# could be done using a nestedSequence


class CellWatch(Serialisable):
    tagname = "cellWatch"

    r = String()

    def __init__(self, r=None):
        self.r = r


class CellWatches(Serialisable):
    tagname = "cellWatches"

    cellWatch = Sequence(expected_type=CellWatch)

    __elements__ = ("cellWatch",)

    def __init__(self, cellWatch=()):
        self.cellWatch = cellWatch
