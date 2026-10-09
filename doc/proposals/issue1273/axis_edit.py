"""Executable design sketch for #1273, NOT integrated into Worksheet.

A pure one-dimensional transform for inclusive references. It deliberately
knows nothing about XML, formulas, merged cells, or workbook object identity.
"""
from dataclasses import dataclass
from typing import Optional, Tuple


@dataclass(frozen=True)
class AxisEdit:
    index: int
    count: int = 1
    deleting: bool = False
    limit: int = 1048576

    def __post_init__(self):
        for name in ("index", "count", "limit"):
            if type(getattr(self, name)) is not int:
                raise TypeError(name + " must be an integer")
        if type(self.deleting) is not bool:
            raise TypeError("deleting must be a bool")
        if self.limit < 1 or not 1 <= self.index <= self.limit:
            raise ValueError("index is outside the worksheet axis")
        if not 0 <= self.count <= self.limit:
            raise ValueError("invalid count")
        if self.deleting and self.index + self.count - 1 > self.limit:
            raise ValueError("deletion is outside the worksheet axis")

    def _check_coordinate(self, value):
        if type(value) is not int:
            raise TypeError("coordinate must be an integer")
        if not 1 <= value <= self.limit:
            raise ValueError("coordinate is outside the worksheet axis")

    def point(self, value: int) -> Optional[int]:
        self._check_coordinate(value)
        if not self.count or value < self.index:
            return value
        if self.deleting:
            if value < self.index + self.count:
                return None
            return value - self.count
        result = value + self.count
        if result > self.limit:
            raise OverflowError("insertion moves a reference outside the worksheet")
        return result

    def interval(self, start: int, end: int) -> Optional[Tuple[int, int]]:
        self._check_coordinate(start)
        self._check_coordinate(end)
        if start > end:
            raise ValueError("inverted interval")
        if not self.count:
            return start, end
        if not self.deleting:
            return self.point(start), self.point(end)
        if end < self.index:
            return start, end
        deletion_end = self.index + self.count - 1
        if start > deletion_end:
            return start - self.count, end - self.count
        first = start if start < self.index else self.index
        last = end - self.count if end > deletion_end else self.index - 1
        if first > last:
            return None
        return first, last
