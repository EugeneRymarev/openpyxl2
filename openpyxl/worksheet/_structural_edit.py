"""Plan row/column edits and synchronize the metadata covered by issue #1273.

Formula text is deliberately never rewritten. All fallible transformations are
prepared before the worksheet, its cells, or workbook registries are changed.
"""

import copy
import re
from dataclasses import dataclass
from typing import Optional
from typing import Tuple

from openpyxl.cell.cell import Cell
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.coordinate import Coordinate
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils.cell import column_index_from_string
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.dimensions import DimensionHolder
from openpyxl.worksheet.merge import MergedCellRange
from openpyxl.worksheet.print_settings import ColRange
from openpyxl.worksheet.print_settings import PrintArea
from openpyxl.worksheet.print_settings import RowRange


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


_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)\Z")
_COL = re.compile(r"(\$?)([A-Za-z]{1,3})\Z")
_ROW = re.compile(r"(\$?)([1-9][0-9]*)\Z")


def _range(edit, axis, cr):
    bounds = list(cr.bounds)
    first = 1 if axis == "row" else 0
    interval = edit.interval(bounds[first], bounds[first + 2])
    if interval is None:
        return None
    bounds[first], bounds[first + 2] = interval
    return CellRange(
        min_col=bounds[0],
        min_row=bounds[1],
        max_col=bounds[2],
        max_row=bounds[3],
        title=cr.title,
    )


def _reference(value):
    """Parse only a literal A1 reference; return None for expressions/aliases."""
    prefix, bang, coordinate = value.rpartition("!")
    if not bang:
        coordinate = value
    if coordinate == "#REF!":
        return prefix, bang, coordinate, None, None
    parts = coordinate.split(":")
    if len(parts) > 2:
        return None
    patterns = [_CELL] if len(parts) == 1 else [_CELL, _COL, _ROW]
    for pattern in patterns:
        matches = [pattern.fullmatch(part) for part in parts]
        if all(matches):
            return prefix, bang, coordinate, pattern, [m.groups() for m in matches]
    return None


def _rewrite_reference(value, parsed, owner, ws, axis, edit):
    prefix, bang, coordinate, pattern, parts = parsed
    title = prefix if bang else owner
    if bang and title.startswith("'") and title.endswith("'"):
        title = title[1:-1].replace("''", "'")
    if title is None or "[" in title or ":" in title:
        return value  # external, 3D, or context-dependent global reference
    if title.casefold() != ws.title.casefold() or pattern is None:
        return value
    if (pattern is _COL and axis == "row") or (pattern is _ROW and axis == "column"):
        return value  # the orthogonal axis is unbounded, not a finite rectangle
    component = 3 if pattern is _CELL and axis == "row" else 1
    values = [
        int(p[component]) if axis == "row" else column_index_from_string(p[component])
        for p in parts
    ]
    result = edit.interval(values[0], values[-1])
    if result is None:
        return "#REF!"
    if result == (values[0], values[-1]):
        return value
    out = []
    for part, position in zip(parts, (result[0], result[1])):
        part = list(part)
        replacement = str(position) if axis == "row" else get_column_letter(position)
        if axis == "column" and part[component].islower():
            replacement = replacement.lower()
        part[component] = replacement
        out.append("".join(part))
    return (prefix + bang if bang else "") + ":".join(out)


def _name_value(value, owner, ws, axis, edit):
    if not isinstance(value, str):
        return value
    leading_equals = value.startswith("=")
    try:
        tokens = Tokenizer(value if leading_equals else "=" + value).items
    except TokenizerError:
        return value
    significant = [t for t in tokens if t.type != "WHITE-SPACE"]
    if not significant or len(significant) % 2 == 0:
        return value
    references = {}
    for index, token in enumerate(significant):
        if index % 2:
            if token.value != "," or token.type != "OPERATOR-INFIX":
                return value
        else:
            if token.type != "OPERAND" or token.subtype not in ("RANGE", "ERROR"):
                return value
            parsed = _reference(token.value)
            if parsed is None:
                return value  # the entire expression is outside the literal subset
            references[id(token)] = parsed
    result = []
    for token in tokens:
        parsed = references.get(id(token))
        result.append(
            _rewrite_reference(token.value, parsed, owner, ws, axis, edit)
            if parsed is not None
            else token.value
        )
    return ("=" if leading_equals else "") + "".join(result)


def _names(ws, axis, edit):
    wb = ws.parent
    scopes = [(getattr(wb, "defined_names", {}), None)]
    sheets = list(getattr(wb, "worksheets", ()))
    if ws not in sheets:
        sheets.append(ws)
    scopes.extend((sheet.defined_names, sheet.title) for sheet in sheets)
    changes = []
    for definitions, owner in scopes:
        for name in definitions.values():
            value = _name_value(name.attr_text, owner, ws, axis, edit)
            if value != name.attr_text:
                changes.append((name, value, name.attr_text))
    return changes


def _copy_dimension(dimension):
    # Dimension.__copy__ currently mutates the source __dict__; avoid it while
    # planning an operation that must leave the original unchanged on failure.
    result = dimension.__class__.__new__(dimension.__class__)
    result.__dict__.update(dimension.__dict__)
    result.parent = dimension.parent
    result._style = copy.copy(dimension._style)
    return result


def _dimensions(ws, axis, edit):
    source = ws.row_dimensions if axis == "row" else ws.column_dimensions
    result = DimensionHolder(
        ws, reference=source.reference, default_factory=source.default_factory
    )
    result.max_outline = source.max_outline
    spans = []
    for key, dimension in source.items():
        if axis == "row":
            position = edit.point(key)
            if position is None:
                continue
            dim = _copy_dimension(dimension)
            dim.index = position
            result[position] = dim
        else:
            start = dimension.min or column_index_from_string(key)
            end = dimension.max or start
            interval = edit.interval(start, end)
            if interval is None:
                continue
            dim = _copy_dimension(dimension)
            dim.min, dim.max = interval
            dim.index = get_column_letter(dim.min)
            spans.append(interval)
            if dim.index in result:
                raise ValueError("Column dimension groups overlap after the edit")
            result[dim.index] = dim
    spans.sort()
    if any(left[1] >= right[0] for left, right in zip(spans, spans[1:])):
        raise ValueError("Overlapping column dimension groups are not supported")
    return result


class _EditPlan:
    def __init__(self, ws, axis, edit):
        self.ws = ws
        self.axis = axis
        self.cells = {}
        self.styles = []
        self.borders = None
        index = 0 if axis == "row" else 1
        for coordinate, cell in ws._cells.items():
            position = edit.point(coordinate[index])
            if position is not None:
                target = list(coordinate)
                target[index] = position
                self.cells[tuple(target)] = cell

        # Bound/reference objects are replaced, never mutated while hashed.
        self.merges = self._merges(edit)
        self.area = PrintArea(
            [
                cr
                for original in ws._print_area
                if (cr := _range(edit, axis, original)) is not None
            ]
        )
        self.rows = ws._print_rows
        self.cols = ws._print_cols
        if axis == "row" and self.rows is not None:
            span = edit.interval(self.rows.min_row, self.rows.max_row)
            self.rows = RowRange(min_row=span[0], max_row=span[1]) if span else None
        elif axis == "column" and self.cols is not None:
            span = edit.interval(
                column_index_from_string(self.cols.min_col),
                column_index_from_string(self.cols.max_col),
            )
            self.cols = (
                ColRange(
                    min_col=get_column_letter(span[0]),
                    max_col=get_column_letter(span[1]),
                )
                if span
                else None
            )
        self.dimensions = _dimensions(ws, axis, edit)
        self.names = _names(ws, axis, edit)
        self.coordinates = [
            (cell, Coordinate(*coordinate), cell._coord)
            for coordinate, cell in self.cells.items()
        ]
        self.current_row = max((row for row, col in self.cells), default=0)

    def _merges(self, edit):
        ws = self.ws
        ranges = []
        claimed = set()
        for original in ws.merged_cells:
            cr = _range(edit, self.axis, original)
            if cr is None:
                continue
            anchor = ws._cells.get((original.min_row, original.min_col))
            if not isinstance(anchor, Cell):
                raise ValueError("A merged range must have a normal anchor cell")
            resized = cr.size != original.size
            start = (cr.min_row, cr.min_col)
            # Keep the original anchor when deletion removes its row/column but
            # leaves part of the merge. Whole-merge deletion removes it normally.
            self.cells[start] = anchor
            projections = {}
            for coordinate in cr.cells:
                if coordinate in claimed:
                    raise ValueError("Overlapping merged ranges are not supported")
                claimed.add(coordinate)
                if coordinate == start:
                    continue
                cell = self.cells.get(coordinate)
                if cell is None:
                    cell = MergedCell(ws, *coordinate)
                    cell._style = copy.copy(anchor._style)
                    self.cells[coordinate] = cell
                if not isinstance(cell, MergedCell):
                    raise ValueError(
                        "Structural edit would overwrite a non-merged cell"
                    )
                if resized:
                    if self.borders is None:
                        self.borders = IndexedList(ws.parent._borders)
                    edges = (
                        coordinate[0] == cr.min_row,
                        coordinate[0] == cr.max_row,
                        coordinate[1] == cr.min_col,
                        coordinate[1] == cr.max_col,
                    )
                    if edges not in projections:
                        border_id = (
                            anchor._style.borderId if anchor._style is not None else 0
                        )
                        border = copy.copy(self.borders[border_id])
                        for side, on_edge in zip(
                            ("top", "bottom", "left", "right"), edges
                        ):
                            if not on_edge:
                                setattr(border, side, None)
                        projections[edges] = self.borders.add(border)
                    style = (
                        copy.copy(cell._style)
                        if cell._style is not None
                        else StyleArray()
                    )
                    style.borderId = projections[edges]
                    self.styles.append((cell, style, cell._style))
            # MergedCellRange.__init__ reads/mutates ws. This prevalidated object
            # is attached only at commit, after the planned cells are installed.
            merged = MergedCellRange.__new__(MergedCellRange)
            CellRange.__init__(merged, range_string=cr.coord)
            merged.ws = ws
            merged.start_cell = anchor
            ranges.append(merged)
        return MultiCellRange(ranges)

    def apply(self):
        ws = self.ws
        dimension_attr = "row_dimensions" if self.axis == "row" else "column_dimensions"
        replacements = {
            "_cells": self.cells,
            "merged_cells": self.merges,
            "_print_area": self.area,
            "_print_rows": self.rows,
            "_print_cols": self.cols,
            dimension_attr: self.dimensions,
            "_current_row": self.current_row,
        }
        previous = {name: getattr(ws, name) for name in replacements}
        old_borders = ws.parent._borders if self.borders is not None else None
        try:
            for cell, coordinate, old in self.coordinates:
                cell._coord = coordinate
            for cell, style, old in self.styles:
                cell._style = style
            if self.borders is not None:
                ws.parent._borders = self.borders
            for name, value in replacements.items():
                setattr(ws, name, value)
            for name, value, old in self.names:
                name.attr_text = value
        except Exception:
            for cell, coordinate, old in self.coordinates:
                cell._coord = old
            for cell, style, old in self.styles:
                cell._style = old
            if self.borders is not None:
                ws.parent._borders = old_borders
            for name, value in previous.items():
                setattr(ws, name, value)
            for name, value, old in self.names:
                name.attr_text = old
            raise


def edit_worksheet(ws, axis, index, amount, deleting=False):
    limit = 1048576 if axis == "row" else 16384
    edit = AxisEdit(index, amount, deleting=deleting, limit=limit)
    if amount:
        _EditPlan(ws, axis, edit).apply()
