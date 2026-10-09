# Copyright (c) 2010-2025 openpyxl
import copy

from openpyxl.cell.cell import MergedCell
from openpyxl.descriptors.base import Integer
from openpyxl.descriptors.sequence import Sequence
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.styles.borders import Border
from openpyxl.worksheet.cell_range import CellRange


class MergeCell(CellRange):
    tagname = "mergeCell"
    ref = CellRange.coord
    __attrs__ = ("ref",)

    def __init__(self, ref=None):
        super().__init__(ref)

    def __copy__(self):
        return self.__class__(self.ref)


class MergeCells(Serialisable):
    tagname = "mergeCells"
    count = Integer(allow_none=True)
    mergeCell = Sequence(expected_type=MergeCell)
    __elements__ = ("mergeCell",)
    __attrs__ = ("count",)

    def __init__(self, count=None, mergeCell=()):
        self.mergeCell = mergeCell

    @property
    def count(self):
        return len(self.mergeCell)


class MergedCellRange(CellRange):
    """
    MergedCellRange stores the border information of a merged cell in the top
    left cell of the merged cell.
    The remaining cells in the merged cell are stored as MergedCell objects and
    get their border information from the upper left cell.
    """

    def __init__(self, worksheet, coord):
        self.ws = worksheet
        super().__init__(range_string=coord)
        self.start_cell = None
        self._get_borders()

    def _get_borders(self):
        """
        If the upper left cell of the merged cell does not yet exist, it is
        created.
        The upper left cell gets the border information of the bottom and right
        border from the bottom right cell of the merged cell, if available.
        """
        # Top-left cell.
        self.start_cell = self.ws._cells.get((self.min_row, self.min_col))
        if self.start_cell is None:
            self.start_cell = self.ws.cell(row=self.min_row, column=self.min_col)
        # Bottom-right cell
        end_cell = self.ws._cells.get((self.max_row, self.max_col))
        if end_cell is not None:
            border = Border(right=end_cell.border.right, bottom=end_cell.border.bottom)
            self.start_cell._style = copy.copy(self.start_cell._style)
            self.start_cell._border += border

    def _apply_border(self, border):
        """Replace the outline after an explicit assignment to the anchor.

        Keep the complete definition on the anchor and project its four sides
        onto the corresponding perimeter cells. Replacing (not adding) sides
        also permits clearing a previously assigned border.
        """
        projections = {}
        for row, column in self.cells:
            if (row, column) == (self.min_row, self.min_col):
                continue
            cell = self.ws._cells.get((row, column))
            if cell is None:
                cell = MergedCell(self.ws, row=row, column=column)
                self.ws._cells[(row, column)] = cell
            edges = (row == self.min_row, row == self.max_row,
                     column == self.min_col, column == self.max_col)
            if edges not in projections:
                projected = copy.copy(border)
                for name, on_edge in zip(("top", "bottom", "left", "right"), edges):
                    if not on_edge:
                        setattr(projected, name, None)
                projections[edges] = projected
            projected = projections[edges]
            cell._border = projected

    def format(self, preserve_styles=False):
        """
        Each cell of the merged cell is created as MergedCell if it does not
        already exist.

        The MergedCells at the edge of the merged cell gets its borders from
        the upper left cell.

         - The top MergedCells get the top border from the top left cell.
         - The bottom MergedCells get the bottom border from the top left cell.
         - The left MergedCells get the left border from the top left cell.
         - The right MergedCells get the right border from the top left cell.
        """
        # Loading must retain explicit per-cell protection as well as other
        # stored styles. Capture this before border access creates style arrays.
        styled = set()
        if preserve_styles:
            styled = {
                coord for coord in self.cells
                if self.ws._cells.get(coord) is not None
                and self.ws._cells[coord]._style is not None
            }
        names = ["top", "left", "right", "bottom"]
        for name in names:
            side = getattr(self.start_cell.border, name)
            if side and side.style is None:
                # don't need to do anything if there is no border style
                continue
            border = Border(**{name: side})
            for coord in getattr(self, name):
                cell = self.ws._cells.get(coord)
                if cell is None:
                    row, col = coord
                    cell = MergedCell(self.ws, row=row, column=col)
                    self.ws._cells[(cell.row, cell.column)] = cell
                cell._border += border
        protected = self.start_cell.protection is not None
        protection = None
        if protected:
            protection = copy.copy(self.start_cell.protection)
        for coord in self.cells:
            cell = self.ws._cells.get(coord)
            if cell is None:
                row, col = coord
                cell = MergedCell(self.ws, row=row, column=col)
                self.ws._cells[(cell.row, cell.column)] = cell
            if protected and coord not in styled:
                cell._protection = protection

    def __contains__(self, coord):
        return coord in CellRange(self.coord)

    def __copy__(self):
        return self.__class__(self.ws, self.coord)
