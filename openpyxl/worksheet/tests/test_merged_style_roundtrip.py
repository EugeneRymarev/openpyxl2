"""Public-API and XLSX round-trip regressions for upstream issue #2024."""
from copy import copy
from io import BytesIO

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Protection, Side
from openpyxl.worksheet.cell_range import CellRange


def roundtrip(workbook, read_only=False):
    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return load_workbook(stream, read_only=read_only)


def border_style(cell, side):
    border = getattr(cell.border, side, None)
    return border.style if border is not None else None


def assert_perimeter(ws, ref, border):
    cr = CellRange(ref)
    for row, col in cr.cells:
        cell = ws.cell(row, col)
        edges = {"top": row == cr.min_row, "bottom": row == cr.max_row,
                 "left": col == cr.min_col, "right": col == cr.max_col}
        for side, on_edge in edges.items():
            # The anchor is the canonical definition of all four outer sides.
            if (row, col) == (cr.min_row, cr.min_col):
                on_edge = True
            expected_side = getattr(border, side) if on_edge else None
            expected = expected_side.style if expected_side is not None else None
            assert border_style(cell, side) == expected, (cell.coordinate, side)
            if expected:
                assert getattr(cell.border, side).color == expected_side.color


def outline(style="thin", color="FFFF0000"):
    side = Side(style=style, color=color)
    return Border(left=side, right=side, top=side, bottom=side)


@pytest.mark.parametrize("ref", ["A1:B3", "B2:D2", "B2:B4", "B2:D4", "B2"])
@pytest.mark.parametrize("when", ["before", "after", "named"])
def test_merged_border_perimeter_and_roundtrip(ref, when):
    wb = Workbook()
    ws = wb.active
    anchor = ws.cell(CellRange(ref).min_row, CellRange(ref).min_col)
    border = outline()
    if when == "before":
        anchor.border = border
    ws.merge_cells(ref)
    anchor.value = "Test"
    if when == "after":
        anchor.border = border
    elif when == "named":
        anchor.style = NamedStyle(name="outlined", border=border)
    assert_perimeter(ws, ref, border)
    for _ in range(2):
        wb = roundtrip(wb)
        ws = wb.active
        assert ws[anchor.coordinate].value == "Test"
        assert_perimeter(ws, ref, border)
    streamed = roundtrip(wb, read_only=True)
    try:
        assert_perimeter(streamed.active, ref, border)
    finally:
        streamed.close()


@pytest.mark.parametrize("update", ["replace", "clear", "delete", "named"])
def test_merged_border_updates_remove_old_edges(update):
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("A1:C3")
    ws["A1"].border = outline()
    expected = Border(right=Side(style="double", color="FF00FF00"))
    if update == "replace":
        ws["A1"].border = expected
    elif update == "named":
        ws["A1"].style = NamedStyle(name="replacement", border=expected)
    else:
        expected = Border()
        if update == "clear":
            ws["A1"].border = expected
        else:
            del ws["A1"].border
    assert_perimeter(ws, "A1:C3", expected)
    assert_perimeter(roundtrip(wb).active, "A1:C3", expected)


def test_loading_preserves_individual_merged_cell_styles():
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("A1:C3")
    ws["A1"].border = outline()
    ws["A1"].fill = PatternFill("solid", fgColor="FFFF0000")
    ws["A1"].font = Font(name="Arial", bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].number_format = "0.00"
    ws["A1"].protection = Protection(locked=False)
    # An existing workbook may have explicit formatting on a non-anchor cell.
    ws["B2"].fill = PatternFill("solid", fgColor="FF00FF00")
    ws["B2"].protection = Protection(locked=True, hidden=True)
    properties = ("font", "fill", "alignment", "number_format", "protection")
    expected = {cell.coordinate: {key: copy(getattr(cell, key)) for key in properties}
                for row in ws for cell in row}
    for _ in range(2):
        wb = roundtrip(wb)
        for coord, styles in expected.items():
            for key, value in styles.items():
                assert getattr(wb.active[coord], key) == value, (coord, key)
        assert_perimeter(wb.active, "A1:C3", outline())


def test_unmerged_cell_and_explicit_placeholder_border_are_local():
    ws = Workbook().active
    ws.merge_cells("A1:B2")
    ws["A1"].border = outline()
    expected = copy(ws["A1"].border)
    ws["B2"].border = Border(bottom=Side(style="double"))
    ws["D4"].border = outline("thick")
    assert ws["A1"].border == expected
    assert ws["B2"].border.bottom.style == "double"


@pytest.mark.parametrize("by_name", [False, True])
def test_named_style_preserves_other_properties_when_border_is_projected(by_name):
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("A1:C3")
    style = NamedStyle(name="filled", border=outline(),
                       fill=PatternFill("solid", fgColor="FF0000FF"),
                       font=Font(bold=True), number_format="0.00")
    wb.add_named_style(style)
    ws["A1"].style = style.name if by_name else style
    loaded = roundtrip(wb).active
    for row in loaded:
        for cell in row:
            assert cell.fill == style.fill
            assert cell.font == style.font
            assert cell.number_format == style.number_format
    assert_perimeter(loaded, "A1:C3", outline())


def test_issue_2024_original_example():
    wb = Workbook()
    ws = wb.active
    side = Side(style="thin", color="000000")
    border = Border(left=side, top=side, right=side, bottom=side,
                    vertical=side, horizontal=side)
    ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=2)
    ws.cell(row=1, column=1, value="Test").border = border
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=3)
    ws.cell(row=4, column=1, value="Test2").border = border
    for read_only in (False, True):
        loaded = roundtrip(wb, read_only=read_only)
        try:
            assert_perimeter(loaded.active, "A1:B3", border)
            assert_perimeter(loaded.active, "A4:C4", border)
        finally:
            loaded.close()
