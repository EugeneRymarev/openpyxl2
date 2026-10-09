"""Boundary, persistence and failure-atomicity checks for structural edits."""

from io import BytesIO

import pytest

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Border
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet


def roundtrip(wb):
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return load_workbook(stream)


@pytest.mark.parametrize(
    "operation,index,amount,expected",
    [
        ("insert_rows", 4, 2, "B3:D7"),
        ("insert_rows", 3, 2, "B5:D7"),
        ("insert_rows", 6, 2, "B3:D5"),
        ("delete_rows", 3, 1, "B3:D4"),
        ("delete_rows", 4, 1, "B3:D4"),
        ("delete_rows", 5, 1, "B3:D4"),
        ("delete_rows", 2, 3, "B2:D2"),
        ("delete_rows", 3, 3, ""),
        ("delete_rows", 1, 6, ""),
        ("insert_cols", 3, 2, "B3:F5"),
        ("insert_cols", 2, 2, "D3:F5"),
        ("insert_cols", 5, 2, "B3:D5"),
        ("delete_cols", 2, 1, "B3:C5"),
        ("delete_cols", 3, 1, "B3:C5"),
        ("delete_cols", 4, 1, "B3:C5"),
        ("delete_cols", 1, 3, "A3:A5"),
        ("delete_cols", 2, 3, ""),
        ("delete_cols", 1, 5, ""),
    ],
)
def test_resize_merge_preserves_anchor_and_outline(operation, index, amount, expected):
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("B3:D5")
    ws["B3"] = "heading"
    anchor = ws["B3"]
    side = Side(style="thin", color="FF123456")
    anchor.border = Border(top=side, left=side, right=side, bottom=side)
    anchor.font = Font(bold=True)
    getattr(ws, operation)(index, amount)
    assert str(ws.merged_cells) == expected
    if not expected:
        assert not ws._cells
        assert ws._current_row == 0
        assert not roundtrip(wb).active.merged_cells
        return
    cr = CellRange(expected)
    assert ws.cell(cr.min_row, cr.min_col) is anchor
    for _ in range(2):
        ws = wb.active
        merged = next(iter(ws.merged_cells))
        assert merged.start_cell is ws.cell(cr.min_row, cr.min_col)
        assert merged.start_cell.value == "heading"
        assert merged.start_cell.font.bold
        for row, col in cr.cells:
            cell = ws.cell(row, col)
            if (row, col) == (cr.min_row, cr.min_col):
                continue
            assert isinstance(cell, MergedCell)
            for name, on_edge in [
                ("top", row == cr.min_row),
                ("bottom", row == cr.max_row),
                ("left", col == cr.min_col),
                ("right", col == cr.max_col),
            ]:
                border = getattr(cell.border, name)
                assert (border.style if border else None) == (
                    "thin" if on_edge else None
                )
        wb = roundtrip(wb)
    wb.active.unmerge_cells(expected)
    assert not wb.active.merged_cells


def test_resize_preserves_individual_styles_and_neighboring_cells():
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("B3:D5")
    ws["B3"] = "heading"
    ws["C4"].fill = PatternFill("solid", fgColor="FF00FF00")
    ws["F4"] = "neighbor"
    neighbor = ws["F4"]
    ws.insert_rows(4, 2)
    assert ws["F6"] is neighbor
    assert neighbor.coordinate == "F6"
    loaded = roundtrip(wb).active
    assert loaded["C6"].fill.fgColor.rgb == "FF00FF00"
    assert loaded["F6"].value == "neighbor"


def test_single_cell_merge_survives_repeated_edits():
    ws = Workbook().active
    ws.merge_cells("B3:C4")
    ws["B3"] = 7
    ws.delete_rows(3)
    ws.delete_cols(2)
    assert str(ws.merged_cells) == "B3"
    assert ws["B3"].value == 7
    ws.insert_cols(1)
    ws.insert_rows(1)
    assert str(ws.merged_cells) == "C4"
    assert ws["C4"].value == 7
    ws.unmerge_cells("C4")


@pytest.mark.parametrize(
    "operation", ["insert_rows", "delete_rows", "insert_cols", "delete_cols"]
)
def test_opt_out_retains_legacy_dependency_behavior(operation):
    ws = Workbook().active
    ws.merge_cells("B3:C4")
    ws["B3"] = "heading"
    ws.print_area = "B3:C4"
    ws.defined_names.add(DefinedName("Block", attr_text="'Sheet'!$B$3:$C$4"))
    getattr(ws, operation)(1, update_dependencies=False)
    assert str(ws.merged_cells) == "B3:C4"
    assert ws.print_area == "'Sheet'!$B$3:$C$4"
    assert ws.defined_names["Block"].attr_text == "'Sheet'!$B$3:$C$4"


@pytest.mark.parametrize(
    "operation", ["insert_rows", "delete_rows", "insert_cols", "delete_cols"]
)
def test_zero_amount_is_identity(operation):
    ws = Workbook().active
    ws.merge_cells("B3:C4")
    original = (
        ws._cells,
        ws.merged_cells,
        ws.row_dimensions,
        ws.column_dimensions,
        ws["B3"]._coord,
        ws["B3"]._style,
    )
    getattr(ws, operation)(1, 0)
    current = (
        ws._cells,
        ws.merged_cells,
        ws.row_dimensions,
        ws.column_dimensions,
        ws["B3"]._coord,
        ws["B3"]._style,
    )
    assert all(left is right for left, right in zip(original, current))


def test_sparse_edit_and_append_after_deleting_last_column():
    ws = Workbook().active
    ws["A1"] = 1
    ws["XFC1000000"] = 2
    ws.insert_rows(2)
    ws.insert_cols(2)
    assert len(ws._cells) == 2
    assert ws["XFD1000001"].value == 2
    ws.delete_cols(16384)
    assert len(ws._cells) == 1
    assert ws._current_row == 1
    ws.append([3])
    assert ws["A2"].value == 3


def test_print_settings_and_names_fully_deleted():
    wb = Workbook()
    ws = wb.active
    ws.print_area = ["B3:C4", "E8:F9"]
    ws.print_title_rows = "3:4"
    ws.print_title_cols = "B:C"
    ws.row_dimensions[3].height = 31
    wb.defined_names.add(DefinedName("Gone", attr_text="'Sheet'!$B$3:$C$4"))
    ws.defined_names.add(DefinedName("Gone", attr_text="B3:C4"))
    ws.delete_rows(3, 2)
    assert ws.print_area == "'Sheet'!$E$6:$F$7"
    assert ws.print_title_rows == ""
    assert ws.print_title_cols == "$B:$C"
    assert not ws.row_dimensions
    assert wb.defined_names["Gone"].attr_text == "#REF!"
    assert ws.defined_names["Gone"].attr_text == "#REF!"
    ws.delete_rows(6, 2)
    loaded = roundtrip(wb)
    assert loaded.active.print_area == ""
    assert loaded.active.print_title_rows == ""
    assert loaded.defined_names["Gone"].attr_text == "#REF!"


@pytest.mark.parametrize(
    "operation,index,amount,key,span",
    [
        ("insert_cols", 1, 2, "D", (4, 6)),
        ("insert_cols", 3, 2, "B", (2, 6)),
        ("delete_cols", 1, 2, "A", (1, 2)),
        ("delete_cols", 3, 1, "B", (2, 3)),
    ],
)
def test_grouped_column_dimensions(operation, index, amount, key, span):
    wb = Workbook()
    ws = wb.active
    ws.column_dimensions.group("B", "D", outline_level=2, hidden=True)
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["B"].font = Font(bold=True)
    getattr(ws, operation)(index, amount)
    for _ in range(2):
        dim = wb.active.column_dimensions[key]
        assert (dim.min, dim.max) == span
        assert dim.index == key and dim.width == 26
        assert dim.hidden and dim.outlineLevel == 2 and dim.font.bold
        wb = roundtrip(wb)


def test_dimensions_without_cells_and_blank_inserted_rows():
    wb = Workbook()
    ws = wb.active
    ws.row_dimensions[3].height = 31
    ws.row_dimensions[3].hidden = True
    ws.column_dimensions["D"].width = 24
    ws.insert_rows(3)
    ws.delete_cols(1)
    assert not ws._cells and ws._current_row == 0
    assert 3 not in ws.row_dimensions
    loaded = roundtrip(wb).active
    assert loaded.row_dimensions[4].height == 31
    assert loaded.row_dimensions[4].hidden
    assert loaded.column_dimensions["C"].width == 24


@pytest.mark.parametrize(
    "value,operation,expected",
    [
        ("'Sheet'!$B$3:$D$6", "insert_rows", "'Sheet'!$B$4:$D$7"),
        ("'Sheet'!B$3:$D6", "insert_cols", "'Sheet'!C$3:$E6"),
        ("'Sheet'!$B:$D", "insert_rows", "'Sheet'!$B:$D"),
        ("'Sheet'!$B:$D", "insert_cols", "'Sheet'!$C:$E"),
        ("'Sheet'!$3:$6", "insert_rows", "'Sheet'!$4:$7"),
        ("'Sheet'!$3:$6", "insert_cols", "'Sheet'!$3:$6"),
        ("'Sheet'!B3, 'Other'!B3", "insert_rows", "'Sheet'!B4, 'Other'!B3"),
        ("=Sheet!B3", "insert_rows", "=Sheet!B4"),
        ("sheet!b3:d6", "insert_cols", "sheet!c3:e6"),
        ("'Sheet'!A1,'Sheet'!B3", "delete_rows", "#REF!,'Sheet'!B2"),
        ("'Sheet'!#REF!", "insert_rows", "'Sheet'!#REF!"),
        ("'[1]Sheet'!B3", "insert_rows", "'[1]Sheet'!B3"),
        ("'Sheet:Other'!B3", "insert_rows", "'Sheet:Other'!B3"),
        ("B3:D6", "insert_rows", "B3:D6"),
    ],
)
def test_static_name_references(value, operation, expected):
    wb = Workbook()
    wb.create_sheet("Other")
    wb.defined_names.add(DefinedName("Block", attr_text=value))
    getattr(wb.active, operation)(1, update_formulas=False)
    assert wb.defined_names["Block"].attr_text == expected


def test_scoped_names_and_escaped_sheet_titles():
    wb = Workbook()
    ws = wb.active
    ws.title = "O'Brien, Data!"
    other = wb.create_sheet("Other")
    ws.defined_names.add(DefinedName("Local", attr_text="$B$3:$D$6"))
    other.defined_names.add(DefinedName("Local", attr_text="$B$3:$D$6"))
    other.defined_names.add(
        DefinedName("Target", attr_text="'O''Brien, Data!'!$B$3:$D$6")
    )
    wb.defined_names.add(DefinedName("Global", attr_text="'O''Brien, Data!'!$B$3:$D$6"))
    ws.insert_rows(4, 2)
    for _ in range(2):
        ws, other = wb.worksheets
        assert ws.defined_names["Local"].attr_text == "$B$3:$D$8"
        assert other.defined_names["Local"].attr_text == "$B$3:$D$6"
        assert other.defined_names["Target"].attr_text == "'O''Brien, Data!'!$B$3:$D$8"
        assert wb.defined_names["Global"].attr_text == "'O''Brien, Data!'!$B$3:$D$8"
        wb = roundtrip(wb)


@pytest.mark.parametrize(
    "expression",
    [
        "SUM(Sheet!B3:D6)",
        "OFFSET(Sheet!B3,0,0,2,2)",
        'INDIRECT("Sheet!B3")',
        "Sheet!B3+Sheet!B4",
        '"Sheet!B3"',
        "42",
        "Sheet!B3 Sheet!C3",
        "OtherName",
    ],
)
def test_formula_names_are_unchanged(expression):
    wb = Workbook()
    wb.defined_names.add(DefinedName("Expression", attr_text=expression))
    wb.active.insert_rows(1, update_formulas=False)
    assert wb.defined_names["Expression"].attr_text == expression


def test_cell_formulas_on_all_sheets_are_unchanged():
    wb = Workbook()
    ws = wb.active
    other = wb.create_sheet("Other")
    ws["B3"] = "=SUM($A$1:A2)"
    other["A1"] = "=Sheet!B3"
    formula = ws["B3"]
    ws.insert_rows(1, update_formulas=False)
    ws.delete_cols(1, update_formulas=False)
    assert ws["A4"] is formula
    loaded = roundtrip(wb)
    assert loaded.active["A4"].value == "=SUM($A$1:A2)"
    assert loaded["Other"]["A1"].value == "=Sheet!B3"


@pytest.mark.parametrize(
    "index,amount,error",
    [
        (0, 1, ValueError),
        (-1, 1, ValueError),
        (1, -1, ValueError),
        (True, 1, TypeError),
        (1, 0.5, TypeError),
        (1048577, 1, ValueError),
    ],
)
def test_invalid_edit_does_not_mutate(index, amount, error):
    ws = Workbook().active
    ws["B3"] = "untouched"
    cells, coord = ws._cells, ws["B3"]._coord
    with pytest.raises(error):
        ws.insert_rows(index, amount)
    assert ws._cells is cells and ws["B3"]._coord is coord
    assert ws["B3"].value == "untouched"


def test_late_name_overflow_keeps_cells_styles_and_dimensions_unchanged():
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("B3:D5")
    ws["B3"].border = Border(bottom=Side(style="thin"))
    ws.row_dimensions[4].height = 42
    wb.defined_names.add(DefinedName("Outside", attr_text="'Sheet'!$A$1048576"))
    cells, merges, dims, borders = (
        ws._cells,
        ws.merged_cells,
        ws.row_dimensions,
        wb._borders,
    )
    before = {key: (cell._coord, cell._style) for key, cell in cells.items()}
    dimensions_before = dict(ws.row_dimensions[4].__dict__)
    with pytest.raises(OverflowError):
        ws.insert_rows(4)
    assert ws._cells is cells and ws.merged_cells is merges
    assert ws.row_dimensions is dims and wb._borders is borders
    assert ws.row_dimensions[4].__dict__ == dimensions_before
    for key, (coord, style) in before.items():
        assert ws._cells[key]._coord is coord
        assert ws._cells[key]._style is style
    assert wb.defined_names["Outside"].attr_text == "'Sheet'!$A$1048576"


def test_cell_overflow_and_overlapping_groups_fail_before_mutation():
    ws = Workbook().active
    ws["XFD1"] = "last"
    cells = ws._cells
    with pytest.raises(OverflowError):
        ws.insert_cols(1)
    assert ws._cells is cells and ws["XFD1"].value == "last"
    ws = Workbook().active
    ws.column_dimensions.group("B", "D")
    ws.column_dimensions["C"].width = 42
    dims = ws.column_dimensions
    with pytest.raises(ValueError, match="[Oo]verlap"):
        ws.insert_cols(1)
    assert ws.column_dimensions is dims


def test_unexpected_commit_error_restores_original_state(monkeypatch):
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("B3:D5")
    ws["B3"] = "heading"
    original_cells, original_merges, original_borders = (
        ws._cells,
        ws.merged_cells,
        wb._borders,
    )
    original_coord = ws["B3"]._coord
    setter = Worksheet.__setattr__
    fail = [True]

    def set_attribute(obj, name, value):
        if obj is ws and name == "_print_area" and fail[0]:
            fail[0] = False
            raise RuntimeError("injected commit error")
        setter(obj, name, value)

    monkeypatch.setattr(Worksheet, "__setattr__", set_attribute)
    with pytest.raises(RuntimeError, match="injected commit error"):
        ws.insert_rows(4)
    assert ws._cells is original_cells and ws.merged_cells is original_merges
    assert wb._borders is original_borders
    assert ws["B3"]._coord is original_coord
    assert ws["B3"].value == "heading"
