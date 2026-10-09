"""Executable acceptance criteria for the #1273 proposal; not yet implemented.

Run with --runxfail to reproduce the current failures. Strict xfails deliberately
make an unexpected fix visible instead of pretending this branch implements it.
"""
from io import BytesIO

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import absolute_coordinate
from openpyxl.workbook.defined_name import DefinedName


OPERATIONS = [
    ("insert_rows", "B4:C5", "B4", 4, "B", "$4:$5"),
    ("delete_rows", "B2:C3", "B2", 2, "B", "$2:$3"),
    ("insert_cols", "C3:D4", "C3", 3, "C", "$C:$D"),
    ("delete_cols", "A3:B4", "A3", 3, "A", "$A:$B"),
]


@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="Proposal #1273: structural edits do not update dependencies yet")
@pytest.mark.parametrize("operation,ref,anchor,row_key,col_key,titles", OPERATIONS)
@pytest.mark.parametrize("feature", ["merge", "global_name", "local_name", "print_area", "titles", "dimensions"])
@pytest.mark.parametrize("reload", [False, True])
def test_structural_edit_updates_dependencies(operation, ref, anchor, row_key,
                                               col_key, titles, feature, reload):
    wb = Workbook()
    ws = wb.active
    ws["B3"] = "anchor"
    source = "'Sheet'!$B$3:$C$4"
    if feature == "merge":
        ws.merge_cells("B3:C4")
    elif feature == "global_name":
        wb.defined_names.add(DefinedName("Block", attr_text=source))
    elif feature == "local_name":
        ws.defined_names.add(DefinedName("Block", attr_text=source))
    elif feature == "print_area":
        ws.print_area = "B3:C4"
    elif feature == "titles":
        ws.print_title_rows = "3:4"
        ws.print_title_cols = "B:C"
    elif feature == "dimensions":
        ws.row_dimensions[3].height = 31
        ws.column_dimensions["B"].width = 23
    getattr(ws, operation)(1)
    assert ws[anchor].value == "anchor", "the cell itself must move"
    if reload:
        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        wb = load_workbook(stream)
        ws = wb.active
    if feature == "merge":
        assert str(ws.merged_cells) == ref
        assert ws[anchor].value == "anchor"
    elif feature == "global_name":
        assert wb.defined_names["Block"].attr_text == "'Sheet'!" + absolute_coordinate(ref)
    elif feature == "local_name":
        assert ws.defined_names["Block"].attr_text == "'Sheet'!" + absolute_coordinate(ref)
    elif feature == "print_area":
        assert str(ws.print_area) == "'Sheet'!" + absolute_coordinate(ref)
    elif feature == "titles":
        actual = ws.print_title_rows if operation.endswith("rows") else ws.print_title_cols
        assert actual == titles
    elif feature == "dimensions":
        if operation.endswith("rows"):
            dimension = ws.row_dimensions.get(row_key)
            assert dimension is not None and dimension.height == 31
        else:
            dimension = ws.column_dimensions.get(col_key)
            assert dimension is not None and dimension.width == 23
