"""Formula dependencies exercised through the four public structural methods."""

from io import BytesIO

import pytest

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.formula.structural import FormulaTranslationError
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.formula import DataTableFormula


def reload(wb):
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return load_workbook(stream)


@pytest.mark.parametrize(
    "operation,index,amount,formula,expected",
    [
        ("insert_rows", 3, 2, "=A1+A3+$A$3+A$3+$A3", "=A1+A5+$A$5+A$5+$A5"),
        ("insert_cols", 3, 2, "=A1+C1+$C$1+C$1+$C1", "=A1+E1+$E$1+E$1+$E1"),
        ("delete_rows", 3, 2, "=A2+A3+A4+A5", "=A2+#REF!+#REF!+A3"),
        ("delete_cols", 3, 2, "=B1+C1+D1+E1", "=B1+#REF!+#REF!+C1"),
        ("insert_rows", 3, 2, "=SUM(A3:A8)", "=SUM(A5:A10)"),
        ("insert_rows", 5, 2, "=SUM(A3:A8)", "=SUM(A3:A10)"),
        ("insert_rows", 9, 2, "=SUM(A3:A8)", "=SUM(A3:A8)"),
        ("delete_rows", 3, 2, "=SUM(A3:A8)", "=SUM(A3:A6)"),
        ("delete_rows", 5, 2, "=SUM(A3:A8)", "=SUM(A3:A6)"),
        ("delete_rows", 7, 2, "=SUM(A3:A8)", "=SUM(A3:A6)"),
        ("delete_rows", 2, 9, "=SUM(A3:A8)", "=SUM(#REF!)"),
        ("insert_cols", 4, 2, "=SUM(B2:F4)", "=SUM(B2:H4)"),
        ("delete_cols", 2, 2, "=SUM(B2:F4)", "=SUM(B2:D4)"),
        ("delete_cols", 5, 2, "=SUM(B2:F4)", "=SUM(B2:D4)"),
        ("delete_cols", 2, 5, "=SUM(B2:F4)", "=SUM(#REF!)"),
        ("insert_rows", 5, 2, "=SUM($3:8,A:C)", "=SUM($3:10,A:C)"),
        ("insert_cols", 2, 2, "=SUM(3:8,$A:C)", "=SUM(3:8,$A:E)"),
        ("delete_rows", 3, 6, "=SUM(3:8,A:C)", "=SUM(#REF!,A:C)"),
        ("delete_cols", 1, 3, "=SUM(3:8,A:C)", "=SUM(3:8,#REF!)"),
        ("insert_rows", 5, 2, "=SUM(A8:$B$3)", "=SUM(A10:$B$3)"),
        ("insert_cols", 2, 2, "=SUM(c8:$a$3)", "=SUM(e8:$a$3)"),
        (
            "insert_rows",
            3,
            2,
            '=IF(A3="A3","say ""A3""",A3)',
            '=IF(A5="A3","say ""A3""",A5)',
        ),
        (
            "insert_rows",
            3,
            2,
            '=INDIRECT("A3")+OFFSET(A3,1,0)',
            '=INDIRECT("A3")+OFFSET(A5,1,0)',
        ),
        (
            "insert_rows",
            3,
            2,
            "=  SUM( A3:A5,  A8 ) +\n  A3",
            "=  SUM( A5:A7,  A10 ) +\n  A5",
        ),
        ("insert_rows", 3, 2, "=SUM((A3:A8,C3:C8))", "=SUM((A5:A10,C5:C10))"),
        ("insert_rows", 3, 2, "=SUM(A3:C8 B4:D9)", "=SUM(A5:C10 B6:D11)"),
        ("insert_rows", 3, 2, "={1,2;3,4}+A3", "={1,2;3,4}+A5"),
        ("insert_rows", 3, 2, "=LOG10(A3)+1E-3", "=LOG10(A5)+1E-3"),
        ("insert_rows", 3, 2, "=XFE1+A1048577+A3", "=XFE1+A1048577+A5"),
        ("insert_rows", 3, 2, "=SUM(Table1[Amount])+A3", "=SUM(Table1[Amount])+A5"),
        ("insert_rows", 3, 2, "=[1]Sheet!A3+A3", "=[1]Sheet!A3+A5"),
        (
            "insert_rows",
            3,
            2,
            "='#REF!'!A3+#REF!+Sheet!#REF!+A3",
            "='#REF!'!A3+#REF!+Sheet!#REF!+A5",
        ),
        ("insert_rows", 3, 2, "=@A3+SUM(A3#)", "=@A5+SUM(A5#)"),
        ("insert_rows", 3, 2, "=@'Sheet'!A3", "=@'Sheet'!A5"),
        ("delete_rows", 3, 2, "=SUM(A3#)", "=SUM(#REF!)"),
        ("insert_rows", 3, 2, "=A1048576", "=#REF!"),
        ("insert_cols", 3, 2, "=XFD1", "=#REF!"),
        ("insert_rows", 3, 2, "=SUM(A1048573:A1048576)", "=SUM(A1048575:A1048576)"),
        ("insert_rows", 3, 2, "=SUM(A1:A1048576)", "=SUM(A1:A1048576)"),
        ("delete_rows", 3, 2, "=SUM(1:1048576)", "=SUM(1:1048576)"),
        ("delete_cols", 3, 2, "=SUM(A:XFD)", "=SUM(A:XFD)"),
    ],
)
def test_reference_semantics(operation, index, amount, formula, expected):
    wb = Workbook()
    ws = wb.active
    # Place the formula below/right of the edited band so it survives and moves.
    ws["Z30"] = formula
    cell = ws["Z30"]
    getattr(ws, operation)(index, amount)
    assert cell.value == expected
    coord = cell.coordinate
    for _ in range(2):
        wb = reload(wb)
        assert wb.active[coord].value == expected
        assert wb.active[coord].data_type == "f"


@pytest.mark.parametrize(
    "operation,expected",
    [
        ("insert_rows", "='O''Brien, Data!'!$B$5+C3"),
        ("delete_rows", "='O''Brien, Data!'!#REF!+C3"),
        ("insert_cols", "='O''Brien, Data!'!$D$3+C3"),
        ("delete_cols", "='O''Brien, Data!'!#REF!+C3"),
    ],
)
def test_references_from_all_sheets(operation, expected):
    wb = Workbook()
    ws = wb.active
    ws.title = "O'Brien, Data!"
    other = wb.create_sheet("Other")
    other["A1"] = "='O''Brien, Data!'!$B$3+C3"
    ws["H20"] = "=Other!A1"
    stable = ws["H20"]
    getattr(ws, operation)(3 if operation.endswith("rows") else 2, 2)
    assert other["A1"].value == expected
    assert stable.value == "=Other!A1"
    assert reload(wb)["Other"]["A1"].value == expected


def test_formula_move_is_not_copy_translation():
    ws = Workbook().active
    ws["A10"] = "=A1+$B$5"
    ws["C1"] = "=A10"
    ws.insert_rows(5, 2)
    assert ws["A12"].value == "=A1+$B$7"
    assert ws["C1"].value == "=A12"


def test_deleted_formula_is_not_translated_but_retained_merge_anchor_is():
    ws = Workbook().active
    ws["F3"] = "=UNSUPPORTED("
    ws.merge_cells("B3:D5")
    ws["B3"] = "=A8"
    ws.delete_rows(3)
    assert ws["B3"].value == "=A7"
    assert ws["F3"].value is None


def test_calculated_names_and_name_identifiers():
    wb = Workbook()
    ws = wb.active
    other = wb.create_sheet("Other")
    definitions = [
        (wb.defined_names, "Total", "SUM(Sheet!$B$3:$B$8)", "SUM(Sheet!$B$3:$B$10)"),
        (ws.defined_names, "Local", "OFFSET($B$5,0,0,2,2)", "OFFSET($B$7,0,0,2,2)"),
        (other.defined_names, "Local", "SUM(B5,Sheet!B5)", "SUM(B5,Sheet!B7)"),
        (wb.defined_names, "Ambiguous", "SUM(B5,Sheet!B5)", "SUM(B5,Sheet!B7)"),
        (wb.defined_names, "TextRef", 'INDIRECT("Sheet!B5")', 'INDIRECT("Sheet!B5")'),
        (wb.defined_names, "Constant", "42", "42"),
        (wb.defined_names, "Alias", "Total", "Total"),
    ]
    for scope, name, value, expected in definitions:
        scope.add(DefinedName(name, attr_text=value))
    other["A1"] = "=Total+Sheet!Local"
    ws.insert_rows(5, 2)
    for scope, name, value, expected in definitions:
        assert scope[name].attr_text == expected
    assert other["A1"].value == "=Total+Sheet!Local"
    assert reload(wb).defined_names["Total"].attr_text == definitions[0][3]


@pytest.mark.parametrize(
    "value",
    [
        "=SUM(Sheet:Other!A3)",
        "=SUM(Name:A3)",
        "=SUM(A3:INDEX(A:A,8))",
        "=SUM(INDEX(A:A,3):A8)",
        "=A3\t+A4",
        "=SUM(A3",
        "=A3)",
        '=A3+"unterminated',
        ArrayFormula("B3:B5", "=A3:A5"),
        DataTableFormula("B3:C5", r1="A1"),
    ],
)
def test_unsupported_formulas_fail_atomically(value):
    wb = Workbook()
    ws = wb.active
    other = wb.create_sheet("Other")
    ws["A3"] = 7
    ws["C3"] = "=A3"
    other["B3"] = value
    original_cells, original_coord, calc = ws._cells, ws["A3"]._coord, wb.calculation
    with pytest.raises(FormulaTranslationError, match="Other!B3"):
        ws.insert_rows(1)
    assert ws._cells is original_cells and ws["A3"]._coord is original_coord
    assert ws["C3"].value == "=A3" and other["B3"].value == value
    assert wb.calculation is calc


@pytest.mark.parametrize(
    "operation", ["insert_rows", "insert_cols", "delete_rows", "delete_cols"]
)
def test_formula_opt_out_still_updates_metadata(operation):
    wb = Workbook()
    ws = wb.active
    ws.merge_cells("B5:D7")
    ws["H20"] = "=SUM(Sheet:Other!A3)"
    wb.defined_names.add(DefinedName("Calc", attr_text="SUM(Sheet!A3)"))
    getattr(ws, operation)(1, update_formulas=False)
    assert str(ws.merged_cells) != "B5:D7"
    assert (
        next(c.value for c in ws._cells.values() if c.data_type == "f")
        == "=SUM(Sheet:Other!A3)"
    )
    assert wb.defined_names["Calc"].attr_text == "SUM(Sheet!A3)"


def test_legacy_dependency_opt_out_and_zero_amount():
    ws = Workbook().active
    ws["B3"] = "=SUM("
    cells = ws._cells
    ws.insert_rows(1, 0)
    assert ws._cells is cells
    ws.insert_rows(1, update_dependencies=False)
    assert ws["B4"].value == "=SUM("


def test_recalculation_flags_preserve_user_settings():
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = False
    wb.calculation.forceFullCalc = False
    wb.calculation.calcMode = "manual"
    wb.calculation.iterate = True
    wb.calculation.iterateCount = 50
    old = wb.calculation
    wb.active["A2"] = "=1+2"
    wb.active.insert_rows(1)
    assert wb.calculation is not old and old.fullCalcOnLoad is False
    loaded = reload(wb)
    assert loaded.calculation.fullCalcOnLoad and loaded.calculation.forceFullCalc
    assert loaded.calculation.calcMode == "manual"
    assert loaded.calculation.iterate and loaded.calculation.iterateCount == 50


def test_cross_sheet_formulas_and_names_rollback_on_commit_error(monkeypatch):
    wb = Workbook()
    ws = wb.active
    other = wb.create_sheet("Other")
    ws["A3"] = 7
    other["B1"] = "=Sheet!A3"
    wb.defined_names.add(DefinedName("Calc", attr_text="SUM(Sheet!A3)"))
    original_cells, original_calc = ws._cells, wb.calculation
    setter = Workbook.__setattr__
    fail = [True]

    def inject(obj, name, value):
        if obj is wb and name == "calculation" and fail[0]:
            fail[0] = False
            raise RuntimeError("injected calculation failure")
        setter(obj, name, value)

    monkeypatch.setattr(Workbook, "__setattr__", inject)
    with pytest.raises(RuntimeError, match="injected"):
        ws.insert_rows(1)
    assert ws._cells is original_cells and ws["A3"].value == 7
    assert ws["A3"].coordinate == "A3"
    assert other["B1"].value == "=Sheet!A3"
    assert wb.defined_names["Calc"].attr_text == "SUM(Sheet!A3)"
    assert wb.calculation is original_calc


def test_sequential_edits_of_spills_and_existing_errors():
    ws = Workbook().active
    ws["D10"] = "=@A3+SUM(A3#)+#REF!"
    ws.delete_rows(3)
    assert ws["D9"].value == "=@#REF!+SUM(#REF!)+#REF!"
    ws.insert_cols(1)
    assert ws["E9"].value == "=@#REF!+SUM(#REF!)+#REF!"


def test_existing_error_in_static_union_does_not_hide_other_references():
    wb = Workbook()
    wb.defined_names.add(DefinedName("Block", attr_text="#REF!,Sheet!A3"))
    wb.active.insert_rows(1)
    assert wb.defined_names["Block"].attr_text == "#REF!,Sheet!A4"


@pytest.mark.parametrize(
    "expression", ["SUM(Sheet:Other!A3)", "SUM(Sheet!A3))", "Sheet!A3,Sheet:Other!A3"]
)
def test_name_error_does_not_mutate_other_names_or_cells(expression):
    wb = Workbook()
    ws = wb.active
    ws["A3"] = 7
    wb.defined_names.add(DefinedName("Good", attr_text="Sheet!A3"))
    ws.defined_names.add(DefinedName("Bad", attr_text=expression))
    cells = ws._cells
    with pytest.raises(FormulaTranslationError, match="Defined name Bad"):
        ws.insert_rows(1)
    assert ws._cells is cells and wb.defined_names["Good"].attr_text == "Sheet!A3"


def test_sparse_workbook_formulas_do_not_materialize_rectangles():
    wb = Workbook()
    ws = wb.active
    other = wb.create_sheet("Other")
    ws["A1"] = 1
    ws["Z900000"] = "=A1"
    other["XFD1048576"] = "=Sheet!Z900000"
    ws.insert_rows(2)
    assert len(ws._cells) == 2 and len(other._cells) == 1
    assert ws["Z900001"].value == "=A1"
    assert other["XFD1048576"].value == "=Sheet!Z900001"


@pytest.mark.parametrize("value", [None, 0, 1, "yes"])
def test_update_formulas_requires_bool(value):
    ws = Workbook().active
    ws["A1"] = 1
    with pytest.raises(TypeError, match="update_formulas"):
        ws.insert_rows(1, update_formulas=value)
    assert ws["A1"].value == 1
