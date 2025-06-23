# Copyright (c) 2010-2025 openpyxl
import datetime

import pytest

from openpyxl.utils.exceptions import ReadOnlyWorkbookException
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.xml.constants import XLSM
from openpyxl.xml.constants import XLSX
from openpyxl.xml.constants import XLTM
from openpyxl.xml.constants import XLTX


@pytest.fixture
def workbook():
    from openpyxl.workbook.workbook import Workbook

    return Workbook


@pytest.fixture
def table():
    from openpyxl.worksheet.table import Table

    return Table


class TestWorkbook:
    @pytest.mark.parametrize(
        "has_vba, as_template, content_type",
        [
            (None, False, XLSX),
            (None, True, XLTX),
            (True, False, XLSM),
            (True, True, XLTM),
        ],
    )
    def test_template(self, has_vba, as_template, content_type, workbook):
        wb = workbook()
        wb._vba = has_vba
        wb.template = as_template
        assert wb.mime_type == content_type

    def test_named_styles(self, workbook):
        wb = workbook()
        assert wb.named_styles == ["Normal"]

    def test_immutable_builtins(self, workbook):
        wb1 = workbook()
        wb2 = workbook()
        normal = wb1._named_styles["Normal"]
        normal.font.color = "FF0000"
        assert wb2._named_styles["Normal"].font.color.index == 1

    def test_duplicate_table_name(self, workbook, table):
        wb = workbook()
        ws = wb.create_sheet()
        ws.add_table(table(displayName="Table1", ref="A1:D10"))
        assert True == wb._duplicate_name("Table1")
        assert True == wb._duplicate_name("TABLE1")

    def test_duplicate_defined_name(self, workbook):
        wb1 = workbook()
        wb1.defined_names["dfn1"] = DefinedName("dfn1")
        assert True == wb1._duplicate_name("dfn1")
        assert True == wb1._duplicate_name("DFN1")


def test_get_active_sheet(workbook):
    wb = workbook()
    assert wb.active == wb.worksheets[0]


def test_set_active_by_sheet(workbook):
    wb = workbook()
    names = ["Sheet", "Sheet1", "Sheet2"]
    for n in names:
        wb.create_sheet(n)
    for n in names:
        sheet = wb[n]
        wb.active = sheet
        assert wb.active == wb[n]


def test_set_active_by_index(workbook):
    wb = workbook()
    names = ["Sheet", "Sheet1", "Sheet2"]
    for n in names:
        wb.create_sheet(n)
    for idx, name in enumerate(names):
        wb.active = idx
        assert wb.active == wb.worksheets[idx]


@pytest.mark.xfail
def test_set_invalid_active_index(workbook):
    wb = workbook()
    with pytest.raises(ValueError):
        wb.active = 1


def test_set_invalid_sheet_by_name(workbook):
    wb = workbook()
    with pytest.raises(TypeError):
        wb.active = "Sheet"


def test_set_invalid_child_as_active(workbook):
    wb1 = workbook()
    wb2 = workbook()
    ws2 = wb2["Sheet"]
    with pytest.raises(ValueError):
        wb1.active = ws2


def test_set_hidden_sheet_as_active(workbook):
    wb = workbook()
    ws = wb.create_sheet()
    ws.sheet_state = "hidden"
    with pytest.raises(ValueError):
        wb.active = ws


def test_no_active(workbook):
    wb = workbook(write_only=True)
    assert wb.active is None


def test_create_sheet(workbook):
    wb = workbook()
    new_sheet = wb.create_sheet()
    assert new_sheet == wb.worksheets[-1]


def test_create_sheet_with_name(workbook):
    wb = workbook()
    new_sheet = wb.create_sheet(title="LikeThisName")
    assert new_sheet == wb.worksheets[-1]


def test_add_correct_sheet(workbook):
    wb = workbook()
    new_sheet = wb.create_sheet()
    wb._add_sheet(new_sheet)
    assert new_sheet == wb.worksheets[2]


def test_add_sheetname(workbook):
    wb = workbook()
    with pytest.raises(TypeError):
        wb._add_sheet("Test")


def test_add_sheet_from_other_workbook(workbook):
    wb1 = workbook()
    wb2 = workbook()
    ws = wb1.active
    with pytest.raises(ValueError):
        wb2._add_sheet(ws)


def test_create_sheet_readonly(workbook):
    wb = workbook()
    wb._read_only = True
    with pytest.raises(ReadOnlyWorkbookException):
        wb.create_sheet()


def test_remove_sheet(workbook):
    wb = workbook()
    new_sheet = wb.create_sheet(0)
    wb.remove(new_sheet)
    assert new_sheet not in wb.worksheets


def test_move_sheet(workbook):
    wb = workbook()
    for i in range(9):
        wb.create_sheet()
    expected = [
        "Sheet",
        "Sheet1",
        "Sheet2",
        "Sheet3",
        "Sheet4",
        "Sheet5",
        "Sheet6",
        "Sheet7",
        "Sheet8",
        "Sheet9",
    ]
    assert wb.sheetnames == expected
    ws = wb["Sheet9"]
    wb.move_sheet(ws, -5)
    expected = [
        "Sheet",
        "Sheet1",
        "Sheet2",
        "Sheet3",
        "Sheet9",
        "Sheet4",
        "Sheet5",
        "Sheet6",
        "Sheet7",
        "Sheet8",
    ]
    assert wb.sheetnames == expected


def test_move_sheet2(workbook):
    wb = workbook()
    for i in range(9):
        wb.create_sheet()
    expected = [
        "Sheet",
        "Sheet1",
        "Sheet2",
        "Sheet3",
        "Sheet4",
        "Sheet5",
        "Sheet6",
        "Sheet7",
        "Sheet8",
        "Sheet9",
    ]
    assert wb.sheetnames == expected
    wb.move_sheet("Sheet9", -5)
    expected = [
        "Sheet",
        "Sheet1",
        "Sheet2",
        "Sheet3",
        "Sheet9",
        "Sheet4",
        "Sheet5",
        "Sheet6",
        "Sheet7",
        "Sheet8",
    ]
    assert wb.sheetnames == expected


def test_getitem(workbook):
    wb = workbook()
    ws = wb["Sheet"]
    assert isinstance(ws, Worksheet)
    with pytest.raises(KeyError):
        _ = wb["NotThere"]


def test_get_chartsheet(workbook):
    wb = workbook()
    cs = wb.create_chartsheet()
    assert wb[cs.title] is cs


def test_del_worksheet(workbook):
    wb = workbook()
    del wb["Sheet"]
    assert wb.worksheets == []


def test_del_chartsheet(workbook):
    wb = workbook()
    cs = wb.create_chartsheet()
    del wb[cs.title]
    assert wb.chartsheets == []


def test_contains(workbook):
    wb = workbook()
    assert "Sheet" in wb
    assert "NotThere" not in wb


def test_iter(workbook):
    wb = workbook()
    ws = None
    for ws in wb:
        pass
    assert ws.title == "Sheet"


def test_index(workbook):
    wb = workbook()
    new_sheet = wb.create_sheet()
    sheet_index = wb.index(new_sheet)
    assert sheet_index == 1


def test_get_sheet_names(workbook):
    wb = workbook()
    names = ["Sheet", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"]
    for count in range(5):
        wb.create_sheet(0)
    assert wb.sheetnames == names


def test_add_invalid_worksheet_class_instance(workbook):
    class AlternativeWorksheet:
        def __init__(self, parent_workbook, title=None):
            self.parent_workbook = parent_workbook
            if not title:
                title = "AlternativeSheet"
            self.title = title

    wb = workbook
    ws = AlternativeWorksheet(parent_workbook=wb)
    with pytest.raises(TypeError):
        wb._add_sheet(worksheet=ws)


class TestCopy:
    def test_worksheet_copy(self, workbook):
        wb = workbook()
        ws1 = wb.active
        ws2 = wb.copy_worksheet(ws1)
        assert ws2 is not None

    @pytest.mark.parametrize(
        "title, copy",
        [("TestSheet", "TestSheet Copy"), ("D\xfcsseldorf", "D\xfcsseldorf Copy")],
    )
    def test_worksheet_copy_name(self, title, copy, workbook):
        wb = workbook()
        ws1 = wb.active
        ws1.title = title
        ws2 = wb.copy_worksheet(ws1)
        assert ws2.title == copy

    def test_cannot_copy_readonly(self, workbook):
        wb = workbook()
        ws = wb.active
        wb._read_only = True
        with pytest.raises(ValueError):
            wb.copy_worksheet(ws)

    def test_cannot_copy_writeonly(self, workbook):
        wb = workbook(write_only=True)
        ws = wb.create_sheet()
        with pytest.raises(ValueError):
            wb.copy_worksheet(ws)

    def test_default_epoch(self, workbook):
        wb = workbook()
        assert wb.epoch == datetime.datetime(1899, 12, 30)

    def test_assign_epoch(self, workbook):
        wb = workbook()
        wb.epoch = datetime.datetime(1904, 1, 1)

    def test_invalid_epoch(self, workbook):
        wb = workbook()
        with pytest.raises(ValueError):
            wb.epoch = datetime.datetime(1970, 1, 1)
