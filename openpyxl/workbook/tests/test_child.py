# Copyright (c) 2010-2025 openpyxl
import pytest


@pytest.fixture
def workbook_child():
    from openpyxl.workbook.child import _WorkbookChild

    return _WorkbookChild


class DummyWorkbook:
    encoding = "utf-8"

    def __init__(self):
        self.sheetnames = ["Sheet 1"]


s = r"[\\*?:/\[\]]"


@pytest.mark.parametrize(
    "value",
    ["Title:", "title?", "title/", "title[", "title]", r"title\\", "title*"],
)
def test_invalid_chars(value):
    from openpyxl.workbook.child import INVALID_TITLE_REGEX

    assert INVALID_TITLE_REGEX.search(value)


@pytest.mark.parametrize(
    "names, value, result",
    [
        ([], "Sheet", "Sheet"),
        (["Sheet2"], "Sheet2", "Sheet21"),  # suggestions are stupid
        (["R\xf3g"], "R\xf3g", "R\xf3g1"),
        (["Sheet", "Sheet1"], "Sheet", "Sheet2"),
        (["Regex Test ("], "Regex Test (", "Regex Test (1"),
        (
            ["Foo", "Baz", "Sheet2", "Sheet3", "Bar", "Sheet4", "Sheet6"],
            "Sheet",
            "Sheet",
        ),
        (["Foo"], "FOO", "FOO1"),
    ],
)
def test_duplicate_title(names, value, result):
    from openpyxl.workbook.child import avoid_duplicate_name

    title = avoid_duplicate_name(names, value)
    assert title == result


class TestWorkbookChild:
    def test_ctor(self, workbook_child):
        wb = DummyWorkbook()
        child = workbook_child(wb)
        assert child.parent == wb
        assert child.encoding == "utf-8"
        assert child.title == "Sheet"

    def test_repr(self, workbook_child):
        wb = DummyWorkbook()
        child = workbook_child(wb)
        assert repr(child) == '<_WorkbookChild "Sheet">'

    def test_invalid_title(self, workbook_child):
        wb = DummyWorkbook()
        child = workbook_child(wb)
        with pytest.raises(ValueError):
            child.title = "title?"

    def test_reassign_title(self, workbook_child):
        wb = DummyWorkbook()
        child = workbook_child(wb, "Sheet")
        assert child.title == "Sheet"

    def test_title_too_long(self, workbook_child, recwarn):
        workbook_child(DummyWorkbook(), "X" * 50)
        w = recwarn.pop()
        assert w.category == UserWarning

    def test_set_encoded_title(self, workbook_child):
        with pytest.raises(ValueError):
            workbook_child(DummyWorkbook(), b"B\xc3\xbcro")

    def test_empty_title(self, workbook_child):
        child = workbook_child(DummyWorkbook())
        with pytest.raises(ValueError):
            child.title = ""
