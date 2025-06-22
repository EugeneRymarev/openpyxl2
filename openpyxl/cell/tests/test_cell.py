# Copyright (c) 2010-2025 openpyxl
import datetime
import itertools

import numpy
import pandas
import pytest
from openpyxl.cell.cell import ERROR_CODES
from openpyxl.cell.cell import get_time_format
from openpyxl.comments.comments import Comment


@pytest.fixture
def dummy_worksheet():
    from openpyxl.cell.cell import Cell
    from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900
    from openpyxl.utils.indexed_list import IndexedList

    class Wb:
        epoch = CALENDAR_WINDOWS_1900
        _fonts = IndexedList()
        _fills = IndexedList()
        _borders = IndexedList()
        _protections = IndexedList()
        _alignments = IndexedList()
        _number_formats = IndexedList()
        _cell_styles = IndexedList()

    class Ws:
        def cell(self, column, row):
            return Cell(self, row=row, column=column)

        encoding = "utf-8"
        parent = Wb()
        title = "Dummy Worksheet"
        _comment_count = 0

    return Ws()


@pytest.fixture
def cell_():
    from openpyxl.cell.cell import Cell

    return Cell


@pytest.fixture
def dummy_cell(dummy_worksheet, cell_):
    ws = dummy_worksheet
    cell = cell_(ws, column=1, row=1)
    return cell


@pytest.fixture
def merged_cell(dummy_worksheet):
    from openpyxl.cell.cell import MergedCell

    return MergedCell(dummy_worksheet, 1, 4)


def test_ctor(dummy_cell):
    cell = dummy_cell
    assert cell.data_type == "n"
    assert cell.column == 1
    assert cell.row == 1
    assert cell.coordinate == "A1"
    assert cell.value is None
    assert cell.comment is None


@pytest.mark.parametrize("datatype", ["n", "d", "s", "b", "f", "e"])
def test_null(dummy_cell, datatype):
    cell = dummy_cell
    cell.data_type = datatype
    assert cell.data_type == datatype
    cell.value = None
    assert cell.data_type == "n"


@pytest.mark.parametrize("value", ["hello", ".", "0800"])
def test_string(dummy_cell, value):
    cell = dummy_cell
    cell.value = "hello"
    assert cell.data_type == "s"


@pytest.mark.parametrize("value", ["=42", "=if(A1<4;-1;1)"])
def test_formula(dummy_cell, value):
    cell = dummy_cell
    cell.value = value
    assert cell.data_type == "f"


def test_not_formula(dummy_cell):
    dummy_cell.value = "="
    assert dummy_cell.data_type == "s"
    assert dummy_cell.value == "="


@pytest.mark.parametrize("value", [True, False])
def test_boolean(dummy_cell, value):
    cell = dummy_cell
    cell.value = value
    assert cell.data_type == "b"


@pytest.mark.parametrize("error_string", ERROR_CODES)
def test_error_codes(dummy_cell, error_string):
    cell = dummy_cell
    cell.value = error_string
    assert cell.data_type == "e"


@pytest.mark.parametrize(
    "value, number_format",
    [
        (datetime.datetime(2010, 7, 13, 6, 37, 41), "yyyy-mm-dd h:mm:ss"),
        (datetime.date(2010, 7, 13), "yyyy-mm-dd"),
        (datetime.time(1, 3), "h:mm:ss"),
    ],
)
def test_insert_date(dummy_cell, value, number_format):
    cell = dummy_cell
    cell.value = value
    assert cell.data_type == "d"
    assert cell.is_date
    assert cell.number_format == number_format


@pytest.mark.pandas_required
def test_timestamp(dummy_cell):
    cell = dummy_cell
    cell.value = pandas.Timestamp("2018-09-05")
    assert cell.number_format == "yyyy-mm-dd h:mm:ss"


def test_time_format_datetime_subclass():
    class TestDatetime(datetime.datetime):
        pass

    number_format = get_time_format(TestDatetime)
    assert number_format == "yyyy-mm-dd h:mm:ss"


def test_time_format_date_subclass():
    class TestDate(datetime.date):
        pass

    number_format = get_time_format(TestDate)
    assert number_format == "yyyy-mm-dd"


def test_time_format_no_date_subclass():
    with pytest.raises(ValueError):
        number_format = get_time_format(object)


def test_not_overwrite_time_format(dummy_cell):
    cell = dummy_cell
    cell.number_format = "mmm-yy"
    cell.value = datetime.date(2010, 7, 13)
    assert cell.number_format == "mmm-yy"


@pytest.mark.parametrize(
    "value, is_date",
    [
        (None, True),
        ("testme", False),
        (True, False),
    ],
)
def test_cell_formatted_as_date(dummy_cell, value, is_date):
    cell = dummy_cell
    cell.value = datetime.datetime.today()
    cell.value = value
    assert cell.is_date == is_date
    assert cell.value == value


def test_illegal_characters(dummy_cell):
    from openpyxl.utils.exceptions import IllegalCharacterError

    cell = dummy_cell
    # The bytes 0x00 through 0x1F inclusive must be manually escaped in values.
    illegal_chrs = itertools.chain(range(9), range(11, 13), range(14, 32))
    for i in illegal_chrs:
        with pytest.raises(IllegalCharacterError):
            cell.value = chr(i)
        with pytest.raises(IllegalCharacterError):
            cell.value = f"A {chr(i)} B"
    cell.value = chr(33)
    cell.value = chr(9)  # Tab
    cell.value = chr(10)  # Newline
    cell.value = chr(13)  # Carriage return
    cell.value = " Leading and trailing spaces are legal "


@pytest.mark.xfail
def test_timedelta(dummy_cell):
    cell = dummy_cell
    cell.value = datetime.timedelta(days=1, hours=3)
    assert cell.value == 1.125
    assert cell.data_type == "n"
    assert cell.is_date is False
    assert cell.number_format == "[hh]:mm:ss"


def test_repr(dummy_cell):
    cell = dummy_cell
    assert repr(cell) == "<Cell 'Dummy Worksheet'.A1>"


def test_repr_object(dummy_cell):
    class Dummy:
        def __str__(self):
            return "something"

    cell = dummy_cell
    try:
        cell.value = Dummy()
    except ValueError as err:
        assert "something" not in str(err)


def test_comment_assignment(dummy_cell):
    assert dummy_cell.comment is None
    comm = Comment("text", "author")
    dummy_cell.comment = comm
    assert dummy_cell.comment == comm


def test_only_one_cell_per_comment(dummy_cell):
    ws = dummy_cell.parent
    comm = Comment("text", "author")
    dummy_cell.comment = comm
    c2 = ws.cell(column=1, row=2)
    c2.comment = comm
    assert c2.comment.parent is c2


def test_remove_comment(dummy_cell):
    comm = Comment("text", "author")
    dummy_cell.comment = comm
    dummy_cell.comment = None
    assert dummy_cell.comment is None


def test_cell_offset(dummy_cell):
    cell = dummy_cell
    assert cell.offset(2, 1).coordinate == "B3"


class TestEncoding:
    pound = chr(163)
    test_string = f"Compound Value {pound}".encode("latin1")

    def test_bad_encoding(self):
        from openpyxl.workbook.workbook import Workbook

        wb = Workbook()
        ws = wb.active
        cell = ws["A1"]
        with pytest.raises(UnicodeDecodeError):
            cell.check_string(self.test_string)
        with pytest.raises(UnicodeDecodeError):
            cell.value = self.test_string

    def test_good_encoding(self):
        from openpyxl.workbook.workbook import Workbook

        wb = Workbook()
        wb.encoding = "latin1"
        ws = wb.active
        cell = ws["A1"]
        cell.value = self.test_string


def test_font(dummy_worksheet, cell_):
    from openpyxl.styles.fonts import Font

    font = Font(bold=True)
    ws = dummy_worksheet
    ws.parent._fonts.add(font)
    cell = cell_(ws, row=1, column=1)
    assert cell.font == font


def test_fill(dummy_worksheet, cell_):
    from openpyxl.styles.fills import PatternFill

    fill = PatternFill(patternType="solid", fgColor="FF0000")
    ws = dummy_worksheet
    ws.parent._fills.add(fill)
    cell = cell_(ws, column="A", row=1)
    assert cell.fill == fill


def test_border(dummy_worksheet, cell_):
    from openpyxl.styles.borders import Border

    border = Border()
    ws = dummy_worksheet
    ws.parent._borders.add(border)
    cell = cell_(ws, column="A", row=1)
    assert cell.border == border


def test_number_format(dummy_worksheet, cell_):
    ws = dummy_worksheet
    ws.parent._number_formats.add("dd--hh--mm")
    cell = cell_(ws, column="A", row=1)
    cell.number_format = "dd--hh--mm"
    assert cell.number_format == "dd--hh--mm"


def test_alignment(dummy_worksheet, cell_):
    from openpyxl.styles.alignment import Alignment

    align = Alignment(wrapText=True)
    ws = dummy_worksheet
    ws.parent._alignments.add(align)
    cell = cell_(ws, column="A", row=1)
    assert cell.alignment == align


def test_protection(dummy_worksheet, cell_):
    from openpyxl.styles.protection import Protection

    prot = Protection(locked=False)
    ws = dummy_worksheet
    ws.parent._protections.add(prot)
    cell = cell_(ws, column="A", row=1)
    assert cell.protection == prot


def test_pivot_button(dummy_worksheet, cell_):
    ws = dummy_worksheet
    cell = cell_(ws, column="A", row=1)
    _ = cell.style_id
    cell._style.pivotButton = 1
    assert cell.pivotButton is True


def test_quote_prefix(dummy_worksheet, cell_):
    ws = dummy_worksheet
    cell = cell_(ws, column="A", row=1)
    _ = cell.style_id
    cell._style.quotePrefix = 1
    assert cell.quotePrefix is True


def test_remove_hyperlink(dummy_cell):
    """Remove a cell hyperlink"""
    cell = dummy_cell
    cell.hyperlink = "http://test.com"
    cell.hyperlink = None
    assert cell.hyperlink is None


class TestMergedCell:
    def test_value(self, merged_cell):
        cell = merged_cell
        assert cell._value is None

    def test_data_type(self, merged_cell):
        cell = merged_cell
        assert cell.data_type == "n"

    def test_comment(self, merged_cell):
        cell = merged_cell
        assert cell.comment is None

    def test_coordinate(self, merged_cell):
        cell = merged_cell
        assert cell.coordinate == "D1"

    def test_repr(self, merged_cell):
        cell = merged_cell
        assert repr(cell) == "<MergedCell 'Dummy Worksheet'.D1>"

    def test_hyperlink(self, merged_cell):
        cell = merged_cell
        assert cell.hyperlink is None


@pytest.mark.numpy_required
def test_write_numpy_to_cell(dummy_cell):
    data = numpy.array([1.0])
    cell = dummy_cell
    cell.value = data[0]
