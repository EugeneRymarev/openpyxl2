# Copyright (c) 2010-2024 openpyxl
import pytest

from ..named_styles import NamedStyle
from ..named_styles import NamedStyleList
from openpyxl.utils.indexed_list import IndexedList


def test_descriptor(Worksheet):
    from ..cell_style import StyleArray
    from ..fonts import Font
    from ..styleable import StyleDescriptor

    class Styled:
        font = StyleDescriptor("_fonts", "fontId")

        def __init__(self):
            self._style = StyleArray()
            self.parent = Worksheet

    styled = Styled()
    styled.font = Font()
    assert styled.font == Font()


@pytest.fixture
def Workbook():
    class DummyWorkbook:
        _fonts = IndexedList()
        _fills = IndexedList()
        _borders = IndexedList()
        _protections = IndexedList()
        _alignments = IndexedList()
        _number_formats = IndexedList()
        _named_styles = NamedStyleList()

        def add_named_style(self, style):
            self._named_styles.append(style)
            style.bind(self)

    return DummyWorkbook()


@pytest.fixture
def Worksheet(Workbook):
    class DummyWorksheet:
        parent = Workbook

    return DummyWorksheet()


@pytest.fixture
def StyleableObject(Worksheet):
    from ..styleable import StyleableObject

    so = StyleableObject(sheet=Worksheet, style_array=[0] * 9)
    return so


def test_has_style(StyleableObject):
    so = StyleableObject
    so._style = None
    assert not so.has_style
    so.number_format = "dd"
    assert so.has_style


class TestNamedStyle:

    def test_assign_name(self, StyleableObject):
        so = StyleableObject
        wb = so.parent.parent
        style = NamedStyle(name="Standard")
        wb.add_named_style(style)

        so.style = "Standard"
        assert so._style.xfId == 0

    def test_assign_style(self, StyleableObject):
        so = StyleableObject
        wb = so.parent.parent
        style = NamedStyle(name="Standard")

        so.style = style
        assert so._style.xfId == 0

    def test_unknown_style(self, StyleableObject):
        so = StyleableObject

        with pytest.raises(ValueError):
            so.style = "Financial"

    def test_read(self, StyleableObject):
        so = StyleableObject
        wb = so.parent.parent

        red = NamedStyle(name="Red")
        wb.add_named_style(red)
        blue = NamedStyle(name="Blue")
        wb.add_named_style(blue)

        so._style.xfId = 1
        assert so.style == "Blue"

    def test_builtin(self, StyleableObject):
        so = StyleableObject
        so.style = "Hyperlink"
        assert so.style == "Hyperlink"

    def test_copy_not_share(self, StyleableObject):
        s1 = StyleableObject
        wb = s1.parent.parent

        from copy import copy

        s2 = copy(s1)
        s1.style = "Hyperlink"
        s2.style = "Hyperlink"
        assert s1._style is not s2._style

    def test_quote_prefix(self, StyleableObject):
        s1 = StyleableObject
        assert s1.quotePrefix is False
        s1.quotePrefix = True
        assert s1.quotePrefix is True

    def test_pivot_button(self, StyleableObject):
        s1 = StyleableObject
        assert s1.pivotButton is False
        s1.pivotButton = True
        assert s1.pivotButton is True


def test_styles_for_merged_cells(tmpdir):
    """
    Test styles for merged cells
    """
    tmpdir.chdir()
    from openpyxl import Workbook
    from openpyxl.styles.alignment import Alignment
    from openpyxl.styles.borders import BORDER_THIN
    from openpyxl.styles.borders import Border
    from openpyxl.styles.borders import Side
    from openpyxl.styles.colors import BLACK
    from openpyxl.styles.fills import FILL_PATTERN_DARKGRAY
    from openpyxl.styles.fills import PatternFill
    from openpyxl.styles.fonts import Font
    from openpyxl.styles.numbers import FORMAT_GENERAL
    from openpyxl.styles.numbers import FORMAT_NUMBER_00
    from openpyxl.styles.protection import Protection

    alignment = Alignment(horizontal="center", vertical="center")

    side = Side(style=BORDER_THIN, color=BLACK)
    border = Border(left=side, right=side, top=side, bottom=side)

    fill = PatternFill(fill_type=FILL_PATTERN_DARKGRAY)

    font = Font(name="Arial", size=9)

    style = NamedStyle("test_style", border=border)

    wb = Workbook()
    ws = wb.active
    ws.merge_cells("A1:B2")
    ws["A1"].alignment = alignment
    ws["A1"].border = border
    ws["A1"].fill = fill
    ws["A1"].font = font
    ws["A1"].number_format = FORMAT_NUMBER_00
    ws["A1"].protection = Protection(False, True)
    ws.merge_cells("A3:B3")
    ws["A3"].style = style

    # TODO
    # A bug that requires mandatory fixing.
    # You can't use write-read, because when reading ReadOnly,
    # the style property is not loaded, and when reading normally,
    # the applied styles for MergedCell objects are lost.
    # from openpyxl import load_workbook
    # xlsx_file = "merged_cells_styles.xlsx"
    # wb.save(xlsx_file)
    # wb = load_workbook(xlsx_file, read_only=True)
    # ws = wb.active

    for row in ws["A1:B2"]:
        for cell in row:
            assert cell.alignment == alignment
            assert cell.border == border
            assert cell.fill == fill
            assert cell.font == font
            assert cell.number_format == FORMAT_NUMBER_00
            assert cell.protection == Protection(False, True)
    assert ws["A3"].style == style
    assert ws["B3"].style == style
    del ws["A1"].alignment
    del ws["A1"].border
    del ws["A1"].fill
    del ws["A1"].font
    del ws["A1"].number_format
    del ws["A1"].protection
    del ws["A3"].style
    for row in ws["A1:B2"]:
        for cell in row:
            assert cell.alignment == Alignment()
            assert cell.border == Border()
            assert cell.fill == PatternFill()
            assert cell.font == Font()
            assert cell.number_format == FORMAT_GENERAL
            assert cell.protection == Protection()
    assert ws["A3"].style == "Normal"
    assert ws["B3"].style == "Normal"

