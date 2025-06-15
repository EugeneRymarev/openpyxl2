# Copyright (c) 2010-2025 openpyxl
import copy

import pytest
from openpyxl.styles.named_styles import NamedStyle
from openpyxl.styles.named_styles import NamedStyleList
from openpyxl.utils.indexed_list import IndexedList


@pytest.fixture
def styleable_object(worksheet):
    from openpyxl.styles.styleable import StyleableObject

    so = StyleableObject(sheet=worksheet, style_array=[0] * 9)
    return so


def test_descriptor(worksheet):
    from openpyxl.styles.cell_style import StyleArray
    from openpyxl.styles.fonts import Font
    from openpyxl.styles.styleable import StyleDescriptor

    class Styled:
        font = StyleDescriptor("_fonts", "fontId")

        def __init__(self):
            self._style = StyleArray()
            self.parent = worksheet

    styled = Styled()
    styled.font = Font()
    assert styled.font == Font()


@pytest.fixture
def workbook():
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
def worksheet(workbook):
    class DummyWorksheet:
        parent = workbook

    return DummyWorksheet()


def test_has_style(styleable_object):
    so = styleable_object
    so._style = None
    assert not so.has_style
    so.number_format = "dd"
    assert so.has_style


class TestNamedStyle:
    def test_assign_name(self, styleable_object):
        so = styleable_object
        wb = so.parent.parent
        style = NamedStyle(name="Standard")
        wb.add_named_style(style)
        so.style = "Standard"
        assert so._style.xfId == 0

    def test_assign_style(self, styleable_object):
        so = styleable_object
        wb = so.parent.parent
        style = NamedStyle(name="Standard")
        so.style = style
        assert so._style.xfId == 0

    def test_unknown_style(self, styleable_object):
        so = styleable_object
        with pytest.raises(ValueError):
            so.style = "Financial"

    def test_read(self, styleable_object):
        so = styleable_object
        wb = so.parent.parent
        red = NamedStyle(name="Red")
        wb.add_named_style(red)
        blue = NamedStyle(name="Blue")
        wb.add_named_style(blue)
        so._style.xfId = 1
        assert so.style == "Blue"

    def test_builtin(self, styleable_object):
        so = styleable_object
        so.style = "Hyperlink"
        assert so.style == "Hyperlink"

    def test_copy_not_share(self, styleable_object):
        s1 = styleable_object
        wb = s1.parent.parent
        s2 = copy.copy(s1)
        s1.style = "Hyperlink"
        s2.style = "Hyperlink"
        assert s1._style is not s2._style

    def test_quote_prefix(self, styleable_object):
        s1 = styleable_object
        assert s1.quotePrefix is False
        s1.quotePrefix = True
        assert s1.quotePrefix is True

    def test_pivot_button(self, styleable_object):
        s1 = styleable_object
        assert s1.pivotButton is False
        s1.pivotButton = True
        assert s1.pivotButton is True
