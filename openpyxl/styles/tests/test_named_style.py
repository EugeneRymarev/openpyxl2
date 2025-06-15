# Copyright (c) 2010-2025 openpyxl
import array

import pytest
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border
from openpyxl.styles.cell_style import CellStyle
from openpyxl.styles.cell_style import StyleArray
from openpyxl.styles.fills import PatternFill
from openpyxl.styles.fonts import Font
from openpyxl.styles.protection import Protection
from openpyxl.tests.helper import compare_xml
from openpyxl.workbook.workbook import Workbook
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def named_style():
    from openpyxl.styles.named_styles import NamedStyle

    return NamedStyle


@pytest.fixture
def _named_cell_style():
    from openpyxl.styles.named_styles import _NamedCellStyle

    return _NamedCellStyle


@pytest.fixture
def _named_cell_style_list():
    from openpyxl.styles.named_styles import _NamedCellStyleList

    return _NamedCellStyleList


@pytest.fixture
def named_style_list():
    from openpyxl.styles.named_styles import NamedStyleList

    return NamedStyleList


class TestNamedStyle:
    def test_ctor(self, named_style):
        style = named_style()
        assert style.font == Font()
        assert style.border == Border()
        assert style.fill == PatternFill()
        assert style.protection == Protection()
        assert style.alignment == Alignment()
        assert style.number_format == "General"
        assert style._wb is None

    def test_dict(self, named_style):
        style = named_style()
        assert dict(style) == {"name": "Normal", "hidden": "0"}

    def test_bind(self, named_style):
        style = named_style()
        wb = Workbook()
        style.bind(wb)
        assert style._wb is wb

    def test_as_tuple(self, named_style):
        style = named_style()
        assert style.as_tuple() == array.array("i", (0, 0, 0, 0, 0, 0, 0, 0, 0))

    def test_as_xf(self, named_style):
        style = named_style()
        style.alignment = Alignment(horizontal="left")
        xf = style.as_xf()
        expected = CellStyle(
            numFmtId=0,
            fontId=0,
            fillId=0,
            borderId=0,
            applyNumberFormat=None,
            applyFont=None,
            applyFill=None,
            applyBorder=None,
            applyAlignment=True,
            applyProtection=None,
            alignment=Alignment(horizontal="left"),
            protection=None,
        )
        assert xf == expected

    def test_as_name(self, named_style, _named_cell_style):
        style = named_style()
        name = style.as_name()
        assert name == _named_cell_style(name="Normal", xfId=0, hidden=False)

    @pytest.mark.parametrize(
        "attr, key, collection, expected",
        [
            ("font", "fontId", "_fonts", 0),
            ("fill", "fillId", "_fills", 0),
            ("border", "borderId", "_borders", 0),
            ("alignment", "alignmentId", "_alignments", 0),
            ("protection", "protectionId", "_protections", 0),
            ("number_format", "numFmtId", "_number_formats", 164),
        ],
    )
    def test_recalculate(self, named_style, attr, key, collection, expected):
        style = named_style()
        wb = Workbook()
        wb._number_formats.append("###")
        style.bind(wb)
        style._style = StyleArray([1, 1, 1, 1, 1, 1, 1, 1, 1])
        obj = getattr(wb, collection)[0]
        setattr(style, attr, obj)
        assert getattr(style._style, key) == expected

    def test_no_mutable_defaults(self, named_style):
        ns1 = named_style()
        ns2 = named_style()
        for attr in ("font", "fill", "border", "alignment", "protection"):
            assert getattr(ns1, attr) is not getattr(ns2, attr)


class TestNamedCellStyle:
    def test_ctor(self, _named_cell_style):
        named_style = _named_cell_style(xfId=0, name="Normal", builtinId=0)
        xml = tostring(named_style.to_tree())
        expected = '<cellStyle name="Normal" xfId="0" builtinId="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, _named_cell_style):
        src = """
        <cellStyle
                name="Followed Hyperlink"
                xfId="10"
                builtinId="9"
                hidden="1"/>
        """
        node = fromstring(src)
        ns = _named_cell_style.from_tree(node)
        expected = _named_cell_style(
            name="Followed Hyperlink",
            xfId=10,
            builtinId=9,
            hidden=True,
        )
        assert ns == expected


class TestNamedCellStyleList:
    def test_ctor(self, _named_cell_style_list):
        styles = _named_cell_style_list()
        xml = tostring(styles.to_tree())
        expected = '<cellStyles count ="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, _named_cell_style_list):
        src = "<cellStyles/>"
        node = fromstring(src)
        styles = _named_cell_style_list.from_tree(node)
        assert styles == _named_cell_style_list()

    def test_duplicate_names(self, _named_cell_style_list):
        src = """
        <cellStyles count="11">
            <cellStyle name="Followed Hyperlink" xfId="2" builtinId="9" hidden="1"/>
            <cellStyle name="Followed Hyperlink" xfId="4" builtinId="9" hidden="1"/>
            <cellStyle name="Followed Hyperlink" xfId="6" builtinId="9" hidden="1"/>
            <cellStyle name="Followed Hyperlink" xfId="8" builtinId="9" hidden="1"/>
            <cellStyle name="Followed Hyperlink" xfId="10" builtinId="9" hidden="1"/>
            <cellStyle name="Hyperlink" xfId="1" builtinId="8" hidden="1"/>
            <cellStyle name="Hyperlink" xfId="3" builtinId="8" hidden="1"/>
            <cellStyle name="Hyperlink" xfId="5" builtinId="8" hidden="1"/>
            <cellStyle name="Hyperlink" xfId="7" builtinId="8" hidden="1"/>
            <cellStyle name="Hyperlink" xfId="9" builtinId="8" hidden="1"/>
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
        """
        node = fromstring(src)
        styles = _named_cell_style_list.from_tree(node)
        cleaned = styles.remove_duplicates()
        expected = ["Normal", "Hyperlink", "Followed Hyperlink"]
        assert [s.name for s in cleaned] == expected

    def test_duplicate_ids(self, _named_cell_style_list):
        src = """
        <cellStyles count="18">
            <cellStyle name="Column0Style" xfId="1"/>
            <cellStyle name="Column10Style" xfId="1"/>
            <cellStyle name="Column11Style" xfId="1"/>
            <cellStyle name="Column12Style" xfId="4"/>
            <cellStyle name="Column13Style" xfId="4"/>
            <cellStyle name="Column1Style" xfId="1"/>
            <cellStyle name="Column2Style" xfId="3"/>
            <cellStyle name="Column3Style" xfId="4"/>
            <cellStyle name="Column4Style" xfId="1"/>
            <cellStyle name="Column5Style" xfId="1"/>
            <cellStyle name="Column6Style" xfId="1"/>
            <cellStyle name="Column7Style" xfId="1"/>
            <cellStyle name="Column8Style" xfId="1"/>
            <cellStyle name="Column9Style" xfId="1"/>
            <cellStyle name="Heading" xfId="2"/>
            <cellStyle name="Hyperlink 2" xfId="6"/>
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
            <cellStyle name="Normal 2" xfId="5"/>
        </cellStyles>
        """
        node = fromstring(src)
        styles = _named_cell_style_list.from_tree(node)
        cleaned = styles.remove_duplicates()
        expected = [
            "Normal",
            "Column0Style",
            "Heading",
            "Column2Style",
            "Column12Style",
            "Normal 2",
            "Hyperlink 2",
        ]
        assert [s.name for s in cleaned] == expected


class TestNamedStyleList:
    def test_append_valid(self, named_style, named_style_list):
        styles = named_style_list()
        style = named_style(name="special")
        styles.append(style)
        assert style in styles

    def test_append_invalid(self, named_style_list):
        styles = named_style_list()
        with pytest.raises(TypeError):
            styles.append(1)

    def test_duplicate(self, named_style_list, named_style):
        styles = named_style_list()
        style = named_style(name="special")
        styles.append(style)
        with pytest.raises(ValueError):
            styles.append(style)

    def test_names(self, named_style_list, named_style):
        styles = named_style_list()
        style = named_style(name="special")
        styles.append(style)
        assert styles.names == ["special"]

    def test_idx(self, named_style_list, named_style):
        styles = named_style_list()
        style = named_style(name="special")
        styles.append(style)
        assert styles[0] == style

    def test_key(self, named_style_list, named_style):
        styles = named_style_list()
        style = named_style(name="special")
        styles.append(style)
        assert styles["special"] == style

    def test_key_error(self, named_style_list):
        styles = named_style_list()
        with pytest.raises(KeyError):
            styles["special"]
