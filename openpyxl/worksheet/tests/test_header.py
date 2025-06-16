# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def _header_footer_part():
    from openpyxl.worksheet.header_footer import _HeaderFooterPart

    return _HeaderFooterPart


@pytest.fixture
def header_footer_item():
    from openpyxl.worksheet.header_footer import HeaderFooterItem

    return HeaderFooterItem


@pytest.fixture
def header_footer():
    from openpyxl.worksheet.header_footer import HeaderFooter

    return HeaderFooter


def test_split_into_parts():
    from openpyxl.worksheet.header_footer import _split_string

    headers = _split_string("&Ltest header")
    assert headers["left"] == "test header"
    s = (
        '&L&"Lucida Grande,Standard"&K000000Left top&C&'
        '"Lucida Grande,Standard"&K000000Middle top&R&'
        '"Lucida Grande,Standard"&K000000Right top'
    )
    headers = _split_string(s)
    assert headers["left"] == '&"Lucida Grande,Standard"&K000000Left top'
    assert headers["center"] == '&"Lucida Grande,Standard"&K000000Middle top'
    assert headers["right"] == '&"Lucida Grande,Standard"&K000000Right top'


def test_cannot_split():
    from openpyxl.worksheet.header_footer import _split_string

    s = "\n "
    parts = _split_string(s)
    assert parts == {"left": "", "right": "", "center": ""}


def test_multiline_string():
    from openpyxl.worksheet.header_footer import _split_string

    s = "&L141023 V1&CRoute - Malls\nSchedules R1201 v R1301&RClient-internal use only"
    headers = _split_string(s)
    expected = {
        "center": "Route - Malls\nSchedules R1201 v R1301",
        "left": "141023 V1",
        "right": "Client-internal use only",
    }
    assert headers == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        ("&9", [("", "", "9")]),
        ('&"Lucida Grande,Standard"', [("Lucida Grande,Standard", "", "")]),
        ("&K000000", [("", "000000", "")]),
    ],
)
def test_parse_format(value, expected):
    from openpyxl.worksheet.header_footer import FORMAT_REGEX

    m = FORMAT_REGEX.findall(value)
    assert m == expected


class TestHeaderFooterPart:
    def test_ctor(self, _header_footer_part):
        hf = _header_footer_part(
            text="secret message",
            font="Calibri,Regular",
            color="000000",
        )
        assert str(hf) == '&"Calibri,Regular"&K000000secret message'

    def test_read(self, _header_footer_part):
        s = '&"Lucida Grande,Standard"&K22BBDDLeft top&12 '
        hf = _header_footer_part.from_str(s)
        assert hf.text == "Left top"
        assert hf.font == "Lucida Grande,Standard"
        assert hf.color == "22BBDD"
        assert hf.size == 12

    def test_bool(self, _header_footer_part):
        hf = _header_footer_part()
        assert bool(hf) is False
        hf.text = "Title"
        assert bool(hf) is True

    def test_str(self, _header_footer_part):
        hf = _header_footer_part()
        hf.text = "D\xfcsseldorf"
        assert str(hf) == "D\xfcsseldorf"


class TestHeaderFooterItem:
    def test_ctor(self, header_footer_item):
        hf = header_footer_item()
        hf.left.text = "yes"
        hf.center.text = "no"
        hf.right.text = "maybe"
        assert str(hf) == "&Lyes&Cno&Rmaybe"

    def test_read(self, header_footer_item):
        s = (
            '&amp;L&amp;"Lucida Grande,Standard"&amp;K000000&amp;'
            '12 Left top&amp;C&amp;"Lucida Grande,Standard"&amp;'
            'K000000Middle top&amp;R&amp;"Lucida Grande,Standard"'
            "&amp;K000000Right top"
        )
        xml = f"<oddHeader>{s}</oddHeader>"
        node = fromstring(xml)
        hf = header_footer_item.from_tree(node)
        assert hf.left.text == "Left top"
        assert hf.center.text == "Middle top"
        assert hf.right.text == "Right top"

    def test_write(self, header_footer_item):
        hf = header_footer_item()
        hf.left.text = "A secret message"
        hf.left.size = 12
        xml = tostring(hf.to_tree("header_or_footer"))
        expected = "<header_or_footer>&amp;L&amp;12 A secret message</header_or_footer>"
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_bool(self, header_footer_item):
        hf = header_footer_item()
        assert bool(hf) is False
        hf.left.text = "Title"
        assert bool(hf) is True

    def test_str(self, header_footer_item):
        hf = header_footer_item()
        hf.left.text = "D\xfcsseldorf"
        assert str(hf) == "&LD\xfcsseldorf"


class TestHeaderFooter:
    def test_ctor(self, header_footer):
        hf = header_footer()
        xml = tostring(hf.to_tree())
        expected = """
        <headerFooter>
            <oddHeader/>
            <oddFooter/>
            <evenHeader/>
            <evenFooter/>
            <firstHeader/>
            <firstFooter/>
        </headerFooter>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, header_footer):
        s1 = (
            '&amp;L&amp;"Lucida Grande,Standard"&amp;K000000Left '
            'top&amp;C&amp;"Lucida Grande,Standard"&amp;K000000Middle '
            'top&amp;R&amp;"Lucida Grande,Standard"&amp;K000000Right top'
        )
        s2 = (
            '&amp;L&amp;"Lucida Grande,Standard"&amp;K000000Left '
            'footer&amp;C&amp;"Lucida Grande,Standard"&amp;'
            'K000000Middle Footer&amp;R&amp;"Lucida Grande,'
            'Standard"&amp;K000000Right Footer'
        )
        src = f"""
        <headerFooter>
            <oddHeader>{s1}</oddHeader>
            <oddFooter>{s2}</oddFooter>
        </headerFooter>
        """
        node = fromstring(src)
        hf = header_footer.from_tree(node)
        assert hf.oddHeader.left.text == "Left top"

    def test_bool(self, header_footer, header_footer_item):
        hf = header_footer()
        assert bool(hf) is False
        hf.oddHeader = header_footer_item()
        hf.oddHeader.left.text = "Title"
        assert bool(hf) is True
