# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.styles.colors import BLACK
from openpyxl.styles.colors import WHITE
from openpyxl.styles.colors import Color
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def gradient_fill():
    from openpyxl.styles.fills import GradientFill

    return GradientFill


@pytest.fixture
def stop():
    from openpyxl.styles.fills import Stop

    return Stop


@pytest.fixture
def pattern_fill():
    from openpyxl.styles.fills import PatternFill

    return PatternFill


class TestGradientFill:
    def test_empty_ctor(self, gradient_fill):
        gf = gradient_fill()
        assert gf.type == "linear"
        assert gf.degree == 0
        assert gf.left == 0
        assert gf.right == 0
        assert gf.top == 0
        assert gf.bottom == 0
        assert gf.stop == []

    def test_ctor(self, gradient_fill):
        gf = gradient_fill(degree=90, left=1, right=2, top=3, bottom=4)
        assert gf.degree == 90
        assert gf.left == 1
        assert gf.right == 2
        assert gf.top == 3
        assert gf.bottom == 4

    @pytest.mark.parametrize("colors", [[Color(BLACK), Color(WHITE)], [BLACK, WHITE]])
    def test_stop_sequence(self, gradient_fill, stop, colors):
        gf = gradient_fill(stop=[stop(colors[0], 0), stop(colors[1], 0.5)])
        assert gf.stop[0].color.rgb == BLACK
        assert gf.stop[1].color.rgb == WHITE
        assert gf.stop[0].position == 0
        assert gf.stop[1].position == 0.5

    @pytest.mark.parametrize(
        "colors,rgbs,positions",
        [
            ([Color(BLACK), Color(WHITE)], [BLACK, WHITE], [0, 1]),
            ([BLACK, WHITE], [BLACK, WHITE], [0, 1]),
            ([BLACK, WHITE, BLACK], [BLACK, WHITE, BLACK], [0, 0.5, 1]),
            ([WHITE], [WHITE], [0]),
        ],
    )
    def test_color_sequence(self, stop, colors, rgbs, positions):
        from openpyxl.styles.fills import _assign_position

        stops = _assign_position(colors)
        assert [stop.color.rgb for stop in stops] == rgbs
        assert [stop.position for stop in stops] == positions

    def test_invalid_stop_color_mix(self, stop):
        from openpyxl.styles.fills import _assign_position

        with pytest.raises(ValueError):
            _assign_position([stop(BLACK, 0.1), WHITE])

    def test_duplicate_position(self, stop):
        from openpyxl.styles.fills import _assign_position

        with pytest.raises(ValueError):
            _assign_position([stop(BLACK, 0.5), stop(BLACK, 0.5)])

    def test_dict_interface(self, gradient_fill):
        gf = gradient_fill(degree=90, left=1, right=2, top=3, bottom=4)
        expected = {
            "bottom": "4",
            "degree": "90",
            "left": "1",
            "right": "2",
            "top": "3",
            "type": "linear",
        }
        assert dict(gf) == expected

    def test_serialise(self, gradient_fill, stop):
        gf = gradient_fill(
            degree=90,
            left=1,
            right=2,
            top=3,
            bottom=4,
            stop=[stop(BLACK, 0), stop(WHITE, 1)],
        )
        xml = tostring(gf.to_tree())
        expected = """
        <fill>
            <gradientFill
                    bottom="4"
                    degree="90"
                    left="1"
                    right="2"
                    top="3"
                    type="linear">
                <stop position="0">
                    <color rgb="00000000"></color>
                </stop>
                <stop position="1">
                    <color rgb="00FFFFFF"></color>
                </stop>
            </gradientFill>
        </fill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_create(self, gradient_fill, stop):
        src = """
        <fill>
            <gradientFill
                    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    degree="90">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        """
        xml = fromstring(src)
        fill = gradient_fill.from_tree(xml)
        expected = [stop(Color(theme=0), position=0), stop(Color(theme=4), position=1)]
        assert fill.stop == expected


class TestPatternFill:
    def test_ctor(self, pattern_fill):
        pf = pattern_fill()
        assert pf.patternType is None
        assert pf.fgColor == Color()
        assert pf.bgColor == Color()

    def test_dict_interface(self, pattern_fill):
        pf = pattern_fill(fill_type="solid")
        assert dict(pf) == {"patternType": "solid"}

    def test_serialise(self, pattern_fill):
        pf = pattern_fill("solid", "FF0000", "FFFF00")
        xml = tostring(pf.to_tree())
        expected = """
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="00FF0000"/>
                <bgColor rgb="00FFFF00"/>
            </patternFill>
        </fill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.parametrize(
        "src, args",
        [
            (
                """
                    <fill>
                        <patternFill
                                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                                patternType="solid">
                            <fgColor theme="0" tint="-0.14999847407452621"/>
                            <bgColor indexed="64"/>
                        </patternFill>
                    </fill>
                    """,
                dict(
                    patternType="solid",
                    start_color=Color(theme=0, tint=-0.14999847407452621),
                    end_color=Color(indexed=64),
                ),
            ),
            (
                """
                    <fill>
                        <patternFill
                                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                                patternType="solid">
                            <fgColor theme="0"/>
                            <bgColor indexed="64"/>
                        </patternFill>
                    </fill>
                    """,
                dict(
                    patternType="solid",
                    start_color=Color(theme=0),
                    end_color=Color(indexed=64),
                ),
            ),
            (
                """
                    <fill>
                        <patternFill
                                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                                patternType="solid">
                            <fgColor indexed="62"/>
                            <bgColor indexed="64"/>
                        </patternFill>
                    </fill>
                    """,
                dict(
                    patternType="solid",
                    start_color=Color(indexed=62),
                    end_color=Color(indexed=64),
                ),
            ),
        ],
    )
    def test_create(self, pattern_fill, src, args):
        xml = fromstring(src)
        assert pattern_fill.from_tree(xml) == pattern_fill(**args)


def test_create_empty_fill():
    from openpyxl.styles.fills import Fill

    src = fromstring("<fill/>")
    assert Fill.from_tree(src) is None


class TestStop:

    def test_ctor(self, stop):
        s = stop("999999", 0.5)
        xml = tostring(s.to_tree())
        expected = """
        <stop position="0.5">
            <color rgb="00999999"></color>
        </stop>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, stop):
        src = """
        <stop position=".5">
            <color rgb="00999999"></color>
        </stop>
        """
        node = fromstring(src)
        s = stop.from_tree(node)
        assert s == stop("999999", 0.5)

    @pytest.mark.parametrize("position", [0, 0.5, 1])
    def test_position_valid(self, stop, position):
        # smoke test
        stop("999999", position)

    @pytest.mark.parametrize(
        "position,exception",
        [(-0.1, ValueError), (1.1, ValueError), (None, TypeError)],
    )
    def test_position_invalid(self, stop, position, exception):
        with pytest.raises(exception):
            stop("999999", position)


def test_read_fills():
    # Make sure we pass the right class
    from openpyxl.styles.fills import Fill

    s = """
    <fills count="3"
           xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
        <fill>
            <gradientFill
                    type="path"
                    left="0.5"
                    right="0.5"
                    top="0.5"
                    bottom="0.5">
                <stop position="0">
                    <color theme="0" tint="-5.0935392315439317E-2"/>
                </stop>
                <stop position="1">
                    <color theme="0" tint="-0.25098422193060094"/>
                </stop>
            </gradientFill>
        </fill>
    </fills>
    """
    xml = fromstring(s)
    for node in xml:
        fill = Fill.from_tree(node)
