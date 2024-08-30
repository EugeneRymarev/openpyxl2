# Copyright (c) 2010-2024 openpyxl
from openpyxl.cell.rich_text import CellRichText
from openpyxl.cell.rich_text import TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.reader.strings import read_rich_text
from openpyxl.reader.strings import read_string_table
from openpyxl.styles.colors import Color


def test_read_string_table(datadir):
    datadir.chdir()
    src = "sharedStrings.xml"
    with open(src, "rb") as content:
        assert read_string_table(content) == [
            "This is cell A1 in Sheet 1",
            "This is cell G5",
        ]


def test_empty_string(datadir):
    datadir.chdir()
    src = "sharedStrings-emptystring.xml"
    with open(src, "rb") as content:
        assert read_string_table(content) == ["Testing empty cell", ""]


def test_formatted_string_table(datadir):
    datadir.chdir()
    src = "shared-strings-rich.xml"
    with open(src, "rb") as content:
        assert repr(read_rich_text(content)) == repr(
            [
                "Welcome",
                CellRichText(
                    [
                        "to the best ",
                        TextBlock(
                            font=InlineFont(
                                rFont="Calibri",
                                sz="11",
                                family="2",
                                scheme="minor",
                                color=Color(theme=1),
                                b=True,
                            ),
                            text="shop in ",
                        ),
                        TextBlock(
                            font=InlineFont(
                                rFont="Calibri",
                                sz="11",
                                family="2",
                                scheme="minor",
                                color=Color(theme=1),
                                b=True,
                                u="single",
                            ),
                            text="town",
                        ),
                    ]
                ),
                "     let's play ",
            ]
        )
