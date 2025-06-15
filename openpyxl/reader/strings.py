# Copyright (c) 2010-2025 openpyxl
from openpyxl.cell.rich_text import CellRichText
from openpyxl.cell.text import Text
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import iterparse


def read_string_table(xml_source):
    """
    Read in all shared strings in the table
    """
    strings = []
    string_tag = f"{{{SHEET_MAIN_NS}}}si"
    for _, node in iterparse(xml_source):
        if node.tag == string_tag:
            text = Text.from_tree(node).content
            text = text.replace("x005F_", "")
            node.clear()
            strings.append(text)
    return strings


def read_rich_text(xml_source):
    """Read in all shared strings in the table"""

    strings = []
    string_tag = f"{{{SHEET_MAIN_NS}}}si"
    for _, node in iterparse(xml_source):
        if node.tag == string_tag:
            text = CellRichText.from_tree(node)
            if len(text) == 0:
                text = ""
            elif len(text) == 1 and isinstance(text[0], str):
                text = text[0]
            node.clear()
            strings.append(text)
    return strings
