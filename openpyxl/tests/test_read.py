# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.reader.excel import load_workbook
from openpyxl.styles.numbers import FORMAT_DATE_TIME3
from openpyxl.styles.numbers import FORMAT_DATE_XLSX14
from openpyxl.styles.numbers import FORMAT_GENERAL
from openpyxl.styles.numbers import FORMAT_NUMBER_00
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00


@pytest.mark.parametrize(
    "cell, number_format",
    [
        ("A1", FORMAT_GENERAL),
        ("A2", FORMAT_DATE_XLSX14),
        ("A3", FORMAT_NUMBER_00),
        ("A4", FORMAT_DATE_TIME3),
        ("A5", FORMAT_PERCENTAGE_00),
    ],
)
def test_read_general_style(datadir, cell, number_format):
    datadir.join("genuine").chdir()
    wb = load_workbook("empty-with-styles.xlsx")
    ws = wb["Sheet1"]
    assert ws[cell].number_format == number_format


def test_read_no_theme(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook("libreoffice_nrt.xlsx")
    assert wb
