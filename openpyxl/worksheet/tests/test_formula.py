# Copyright (c) 2010-2025 openpyxl
import pytest


@pytest.fixture
def data_table_formula():
    from openpyxl.worksheet.formula import DataTableFormula

    return DataTableFormula


@pytest.fixture
def array_formula():
    from openpyxl.worksheet.formula import ArrayFormula

    return ArrayFormula


class TestDataTableFormula:
    def test_ctor(self, data_table_formula):
        dt = data_table_formula(
            t="dataTable",
            ref="I9:S24",
            dt2D="1",
            dtr="1",
            r1="I5",
            r2="I4",
        )
        assert dt.ref == "I9:S24"

    def test_dict(self, data_table_formula):
        dt = data_table_formula(ref="A1:B6", r1="G5", dt2D=True)
        assert dict(dt) == {"ref": "A1:B6", "r1": "G5", "dt2D": "1", "t": "dataTable"}


class TestDataTable:
    def test_ctor(self, array_formula):
        af = array_formula(ref="I9:S24")
        assert af.ref == "I9:S24"

    def test_dict(self, array_formula):
        af = array_formula(ref="A1:B6")
        assert dict(af) == {"ref": "A1:B6", "t": "array"}
