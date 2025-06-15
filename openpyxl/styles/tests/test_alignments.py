# Copyright (c) 2010-2025 openpyxl
import pytest


@pytest.fixture
def alignment():
    from openpyxl.styles.alignment import Alignment

    return Alignment


def test_default(alignment):
    al = alignment()
    assert dict(al) == {}


def test_round_trip(alignment):
    args = {
        "horizontal": "center",
        "vertical": "top",
        "textRotation": "45",
        "indent": "4",
    }
    al = alignment(**args)
    assert dict(al) == args


def test_alias(alignment):
    al = alignment(text_rotation=90, shrink_to_fit=True, wrap_text=True)
    assert dict(al) == {"textRotation": "90", "shrinkToFit": "1", "wrapText": "1"}
