# Copyright (c) 2010-2025 openpyxl
import pytest


@pytest.fixture
def protection():
    from openpyxl.styles.protection import Protection

    return Protection


def test_default(protection):
    pt = protection()
    assert dict(pt) == {"hidden": "0", "locked": "1"}


def test_round_trip(protection):
    args = {"hidden": "1", "locked": "1"}
    pt = protection(**args)
    assert dict(pt) == args
