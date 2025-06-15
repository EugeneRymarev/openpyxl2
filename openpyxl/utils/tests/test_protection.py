# Copyright (c) 2010-2025 openpyxl
from openpyxl.utils.protection import hash_password


def test_password():
    enc = hash_password("secret")
    assert enc == "DAA7"
