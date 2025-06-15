# Copyright (c) 2010-2025 openpyxl
import pytest


@pytest.fixture
def bound_dictionary():
    from openpyxl.utils.bound_dictionary import BoundDictionary

    return BoundDictionary


@pytest.mark.parametrize("default", (None, int))
def test_ctor(bound_dictionary, default):
    bd = bound_dictionary("parent", default)
    assert bd.reference == "parent"
    assert bd.default_factory == default


def test_coupling(bound_dictionary):
    class Child:
        def __init__(self, parent, index=None):
            self.parent = parent
            self.index = index

    class Parent:
        def __init__(self):
            self.children = bound_dictionary("index", self._add_child)

        def _add_child(self):
            return Child(self)

    p = Parent()
    child = p.children["A"]
    assert child.parent == p
    assert child.index == "A"
