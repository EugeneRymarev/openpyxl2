import pytest


@pytest.fixture
def indexed_list():
    from openpyxl.utils.indexed_list import IndexedList

    return IndexedList


def test_ctor(indexed_list):
    l = indexed_list(["b", "a"])
    assert l == ["b", "a"]
    assert l.clean is False


def test_allow_duplicate_ctor(indexed_list):
    l = indexed_list(["b", "a", "b"])
    assert l == ["b", "a", "b"]
    l.append("a")
    assert l == ["b", "a", "b"]


def test_function(indexed_list):
    l = indexed_list()
    l.append("b")
    l.append("a")
    assert l == ["b", "a"]


def test_contains(indexed_list):
    l = indexed_list(["a", "b", "a"])
    assert l.clean is False
    assert "a" in l
    assert l.clean is True


def test_index(indexed_list):
    l = indexed_list(["a", "b"])
    l.append("a")
    assert l == ["a", "b"]
    l.append("c")
    assert l.index("c") == 2
    assert l.clean is True


def test_table_builder(indexed_list):
    sb = indexed_list()
    result = {"a": 0, "b": 1, "c": 2, "d": 3}
    for letter in sorted(result):
        for x in range(5):
            sb.append(letter)
        assert sb.index(letter) == result[letter]
    assert sb == ["a", "b", "c", "d"]
