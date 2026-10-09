"""Checks of the proposal's pure coordinate model, not a Worksheet fix."""
import pytest

from axis_edit import AxisEdit


@pytest.mark.parametrize("index,count,expected", [
    (1, 2, (5, 9)), (3, 2, (5, 9)), (4, 2, (3, 9)),
    (7, 2, (3, 9)), (8, 2, (3, 7)), (3, 0, (3, 7)),
])
def test_insertion(index, count, expected):
    assert AxisEdit(index, count).interval(3, 7) == expected


@pytest.mark.parametrize("index,count,expected", [
    (1, 2, (1, 5)), (3, 2, (3, 5)), (4, 2, (3, 5)),
    (6, 5, (3, 5)), (2, 8, None), (8, 1, (3, 7)), (3, 0, (3, 7)),
])
def test_deletion(index, count, expected):
    assert AxisEdit(index, count, deleting=True).interval(3, 7) == expected


def test_small_domain_against_independent_point_oracle():
    # Enumerate points, then take the surviving span. For insertion this span
    # includes new coordinates between the original endpoints.
    for deleting in (False, True):
        for index in range(1, 13):
            for count in range(5):
                edit = AxisEdit(index, count, deleting=deleting, limit=30)
                for start in range(1, 13):
                    for end in range(start, 13):
                        survivors = []
                        for value in range(start, end + 1):
                            if deleting:
                                if index <= value < index + count:
                                    continue
                                mapped = value - count if value >= index + count else value
                            else:
                                mapped = value + count if value >= index else value
                            survivors.append(mapped)
                        expected = (min(survivors), max(survivors)) if survivors else None
                        assert edit.interval(start, end) == expected


@pytest.mark.parametrize("kwargs,error", [
    ({"index": 0}, ValueError), ({"index": True}, TypeError),
    ({"index": 1, "count": -1}, ValueError),
    ({"index": 1, "count": 1.5}, TypeError),
    ({"index": 10, "count": 2, "deleting": True, "limit": 10}, ValueError),
    ({"index": 1, "deleting": 1}, TypeError),
])
def test_invalid_edits(kwargs, error):
    with pytest.raises(error):
        AxisEdit(**kwargs)


def test_overflow_and_invalid_references():
    edit = AxisEdit(4, 2, limit=10)
    with pytest.raises(OverflowError):
        edit.interval(2, 10)
    with pytest.raises(ValueError):
        edit.interval(5, 2)
    with pytest.raises(ValueError):
        edit.point(0)
    with pytest.raises(TypeError):
        edit.point(True)
    assert AxisEdit(4, 2, deleting=True, limit=10).point(4) is None
