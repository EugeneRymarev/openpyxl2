# Copyright (c) 2010-2025 openpyxl
import itertools

import numpy
import pandas
import pytest


@pytest.fixture
def sample_data():
    data = {
        "A": [0.0, 1.0, 2.0, 3.0, 4.0],
        "B": [0.0, 1.0, 0.0, 1.0, 0.0],
        "C": ["foo1", "foo2", "foo3", "foo4", "foo5"],
        "D": pandas.date_range("2009-01-01", periods=5),
    }
    df = pandas.DataFrame(data)
    df.index.name = "openpyxl test"
    df.iloc[0] = numpy.nan
    return df


@pytest.mark.pandas_required
def test_dataframe(sample_data):
    from openpyxl.utils.dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, index=False, header=False))
    assert rows[2] == [2.0, 0.0, "foo3", pandas.Timestamp("2009-01-03 00:00:00")]


@pytest.mark.pandas_required
def test_dataframe_header(sample_data):
    from openpyxl.utils.dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, index=False))
    assert rows[0] == ["A", "B", "C", "D"]


@pytest.mark.pandas_required
def test_dataframe_index(sample_data):
    from openpyxl.utils.dataframe import dataframe_to_rows

    rows = tuple(dataframe_to_rows(sample_data, header=False))
    assert rows[0] == ["openpyxl test"]


@pytest.mark.pandas_required
def test_dataframe_multiindex():
    from openpyxl.utils.dataframe import dataframe_to_rows

    arrays = [
        ["bar", "bar", "bar", "baz", "foo", "foo", "qux", "qux"],
        ["one", "two", "three", "one", "one", "two", "one", "two"],
    ]
    tuples = list(zip(*arrays))
    index = pandas.MultiIndex.from_tuples(tuples, names=["first", "second"])
    df = pandas.Series(0, index=index)
    df = pandas.DataFrame(df)
    rows = list(dataframe_to_rows(df, header=False))
    expected = [
        ["first", "second"],
        ["bar", "one", 0],
        [None, "two", 0],
        [None, "three", 0],
        ["baz", "one", 0],
        ["foo", "one", 0],
        [None, "two", 0],
        ["qux", "one", 0],
        [None, "two", 0],
    ]
    assert rows == expected


@pytest.mark.pandas_required
def test_expand_index_vertically():
    from openpyxl.utils.dataframe import expand_index

    arrays = [
        [2019, 2019, 2019, 2019, 2020, 2020, 2020, 2021, 2021, 2021, 2021],
        [
            "Major",
            "Major",
            "Minor",
            "Minor",
            "Major",
            "Major",
            "Minor",
            "Minor",
            "Major",
            "Major",
            "Minor",
            "Minor",
        ],
        ["a", "b", "a", "b", "a", "b", "a", "b", "a", "b", "a", "b"],
    ]

    tuples = list(zip(*arrays))
    index = pandas.MultiIndex.from_tuples(tuples, names=["first", "second", "third"])
    rows = list(expand_index(index))
    assert rows[0] == [2019, "Major", "a"]
    assert rows[1] == [None, None, "b"]


@pytest.mark.pandas_required
def test_expand_levels_horizontally():
    from openpyxl.utils.dataframe import expand_index

    levels = [["2016", "2017", "2018"], ["Major", "Minor"], ["a", "b"]]
    tuples = itertools.product(*levels)
    index = pandas.MultiIndex.from_tuples(tuples, names=["first", "second", "third"])
    expanded = list(expand_index(index, header=True))
    expected = [
        "2016",
        None,
        None,
        None,
        "2017",
        None,
        None,
        None,
        "2018",
        None,
        None,
        None,
    ]
    assert expanded[0] == expected
    expected = [
        "Major",
        None,
        "Minor",
        None,
        "Major",
        None,
        "Minor",
        None,
        "Major",
        None,
        "Minor",
        None,
    ]
    assert expanded[1] == expected
    assert expanded[2] == ["a", "b", "a", "b", "a", "b", "a", "b", "a", "b", "a", "b"]


@pytest.mark.pandas_required
def test_dataframe_categorical():
    from openpyxl.utils.dataframe import dataframe_to_rows

    arrays = [
        [2019, 2019, 2019, 2019, 2020, 2020, 2020, 2021, 2021, 2021, 2021, 2022],
        [
            "Major",
            "Major",
            "Minor",
            "Minor",
            "Major",
            "Major",
            "Minor",
            "Minor",
            "Major",
            "Major",
            "Minor",
            "Minor",
        ],
        ["a", "b", "a", "b", "a", "b", "a", "b", "a", "b", "a", "b"],
    ]
    df = pandas.DataFrame(arrays)
    df = df.apply(lambda col: col.astype("category"))
    rows = list(dataframe_to_rows(df, header=False, index=False))
    assert rows == arrays
