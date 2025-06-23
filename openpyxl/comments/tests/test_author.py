# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def author_list():
    from openpyxl.comments.author import AuthorList

    return AuthorList


class TestAuthor:
    def test_ctor(self, author_list):
        vals = ["Bob", "Alice", "Eve"]
        authors = author_list(author=vals)
        xml = tostring(authors.to_tree())
        expected = """
        <authors>
            <author>Bob</author>
            <author>Alice</author>
            <author>Eve</author>
        </authors>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, author_list):
        src = """
        <authors>
            <author>author2</author>
            <author>author</author>
            <author>author3</author>
        </authors>
        """
        node = fromstring(src)
        author = author_list.from_tree(node)
        assert author.author == ["author2", "author", "author3"]
