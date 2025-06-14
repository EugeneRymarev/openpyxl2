# Copyright (c) 2010-2025 openpyxl
import copy

import pytest


@pytest.fixture
def comment_():
    from openpyxl.comments.comments import Comment

    return Comment


class TestComment:
    def test_ctor(self, comment_):
        comment = comment_(author="Charlie", text="A comment")
        assert comment.author == "Charlie"
        assert comment.text == "A comment"
        assert comment.parent is None
        assert comment.height == 79
        assert comment.width == 144
        assert repr(comment) == "Comment: A comment by Charlie"

    def test_bind(self, comment_):
        comment = comment_("", "")
        comment.bind("ws")
        assert comment.parent == "ws"

    def test_unbind(self, comment_):
        comment = comment_("", "")
        comment.bind("ws")
        comment.unbind()
        assert comment.parent is None

    def test_copy(self, comment_):
        comment = comment_("", "")
        clone = copy.copy(comment)
        assert clone is not comment
        assert comment.text == clone.text
        assert comment.author == clone.author
        assert comment.height == clone.height
        assert comment.width == clone.width
