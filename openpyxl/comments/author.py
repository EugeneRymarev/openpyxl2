# Copyright (c) 2010-2025 openpyxl
from openpyxl.descriptors.base import Alias
from openpyxl.descriptors.sequence import Sequence
from openpyxl.descriptors.serialisable import Serialisable


class AuthorList(Serialisable):
    tagname = "authors"
    author = Sequence(expected_type=str)
    authors = Alias("author")

    def __init__(self, author=()):
        self.author = author
