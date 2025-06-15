# Copyright (c) 2010-2025 openpyxl
from lxml.doctestcompare import PARSE_XML
from lxml.doctestcompare import LXMLOutputChecker


def compare_xml(generated, expected):
    """
    Use doctest checking from lxml for comparing XML trees.
    Returns diff if the two are not the same
    """

    class DummyDocTest:
        pass

    checker = LXMLOutputChecker()
    ob = DummyDocTest()
    ob.want = expected
    check = checker.check_output(expected, generated, PARSE_XML)
    if not check:
        diff = checker.output_difference(ob, generated, PARSE_XML)
        return diff
    return None
