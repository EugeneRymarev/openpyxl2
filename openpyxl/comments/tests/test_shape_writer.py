# Copyright (c) 2010-2024 openpyxl
from lxml.etree import Element

from ..comment_sheet import CommentRecord
from ..comments import Comment
from ..shape_writer import excelns
from ..shape_writer import ShapeWriter
from ..shape_writer import vmlns
from openpyxl.tests.helper import compare_xml
from openpyxl.workbook import Workbook
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


def create_comments():
    wb = Workbook()
    ws = wb.active
    comment1 = Comment("text", "author")
    comment2 = Comment("text2", "author2")
    comment3 = Comment("text3", "author3")
    ws["B2"].comment = comment1
    ws["C7"].comment = comment2
    ws["D9"].comment = comment3

    comments = []
    for coord, cell in sorted(ws._cells.items()):
        if cell._comment is not None:
            comment = CommentRecord.from_cell(cell)
            comments.append((cell.coordinate, comment))

    return comments


def test_merge_comments_vml(datadir):
    datadir.chdir()
    cw = ShapeWriter(create_comments())

    with open("control+comments.vml", "rb") as existing:
        content = fromstring(cw.write(fromstring(existing.read())))
    assert len(content.findall(f"{{{vmlns}}}shape")) == 5
    assert len(content.findall(f"{{{vmlns}}}shapetype")) == 2


def test_write_comments_vml(datadir):
    datadir.chdir()
    cw = ShapeWriter(create_comments())

    content = cw.write(Element("xml"))
    with open("commentsDrawing1.vml", "rb") as expected:
        correct = fromstring(expected.read())
    check = fromstring(content)
    correct_ids = []
    correct_coords = []
    check_ids = []
    check_coords = []

    for i in correct.findall(f"{{{vmlns}}}shape"):
        correct_ids.append(i.attrib["id"])
        row = i.find(f"{{{excelns}}}ClientData").find(f"{{{excelns}}}Row").text
        col = i.find(f"{{{excelns}}}ClientData").find(f"{{{excelns}}}Column").text
        correct_coords.append((row, col))
        # blank the data we are checking separately
        i.attrib["id"] = "0"
        i.find(f"{{{excelns}}}ClientData").find(f"{{{excelns}}}Row").text = "0"
        i.find(f"{{{excelns}}}ClientData").find(f"{{{excelns}}}Column").text = "0"

    for i in check.findall(f"{{{vmlns}}}shape"):
        check_ids.append(i.attrib["id"])
        row = i.find(f"{{{excelns}}}ClientData").find(f"{{{excelns}}}Row").text
        col = i.find(f"{{{excelns}}}ClientData").find(f"{{{excelns}}}Column").text
        check_coords.append((row, col))
        # blank the data we are checking separately
        i.attrib["id"] = "0"
        i.find(f"{{{excelns}}}ClientData").find(f"{{{excelns}}}Row").text = "0"
        i.find(f"{{{excelns}}}ClientData").find(f"{{{excelns}}}Column").text = "0"

    assert set(correct_coords) == set(check_coords)
    assert set(correct_ids) == set(check_ids)
    diff = compare_xml(tostring(correct), tostring(check))
    assert diff is None, diff


def test_shape(datadir):
    from ..shape_writer import _shape_factory

    datadir.chdir()
    with open("size+comments.vml", "rb") as existing:
        expected = existing.read()

    shape = _shape_factory(2, 3, 79, 144)
    xml = tostring(shape)

    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_shape_with_custom_size(datadir):
    from ..shape_writer import _shape_factory

    datadir.chdir()
    with open("size+comments.vml", "rb") as existing:
        expected = existing.read()
        # Change our source document for this test
        expected = expected.replace(b"width:144px;", b"width:80px;")
        expected = expected.replace(b"height:79px;", b"height:20px;")

    shape = _shape_factory(2, 3, 20, 80)
    xml = tostring(shape)

    diff = compare_xml(xml, expected)
    assert diff is None, diff
