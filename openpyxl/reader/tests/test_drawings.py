# Copyright (c) 2010-2025 openpyxl
import io
import zipfile


def test_read_charts(datadir):
    from openpyxl.reader.drawings import find_images

    datadir.chdir()
    archive = zipfile.ZipFile("sample.xlsx")
    path = "xl/drawings/drawing1.xml"
    charts = find_images(archive, path)[0]
    assert len(charts) == 6


def test_read_drawing(datadir):
    from openpyxl.reader.drawings import find_images

    datadir.chdir()
    archive = zipfile.ZipFile("sample_with_images.xlsx")
    path = "xl/drawings/drawing1.xml"
    images = find_images(archive, path)[1]
    assert len(images) == 3


def test_unsupported_drawing(datadir):
    from openpyxl.reader.drawings import find_images

    datadir.chdir()
    out = io.BytesIO()
    archive = zipfile.ZipFile(out, mode="w")
    archive.write("unsupported_drawing.xml", "drawing1.xml")
    charts, images, shapes = find_images(archive, "drawing1.xml")
    assert charts == images == shapes == []


def test_unsupported_image_format(datadir):
    from openpyxl.reader.drawings import find_images

    datadir.chdir()
    archive = zipfile.ZipFile("sample_with_unsupported_image_format.xlsx", "r")
    path = "xl/drawings/drawing1.xml"
    images = find_images(archive, path)
    assert images == ([], [], [])


def test_hyperlink(datadir):
    from openpyxl.reader.drawings import find_images

    datadir.chdir()
    archive = zipfile.ZipFile("drawing_with_hyperlink.xlsx", "r")
    path = "xl/drawings/drawing1.xml"
    shapes = find_images(archive, path)[-1]
    assert shapes[0].nvSpPr.cNvPr.hlinkClick.target == "http://www.example.org"
