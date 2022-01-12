from __future__ import absolute_import
# Copyright (c) 2010-2021 openpyxl


from io import BytesIO
from warnings import warn

from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import IMAGE_NS, REL_NS
from openpyxl.packaging.relationship import get_rel, get_rels_path, get_dependents
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.drawing.image import Image, PILImage, ImageGroup
from openpyxl.chart.chartspace import ChartSpace
from openpyxl.chart.reader import read_chart


def find_images(archive, path):
    """
    Given the path to a drawing file extract charts and images and supported shapes

    Ingore errors due to unsupported parts of DrawingML
    """

    charts = []
    images = []
    shapes = []

    rels_path = get_rels_path(path)
    deps = []
    if rels_path in archive.namelist():
        deps = get_dependents(archive, rels_path)

    src = archive.read(path)
    tree = fromstring(src)
    try:
        drawing = SpreadsheetDrawing.from_tree(tree)
    except TypeError:
        warn(f"DrawingML support is incomplete and limited to charts and images only." +
             "Shapes and other elements may be lost from {path}.")
        return charts, images, shapes

    shapes = drawing._shapes

    for shape in shapes:
        link = getattr(shape.nvSpPr.cNvPr, "hlinkClick")
        if link:
            link.target = deps[link.id].Target
            link.mode = deps[link.id].TargetMode
            link.id = None

    for rel in drawing._chart_rels:
        cs = get_rel(archive, deps, rel.id, ChartSpace)
        chart = read_chart(cs)
        chart.anchor = rel.anchor
        charts.append(chart)

    if not PILImage: # Pillow not installed, drop images
        return charts, images, shapes

    for rel in drawing._blip_rels:
        dep = deps[rel.embed]
        if dep.Type == IMAGE_NS:
            try:
                image = Image(BytesIO(archive.read(dep.target)))
            except OSError:
                msg = "The image {0} will be removed because it cannot be read".format(dep.target)
                warn(msg)
                continue
            #if image.format.upper() == "WMF": # cannot save
                #msg = f"{image.format} image format is not supported so the image {dep.target} is being dropped"
                #warn(msg)
                #continue
            image.anchor = rel.anchor
            images.append(image)

    for group in drawing._group_rels:
        img_group = ImageGroup()
        img_group.anchor = group.pop(0)
        for rel in group:
            dep = deps[rel.embed]
            try:
                image = Image(BytesIO(archive.read(dep.target)))
            except OSError:
                msg = "The image {0} will be removed because it cannot be read".format(dep.target)
                warn(msg)
                continue
            image.properties = rel.properties # need xfrm to position the image within the anchor
            img_group.append(image)
        images.append(img_group)

    return charts, images, shapes
