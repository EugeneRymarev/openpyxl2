# Copyright (c) 2010-2025 openpyxl
from warnings import warn

from openpyxl.chart.chartspace import ChartSpace
from openpyxl.chart.reader import read_chart
from openpyxl.drawing.image import ImageGroup
from openpyxl.drawing.image import PILImage
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.packaging.relationship import get_dependents
from openpyxl.packaging.relationship import get_rel
from openpyxl.packaging.relationship import get_rels_path
from openpyxl.xml.functions import fromstring


def find_images(archive, path):
    """
    Given the path to a drawing file extract charts and images and supported shapes

    Ignore errors due to unsupported parts of DrawingML
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
        warn(
            f"DrawingML support is incomplete and limited to charts and images only."
            + "Shapes and other elements may be lost from {path}."
        )
        return charts, images, shapes

    shapes = drawing._shapes

    for shape in shapes:
        link = getattr(shape.nvSpPr.cNvPr, "hlinkClick")
        if link:
            link.target = deps.get(link.id).Target
            link.mode = deps.get(link.id).TargetMode
            link.id = None

    for rel in drawing._chart_rels:
        try:
            cs = get_rel(archive, deps, rel.id, ChartSpace)
        except TypeError as e:
            warn(f"Unable to read chart {rel.id} from {path} {e}")
            continue
        chart = read_chart(cs)
        chart.anchor = rel.anchor
        if rel.anchor.graphicFrame.props.non_visual_props.hidden:
            chart.hidden = True
        charts.append(chart)

    if not PILImage:  # Pillow not installed, drop images
        return charts, images, shapes

    for blip in drawing._blip_rels:
        image = blip._read(deps, archive)
        if image is not None:
            image.anchor = blip.anchor
            images.append(image)

    for group in drawing._group_rels:
        img_group = ImageGroup()
        img_group.anchor = group.pop(0)
        for blip in group:
            image = blip._read(deps, archive)
            image.properties = (
                blip.properties
            )  # need xfrm to position the image within the anchor
            img_group.append(image)
        images.append(img_group)

    return charts, images, shapes
