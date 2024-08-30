# Copyright (c) 2010-2024 openpyxl
from lxml.etree import Element
from lxml.etree import SubElement

from openpyxl.utils import coordinate_to_tuple
from openpyxl.xml.functions import tostring

vmlns = "urn:schemas-microsoft-com:vml"
officens = "urn:schemas-microsoft-com:office:office"
excelns = "urn:schemas-microsoft-com:office:excel"


class ShapeWriter:
    """
    Create VML for comments
    """

    vml = None
    vml_path = None

    def __init__(self, comments):
        self.comments = comments

    def add_comment_shapetype(self, root):
        shape_layout = SubElement(
            root, f"{{{officens}}}shapelayout", {f"{{{vmlns}}}ext": "edit"}
        )
        SubElement(
            shape_layout,
            f"{{{officens}}}idmap",
            {f"{{{vmlns}}}ext": "edit", "data": "1"},
        )
        shape_type = SubElement(
            root,
            f"{{{vmlns}}}shapetype",
            {
                "id": "_x0000_t202",
                "coordsize": "21600,21600",
                f"{{{officens}}}spt": "202",
                "path": "m,l,21600r21600,l21600,xe",
            },
        )
        SubElement(shape_type, f"{{{vmlns}}}stroke", {"joinstyle": "miter"})
        SubElement(
            shape_type,
            f"{{{vmlns}}}path",
            {"gradientshapeok": "t", f"{{{officens}}}connecttype": "rect"},
        )

    def add_comment_shape(self, root, idx, coord, height, width):
        row, col = coordinate_to_tuple(coord)
        row -= 1
        col -= 1
        shape = _shape_factory(row, col, height, width)

        shape.set("id", f"_x0000_s{idx:04d}")
        root.append(shape)

    def write(self, root):

        if not hasattr(root, "findall"):
            root = Element("xml")

        # Remove any existing comment shapes
        comments = root.findall(f"{{{vmlns}}}shape[@type='#_x0000_t202']")
        for c in comments:
            root.remove(c)

        # check whether comments shape type already exists
        shape_types = root.find(f"{{{vmlns}}}shapetype[@id='_x0000_t202']")
        if shape_types is None:
            self.add_comment_shapetype(root)

        for idx, (coord, comment) in enumerate(self.comments, 1026):
            self.add_comment_shape(root, idx, coord, comment.height, comment.width)

        return tostring(root)


def _shape_factory(row, column, height, width):
    style = (
        "position:absolute; "
        "margin-left:59.25pt;"
        "margin-top:1.5pt;"
        f"width:{width}px;"
        f"height:{height}px;"
        "z-index:1;"
        "visibility:hidden"
    )
    attrs = {
        "type": "#_x0000_t202",
        "style": style,
        "fillcolor": "#ffffe1",
        f"{{{officens}}}insetmode": "auto",
    }
    shape = Element(f"{{{vmlns}}}shape", attrs)

    SubElement(shape, f"{{{vmlns}}}fill", {"color2": "#ffffe1"})
    SubElement(shape, f"{{{vmlns}}}shadow", {"color": "black", "obscured": "t"})
    SubElement(shape, f"{{{vmlns}}}path", {f"{{{officens}}}connecttype": "none"})
    textbox = SubElement(
        shape, f"{{{vmlns}}}textbox", {"style": "mso-direction-alt:auto"}
    )
    SubElement(textbox, "div", {"style": "text-align:left"})
    client_data = SubElement(shape, f"{{{excelns}}}ClientData", {"ObjectType": "Note"})
    SubElement(client_data, f"{{{excelns}}}MoveWithCells")
    SubElement(client_data, f"{{{excelns}}}SizeWithCells")
    SubElement(client_data, f"{{{excelns}}}AutoFill").text = "False"
    SubElement(client_data, f"{{{excelns}}}Row").text = str(row)
    SubElement(client_data, f"{{{excelns}}}Column").text = str(column)
    return shape
