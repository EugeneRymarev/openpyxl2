# Copyright (c) 2010-2024 openpyxl
"""
XML compatibility functions
"""
import re
from functools import partial

from openpyxl import DEFUSEDXML
from openpyxl import LXML

if LXML is True:
    from lxml.etree import XMLParser
    from lxml.etree import fromstring
    from lxml.etree import register_namespace
    from lxml.etree import tostring

    # do not resolve entities
    safe_parser = XMLParser(resolve_entities=False)
    fromstring = partial(fromstring, parser=safe_parser)

else:
    from xml.etree.ElementTree import fromstring
    from xml.etree.ElementTree import register_namespace
    from xml.etree.ElementTree import tostring

    if DEFUSEDXML is True:
        from defusedxml.ElementTree import fromstring

from openpyxl.xml.constants import CHART_DRAWING_NS
from openpyxl.xml.constants import CHART_NS
from openpyxl.xml.constants import COREPROPS_NS
from openpyxl.xml.constants import CUSTPROPS_NS
from openpyxl.xml.constants import DCTERMS_NS
from openpyxl.xml.constants import DCTERMS_PREFIX
from openpyxl.xml.constants import DRAWING_NS
from openpyxl.xml.constants import REL_NS
from openpyxl.xml.constants import SHEET_DRAWING_NS
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.constants import VTYPES_NS
from openpyxl.xml.constants import XML_NS

register_namespace(DCTERMS_PREFIX, DCTERMS_NS)
register_namespace("dcmitype", "http://purl.org/dc/dcmitype/")
register_namespace("cp", COREPROPS_NS)
register_namespace("c", CHART_NS)
register_namespace("a", DRAWING_NS)
register_namespace("s", SHEET_MAIN_NS)
register_namespace("r", REL_NS)
register_namespace("vt", VTYPES_NS)
register_namespace("xdr", SHEET_DRAWING_NS)
register_namespace("cdr", CHART_DRAWING_NS)
register_namespace("xml", XML_NS)
register_namespace("cust", CUSTPROPS_NS)

tostring = partial(tostring, encoding="utf-8")

NS_REGEX = re.compile(r"({(?P<namespace>.*)})?(?P<localname>.*)")


def localname(node):
    if callable(node.tag):
        return "comment"
    m = NS_REGEX.match(node.tag)
    return m.group("localname")


def whitespace(node):
    stripped = node.text.strip()
    if stripped and node.text != stripped:
        node.set(f"{{{XML_NS}}}space", "preserve")
