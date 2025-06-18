# Copyright (c) 2010-2025 openpyxl
"""
XML compatibility functions
"""
import functools
import re

from openpyxl.xml import DEFUSEDXML
from openpyxl.xml import LXML
from openpyxl.xml.constants import ACTIVEX_NS
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

if LXML is True:
    from lxml.etree import Element
    from lxml.etree import QName
    from lxml.etree import SubElement
    from lxml.etree import XMLParser
    from lxml.etree import fromstring
    from lxml.etree import register_namespace
    from lxml.etree import tostring
    from lxml.etree import xmlfile

    # do not resolve entities
    safe_parser = XMLParser(resolve_entities=False)
    fromstring = functools.partial(fromstring, parser=safe_parser)
else:
    from xml.etree.ElementTree import Element
    from xml.etree.ElementTree import QName
    from xml.etree.ElementTree import SubElement
    from xml.etree.ElementTree import fromstring
    from xml.etree.ElementTree import register_namespace
    from xml.etree.ElementTree import tostring

    from et_xmlfile import xmlfile

    if DEFUSEDXML is True:
        from defusedxml.ElementTree import fromstring
if DEFUSEDXML is True:
    from defusedxml.ElementTree import iterparse

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
register_namespace("ax", ACTIVEX_NS)
tostring = functools.partial(tostring, encoding="utf-8")
NS_REGEX = re.compile("({(?P<namespace>.*)})?(?P<localname>.*)")


def localname(node):
    if callable(node.tag):
        return "comment"
    m = NS_REGEX.match(node.tag)
    return m.group("localname")


def whitespace(node):
    stripped = node.text.strip()
    if stripped and node.text != stripped:
        node.set(f"{{{XML_NS}}}space", "preserve")
