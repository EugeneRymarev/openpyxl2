# Copyright (c) 2010-2025 openpyxl
"""
Generic serialisable classes
"""
from openpyxl.compat.strings import safe_string
from openpyxl.descriptors.base import Bool
from openpyxl.descriptors.base import Convertible
from openpyxl.descriptors.base import Descriptor
from openpyxl.descriptors.base import Float
from openpyxl.descriptors.base import Integer
from openpyxl.descriptors.base import MinMax
from openpyxl.descriptors.base import NoneSet
from openpyxl.descriptors.base import Set
from openpyxl.descriptors.base import String
from openpyxl.xml.functions import Element
from openpyxl.xml.functions import localname
from openpyxl.xml.functions import whitespace


class Nested(Descriptor):
    nested = True
    attribute = "val"

    def __set__(self, instance, value):
        if hasattr(value, "tag"):
            tag = localname(value)
            if tag != self.name:
                raise ValueError("Tag does not match attribute")
            value = self.from_tree(value)
        super().__set__(instance, value)

    def from_tree(self, node):
        return node.get(self.attribute)

    def to_tree(self, tagname=None, value=None, namespace=None):
        namespace = getattr(self, "namespace", namespace)
        if value is not None:
            if namespace is not None:
                tagname = f"{{{namespace}}}{tagname}"
            value = safe_string(value)
            return Element(tagname, {self.attribute: value})
        return None


class NestedValue(Nested, Convertible):
    """
    Nested tag storing the value on the 'val' attribute
    """

    pass


class NestedText(NestedValue):
    """
    Represents any nested tag with the value as the contents of the tag
    """

    def from_tree(self, node):
        return node.text

    def to_tree(self, tagname=None, value=None, namespace=None):
        namespace = getattr(self, "namespace", namespace)
        if value is not None:
            if namespace is not None:
                tagname = f"{{{namespace}}}{tagname}"
            el = Element(tagname)
            el.text = safe_string(value)
            whitespace(el)
            return el
        return None


class NestedFloat(NestedValue, Float):
    pass


class NestedInteger(NestedValue, Integer):
    pass


class NestedString(NestedValue, String):
    pass


class NestedBool(NestedValue, Bool):
    def from_tree(self, node):
        return node.get("val", True)


class NestedNoneSet(Nested, NoneSet):
    pass


class NestedSet(Nested, Set):
    pass


class NestedMinMax(Nested, MinMax):
    pass


class EmptyTag(Nested, Bool):
    """
    Boolean if a tag exists or not.
    """

    def from_tree(self, node):
        return True

    def to_tree(self, tagname=None, value=None, namespace=None):
        if value:
            namespace = getattr(self, "namespace", namespace)
            if namespace is not None:
                tagname = f"{{{namespace}}}{tagname}"
            return Element(tagname)
        return None
