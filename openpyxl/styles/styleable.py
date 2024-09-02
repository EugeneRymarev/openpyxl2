# Copyright (c) 2010-2024 openpyxl
from copy import copy

from .builtins import styles
from .cell_style import StyleArray
from .named_styles import NamedStyle
from .numbers import BUILTIN_FORMATS
from .numbers import BUILTIN_FORMATS_MAX_SIZE
from .numbers import BUILTIN_FORMATS_REVERSE
from .proxy import StyleProxy


class StyleDescriptor:

    def __init__(self, collection, key):
        self.collection = collection
        self.key = key

    def __set__(self, instance, value):
        coll = getattr(instance.parent.parent, self.collection)
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        setattr(instance._style, self.key, coll.add(value))

    def __get__(self, instance, cls):
        coll = getattr(instance.parent.parent, self.collection)
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        return StyleProxy(coll[idx])


class NumberFormatDescriptor:
    key = "numFmtId"
    collection = "_number_formats"

    def __set__(self, instance, value):
        coll = getattr(instance.parent.parent, self.collection)
        if value in BUILTIN_FORMATS_REVERSE:
            idx = BUILTIN_FORMATS_REVERSE[value]
        else:
            idx = coll.add(value) + BUILTIN_FORMATS_MAX_SIZE

        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        setattr(instance._style, self.key, idx)

    def __get__(self, instance, cls):
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        if idx < BUILTIN_FORMATS_MAX_SIZE:
            return BUILTIN_FORMATS.get(idx, "General")
        coll = getattr(instance.parent.parent, self.collection)
        return coll[idx - BUILTIN_FORMATS_MAX_SIZE]


class NamedStyleDescriptor:
    key = "xfId"
    collection = "_named_styles"

    def __set__(self, instance, value):
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        coll = getattr(instance.parent.parent, self.collection)
        if isinstance(value, NamedStyle):
            style = value
            if style not in coll:
                instance.parent.parent.add_named_style(style)
        elif value not in coll.names:
            if value in styles:  # is it builtin?
                style = styles[value]
                if style not in coll:
                    instance.parent.parent.add_named_style(style)
            else:
                raise ValueError(f"{value} is not a known style")
        else:
            style = coll[value]
        instance._style = copy(style.as_tuple())

    def __get__(self, instance, cls):
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        coll = getattr(instance.parent.parent, self.collection)
        return coll.names[idx]


class StyleArrayDescriptor:

    def __init__(self, key):
        self.key = key

    def __set__(self, instance, value):
        if instance._style is None:
            instance._style = StyleArray()
        setattr(instance._style, self.key, value)

    def __get__(self, instance, cls):
        if instance._style is None:
            return False
        return bool(getattr(instance._style, self.key))


class StyleableObject:
    """
    Base class for styleble objects implementing proxy and lookup functions
    """

    _alignment = StyleDescriptor("_alignments", "alignmentId")
    _border = StyleDescriptor("_borders", "borderId")
    _fill = StyleDescriptor("_fills", "fillId")
    _font = StyleDescriptor("_fonts", "fontId")
    _number_format = NumberFormatDescriptor()
    _protection = StyleDescriptor("_protections", "protectionId")
    quotePrefix = StyleArrayDescriptor("quotePrefix")
    pivotButton = StyleArrayDescriptor("pivotButton")

    __slots__ = ("parent", "_style")

    def __init__(self, sheet, style_array=None):
        self.parent = sheet
        if style_array is not None:
            style_array = StyleArray(style_array)
        self._style = style_array

    @property
    def style_id(self):
        if self._style is None:
            self._style = StyleArray()
        return self.parent.parent._cell_styles.add(self._style)

    @property
    def has_style(self):
        if self._style is None:
            return False
        return any(self._style)

    @property
    def alignment(self):
        return self._alignment

    @alignment.setter
    def alignment(self, value):
        self._alignment = value

    @alignment.deleter
    def alignment(self):
        self._alignment = StyleDescriptor("_alignments", "alignmentId")

    @property
    def border(self):
        return self._border

    @border.setter
    def border(self, value):
        self._border = value

    @border.deleter
    def border(self):
        self._border = StyleDescriptor("_borders", "borderId")

    @property
    def fill(self):
        return self._fill

    @fill.setter
    def fill(self, value):
        self._fill = value

    @fill.deleter
    def fill(self):
        self._fill = StyleDescriptor("_fills", "fillId")

    @property
    def font(self):
        return self._font

    @font.setter
    def font(self, value):
        self._font = value

    @font.deleter
    def font(self):
        self._font = StyleDescriptor("_fonts", "fontId")

    @property
    def number_format(self):
        return self._number_format

    @number_format.setter
    def number_format(self, value):
        self._number_format = value

    @number_format.deleter
    def number_format(self):
        self._number_format = NumberFormatDescriptor()

    @property
    def protection(self):
        return self._protection

    @protection.setter
    def protection(self, value):
        self._protection = value

    @protection.deleter
    def protection(self):
        self._protection = StyleDescriptor("_protections", "protectionId")

    @property
    def style(self):
        return NamedStyleDescriptor().__get__(self, None)

    @style.setter
    def style(self, value):
        NamedStyleDescriptor().__set__(self, value)

    @style.deleter
    def style(self):
        NamedStyleDescriptor().__set__(self, "Normal")
