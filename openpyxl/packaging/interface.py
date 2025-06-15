# Copyright (c) 2010-2025 openpyxl
import abc

from openpyxl.compat.abc import ABC


class ISerialisableFile(ABC):
    """
    Interface for Serialisable classes that represent files in the archive
    """

    @abc.abstractproperty
    def id(self):
        """
        Object id making it unique
        """
        pass

    @abc.abstractproperty
    def _path(self):
        """
        File path in the archive
        """
        pass

    @abc.abstractproperty
    def _namespace(self):
        """
        Qualified namespace when serialised
        """
        pass

    @abc.abstractproperty
    def _type(self):
        """
        The content type for the manifest
        """

    @abc.abstractproperty
    def _rel_type(self):
        """
        The content type for relationships
        """

    @abc.abstractproperty
    def _rel_id(self):
        """
        Links object with parent
        """
