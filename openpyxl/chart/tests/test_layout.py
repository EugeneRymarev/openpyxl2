# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def ManualLayout():
    from ..layout import ManualLayout

    return ManualLayout


class TestManualLayout:

    def test_ctor(self, ManualLayout):
        layout = ManualLayout(
            layoutTarget="inner",
            xMode="edge",
            yMode="factor",
            wMode="factor",
            hMode="edge",
            x=0.1,
            y=0.5,
            w=0.5,
            h=0.1,
        )
        xml = tostring(layout.to_tree())
        expected = """
        <manualLayout>
          <layoutTarget val="inner"></layoutTarget>
          <xMode val="edge"></xMode>
          <yMode val="factor"></yMode>
          <wMode val="factor"></wMode>
          <hMode val="edge"></hMode>
          <x val="0.1"></x>
          <y val="0.5"></y>
          <w val="0.5"></w>
          <h val="0.1"></h>
        </manualLayout>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, ManualLayout):
        src = """
        <manualLayout>
          <layoutTarget val="inner"></layoutTarget>
          <xMode val="edge"></xMode>
          <yMode val="factor"></yMode>
          <wMode val="factor"></wMode>
          <hMode val="edge"></hMode>
          <x val="0.1"></x>
          <y val="0.5"></y>
          <w val="0.5"></w>
          <h val="0.1"></h>
        </manualLayout>
        """
        node = fromstring(src)
        layout = ManualLayout.from_tree(node)
        assert layout == ManualLayout(
            layoutTarget="inner",
            xMode="edge",
            yMode="factor",
            wMode="factor",
            hMode="edge",
            x=0.1,
            y=0.5,
            w=0.5,
            h=0.1,
        )


class TestLayout:

    def test_ctor(self):
        from ..layout import Layout

        layout = Layout()
        xml = tostring(layout.to_tree())
        diff = compare_xml(xml, "<layout />")
        assert diff is None, diff
