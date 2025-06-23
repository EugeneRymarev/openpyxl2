# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def vol_topic_ref():
    from openpyxl.volatile.volatile import VolTopicRef

    return VolTopicRef


@pytest.fixture
def vol_main():
    from openpyxl.volatile.volatile import VolMain

    return VolMain


@pytest.fixture
def vol_type():
    from openpyxl.volatile.volatile import VolType

    return VolType


@pytest.fixture
def vol_topic():
    from openpyxl.volatile.volatile import VolTopic

    return VolTopic


@pytest.fixture
def vol_types():
    from openpyxl.volatile.volatile import VolTypesList

    return VolTypesList


class TestVolTopicRef:
    def test_ctor(self, vol_topic_ref):
        ref = vol_topic_ref(r="A1", s=3)
        xml = tostring(ref.to_tree())
        expected = '<tr r="A1" s="3"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, vol_topic_ref):
        src = '<tr r="A1" s="1"/>'
        node = fromstring(src)
        ref = vol_topic_ref.from_tree(node)
        assert ref == vol_topic_ref(r="A1", s=1)


class TestVolMain:
    def test_ctor(self, vol_main):
        main = vol_main(first="ThisDataModel")
        xml = tostring(main.to_tree())
        expected = '<main first="ThisDataModel"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, vol_main):
        from openpyxl.volatile.volatile import VolTopic

        src = """
        <main first="ThisWorkbookDataModel">
            <tp t="e">
                <v>#N/A</v>
            </tp>
        </main>
        """
        node = fromstring(src)
        main = vol_main.from_tree(node)
        expected = vol_main(
            first="ThisWorkbookDataModel",
            tp=[VolTopic(t="e", v="#N/A")],
        )
        assert main == expected


class TestVolType:
    def test_ctor(self, vol_type):
        from openpyxl.volatile.volatile import VolMain
        from openpyxl.volatile.volatile import VolTopic

        typ = vol_type(
            main=[VolMain(first="teststring", tp=[VolTopic(t="s", v="aaa: 4447")])],
            type="realTimeData",
        )
        xml = tostring(typ.to_tree())
        expected = """
        <volType type="realTimeData">
            <main first="teststring">
                <tp t="s">
                <v>aaa: 4447</v>
                </tp>
            </main>
        </volType>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, vol_type):
        src = """
        <volType type="olapFunctions">
            <main first="ThisWorkbookDataModel">
                <tp t="e">
                    <v>#N/A</v>
                    <stp>1</stp>
                    <tr r="A1" s="1"/>
                </tp>
            </main>
        </volType>
        """
        node = fromstring(src)
        typ = vol_type.from_tree(node)
        assert typ.type == "olapFunctions"
        assert typ.main[0].first == "ThisWorkbookDataModel"


class TestVolTopic:
    def test_ctor(self, vol_topic):
        topic = vol_topic(t="s", v="aaa: 4447")
        xml = tostring(topic.to_tree())
        expected = """
        <tp t="s">
            <v>aaa: 4447</v>
        </tp>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, vol_topic):
        from openpyxl.volatile.volatile import VolTopicRef

        src = """
        <tp t="e">
            <v>#N/A</v>
            <stp>1</stp>
            <tr r="A1" s="1"></tr>
        </tp>
        """
        node = fromstring(src)
        topic = vol_topic.from_tree(node)
        expected = vol_topic(t="e", v="#N/A", stp="1", tr=[VolTopicRef(r="A1", s=1)])
        assert topic == expected


class TestVolTypes:
    def test_ctor(self, vol_types):
        from openpyxl.volatile.volatile import VolMain
        from openpyxl.volatile.volatile import VolTopic
        from openpyxl.volatile.volatile import VolType

        typ = vol_types(
            volType=[
                VolType(
                    main=[
                        VolMain(
                            first="teststring",
                            tp=[VolTopic(t="s", v="aaa: 4447")],
                        )
                    ],
                    type="realTimeData",
                )
            ]
        )
        xml = tostring(typ.to_tree())
        expected = """
        <volTypes xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <volType type="realTimeData">
                <main first="teststring">
                    <tp t="s">
                        <v>aaa: 4447</v>
                    </tp>
                </main>
            </volType>
        </volTypes>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, vol_types):
        src = """
        <volTypes xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <volType type="olapFunctions">
                <main first="ThisWorkbookDataModel">
                    <tp t="e">
                        <v>#N/A</v>
                        <stp>1</stp>
                        <tr r="A1" s="1"/>
                    </tp>
                </main>
            </volType>
        </volTypes>
        """
        node = fromstring(src)
        typ = vol_types.from_tree(node)
        assert typ.volType[0].type == "olapFunctions"
        assert typ.volType[0].main[0].first == "ThisWorkbookDataModel"
