# Copyright (c) 2010-2025 openpyxl
import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def input_cells():
    from openpyxl.worksheet.scenario import InputCells

    return InputCells


@pytest.fixture
def scenario():
    from openpyxl.worksheet.scenario import Scenario

    return Scenario


@pytest.fixture
def scenario_list():
    from openpyxl.worksheet.scenario import ScenarioList

    return ScenarioList


class TestInputCells:
    def test_ctor(self, input_cells):
        fut = input_cells(r="B2", val="50000")
        xml = tostring(fut.to_tree())
        expected = '<inputCells r="B2" val="50000" deleted="0" undone="0"/>'
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, input_cells):
        src = '<inputCells r="B3" val="12000"/>'
        node = fromstring(src)
        fut = input_cells.from_tree(node)
        assert fut == input_cells(r="B3", val="12000")


class TestScenario:
    def test_ctor(self, scenario, input_cells):
        c1 = input_cells(r="B2", val="50000")
        c2 = input_cells(r="B3", val="12200")
        fut = scenario(inputCells=[c1, c2], name="Worst case", locked=True)
        xml = tostring(fut.to_tree())
        expected = """
        <scenario name="Worst case" locked="1" count="2" hidden="0">
            <inputCells r="B2" val="50000" deleted="0" undone="0"/>
            <inputCells r="B3" val="12200" deleted="0" undone="0"/>
        </scenario>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, scenario, input_cells):
        src = """
        <scenario name="Worst case" locked="1" count="2">
            <inputCells r="B2" val="50000"/>
            <inputCells r="B3" val="12200"/>
        </scenario>
        """
        node = fromstring(src)
        fut = scenario.from_tree(node)
        c1 = input_cells(r="B2", val="50000")
        c2 = input_cells(r="B3", val="12200")
        assert fut == scenario(inputCells=[c1, c2], name="Worst case", locked=True)


class TestScenarios:
    def test_ctor(self, scenario_list, scenario, input_cells):
        c1 = input_cells(r="B2", val="50000")
        s = scenario(name="Worst case", inputCells=[c1], locked=True, user="User")
        fut = scenario_list(scenario=[s])
        xml = tostring(fut.to_tree())
        expected = """
        <scenarios>
            <scenario name="Worst case" locked="1" hidden="0" count="1" user="User">
                <inputCells r="B2" val="50000" deleted="0" undone="0"/>
            </scenario>
        </scenarios>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, scenario_list, scenario, input_cells):
        src = """
        <scenarios current="0" show="0">
            <scenario name="Best case" locked="1" count="1" user="User">
                <inputCells r="B2" val="50000"/>
            </scenario>
        </scenarios>
        """
        node = fromstring(src)
        fut = scenario_list.from_tree(node)
        c1 = input_cells(r="B2", val="50000")
        s = scenario(name="Best case", inputCells=[c1], locked=True, user="User")
        assert fut == scenario_list(scenario=[s], current=0, show=0)
