# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.chart.pivot import PivotFormat
from openpyxl.chart.pivot import PivotSource
from openpyxl.chart.plotarea import PlotArea
from openpyxl.chart.series import Series
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import tostring


@pytest.fixture
def chart_base():
    from openpyxl.chart._chart import ChartBase

    return ChartBase


class TestChartBase:
    def test_ctor(self, chart_base):
        chart = chart_base()
        with pytest.raises(NotImplementedError):
            xml = tostring(chart.to_tree())

    def test_iadd(self, chart_base):
        chart1 = chart_base()
        chart2 = chart_base()
        chart1 += chart2
        assert chart1._charts == [chart1, chart2]

    def test_invalid_add(self, chart_base):
        chart = chart_base()
        s = Series()
        with pytest.raises(TypeError):
            chart += s

    def test_set_categories(self, chart_base):
        from openpyxl.chart.series import Series

        s1 = Series()
        s1.__elements__ = ("cat",)
        chart = chart_base()
        chart.ser = [s1]
        chart.set_categories("Sheet!A1:A4")
        xml = tostring(s1.to_tree())
        expected = """
        <ser>
            <cat>
                <numRef>
                    <f>'Sheet'!$A$1:$A$4</f>
                </numRef>
            </cat>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_add_data_cols(self, chart_base):
        chart = chart_base()
        chart.ser = []
        chart.add_data("Sheet!A1:E4")
        assert len(chart.ser) == 5
        assert chart.ser[0].val.numRef.f == "'Sheet'!$A$1:$A$4"
        assert chart.ser[-1].val.numRef.f == "'Sheet'!$E$1:$E$4"

    def test_add_data_rows(self, chart_base):
        chart = chart_base()
        chart.ser = []
        chart.add_data("Sheet!A1:E4", from_rows=True)
        assert len(chart.ser) == 4
        assert chart.ser[0].val.numRef.f == "'Sheet'!$A$1:$E$1"
        assert chart.ser[-1].val.numRef.f == "'Sheet'!$A$4:$E$4"

    def test_hash_function(self, chart_base):
        chart = chart_base()
        assert hash(chart) == hash(id(chart))

    def test_path(self, chart_base):
        chart = chart_base()
        assert chart.path == "/xl/charts/chart1.xml"

    def test_plot_area(self, chart_base):
        chart = chart_base()
        assert type(chart.plot_area) is PlotArea

    def test_save_twice(self, chart_base):
        chart_base.tagname = "DummyChart"
        chart = chart_base()
        chart._write()
        chart._write()
        area = chart.plot_area
        assert len(area._charts) == 1
        assert area._axes == []

    def test_axIds(self, chart_base):
        chart = chart_base()
        assert chart.axId == []

    def test_plot_visible_cells(self, chart_base):
        chart = chart_base()
        assert chart.visible_cells_only is True

    def test_plot_visible_cells2(self, chart_base):
        chart = chart_base()
        chart.visible_cells_only = False
        tree = chart._write()
        expected = """
        <chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <chart>
                <plotArea>
                    <DummyChart visible_cells_only="0" display_blanks="gap"/>
                </plotArea>
                <legend>
                    <legendPos val="r"></legendPos>
                    <overlay val="0"/>
                </legend>
                <plotVisOnly val="0"></plotVisOnly>
                <dispBlanksAs val="gap"></dispBlanksAs>
            </chart>
        </chartSpace>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_pivot_source(self, chart_base):
        chart = chart_base()
        chart.pivotSource = PivotSource(name="some pivot", fmtId=5)
        expected = """
        <chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <pivotSource>
                <name>some pivot</name>
                <fmtId val="5"/>
            </pivotSource>
            <chart>
                <plotArea>
                    <DummyChart visible_cells_only="1" display_blanks="gap"/>
                </plotArea>
                <legend>
                    <legendPos val="r"></legendPos>
                    <overlay val="0"/>
                </legend>
                <plotVisOnly val="1"></plotVisOnly>
                <dispBlanksAs val="gap"></dispBlanksAs>
            </chart>
        </chartSpace>
        """
        tree = chart._write()
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_pivot_format(self, chart_base):
        chart = chart_base()
        fmt = PivotFormat()
        chart.pivotFormats = [fmt]
        expected = """
        <chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <chart>
                <pivotFmts>
                    <pivotFmt>
                        <idx val="0"/>
                    </pivotFmt>
                </pivotFmts>
                <plotArea>
                    <DummyChart visible_cells_only="1" display_blanks="gap"/>
                </plotArea>
                <legend>
                    <legendPos val="r"></legendPos>
                    <overlay val="0"/>
                </legend>
                <plotVisOnly val="1"></plotVisOnly>
                <dispBlanksAs val="gap"></dispBlanksAs>
            </chart>
        </chartSpace>
        """
        tree = chart._write()
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_graphical_properties(self, chart_base):
        from openpyxl.chart.shapes import GraphicalProperties

        chart = chart_base()
        chart.graphical_properties = GraphicalProperties()
        tree = chart._write()
        expected = """
        <chartSpace
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <chart>
                <plotArea>
                    <DummyChart visible_cells_only="1" display_blanks="gap"/>
                </plotArea>
                <legend>
                    <legendPos val="r"></legendPos>
                    <overlay val="0"/>
                </legend>
                <plotVisOnly val="1"></plotVisOnly>
                <dispBlanksAs val="gap"></dispBlanksAs>
            </chart>
            <spPr>
                <a:ln>
                    <a:prstDash val="solid"/>
                </a:ln>
            </spPr>
        </chartSpace>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_no_fill(self, chart_base):
        from openpyxl.chart.shapes import GraphicalProperties

        chart = chart_base()
        chart.graphical_properties = GraphicalProperties()
        chart.graphical_properties.noFill = True
        chart.graphical_properties.line = None
        tree = chart._write()
        expected = """
        <chartSpace
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <chart>
                <plotArea>
                    <DummyChart visible_cells_only="1" display_blanks="gap"/>
                </plotArea>
                <legend>
                    <legendPos val="r"></legendPos>
                    <overlay val="0"/>
                </legend>
                <plotVisOnly val="1"></plotVisOnly>
                <dispBlanksAs val="gap"></dispBlanksAs>
            </chart>
            <spPr>
                <a:noFill/>
            </spPr>
        </chartSpace>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_no_border(self, chart_base):
        from openpyxl.chart.shapes import GraphicalProperties

        chart = chart_base()
        chart.graphical_properties = GraphicalProperties()
        chart.graphical_properties.line.noFill = True
        chart.graphical_properties.line.prstDash = None
        tree = chart._write()
        expected = """
        <chartSpace
                xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <chart>
                <plotArea>
                    <DummyChart visible_cells_only="1" display_blanks="gap"/>
                </plotArea>
                <legend>
                    <legendPos val="r"></legendPos>
                    <overlay val="0"/>
                </legend>
                <plotVisOnly val="1"></plotVisOnly>
                <dispBlanksAs val="gap"></dispBlanksAs>
            </chart>
            <spPr>
                <a:ln>
                    <a:noFill/>
                </a:ln>
            </spPr>
        </chartSpace>
        """
        xml = tostring(tree)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_reindex(self, chart_base):
        chart = chart_base()
        chart.ser = []
        chart.add_data("Sheet!D1:D4")
        chart.add_data("Sheet!B1:B4")
        chart.add_data("Sheet!C1:C4")
        chart.add_data("Sheet!A1:A4")
        orders = [40, 20, 34, 11]
        for o, s in zip(orders, chart.series):
            s.order = o
        ordered = [s.order for s in chart.series]
        assert ordered == [40, 20, 34, 11]
        chart._reindex()
        reordered = [s.order for s in chart.series]
        assert reordered == [0, 1, 2, 3]
