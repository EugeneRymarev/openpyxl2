# Copyright (c) 2010-2024 openpyxl
from ..axis import DateAxis
from ..axis import NumericAxis
from ..bar_chart import BarChart
from ..chartspace import ChartContainer
from ..chartspace import ChartSpace
from ..line_chart import LineChart
from openpyxl.xml.functions import fromstring


def test_read(datadir):
    datadir.chdir()
    from ..reader import read_chart

    with open("chart1.xml") as src:
        xml = src.read()
    tree = fromstring(xml)
    cs = ChartSpace.from_tree(tree)
    chart = read_chart(cs)

    assert isinstance(chart, LineChart)
    assert chart.title.tx.rich.p[0].r[0].t == "Website Performance"
    assert chart.display_blanks == "span"

    assert isinstance(chart.y_axis, NumericAxis)
    assert chart.y_axis.title.tx.rich.p[0].r[0].t == "Time in seconds"

    assert isinstance(chart.x_axis, DateAxis)
    assert chart.x_axis.title is None

    assert len(chart.series) == 10

    assert chart.pivotSource.name == "[files.xlsx]PIVOT!PivotTable1"
    assert len(chart.pivotFormats) == 1

    assert chart.idx_base == 0


def test_read_chart_with_no_series():
    container = ChartContainer()
    cs = ChartSpace(chart=container)
    cs.chart.plotArea.barChart = BarChart()

    from ..reader import read_chart

    chart = read_chart(cs)

    assert isinstance(chart, BarChart)
    assert len(chart.series) == 0
    assert chart.idx_base == 0


def test_read_chart_with_shapes(datadir):
    datadir.chdir()

    from ..reader import read_chart

    with open("chart_no_border.xml") as src:
        xml = src.read()

    tree = fromstring(xml)
    cs = ChartSpace.from_tree(tree)
    chart = read_chart(cs)

    assert chart.graphical_properties is not None
    assert chart.graphical_properties.noFill
    assert chart.graphical_properties.line.noFill
