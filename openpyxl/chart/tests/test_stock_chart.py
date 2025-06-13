# Copyright (c) 2010-2025 openpyxl
import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring
from openpyxl.xml.functions import tostring


@pytest.fixture
def stock_chart():
    from openpyxl.chart.stock_chart import StockChart

    return StockChart


class TestStockChart:
    def test_ctor(self, stock_chart):
        from openpyxl.chart.series import Series

        chart = stock_chart(ser=[Series(), Series(), Series()])
        xml = tostring(chart.to_tree())
        expected = """
        <stockChart
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <ser>
                <idx val="0"/>
                <order val="0"/>
                <spPr>
                    <a:ln>
                        <a:prstDash val="solid"/>
                    </a:ln>
                </spPr>
                <marker>
                    <symbol val="none"/>
                    <spPr>
                        <a:ln>
                            <a:prstDash val="solid"/>
                        </a:ln>
                    </spPr>
                </marker>
            </ser>
            <ser>
                <idx val="1"/>
                <order val="1"/>
                <spPr>
                    <a:ln>
                        <a:prstDash val="solid"/>
                    </a:ln>
                </spPr>
                <marker>
                    <symbol val="none"/>
                    <spPr>
                        <a:ln>
                            <a:prstDash val="solid"/>
                        </a:ln>
                    </spPr>
                </marker>
            </ser>
            <ser>
                <idx val="2"></idx>
                <order val="2"></order>
                <spPr>
                    <a:ln>
                        <a:prstDash val="solid"/>
                    </a:ln>
                </spPr>
                <marker>
                    <symbol val="none"/>
                    <spPr>
                        <a:ln>
                            <a:prstDash val="solid"/>
                        </a:ln>
                    </spPr>
                </marker>
            </ser>
            <axId val="10"></axId>
            <axId val="100"></axId>
        </stockChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_from_xml(self, stock_chart):
        src = """
        <stockChart>
            <axId val="10"></axId>
            <axId val="100"></axId>
        </stockChart>
        """
        node = fromstring(src)
        chart = stock_chart.from_tree(node)
        assert chart.axId == [10, 100]
