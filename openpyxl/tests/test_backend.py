#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file
"""Make sure we're using the fastest backend available"""
from xml.etree.ElementTree import Element as pyElement

from openpyxl.xml import LXML

try:
    from xml.etree.cElementTree import Element as cElement

    C = True
except ImportError:
    C = False

try:
    from lxml.etree import Element as lElement
except ImportError:
    lElement = None


def test_backend():
    from openpyxl.xml.functions import Element

    if LXML:
        assert Element == lElement
    elif C:
        assert Element == cElement
    else:
        assert Element == pyElement


def test_namespace_register():
    from openpyxl.xml.constants import SHEET_MAIN_NS
    from openpyxl.xml.functions import Element
    from openpyxl.xml.functions import tostring

    e = Element(f"{{{SHEET_MAIN_NS}}}sheet")
    xml = tostring(e)
    if hasattr(xml, "decode"):
        xml = xml.decode("utf-8")
    assert xml.startswith("<s:sheet")
