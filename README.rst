.. image:: https://coveralls.io/repos/github/EugeneRymarev/openpyxl2/badge.svg?branch=master
    :target: https://coveralls.io/github/EugeneRymarev/openpyxl2?branch=master
    :alt: coverage status


About This Fork
---------------

This is a fork of the `original library <https://foss.heptapod.net/openpyxl/openpyxl>`_.

I made it because I disagree with the author's opinion on some issues regarding the result of some functions. The list of differences will be provided below.


Differences
-----------

* All code (except documentation) is processed by Black and reorder-python-imports modules
* MultiCellRange can now return a CellRange for any of the merged cells.
* Now you can compare NamedStyle styles and style names
* Now Alignment, Border, Font, Fill, Protection, NumberFormat and Style are applied to all merged cells. Explanation below.

Explanation
-----------
Styles assigned to the anchor of a merged range are applied to its cells.
Borders are projected onto the outer perimeter, while the anchor retains the
complete border definition. Replacing or deleting the anchor border also updates
the perimeter. Named styles follow the same rule.

Loading a workbook preserves the stored styles of non-anchor merged cells,
including individual formatting, through subsequent save/load cycles.
Explicit border assignments to a non-anchor MergedCell remain local.


In progress
------------
* Fix the problem with the insert_rows, insert_cols, delete_rows, delete_cols functions so that the result of executing these functions matches the result of executing similar actions in Excel.


Introduction
------------

openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.

It was born from lack of existing library to read/write natively from Python
the Office Open XML format.

All kudos to the PHPExcel team as openpyxl was initially based on PHPExcel.


Security
--------

By default openpyxl does not guard against quadratic blowup or billion laughs
xml attacks. To guard against these attacks install defusedxml.

Mailing List
------------

The user list can be found on http://groups.google.com/group/openpyxl-users


Sample code::

    from openpyxl import Workbook
    wb = Workbook()

    # grab the active worksheet
    ws = wb.active

    # Data can be assigned directly to cells
    ws['A1'] = 42

    # Rows can also be appended
    ws.append([1, 2, 3])

    # Python types will automatically be converted
    import datetime
    ws['A2'] = datetime.datetime.now()

    # Save the file
    wb.save("sample.xlsx")


Documentation
-------------

The documentation is at: https://openpyxl.pages.heptapod.net/openpyxl/index.html

* installation methods
* code examples
* instructions for contributing

Release notes: https://openpyxl.pages.heptapod.net/openpyxl/changes.html
