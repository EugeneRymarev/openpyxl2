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
At the moment, the function of applying styles to all merged cells works only when saving and when reading in read-only mode.

If you read a file with merged cells, then MergedCell will not have any of the properties - they are replaced by the _clean_merge_range function.

As a result, after reading and saving, the styles of merged cells are lost.

This behavior will be fixed in the future.

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

The documentation is at: https://openpyxl.readthedocs.io

* installation methods
* code examples
* instructions for contributing

Release notes: https://openpyxl.readthedocs.io/en/stable/changes.html
