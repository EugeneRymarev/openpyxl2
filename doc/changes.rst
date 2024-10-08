3.1.5 (2024-06-28)
==================


* `#2187 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2187>`_ Test fails due to change in Numpy API
* `#2198 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2198>`_ Excel is very fussy about the version number
* `#2200 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2200>`_ Poor perfomance when reading workbooks with lots of named styles


3.1.4 (2024-06-12)
==================

* `#2189 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2189>`_ Assigning named styles doesn't work
* `#2190 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2190>`_ Problems caused when saving workbooks created by LibreOffice


Changes
-------

* Python 3.6 and 3.7 are no longer supported.


3.1.3 (2024-05-29)
==================


Bugfixes
--------

* `#1401 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1401>`_ Column name caches are slow and use a lot of memory
* `#1457 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1457>`_ Improved handling of duplicate named styles
* `#1842 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1842>`_ Rich-text can be saved if lxml is not installed
* `#1954 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1954>`_ Documentation for sheet views is incorrect
* `#1973 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1973>`_ Timedeltas not read properly in read-only mode
* `#1987 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1987>`_ List of formulae names contains mistakes
* `#1967 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1967>`_ Filters does not handle non-numerical filters
* `#2054 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2054>`_ Type checking increases exponentially
* `#2057 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2057>`_ Loading pivot tables can be unnecessarily slow
* `#2102 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2102>`_ Improve performance when reading files with lots of custom properties
* `#2106 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2106>`_ Setting Trendline.name attribute raises exception when saving
* `#2120 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2120>`_ Timezone and Zombie formatting cannot be combined.
* `#2107 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2107>`_ Column name generation is inefficient and slow
* `#2122 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2122>`_ File handlers not always released in read-only mode
* `#2149 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2149>`_ Workbook files not properly closed on Python ≥ 3.11.8 and Windows
* `#2161 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2161>`_ Pivot cache definitions using tupleCache had serialisation issues

Changes
-------

* Add a `__repr__` method for Row and Column dimension objects so you don't need to check every time.


3.1.2 (2023-03-11)
==================

* `#1963 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1963>`_ Cannot read worksheets in read-only mode with locally scoped definitions
* `#1974 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1974>`_ Empty custom properties cause invalid files




3.1.1 (2023-02-13)
==================


Bugfixes
--------

* `#1881 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1881>`_ DocumentProperties times set by module import only
* `#1947 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1947>`_ Worksheet-specific definitions are missing


3.1.0 (2023-01-31)
==================


New Features
------------

* Added support for data table formulae
* Mapped chartspace graphical properties to charts for advanced formatting


Bugfixes
--------

* `#1156 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1156>`_ Table filters are always overriden
* `#1360 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1360>`_ Can't read some ScatterCharts if n
* `#1724 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1724>`_ Problem with multilevel indices in dataframes
* `#1763 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1763>`_ Make calculating worksheet sizes slightly faster
* `#1772 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1772>`_ Problem with category indices in dataframes
* `#1786 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1786>`_ NamedStyles share attributes - mutables gotcha
* `#1851 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1851>`_ Allow print area to be set to None
* `#1852 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1852>`_ Worksheet for print title and print areas can't be found
* `#1853 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1853>`_ Custom document properties that are strings can be empty
* `#1858 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1858>`_ ConditionalFormatting lost when pivot table updated
* `#1864 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1864>`_ Better handling of defined names
* `#1904 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1904>`_ dataframe_to_rows() misalignment on multiindex
* `#1908 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1908>`_ Ditto
* `#1912 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1912>`_ Excel doesn't like xmlns:space on nodes with only whitespace, which it treats as empty.
* `#1942 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1942>`_ Exception when print areas use table references.


Pull Requests
-------------

* `PR409 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/409/>`_ Support for Rich Text in cells
* `PR411 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/411/>`_ Provide more information when workbook cannot be loaded
* `PR407 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/407/>`_ Support for Custom Document Properties


Deprecations
------------

The following properties have been removed from worksheets: formula_attributes, page_breaks, show_summary_below, show_summary_right, page_size orientation. Client code should use the relevant objects.


Removals
--------

The following deprecated methods have been removed from workbooks: get_named_range, add_named_range, remove_named_range. And the get_emu_dimesions from images.


3.0.10 (2022-05-19)
===================


Bugfixes
--------

* `#1684 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1684>`_ Image files not closed when workbooks are saved
* `#1778 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1778>`_ Problem with missing scope attribute in Pivot Table formats
* `#1821 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1821>`_ Excel unhappy when multiple sorts are defined
* `#2014 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/2014>`_ Accounting format interpreted as datetime


3.0.9 (2021-09-22)
==================


Bugfixes
--------

* `#1284 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1284>`_ Ignore blank ignored in existing Data Validations
* `#1539 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1539>`_ Add support for cell protection for merged cell ranges
* `#1645 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1645>`_ Timezone-aware datetimes raise an Exception
* `#1666 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1666>`_ Improved normalisation of chart series
* `#1670 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1670>`_ Catch OverflowError for out of range datetimes
* `#1708 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1708>`_ Alignment.relativeIndent can be negative
* `#1736 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1769>`_ Incorrect default value `groupBy` attribute


3.0.8 (brown bag)
==================

Deleted because it contained breaking changes from 3.1


3.0.7 (2021-03-09)
==================


Bugfixes
--------

* `#1510 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1510>`_ Problems with zero time values
* `#1588 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1588>`_ Not possible to correctly convert excel dates to timedelta
* `#1589 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1589>`_ Exception raised when merging cells which do not have borders all the way round.
* `#1594 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1594>`_ Python 2 print statement in the tutorial


Pull Requests
-------------

* `PR392 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/392/>`_ Add documentation on datetime handling
* `PR393 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/393/>`_ Drop dependency on jdcal
* `PR394 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/394/>`_ Datetime rounding
* `PR395 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/395/>`_ Unify handling of 1900 epoch
* `PR397 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/397/>`_ Add explicit support for reading datetime deltas
* `PR399 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/399/>`_ Millisecond precision for datetimes


3.0.6 (2021-01-14)
==================


Bugfixes
--------

* `#1154 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1154>`_ Borders in differential styles are incorrect
* `#1287 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1528>`_ Error when opening some pivot tables
* `#1366 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1366>`_ Resave breaks the border format in conditional formatting rules
* `#1450 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1450>`_ Read-only workbook not closed properly if generator interrupted
* `#1547 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1547>`_ Pandas.Multiindex.labels deprecated
* `#1552 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1557>`_ Pandas.Multiinex not expanded correctly
* `#1557 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1557>`_ Cannot read rows with exponents
* `#1568 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1568>`_ numpy.float is deprecated
* `#1571 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1571>`_ Cells without coordinate attributes not always correctly handled


Pull Requests
-------------

* `PR385 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/385/>`_ Improved handling of borders for differential styles
* `PR386 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/386/>`_ Support subclasses of datetime objects
* `PR387 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/387/>`_ Improved handling of cells without coordinates


3.0.5 (2020-08-21)
==================


Bugfixes
--------

* `#1413 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1413>`_ Incorrectly consider currency format as datetime
* `#1490 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1490>`_ Cannot copy worksheets with merged cells
* `#1492 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1492>`_ Empty worksheets do not return generators when looping.
* `#1496 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1496>`_ Hyperlinks duplicated on multiple saves
* `#1500 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1500>`_ Incorrectly literal format as datetime
* `#1502 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1502>`_ Links set to range of cells not preserved
* `#1507 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1507>`_ Exception when opening workbook with chartsheets and tables


3.0.4 (2020-06-24)
==================


Bugfixes
--------

* `#844 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/844>`_ Find tables by name
* `#1414 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1414>`_ Worksheet protection missing in existing files
* `#1439 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1439>`_ Exception when reading files with external images
* `#1452 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1452>`_ Reading lots of merged cells is very slow.
* `#1455 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1455>`_ Read support for Bubble Charts.
* `#1458 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1458>`_ Preserve any indexed colours
* `#1473 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1473>`_ Reading many thousand of merged cells is really slow.
* `#1474 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1474>`_ Adding tables in write-only mode raises an exception.


Pull Requests
-------------

* `PR377 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/377/>`_ Add support for finding tables by name or range.


3.0.3 (2020-01-20)
==================


Bugfixes
--------

* `#1260 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1260>`_ Exception when handling merged cells with hyperlinks
* `#1373 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1373>`_ Problems when both lxml and defusedxml are installed
* `#1385 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1385>`_ CFVO with incorrect values cannot be processed


3.0.2 (2019-11-25)
==================


Bug fixes
---------

* `#1267 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1267>`_ DeprecationError if both defusedxml and lxml are installed
* `#1345 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1345>`_ ws._current_row is higher than ws.max_row
* `#1365 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1365>`_ Border bottom style is not optional when it should be
* `#1367 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1367>`_ Empty cells in read-only, values-only mode are sometimes returned as ReadOnlyCells
* `#1368 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1368>`_ Cannot add page breaks to existing worksheets if none exist already


Pull Requests
-------------

* `PR359 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/359/>`_ Improvements to the documentation


3.0.1 (2019-11-14)
==================

Bugfixes
--------

* `#1250 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1250>`_ Cannot read empty charts.


Pull Requests
-------------

* `PR354 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/354/>`_ Fix for #1250
* `PR352 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/354/>`_ TableStyleElement is a sequence


3.0.0 (2019-09-25)
==================

Python 3.6+ only release
------------------------


2.6.4 (2019-09-25)
==================


Final release for Python 2.7 and 3.5
------------------------------------

Bugfixes
--------

* ` #1330 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1330>`_ Cannot save workbooks with comments more than once.


2.6.3 (2019-08-19)
==================


Bugfixes
--------

* `#1237 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1237>`_ Fix 3D charts.
* `#1290 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1290>`_ Minimum for holeSize in Doughnut charts too high
* `#1291 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1291>`_ Warning for MergedCells with comments
* `#1296 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1296>`_ Pagebreaks duplicated
* `#1309 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1309>`_ Workbook has no default CellStyle
* `#1330 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1330>`_ Workbooks with comments cannot be saved multiple times


Pull Requests
-------------

* `PR344 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/345/>`_ Make sure NamedStyles number formats are correctly handled


2.6.2 (2019-03-29)
==================


Bugfixes
--------

* `#1173 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1173>`_ Workbook has no _date_formats attribute
* `#1190 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1190>`_ Cannot create charts for worksheets with quotes in the title
* `#1228 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1228>`_ MergedCells not removed when range is unmerged
* `#1232 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1232>`_ Link to pivot table lost from charts
* `#1233 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1233>`_ Chart colours change after saving
* `#1236 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1236>`_ Cannot use ws.cell in read-only mode with Python 2.7



2.6.1 (2019-03-04)
==================


Bugfixes
--------

* `#1174 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1174>`_ ReadOnlyCell.is_date does not work properly
* `#1175 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1175>`_ Cannot read Google Docs spreadsheet with a Pivot Table
* `#1180 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1180>`_ Charts created with openpyxl cannot be styled
* `#1181 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1181>`_ Cannot handle some numpy number types
* `#1182 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1182>`_ Exception when reading unknowable number formats
* `#1186 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1186>`_ Only last formatting rule for a range loaded
* `#1191 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1191>`_ Give MergedCell a `value` attribute
* `#1193 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1193>`_ Cannot process worksheets with comments
* `#1197 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1197>`_ Cannot process worksheets with both row and page breaks
* `#1204 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1204>`_ Cannot reset dimensions in ReadOnlyWorksheets
* `#1211 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1211>`_ Incorrect descriptor in ParagraphProperties
* `#1213 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1213>`_ Missing `hier` attribute in PageField raises an exception


2.6.0 (2019-02-06)
==================


Bugfixes
--------

* `#1162 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1162>`_ Exception on tables with names containing spaces.
* `#1170 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1170>`_ Cannot save files with existing images.


2.6.-b1 (2019-01-08)
====================


Bugfixes
--------

* `#1141 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1141>`_ Cannot use read-only mode with stream
* `#1143 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1143>`_ Hyperlinks always set on A1
* `#1151 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1151>`_ Internal row counter not initialised when reading files
* `#1152 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1152>`_ Exception raised on out of bounds date


2.6-a1 (2018-11-21)
===================


Major changes
-------------

* Implement robust for merged cells so that these can be formatted the way
  Excel does without confusion. Thanks to Magnus Schieder.


Minor changes
-------------

* Add support for worksheet scenarios
* Add read support for chartsheets
* Add method for moving ranges of cells on a worksheet
* Drop support for Python 3.4
* Last version to support Python 2.7


Deprecations
------------

* Type inference and coercion for cell values


2.5.14 (2019-01-23)
===================


Bugfixes
--------

* `#1150 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1150>`_ Correct typo in LineProperties
* `#1142 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1142>`_ Exception raised for unsupported image files
* `#1159 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1159>`_ Exception raised when cannot find source for non-local cache object


Pull Requests
-------------

* `PR301 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/301/>`_ Add support for nested brackets to the tokeniser
* `PR303 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/301/>`_ Improvements on handling nested brackets in the tokeniser


2.5.13 (brown bag)
==================


2.5.12 (2018-11-29)
===================


Bugfixes
--------

* `#1130 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1130>`_ Overwriting default font in Normal style affects library default
* `#1133 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1133>`_ Images not added to anchors.
* `#1134 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1134>`_ Cannot read pivot table formats without dxId
* `#1138 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1138>`_ Repeated registration of simple filter could lead to memory leaks


Pull Requests
-------------

* `PR300 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/300/>`_ Use defusedxml if available


2.5.11 (2018-11-21)
===================


Pull Requests
-------------

* `PR295 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/295>`_ Improved handling of missing rows
* `PR296 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/296>`_ Add support for defined names to tokeniser


2.5.10 (2018-11-13)
===================


Bugfixes
--------

* `#1114 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1114>`_ Empty column dimensions should not be saved.


Pull Requests
-------------

* `PR285 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/285>`_ Tokenizer failure for quoted sheet name in second half of range
* `PR289 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/289>`_ Improved error detection in ranges.


2.5.9 (2018-10-19)
==================


Bugfixes
--------

* `#1000 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1000>`_ Clean AutoFilter name definitions
* `#1106 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1106>`_ Attribute missing from Shape object
* `#1109 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1109>`_ Failure to read all DrawingML means workbook can't be read


Pull Requests
-------------

* `PR281 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/281>`_ Allow newlines in formulae
* `PR284 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/284>`_ Fix whitespace in front of infix operator in formulae


2.5.8 (2018-09-25)
==================


* `#877 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/877>`_ Cannot control how missing values are displayed in charts.
* `#948 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/948>`_ Cell references can't be used for chart titles
* `#1095 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1095>`_ Params in iter_cols and iter_rows methods are slightly wrong.


2.5.7 (2018-09-13)
==================


* `#954 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/954>`_ Sheet title containing % need quoting in references
* `#1047 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1047>`_ Cannot set quote prefix
* `#1093 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1093>`_ Pandas timestamps raise KeyError


2.5.6 (2018-08-30)
==================


* `#832 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/832>`_ Read-only mode can leave find-handles open when reading dimensions
* `#933 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/933>`_ Set a worksheet directly as active
* `#1086 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1086>`_ Internal row counter not adjusted when rows are deleted or inserted


2.5.5 (2018-08-04)
==================


Bugfixes
--------

* `#1049 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1049>`_ Files with Mac epoch are read incorrectly
* `#1058 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1058>`_ Cannot copy merged cells
* `#1066 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1066>`_ Cannot access ws.active_cell


Pull Requests
-------------

* `PR267 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/267/image-read>`_ Introduce read-support for images


2.5.4 (2018-06-07)
==================


Bugfixes
--------

* `#1025 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1025>`_ Cannot read files with 3D charts.
* `#1030 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1030>`_ Merged cells take a long time to parse


Minor changes
-------------

* Improve read support for pivot tables and don't always create a Filters child for filterColumn objects.
* `Support folding rows` <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/259/fold-rows>`_


2.5.3 (2018-04-18)
==================


Bugfixes
--------

* `#983 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/983>`_ Warning level too aggressive.
* `#1015 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1015>`_ Alignment and protection values not saved for named styles.
* `#1017 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1017>`_ Deleting elements from a legend doesn't work.
* `#1018 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1018>`_ Index names repeated for every row in dataframe.
* `#1020 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1020>`_ Worksheet protection not being stored.
* `#1023 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1023>`_ Exception raised when reading a tooltip.


2.5.2 (2018-04-06)
==================


Bugfixes
--------

* `#949 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/949>`_ High memory use when reading text-heavy files.
* `#970 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/970>`_ Copying merged cells copies references.
* `#978 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/978>`_ Cannot set comment size.
* `#985 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/895>`_ Exception when trying to save workbooks with no views.
* `#995 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/995>`_ Cannot delete last row or column.
* `#1002 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1002>`_ Cannot read Drawings containing embedded images.


Minor changes
-------------

* Support for dataframes with multiple columns and multiple indices.


2.5.1 (2018-03-12)
==================


Bugfixes
--------

* `#934 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/934>`_ Headers and footers not included in write-only mode.
* `#960 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/960>`_ Deprecation warning raised when using ad-hoc access in read-only mode.
* `#964 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/964>`_ Not all cells removed when deleting multiple rows.
* `#966 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/966>`_ Cannot read 3d bar chart correctly.
* `#967 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/967>`_ Problems reading some charts.
* `#968 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/968>`_ Worksheets with SHA protection become corrupted after saving.
* `#974 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/974>`_ Problem when deleting ragged rows or columns.
* `#976 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/976>`_ GroupTransforms and GroupShapeProperties have incorrect descriptors
* Make sure that headers and footers in chartsheets are included in the file



2.5.0 (2018-01-24)
==================


Minor changes
-------------

* Correct definition for Connection Shapes. Related to # 958


2.5.0-b2 (2018-01-19)
=====================


Bugfixes
--------

* `#915 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/915>`_ TableStyleInfo has no required attributes
* `#925 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/925>`_ Cannot read files with 3D drawings
* `#926 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/926>`_ Incorrect version check in installer
* Cell merging uses transposed parameters
* `#928 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/928>`_ ExtLst missing keyword for PivotFields
* `#932 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/932>`_ Inf causes problems for Excel
* `#952 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/952>`_ Cannot load table styles with custom names


Major Changes
-------------

* You can now insert and delete rows and columns in worksheets


Minor Changes
-------------

* pip now handles which Python versions can be used.


2.5.0-b1 (2017-10-19)
=====================


Bugfixes
--------
* `#812 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/812>`_ Explicitly support for multiple cell ranges in conditonal formatting
* `#827 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/827>`_ Non-contiguous cell ranges in validators get merged
* `#837 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/837>`_ Empty data validators create invalid Excel files
* `#860 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/860>`_ Large validation ranges use lots of memory
* `#876 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/876>`_ Unicode in chart axes not handled correctly in Python 2
* `#882 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/882>`_ ScatterCharts have defective axes
* `#885 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/885>`_ Charts with empty numVal elements cannot be read
* `#894 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/894>`_ Scaling options from existing files ignored
* `#895 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/895>`_ Charts with PivotSource cannot be read
* `#903 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/903>`_ Cannot read gradient fills
* `#904 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/904>`_ Quotes in number formats could be treated as datetimes


Major Changes
-------------

`worksheet.cell()` no longer accepts a `coordinate` parameter. The syntax is now `ws.cell(row, column, value=None)`


Minor Changes
-------------

Added CellRange and MultiCellRange types (thanks to Laurent LaPorte for the
suggestion) as a utility type for things like data validations, conditional
formatting and merged cells.


Deprecations
------------

ws.merged_cell_ranges has been deprecated because MultiCellRange provides sufficient functionality


2.5.0-a3 (2017-08-14)
=====================


Bugfixes
--------
* `#848 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/848>`_ Reading workbooks with Pie Charts raises an exception
* `#857 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/857>`_ Pivot Tables without Worksheet Sources raise an exception


2.5.0-a2 (2017-06-25)
=====================


Major Changes
-------------

* Read support for charts


Bugfixes
--------
* `#833 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/833>`_ Cannot access chartsheets by title
* `#834 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/834>`_ Preserve workbook views
* `#841 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/841>`_ Incorrect classification of a datetime


2.5.0-a1 (2017-05-30)
=====================


Compatibility
-------------

* Dropped support for Python 2.6 and 3.3. openpyxl will not run with Python 2.6


Major Changes
-------------

* Read/write support for pivot tables


Deprecations
------------

* Dropped the anchor method from images and additional constructor arguments


Bugfixes
--------
* `#779 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/779>`_ Fails to recognise Chinese date format`
* `#828 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/828>`_ Include hidden cells in charts`


Pull requests
-------------
* `163 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/163>`_ Improved GradientFill


Minor changes
-------------

* Remove deprecated methods from Cell
* Remove deprecated methods from Worksheet
* Added read/write support for the datetime type for cells


2.4.11 (2018-01-24)
===================

* #957 `<https://foss.heptapod.net/openpyxl/openpyxl/-/issues/957>`_ Relationship type for tables is borked


2.4.10 (2018-01-19)
===================

Bugfixes
--------

* #912 `<https://foss.heptapod.net/openpyxl/openpyxl/-/issues/912>`_ Copying objects uses shallow copy
* #921 `<https://foss.heptapod.net/openpyxl/openpyxl/-/issues/921>`_ API documentation not generated automatically
* #927 `<https://foss.heptapod.net/openpyxl/openpyxl/-/issues/927>`_ Exception raised when adding coloured borders together
* #931 `<https://foss.heptapod.net/openpyxl/openpyxl/-/issues/931>`_ Number formats not correctly deduplicated


Pull requests
-------------

* 203 `<https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/203/>`_ Correction to worksheet protection description
* 210 `<https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/210/>`_ Some improvements to the API docs
* 211 `<https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/211/>`_ Improved deprecation decorator
* 218 `<https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/218/>`_ Fix problems with deepcopy


2.4.9 (2017-10-19)
==================


Bugfixes
--------

* `#809 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/809>`_ Incomplete documentation of `copy_worksheet` method
* `#811 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/811>`_ Scoped definedNames not removed when worksheet is deleted
* `#824 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/824>`_ Raise an exception if a chart is used in multiple sheets
* `#842 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/842>`_ Non-ASCII table column headings cause an exception in Python 2
* `#846 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/846>`_ Conditional formats not supported in write-only mode
* `#849 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/849>`_ Conditional formats with no sqref cause an exception
* `#859 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/859>`_ Headers that start with a number conflict with font size
* `#902 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/902>`_ TableStyleElements don't always have a condtional format
* `#908 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/908>`_ Read-only mode sometimes returns too many cells



Pull requests
-------------

* `#179 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/179>`_ Cells kept in a set
* `#180 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/180>`_ Support for Workbook protection
* `#182 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/182>`_ Read support for page breaks
* `#183 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/183>`_ Improve documentation of `copy_worksheet` method
* `#198 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/198>`_ Fix for #908


2.4.8 (2017-05-30)
==================


Bugfixes
--------

* AutoFilter.sortState being assignd to the ws.sortState
* `#766 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/666>`_ Sheetnames with apostrophes need additional escaping
* `#729 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/729>`_ Cannot open files created by Microsoft Dynamics
* `#819 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/819>`_ Negative percents not case correctly
* `#821 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/821>`_ Runtime imports can cause deadlock
* `#855 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/855>`_ Print area containing only columns leads to corrupt file


Minor changes
-------------
* Preserve any table styles


2.4.7 (2017-04-24)
==================


Bugfixes
--------
* `#807 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/807>`_ Sample files being included by mistake in sdist


2.4.6 (2017-04-14)
==================


Bugfixes
--------
* `#776 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/776>`_ Cannot apply formatting to plot area
* `#780 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/780>`_ Exception when element attributes are Python keywords
* `#781 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/781>`_ Exception raised when saving files with styled columns
* `#785 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/785>`_ Number formats for data labels are incorrect
* `#788 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/788>`_ Worksheet titles not quoted in defined names
* `#800 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/800>`_ Font underlines not read correctly


2.4.5 (2017-03-07)
==================


Bugfixes
--------
* `#750 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/750>`_ Adding images keeps file handles open
* `#772 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/772>`_ Exception for column-only ranges
* `#773 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/773>`_ Cannot copy worksheets with non-ascii titles on Python 2


Pull requests
-------------

* `161 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/161>`_ Support for non-standard names for Workbook part.
* `162 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/162>`_ Documentation correction


2.4.4 (2017-02-23)
==================


Bugfixes
--------

* `#673 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/673>`_ Add close method to workbooks
* `#762 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/762>`_ openpyxl can create files with invalid style indices
* `#729 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/729>`_ Allow images in write-only mode
* `#744 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/744>`_ Rounded corners for charts
* `#747 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/747>`_ Use repr when handling non-convertible objects
* `#764 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/764>`_ Hashing function is incorrect
* `#765 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/765>`_ Named styles share underlying array


Minor Changes
-------------

* Add roundtrip support for worksheet tables.


Pull requests
-------------

* `160 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/160>`_ Don't init mimetypes more than once.


2.4.3 (unreleased)
==================
bad release


2.4.2 (2017-01-31)
==================


Bug fixes
---------

* `#727 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/727>`_ DeprecationWarning is incorrect
* `#734 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/734>`_ Exception raised if userName is missing
* `#739 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/739>`_ Always provide a date1904 attribute
* `#740 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/740>`_ Hashes should be stored as Base64
* `#743 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/743>`_ Print titles broken on sheetnames with spaces
* `#748 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/748>`_ Workbook breaks when active sheet is removed
* `#754 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/754>`_ Incorrect descriptor for Filter values
* `#756 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/756>`_ Potential XXE vulerability
* `#758 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/758>`_ Cannot create files with page breaks and charts
* `#759 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/759>`_ Problems with worksheets with commas in their titles


Minor Changes
-------------

* Add unicode support for sheet name incrementation.


2.4.1 (2016-11-23)
==================


Bug fixes
---------

* `#643 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/643>`_ Make checking for duplicate sheet titles case insensitive
* `#647 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/647>`_ Trouble handling LibreOffice files with named styles
* `#687 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/682>`_ Directly assigned new named styles always refer to "Normal"
* `#690 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/690>`_ Cannot parse print titles with multiple sheet names
* `#691 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/691>`_ Cannot work with macro files created by LibreOffice
* Prevent duplicate differential styles
* `#694 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/694>`_ Allow sheet titles longer than 31 characters
* `#697 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/697>`_ Cannot unset hyperlinks
* `#699 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/699>`_ Exception raised when format objects use cell references
* `#703 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/703>`_ Copy height and width when copying comments
* `#705 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/705>`_ Incorrect content type for VBA macros
* `#707 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/707>`_ IndexError raised in read-only mode when accessing individual cells
* `#711 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/711>`_ Files with external links become corrupted
* `#715 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/715>`_ Cannot read files containing macro sheets
* `#717 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/717>`_ Details from named styles not preserved when reading files
* `#722 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/722>`_ Remove broken Print Title and Print Area definitions


Minor changes
-------------

* Add support for Python 3.6
* Correct documentation for headers and footers


Deprecations
------------

Worksheet methods `get_named_range()` and `get_sqaured_range()`


Bug fixes
---------


2.4.0 (2016-09-15)
==================


Bug fixes
---------

* `#652 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/652>`_ Exception raised when epoch is 1904
* `#642 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/642>`_ Cannot handle unicode in headers and footers in Python 2
* `#646 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/646>`_ Cannot handle unicode sheetnames in Python 2
* `#658 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/658>`_ Chart styles, and axis units should not be 0
* `#663 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/663>`_ Strings in external workbooks not unicode


Major changes
-------------

* Add support for builtin styles and include one for Pandas


Minor changes
-------------

* Add a `keep_links` option to `load_workbook`. External links contain cached
  copies of the external workbooks. If these are big it can be advantageous to
  be able to disable them.
* Provide an example for using cell ranges in DataValidation.
* PR 138 - add copy support to comments.


2.4.0-b1 (2016-06-08)
=====================


Minor changes
-------------

* Add an the alias `hide_drop_down` to DataValidation for `showDropDown` because that is how Excel works.


Bug fixes
---------

* `#625 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/625>`_ Exception raises when inspecting EmptyCells in read-only mode
* `#547 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/547>`_ Functions for handling OOXML "escaped" ST_XStrings
* `#629 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/629>`_ Row Dimensions not supported in write-only mode
* `#530 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/530>`_ Problems when removing worksheets with charts
* `#630 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/630>`_ Cannot use SheetProtection in write-only mode


Features
--------

* Add write support for worksheet tables


2.4.0-a1 (2016-04-11)
=====================


Minor changes
-------------

* Remove deprecated methods from DataValidation
* Remove deprecated methods from PrintPageSetup
* Convert AutoFilter to Serialisable and extend support for filters
* Add support for SortState
* Removed `use_iterators` keyword when loading workbooks. Use `read_only` instead.
* Removed `optimized_write` keyword for new workbooks. Use `write_only` instead.
* Improve print title support
* Add print area support
* New implementation of defined names
* New implementation of page headers and footers
* Add support for Python's NaN
* Added iter_cols method for worksheets
* ws.rows and ws.columns now always return generators and start at the top of the worksheet
* Add a `values` property for worksheets
* Default column width changed to 8 as per the specification


Deprecations
------------

* Cell anchor method
* Worksheet point_pos method
* Worksheet add_print_title method
* Worksheet HeaderFooter attribute, replaced by individual ones
* Flatten function for cells
* Workbook get_named_range, add_named_range, remove_named_range, get_sheet_names, get_sheet_by_name
* Comment text attribute
* Use of range strings deprecated for ws.iter_rows()
* Use of coordinates deprecated for ws.cell()
* Deprecate .copy() method for StyleProxy objects


Bug fixes
---------

* `#152 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/152>`_ Hyperlinks lost when reading files
* `#171 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/171>`_ Add function for copying worksheets
* `#386 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/386>`_ Cells with inline strings considered empty
* `#397 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/397>`_ Add support for ranges of rows and columns
* `#446 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/446>`_ Workbook with definedNames corrupted by openpyxl
* `#481 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/481>`_ "safe" reserved ranges are not read from workbooks
* `#501 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/501>`_ Discarding named ranges can lead to corrupt files
* `#574 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/574>`_ Exception raised when using the class method to parse Relationships
* `#579 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/579>`_ Crashes when reading defined names with no content
* `#597 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/597>`_ Cannot read worksheets without coordinates
* `#617 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/617>`_ Customised named styles not correctly preserved


2.3.5 (2016-04-11)
==================


Bug fixes
---------

* `#618 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/618>`_ Comments not written in write-only mode


2.3.4 (2016-03-16)
==================


Bug fixes
---------

* `#594 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/594>`_ Content types might be missing when keeping VBA
* `#599 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/599>`_ Cells with only one cell look empty
* `#607 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/607>`_ Serialise NaN as ''


Minor changes
-------------

* Preserve the order of external references because formualae use numerical indices.
* Typo corrected in cell unit tests (PR 118)


2.3.3 (2016-01-18)
==================


Bug fixes
---------

* `#540 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/540>`_ Cannot read merged cells in read-only mode
* `#565 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/565>`_ Empty styled text blocks cannot be parsed
* `#569 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/569>`_ Issue warning rather than raise Exception raised for unparsable definedNames
* `#575 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/575>`_ Cannot open workbooks with embdedded OLE files
* `#584 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/584>`_ Exception when saving borders with attributes


Minor changes
-------------

* `PR 103 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/103/>`_ Documentation about chart scaling and axis limits
* Raise an exception when trying to copy cells from other workbooks.


2.3.2 (2015-12-07)
==================


Bug fixes
---------

* `#554 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/554>`_ Cannot add comments to a worksheet when preserving VBA
* `#561 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/561>`_ Exception when reading phonetic text
* `#562 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/562>`_ DARKBLUE is the same as RED
* `#563 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/563>`_ Minimum for row and column indexes not enforced


Minor changes
-------------

* `PR 97 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/97/>`_ One VML file per worksheet.
* `PR 96 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/96/>`_ Correct descriptor for CharacterProperties.rtl
* `#498 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/498>`_ Metadata is not essential to use the package.


2.3.1 (2015-11-20)
==================


Bug fixes
---------

* `#534 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/534>`_ Exception when using columns property in read-only mode.
* `#536 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/536>`_ Incorrectly handle comments from Google Docs files.
* `#539 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/539>`_ Flexible value types for conditional formatting.
* `#542 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/542>`_ Missing content types for images.
* `#543 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/543>`_ Make sure images fit containers on all OSes.
* `#544 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/544>`_ Gracefully handle missing cell styles.
* `#546 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/546>`_ ExternalLink duplicated when editing a file with macros.
* `#548 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/548>`_ Exception with non-ASCII worksheet titles
* `#551 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/551>`_ Combine multiple LineCharts


Minor changes
-------------

* `PR 88 <https://foss.heptapod.net/openpyxl/openpyxl/-/merge_requests/88/>`_ Fix page margins in parser.


2.3.0 (2015-10-20)
==================


Major changes
-------------

* Support the creation of chartsheets


Bug fixes
---------

* `#532 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/532>`_ Problems when cells have no style in read-only mode.


Minor changes
-------------

* PR 79 Make PlotArea editable in charts
* Use graphicalProperties as the alias for spPr


2.3.0-b2 (2015-09-04)
=====================


Bug fixes
---------

* `#488 <https://bitbucket.org/openpyxl/openpyxl/issue/488>`_ Support hashValue attribute for sheetProtection
* `#493 <https://bitbucket.org/openpyxl/openpyxl/issue/493>`_ Warn that unsupported extensions will be dropped
* `#494 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/494/>`_ Cells with exponentials causes a ValueError
* `#497 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/497/>`_ Scatter charts are broken
* `#499 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/499/>`_ Inconsistent conversion of localised datetimes
* `#500 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/500/>`_ Adding images leads to unreadable files
* `#509 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/509/>`_ Improve handling of sheet names
* `#515 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/515/>`_ Non-ascii titles have bad repr
* `#516 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/516/>`_ Ignore unassigned worksheets


Minor changes
-------------

* Worksheets are now iterable by row.
* Assign individual cell styles only if they are explicitly set.


2.3.0-b1 (2015-06-29)
=====================


Major changes
-------------

* Shift to using (row, column) indexing for cells. Cells will at some point *lose* coordinates.
* New implementation of conditional formatting. Databars now partially preserved.
* et_xmlfile is now a standalone library.
* Complete rewrite of chart package
* Include a tokenizer for fomulae to be able to adjust cell references in them. PR 63


Minor changes
-------------

* Read-only and write-only worksheets renamed.
* Write-only workbooks support charts and images.
* `PR76 <https://bitbucket.org/openpyxl/openpyxl/pull-request/76>`_ Prevent comment images from conflicting with VBA


Bug fixes
---------

* `#81 <https://bitbucket.org/openpyxl/openpyxl/issue/81>`_ Support stacked bar charts
* `#88 <https://bitbucket.org/openpyxl/openpyxl/issue/88>`_ Charts break hyperlinks
* `#97 <https://bitbucket.org/openpyxl/openpyxl/issue/97>`_ Pie and combination charts
* `#99 <https://bitbucket.org/openpyxl/openpyxl/issue/99>`_ Quote worksheet names in chart references
* `#150 <https://bitbucket.org/openpyxl/openpyxl/issue/150>`_ Support additional chart options
* `#172 <https://bitbucket.org/openpyxl/openpyxl/issue/172>`_ Support surface charts
* `#381 <https://bitbucket.org/openpyxl/openpyxl/issue/381>`_ Preserve named styles
* `#470 <https://bitbucket.org/openpyxl/openpyxl/issue/470>`_ Adding more than 10 worksheets with the same name leads to duplicates sheet names and an invalid file


2.2.6 (unreleased)
==================


Bug fixes
---------

* `#502 <https://bitbucket.org/openpyxl/openpyxl/issue/502>`_ Unexpected keyword "mergeCell"
* `#503 <https://bitbucket.org/openpyxl/openpyxl/issue/503>`_ tostring missing in dump_worksheet
* `#506 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/506>`_ Non-ASCII formulae cannot be parsed
* `#508 <https://foss.heptapod.net/openpyxl/openpyxl/-/issues/508>`_ Cannot save files with coloured tabs
* Regex for ignoring named ranges is wrong (character class instead of prefix)


2.2.5 (2015-06-29)
==================


Bug fixes
---------

* `#463 <https://bitbucket.org/openpyxl/openpyxl/issue/463>`_ Unexpected keyword "mergeCell"
* `#484 <https://bitbucket.org/openpyxl/openpyxl/issue/484>`_ Unusual dimensions breaks read-only mode
* `#485 <https://bitbucket.org/openpyxl/openpyxl/issue/485>`_ Move return out of loop


2.2.4 (2015-06-17)
==================


Bug fixes
---------

* `#464 <https://bitbucket.org/openpyxl/openpyxl/issue/464>`_ Cannot use images when preserving macros
* `#465 <https://bitbucket.org/openpyxl/openpyxl/issue/465>`_ ws.cell() returns an empty cell on read-only workbooks
* `#467 <https://bitbucket.org/openpyxl/openpyxl/issue/467>`_ Cannot edit a file with ActiveX components
* `#471 <https://bitbucket.org/openpyxl/openpyxl/issue/471>`_ Sheet properties elements must be in order
* `#475 <https://bitbucket.org/openpyxl/openpyxl/issue/475>`_ Do not redefine class __slots__ in subclasses
* `#477 <https://bitbucket.org/openpyxl/openpyxl/issue/477>`_ Write-only support for SheetProtection
* `#478 <https://bitbucket.org/openpyxl/openpyxl/issue/477>`_ Write-only support for DataValidation
* Improved regex when checking for datetime formats


2.2.3 (2015-05-26)
==================


Bug fixes
---------

* `#451 <https://bitbucket.org/openpyxl/openpyxl/issue/451>`_ fitToPage setting ignored
* `#458 <https://bitbucket.org/openpyxl/openpyxl/issue/458>`_ Trailing spaces lost when saving files.
* `#459 <https://bitbucket.org/openpyxl/openpyxl/issue/459>`_ setup.py install fails with Python 3
* `#462 <https://bitbucket.org/openpyxl/openpyxl/issue/462>`_ Vestigial rId conflicts when adding charts, images or comments
* `#455 <https://bitbucket.org/openpyxl/openpyxl/issue/455>`_ Enable Zip64 extensions for all versions of Python


2.2.2 (2015-04-28)
==================


Bug fixes
---------

* `#447 <https://bitbucket.org/openpyxl/openpyxl/issue/447>`_ Uppercase datetime number formats not recognised.
* `#453 <https://bitbucket.org/openpyxl/openpyxl/issue/453>`_ Borders broken in shared_styles.


2.2.1 (2015-03-31)
==================


Minor changes
-------------

* `PR54 <https://bitbucket.org/openpyxl/openpyxl/pull-request/54>`_ Improved precision on times near midnight.
* `PR55 <https://bitbucket.org/openpyxl/openpyxl/pull-request/55>`_ Preserve macro buttons


Bug fixes
---------

* `#429 <https://bitbucket.org/openpyxl/openpyxl/issue/429>`_ Workbook fails to load because header and footers cannot be parsed.
* `#433 <https://bitbucket.org/openpyxl/openpyxl/issue/433>`_ File-like object with encoding=None
* `#434 <https://bitbucket.org/openpyxl/openpyxl/issue/434>`_ SyntaxError when writing page breaks.
* `#436 <https://bitbucket.org/openpyxl/openpyxl/issue/436>`_ Read-only mode duplicates empty rows.
* `#437 <https://bitbucket.org/openpyxl/openpyxl/issue/437>`_ Cell.offset raises an exception
* `#438 <https://bitbucket.org/openpyxl/openpyxl/issue/438>`_ Cells with pivotButton and quotePrefix styles cannot be read
* `#440 <https://bitbucket.org/openpyxl/openpyxl/issue/440>`_ Error when customised versions of builtin formats
* `#442 <https://bitbucket.org/openpyxl/openpyxl/issue/442>`_ Exception raised when a fill element contains no children
* `#444 <https://bitbucket.org/openpyxl/openpyxl/issue/442>`_ Styles cannot be copied


2.2.0 (2015-03-11)
==================


Bug fixes
---------
* `#415 <https://bitbucket.org/openpyxl/openpyxl/issue/415>`_ Improved exception when passing in invalid in memory files.


2.2.0-b1 (2015-02-18)
=====================


Major changes
-------------
* Cell styles deprecated, use formatting objects (fonts, fills, borders, etc.) directly instead
* Charts will no longer try and calculate axes by default
* Support for template file types - PR21
* Moved ancillary functions and classes into utils package - single place of reference
* `PR 34 <https://bitbucket.org/openpyxl/openpyxl/pull-request/34/>`_ Fully support page setup
* Removed SAX-based XML Generator. Special thanks to Elias Rabel for implementing xmlfile for xml.etree
* Preserve sheet view definitions in existing files (frozen panes, zoom, etc.)


Bug fixes
---------
* `#103 <https://bitbucket.org/openpyxl/openpyxl/issue/103>`_ Set the zoom of a sheet
* `#199 <https://bitbucket.org/openpyxl/openpyxl/issue/199>`_ Hide gridlines
* `#215 <https://bitbucket.org/openpyxl/openpyxl/issue/215>`_ Preserve sheet view setings
* `#262 <https://bitbucket.org/openpyxl/openpyxl/issue/262>`_ Set the zoom of a sheet
* `#392 <https://bitbucket.org/openpyxl/openpyxl/issue/392>`_ Worksheet header not read
* `#387 <https://bitbucket.org/openpyxl/openpyxl/issue/387>`_ Cannot read files without styles.xml
* `#410 <https://bitbucket.org/openpyxl/openpyxl/issue/410>`_ Exception when preserving whitespace in strings
* `#417 <https://bitbucket.org/openpyxl/openpyxl/issue/417>`_ Cannot create print titles
* `#420 <https://bitbucket.org/openpyxl/openpyxl/issue/420>`_ Rename confusing constants
* `#422 <https://bitbucket.org/openpyxl/openpyxl/issue/422>`_ Preserve color index in a workbook if it differs from the standard


Minor changes
-------------
* Use a 2-way cache for column index lookups
* Clean up tests in cells
* `PR 40 <https://bitbucket.org/openpyxl/openpyxl/pull-request/40/>`_ Support frozen panes and autofilter in write-only mode
* Use ws.calculate_dimension(force=True) in read-only mode for unsized worksheets


2.1.5 (2015-02-18)
==================


Bug fixes
---------
* `#403 <https://bitbucket.org/openpyxl/openpyxl/issue/403>`_ Cannot add comments in write-only mode
* `#401 <https://bitbucket.org/openpyxl/openpyxl/issue/401>`_ Creating cells in an empty row raises an exception
* `#408 <https://bitbucket.org/openpyxl/openpyxl/issue/408>`_ from_excel adjustment for Julian dates 1 < x < 60
* `#409 <https://bitbucket.org/openpyxl/openpyxl/issue/409>`_ refersTo is an optional attribute


Minor changes
-------------
* Allow cells to be appended to standard worksheets for code compatibility with write-only mode.


2.1.4 (2014-12-16)
==================


Bug fixes
---------

* `#393 <https://bitbucket.org/openpyxl/openpyxl/issue/393>`_ IterableWorksheet skips empty cells in rows
* `#394 <https://bitbucket.org/openpyxl/openpyxl/issue/394>`_ Date format is applied to all columns (while only first column contains dates)
* `#395 <https://bitbucket.org/openpyxl/openpyxl/issue/395>`_ temporary files not cleaned properly
* `#396 <https://bitbucket.org/openpyxl/openpyxl/issue/396>`_ Cannot write "=" in Excel file
* `#398 <https://bitbucket.org/openpyxl/openpyxl/issue/398>`_ Cannot write empty rows in write-only mode with LXML installed


Minor changes
-------------
* Add relation namespace to root element for compatibility with iWork
* Serialize comments relation in LXML-backend


2.1.3 (2014-12-09)
==================


Minor changes
-------------
* `PR 31 <https://bitbucket.org/openpyxl/openpyxl/pull-request/31/>`_ Correct tutorial
* `PR 32 <https://bitbucket.org/openpyxl/openpyxl/pull-request/32/>`_ See #380
* `PR 37 <https://bitbucket.org/openpyxl/openpyxl/pull-request/37/>`_ Bind worksheet to ColumnDimension objects


Bug fixes
---------
* `#379 <https://bitbucket.org/openpyxl/openpyxl/issue/379>`_ ws.append() doesn't set RowDimension Correctly
* `#380 <https://bitbucket.org/openpyxl/openpyxl/issue/379>`_ empty cells formatted as datetimes raise exceptions


2.1.2 (2014-10-23)
==================


Minor changes
-------------
* `PR 30 <https://bitbucket.org/openpyxl/openpyxl/pull-request/30/>`_ Fix regex for positive exponentials
* `PR 28 <https://bitbucket.org/openpyxl/openpyxl/pull-request/28/>`_ Fix for #328


Bug fixes
---------
* `#120 <https://bitbucket.org/openpyxl/openpyxl/issue/120>`_, `#168 <https://bitbucket.org/openpyxl/openpyxl/issue/168>`_ defined names with formulae raise exceptions, `#292 <https://bitbucket.org/openpyxl/openpyxl/issue/292>`_
* `#328 <https://bitbucket.org/openpyxl/openpyxl/issue/328/>`_ ValueError when reading cells with hyperlinks
* `#369 <https://bitbucket.org/openpyxl/openpyxl/issue/369>`_ IndexError when reading definedNames
* `#372 <https://bitbucket.org/openpyxl/openpyxl/issue/372>`_ number_format not consistently applied from styles


2.1.1 (2014-10-08)
==================


Minor changes
-------------
* PR 20 Support different workbook code names
* Allow auto_axis keyword for ScatterCharts


Bug fixes
---------

* `#332 <https://bitbucket.org/openpyxl/openpyxl/issue/332>`_ Fills lost in ConditionalFormatting
* `#360 <https://bitbucket.org/openpyxl/openpyxl/issue/360>`_ Support value="none" in attributes
* `#363 <https://bitbucket.org/openpyxl/openpyxl/issue/363>`_ Support undocumented value for textRotation
* `#364 <https://bitbucket.org/openpyxl/openpyxl/issue/364>`_ Preserve integers in read-only mode
* `#366 <https://bitbucket.org/openpyxl/openpyxl/issue/366>`_ Complete read support for DataValidation
* `#367 <https://bitbucket.org/openpyxl/openpyxl/issue/367>`_ Iterate over unsized worksheets


2.1.0 (2014-09-21)
==================

Major changes
-------------
* "read_only" and "write_only" new flags for workbooks
* Support for reading and writing worksheet protection
* Support for reading hidden rows
* Cells now manage their styles directly
* ColumnDimension and RowDimension object manage their styles directly
* Use xmlfile for writing worksheets if available - around 3 times faster
* Datavalidation now part of the worksheet package


Minor changes
-------------
* Number formats are now just strings
* Strings can be used for RGB and aRGB colours for Fonts, Fills and Borders
* Create all style tags in a single pass
* Performance improvement when appending rows
* Cleaner conversion of Python to Excel values
* PR6 reserve formatting for empty rows
* standard worksheets can append from ranges and generators


Bug fixes
---------
* `#153 <https://bitbucket.org/openpyxl/openpyxl/issue/153>`_ Cannot read visibility of sheets and rows
* `#181 <https://bitbucket.org/openpyxl/openpyxl/issue/181>`_ No content type for worksheets
* `241 <https://bitbucket.org/openpyxl/openpyxl/issue/241>`_ Cannot read sheets with inline strings
* `322 <https://bitbucket.org/openpyxl/openpyxl/issue/322>`_ 1-indexing for merged cells
* `339 <https://bitbucket.org/openpyxl/openpyxl/issue/339>`_ Correctly handle removal of cell protection
* `341 <https://bitbucket.org/openpyxl/openpyxl/issue/341>`_ Cells with formulae do not round-trip
* `347 <https://bitbucket.org/openpyxl/openpyxl/issue/347>`_ Read DataValidations
* `353 <https://bitbucket.org/openpyxl/openpyxl/issue/353>`_ Support Defined Named Ranges to external workbooks


2.0.5 (2014-08-08)
==================


Bug fixes
---------
* `#348 <https://bitbucket.org/openpyxl/openpyxl/issue/348>`_ incorrect casting of boolean strings
* `#349 <https://bitbucket.org/openpyxl/openpyxl/issue/349>`_ roundtripping cells with formulae


2.0.4 (2014-06-25)
==================

Minor changes
-------------
* Add a sample file illustrating colours


Bug fixes
---------

* `#331 <https://bitbucket.org/openpyxl/openpyxl/issue/331>`_ DARKYELLOW was incorrect
* Correctly handle extend attribute for fonts


2.0.3 (2014-05-22)
==================

Minor changes
-------------

* Updated docs


Bug fixes
---------

* `#319 <https://bitbucket.org/openpyxl/openpyxl/issue/319>`_ Cannot load Workbooks with vertAlign styling for fonts


2.0.2 (2014-05-13)
==================

2.0.1 (2014-05-13)  brown bag
=============================

2.0.0 (2014-05-13)  brown bag
=============================


Major changes
-------------

* This is last release that will support Python 3.2
* Cells are referenced with 1-indexing: A1 == cell(row=1, column=1)
* Use jdcal for more efficient and reliable conversion of datetimes
* Significant speed up when reading files
* Merged immutable styles
* Type inference is disabled by default
* RawCell renamed ReadOnlyCell
* ReadOnlyCell.internal_value and ReadOnlyCell.value now behave the same as Cell
* Provide no size information on unsized worksheets
* Lower memory footprint when reading files


Minor changes
-------------

* All tests converted to pytest
* Pyflakes used for static code analysis
* Sample code in the documentation is automatically run
* Support GradientFills
* BaseColWidth set


Pull requests
-------------
* #70 Add filterColumn, sortCondition support to AutoFilter
* #80 Reorder worksheets parts
* #82 Update API for conditional formatting
* #87 Add support for writing Protection styles, others
* #89 Better handling of content types when preserving macros


Bug fixes
---------
* `#46 <https://bitbucket.org/openpyxl/openpyxl/issue/46>`_ ColumnDimension style error
* `#86 <https://bitbucket.org/openpyxl/openpyxl/issue/86>`_ reader.worksheet.fast_parse sets booleans to integers
* `#98 <https://bitbucket.org/openpyxl/openpyxl/issue/98>`_ Auto sizing column widths does not work
* `#137 <https://bitbucket.org/openpyxl/openpyxl/issue/137>`_ Workbooks with chartsheets
* `#185 <https://bitbucket.org/openpyxl/openpyxl/issue/185>`_  Invalid PageMargins
* `#230 <https://bitbucket.org/openpyxl/openpyxl/issue/230>`_ Using \v in cells creates invalid files
* `#243 <https://bitbucket.org/openpyxl/openpyxl/issue/243>`_ - IndexError when loading workbook
* `#263 <https://bitbucket.org/openpyxl/openpyxl/issue/263>`_ - Forded conversion of line breaks
* `#267 <https://bitbucket.org/openpyxl/openpyxl/issue/267>`_ - Raise exceptions when passed invalid types
* `#270 <https://bitbucket.org/openpyxl/openpyxl/issue/270>`_ - Cannot open files which use non-standard sheet names or reference Ids
* `#269 <https://bitbucket.org/openpyxl/openpyxl/issue/269>`_ - Handling unsized worksheets in IterableWorksheet
* `#270 <https://bitbucket.org/openpyxl/openpyxl/issue/270>`_ - Handling Workbooks with non-standard references
* `#275 <https://bitbucket.org/openpyxl/openpyxl/issue/275>`_ - Handling auto filters where there are only custom filters
* `#277 <https://bitbucket.org/openpyxl/openpyxl/issue/277>`_ - Harmonise chart and cell coordinates
* `#280 <https://bitbucket.org/openpyxl/openpyxl/issue/280>`_- Explicit exception raising for invalid characters
* `#286 <https://bitbucket.org/openpyxl/openpyxl/issue/286>`_ - Optimized writer can not handle a datetime.time value
* `#296 <https://bitbucket.org/openpyxl/openpyxl/issue/296>`_ - Cell coordinates not consistent with documentation
* `#300 <https://bitbucket.org/openpyxl/openpyxl/issue/300>`_ - Missing column width causes load_workbook() exception
* `#304 <https://bitbucket.org/openpyxl/openpyxl/issue/304>`_ - Handling Workbooks with absolute paths for worksheets (from Sharepoint)


1.8.6 (2014-05-05)
==================

Minor changes
-------------
Fixed typo for import Elementtree

Bugfixes
--------
* `#279 <https://bitbucket.org/openpyxl/openpyxl/issue/279>`_ Incorrect path for comments files on Windows


1.8.5 (2014-03-25)
==================

Minor changes
-------------
* The '=' string is no longer interpreted as a formula
* When a client writes empty xml tags for cells (e.g. <c r='A1'></c>), reader will not crash


1.8.4 (2014-02-25)
==================

Bugfixes
--------
* `#260 <https://bitbucket.org/openpyxl/openpyxl/issue/260>`_ better handling of undimensioned worksheets
* `#268 <https://bitbucket.org/openpyxl/openpyxl/issue/268>`_ non-ascii in formualae
* `#282 <https://bitbucket.org/openpyxl/openpyxl/issue/282>`_ correct implementation of register_namepsace for Python 2.6


1.8.3 (2014-02-09)
==================

Major changes
-------------
Always parse using cElementTree

Minor changes
-------------
Slight improvements in memory use when parsing

* `#256 <https://bitbucket.org/openpyxl/openpyxl/issue/256>`_ - error when trying to read comments with optimised reader
* `#260 <https://bitbucket.org/openpyxl/openpyxl/issue/260>`_ - unsized worksheets
* `#264 <https://bitbucket.org/openpyxl/openpyxl/issue/264>`_ - only numeric cells can be dates


1.8.2 (2014-01-17)
==================

* `#247 <https://bitbucket.org/openpyxl/openpyxl/issue/247>`_ - iterable worksheets open too many files
* `#252 <https://bitbucket.org/openpyxl/openpyxl/issue/252>`_ - improved handling of lxml
* `#253 <https://bitbucket.org/openpyxl/openpyxl/issue/253>`_ - better handling of unique sheetnames


1.8.1 (2014-01-14)
==================

* `#246 <https://bitbucket.org/openpyxl/openpyxl/issue/246>`_


1.8.0 (2014-01-08)
==================

Compatibility
-------------

Support for Python 2.5 dropped.

Major changes
-------------

* Support conditional formatting
* Support lxml as backend
* Support reading and writing comments
* pytest as testrunner now required
* Improvements in charts: new types, more reliable


Minor changes
-------------

* load_workbook now accepts data_only to allow extracting values only from
  formulae. Default is false.
* Images can now be anchored to cells
* Docs updated
* Provisional benchmarking
* Added convenience methods for accessing worksheets and cells by key


1.7.0 (2013-10-31)
==================


Major changes
-------------

Drops support for Python < 2.5 and last version to support Python 2.5


Compatibility
-------------

Tests run on Python 2.5, 2.6, 2.7, 3.2, 3.3


Merged pull requests
--------------------

* 27 Include more metadata
* 41 Able to read files with chart sheets
* 45 Configurable Worksheet classes
* 3 Correct serialisation of Decimal
* 36 Preserve VBA macros when reading files
* 44 Handle empty oddheader and oddFooter tags
* 43 Fixed issue that the reader never set the active sheet
* 33 Reader set value and type explicitly and TYPE_ERROR checking
* 22 added page breaks, fixed formula serialization
* 39 Fix Python 2.6 compatibility
* 47 Improvements in styling


Known bugfixes
--------------

* `#109 <https://bitbucket.org/openpyxl/openpyxl/issue/109>`_
* `#165 <https://bitbucket.org/openpyxl/openpyxl/issue/165>`_
* `#209 <https://bitbucket.org/openpyxl/openpyxl/issue/209>`_
* `#112 <https://bitbucket.org/openpyxl/openpyxl/issue/112>`_
* `#166 <https://bitbucket.org/openpyxl/openpyxl/issue/166>`_
* `#109 <https://bitbucket.org/openpyxl/openpyxl/issue/109>`_
* `#223 <https://bitbucket.org/openpyxl/openpyxl/issue/223>`_
* `#124 <https://bitbucket.org/openpyxl/openpyxl/issue/124>`_
* `#157 <https://bitbucket.org/openpyxl/openpyxl/issue/157>`_


Miscellaneous
-------------

Performance improvements in optimised writer

Docs updated
