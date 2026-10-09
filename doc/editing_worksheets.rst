Inserting and deleting rows and columns, moving ranges of cells
===============================================================


Inserting rows and columns
--------------------------

You can insert rows or columns using the relevant worksheet methods:

    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.insert_cols`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_rows`
    * :func:`openpyxl.worksheet.worksheet.Worksheet.delete_cols`

The default is one row or column. For example to insert a row at 7 (before
the existing row 7)::

    >>> ws.insert_rows(7)


Deleting rows and columns
--------------------------

To delete the columns ``F:H``::

    >>> ws.delete_cols(6, 3)

Updating ranges and dimensions
------------------------------

In this fork, all four insert/delete methods update these objects by default:

* merged ranges and their anchor/placeholder cells;
* static workbook-level and worksheet-local defined names referring to the
  edited worksheet, including names scoped to a different worksheet;
* print areas and repeated print rows/columns;
* row heights, column widths, hidden state, grouping and dimension styles.

Cells are moved without filling gaps in sparse worksheets. Surviving Cell
objects retain their identity. The append position is recalculated after both
row and column edits, including deletion of the last populated column.

Applications that already adjust ranges themselves can keep the previous
cell-only behavior::

    >>> ws.insert_rows(7, update_dependencies=False)

Range rules
~~~~~~~~~~~

An insertion before or at the first coordinate shifts the whole range. An
insertion inside it extends its end; one immediately after it leaves the range
unchanged. A partial deletion contracts the range to its surviving part.

When part of a merged range survives, the original anchor's value and style are
retained at the new upper-left coordinate, even when its original row/column is
deleted. A completely deleted merge loses its cells and definition. A merge
that shrinks to one cell is retained as a one-cell merge. Resizing reconstructs
the outline from the anchor, preserving other styles on surviving placeholders;
a pure shift preserves their individual borders too.

Empty print areas and title ranges are cleared. Deleted dimension entries are
removed. Completely deleted named references become ``#REF!``; names, scope and
other attributes are retained. In a multi-area name, each deleted component
becomes ``#REF!`` rather than silently disappearing from the definition.

Inserted rows/independent columns have default dimensions. Insertion inside an
existing grouped column span extends that span and inherits its settings.
Overlapping explicit column spans are rejected before any changes are applied.
These are defined policies of this fork, not a guarantee of identical behavior
for every desktop Excel editing command.

Supported defined names
~~~~~~~~~~~~~~~~~~~~~~~

Literal A1 cells, rectangular ranges, unions, whole-row and whole-column
references are updated while preserving dollar signs and sheet quoting. Sheet
names are matched case-insensitively. A local unqualified reference belongs to
its owning worksheet. Global unqualified references have no reliable worksheet
context and remain unchanged. External-workbook and 3D references remain
unchanged as well.

.. important::

    Formulas are not rewritten or evaluated, even when their cells move.
    Calculated names, such as OFFSET/INDIRECT expressions, are also unchanged.
    This release does not update tables, filters, charts, pivot caches, data
    validation, conditional formatting, hyperlinks, drawings, controls or view
    references. Applications using those objects must continue managing their
    dependencies. ``move_range`` and the private ``_move_cells`` helper retain
    their existing behavior.

Validation and failure handling
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

With dependency updates enabled, indices must be positive integers within the
Excel row/column limits; amounts must be non-negative integers. Booleans are not
accepted as integers. Zero amount is a no-op. Deletion beyond the axis limit is
rejected. Insertion that would move a supported reference or cell beyond the
axis limit raises OverflowError.

The operation prepares replacements before modifying cells or metadata. Invalid
ranges, overlapping merges/dimension groups and overflow leave the workbook
unchanged, including its style registries. An unexpected exception during commit
restores the prior state. There is no recovery guarantee for process termination
or exhaustion of memory.


Moving ranges of cells
----------------------

You can also move ranges of cells within a worksheet::

    >>> ws.move_range("D4:F10", rows=-1, cols=2)

This will move the cells in the range ``D4:F10`` up one row, and right two
columns. The cells will overwrite any existing cells.

If cells contain formulae you can let openpyxl translate these for you, but
as this is not always what you want it is disabled by default. Also only the
formulae in the cells themselves will be translated. References to the cells
from other cells or defined names will not be updated; you can use the
:doc:`formula` translator to do this::

    >>> ws.move_range("G4:H10", rows=1, cols=1, translate=True)

This will move the relative references in formulae in the range by one row and one column.


Merge / Unmerge cells
---------------------

When you merge cells all cells but the top-left one are **removed** from the
worksheet. To carry the border-information of the merged cell, the boundary cells of the
merged cell are created as MergeCells which always have the value None.
See :ref:`styling-merged-cells` for information on formatting merged cells.

.. :: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.merge_cells('A2:D2')
>>> ws.unmerge_cells('A2:D2')
>>>
>>> # or equivalently
>>> ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=4)
>>> ws.unmerge_cells(start_row=2, start_column=1, end_row=4, end_column=4)
