# openpyxl2

[Русская версия](README.ru.md)

**openpyxl2** is a fork of [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl), a Python library for reading and writing Excel `.xlsx`, `.xlsm`, `.xltx` and `.xltm` files. It improves row and column editing, maintains formula references, and preserves formatting of merged cells.

The Python package and import name remain **`openpyxl`**. Python 3.8 or newer is required; the repository's CI matrix covers Python 3.8–3.14. The code is based on the upstream 3.2 development line and reports version `3.2.0b1`.

## Installation and a first workbook

Install this repository in a virtual environment to use the fork:

```sh
git clone https://github.com/EugeneRymarev/openpyxl2.git
cd openpyxl2
python -m pip install .
```

`pip install openpyxl` from PyPI installs the original distribution. The fork uses the same distribution name, so keep the two in separate environments when comparing them. Excel is not required to generate or read files; a spreadsheet application is needed to calculate formulas and view the screenshots' results.

```python
from openpyxl import Workbook, load_workbook

wb = Workbook()
ws = wb.active
ws.title = "Sales"
ws.append(["Item", "Units"])
ws.append(["Pens", 10])
ws.append(["Folders", 20])
ws["B4"] = "=SUM(B2:B3)"
ws.insert_rows(2)  # B5 now contains =SUM(B3:B4)
wb.save("sales.xlsx")

loaded = load_workbook("sales.xlsx")
assert loaded["Sales"]["B5"].value == "=SUM(B3:B4)"
loaded.close()
```

## Differences from original openpyxl

This overview compares the fork with the upstream revision imported into its history, [`1fa786a9d`](https://github.com/EugeneRymarev/openpyxl2/commit/1fa786a9dd5d77754d703f967832f5243e08ba0a). The examples below use fork revision `4c2bce2eb6f4682a1c3bafbb00e8802e3aac88a5`. This is a reproducible comparison of two source revisions, not a claim that every upstream release behaves identically.

- **Row and column edits maintain dependencies.** `insert_rows`, `insert_cols`, `delete_rows` and `delete_cols` update merged ranges, static workbook and worksheet names, print areas, print titles, row heights and column widths, including stored hidden/grouping/style properties. Supported literal names include unions and whole-row/whole-column references.
- **Formulas follow structural edits.** Direct A1 references in ordinary cell formulas throughout the workbook and in calculated defined names are adjusted, including absolute and mixed references. Ranges expand or shrink; references that are completely deleted become `#REF!`. Dollar signs still control copying, but do not prevent structural updates.
- **Merged ranges survive partial deletion.** A partially deleted merge shrinks and retains the original anchor's content and style at the surviving upper-left cell. A fully deleted merge is removed. Single-cell merged ranges are supported, and resized ranges receive the corresponding outer borders.
- **Edits are planned before mutation.** Supported-reference and bounds checks happen before applying the edit. The implementation stages changes and rolls them back on commit failures. Sparse worksheets stay sparse, surviving `Cell` objects retain their identity, and subsequent `append()` uses the updated worksheet position.
- **Styles assigned after merging propagate.** Assigning font, fill, alignment, protection, number format or a named style to the top-left cell updates the merged range. Borders are projected onto its perimeter; the anchor retains the full border definition. Replacing or clearing the anchor border updates previously projected edges. Explicit edits to a non-anchor `MergedCell` remain local.
- **Stored merged-cell styles survive loading.** Individual styles and explicit protection of non-anchor merged cells are preserved through save/load cycles. Missing border information is reconstructed without replacing those stored styles. Merging still discards non-anchor values; this feature does not recover data when unmerging.
- **Merged-range lookup by coordinate.** `MultiCellRange.__getitem__`, including `ws.merged_cells["C3"]`, returns the containing range as a coordinate **string**. A coordinate outside all ranges raises `CellNotMergedException`.
- **Named styles compare by name.** A `NamedStyle` can be compared with a string or another `NamedStyle`. Equality compares names, not the formatting properties. Other operand types raise `NotNamedStyleOrStrException`.
- **Named styles are rebound before serialization.** Their font, fill, border and other component indexes are resolved against the workbook being saved, avoiding indexes from another workbook when a style object is reused. This does not make shared mutable styles thread-safe.
- **The `openpyxl.open` alias was removed.** Use `openpyxl.load_workbook`. Imports were also made explicit; code should import objects from their defining modules rather than rely on incidental re-exports.
- **Runtime corrections.** Fixes cover mutable default arguments and shared state, missing returns and ineffective statements, mismatched method signatures, incorrect variable/class references, typos, built-in name shadowing, Decimal imports and numeric-type tuple construction, and theme XML namespace serialization.
- **XML backend corrections.** The standard-library `iterparse` fallback is available without optional `defusedxml`. XML tests exercise the selected backend and use appropriate byte input for lxml. Dependency constraints were updated, including Python 3.14-compatible lxml and NumPy selections.
- **Upstream maintenance fixes were ported.** Worksheet ZIP streams and normal-mode workbook archives are closed on the covered success/error paths, reducing resource leaks and Windows file locks. MIME lookup is case-insensitive while preserving the extension written into the package. These are upstream fixes carried into this fork, not features unique to it. Successful read-only loads still require the caller to close the workbook.
- **Development and compatibility work.** Code and imports were reformatted with Black and import-sorting tools; line endings and repository configuration were normalized. Regression coverage was added for structural edits, formulas, merged styles, resource handling and XML backends. CI includes Python 3.8–3.14, Windows/Linux checks for 3.14, and installation of a built wheel without optional XML packages.

The fork also inherits upstream 3.2 development changes, such as the cell `Coordinate` model, external connections, controls and volatile dependencies. Those are part of its upstream base, not original additions by this fork. For the history of that comparison, see the [upstream audit](doc/upstream-audit-2026-10-09.md).

## Before and after in Microsoft Excel

All ten images below are screenshots of **actual workbooks opened in Microsoft Excel for Windows**. Both versions run the same Python operations on the same data. “Before” means the pinned original openpyxl revision; “After” means this fork. The comparison shows library-generated files opened in Excel, not Excel's own Insert/Delete commands.

The starting table contains 10 pens and 20 folders. `C9` contains `=SUM(C6:C7)`, so the expected total is **30**. Its title is merged across `B3:E3`. Each operation starts from a fresh copy of this table:

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.merge_cells("B3:E3")
ws["B3"] = "OFFICE SUPPLIES"
for row, values in (
    (5, ["Item", "Units", "Unit price", "Batch"]),
    (6, ["Pens", 10, 2, "A"]),
    (7, ["Folders", 20, 3, "B"]),
):
    for column, value in enumerate(values, 2):
        ws.cell(row, column, value)
ws["B9"] = "Total units"
ws["C9"] = "=SUM(C6:C7)"
# Apply one of the operations below, then save the workbook.
```

The [complete generator](doc/readme/build_examples.py) adds the formatting, dimensions, defined names, print areas and explanatory captions used in the screenshots. The formula bar in this Excel installation displays the localized name `СУММ`; the stored XLSX formula is `SUM`.

### 1. `insert_rows`

```python
ws.insert_rows(3)
wb.save("insert_rows.xlsx")
```

Previously, the cells moved but the formula in `C10` stayed `=SUM(C6:C7)`, giving **10**, and the merge stayed at `B3:E3`. Now the formula becomes `=SUM(C7:C8)`, giving **30**, and the title's merge moves to `B4:E4`.

| Before: original openpyxl                                                                     | After: openpyxl2                                                                                   |
|-----------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------|
| ![insert_rows before: total 10 and displaced title](doc/readme/images/insert_rows_before.jpg) | ![insert_rows after: total 30 and correctly merged title](doc/readme/images/insert_rows_after.jpg) |

### 2. `insert_cols`

```python
ws.insert_cols(2)
wb.save("insert_cols.xlsx")
```

Previously, `D9` still summed `C6:C7`, now containing item names, and displayed **0**. The old `B3:E3` merge also hid the moved title when Excel opened the file. Now `D9` sums `D6:D7` and displays **30**; the merge moves to `C3:F3`, and the column dimensions follow the table.

| Before: original openpyxl                                                                  | After: openpyxl2                                                                          |
|--------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------|
| ![insert_cols before: total 0 and missing title](doc/readme/images/insert_cols_before.jpg) | ![insert_cols after: total 30 and visible title](doc/readme/images/insert_cols_after.jpg) |

### 3. `delete_rows`

```python
ws.delete_rows(2)
wb.save("delete_rows.xlsx")
```

Previously, the formula moved to `C8` but kept `=SUM(C6:C7)`, giving **20**. The title moved away from its merge. Now `C8` contains `=SUM(C5:C6)`, giving **30**, and the merge moves to `B2:E2`.

| Before: original openpyxl                                                                     | After: openpyxl2                                                                                   |
|-----------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------|
| ![delete_rows before: total 20 and displaced title](doc/readme/images/delete_rows_before.jpg) | ![delete_rows after: total 30 and correctly merged title](doc/readme/images/delete_rows_after.jpg) |

### 4. `delete_cols`

```python
ws.delete_cols(1)
wb.save("delete_cols.xlsx")
```

Previously, `B9` still summed `C6:C7`, now the unit prices, giving **5**. The original narrow column A clipped the moved labels, and the merge stayed at `B3:E3`. Now `B9` sums `B6:B7`, giving **30**; the merge becomes `A3:D3`, and the widths move with the columns.

| Before: original openpyxl                                                                   | After: openpyxl2                                                                             |
|---------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------|
| ![delete_cols before: total 5 and clipped labels](doc/readme/images/delete_cols_before.jpg) | ![delete_cols after: total 30 and preserved widths](doc/readme/images/delete_cols_after.jpg) |

### 5. `merge_cells` and formatting after merging

This independent example assigns formatting to the anchor **after** creating each merge:

```python
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

wb = Workbook()
ws = wb.active
blue = Side(style="medium", color="FF2563EB")
for area, anchor in (("B3:E6", "B3"), ("B9:E9", "B9")):
    ws.merge_cells(area)
    cell = ws[anchor]
    cell.value = f"Merged {area}"
    cell.font = Font(name="Calibri", size=18, bold=True, color="FF17365D")
    cell.fill = PatternFill("solid", fgColor="FFDBEAFE")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(left=blue, right=blue, top=blue, bottom=blue)
wb.save("merged.xlsx")
```

Previously, only the anchor stored the assigned border, so Excel displayed an incomplete outline. Now both the rectangular and single-row merges have a **complete outer border**. Excel already displays the anchor's fill and centered text across the old merge; the visible correction here is the border. The fork additionally propagates the other style properties in the underlying cells and preserves stored merged-cell styles on reload.

| Before: original openpyxl                                                                | After: openpyxl2                                                                     |
|------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------|
| ![merge_cells before: incomplete blue borders](doc/readme/images/merge_cells_before.jpg) | ![merge_cells after: complete blue borders](doc/readme/images/merge_cells_after.jpg) |

### Reproduce the comparison

Download the [original-version workbook](doc/readme/workbooks/before.xlsx) and the [fork workbook](doc/readme/workbooks/after.xlsx). Each has five sheets named after the demonstrated operations. Open them in Excel and allow formula recalculation to see the totals; openpyxl itself does not calculate cached results.

To regenerate the workbooks from a checkout containing the baseline Git history and installed dependencies:

```sh
python doc/readme/build_examples.py
```

The script extracts the pinned upstream source into a temporary directory, runs it and the current checkout separately, and checks the resulting formulas and border differences. It records source revisions, merges, names and print areas in [provenance.json](doc/readme/provenance.json). Screenshots are captured separately in Excel; running the script does not capture or refresh them. The [capture manifest](doc/readme/screenshots.json) records image dimensions and SHA-256 hashes of the screenshots and displayed workbooks. Both books contain synthetic sample data only.

## Additional API examples

```python
from openpyxl import Workbook
from openpyxl.styles import NamedStyle

wb = Workbook()
ws = wb.active
ws.merge_cells("B2:D4")
assert ws.merged_cells["C3"] == "B2:D4"  # str, not CellRange

style = NamedStyle(name="Heading")
assert style == "Heading"
assert style == NamedStyle(name="Heading")  # compares names only
```

By default, all four structural methods use `update_dependencies=True` and `update_formulas=True`. These switches let existing applications keep responsibility for part or all of the update:

```python
ws.insert_rows(5, amount=2)  # Update supported metadata and formulas.
ws.insert_rows(5, amount=2, update_formulas=False)  # Keep formula text.
ws.insert_rows(5, amount=2, update_dependencies=False)  # Legacy cell movement.
```

These are alternative calls, not a recipe to apply all three to the same sheet. `update_dependencies=False` also disables formula rewriting. The same options are available for `insert_cols`, `delete_rows` and `delete_cols`.

## Scope and limitations

- Structural edits do not update tables, chart/pivot references, filters, data validation, conditional formatting, hyperlinks, drawings/controls or worksheet views. `move_range` retains its separate behavior. Merged-range and dimension policies are documented in [editing worksheets](doc/editing_worksheets.rst).
- Formula updates support direct A1 cells/ranges, whole axes, unions/intersections and ordinary-string `@`/spill references. External-workbook references and structured table references to remain unchanged. Text inside `INDIRECT` is literal; dynamic addressing is not evaluated.
- Detected unsupported constructs, including 3D references, named/function range endpoints, unparseable formulas and array/data-table formula objects, raise `FormulaTranslationError` before mutation. Array/data-table objects anywhere in the workbook block formula-enabled edits. The tokenizer is not a complete Excel grammar validator.
- Unqualified workbook-level names have no reliable sheet context and remain unchanged. Relative names use their stored A1 coordinates; prefer absolute references for predictable maintenance. Recognizing spill syntax does not maintain dynamic-array output metadata.
- The library rewrites formula text and requests recalculation; it does not evaluate formulas or write their cached results. Existing manual/iterative calculation settings are preserved. Each formula-enabled edit scans materialized cells across the workbook; there is no dependency graph.
- XML hardening is optional: install `defusedxml` when processing untrusted XML. `lxml` is an optional backend, not a requirement for using the library.

See [formula support](doc/issue-1273-formulas.md), [structural edits](doc/issue-1273.md) and [merged-cell borders](doc/issue-2024.md) for implementation details and recorded regression results. Those records describe their respective revisions; the README describes the combined behavior.

## Development, documentation and license

```sh
python -m pip install -r requirements.txt
python -m pip install -e .
python -m pytest
```

The [upstream documentation](https://openpyxl.pages.heptapod.net/openpyxl/) explains the shared API. Use this README and the repository's `doc/` notes for fork-specific behavior. Report fork issues in the [openpyxl2 tracker](https://github.com/EugeneRymarev/openpyxl2/issues).

Distributed under the [MIT license](LICENCE.rst). Original openpyxl authors and contributors are listed in [AUTHORS.rst](AUTHORS.rst); the project was initially inspired by PHPExcel.
