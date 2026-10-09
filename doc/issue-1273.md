# Issue 1273: structural editing

The initial implementation below was integrated into master without formula
rewriting. A subsequent implementation from `codex/fix-1273-formula-references`,
now also included in local master, adds formula maintenance by default,
with `update_formulas=False` retaining the
earlier behavior. See [formula support and validation](issue-1273-formulas.md).

Implemented in `codex/fix-1273-structural-edits` and integrated into local
`master` by fast-forward through `b63fe1faa`. Master includes the upstream fixes,
#2024, the initial #1273 implementation and formula maintenance from `315acec29`.

All four public insert/delete methods now synchronize merged ranges, static
global/local names (including names in another worksheet that refer to this
one), print areas/titles and row/column dimensions. `update_dependencies=True`
is the default; False retains the previous implementation for applications that
already adjust these objects themselves.

A common immutable axis transform handles shifts, contraction, complete
deletion and bounds. A private edit plan builds replacement cell/range maps,
dimensions and literal references before committing them. It avoids filling
rectangular gaps in sparse sheets, preserves surviving Cell identities and
rebuilds hash-based range collections rather than mutating their keys.

Planning does not change existing style registries or dimensions. Border
variants for resized merges are staged independently. Unexpected commit errors
restore previous collections, coordinates, styles and names; validation failures
occur before any mutation. A regression injects an error after partial commit
to verify the rollback path.

Partial merge deletion deliberately retains the original anchor content/style
at the surviving upper-left cell. Entire merge deletion discards it; a one-cell
merge is retained. Resizing restores the perimeter using the anchor's border.
Unqualified global names, calculated names and external/3D references remain
unchanged. Wholly deleted literal references become #REF!.

In the initial implementation, formula text and calculated names were unchanged. This scope
does not synchronize tables, charts/pivots, filters, validation/conditional
formatting, hyperlinks, drawings/controls or views. `move_range` and private
cell-moving primitives retain their earlier semantics. These limitations and
the dimension inheritance/deletion policies are documented in
[editing_worksheets.rst](editing_worksheets.rst).

Testing includes all 48 originally failing issue cases, additional public-method
boundary/round-trip tests, formula preservation, failure atomicity, injected
rollback and the 9,360-case mathematical oracle. Desktop Excel behavior was not
used as an oracle; the documented merge/dimension policies are explicit choices.


## Validation results

Windows, Python 3.12.10, pytest 9.1.1:

- Full suite with lxml 5.4.0: **2865 passed, 10 skipped, 12 xfailed**.
- Full suite with lxml disabled / defusedxml 0.7.1: **2869 passed, 6 skipped,
  12 xfailed**.
- The 12 expected failures pre-date this work. None of the #1273 acceptance
  cases remain marked xfail.
- The new module and regression files pass Black's Python 3.8 formatting check.
- Three visual scenarios were checked using Microsoft Excel 16.0, build 20430:
  borders assigned after merging, row insertion inside a merge and column
  deletion including the original anchor. All three native PDF-rendered pages
  were inspected: outlines and text were intact, without clipping or overlap.
  Excel also verified the resulting merges, static names, dimensions and print
  settings. The workbook was opened read-only and its SHA-256 remained unchanged.
  This checks the generated workbook, not equivalence with Excel's own editing
  policies. Two openpyxl save/load cycles also passed before the Excel check.

No branches were pushed to GitHub. The implementation branch is retained;
all three requested code changes are now integrated into local master.
