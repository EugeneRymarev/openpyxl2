# Issue 1273: structural editing without formula rewriting

Implemented in `codex/fix-1273-structural-edits`. The earlier upstream and #2024
commits have been integrated into local `master` at `ad9d318f6`. This change
remains separate for review.

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

Formula text in cells and calculated name expressions is unchanged. This scope
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
- Local master exactly matches `codex/fix-2024-merged-borders` (`ad9d318f6`),
  verified with an ancestry check and an empty tree diff.

No branches were pushed to GitHub. The current implementation is in its own
branch; master contains only the earlier upstream and #2024 changes.
