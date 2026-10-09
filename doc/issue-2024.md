# Issue 2024: merged-cell borders

Source: [upstream #2024](https://foss.heptapod.net/openpyxl/openpyxl/-/work_items/2024).

The fork already propagated styles assigned after merging, but propagated the
whole border to every constituent cell. Internal formatting during merge also
called the public setter and triggered that propagation again. Loading discarded
the style arrays of non-anchor cells. A single-cell merged range additionally
failed because worksheet indexing returned a Cell rather than iterable rows.

The fix uses the actual merged range, projects each explicit anchor border
assignment onto its perimeter, and bypasses public propagation during internal
border reconstruction. The anchor continues storing the full border definition.
Repeated assignments and deletion replace the old projected sides. NamedStyle
assignments retain all other properties while using the same border projection.
Only a bounded number of border variants are copied for a range.

Loading preserves stored placeholder style arrays and explicit protection,
while rebuilding missing border information. Fresh merge operations retain
existing value-discarding semantics; this patch is not an unmerge-data recovery
feature. Explicit edits to a MergedCell remain local. This intentionally changes
the fork's earlier behavior of putting all four sides on every placeholder.

Regression coverage uses the public Workbook/Worksheet interface, the original
issue example, 1x1 / 1xN / Nx1 / NxM ranges, borders before/after merging,
NamedStyle by name/object, replacement/clearing, two save/load cycles, and
read-only loading of the written XML. The existing merged-style test now really
saves and reloads instead of leaving that step commented out.

Validation on Windows / Python 3.12.10:

- Before fix: 18 failures, 3 passes in the first 21-case regression set.
- After fix: 24 new regression cases pass.
- Full suite, lxml 5.4.0: 2728 passed, 10 skipped, 12 xfailed.
- Full suite, ElementTree + defusedxml 0.7.1: 2732 passed, 6 skipped, 12 xfailed.
- Merged-style/XML selection without lxml or defusedxml: 29 passed, 12 skipped.

Two existing assertions were updated: full borders on every merged cell are no
longer expected; loading complex-styles.xlsx retains existing border records
rather than creating three redundant variants. Workbook data and styles were
validated through XML round trips; desktop Excel visual rendering was not tested.
