# Proposal: consistent structural worksheet edits (#1273)

**Status: design proposal and executable coordinate prototype. Worksheet behavior
is not changed by this branch. #1273 is not claimed fixed.**

Source: [issue #1273](https://foss.heptapod.net/openpyxl/openpyxl/-/work_items/1273).
The issue description was available through the public API; comments returned
HTTP 401. This proposal addresses the published description and inspected code,
not an assumed consensus in the inaccessible comment thread.

## What is broken and why

The four public methods in `openpyxl/worksheet/worksheet.py` ultimately move
entries in `_cells` and update `Cell._coord`. References are separate objects:

| Dependency | Representation in this fork |
| --- | --- |
| Merged ranges | `ws.merged_cells`: MultiCellRange containing mutable MergedCellRange objects, each with a cached start_cell |
| Workbook names | `wb.defined_names`: DefinedNameDict; expressions in attr_text |
| Worksheet-local names | `ws.defined_names`: separate DefinedNameDict |
| Print area | `ws._print_area`: PrintArea, another MultiCellRange |
| Repeated print rows/columns | `ws._print_rows` / `ws._print_cols`: RowRange and ColRange |
| Row dimensions | DimensionHolder indexed by row number |
| Column dimensions | DimensionHolder indexed by letters, possibly spanning a min/max group |

No shared edit notification or transform updates these objects. Forty-eight
acceptance cases confirm the failure through public methods, both in memory and
after a save/load cycle. The cell itself moves. The reference does not. Wrong
merge metadata can subsequently delete moved values during loading.

This rules out a serialization-only explanation: the reference is already stale
in memory. It also rules out `_move_cell` failing to move ordinary cell values.
The missing behavior belongs at the structural-operation level.

## Recommended interface

Keep the current signatures backward compatible, adding a keyword-only opt-in
on **all four** methods, for example:

```python
# Proposed interface; it is not available in this branch.
ws.insert_rows(3, amount=2, update_dependencies=True)
ws.delete_cols(4, amount=1, update_dependencies=True)
```

Initially default to False to avoid double-shifting references in applications
that already compensate for upstream behavior. After a deprecation period, the
fork could make dependency updates the default in a release that announces the
change. A generic claim that these operations are identical to Excel should wait
for the supported-object matrix and Excel fixtures below.

The new implementation should be a private module such as
`openpyxl/worksheet/_structural_edit.py`. A caller selects one behavior flag;
the implementation owns validation, dependency discovery, planning, sparse cell
movement and reference updates. Callers should not need to coordinate a series
of merge/print/formula repair functions themselves.

Do not put the orchestration in `_move_cells`: deletion's negative offset starts
at `idx + amount`, after the deleted band. That helper alone has neither the full
intent nor a safe place to update everything before the gutter is removed.
Route public insert/delete methods to the new operation. Keep `_move_cell` a
private primitive. `move_range` requires separate cut/move semantics, overlap
and target-overwrite policies and should not be silently routed to this model.

## One mathematical transform

Represent a single axis edit with immutable `(index, count, deleting, limit)`.
The included `axis_edit.py` implements **only** this mathematical model.
Coordinates and interval endpoints are 1-based and inclusive.

For insertion at k of n coordinates:

- A point before k is unchanged; a point at/after k moves by +n.
- Insertion at/before a range's first coordinate shifts the whole range.
- Insertion strictly inside a range, including immediately before its final
  coordinate, extends its end.
- Insertion immediately after the range leaves it unchanged.

For deletion of `[k, k+n-1]`:

- Points inside the band become absent; points after it move by -n.
- A partially intersecting interval contracts to its surviving span.
- An interval with no surviving coordinates returns None.
- A zero count is an identity operation.

Examples for the interval `[3, 7]`:

| Operation | Result |
| --- | --- |
| Insert 2 at 3 | [5, 9] |
| Insert 2 at 4 | [3, 9] |
| Insert 2 at 8 | [3, 7] |
| Delete 2 at 4 | [3, 5] |
| Delete 8 at 2 | None |

None is a domain result, not automatically the string `#REF!`. Each adapter
chooses the correct representation for its object. Invalid indices/counts and
reference overflow must be detected during planning. Use the appropriate limit
(1,048,576 rows or 16,384 columns). Do not materialize a full row/column reference
such as `A:A` as a finite rectangle: its orthogonal axis is **all rows**, which
must stay unbounded when rows are inserted. The included prototype covers finite
intervals; whole-axis references need a separate sentinel in the reference model.

## Plan first, then apply

1. Validate the edit and snapshot the relevant object identities and references.
2. Discover every dependent object in the supported document profile, including
   names and formulas on other worksheets that refer to this worksheet.
3. Produce immutable replacement coordinates/attributes without changing the
   workbook or invoking side-effecting constructors like MergedCellRange(ws, ...).
4. Reject unsupported references, overlapping resulting merges and overflows
   before mutating anything. Diagnostics must identify the object and reason.
5. Build a replacement `_cells` dictionary by iterating existing entries only.
   Do not use iter_rows/iter_cols to fill the entire occupied rectangle.
6. Apply cell coordinates, dimensions, rebuilt range collections and rewritten
   references, then rebuild merged placeholders and their anchor pointers.
   Preserve styles using the #2024 fix. Recompute `_current_row` consistently
   after either axis changes so a following append uses the correct row.

This avoids the present sparse-sheet amplification and transient key collisions
while moving cells. Validation failures must leave the workbook unchanged,
including style registries. Commit-time operations must not reparse/validate
strings or run public style setters; those can change other cells. Unexpected
commit exceptions require restoring the captured state, with a rollback test.
No in-memory transaction promises recovery from process termination or MemoryError.

Do not mutate CellRange bounds while that object is in a set or a dictionary key:
its hash depends on the coordinates. Build fresh collections. Likewise, update
both the key and the internal `index`/`min`/`max` of DimensionHolder entries.
Merely moving a dictionary entry is not enough for a correct saved workbook.

## Adapters and deletion policies

| Object | Required behavior |
| --- | --- |
| Merge | Transform bounds, remove wholly deleted merges, rebuild start_cell and placeholders; preserve anchor value/style for ordinary shifts and interior edits. Reject partial deletion of the anchor until its policy is verified against Excel. |
| Print area | Transform each disjoint area, drop empty components, clear the area if none remain. |
| Print titles | Transform the affected row/column interval, clear only the deleted axis and preserve the other one. |
| Defined name | Preserve name, scope and attributes; transform statically parsed references; retain a wholly deleted name with `#REF!`, rather than silently rebinding or deleting it. |
| Row dimensions | Relocate height, hidden, outline and style with the original rows; drop deleted entries. New rows use an explicit documented inheritance policy. |
| Column dimensions | Normalize spans, transform/split groups as needed, rebuild keys and min/max without expanding all 16,384 columns. Preserve widths, hidden state, styles and outline levels. |
| Data validation / conditional formatting | Transform each sqref component and relevant formula expressions; rekey conditional-formatting dictionaries. |
| Tables / filters | Update ref, sort/filter ranges and column metadata together. Header/totals-row deletion or ambiguous table resize must be rejected until implemented. |
| Hyperlinks / worksheet views | Update hyperlink refs and internal locations, selection/sqref, activeCell, freeze panes and pane coordinates where applicable. |
| Drawings / comments / controls | Handle their specific anchors and attachment semantics; these are not ordinary A1 strings. |
| Charts / pivots / connections | Update source references, caches and anchors coherently or reject before edits; do not silently preserve stale caches. |

An initial implementation can support the first six rows on a restricted,
formula-free document profile. It must reject unsupported dependency-bearing
objects that may refer to the edited sheet. Macro-enabled workbooks and opaque
control/external-connection content are outside that initial safe profile.
Even a small metadata fix should not be advertised as complete Excel behavior.

Unverified choices requiring Excel-generated before/after fixtures include:
partial removal of the merged anchor, insertion adjacent to range boundaries,
format inheritance for new rows/columns, grouped column splitting, single-cell
merge normalization and table header/totals behavior. Rejection is preferable
to silently selecting a data-loss policy. The prototype's interval examples are
proposed rules, not evidence from a desktop Excel oracle.

## Formulas need reference rewriting, not evaluation

The existing Translator implements copy/move translation. It is insufficient for
structural edits: `$A$5` must become `$A$6` when a row is inserted above row 5,
although the dollar signs prevent that change during formula copying.

Use Tokenizer as a lexer, then parse RANGE operands into a reference model with
sheet/workbook identity, optional bounds and absolute/relative flags. Preserve
sheet quoting/escaped apostrophes and dollar signs. Visit formulas throughout the
workbook, not just the moved cells. Account for sheet-local name shadowing.
Constants and text literals remain unchanged. A fully deleted reference becomes
`#REF!`; formulas are not evaluated and recalculation flags must be set.

Static single-sheet A1 references and unions are a sensible first target. Dynamic
names (OFFSET/INDIRECT), external workbooks, 3D references, structured table
references, array/data-table formulas and spill references need dedicated
handling or an explicit rejection. Unqualified global relative names lack an
unambiguous worksheet context and must not be guessed. Broad regex replacement
of every apparent cell address would corrupt strings and scoped references.

A shared-coordinate object or an observer attached to Cell is insufficient:
most references are ranges, some target empty cells, and many exist only as text
in other objects. A central operation with explicit adapters keeps knowledge of
edit semantics in one place without forcing every cell access to maintain a
workbook-wide graph. Start with scanning dependencies per edit; add an index only
if profiling demonstrates the need, with clear invalidation rules for mutations.

## Implementation increments and acceptance gates

1. Implement immutable transforms and parsers with bounded/unbounded references,
   scope resolution and overflow checks. Prototype tests are a starting point.
2. Integrate sparse cell movement, merged ranges, both name scopes, print settings
   and dimensions for the restricted profile. Turn the 48 contract cases below
   into ordinary passing tests when the opt-in interface is added.
3. Add workbook-wide static formula rewriting, validators and conditional rules;
   keep unsupported syntax/object rejection explicit and atomic.
4. Add tables, views, drawings, charts and pivots one dependency type at a time,
   backed by Excel-generated fixtures. Expand the supported profile accordingly.
5. Consider changing the default only with release notes and migration guidance.

Every increment needs insert/delete on both axes, before/inside/after ranges,
complete deletion, amount=0, invalid/overflow input, cross-sheet references,
escaped titles, repeated save/load cycles and unchanged-state-on-rejection tests.
Add a sparse case with only two far-apart cells and assert no intervening cells
are created. Check both XML backends. Existing unrelated tests must remain green.

## Executable evidence in this branch

- `openpyxl/worksheet/tests/test_structural_edit_contract.py`: 48 strict xfails
  for the **current public methods**, covering all four operations and six
  dependency types, in memory and after serialization. They fail on assertion,
  not import/setup errors. They intentionally remain failing under `--runxfail`.
- `doc/proposals/issue1273/axis_edit.py`: pure, unintegrated finite-interval model.
- `doc/proposals/issue1273/test_axis_edit.py`: 21 passing tests, including an
  independent enumerated-point oracle over 9,360 interval/edit combinations.

From the repository root, using a Python environment with project dependencies:

```text
python -m pytest -q doc/proposals/issue1273/test_axis_edit.py
python -m pytest -q openpyxl/worksheet/tests/test_structural_edit_contract.py
python -m pytest -q openpyxl/worksheet/tests/test_structural_edit_contract.py --runxfail
```

The last command is supposed to fail until the worksheet integration exists.
The first two gave **21 passed, 48 xfailed**; forcing the acceptance cases with
`--runxfail` gave **48 failed**. That distinction prevents the mathematical sketch
from being mistaken for a completed workbook-editing implementation.


Final full repository run in this branch (Python 3.12.10, lxml 5.4.0):
**2749 passed, 10 skipped, 60 xfailed**. The 60 expected failures comprise the
12 pre-existing ones plus this proposal's 48 documented acceptance cases.
