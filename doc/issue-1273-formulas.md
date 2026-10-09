# Formula references during structural edits

Branch: `codex/fix-1273-formula-references`, based on master `54ddb6cfd`.
This follow-up extends #1273 without changing `move_range` or copy translation.

All four insert/delete operations default to `update_formulas=True`. References
to the edited sheet are adjusted in surviving ordinary formulas across the
workbook and in calculated defined names. Dollar signs restrict copying, not
structural editing: inserting two rows at 5 changes `=A1+$B$5` in A10 to
`=A1+$B$7` in A12. A formula on another sheet referring to A10 changes to A12.

The parser works on reference operands, preserving strings, functions, operators,
sheet quoting and whitespace. It shares literal reference parsing and axis
mapping with the metadata implementation. It supports finite cells/ranges,
whole rows/columns, intersections, unions and ordinary-string @/spill references.
Ranges expand or contract; completely deleted references become #REF!. Full-axis
references remain full-axis; overflowing formula references become #REF! or are
clipped to the surviving rectangle. Literal-name and cell/metadata overflow
retain their existing strict policy.

The workbook edit plan stages formula strings before changing any cell. Commit
and rollback cover references on other sheets, calculated names and calculation
properties as well as the previous metadata. New full-calculation flags do not
override manual calculation mode or iterative calculation settings. openpyxl2
does not calculate formula values or generate their cached results.

`update_formulas=False` retains the preceding formula-preserving behavior while
updating metadata. `update_dependencies=False` retains legacy cell-only edits.

## Limits

- 3D references, ranges with named/function endpoints, unparseable formulas and
  array/data-table formula objects raise `FormulaTranslationError` before the
  workbook is changed. The error identifies the cell or defined name. Array and
  data-table objects anywhere in the workbook block formula-enabled edits.
- External-book links and structured table references are unchanged. This work
  does not update table definitions, chart formulas, validation, conditional
  formatting or other object types outside the initial #1273 implementation.
- INDIRECT strings stay literal. OFFSET's direct reference is updated, but its
  numeric arguments are not interpreted. Dynamic addressing is not evaluated.
- Unqualified workbook-level references have no reliable sheet context and
  remain unchanged. Names use their stored A1 coordinates; active-cell/caller
  semantics of relative names are not reconstructed. Use absolute references
  in names for predictable automatic maintenance.
- Recognising `A1#` in an ordinary formula does not maintain dynamic-array
  output metadata. Shared formulas expanded to ordinary strings by the loader
  use the same ordinary-formula path.
- The tokenizer is not a full Excel grammar validator. Unsupported reference
  syntax that is explicitly detected fails early; this is not support for every
  language construct or every future Excel formula extension.
- Each edit scans materialized cells throughout the workbook, plus formula
  tokens; it does not create a dense grid or maintain a dependency graph.

## Validation

The public-method regression suite covers all four operations, ranges and
boundaries, dollar signs, quoting/whitespace/string literals, formulas on other
sheets, named expressions, two save/load cycles, sparse sheets, opt-outs,
unsupported constructs and rollback after a cross-sheet commit failure.

Microsoft Excel 16.0 was used as an independent comparison: **34 structural
editing cases matched** actual Excel insert/delete operations. Formula strings
were normalized through Excel to ignore equivalent spelling such as reversed
range endpoints. A separate generated workbook was recalculated in Excel to
check seven local/cross-sheet and named-expression numeric results.

Final validation on Windows / Python 3.12.10:

- 75 new parametrized regression cases pass. The earlier formula-preservation
  tests now explicitly select `update_formulas=False`.
- Full suite with lxml 5.4.0: **2940 passed, 10 skipped, 12 xfailed**.
- Full suite with lxml disabled / defusedxml 0.7.1: **2944 passed, 6 skipped,
  12 xfailed**.
- All 12 expected failures pre-date these changes; there are no unexpected
  failures or new expected-failure markers.
- Black's Python 3.8 formatting check passes for all five changed Python files.
- `git diff --check` passes. Native comparison and recalculation logs are kept
  locally under the ignored `.venv/codex-audit/` directory.

The implementation remains in its separate branch. Master stays at `54ddb6cfd`;
no branch has been pushed to GitHub.
