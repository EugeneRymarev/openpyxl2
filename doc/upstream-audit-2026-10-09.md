# openpyxl2 / upstream audit (2026-10-09)

## Revisions and method

- Fork inspected: Git `f35829b05431f899993a2e71ca35494acf6fb3f9`.
- Imported upstream baseline: Git `1fa786a9d` (2025-02-13), merged into the
  fork by `e1588309e` on 2025-06-09. The fork already reports **3.2.0b1**.
- Current upstream 3.2: Mercurial `8ea4ebcb2440853e0a85e3b990678730bbc78534`
  (2026-07-11). Mercurial and converted Git hashes are different identities.
- Current upstream 3.1: Mercurial `c7b9026dab21c7fae28c45c2366c0fb3edc0319f`.
- Actual configured default branch: **branch/default**, Mercurial
  `52c77fdee169dceefc11aceef956aa0367e9759a` (2025-10-01), version 3.1.5.
  The project API, not an assumption about branch names, identifies this default.

Sources: [3.2 revision](https://foss.heptapod.net/openpyxl/openpyxl/-/commit/8ea4ebcb2440853e0a85e3b990678730bbc78534),
[public branch API](https://foss.heptapod.net/api/v4/projects/322/repository/branches),
[issue 2024](https://foss.heptapod.net/openpyxl/openpyxl/-/work_items/2024),
[issue 1273](https://foss.heptapod.net/openpyxl/openpyxl/-/work_items/1273).
The issue descriptions and repository snapshot were retrieved via the public
Heptapod API. Comment/discussion endpoints returned HTTP 401; the full comment
threads could not be verified. No conclusions about the maintainer's motives
or release schedule follow from the age of the branch.

Compared the pinned 3.2 archive with the imported Git baseline, normalizing
line endings and comparing Python ASTs to exclude formatting-only changes.
Inspected fork history, relevant implementations, and exercised public workbook
operations. This is a targeted maintenance audit, not a proof of every module.

## Is 3.2 significant?

First distinguish the actual default branch, development heads, and releases.
Against **branch/default**, 3.2 has seven new non-test Python modules and 43
existing non-test Python files with AST differences (including metadata). New
modules include cell coordinates, external connections, drawing anchors/legacy
content and volatile dependencies. This is a substantial difference, and the
user's observation about changes absent from the actual default branch is valid.
Both archives' embedded Mercurial node/branch metadata was checked.

Separately, comparing the two development branches: The verified
archives of current **3.1** (`c7b9026dab21`) and **3.2** (`8ea4ebcb2440`) differ
in only three files: `.hg_archival.txt`, `openpyxl/_constants.py`, and `tox.ini`.
The embedded `.hg_archival.txt` confirms each requested node and branch. All
workbook implementation files are byte-for-byte identical; versions are 3.1.6
and 3.2.0b1. Thus these developments are already present on the current 3.1
branch, even though a 3.2 release has not been published. Branch age alone is
not evidence that the work was rejected or never integrated.

Compared with the older published feature set, the development work is significant. Its existing changelog and code include ActiveX and Form
Controls, volatile dependencies and external connections, a Coordinate object
for cells, removal of Cell.internal_value, ISO dates by default and disabling
external-link loading by default. These are compatibility-relevant changes,
not just cosmetic refactoring. **They are already in this fork.** Replacing
this project with today's upstream archive would discard the fork's work.

The delta since this fork's upstream baseline is much smaller:

| Change | Decision |
| --- | --- |
| Close worksheet ZIP streams; close normal-mode archive on success/error (`d53ae9b3cc98`, July 2026) | Ported. Reproduced Windows file lock and archive leak. |
| Case-insensitive MIME lookup (`076cb4cfb132`, September 2025) | Ported; preserve original extension in the package manifest. |
| Unmask XML tests and pass bytes to lxml (`a8da54586190`, September 2025) | Adapted; assert entities are not expanded rather than merely asserting parsing succeeds. |
| Sphinx source inclusion, authors, release notes, copyright | Not copied: publishing/attribution maintenance, no missing workbook behavior. Sources attributed here. |
| Remove Python-version classifiers, add Python 3.14 to tox | Not copied: these do not establish actual compatibility and the Python 3.14 matrix was not run. |

Other open topic branches listed by Heptapod are not merged 3.2 changes and were
not silently imported. Neither #2024 nor #1273 is fixed by the current 3.2 delta.

## Findings in the fork

- Formatting/import rewrites dominate the diff. Functional changes include
  merged-range lookup, NamedStyle/name comparison, merged-cell style propagation,
  removal of the `open` alias, and rebinding named styles before serialization.
  These are retained.
- `MultiCellRange.__getitem__` returns a coordinate **string**, despite the
  README calling it a CellRange. Its existing callers depend on that string.
- Border propagation currently copies the entire border to every cell, including
  interior edges. Loading replaces non-anchor cells and loses their style arrays.
  This is the follow-up scope of `codex/fix-2024-merged-borders`.
- Row/column editing moves cells but leaves reference-bearing objects in place.
  `_move_cells` also materializes the rectangular occupied area before moving,
  which can be costly for sparse sheets. #1273 needs a shared transformation
  model and explicit unsupported-case handling; see its separate proposal branch.
- Restored the ElementTree `iterparse` fallback in `xml.functions`: the fork had
  an export only when optional defusedxml was installed. Tests now use the selected
  parser rather than accidentally testing the unprotected stdlib import.
- The recent NamedStyle rebinding fix was preserved. It does not establish
  thread-safety for sharing mutable NamedStyle objects across workbooks.
- There is an additional pre-existing Coordinate migration inconsistency:
  `Worksheet.append([Cell(...)])` assigns a tuple to `_coord`. This needs a
  separate regression/fix; it is unrelated to the selected upstream delta.

## Validation

Original environment: Windows, Python 3.12.10, pytest 9.1.1, lxml 6.1.1
(outside requirements.txt's lxml < 6 range). Initial full run stopped at three
failures: 2689 passed, 6 skipped, 12 xfailed. Failures: Windows file lock and two
invocations of the invalid lxml test. A focused regression run before porting
failed 7 cases and passed the read-only ownership check.

After porting, the reader/manifest/XML selection passes: 74 passed, 8 skipped,
1 xfailed on that environment. Full-suite validation on the original environment: **2704 passed, 10 skipped,
12 xfailed**, no failures (42.15 seconds). Separate backend validation is recorded
with the final branch results. The tests for file handles
use pytest temporary directories and disable cyclic GC to prevent an accidental
collection from hiding the Windows regression.

Read-only archive ownership remains with Workbook.close() on successful loads.
The upstream patch does not cover read-only failures before returning a workbook;
that is a remaining resource-lifecycle limitation, not a guarantee of this port.


## Branch layout and validation at the end of the initial review

Local branches are stacked to preserve the fork and make each task reviewable:

```text
master (f35829b05, unchanged)
  -> codex/upstream-3.2-fixes
     -> codex/fix-2024-merged-borders
        -> codex/proposal-1273-structural-edits
```

The last branch includes the completed earlier work. Its new #1273 content is a
proposal, pure coordinate prototype and 48 strict expected-failure acceptance
cases; it does not change the worksheet editing behavior. See
[the proposal](proposals/issue1273/README.md) and [#2024 results](issue-2024.md).

Final full-suite result on Windows/Python 3.12.10/lxml 5.4.0: **2749 passed,
10 skipped, 60 xfailed**. Of the 60 xfails, 12 pre-existed and 48 document #1273.
The #2024 branch independently passes both lxml and defusedxml configurations.
No branch was pushed and no GitHub PR was opened. Pre-existing `.idea/` files
were not changed. Test environments and raw comparison data live in the ignored
`.venv/codex-audit/` directory; they are not runtime dependencies or committed code.


## Follow-up: requested implementation

The owner subsequently requested integration of the first two changes and an
implementation of #1273 without formula rewriting. Local `master` first advanced
to `ad9d318f6` (upstream fixes plus #2024). The implementation was developed on
`codex/fix-1273-structural-edits`, based on the proposal branch. Its 48 formerly
expected failures are now ordinary passing regression tests. The earlier branch
layout and test counts above describe the initial review, not the current state.

See [#1273 implementation and validation](issue-1273.md) for supported references,
declaration of unchanged formula text and remaining object-type limitations.
Full validation: 2865 passed with lxml, 2869 with defusedxml; 12 pre-existing
expected failures remain. No changes have been pushed to GitHub.

After three successful visual scenarios rendered by Microsoft Excel 16.0, the
owner requested integration of #1273. Local `master` was fast-forwarded through
`b63fe1faa`; all three requested code changes are now included. See the
implementation notes for the visual checks and their scope.
