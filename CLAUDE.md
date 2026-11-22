# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

term_xl is a Textual-based terminal spreadsheet that opens, edits and saves real Excel files
(`.xlsx/.xlsm/.xltx/.xltm`, plus CSV/TSV) with its own formula parser and recalculation engine.
See README.md for user-facing features and compatibility limits.

## Environment

- Python 3.11+, dependencies `textual` and `openpyxl` (optional `lxml`). Use the Arch system
  packages (`python-textual`, `python-openpyxl`, `python-pytest`, `python-pytest-asyncio`); do not
  pip-install or create a venv. Ask the user if something is missing.
- There is no build step and no linter configured.

## Commands

```
./txl file.xlsx                      # run from the source tree (or: python3 -m termxl file.xlsx)
python3 -m pytest                    # full suite (~25s)
python3 -m pytest tests/test_functions.py
python3 -m pytest tests/test_ui.py -k test_type_values_and_formulas
python3 tests/fixtures/make_sample.py  # rebuild tests/fixtures/sample.xlsx
```

`tests/test_libreoffice.py` cross-checks ~150 formulas against LibreOffice headless and is skipped
when `soffice` isn't installed. When adding a function, adding a formula to its `FORMULAS` list is
the best way to verify Excel-compatible semantics.

## Architecture

### openpyxl is the source of truth
`model/workbook.py` `Workbook` wraps an openpyxl workbook; each `Sheet` wraps an openpyxl
worksheet. Raw cell contents and styles live in openpyxl objects (accessed directly via
`ws._cells` for speed), so anything term_xl doesn't model (validation, conditional formats,
charts, …) round-trips untouched. term_xl's own state is layered on top per sheet:
`formulas` (`(r,c) -> Formula` with parsed AST), `values` (computed results), `cached`
(Excel's saved results), and `stale` (cells showing a cached value that couldn't be recomputed).
Keys throughout are `(sheet_obj, row, col)` with 1-based rows/cols.

### Formula text has two forms
Storage form (in the file / openpyxl) carries `_xlfn.`/`_xlfn._xlws.` prefixes; display form (what
the user sees/edits, and `Formula.text`) does not. Convert with `parser.to_storage` /
`parser.to_display`. Any code that writes formulas into cells must store the storage form.

### Load → calc → save pipeline (`io/xlsx.py`)
1. `openpyxl.load_workbook` (keeps VBA for macro formats, keeps external links).
2. `Workbook.register_all()` parses every formula and builds the `DepGraph`.
3. `_load_cached_values` regex-scans the sheet XML for Excel's cached `<v>` results (openpyxl can't
   give both formula and value), falling back to a second openpyxl read.
4. `initial_calc()` keeps cached values and only computes formulas with no cache, volatile ones,
   and their dependents. Formulas with unsupported functions keep the cached value (marked stale).
5. `save()` writes via a temp file + `os.replace`, then `_inject_cached_values` patches computed
   results into the saved XML so other programs see values; `fullCalcOnLoad` is set for Excel.

The GC is paused and the heap frozen during bulk work (`model/gcutil.paused_gc`, `xlsx.load`) — a
deliberate performance measure for large workbooks; keep it when touching those paths.

### Formula engine (`formula/`)
- `parser.py`: tokenizer + recursive-descent parser → AST nodes (`Const`, `RefNode`, `NameNode`,
  `Func`, `Binary`, …). Also provides token-level reference rewriting (`transform_refs`,
  `rewrite_tokens`, `translate`) used for copy/paste, structural edits and name renames.
- `evaluator.py`: evaluates ASTs with Excel coercion, error propagation and elementwise array
  lifting. Errors are raised as `XLError` and turned into `ExcelError` values; `Unsupported` means
  "keep the cached value".
- `registry.py`: the `@fn(...)` decorator registers functions in `FUNCTIONS`. Three kinds:
  `scalar` (auto-lifted over arrays/ranges), `raw` (gets `ctx` and unevaluated `RefValue`/`Array`
  args), `lazy` (gets thunks; for IF/IFERROR-style short-circuiting). Shared helpers (`numbers`,
  `flatten`, `make_criterion`, wildcard matching) live here.
- `functions/*_fns.py`: implementations, all imported by `functions/__init__.py` (a new module
  must be added there).
- `deps.py` `DepGraph`: single-cell deps, narrow ranges indexed per column, wide ranges kept per
  sheet; volatile formulas are always recalculated. `Workbook.recalc(changed)` dirties dependents
  and evaluates in topological order.
- `values.py`: runtime types (`RefValue`, `Array`, `MISSING`) and coercion/comparison rules.

### Edits and undo
- All UI cell edits go through `Workbook` methods (`set_input`, `set_style`, `clear`, …) inside
  `with book.transaction(label, where):`. Writes call `_touch` to snapshot the cell before change;
  on exit the transaction pushes one `UndoEntry` (before/after snapshots) and recalcs once.
- `model/structure.py` (insert/delete rows/cols, add/rename/remove sheets) and `model/names.py`
  (defined names) do their own bookkeeping: rewrite formulas across all sheets and names, shift
  openpyxl dimensions/merges/validations/etc., re-register formulas, and push custom
  `UndoEntry(undo=..., redo=...)` closures. Functions take `record=True` to control undo recording.
- `model/outline.py` handles grouped rows/columns. Outline state lives only in openpyxl row/column
  dimensions (`outlineLevel`, `hidden`, `collapsed` on the summary row/col); groups are derived on
  demand. Column dims are read through `Sheet._build_colw` (cached, reset by
  `invalidate_layout()`, which also bumps `Sheet.outline_gen`); write them via
  `structure.isolate_cols` so ranged dims like `B:XFD` are split rather than exploded.
- `model/tables.py` handles Excel tables (openpyxl `ws.tables`). Structured references parse to
  `StructRefNode` (`parser.StructRef`) and resolve through `Workbook.table_index()` /
  `tables.resolve`. Stored formulas always name the table and spell `[@Col]` as
  `[[#This Row],[Col]]` (`tables.qualify` in `set_raw`; display shows `[@Col]` inside the table).
  Inside a transaction, call `book._touch_tables(sheet)` / `book._touch_names()` before changing
  table or name metadata so undo restores it; `tables.after_edit` runs at the end of every
  transaction to sync header names, totals-row functions, auto-expansion and calculated columns.
  Changing tables needs `invalidate_tables()`; formulas with `structured`/`uses_names` are
  re-registered by `tables_changed()`.
- `model/inputparse.py` turns typed text (`12%`, `$5.50`, dates, times) into typed values + number
  format.

### UI (`ui/`)
- `app.py` `TermXLApp`: key bindings, edit mode (including "point mode" where arrows insert
  references into a formula), clipboard, dialogs, status bar; calls into model functions.
- `grid.py` `Grid`: virtualized line-by-line rendering (`render_line`), selection, frozen panes.
  When a sheet has outlines, `Layout.co_h` band lines sit above the column header and a
  `Layout.ro_w` gutter left of the row numbers, so screen y/x don't map 1:1 to the header/cells;
  use `hit()` and `Layout.cells_x`.
- `format/numfmt.py` renders Excel number formats; `format/style.py` maps cell styles to terminal
  colors for the dark/light palettes in `ui/theme.py`.

## Tests

`tests/conftest.py` provides helpers: `ws` (a `Sheetish` wrapper over a new workbook:
`ws["A1"] = "=1+1"; ws["A1"] -> 2`), `ev` (evaluate one formula with optional cell setup), and
`sample_path`. UI tests drive `TermXLApp` with Textual's `run_test()`/Pilot (`tests/test_ui.py`
has `_ready`, `type_text`, `screen_text` helpers). Round-trip tests save and reload through
openpyxl to verify preservation.
