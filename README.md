# term_xl

A spreadsheet for the terminal that opens, edits and saves real Excel files.

```
./txl budget.xlsx          # open a workbook
./txl new.xlsx             # start a new one (created on first Ctrl+S)
./txl data.csv --sheet X   # CSV/TSV work too
```

It needs Python 3.11+ with `textual` and `openpyxl`. On Arch these are the system packages
`python-textual` and `python-openpyxl`; nothing is installed with pip. `lxml` is used if it's
available. Run it from the source tree with `./txl` or `python3 -m termxl`.

## What it does

- **Excel files in and out.** `.xlsx`, `.xlsm` (macros kept), `.xltx/.xltm`, `.csv`, `.tsv`.
  Saving writes your changes into the original workbook object, so formatting, merged cells,
  column widths, row heights, hidden rows/columns/sheets, frozen panes, defined names, data
  validation, conditional formatting, notes, hyperlinks, tab colors and sheet order are kept.
  Formula results are written into the file too, so other programs show values straight away,
  and Excel recalculates when it opens the file.
- **Live formulas.** It has its own parser and recalculation engine with about 190 functions
  (math, statistics, logical, lookup, text, date/time, information), Excel's type coercion and
  error values, ranges across sheets, defined names, array arithmetic (`SUMPRODUCT((A:A>5)*B:B)`),
  and circular-reference detection. Results Excel saved in the file are shown as-is until
  something they depend on changes.
- **Excel-style editing.** Type to replace a cell, F2 to edit it. Arrow keys (or the mouse)
  insert references while you write a formula, and F4 toggles `$`. Excel number formats are
  rendered exactly (`#,##0.00;[Red](#,##0.00)`, dates, times, fractions, percentages…).
  Typing `12%`, `$5.50`, `2026-11-21` or `1:30 PM` gives a typed value with a matching format.
- **Everything else you'd expect:** copy, cut and paste with reference adjustment, fill
  down/right, unlimited undo/redo, insert/delete rows and columns (references update,
  `#REF!` where needed), find and replace with wildcards, go to, sort, bold/italic/underline,
  colors, alignment, column widths, hide/unhide, freeze panes, adding/renaming/moving/deleting
  sheets, a Sum/Average/Count readout for the selection, mouse support, and dark or light grids.

Press **F1** for the key reference and **Ctrl+P** for the command palette.

## Compatibility notes

- Files are read and written with openpyxl. A few things openpyxl can't write back:
  **slicers, timelines, shapes/text boxes, sparklines, form/ActiveX controls, Power Pivot
  models and some newer conditional formats**. If a file contains any of these you're warned
  when it opens, and saving over it asks first and offers to save a copy instead.
  Charts are kept, but some chart formatting may change.
- If a formula uses a function term_xl doesn't implement (for example `BESSELJ`, or structured
  table references like `Table1[Amount]`), the formula is kept unchanged and the cell keeps
  the value Excel saved. The status bar tells you when you're on such a cell.
- Dynamic-array formulas (`FILTER`, `UNIQUE`, `SORT`, `SEQUENCE`…) work inside other functions
  (`=SUM(FILTER(...))`), but results don't spill into neighbouring cells. A cell shows the
  first value.
- Legacy `.xls` and `.xlsb` aren't supported. Convert them first:
  `soffice --headless --convert-to xlsx file.xls`.
- Ctrl+I only reaches the app in terminals with the kitty keyboard protocol; elsewhere use Alt+I.
  To paste text from other programs, use your terminal's paste (often Ctrl+Shift+V).

## Performance

A 50,000-row sheet with 150,000 formulas opens in about 6 seconds. Edits recalculate only what
depends on them, usually in milliseconds. Inserting or deleting rows near the top of a sheet that
size takes several seconds, because every formula below has to be rewritten. Sheets with a few
thousand formulas are effectively instant.

## Layout

```
termxl/
  model/      workbook + sheets, typed input, dates, undo, structural edits
  formula/    lexer/parser, evaluator, function library, dependency graph
  format/     number-format renderer, Excel style -> terminal colors
  io/         xlsx/csv load and save
  ui/         Textual app, grid widget, dialogs
tests/        pytest suite (engine, formats, round-trips, scripted UI sessions,
              cross-check against LibreOffice)
```

## Tests

```
python3 -m pytest
```

`tests/test_libreoffice.py` recalculates about 150 formulas with LibreOffice (`soffice`) and
compares the results with term_xl's. It's skipped if LibreOffice isn't installed.
`tests/fixtures/make_sample.py` builds the demo workbook `tests/fixtures/sample.xlsx`.
