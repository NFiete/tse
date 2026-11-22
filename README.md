# term_xl

A spreadsheet for the terminal that opens, edits and saves real Excel files.

```
./txl budget.xlsx          # open a workbook
./txl new.xlsx             # start a new one (created on first Ctrl+S)
./txl data.csv --sheet X   # CSV/TSV work too
./txl --vi budget.xlsx     # vi-style keys
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
- **Live formulas.** It has its own parser and recalculation engine with over 300 functions
  (math, statistics, probability distributions and hypothesis tests, regression, logical, lookup,
  text, date/time, information), Excel's type coercion and error values, ranges across sheets,
  defined names, array arithmetic (`SUMPRODUCT((A:A>5)*B:B)`), dynamic arrays that spill into
  neighbouring cells (`=SORT(FILTER(A2:C99,C2:C99>0))`, `=A2:A10*2`, spill references like
  `=SUM(E2#)`, `@` for implicit intersection), `LET` and `LAMBDA` (including LAMBDAs stored in
  defined names, and the helpers MAP, REDUCE, SCAN, BYROW, BYCOL and MAKEARRAY), and
  circular-reference detection. Results Excel saved in the file are shown as-is until
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
- **Grouped rows and columns.** Outlines in a file show as bars with `[-]`/`[+]` buttons above
  the column letters and left of the row numbers; click a button to collapse or expand, or the
  level numbers (`1 2`) to show one outline level. Alt+Shift+Right/Left groups/ungroups the
  selection, Alt+Shift+Up/Down collapses/expands the group at the cursor.
- **Named ranges.** Select cells, click the address box (top left), type a name and press Enter.
  Ctrl+F3 opens a Name Manager to create, edit, rename (formulas follow), delete and jump to
  names, with workbook or sheet scope. Names can also be constants or formulas (`Rate = 0.0825`).
- **Tables.** Tables in a file work in formulas: `=SUM(Sales[Amount])`, `Sales[#Totals]`,
  `Sales[[#Headers],[Qty]:[Price]]`, and `[@Qty]*[@Price]` for the same row inside the table.
  Ctrl+T turns the selection (or the data around the cursor) into a table, or, inside a table,
  renames or resizes it, adds a totals row, converts it back to a range or deletes it. As in
  Excel, typing just below or right of a table extends it, a formula typed into an empty column
  fills the whole column, renaming a header updates the formulas that use it, and inserted rows
  get the column formulas. Table styles aren't drawn: cells show their own formatting.

- **Vi keys, if you prefer them.** Start with `--vi`, or pick "Use vi keys" in the command
  palette (remembered in `~/.config/term_xl/settings.json`). `hjkl`, `w`/`b`/`{`/`}`, `gg`/`G`,
  counts, `v`/`V` selection, `i`/`a`/`=` to edit (Esc confirms), `dd`/`x`, `yy`, `p`, `dr`/`yr`
  for whole rows, `u`/`Ctrl+R`, `/` search with `n`/`N`, marks, and `:w`, `:q`, `:wq`, `:e`,
  `:B12`. Ctrl shortcuts vi doesn't use keep their Excel meaning. `:set novi` switches back.

Press **F1** for the key reference and **Ctrl+P** for the command palette.

## Compatibility notes

- Files are read and written with openpyxl. A few things openpyxl can't write back:
  **slicers, timelines, shapes/text boxes, sparklines, form/ActiveX controls, Power Pivot
  models and some newer conditional formats**. If a file contains any of these you're warned
  when it opens, and saving over it asks first and offers to save a copy instead.
  Charts are kept, but some chart formatting may change.
- If a formula uses a function term_xl doesn't implement (for example `BESSELJ`, or references
  to other workbooks), the formula is kept unchanged and the cell keeps the value Excel saved.
  The status bar tells you when you're on such a cell.
- Cutting and pasting a whole table moves its cells but not the table itself. Deleting a table
  column leaves formulas that used it showing `#REF!` without rewriting them.
- Formulas you type use Excel 365 dynamic-array rules: an array result spills, and shows
  `#SPILL!` if something is in the way (the status bar says where). Ordinary formulas in a file
  that predate dynamic arrays keep the older behaviour (a range gives the value in the
  formula's row, as Excel shows with `=@...`), and old Ctrl+Shift+Enter array formulas fill their
  fixed range.
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
