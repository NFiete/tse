"""Excel tables (ListObjects): lookup, structured-reference resolution and editing.

Tables live in openpyxl (`ws.tables`); this module keeps them consistent with the cells:
header cells are the column names, the totals row mirrors each column's totals function,
and formulas that use structured references are rewritten when tables or columns are renamed.
Edits made inside `Workbook.transaction` record the table state for undo via `_touch_tables`.
"""
from __future__ import annotations

import re
from copy import deepcopy
from dataclasses import dataclass

from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

from ..formula.errors import REF, VALUE, ExcelError, Unsupported
from ..formula.parser import (
    NAMET, SREFT, FormulaSyntaxError, StructRef, rewrite_tokens, sref_text, to_storage, tokenize, translate,
)
from .refs import Rect, col_letter, parse_addr, quote_sheet

# --------------------------------------------------------------------------- geometry


@dataclass(frozen=True)
class Geom:
    r1: int
    c1: int
    r2: int
    c2: int
    hdr: int  # 1 if the table has a header row
    tot: int  # 1 if the table has a totals row

    @property
    def data_r1(self) -> int:
        return self.r1 + self.hdr

    @property
    def data_r2(self) -> int:
        return max(self.r2 - self.tot, self.data_r1)

    @property
    def rect(self) -> Rect:
        return Rect(self.r1, self.c1, self.r2, self.c2)

    @property
    def data_rect(self) -> Rect:
        return Rect(self.data_r1, self.c1, self.data_r2, self.c2)

    def contains(self, r: int, c: int) -> bool:
        return self.r1 <= r <= self.r2 and self.c1 <= c <= self.c2


def geom(table) -> Geom:
    c1, r1, c2, r2 = range_boundaries(table.ref)
    hdr = 1 if table.headerRowCount is None else (1 if table.headerRowCount else 0)
    tot = 1 if table.totalsRowCount else 0
    return Geom(r1, c1, r2, c2, hdr, tot)


def _ref_text(r1, c1, r2, c2) -> str:
    return f"{col_letter(c1)}{r1}:{col_letter(c2)}{r2}"


@dataclass
class TableHit:
    sheet: object
    table: Table
    g: Geom

    @property
    def name(self) -> str:
        return self.table.displayName

    def col_names(self) -> list[str]:
        return [tc.name or "" for tc in self.table.tableColumns]

    def col_index(self, name: str) -> int | None:
        low = name.lower()
        for i, tc in enumerate(self.table.tableColumns):
            if (tc.name or "").lower() == low:
                return i
        return None


class TableIndex:
    """Lookup of every table in a workbook, by name and by position."""

    def __init__(self, book):
        self.by_name: dict[str, TableHit] = {}
        self.by_sheet: dict[object, list[TableHit]] = {}
        for s in book.sheets:
            hits = []
            for t in s.ws.tables.values():
                try:
                    hit = TableHit(s, t, geom(t))
                except (TypeError, ValueError):
                    continue
                hits.append(hit)
                self.by_name[(t.displayName or t.name or "").lower()] = hit
            if hits:
                self.by_sheet[s] = hits

    def at(self, sheet, r: int, c: int) -> TableHit | None:
        for hit in self.by_sheet.get(sheet, ()):
            if hit.g.contains(r, c):
                return hit
        return None


def tables_of(book) -> list[TableHit]:
    return [h for hits in book.table_index().by_sheet.values() for h in hits]


def table_at(book, sheet, r: int, c: int) -> TableHit | None:
    return book.table_index().at(sheet, r, c)


def find_table(book, name: str) -> TableHit | None:
    return book.table_index().by_name.get(name.lower())


# --------------------------------------------------------------------------- structured references


def resolve(book, sr: StructRef, sheet, r: int, c: int):
    """(sheet, r1, c1, r2, c2) a structured reference points at, or an ExcelError."""
    idx = book.table_index()
    if sr.table is None:
        hit = idx.at(sheet, r, c)
    else:
        hit = idx.by_name.get(sr.table.lower())
    if hit is None:
        return REF
    g = hit.g
    if sr.col1 is None:
        c1, c2 = g.c1, g.c2
    else:
        i = hit.col_index(sr.col1)
        j = i if sr.col2 is None else hit.col_index(sr.col2)
        if i is None or j is None:
            return REF
        c1, c2 = g.c1 + min(i, j), g.c1 + max(i, j)
        if c2 > g.c2:
            return REF
    items = sr.items
    if not items or items == {"data"}:
        r1, r2 = g.data_r1, g.data_r2
    elif "thisrow" in items:
        if hit.sheet is not sheet or not g.data_r1 <= r <= g.data_r2:
            return VALUE
        r1 = r2 = r
    elif "all" in items:
        r1, r2 = g.r1, g.r2
    elif items == {"headers"}:
        if not g.hdr:
            return REF
        r1 = r2 = g.r1
    elif items == {"totals"}:
        if not g.tot:
            return REF
        r1 = r2 = g.r2
    elif items == {"headers", "data"}:
        r1, r2 = g.r1, g.data_r2
    elif items == {"data", "totals"}:
        r1, r2 = g.data_r1, g.r2
    else:
        return REF
    return hit.sheet, r1, c1, r2, c2


def _formula_cells(book, needle: str | None = None):
    """(sheet, r, c, cell) of formula cells whose storage text might mention needle."""
    low = needle.lower() if needle else None
    for s in book.sheets:
        for (r, c), cell in list(s.ws._cells.items()):
            v = cell._value
            if cell.data_type != "f" or not isinstance(v, str):
                continue
            if low is not None and low not in v.lower():
                continue
            yield s, r, c, cell


def qualify(book, text: str, sheet, r: int, c: int) -> str:
    """Storage form of a formula typed at (r, c): unqualified references such as [@Qty]
    get the name of the table they're in (as Excel saves them). Raises FormulaSyntaxError
    when an unqualified reference is used outside any table."""
    if "[" not in text:
        return text
    try:
        toks = tokenize(text[1:])
    except (FormulaSyntaxError, Unsupported):
        return text
    if any(t.kind == SREFT and t.value.table is None for t in toks):
        hit = table_at(book, sheet, r, c)
        if hit is None and book._tx_expand:
            hit = _adjacent_table(book, sheet, r, c)  # typing next to a table extends it
        if hit is None:
            raise FormulaSyntaxError("References like [@Column] need a table name outside a table, "
                                     "e.g. Table1[@Column]")

        def fn(t):
            if t.kind == SREFT and t.value.table is None:
                v = t.value
                return sref_text(StructRef(hit.name, v.items, v.col1, v.col2), storage=True)
            return None

        text = rewrite_tokens(text, fn)
    return to_storage(text)


def _adjacent_table(book, sheet, r: int, c: int) -> TableHit | None:
    for hit in book.table_index().by_sheet.get(sheet, ()):
        g = hit.g
        if (c == g.c2 + 1 and g.r1 <= r <= g.r2) or (not g.tot and r == g.r2 + 1 and g.c1 <= c <= g.c2):
            return hit
    return None


def display_in_place(book, display: str, sheet, r: int, c: int) -> str:
    """Show Table1[@Qty] as [@Qty] inside Table1 (as Excel does)."""
    hit = table_at(book, sheet, r, c)
    if hit is None:
        return display
    low = hit.name.lower()

    def fn(t):
        if t.kind == SREFT and "thisrow" in t.value.items and (t.value.table or "").lower() == low:
            return sref_text(t.value, table=False)
        return None

    return rewrite_tokens(display, fn)


def _rewrite_srefs(text: str, fn_sref, fn_name=None) -> str:
    """Rewrite structured references (and bare table names) in formula text."""

    def fn(t):
        if t.kind == SREFT:
            return fn_sref(t.value, t)
        if fn_name is not None and t.kind == NAMET and t.sheet is None:
            return fn_name(t.value, t)
        return None

    return rewrite_tokens(text, fn)


def _cell_table(book, s, r, c) -> str | None:
    hit = table_at(book, s, r, c)
    return hit.name.lower() if hit else None


def _rewrite_everywhere(book, make_fn, needle: str):
    """Apply a structured-reference rewriter to every formula cell (via set_raw, so it's undoable),
    to the calculated/totals formulas of every table, and to defined names.

    make_fn(table_of_formula_or_None) -> (fn_sref, fn_name)."""
    for s, r, c, cell in _formula_cells(book, needle):
        fs, fnm = make_fn(_cell_table(book, s, r, c))
        new = _rewrite_srefs(cell._value, fs, fnm)
        if new != cell._value:
            book.set_raw(s, r, c, new, keep_cached=True)
    for hit in tables_of(book):
        touched = False
        fs, fnm = make_fn(hit.name.lower())
        for tc in hit.table.tableColumns:
            for attr in ("calculatedColumnFormula", "totalsRowFormula"):
                tf = getattr(tc, attr)
                if tf is None or not tf.attr_text or needle.lower() not in tf.attr_text.lower():
                    continue
                new = _rewrite_srefs("=" + tf.attr_text, fs, fnm)[1:]
                if new != tf.attr_text:
                    if not touched:
                        book._touch_tables(hit.sheet)
                        touched = True
                    tf.attr_text = new
    stores = [book.wb.defined_names] + [s.ws.defined_names for s in book.sheets]
    fs, fnm = make_fn(None)
    for store in stores:
        for dn in store.values():
            if dn.attr_text and needle.lower() in dn.attr_text.lower():
                new = _rewrite_srefs("=" + dn.attr_text, fs, fnm)[1:]
                if new != dn.attr_text:
                    book._touch_names()
                    dn.attr_text = new


def _rename_column_refs(book, table_name: str, old: str, new: str):
    low_t, low_c = table_name.lower(), old.lower()

    def make_fn(home):
        def fs(sr: StructRef, tok):
            target = (sr.table or home or "").lower()
            if target != low_t:
                return None
            c1 = new if sr.col1 is not None and sr.col1.lower() == low_c else sr.col1
            c2 = new if sr.col2 is not None and sr.col2.lower() == low_c else sr.col2
            if (c1, c2) == (sr.col1, sr.col2):
                return None
            return sref_text(StructRef(sr.table, sr.items, c1, c2), storage=True)
        return fs, None

    _rewrite_everywhere(book, make_fn, old)


def _rename_table_refs(book, old: str, new: str):
    low = old.lower()

    def make_fn(home):
        def fs(sr: StructRef, tok):
            if sr.table is None or sr.table.lower() != low:
                return None
            return sref_text(StructRef(new, sr.items, sr.col1, sr.col2), storage=True)

        def fnm(name, tok):
            return new if name.lower() == low else None
        return fs, fnm

    _rewrite_everywhere(book, make_fn, old)


# --------------------------------------------------------------------------- snapshots (undo)


def snapshot(sheet) -> list:
    return [deepcopy(t) for t in sheet.ws.tables.values()]


def restore(book, sheet, snap: list):
    tl = sheet.ws.tables
    tl.clear()
    for t in snap:
        t = deepcopy(t)
        tl[t.name] = t
    book.invalidate_tables()


# --------------------------------------------------------------------------- validation helpers

_TABLE_NAME_RE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.\\]*$")


def validate_table_name(book, name: str, exclude=None) -> str | None:
    from . import names

    if not name:
        return "Enter a table name"
    if len(name) > 255:
        return "Table names can be at most 255 characters"
    if " " in name:
        return "Table names can't contain spaces (use _ instead)"
    if not _TABLE_NAME_RE.match(name):
        return "Table names must start with a letter or _ and contain only letters, digits, _ and ."
    if parse_addr(name) is not None or re.fullmatch(r"[A-Za-z]{1,3}\d+", name) or names._R1C1_RE.match(name):
        return "A table name can't look like a cell reference"
    if name.upper() in ("TRUE", "FALSE"):
        return "TRUE and FALSE can't be used as names"
    other = find_table(book, name)
    if other is not None and other.table is not exclude:
        return f"There's already a table named '{other.name}'"
    if names.find(book, name, None) is not None or any(names.find(book, name, s) is not None for s in book.sheets):
        return f"'{name}' is already a defined name"
    return None


def default_name(book) -> str:
    i = 1
    while find_table(book, f"Table{i}") is not None or validate_table_name(book, f"Table{i}"):
        i += 1
    return f"Table{i}"


def _unique(name: str, taken: set) -> str:
    if name.lower() not in taken:
        return name
    i = 2
    while f"{name}{i}".lower() in taken:
        i += 1
    return f"{name}{i}"


def _fresh_column_name(taken: set) -> str:
    i = 1
    while f"Column{i}".lower() in taken:
        i += 1
    return f"Column{i}"


def _header_text(book, sheet, r, c) -> str:
    """Text a header cell holds (header cells are always text in Excel)."""
    from ..format.numfmt import format_value

    v = book.get_value(sheet, r, c)
    if v is None or isinstance(v, ExcelError):
        return "" if v is None else v.code
    if isinstance(v, str):
        return v
    return format_value(v, sheet.number_format(r, c), book.epoch1904)[0].strip()


def _set_header(book, sheet, r, c, text: str):
    cell = sheet.ws._cells.get((r, c))
    if cell is not None and cell.data_type == "s" and cell._value == text:
        return
    book.set_raw(sheet, r, c, text, "s")


def _overlaps_other(book, sheet, rect: Rect, exclude=None) -> TableHit | None:
    for hit in book.table_index().by_sheet.get(sheet, ()):
        if hit.table is exclude:
            continue
        g = hit.g
        if not (rect.r2 < g.r1 or rect.r1 > g.r2 or rect.c2 < g.c1 or rect.c1 > g.c2):
            return hit
    return None


def _check_merged(sheet, rect: Rect):
    for mr in sheet.ws.merged_cells.ranges:
        if not (rect.r2 < mr.min_row or rect.r1 > mr.max_row or rect.c2 < mr.min_col or rect.c1 > mr.max_col):
            raise ValueError("A table can't contain merged cells; unmerge them first")


def set_ref(table, r1, c1, r2, c2):
    """Change a table's range, keeping its auto-filter in step."""
    table.ref = _ref_text(r1, c1, r2, c2)
    if table.autoFilter is not None:
        tot = 1 if table.totalsRowCount else 0
        table.autoFilter.ref = _ref_text(r1, c1, r2 - tot, c2)
    table.sortState = None


def _next_col_id(table) -> int:
    return max((tc.id or 0 for tc in table.tableColumns), default=0) + 1


# --------------------------------------------------------------------------- create / edit / remove


def create(book, sheet, rect: Rect, has_headers: bool = True, name: str | None = None) -> Table:
    """Turn a range into a table. Without headers, the row above gets Column1, Column2, ..."""
    name = (name or "").strip() or default_name(book)
    err = validate_table_name(book, name)
    if err:
        raise ValueError(err)
    rect = Rect(rect.r1, rect.c1, rect.r2, rect.c2)
    if not has_headers:
        if rect.r1 == 1:
            raise ValueError("There's no row above the range for the headers; insert a row first")
        above = Rect(rect.r1 - 1, rect.c1, rect.r1 - 1, rect.c2)
        if any(book.get_value(sheet, above.r1, c) not in (None, "") for c in range(rect.c1, rect.c2 + 1)):
            raise ValueError("The row above the range must be empty so the headers can go there")
        rect = Rect(rect.r1 - 1, rect.c1, rect.r2, rect.c2)
    if rect.rows < 2:
        raise ValueError("A table needs a header row and at least one row of data")
    if rect.c2 - rect.c1 + 1 > 16384:
        raise ValueError("That range is too wide")
    other = _overlaps_other(book, sheet, rect)
    if other is not None:
        raise ValueError(f"That range overlaps table {other.name}")
    _check_merged(sheet, rect)
    with book.transaction("Create table", (sheet.name, rect)):
        book._touch_tables(sheet)
        cols, taken = [], set()
        for i, c in enumerate(range(rect.c1, rect.c2 + 1), 1):
            text = _header_text(book, sheet, rect.r1, c) if has_headers else ""
            text = _unique(text, taken) if text else _fresh_column_name(taken)
            taken.add(text.lower())
            _set_header(book, sheet, rect.r1, c, text)
            cols.append(TableColumn(id=i, name=text))
        ref = rect_text(rect)
        t = Table(displayName=name, name=name, ref=ref, tableColumns=cols,
                  autoFilter=AutoFilter(ref=ref),
                  tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                                showLastColumn=False, showRowStripes=True,
                                                showColumnStripes=False))
        t.id = 1 + max((h.table.id or 0 for h in tables_of(book)), default=0)
        sheet.ws.tables[name] = t
        book.invalidate_tables()
    return t


def rect_text(rect: Rect) -> str:
    return _ref_text(rect.r1, rect.c1, rect.r2, rect.c2)


def rename(book, hit: TableHit, new: str):
    new = new.strip()
    old = hit.name
    if new == old:
        return
    err = validate_table_name(book, new, exclude=hit.table)
    if err:
        raise ValueError(err)
    with book.transaction("Rename table", (hit.sheet.name, hit.g.rect)):
        book._touch_tables(hit.sheet)
        _rename_table_refs(book, old, new)
        tl = hit.sheet.ws.tables
        for key in [k for k, t in dict(tl).items() if t is hit.table]:
            del tl[key]
        hit.table.displayName = new
        hit.table.name = new
        tl[new] = hit.table
        book.invalidate_tables()


def _totals_formula(table_name: str, col: str, func: str) -> str | None:
    code = SUBTOTAL_CODES.get(func)
    if code is None:
        return None
    return f"=SUBTOTAL({code},{sref_text(StructRef(table_name, frozenset(), col))})"


SUBTOTAL_CODES = {"average": 101, "countNums": 102, "count": 103, "max": 104, "min": 105,
                  "stdDev": 107, "sum": 109, "var": 110}
_CODE_FUNCS = {v: k for k, v in SUBTOTAL_CODES.items()}
_CODE_FUNCS.update({v - 100: k for k, v in SUBTOTAL_CODES.items()})


def set_totals(book, hit: TableHit, on: bool):
    """Show or hide the totals row (it goes in the row just below the data)."""
    t, g, sheet = hit.table, hit.g, hit.sheet
    if bool(g.tot) == on:
        return
    with book.transaction("Totals row", (sheet.name, g.rect)):
        book._touch_tables(sheet)
        if on:
            row = g.r2 + 1
            from .refs import MAX_ROW

            if row > MAX_ROW:
                raise ValueError("There's no room below the table for a totals row")
            if any(book.get_value(sheet, row, c) not in (None, "") for c in range(g.c1, g.c2 + 1)) or \
                    _overlaps_other(book, sheet, Rect(row, g.c1, row, g.c2), exclude=t):
                raise ValueError("The row below the table must be empty to add a totals row")
            cols = t.tableColumns
            if not any(tc.totalsRowFunction or tc.totalsRowLabel for tc in cols):
                if len(cols) > 1:
                    cols[0].totalsRowLabel = "Total"
                last = cols[-1]
                numeric = any(isinstance(book.get_value(sheet, r, g.c2), (int, float))
                              and not isinstance(book.get_value(sheet, r, g.c2), bool)
                              for r in range(g.data_r1, g.data_r2 + 1))
                last.totalsRowFunction = "sum" if numeric else "count"
            t.totalsRowCount = 1
            t.totalsRowShown = True
            set_ref(t, g.r1, g.c1, row, g.c2)
            for i, tc in enumerate(cols):
                c = g.c1 + i
                if tc.totalsRowFunction == "custom" and tc.totalsRowFormula is not None:
                    book.set_raw(sheet, row, c, "=" + tc.totalsRowFormula.attr_text)
                elif tc.totalsRowFunction and tc.totalsRowFunction != "none":
                    f = _totals_formula(t.displayName, tc.name, tc.totalsRowFunction)
                    if f:
                        book.set_raw(sheet, row, c, f)
                elif tc.totalsRowLabel:
                    book.set_raw(sheet, row, c, tc.totalsRowLabel, "s")
        else:
            row = g.r2
            for c in range(g.c1, g.c2 + 1):
                book.set_raw(sheet, row, c, None)
            t.totalsRowCount = None
            set_ref(t, g.r1, g.c1, row - 1, g.c2)
        book.invalidate_tables()


def resize(book, hit: TableHit, rect: Rect):
    """Change a table's range. The header row must stay where it is."""
    t, g, sheet = hit.table, hit.g, hit.sheet
    if rect == g.rect:
        return
    if rect.r1 != g.r1:
        raise ValueError(f"The header row has to stay in row {g.r1}")
    if rect.rows < 1 + g.hdr + g.tot:
        raise ValueError("A table needs at least one row of data")
    other = _overlaps_other(book, sheet, rect, exclude=t)
    if other is not None:
        raise ValueError(f"That range overlaps table {other.name}")
    _check_merged(sheet, rect)
    totals = bool(g.tot)
    if totals:
        # the range includes the totals row, which moves to the new last row
        for c in range(rect.c1, rect.c2 + 1):
            if not (rect.r2 == g.r2 and g.c1 <= c <= g.c2) and book.get_value(sheet, rect.r2, c) not in (None, ""):
                raise ValueError(f"Row {rect.r2} must be empty: the totals row moves there")
        rect = Rect(rect.r1, rect.c1, rect.r2 - 1, rect.c2)
    with book.transaction("Resize table", (sheet.name, rect)):
        if totals:
            set_totals(book, hit, False)
            hit = find_table(book, t.displayName)
        book._touch_tables(sheet)
        old_cols = {g.c1 + i: tc for i, tc in enumerate(t.tableColumns)}
        new_cols, taken = [], set()
        next_id = _next_col_id(t)
        for c in range(rect.c1, rect.c2 + 1):
            tc = old_cols.get(c)
            if tc is not None:
                new_cols.append(tc)
                taken.add((tc.name or "").lower())
        out = []
        for c in range(rect.c1, rect.c2 + 1):
            tc = old_cols.get(c)
            if tc is None:
                text = _header_text(book, sheet, rect.r1, c) if g.hdr else ""
                text = _unique(text, taken) if text else _fresh_column_name(taken)
                taken.add(text.lower())
                if g.hdr:
                    _set_header(book, sheet, rect.r1, c, text)
                tc = TableColumn(id=next_id, name=text)
                next_id += 1
            out.append(tc)
        if [id(x) for x in out] != [id(x) for x in t.tableColumns] and t.autoFilter is not None:
            t.autoFilter.filterColumn = []
        t.tableColumns = out
        set_ref(t, rect.r1, rect.c1, rect.r2, rect.c2)
        book.invalidate_tables()
        if totals:
            set_totals(book, find_table(book, t.displayName), True)


def _a1_for(book, sr: StructRef, home_sheet, r, c, formula_sheet) -> str:
    res = resolve(book, sr, home_sheet, r, c)
    if isinstance(res, ExcelError):
        return "#REF!"
    s, r1, c1, r2, c2 = res
    prefix = "" if s is formula_sheet else quote_sheet(s.name) + "!"
    if "thisrow" in sr.items:
        a, b = f"${col_letter(c1)}{r1}", f"${col_letter(c2)}{r2}"
    else:
        a, b = f"${col_letter(c1)}${r1}", f"${col_letter(c2)}${r2}"
    return prefix + (a if a == b else f"{a}:{b}")


def convert_to_range(book, hit: TableHit):
    """Remove the table but keep its cells; structured references become A1 references."""
    sheet, t = hit.sheet, hit.table
    low = hit.name.lower()
    with book.transaction("Convert to range", (sheet.name, hit.g.rect)):
        for s, r, c, cell in _formula_cells(book):
            home = _cell_table(book, s, r, c)

            def fs(sr, tok, s=s, r=r, c=c, home=home):
                if (sr.table or home or "").lower() != low:
                    return None
                return _a1_for(book, sr, s, r, c, s)

            def fnm(name, tok, s=s):
                if name.lower() != low or book.resolve_name(name, s, None, tables=False) is not None:
                    return None
                return _a1_for(book, StructRef(hit.name), s, 1, 1, s)

            new = _rewrite_srefs(cell._value, fs, fnm)
            if new != cell._value:
                book.set_raw(s, r, c, new, keep_cached=True)
        for store in [book.wb.defined_names] + [s.ws.defined_names for s in book.sheets]:
            for dn in store.values():
                if dn.attr_text and low in dn.attr_text.lower():
                    new = _rewrite_srefs("=" + dn.attr_text,
                                         lambda sr, tok: None if (sr.table or "").lower() != low else
                                         _a1_for(book, sr, sheet, 1, 1, None))[1:]
                    if new != dn.attr_text:
                        book._touch_names()
                        dn.attr_text = new
        book._touch_tables(sheet)
        tl = sheet.ws.tables
        for key in [k for k, x in dict(tl).items() if x is t]:
            del tl[key]
        book.invalidate_tables()


def delete_table(book, hit: TableHit):
    """Delete a table and clear its cells."""
    with book.transaction("Delete table", (hit.sheet.name, hit.g.rect)):
        convert_to_range(book, hit)
        g = hit.g
        for r in range(g.r1, g.r2 + 1):
            for c in range(g.c1, g.c2 + 1):
                cell = hit.sheet.ws._cells.get((r, c))
                if cell is not None and not isinstance(cell, MergedCell) and cell._value is not None:
                    book.set_raw(hit.sheet, r, c, None)


# --------------------------------------------------------------------------- keeping tables in step with edits


def _fill_calculated(book, hit: TableHit, rows):
    """Write each calculated column's formula into the given (empty) data rows."""
    g = hit.g
    for i, tc in enumerate(hit.table.tableColumns):
        tf = tc.calculatedColumnFormula
        if tf is None or not tf.attr_text:
            continue
        c = g.c1 + i
        for r in rows:
            cell = hit.sheet.ws._cells.get((r, c))
            if cell is not None and cell._value is not None:
                continue
            book.set_raw(hit.sheet, r, c, translate("=" + tf.attr_text, r - g.data_r1, 0))


def _expand(book, keys_by_sheet: dict):
    """Typing just below or right of a table makes the table grow (Excel's auto-expansion)."""
    for sheet, rcs in keys_by_sheet.items():
        for hit in list(book.table_index().by_sheet.get(sheet, ())):
            g = hit.g
            t = hit.table
            filled = {rc for rc in rcs if sheet.ws._cells.get(rc) is not None
                      and sheet.ws._cells[rc]._value is not None}
            below = not g.tot and any(r == g.r2 + 1 and g.c1 <= c <= g.c2 for r, c in filled)
            right = any(c == g.c2 + 1 and g.r1 <= r <= g.r2 for r, c in filled)
            if not below and not right:
                continue
            new = Rect(g.r1, g.c1, g.r2 + (1 if below else 0), g.c2 + (1 if right else 0))
            if _overlaps_other(book, sheet, new, exclude=t) is not None:
                continue
            if any(not (mr.max_row < new.r1 or mr.min_row > new.r2 or mr.max_col < new.c1 or mr.min_col > new.c2)
                   for mr in sheet.ws.merged_cells.ranges):
                continue
            book._touch_tables(sheet)
            if right:
                taken = {(tc.name or "").lower() for tc in t.tableColumns}
                text = _header_text(book, sheet, g.r1, new.c2) if g.hdr else ""
                text = _unique(text, taken) if text else _fresh_column_name(taken)
                if g.hdr:
                    _set_header(book, sheet, g.r1, new.c2, text)
                t.tableColumns.append(TableColumn(id=_next_col_id(t), name=text))
                if g.tot:
                    # the totals row grows too; move nothing, it's just one more (empty) cell
                    pass
            set_ref(t, new.r1, new.c1, new.r2, new.c2)
            book.invalidate_tables()
            if below:
                _fill_calculated(book, find_table(book, t.displayName), [new.r2])


def _calculated_column(book, hit: TableHit, r0: int, c0: int):
    """A formula typed into an otherwise empty table column fills the whole column."""
    g = hit.g
    cell = hit.sheet.ws._cells.get((r0, c0))
    if cell is None or cell.data_type != "f" or not isinstance(cell._value, str):
        return
    text = cell._value
    tc = hit.table.tableColumns[c0 - g.c1]
    prev = tc.calculatedColumnFormula.attr_text if tc.calculatedColumnFormula is not None else None
    others = [r for r in range(g.data_r1, g.data_r2 + 1) if r != r0]
    if not others:
        return
    cells = hit.sheet.ws._cells

    def consistent(r):
        other = cells.get((r, c0))
        if other is None or other._value is None:
            return True
        return prev is not None and other._value == translate("=" + prev, r - g.data_r1, 0)

    if not all(consistent(r) for r in others):
        return
    for r in others:
        book.set_raw(hit.sheet, r, c0, translate(text, r - r0, 0))
    book._touch_tables(hit.sheet)
    tc.calculatedColumnFormula = TableFormula(attr_text=translate(text, g.data_r1 - r0, 0)[1:])


def _sync_header(book, hit: TableHit, c: int):
    sheet, g, t = hit.sheet, hit.g, hit.table
    i = c - g.c1
    tc = t.tableColumns[i]
    cell = sheet.ws._cells.get((g.r1, c))
    text = _header_text(book, sheet, g.r1, c)
    taken = {(x.name or "").lower() for j, x in enumerate(t.tableColumns) if j != i}
    text = _unique(text, taken) if text else _fresh_column_name(taken)
    if cell is None or cell.data_type != "s" or cell._value != text:
        _set_header(book, sheet, g.r1, c, text)
    if tc.name != text:
        old = tc.name or ""
        book._touch_tables(sheet)
        tc.name = text
        if old:
            _rename_column_refs(book, hit.name, old, text)
        book.invalidate_tables()


_SUBTOTAL_RE = re.compile(r"^=\s*SUBTOTAL\(\s*(\d+)\s*,\s*(.+?)\s*\)$", re.I)


def _sync_totals(book, hit: TableHit, c: int):
    sheet, g, t = hit.sheet, hit.g, hit.table
    tc = t.tableColumns[c - g.c1]
    cell = sheet.ws._cells.get((g.r2, c))
    func = label = formula = None
    if cell is not None and cell._value is not None:
        v = cell._value
        if cell.data_type == "f" and isinstance(v, str):
            m = _SUBTOTAL_RE.match(v)
            fn = _CODE_FUNCS.get(int(m.group(1))) if m else None
            if fn is not None and m.group(2).lower() == sref_text(StructRef(hit.name, frozenset(), tc.name)).lower():
                func = fn
            else:
                func, formula = "custom", TableFormula(attr_text=v[1:])
        else:
            label = str(v)
    state = (tc.totalsRowFunction, tc.totalsRowLabel,
             tc.totalsRowFormula.attr_text if tc.totalsRowFormula is not None else None)
    if state != (func, label, formula.attr_text if formula is not None else None):
        book._touch_tables(sheet)
        tc.totalsRowFunction, tc.totalsRowLabel, tc.totalsRowFormula = func, label, formula


def after_edit(book, keys, expand: bool = False):
    """Called at the end of a transaction (still inside it) with the cells that changed."""
    by_sheet: dict = {}
    idx = book.table_index()
    for s, r, c in keys:
        if s in idx.by_sheet:
            by_sheet.setdefault(s, set()).add((r, c))
    if not by_sheet:
        return
    if expand:
        _expand(book, by_sheet)
        if len(keys) == 1:
            (s, r, c), = keys
            hit = table_at(book, s, r, c)
            if hit is not None and hit.g.data_r1 <= r <= hit.g.data_r2:
                _calculated_column(book, hit, r, c)
    for sheet, rcs in by_sheet.items():
        for hit in list(book.table_index().by_sheet.get(sheet, ())):
            g = hit.g
            if g.hdr:
                for c in sorted({c for r, c in rcs if r == g.r1 and g.c1 <= c <= g.c2}):
                    _sync_header(book, find_table(book, hit.name) or hit, c)
            if g.tot:
                for c in sorted({c for r, c in rcs if r == g.r2 and g.c1 <= c <= g.c2}):
                    _sync_totals(book, find_table(book, hit.name) or hit, c)


# --------------------------------------------------------------------------- rows / columns inserted or deleted


def check_delete(book, sheet, axis: str, idx: int, n: int):
    """Raise ValueError if deleting these rows/columns would break a table apart."""
    end = idx + n - 1
    for hit in book.table_index().by_sheet.get(sheet, ()):
        g = hit.g
        if axis == "row":
            if idx <= g.r1 and end >= g.r2:
                continue  # the whole table goes
            if g.hdr and idx <= g.r1 <= end:
                raise ValueError(f"That would delete the header row of table {hit.name} "
                                 "(delete the whole table, or convert it to a range first)")
            data_left = sum(1 for r in range(g.data_r1, g.r2 - g.tot + 1) if not idx <= r <= end)
            if data_left < 1:
                raise ValueError(f"Table {hit.name} needs at least one row of data")


def shift(book, sheet, axis: str, idx: int, n: int, fill: bool = True) -> list:
    """Move/resize the sheet's tables for inserted (n > 0) or deleted (n < 0) rows/columns.
    Returns (r, c) of calculated-column formulas written into inserted table rows."""
    tl = sheet.ws.tables
    new_formulas: list = []
    if not tl:
        return new_formulas
    end = idx - n - 1 if n < 0 else None
    for key, t in list(dict(tl).items()):
        g = geom(t)
        if axis == "row":
            if n > 0:
                r1 = g.r1 + n if g.r1 >= idx else g.r1
                r2 = g.r2 + n if g.r2 >= idx else g.r2
            else:
                if idx <= g.r1 and end >= g.r2:
                    del tl[key]
                    continue
                gone = sum(1 for r in range(g.r1, g.r2 + 1) if idx <= r <= end)
                r1 = g.r1 if g.r1 < idx else (idx if g.r1 <= end else g.r1 + n)
                r2 = r1 + (g.r2 - g.r1 + 1 - gone) - 1
                if g.tot and idx <= g.r2 <= end:
                    t.totalsRowCount = None
            set_ref(t, r1, g.c1, r2, g.c2)
            if n > 0 and fill and g.r1 < idx <= g.r2:
                new_formulas += _new_rows_calculated(sheet, t, idx, n)
        else:
            if n > 0:
                c1 = g.c1 + n if g.c1 >= idx else g.c1
                c2 = g.c2 + n if g.c2 >= idx else g.c2
                if g.c1 < idx <= g.c2:
                    pos = idx - g.c1
                    taken = {(tc.name or "").lower() for tc in t.tableColumns}
                    nid = _next_col_id(t)
                    new_cols = []
                    for k in range(n):
                        name = _fresh_column_name(taken)
                        taken.add(name.lower())
                        new_cols.append(TableColumn(id=nid + k, name=name))
                        if g.hdr:
                            cell = sheet.ws.cell(row=g.r1, column=idx + k)
                            cell.value = name
                    t.tableColumns = list(t.tableColumns[:pos]) + new_cols + list(t.tableColumns[pos:])
                    if t.autoFilter is not None:
                        t.autoFilter.filterColumn = []
            else:
                if idx <= g.c1 and end >= g.c2:
                    del tl[key]
                    continue
                keep = [tc for i, tc in enumerate(t.tableColumns) if not idx <= g.c1 + i <= end]
                if len(keep) != len(t.tableColumns) and t.autoFilter is not None:
                    t.autoFilter.filterColumn = []
                t.tableColumns = keep
                c1 = g.c1 if g.c1 < idx else (idx if g.c1 <= end else g.c1 + n)
                c2 = c1 + len(keep) - 1
            set_ref(t, g.r1, c1, g.r2, c2)
    book.invalidate_tables()
    return new_formulas


def _new_rows_calculated(sheet, t, idx, n) -> list:
    """Inserted table rows get the table's calculated column formulas. Returns their (r, c)."""
    g = geom(t)
    out = []
    for i, tc in enumerate(t.tableColumns):
        tf = tc.calculatedColumnFormula
        if tf is None or not tf.attr_text:
            continue
        c = g.c1 + i
        for r in range(idx, idx + n):
            if not g.data_r1 <= r <= g.data_r2:
                continue
            cell = sheet.ws.cell(row=r, column=c)
            cell.value = translate("=" + tf.attr_text, r - g.data_r1, 0)
            out.append((r, c))
    return out
