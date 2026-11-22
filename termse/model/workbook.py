"""Workbook model: wraps an openpyxl Workbook and adds formulas, recalculation and undo.

The openpyxl workbook stays the source of truth for raw cell contents and styles,
so everything we do not touch (styles, merged cells, validations, charts...) is
written back unchanged when saving.
"""
from __future__ import annotations

import datetime as dt
import sys
from contextlib import contextmanager
from copy import copy

from openpyxl.cell.cell import MergedCell
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils.exceptions import IllegalCharacterError
from openpyxl.utils.datetime import MAC_EPOCH
from openpyxl.worksheet.formula import ArrayFormula

from ..formula import evaluator as ev
from ..formula.deps import DepGraph
from ..formula.errors import CIRC, ERROR_CODES, NAME, REF, SPILL, VALUE, ExcelError, Unsupported, XLError
from ..formula.functions import date_fns
from ..formula.parser import (
    Binary, Func, FormulaSyntaxError, NameNode, RefNode, StructRef, StructRefNode, normalize, parse,
    parse_tokens, to_display, to_storage,
)
from .dates import to_serial
from .inputparse import parse_input
from .refs import MAX_COL, MAX_ROW, Rect, col_letter
from .spill import SpillIndex
from .undo import UndoEntry, UndoTree

sys.setrecursionlimit(max(sys.getrecursionlimit(), 20000))

_DATE_TYPES = (dt.datetime, dt.date, dt.time, dt.timedelta)

# Functions whose result takes the number format of their first referenced cell.
_FORMAT_INHERITING = {"SUM", "AVERAGE", "MIN", "MAX", "MEDIAN", "ROUND", "ROUNDUP", "ROUNDDOWN",
                      "ABS", "SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS", "MAXIFS", "MINIFS",
                      "EDATE", "EOMONTH", "WORKDAY", "LARGE", "SMALL", "INDEX", "VLOOKUP",
                      "HLOOKUP", "XLOOKUP", "LOOKUP", "MROUND", "CEILING", "FLOOR", "TRUNC"}


# Returned by Workbook._compute instead of spill rows: leave the formula's spill as it is.
_KEEP = object()


class Formula:
    __slots__ = ("text", "ast", "volatile", "unsupported", "error", "array", "positional", "structured",
                 "uses_names", "legacy", "arrayish")

    def __init__(self, text: str):
        self.text = text  # display form, with leading '='
        self.ast = None
        self.volatile = False
        self.unsupported: list[str] = []
        self.error: str | None = None
        self.array = False
        self.positional = False  # uses ROW()/COLUMN(): value depends on where the cell is
        self.structured = False  # uses table references: re-resolve when tables change
        self.uses_names = False
        # From a file, saved as an ordinary formula: no spilling, ranges use implicit intersection.
        self.legacy = False
        # The last result was an array or range (saved as a dynamic-array formula).
        self.arrayish = False


class Sheet:
    def __init__(self, book: "Workbook", ws):
        self.book = book
        self.ws = ws
        self.formulas: dict[tuple, Formula] = {}
        self.values: dict[tuple, object] = {}
        self.cached: dict[tuple, tuple] = {}  # (r,c) -> (storage text, cached value from file)
        self.stale: set[tuple] = set()  # cells showing a cached value we could not recompute
        self.spill = SpillIndex()
        self.max_row = 0
        self.max_col = 0
        self._colw: dict[int, tuple] | None = None
        self._colo: dict[int, tuple] = {}  # col -> (outline level, collapsed flag), outlined columns only
        self.outline_gen = 0  # bumped whenever row/column dimensions may have changed
        self._row_outline_cache = None
        self.recompute_bounds()

    @property
    def name(self) -> str:
        return self.ws.title

    def __repr__(self):
        return f"<Sheet {self.name}>"

    def recompute_bounds(self):
        mr, mc = self.spill.extent()
        for (r, c), cell in self.ws._cells.items():
            if cell._value is not None:
                if r > mr:
                    mr = r
                if c > mc:
                    mc = c
        self.max_row, self.max_col = mr, mc

    # ------------------------------------------------------------------ layout helpers

    def _build_colw(self):
        default = self.ws.sheet_format.defaultColWidth
        if default is None:
            base = self.ws.sheet_format.baseColWidth or 8
            default = base + 0.43 if base == 8 else base + 0.71
        self._default_w = float(default)
        cw: dict[int, tuple] = {}
        co: dict[int, tuple] = {}
        for key, dim in self.ws.column_dimensions.items():
            lo = dim.min or 0
            hi = dim.max or 0
            if not lo:
                try:
                    from .refs import col_index

                    lo = hi = col_index(key)
                except Exception:
                    continue
            w = dim.width if (dim.customWidth or dim.width) else None
            outline = (dim.outlineLevel or 0, bool(dim.collapsed))
            for c in range(lo, min(hi, MAX_COL) + 1):
                cw[c] = (w if w else self._default_w, bool(dim.hidden))
                if outline != (0, False):
                    co[c] = outline
        self._colw = cw
        self._colo = co

    def col_width(self, c: int) -> tuple[float, bool]:
        if self._colw is None:
            self._build_colw()
        return self._colw.get(c, (self._default_w, False))

    def col_outline(self, c: int) -> tuple[int, bool]:
        """(outline level, collapsed flag) of column c."""
        if self._colw is None:
            self._build_colw()
        return self._colo.get(c, (0, False))

    def row_hidden(self, r: int) -> bool:
        d = self.ws.row_dimensions.get(r)
        return bool(d is not None and d.hidden)

    def invalidate_layout(self):
        self._colw = None
        self.outline_gen += 1

    def merged_ranges(self):
        return self.ws.merged_cells.ranges

    def merged_at(self, r: int, c: int):
        """Return the merged CellRange covering (r, c), if any."""
        for mr in self.ws.merged_cells.ranges:
            if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                return mr
        return None

    # ------------------------------------------------------------------ cell content

    def cell(self, r: int, c: int):
        return self.ws._cells.get((r, c))

    def is_formula(self, r: int, c: int) -> bool:
        return (r, c) in self.formulas

    def value(self, r: int, c: int):
        """The displayed value of a cell (computed value for formulas)."""
        return self.book.get_value(self, r, c)

    def number_format(self, r: int, c: int) -> str:
        cell = self.ws._cells.get((r, c))
        fmt = cell.number_format if cell is not None else "General"
        if fmt == "General" and self.spill.cols:
            anchor = self.spill.owner(r, c)
            if anchor is not None and anchor != (r, c):
                src = self.ws._cells.get(anchor)
                if src is not None:
                    return src.number_format
        return fmt

    def spill_anchor(self, r: int, c: int):
        """The formula cell whose spilled result covers (r, c) (not counting the formula cell itself)."""
        if not self.spill.cols:
            return None
        anchor = self.spill.owner(r, c)
        return None if anchor == (r, c) else anchor

    def input_text(self, r: int, c: int) -> str:
        """Text to show in the formula bar / editor for a cell."""
        f = self.formulas.get((r, c))
        if f is not None:
            return f.text
        cell = self.ws._cells.get((r, c))
        if cell is None or cell._value is None:
            return ""
        v = cell._value
        if cell.data_type == "s" and isinstance(v, str):
            needs_quote = v.startswith("=") or (parse_input(v).kind not in ("text",))
            return ("'" + v) if needs_quote else v
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, _DATE_TYPES) or self._is_date_fmt(cell):
            from ..format.numfmt import edit_text_for_date

            serial = v if isinstance(v, (int, float)) else to_serial(v, self.book.epoch1904)
            return edit_text_for_date(serial, cell.number_format, self.book.epoch1904)
        if isinstance(v, (int, float)):
            from ..formula.values import format_general_number

            if cell.number_format.endswith("%") and isinstance(v, (int, float)):
                return format_general_number(round(v * 100, 13)) + "%"
            return format_general_number(v)
        return str(v)

    @staticmethod
    def _is_date_fmt(cell) -> bool:
        from ..format.numfmt import is_date_format

        return is_date_format(cell.number_format)


class Workbook:
    def __init__(self, wb, path: str | None = None):
        self.wb = wb
        self.path = path
        self.epoch1904 = wb.epoch == MAC_EPOCH
        self.sheets: list[Sheet] = [Sheet(self, ws) for ws in wb.worksheets]
        self.graph = DepGraph()
        self.dirty: set = set()
        self.evaluating: set = set()
        self.names: dict[str, object] = {}
        self.sheet_names: dict[tuple, object] = {}  # (sheet lower, name lower) -> ast
        self.undo = UndoTree()
        self.modified = False
        self.warnings: list[str] = []  # content we cannot preserve on save
        self._tx: dict | None = None
        self._tx_keys: set | None = None
        self._pending_tokens: dict = {}
        self._damaged: set = set()
        self._damaged_pre: dict = {}
        self._tables = None  # TableIndex, built on demand
        self._tx_tables: dict | None = None
        self._tx_names: dict | None = None
        self._tx_expand = False
        self.legacy: set[int] = set()  # id()s of cells holding formulas loaded without dynamic-array support
        self._spill_changed: list = []  # (sheet, r1, c1, r2, c2, anchor) whose spilled values changed
        self._load_names()

    # ------------------------------------------------------------------ setup

    def _load_names(self):
        self.names.clear()
        self.sheet_names.clear()

        def parse_name(dn):
            text = dn.attr_text or ""
            try:
                return parse("=" + to_display("=" + text)[1:])
            except (FormulaSyntaxError, Unsupported):
                return None

        try:
            items = list(self.wb.defined_names.items())
        except AttributeError:
            items = []
        for name, dn in items:
            ast = parse_name(dn)
            if ast is not None:
                self.names[name.lower()] = ast
        for s in self.sheets:
            for name, dn in getattr(s.ws, "defined_names", {}).items():
                ast = parse_name(dn)
                if ast is not None:
                    self.sheet_names[(s.name.lower(), name.lower())] = ast

    def register_all(self, use_cached: bool = True):
        """Parse every formula in the workbook and build the dependency graph."""
        self.graph.clear()
        self._tables = None
        for s in self.sheets:
            s.formulas.clear()
            for (r, c), cell in s.ws._cells.items():
                if cell.data_type == "f" and cell._value is not None:
                    self._register(s, r, c, _formula_text(cell._value), array=isinstance(cell._value, ArrayFormula))

    def initial_calc(self):
        """Compute formulas with no cached value (and volatile ones); keep Excel's cached results otherwise."""
        dirty = set()
        for s in self.sheets:
            for rc, f in s.formulas.items():
                cached = s.cached.get(rc)
                if cached is None or cached[1] is None or f.volatile or (cached[1] is SPILL and not f.legacy):
                    dirty.add((s, rc[0], rc[1]))
                else:
                    s.values[rc] = cached[1]
                    if f.ast is None or f.unsupported:
                        s.stale.add(rc)
        if dirty:
            dirty |= self.graph.all_dependents(dirty)
            self._run(dirty)

    # ------------------------------------------------------------------ sheets

    def sheet_by_name(self, name: str) -> Sheet | None:
        low = name.lower()
        for s in self.sheets:
            if s.ws.title == name or s.ws.title.lower() == low:
                return s
        return None

    def sheet_index(self, sheet: Sheet) -> int:
        return self.sheets.index(sheet)

    # ------------------------------------------------------------------ engine API

    def bounds(self, sheet: Sheet) -> tuple[int, int]:
        return sheet.max_row, sheet.max_col

    def is_formula(self, sheet: Sheet, r: int, c: int) -> bool:
        return (r, c) in sheet.formulas

    def formula_text(self, sheet: Sheet, r: int, c: int) -> str | None:
        """A cell's formula as the user sees it (with '='), or None if it holds no formula."""
        f = sheet.formulas.get((r, c))
        return None if f is None else f.text

    def resolve_name(self, name: str, sheet: Sheet, explicit_sheet: str | None = None, tables: bool = True):
        low = name.lower()
        if explicit_sheet is not None:
            return self.sheet_names.get((explicit_sheet.lower(), low))
        ast = self.sheet_names.get((sheet.name.lower(), low))
        if ast is not None:
            return ast
        ast = self.names.get(low)
        if ast is None and tables:
            hit = self.table_index().by_name.get(low)
            if hit is not None:
                return StructRefNode(StructRef(hit.name))  # a bare table name is its data
        return ast

    # ------------------------------------------------------------------ tables

    def table_index(self):
        if self._tables is None:
            from .tables import TableIndex

            self._tables = TableIndex(self)
        return self._tables

    def invalidate_tables(self):
        self._tables = None

    def resolve_sref(self, sr, sheet: Sheet, r: int, c: int):
        from ..formula.values import RefValue
        from .tables import resolve

        res = resolve(self, sr, sheet, r, c)
        if isinstance(res, tuple):
            return RefValue(*res)
        return res

    def reregister_structured(self) -> set:
        """Re-parse formulas that use table references (after tables changed). Returns their keys."""
        self._tables = None
        keys = set()
        for s in self.sheets:
            cells = s.ws._cells
            for rc, f in list(s.formulas.items()):
                if not (f.structured or f.uses_names):
                    continue
                cell = cells.get(rc)
                if cell is None or cell.data_type != "f" or cell._value is None:
                    continue
                value = s.values.get(rc)
                self._unregister(s, rc[0], rc[1], keep_spill=value is not None)
                self._register(s, rc[0], rc[1], _formula_text(cell._value), array=isinstance(cell._value, ArrayFormula))
                if value is not None:
                    s.values[rc] = value
                keys.add((s, rc[0], rc[1]))
        return keys

    def tables_changed(self, extra=()) -> set:
        """Re-resolve table references and recalculate what depends on them."""
        changed = self.reregister_structured() | set(extra)
        changed |= self._spill_watchers(changed)
        dirty = self.graph.all_dependents(changed)
        dirty |= {k for k in changed if (k[1], k[2]) in k[0].formulas}
        vol = self.graph.volatile
        if vol:
            dirty |= vol | self.graph.all_dependents(vol)
        if dirty:
            self._run(dirty)
        elif self._spill_changed:
            dirty = self._settle_spills()
        self.modified = True
        return dirty | set(extra)

    def get_value(self, sheet: Sheet, r: int, c: int):
        rc = (r, c)
        if rc in sheet.formulas:
            if self.dirty:
                key = (sheet, r, c)
                if key in self.dirty:
                    if key in self.evaluating:
                        return CIRC
                    self._eval(key)
            return sheet.values.get(rc, 0)
        cell = sheet.ws._cells.get(rc)
        if cell is None:
            return self._spilled(sheet, r, c) if sheet.spill.cols else None
        v = cell._value
        t = type(v)
        if v is None:
            return self._spilled(sheet, r, c) if sheet.spill.cols else None
        if t is str:
            if cell.data_type == "e":
                return ERROR_CODES.get(v.upper(), VALUE)
            return v
        if t is int or t is float or t is bool:
            return v
        if isinstance(v, _DATE_TYPES):
            return to_serial(v, self.epoch1904)
        return str(v)

    def range_rows(self, sheet: Sheet, r1: int, c1: int, r2: int, c2: int) -> list[list]:
        """Values of a rectangle as a list of rows (fast path for range-heavy formulas)."""
        formulas, values, cells = sheet.formulas, sheet.values, sheet.ws._cells
        dirty = self.dirty
        get_value = self.get_value
        spilled = self._spilled if sheet.spill.cols else None
        rows = []
        cols = range(c1, c2 + 1)
        for r in range(r1, r2 + 1):
            row = []
            ap = row.append
            for c in cols:
                rc = (r, c)
                if rc in formulas:
                    if dirty and (sheet, r, c) in dirty:
                        ap(get_value(sheet, r, c))
                    else:
                        ap(values.get(rc, 0))
                    continue
                cell = cells.get(rc)
                if cell is None:
                    ap(None if spilled is None else spilled(sheet, r, c))
                    continue
                v = cell._value
                t = type(v)
                if v is None and spilled is not None:
                    ap(spilled(sheet, r, c))
                elif t is int or t is float or v is None or t is bool or (t is str and cell.data_type != "e"):
                    ap(v)
                else:
                    ap(get_value(sheet, r, c))
            rows.append(row)
        return rows

    def range_values(self, sheet: Sheet, r1: int, c1: int, r2: int, c2: int) -> list:
        """Non-blank values of a rectangle in row-major order."""
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 4 * len(sheet.ws._cells) + 64:
            # sparse: walk the existing cells instead of every position
            keys = {k for k in sheet.ws._cells if r1 <= k[0] <= r2 and c1 <= k[1] <= c2}
            if sheet.spill.anchors:
                keys.update((r, c) for r, c, _ in sheet.spill.cells_in(r1, c1, r2, c2))
            keys = sorted(keys)
            out = []
            for r, c in keys:
                v = self.get_value(sheet, r, c)
                if v is not None:
                    out.append(v)
            return out
        out = []
        for row in self.range_rows(sheet, r1, c1, r2, c2):
            for v in row:
                if v is not None:
                    out.append(v)
        return out

    # ------------------------------------------------------------------ spilling

    def _spilled(self, sheet: Sheet, r: int, c: int):
        """Value of an empty cell that may be covered by a spilled array."""
        sp = sheet.spill
        anchor = sp.owner(r, c)
        if anchor is None:
            return None
        key = (sheet, anchor[0], anchor[1])
        if key in self.evaluating:
            return CIRC  # the formula reads its own spill range
        if self.dirty and key in self.dirty:
            self._eval(key)
            anchor = sp.owner(r, c)
            if anchor is None:
                return None
        return sp.value(anchor, r, c)

    def spill_ref(self, sheet: Sheet, r: int, c: int):
        """The range spilled by the formula in (r, c), for A1# references."""
        from ..formula.values import RefValue

        if (r, c) not in sheet.formulas:
            return REF
        self.get_value(sheet, r, c)  # make sure it's up to date
        sp = sheet.spill
        if (r, c) in sp.anchors:
            return RefValue(sheet, *sp.rect((r, c)))
        if (r, c) in sp.blocked:
            return REF
        return RefValue(sheet, r, c, r, c)

    def _spill_blocked(self, sheet: Sheet, r: int, c: int, r2: int, c2: int) -> bool:
        """Whether anything is in the way of spilling from (r, c) to (r2, c2)."""
        if r2 > MAX_ROW or c2 > MAX_COL:
            return True
        cells = sheet.ws._cells
        if (r2 - r + 1) * (c2 - c + 1) <= 4 * len(cells):
            for rr in range(r, r2 + 1):
                for cc in range(c, c2 + 1):
                    cell = cells.get((rr, cc))
                    if cell is not None and (cell._value is not None or isinstance(cell, MergedCell)) \
                            and (rr, cc) != (r, c):
                        return True
        else:
            for (rr, cc), cell in cells.items():
                if r <= rr <= r2 and c <= cc <= c2 and (rr, cc) != (r, c) and (
                        cell._value is not None or isinstance(cell, MergedCell)):
                    return True
        for mr in sheet.ws.merged_cells.ranges:
            if mr.min_row <= r2 and mr.max_row >= r and mr.min_col <= c2 and mr.max_col >= c:
                return True
        return sheet.spill.overlaps(r, c, r2, c2, exclude=(r, c))

    def _place_spill(self, sheet: Sheet, r: int, c: int, rows, fixed: bool):
        """Record a formula's spilled result (rows=None: it no longer spills). Returns the value
        for the formula cell itself."""
        sp = sheet.spill
        anchor = (r, c)
        old = sp.remove(anchor)
        sp.blocked.pop(anchor, None)
        if rows is None:
            if old is not None:
                self._spill_changed.append((sheet, r, c, r + len(old) - 1, c + len(old[0]) - 1, anchor))
            return None
        r2, c2 = r + len(rows) - 1, c + len(rows[0]) - 1
        if not fixed and self._spill_blocked(sheet, r, c, r2, c2):
            sp.blocked[anchor] = (len(rows), len(rows[0]))
            if old is not None:
                self._spill_changed.append((sheet, r, c, r + len(old) - 1, c + len(old[0]) - 1, anchor))
            return SPILL
        sp.add(anchor, rows)
        if old != rows:
            if old is not None:
                r2 = max(r2, r + len(old) - 1)
                c2 = max(c2, c + len(old[0]) - 1)
            self._spill_changed.append((sheet, r, c, r2, c2, anchor))
        if r2 > sheet.max_row:
            sheet.max_row = min(r2, MAX_ROW)
        if c2 > sheet.max_col:
            sheet.max_col = min(c2, MAX_COL)
        return rows[0][0]

    def _drop_spill(self, sheet: Sheet, r: int, c: int):
        sp = sheet.spill
        if sp:
            sp.blocked.pop((r, c), None)
            if (r, c) in sp.anchors:
                self._place_spill(sheet, r, c, None, False)

    def _spill_watchers(self, keys) -> set:
        """Formulas whose spill area (actual, or wanted but blocked) contains any of the cells."""
        out = set()
        for key in keys:
            sheet, r, c = key
            sp = sheet.spill
            if not sp:
                continue
            anchor = sp.owner(r, c) if sp.cols else None
            if anchor is not None and anchor != (r, c):
                out.add((sheet, anchor[0], anchor[1]))
            for ar, ac in sp.blocked_in(r, c, r, c):
                if (ar, ac) != (r, c):
                    out.add((sheet, ar, ac))
        return out

    def take_spills(self, sheet: Sheet) -> list:
        """Forget a sheet's spills (before its cells move); returns the formula positions that
        spilled or were blocked, so they can be recalculated afterwards."""
        sp = sheet.spill
        anchors = list(sp.anchors) + list(sp.blocked)
        sp.clear()
        return anchors

    def mark_legacy(self):
        """Note that the ordinary formulas just loaded from a file use pre-dynamic-array semantics."""
        legacy = self.legacy
        for s in self.sheets:
            cells = s.ws._cells
            for rc, f in s.formulas.items():
                cell = cells.get(rc)
                if cell is not None and isinstance(cell._value, str):
                    legacy.add(id(cell))
                    f.legacy = True

    def adopt_arrays(self, dynamic: dict):
        """Take over array formulas loaded from a file. The other cells of an array formula's
        range hold its saved results as plain values: move those into the spill index.
        dynamic: {sheet: {(r, c), ...}} anchors of dynamic-array (spilling) formulas; the rest are
        legacy Ctrl+Shift+Enter arrays with a fixed size."""
        from openpyxl.utils.cell import range_boundaries

        for s in self.sheets:
            cells = s.ws._cells
            dyn = dynamic.get(s, ())
            for (r, c), cell in list(cells.items()):
                value = cell._value
                if cell.data_type != "f" or not isinstance(value, ArrayFormula):
                    continue
                try:
                    c1, r1, c2, r2 = range_boundaries(value.ref) if value.ref else (c, r, c, r)
                except (TypeError, ValueError):
                    c1, r1, c2, r2 = c, r, c, r
                if (r1, c1) != (r, c):
                    c2, r2 = c + (c2 - c1), r + (r2 - r1)
                cached = s.cached.get((r, c))
                rows = []
                for rr in range(r, r2 + 1):
                    row = []
                    for cc in range(c, c2 + 1):
                        if (rr, cc) == (r, c):
                            row.append(cached[1] if cached is not None else None)
                            continue
                        other = cells.get((rr, cc))
                        if other is None or other._value is None or other.data_type == "f":
                            row.append(None)
                            continue
                        row.append(self.get_value(s, rr, cc))
                        other._value = None
                        other.data_type = "n"
                    rows.append(row)
                if (r, c) in dyn:
                    self.legacy.discard(id(cell))
                    cell._value = _formula_text(value)
                value_before = s.values.get((r, c))
                self._unregister(s, r, c)
                f = self._register(s, r, c, _formula_text(cell._value), array=isinstance(cell._value, ArrayFormula))
                f.arrayish = (r, c) in dyn
                if value_before is not None:
                    s.values[(r, c)] = value_before
                if len(rows) > 1 or len(rows[0]) > 1:
                    s.spill.add((r, c), rows)
            s.recompute_bounds()

    def _array_size(self, sheet: Sheet, r: int, c: int):
        """(rows, cols) of a legacy array formula's fixed range."""
        from openpyxl.utils.cell import range_boundaries

        cell = sheet.ws._cells.get((r, c))
        value = cell._value if cell is not None else None
        if not isinstance(value, ArrayFormula) or not value.ref:
            return (1, 1)
        try:
            c1, r1, c2, r2 = range_boundaries(value.ref)
        except (TypeError, ValueError):
            return (1, 1)
        return (r2 - r1 + 1, c2 - c1 + 1)

    def array_part(self, sheet: Sheet, r: int, c: int):
        """The anchor of the legacy (fixed-size) array formula covering (r, c), if any."""
        anchor = sheet.spill_anchor(r, c)
        if anchor is not None:
            f = sheet.formulas.get(anchor)
            if f is not None and f.array:
                return anchor
        return None

    # ------------------------------------------------------------------ formula registration

    def _register(self, sheet: Sheet, r: int, c: int, storage_text: str, array: bool = False, tokens=None):
        display = to_display(storage_text)
        f = Formula(display)
        if "[" in display and self.table_index().by_sheet.get(sheet):
            from .tables import display_in_place

            f.text = display_in_place(self, display, sheet, r, c)
        f.array = array
        if self.legacy:
            cell = sheet.ws._cells.get((r, c))
            f.legacy = cell is not None and id(cell) in self.legacy and not array
        key = (sheet, r, c)
        try:
            f.ast = parse_tokens(tokens) if tokens is not None else parse(display)
        except FormulaSyntaxError as e:
            f.error = str(e)
        except Unsupported as e:
            f.error = str(e)
            f.unsupported = [str(e)]
        refs: list = []
        if f.ast is not None:
            ref_nodes, name_nodes, funcs, dyn, sref_nodes = ev.analyze(f.ast)
            functions = ev.FUNCTIONS
            unsupported = [n for n in funcs if n not in functions]
            if unsupported:
                # MYFUNC(1) calls the LAMBDA in a defined name: depend on what the name uses
                named = [n for n in unsupported if self.resolve_name(n, sheet, None, tables=False) is not None]
                if named:
                    unsupported = [n for n in unsupported if n not in named]
                    name_nodes = name_nodes + [NameNode(n) for n in named]
            f.unsupported = sorted(unsupported) if unsupported else []
            f.volatile = dyn or any(functions[n].volatile for n in funcs if n in functions)
            f.positional = "ROW" in funcs or "COLUMN" in funcs or "CELL" in funcs
            by_name = self.sheet_by_name
            for n in ref_nodes:
                s = sheet if n.sheet is None else by_name(n.sheet)
                if s is not None:
                    refs.append((s, n.r1, n.c1, n.r2, n.c2))
            if sref_nodes:
                f.structured = True
                for n in sref_nodes:
                    self._sref_dep(n.ref, sheet, r, c, refs)
            if name_nodes:
                f.uses_names = True
            for n in name_nodes:
                target = self.resolve_name(n.name, sheet, n.sheet)
                if target is not None:
                    f.volatile |= self._collect_refs(target, sheet, refs, 1, r, c)
        sheet.formulas[(r, c)] = f
        self.graph.add(key, refs, f.volatile)
        return f

    def _sref_dep(self, sr, sheet: Sheet, r: int, c: int, out: list):
        from .tables import resolve

        res = resolve(self, sr, sheet, r, c)
        if isinstance(res, tuple):
            out.append(res)

    def _collect_refs(self, node, sheet: Sheet, out: list, depth: int, r: int = 1, c: int = 1) -> bool:
        """Append (sheet, r1, c1, r2, c2) refs used by node. Returns True if refs are dynamic."""
        dynamic = False
        for n in ev.walk(node):
            t = type(n)
            if t is RefNode:
                s = sheet if n.sheet is None else self.sheet_by_name(n.sheet)
                if s is not None:
                    out.append((s, n.r1, n.c1, n.r2, n.c2))
            elif t is StructRefNode:
                self._sref_dep(n.ref, sheet, r, c, out)
            elif t is NameNode and depth < 20:
                target = self.resolve_name(n.name, sheet, n.sheet)
                if target is not None:
                    dynamic |= self._collect_refs(target, sheet, out, depth + 1, r, c)
            elif t is Func and depth < 20 and n.name not in ev.FUNCTIONS:
                target = self.resolve_name(n.name, sheet, None, tables=False)  # a LAMBDA name
                if target is not None:
                    dynamic |= self._collect_refs(target, sheet, out, depth + 1, r, c)
            elif t is Binary and n.op == ":" and not (type(n.left) in ev._STATIC_REFS
                                                      and type(n.right) in ev._STATIC_REFS):
                dynamic = True
        return dynamic

    def _unregister(self, sheet: Sheet, r: int, c: int, keep_spill: bool = False):
        """Forget a cell's formula. keep_spill=True when it's re-registered straight away with its
        old value kept."""
        if sheet.formulas.pop((r, c), None) is not None:
            sheet.values.pop((r, c), None)
            sheet.stale.discard((r, c))
            self.graph.remove((sheet, r, c))
            if not keep_spill:
                self._drop_spill(sheet, r, c)

    # ------------------------------------------------------------------ evaluation

    def _eval(self, key):
        sheet, r, c = key
        f = sheet.formulas.get((r, c))
        if f is None:
            self.dirty.discard(key)
            return
        self.evaluating.add(key)
        try:
            v, rows = self._compute(sheet, r, c, f)
        finally:
            self.evaluating.discard(key)
        if rows is not _KEEP and (rows is not None or sheet.spill):
            top = self._place_spill(sheet, r, c, rows, f.array)
            if rows is not None:
                v = top
        sheet.values[(r, c)] = v
        self.dirty.discard(key)

    def _compute(self, sheet: Sheet, r: int, c: int, f: Formula):
        """-> (value, spill rows or None, or _KEEP to leave the spill as it is)"""
        cached = sheet.cached.get((r, c))
        has_cache = cached is not None and cached[0] == f.text
        if f.ast is None or f.unsupported:
            if has_cache:
                sheet.stale.add((r, c))
                return cached[1], _KEEP
            return NAME, None
        ctx = ev.Ctx(self, sheet, r, c)
        ctx.legacy = f.legacy
        rows = None
        try:
            res = ev.evaluate(f.ast, ctx)
            if f.legacy:
                v = ev.finalize(ctx, res, legacy=True)
            else:
                f.arrayish = ctx.arrays or ev.is_multi(res)
                size = self._array_size(sheet, r, c) if f.array else None
                if size == (1, 1):
                    size = None
                rows = ev.spill_rows(ctx, res, size)
                v = rows[0][0] if rows is not None else ev.finalize(ctx, res)
        except XLError as e:
            v = e.err
        except Unsupported:
            if has_cache:
                sheet.stale.add((r, c))
                return cached[1], _KEEP
            v = NAME
        except RecursionError:
            v = CIRC
        except Exception:  # never let one bad formula take the app down
            v = VALUE
        sheet.stale.discard((r, c))
        return v, rows

    def _run(self, dirty: set):
        if len(dirty) > 20_000:
            from .gcutil import paused_gc

            with paused_gc():
                return self._run_inner(dirty)
        return self._run_inner(dirty)

    def _run_inner(self, dirty: set):
        date_fns.STATE["epoch1904"] = self.epoch1904
        self.dirty |= dirty
        for key in self.graph.topo_order(set(self.dirty)):
            if key in self.dirty:
                self._eval(key)
        self.dirty.clear()
        self._settle_spills()

    def _settle_spills(self) -> set:
        """Recalculate formulas that read cells whose spilled values changed (they may have been
        evaluated before the spill happened, or not been dirty at all). Returns their keys."""
        done: set = set()
        for _ in range(20):
            if not self._spill_changed:
                break
            changed, self._spill_changed = self._spill_changed, []
            more: set = set()
            for sheet, r1, c1, r2, c2, anchor in changed:
                if sheet not in self.sheets:
                    continue
                more |= self.graph.dependents_rect(sheet, r1, c1, r2, c2)
                more.update((sheet, ar, ac) for ar, ac in sheet.spill.blocked_in(r1, c1, r2, c2)
                            if (ar, ac) != anchor)
            if not more:
                continue
            more |= self.graph.all_dependents(more)
            done |= more
            self.dirty |= more
            for key in self.graph.topo_order(set(more)):
                if key in self.dirty:
                    self._eval(key)
            self.dirty.clear()
        self._spill_changed = []
        return done

    def recalc(self, changed) -> set:
        """Recalculate everything that depends on the changed cells (and volatile formulas)."""
        changed = set(changed)
        changed |= self._spill_watchers(changed)
        dirty = self.graph.all_dependents(changed)
        for key in changed:
            if (key[1], key[2]) in key[0].formulas:
                dirty.add(key)
        vol = self.graph.volatile
        if vol:
            dirty |= vol
            dirty |= self.graph.all_dependents(vol)
        if dirty:
            self._run(dirty)
        elif self._spill_changed:
            dirty = self._settle_spills()
        return dirty

    def recalc_all(self):
        dirty = {(s, r, c) for s in self.sheets for (r, c) in s.formulas}
        self._run(dirty)

    # ------------------------------------------------------------------ low-level writes (undo-aware)

    def _snap(self, sheet: Sheet, r: int, c: int):
        cell = sheet.ws._cells.get((r, c))
        if cell is None or isinstance(cell, MergedCell):
            return None
        return (cell._value, cell.data_type, copy(cell._style))

    def _touch(self, sheet: Sheet, r: int, c: int):
        if self._tx is not None:
            key = (sheet, r, c)
            if key not in self._tx:
                self._tx[key] = self._snap(sheet, r, c)
            self._tx_keys.add(key)

    def _restore(self, key, snap):
        sheet, r, c = key
        ws = sheet.ws
        self._unregister(sheet, r, c)
        if snap is None:
            cell = ws._cells.get((r, c))
            if cell is not None and not isinstance(cell, MergedCell):
                del ws._cells[(r, c)]
            return
        value, data_type, style = snap
        cell = ws._cells.get((r, c))
        if isinstance(cell, MergedCell):
            return
        if cell is None:
            cell = ws.cell(row=r, column=c)
            self.legacy.discard(id(cell))  # a new object may reuse a forgotten cell's id
        cell._value = value
        cell.data_type = data_type
        cell._style = copy(style)
        if data_type == "f" and value is not None:
            self._register(sheet, r, c, _formula_text(value), array=isinstance(value, ArrayFormula))
        if value is not None:
            sheet.max_row = max(sheet.max_row, r)
            sheet.max_col = max(sheet.max_col, c)

    def _apply_snaps(self, snaps: dict, tables: dict | None = None, names: dict | None = None) -> set:
        if tables:
            from .tables import restore

            for sheet, snap in tables.items():
                restore(self, sheet, snap)
        if names:
            self._restore_names(names)
        for key, snap in snaps.items():
            self._restore(key, snap)
        self.modified = True
        if tables or names:
            return self.tables_changed(snaps.keys())
        return self.recalc(snaps.keys())

    def _touch_tables(self, sheet: Sheet):
        """Record a sheet's tables before changing them (inside a transaction) for undo."""
        if self._tx_tables is not None and sheet not in self._tx_tables:
            from .tables import snapshot

            self._tx_tables[sheet] = snapshot(sheet)

    def _names_state(self) -> dict:
        out = {}
        for key, dn in self.wb.defined_names.items():
            out[(None, key)] = dn.attr_text
        for s in self.sheets:
            for key, dn in s.ws.defined_names.items():
                out[(s, key)] = dn.attr_text
        return out

    def _touch_names(self):
        if self._tx_names is not None and not self._tx_names:
            self._tx_names.update(self._names_state())

    def _restore_names(self, state: dict):
        for (s, key), text in state.items():
            store = self.wb.defined_names if s is None else s.ws.defined_names
            dn = store.get(key)
            if dn is not None:
                dn.attr_text = text
        self._load_names()

    @contextmanager
    def transaction(self, label: str, where=None):
        """Group cell writes into one undo step and recalculate once at the end."""
        if self._tx is not None:
            yield
            return
        self._tx, self._tx_keys, self._tx_tables, self._tx_names = {}, set(), {}, {}
        try:
            yield
            if self._tx_keys and self.table_index().by_sheet:
                from .tables import after_edit

                after_edit(self, set(self._tx_keys), self._tx_expand)
        finally:
            before, keys, tbefore, nbefore = self._tx, self._tx_keys, self._tx_tables, self._tx_names
            self._tx = self._tx_keys = self._tx_tables = self._tx_names = None
            self._tx_expand = False
        if keys or tbefore or nbefore:
            from .tables import snapshot

            after = {k: self._snap(*k) for k in before}
            tafter = {s: snapshot(s) for s in tbefore}
            nafter = self._names_state() if nbefore else {}
            self.undo.push(UndoEntry(label, lambda: self._apply_snaps(before, tbefore, nbefore),
                                     lambda: self._apply_snaps(after, tafter, nafter), where))
            self.modified = True
            if tbefore or nbefore:
                if nbefore:
                    self._load_names()
                self.last_changed = self.tables_changed(keys) | keys
            else:
                self.last_changed = self.recalc(keys) | keys
        else:
            self.last_changed = set()

    def writable_cell(self, sheet: Sheet, r: int, c: int):
        cell = sheet.ws._cells.get((r, c))
        if isinstance(cell, MergedCell):
            return None
        if cell is None:
            cell = sheet.ws.cell(row=r, column=c)
        return cell

    def set_raw(self, sheet: Sheet, r: int, c: int, value, data_type: str | None = None):
        """Write a raw value (formulas as '=...' storage text). Must be inside a transaction."""
        existing = sheet.ws._cells.get((r, c))
        if isinstance(existing, MergedCell):
            return
        if value is None and existing is None:
            return
        if isinstance(value, str) and value.startswith("=") and "[" in value and data_type is None:
            from .tables import qualify

            value = qualify(self, value, sheet, r, c)
        self._touch(sheet, r, c)
        self._unregister(sheet, r, c)
        cell = existing if existing is not None else sheet.ws.cell(row=r, column=c)
        if self.legacy:
            self.legacy.discard(id(cell))
        try:
            cell.value = value
        except IllegalCharacterError:
            # Control characters (e.g. from CHAR(7)): openpyxl refuses them, Excel doesn't.
            # They're escaped when the file is saved.
            cell._value, cell.data_type = value, "s"
        if data_type is not None:
            cell.data_type = data_type
        if cell.data_type == "f":
            self._register(sheet, r, c, _formula_text(cell._value))
        if value is not None:
            if r > sheet.max_row:
                sheet.max_row = r
            if c > sheet.max_col:
                sheet.max_col = c

    def set_style(self, sheet: Sheet, r: int, c: int, **attrs):
        """Change number_format / font / fill / alignment of a cell (inside a transaction)."""
        cell = self.writable_cell(sheet, r, c)
        if cell is None:
            return
        self._touch(sheet, r, c)
        for k, v in attrs.items():
            setattr(cell, k, v)

    # ------------------------------------------------------------------ user-level edits

    def set_input(self, sheet: Sheet, r: int, c: int, text: str):
        """Apply text typed by the user. Raises FormulaSyntaxError for bad formulas."""
        part_of = self.array_part(sheet, r, c)
        if part_of is not None:
            raise ValueError("You can't change part of an array formula; edit the formula in "
                             f"{col_letter(part_of[1])}{part_of[0]}")
        p = parse_input(text, self.epoch1904)
        with self.transaction("Edit cell", (sheet.name, Rect(r, c, r, c))):
            self._tx_expand = True
            self._apply_parsed(sheet, r, c, p)

    def _apply_parsed(self, sheet, r, c, p):
        if p.kind == "formula":
            text = normalize(p.value)
            try:
                ast = parse(text)
            except Unsupported:
                ast = None
            self.set_raw(sheet, r, c, to_storage(text))
            if ast is not None:
                self._inherit_format(sheet, r, c, ast)
            return
        if p.kind == "empty":
            self.set_raw(sheet, r, c, None)
            return
        if p.kind == "text":
            v = p.value
            self.set_raw(sheet, r, c, v, "s" if isinstance(v, str) and v.startswith("=") else None)
            return
        if p.kind == "error":
            self.set_raw(sheet, r, c, p.value, "e")
            return
        self.set_raw(sheet, r, c, p.value)
        if p.number_format:
            cell = sheet.ws._cells.get((r, c))
            if cell is not None and cell.number_format == "General":
                self.set_style(sheet, r, c, number_format=p.number_format)

    def _inherit_format(self, sheet, r, c, ast):
        cell = sheet.ws._cells.get((r, c))
        if cell is None or cell.number_format != "General":
            return
        fmt = None
        root = ast
        if type(root) is Func and root.name in ("TODAY", "DATE", "DATEVALUE"):
            fmt = "m/d/yyyy"
        elif type(root) is Func and root.name == "NOW":
            fmt = "m/d/yyyy h:mm"
        elif type(root) is Func and root.name == "TIME":
            fmt = "h:mm AM/PM"
        else:
            ref = _first_ref(root)
            if ref is not None:
                s = sheet if ref.sheet is None else self.sheet_by_name(ref.sheet)
                if s is not None:
                    src = s.ws._cells.get((ref.r1, ref.c1))
                    if src is not None and src.number_format != "General":
                        fmt = src.number_format
        if fmt:
            self.set_style(sheet, r, c, number_format=fmt)

    def clear(self, sheet: Sheet, rect: Rect, formats: bool = False):
        rect = self.clip(sheet, rect)
        keys = [rc for rc in list(sheet.ws._cells.keys()) if rect.contains(*rc)] \
            if rect.rows * rect.cols > len(sheet.ws._cells) else list(rect.cells())
        with self.transaction("Clear", (sheet.name, rect)):
            for (r, c) in keys:
                cell = sheet.ws._cells.get((r, c))
                if cell is None or isinstance(cell, MergedCell):
                    continue
                if formats:
                    self._touch(sheet, r, c)
                    self._unregister(sheet, r, c)
                    del sheet.ws._cells[(r, c)]
                elif cell._value is not None:
                    self.set_raw(sheet, r, c, None)

    def clip(self, sheet: Sheet, rect: Rect) -> Rect:
        """Clip whole-row/column selections to the used area."""
        r2 = rect.r2 if rect.r2 < MAX_ROW else max(rect.r1, sheet.max_row)
        c2 = rect.c2 if rect.c2 < MAX_COL else max(rect.c1, sheet.max_col)
        return Rect(rect.r1, rect.c1, r2, c2)

    # ------------------------------------------------------------------ styles

    def toggle_font(self, sheet: Sheet, rect: Rect, attr: str):
        rect = self.clip(sheet, rect)
        anchor = sheet.ws._cells.get((rect.r1, rect.c1))
        cur = getattr(anchor.font, attr) if anchor is not None else None
        if attr == "underline":
            new = None if cur else "single"
        else:
            new = not bool(cur)
        with self.transaction(f"Toggle {attr}", (sheet.name, rect)):
            for r, c in rect.cells():
                cell = self.writable_cell(sheet, r, c)
                if cell is None:
                    continue
                font = copy(cell.font)
                setattr(font, attr, new)
                self.set_style(sheet, r, c, font=font)

    def set_number_format(self, sheet: Sheet, rect: Rect, fmt: str):
        rect = self.clip(sheet, rect)
        with self.transaction("Number format", (sheet.name, rect)):
            for r, c in rect.cells():
                self.set_style(sheet, r, c, number_format=fmt)

    def set_alignment(self, sheet: Sheet, rect: Rect, horizontal: str | None):
        rect = self.clip(sheet, rect)
        with self.transaction("Align", (sheet.name, rect)):
            for r, c in rect.cells():
                cell = self.writable_cell(sheet, r, c)
                if cell is None:
                    continue
                al = copy(cell.alignment)
                al.horizontal = horizontal
                self.set_style(sheet, r, c, alignment=al)

    def set_color(self, sheet: Sheet, rect: Rect, which: str, rgb: str | None):
        from openpyxl.styles import PatternFill

        rect = self.clip(sheet, rect)
        with self.transaction("Color", (sheet.name, rect)):
            for r, c in rect.cells():
                cell = self.writable_cell(sheet, r, c)
                if cell is None:
                    continue
                if which == "fill":
                    fill = PatternFill() if rgb is None else PatternFill("solid", fgColor="FF" + rgb)
                    self.set_style(sheet, r, c, fill=fill)
                else:
                    font = copy(cell.font)
                    font.color = None if rgb is None else "FF" + rgb
                    self.set_style(sheet, r, c, font=font)

    def set_col_width(self, sheet: Sheet, c: int, width: float | None, hidden: bool | None = None):
        from .structure import split_col_dims

        prev = sheet.col_width(c)
        new = (prev[0] if width is None else width, prev[1] if hidden is None else hidden)

        def apply(state):
            split_col_dims(sheet.ws, c)
            dim = sheet.ws.column_dimensions[col_letter(c)]
            dim.width = state[0]
            dim.hidden = state[1]
            sheet.invalidate_layout()
            self.modified = True

        apply(new)
        self.undo.push(UndoEntry("Column width", lambda: apply(prev), lambda: apply(new),
                                 (sheet.name, Rect(1, c, 1, c))))

    # ------------------------------------------------------------------ clipboard / fill

    def capture(self, sheet: Sheet, rect: Rect) -> dict:
        rect = self.clip(sheet, rect)
        cells = {}
        for (r, c), cell in sheet.ws._cells.items():
            if rect.contains(r, c) and not isinstance(cell, MergedCell):
                cells[(r - rect.r1, c - rect.c1)] = (cell._value, cell.data_type, copy(cell._style))
        if sheet.spill.anchors:
            # Spilled values are pasted as values, except where the formula itself is copied too
            # (then it spills again at the destination): those are marked "spill" and only pasted
            # by Paste Values.
            for r, c, anchor in sheet.spill.cells_in(rect.r1, rect.c1, rect.r2, rect.c2):
                v = sheet.spill.value(anchor, r, c)
                if v is None:
                    continue
                cell = sheet.ws._cells.get((r, c))
                style = copy(cell._style) if cell is not None else StyleArray()
                if rect.contains(*anchor):
                    dtype = "spill"
                elif isinstance(v, ExcelError):
                    v, dtype = v.code, "e"
                else:
                    dtype = "s" if isinstance(v, str) else None
                cells[(r - rect.r1, c - rect.c1)] = (v, dtype, style)
        return {"sheet": sheet, "rect": rect, "cells": cells}

    def paste(self, clip: dict, sheet: Sheet, r0: int, c0: int, cut: bool = False,
              values_only: bool = False, tile_to: Rect | None = None) -> Rect:
        from ..formula.parser import translate

        src: Rect = clip["rect"]
        h, w = src.rows, src.cols
        reps_r = reps_c = 1
        if tile_to is not None and tile_to.rows % h == 0 and tile_to.cols % w == 0:
            reps_r, reps_c = tile_to.rows // h, tile_to.cols // w
        dest = Rect(r0, c0, r0 + h * reps_r - 1, c0 + w * reps_c - 1)
        if dest.r2 > MAX_ROW or dest.c2 > MAX_COL:
            raise ValueError("Paste area extends beyond the sheet")
        same_wb = clip["sheet"].book is self
        with self.transaction("Cut" if cut else "Paste", (sheet.name, dest)):
            src_sheet: Sheet = clip["sheet"]
            if cut and src_sheet.book is self:
                for (dr, dc) in clip["cells"]:
                    r, c = src.r1 + dr, src.c1 + dc
                    if not dest.contains(r, c) or src_sheet is not sheet:
                        self._touch(src_sheet, r, c)
                        self._unregister(src_sheet, r, c)
                        src_sheet.ws._cells.pop((r, c), None)
            for i in range(reps_r):
                for j in range(reps_c):
                    br, bc = r0 + i * h, c0 + j * w
                    for dr in range(h):
                        for dc in range(w):
                            r, c = br + dr, bc + dc
                            data = clip["cells"].get((dr, dc))
                            if data is not None and data[1] == "spill":
                                if values_only:
                                    v = data[0]
                                    if isinstance(v, ExcelError):
                                        self.set_raw(sheet, r, c, v.code, "e")
                                    else:
                                        self.set_raw(sheet, r, c, v, "s" if isinstance(v, str) else None)
                                    continue
                                data = None
                            if data is None:
                                if sheet.ws._cells.get((r, c)) is not None:
                                    self.set_raw(sheet, r, c, None)
                                    if not values_only and same_wb:
                                        self.set_style(sheet, r, c, _style=StyleArray())
                                continue
                            value, dtype, style = data
                            if dtype == "f" and isinstance(value, (str, ArrayFormula)):
                                text = _formula_text(value)
                                if values_only:
                                    v = self.get_value(src_sheet, src.r1 + dr, src.c1 + dc)
                                    if isinstance(v, ExcelError):
                                        self.set_raw(sheet, r, c, v.code, "e")
                                    else:
                                        self.set_raw(sheet, r, c, v)
                                elif cut:
                                    self.set_raw(sheet, r, c, text)
                                else:
                                    self.set_raw(sheet, r, c, translate(text, r - (src.r1 + dr), c - (src.c1 + dc)))
                            else:
                                self.set_raw(sheet, r, c, value, dtype if dtype in ("s", "e") else None)
                            if not values_only and same_wb:
                                cell = self.writable_cell(sheet, r, c)
                                if cell is not None:
                                    self.set_style(sheet, r, c, _style=copy(style))
            if cut and src_sheet.book is self:
                self._move_refs(src_sheet, src, sheet, dest.r1 - src.r1, dest.c1 - src.c1)
        return dest

    def _move_refs(self, src_sheet: Sheet, src: Rect, dst_sheet: Sheet, dr: int, dc: int):
        """After a cut/paste, point references to the moved block at its new location."""
        from ..formula.parser import RefInfo, transform_refs
        from .refs import quote_sheet

        for s in self.sheets:
            for (r, c), f in list(s.formulas.items()):
                def fn(info: RefInfo, s=s):
                    target = s if info.sheet is None else self.sheet_by_name(info.sheet)
                    if target is not src_sheet or info.kind not in ("cell", "range"):
                        return info
                    if not (src.contains(info.r1, info.c1) and src.contains(info.r2, info.c2)):
                        return info
                    new = RefInfo(**info.__dict__)
                    new.r1 += dr
                    new.r2 += dr
                    new.c1 += dc
                    new.c2 += dc
                    if dst_sheet is not s or info.sheet is not None:
                        if dst_sheet is not target:
                            new.sheet = dst_sheet.name
                            new.sheet_raw = quote_sheet(dst_sheet.name) + "!"
                    return new

                text = f.text
                new_text = transform_refs(text, fn)
                if new_text != text:
                    self.set_raw(s, r, c, to_storage(new_text))

    def fill(self, sheet: Sheet, rect: Rect, direction: str):
        """Ctrl+D / Ctrl+R: copy the first row/column of rect across the rest."""
        rect = self.clip(sheet, rect)
        if direction == "down":
            if rect.rows < 2:
                return
            clip = self.capture(sheet, Rect(rect.r1, rect.c1, rect.r1, rect.c2))
            clip["rect"] = Rect(rect.r1, rect.c1, rect.r1, rect.c2)
            with self.transaction("Fill down", (sheet.name, rect)):
                for r in range(rect.r1 + 1, rect.r2 + 1):
                    self.paste(clip, sheet, r, rect.c1)
        else:
            if rect.cols < 2:
                return
            clip = self.capture(sheet, Rect(rect.r1, rect.c1, rect.r2, rect.c1))
            clip["rect"] = Rect(rect.r1, rect.c1, rect.r2, rect.c1)
            with self.transaction("Fill right", (sheet.name, rect)):
                for c in range(rect.c1 + 1, rect.c2 + 1):
                    self.paste(clip, sheet, rect.r1, c)

    def paste_text(self, sheet: Sheet, r0: int, c0: int, text: str) -> Rect:
        """Paste tab/newline separated text (e.g. from the system clipboard)."""
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        rows = [ln.split("\t") for ln in lines]
        h = len(rows)
        w = max((len(r) for r in rows), default=1)
        dest = Rect(r0, c0, r0 + max(h, 1) - 1, c0 + w - 1)
        with self.transaction("Paste", (sheet.name, dest)):
            for i, row in enumerate(rows):
                for j, val in enumerate(row):
                    self._apply_parsed(sheet, r0 + i, c0 + j, parse_input(val, self.epoch1904))
        return dest

    def text_of(self, sheet: Sheet, rect: Rect) -> str:
        """Formatted values of a range as TSV (for the system clipboard)."""
        from ..format.numfmt import format_value

        rect = self.clip(sheet, rect)
        lines = []
        for r in range(rect.r1, rect.r2 + 1):
            vals = []
            for c in range(rect.c1, rect.c2 + 1):
                v = self.get_value(sheet, r, c)
                vals.append(format_value(v, sheet.number_format(r, c), self.epoch1904)[0] if v is not None else "")
            lines.append("\t".join(vals))
        return "\n".join(lines)

    # ------------------------------------------------------------------ undo/redo

    def do_undo(self):
        e = self.undo.undo()
        if e is not None:
            self.modified = True
        return e

    def do_redo(self):
        e = self.undo.redo()
        if e is not None:
            self.modified = True
        return e

    def undo_goto(self, node) -> list:
        """Move to any state in the undo tree; returns the entries undone/redone on the way."""
        done = self.undo.goto(node)
        if done:
            self.modified = True
        return done


def _formula_text(value) -> str:
    if isinstance(value, ArrayFormula):
        t = value.text or ""
        return t if t.startswith("=") else "=" + t
    return value if isinstance(value, str) else str(value)


def _first_ref(node):
    if type(node) is RefNode:
        return node
    if type(node) is Binary and node.op in ("+", "-", "*", "/"):
        return _first_ref(node.left) or (_first_ref(node.right) if node.op in ("+", "-") else None)
    if type(node) is Func and node.name in _FORMAT_INHERITING:
        for a in node.args:
            r = _first_ref(a)
            if r is not None:
                return r
    return None

