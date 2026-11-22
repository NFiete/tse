"""Structural edits: insert/delete rows and columns, add/rename/delete sheets."""
from __future__ import annotations

from copy import copy

from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.formula import ArrayFormula

from ..formula.parser import RefInfo, to_display, to_storage, transform_refs, transform_refs_tokens
from . import tables
from .refs import MAX_COL, MAX_ROW, Rect, col_index, col_letter, quote_sheet
from .gcutil import paused_gc
from .undo import UndoEntry


def split_col_dims(ws, c: int):
    """Make sure column c has its own ColumnDimension (openpyxl groups ranges like B:D)."""
    isolate_cols(ws, c, c)


def isolate_cols(ws, c1: int, c2: int):
    """Give every column in c1..c2 its own ColumnDimension. A grouped dimension such as B:XFD is
    split into the part before the range, one per column inside it, and the part after."""
    for key, dim in list(ws.column_dimensions.items()):
        lo, hi = dim.min, dim.max
        if not lo or not hi or lo == hi or hi < c1 or lo > c2:
            continue
        del ws.column_dimensions[key]
        if lo < c1:
            _clone_coldim(ws, dim, lo, c1 - 1)
        for cc in range(max(lo, c1), min(hi, c2) + 1):
            _clone_coldim(ws, dim, cc)
        if hi > c2:
            _clone_coldim(ws, dim, c2 + 1, hi)
    for cc in range(c1, c2 + 1):
        letter = col_letter(cc)
        if letter not in ws.column_dimensions:
            # width 0 means "default width" and isn't written to the file
            ws.column_dimensions[letter] = ColumnDimension(ws, index=letter, width=0, min=cc, max=cc)


def _clone_coldim(ws, dim, lo: int, hi: int | None = None):
    hi = lo if hi is None else hi
    letter = col_letter(lo)
    nd = ColumnDimension(ws, index=letter, width=dim.width, bestFit=dim.bestFit, hidden=dim.hidden,
                         outlineLevel=dim.outlineLevel, collapsed=dim.collapsed, min=lo, max=hi)
    nd._style = copy(dim._style)
    ws.column_dimensions[letter] = nd
    return nd


def _explode_col_dims(ws) -> dict[int, ColumnDimension]:
    out = {}
    for key, dim in list(ws.column_dimensions.items()):
        lo, hi = dim.min, dim.max
        if not lo or not hi:
            try:
                lo = hi = col_index(key)
            except Exception:
                continue
        for cc in range(lo, min(hi, MAX_COL) + 1):
            out[cc] = dim
    return out


def _rebuild_col_dims(ws, mapping: dict[int, ColumnDimension]):
    ws.column_dimensions.clear()
    for cc in sorted(mapping):
        if 1 <= cc <= MAX_COL:
            _clone_coldim(ws, mapping[cc], cc)


# --------------------------------------------------------------------------- reference shifting


def _formula_of(cell) -> str | None:
    """Storage text of a cell's formula (plain or legacy array formula), or None."""
    v = cell._value
    if isinstance(v, str):
        return v if cell.data_type == "f" else None
    if isinstance(v, ArrayFormula):
        t = v.text or ""
        return t if t.startswith("=") else "=" + t
    return None


def _with_text(value, text: str):
    """A cell value like `value` holding formula text `text`. Array formulas get a new object:
    undo snapshots may still share the old one."""
    return ArrayFormula(value.ref, text) if isinstance(value, ArrayFormula) else text


def _shift_fn(axis: str, idx: int, n: int):
    """Coordinate mapping for inserting (n>0) or deleting (n<0) rows/cols starting at idx.

    Returns f(lo, hi) -> (lo, hi) or None if the whole span was deleted.
    """
    if n > 0:
        def f(lo, hi):
            return (lo + n if lo >= idx else lo), (hi + n if hi >= idx else hi)
    else:
        cnt = -n
        end = idx + cnt - 1

        def f(lo, hi):
            if lo >= idx and hi <= end:
                return None
            nlo = lo if lo < idx else (idx if lo <= end else lo - cnt)
            nhi = hi if hi < idx else (idx - 1 if hi <= end else hi - cnt)
            if nhi < nlo:
                return None
            return nlo, nhi
    return f


def _transform_formula(text: str, formula_sheet, target_sheet, book, axis: str, f, with_tokens: bool = False,
                       damage: list | None = None):
    def fn(info: RefInfo):
        res = _fn(info)
        if damage is not None and res is not info:
            if res is None:
                damage.append(True)
            else:
                span_old = (info.r2 - info.r1) if axis == "row" else (info.c2 - info.c1)
                span_new = (res.r2 - res.r1) if axis == "row" else (res.c2 - res.c1)
                if span_old != span_new:
                    damage.append(True)
        return res

    def _fn(info: RefInfo):
        s = formula_sheet if info.sheet is None else book.sheet_by_name(info.sheet)
        if s is not target_sheet:
            return info
        if axis == "row":
            if info.kind == "col":
                return info
            res = f(info.r1, info.r2)
            if res is None:
                return None
            if res == (info.r1, info.r2):
                return info
            new = RefInfo(**info.__dict__)
            new.r1, new.r2 = res
            if info.kind == "cell":
                if res[0] != res[1]:
                    return None
            return new
        if info.kind == "row":
            return info
        res = f(info.c1, info.c2)
        if res is None:
            return None
        if res == (info.c1, info.c2):
            return info
        new = RefInfo(**info.__dict__)
        new.c1, new.c2 = res
        if info.kind == "cell" and res[0] != res[1]:
            return None
        return new

    if with_tokens:
        return transform_refs_tokens(text, fn)
    return transform_refs(text, fn)


def _shift_multicell(mcr: MultiCellRange, axis: str, f) -> MultiCellRange:
    out = []
    for cr in mcr.ranges:
        if axis == "row":
            res = f(cr.min_row, cr.max_row)
            if res is None:
                continue
            out.append(CellRange(min_col=cr.min_col, min_row=res[0], max_col=cr.max_col, max_row=res[1]))
        else:
            res = f(cr.min_col, cr.max_col)
            if res is None:
                continue
            out.append(CellRange(min_col=res[0], min_row=cr.min_row, max_col=res[1], max_row=cr.max_row))
    return MultiCellRange(out)


# --------------------------------------------------------------------------- insert / delete


def _move_cells(sheet, axis: str, idx: int, n: int) -> dict:
    """Shift cells; returns snapshots of deleted cells keyed by original position."""
    ws = sheet.ws
    deleted = {}
    ai = 0 if axis == "row" else 1
    if n < 0:
        cnt = -n
        end = idx + cnt - 1
        for key in [k for k in ws._cells if idx <= k[ai] <= end]:
            cell = ws._cells.pop(key)
            if not isinstance(cell, MergedCell):
                deleted[key] = (cell._value, cell.data_type, copy(cell._style))
    keys = sorted((k for k in ws._cells if k[ai] >= idx), key=lambda k: k[ai], reverse=n > 0)
    moved = {}
    for key in keys:
        cell = ws._cells.pop(key)
        nr, nc = (key[0] + n, key[1]) if axis == "row" else (key[0], key[1] + n)
        cell.row, cell.column = nr, nc
        moved[(nr, nc)] = cell
        if isinstance(cell._value, ArrayFormula):
            cell._value = _moved_array(cell._value, nr, nc)
        if getattr(cell, "_hyperlink", None) is not None:
            cell._hyperlink.ref = cell.coordinate
    ws._cells.update(moved)

    def remap(d: dict):
        out = {}
        for (r, c), v in d.items():
            k = (r, c)[ai]
            if n < 0 and idx <= k <= idx - n - 1:
                continue
            if k >= idx:
                out[(r + n, c) if axis == "row" else (r, c + n)] = v
            else:
                out[(r, c)] = v
        return out

    sheet.values = remap(sheet.values)
    sheet.cached = remap(sheet.cached)
    sheet.stale = set(remap({k: 1 for k in sheet.stale}))
    return deleted


def _moved_array(value: ArrayFormula, r: int, c: int) -> ArrayFormula:
    """A legacy array formula whose anchor cell moved to (r, c): its range moves along."""
    from openpyxl.utils.cell import range_boundaries

    try:
        c1, r1, c2, r2 = range_boundaries(value.ref)
    except (TypeError, ValueError):
        c1 = r1 = c2 = r2 = 1
    ref = f"{col_letter(c)}{r}"
    if (r2, c2) != (r1, c1):
        ref += f":{col_letter(c + c2 - c1)}{r + r2 - r1}"
    return ArrayFormula(ref, value.text)


def _copy_formats(sheet, axis: str, src: int, idx: int, n: int):
    """New rows/columns take the formatting of the one above/left, like Excel."""
    ws = sheet.ws
    ai = 0 if axis == "row" else 1
    for key, cell in list(ws._cells.items()):
        if key[ai] != src or isinstance(cell, MergedCell) or cell._style is None or not any(cell._style):
            continue
        for k in range(n):
            r, c = (idx + k, key[1]) if axis == "row" else (key[0], idx + k)
            new = ws.cell(row=r, column=c)
            new._style = copy(cell._style)


def _shift_dims(sheet, axis: str, idx: int, n: int) -> dict:
    ws = sheet.ws
    removed = {}
    if axis == "row":
        dims = ws.row_dimensions
        items = sorted(dims.items(), key=lambda kv: kv[0])
        new = {}
        for r, d in items:
            if n < 0 and idx <= r <= idx - n - 1:
                removed[r] = d
                continue
            nr = r + n if r >= idx else r
            d.index = nr
            new[nr] = d
        dims.clear()
        dims.update(new)
        sheet.invalidate_layout()  # row outline levels moved
    else:
        mapping = _explode_col_dims(ws)
        new = {}
        for c, d in mapping.items():
            if n < 0 and idx <= c <= idx - n - 1:
                removed[c] = d
                continue
            new[c + n if c >= idx else c] = d
        _rebuild_col_dims(ws, new)
        sheet.invalidate_layout()
    return removed


def _shift_merged(sheet, axis: str, idx: int, n: int) -> list:
    ws = sheet.ws
    f = _shift_fn(axis, idx, n)
    removed = []
    for mr in list(ws.merged_cells.ranges):
        lo, hi = (mr.min_row, mr.max_row) if axis == "row" else (mr.min_col, mr.max_col)
        old = mr.coord
        res = f(lo, hi)
        anchor_deleted = n < 0 and idx <= lo <= idx - n - 1
        if res is None or anchor_deleted or (res[0] == res[1] and (
                (mr.min_col == mr.max_col) if axis == "row" else (mr.min_row == mr.max_row))):
            ws.merged_cells.remove(mr)
            removed.append(old)
            # drop leftover placeholders
            for key in list(ws._cells):
                cell = ws._cells[key]
                if isinstance(cell, MergedCell) and not any(
                        m.min_row <= key[0] <= m.max_row and m.min_col <= key[1] <= m.max_col
                        for m in ws.merged_cells.ranges):
                    del ws._cells[key]
            continue
        if axis == "row":
            mr.min_row, mr.max_row = res
        else:
            mr.min_col, mr.max_col = res
        if n > 0 and lo < idx <= hi:
            ws._clean_merge_range(mr)
        if mr.coord != old:
            removed.append(None)
    return removed


def _shift_misc(sheet, axis: str, idx: int, n: int):
    ws = sheet.ws
    f = _shift_fn(axis, idx, n)
    try:
        for dv in ws.data_validations.dataValidation:
            dv.sqref = _shift_multicell(dv.sqref, axis, f)
        ws.data_validations.dataValidation = [dv for dv in ws.data_validations.dataValidation if dv.sqref.ranges]
    except AttributeError:
        pass
    try:
        cfl = ws.conditional_formatting
        rules = list(cfl._cf_rules.items())
        cfl._cf_rules.clear()
        for cf, rl in rules:
            new = _shift_multicell(cf.sqref, axis, f)
            if new.ranges:
                cf.sqref = new
                cfl._cf_rules[cf] = rl
    except AttributeError:
        pass
    fp = ws.freeze_panes
    if fp:
        from .refs import parse_addr

        p = parse_addr(fp)
        if p:
            r, c = p
            if axis == "row" and r > idx and n > 0:
                ws.freeze_panes = f"{col_letter(c)}{r + n}"
            elif axis == "col" and c > idx and n > 0:
                ws.freeze_panes = f"{col_letter(c + n)}{r}"
            elif n < 0:
                res = f(r, r) if axis == "row" else f(c, c)
                if res:
                    ws.freeze_panes = f"{col_letter(c)}{max(res[0], 1)}" if axis == "row" else \
                        f"{col_letter(max(res[0], 1))}{r}"


def _shift_sheet_settings(book, target_sheet, axis: str, f):
    """References outside cells: data-validation and conditional-format formulas on every sheet
    (they may point at the target sheet), and the target sheet's auto-filter, print titles and
    print area. (The ranges validations and formats apply to are moved by _shift_misc.)"""

    def tf(text, s):
        return _transform_formula("=" + text, s, target_sheet, book, axis, f)[1:]

    for s in book.sheets:
        ws = s.ws
        for dv in ws.data_validations.dataValidation:
            if dv.formula1:
                dv.formula1 = tf(dv.formula1, s)
            if dv.formula2:
                dv.formula2 = tf(dv.formula2, s)
        for rules in ws.conditional_formatting._cf_rules.values():
            for rule in rules:
                if rule.formula:
                    rule.formula = [tf(x, s) for x in rule.formula]
    ws = target_sheet.ws
    af = ws.auto_filter
    if af.ref:
        new = _shift_multicell(MultiCellRange(af.ref), axis, f)
        af.ref = next(iter(new.ranges)).coord if new.ranges else None
    if ws._print_area:
        new = _shift_multicell(ws._print_area, axis, f)
        ws.print_area = [r.coord for r in new.ranges] or None
    if axis == "row" and ws._print_rows:
        res = f(int(ws._print_rows.min_row), int(ws._print_rows.max_row))
        _set_print_titles(ws, None if res is None else f"{res[0]}:{res[1]}", ws.print_title_cols)
    elif axis == "col" and ws._print_cols:
        res = f(col_index(ws._print_cols.min_col), col_index(ws._print_cols.max_col))
        _set_print_titles(ws, ws.print_title_rows,
                          None if res is None else f"{col_letter(res[0])}:{col_letter(res[1])}")


def _set_print_titles(ws, rows: str | None, cols: str | None):
    # openpyxl's setters ignore None, so they can't remove titles
    from openpyxl.worksheet.print_settings import ColRange, RowRange

    ws._print_rows = RowRange(rows.replace("$", "")) if rows else None
    ws._print_cols = ColRange(cols.replace("$", "")) if cols else None


def _settings_state(book) -> dict:
    """Snapshot of what _shift_sheet_settings and _shift_misc change, for undoing a deletion."""
    from copy import deepcopy

    out = {}
    for s in book.sheets:
        ws = s.ws
        out[s] = (deepcopy(ws.data_validations.dataValidation), deepcopy(ws.conditional_formatting._cf_rules),
                  ws.auto_filter.ref, [r.coord for r in ws._print_area.ranges] if ws._print_area else None,
                  ws.print_title_rows, ws.print_title_cols)
    return out


def _restore_settings(book, state: dict):
    for s, (dvs, cfs, af, area, rows, cols) in state.items():
        ws = s.ws
        ws.data_validations.dataValidation = dvs
        ws.conditional_formatting._cf_rules = cfs
        ws.auto_filter.ref = af
        ws.print_area = area
        _set_print_titles(ws, rows, cols)


def _affected(refs, target_sheet, axis: str, idx: int) -> bool:
    """Could a formula with these (resolved) refs need rewriting for a change at idx?"""
    if axis == "row":
        return any(ref[0] is target_sheet and ref[3] >= idx for ref in refs)
    return any(ref[0] is target_sheet and ref[4] >= idx for ref in refs)


def _rewrite_all(book, target_sheet, axis: str, idx: int, f) -> tuple[dict, bool]:
    """Rewrite formula text before cells move.

    Returns ({(sheet, r, c): old_storage_text} in pre-move positions, names_changed). Formulas
    left with #REF! lose their saved value; book._damaged_pre keeps (old value, saved value).
    Formulas whose references can't be affected are skipped without re-tokenizing.
    """
    changed = {}
    book._pending_tokens = {}
    book._damaged = set()
    book._damaged_pre = {}
    precedents = book.graph.precedents
    for s in book.sheets:
        formulas = s.formulas
        cells = s.ws._cells
        for rc, fobj in formulas.items():
            refs = precedents.get((s, rc[0], rc[1]))
            if fobj.ast is not None and refs is not None and not _affected(refs, target_sheet, axis, idx):
                continue
            cell = cells.get(rc)
            old = None if cell is None else _formula_of(cell)
            if old is None:
                continue
            damage: list = []
            new, toks = _transform_formula(old, s, target_sheet, book, axis, f, with_tokens=True, damage=damage)
            if new != old:
                old = cell._value
                cell._value = _with_text(old, new)
                changed[(s, rc[0], rc[1])] = old
                if toks is not None:
                    book._pending_tokens[cell] = toks
                if damage:
                    # A cut reference makes Excel's saved value wrong: drop it (undo restores it)
                    book._damaged.add(cell)
                    book._damaged_pre[(s, rc[0], rc[1])] = (old, s.cached.pop(rc, None))
    names_changed = False
    for dn in _defined_names(book):
        txt = "=" + (dn.attr_text or "")
        new = _transform_formula(txt, None, target_sheet, book, axis, f)
        if new != txt:
            dn.attr_text = new[1:]
            names_changed = True
    return changed, names_changed


def _defined_names(book) -> list:
    """Every defined name: workbook-level ones and those scoped to a sheet."""
    out = list(book.wb.defined_names.values())
    for s in book.sheets:
        out.extend(s.ws.defined_names.values())
    return out


def _post_key(sheet, axis, idx, n, key):
    s, r, c = key
    if s is not sheet:
        return key
    if axis == "row":
        return (s, r + n, c) if r >= idx else key
    return (s, r, c + n) if c >= idx else key


def _reindex(book, sheet, axis: str, idx: int, n: int, changed_pre: dict, names_changed: bool) -> set:
    """Move formula registrations along with their cells; re-parse only changed formulas.

    Returns the post-move keys of formulas whose text changed."""
    from .workbook import _formula_text

    if names_changed:
        book._load_names()
        book.register_all()
        return {_post_key(sheet, axis, idx, n, k) for k in changed_pre}
    graph = book.graph
    ai = 0 if axis == "row" else 1
    formulas = sheet.formulas
    moved = []
    for rc in [rc for rc in formulas if rc[ai] >= idx]:
        key = (sheet, rc[0], rc[1])
        refs = graph.precedents.get(key)
        fobj = formulas.pop(rc)
        graph.remove(key)
        if n < 0 and rc[ai] <= idx - n - 1:
            continue  # deleted
        new_rc = (rc[0] + n, rc[1]) if axis == "row" else (rc[0], rc[1] + n)
        moved.append((key, new_rc, fobj, refs))
    changed_post = set()
    for key, new_rc, fobj, refs in moved:
        if key in changed_pre or refs is None:
            cell = sheet.ws._cells.get(new_rc)
            if cell is not None and cell._value is not None:
                book._register(sheet, new_rc[0], new_rc[1], _formula_text(cell._value),
                               array=isinstance(cell._value, ArrayFormula), tokens=book._pending_tokens.get(cell))
            changed_post.add((sheet, new_rc[0], new_rc[1]))
        else:
            formulas[new_rc] = fobj
            graph.add((sheet, new_rc[0], new_rc[1]), refs, fobj.volatile)
    for key in changed_pre:
        s, r, c = key
        if s is sheet and (r, c)[ai] >= idx:
            continue
        cell = s.ws._cells.get((r, c))
        value = s.values.get((r, c))
        book._unregister(s, r, c, keep_spill=value is not None)
        if cell is not None and cell._value is not None:
            book._register(s, r, c, _formula_text(cell._value), array=isinstance(cell._value, ArrayFormula),
                           tokens=book._pending_tokens.get(cell))
            if value is not None:
                s.values[(r, c)] = value  # shifting references doesn't change the result
        changed_post.add(key)
    return changed_post


def _dirty_after(book, sheet, axis: str, idx: int, n: int, changed_post: set) -> set:
    """Formulas whose value can change: ranges spanning the edit, ROW()/COLUMN() users that
    moved, and (for deletions) anything whose references changed."""
    dirty = set()
    lo = idx
    for fkey, refs in book.graph.precedents.items():
        for ref in refs:
            if ref[0] is not sheet:
                continue
            a, b = (ref[1], ref[3]) if axis == "row" else (ref[2], ref[4])
            if a < lo <= b or (n > 0 and a <= lo + n - 1 and b >= lo and a != b):
                dirty.add(fkey)
                break
    ai = 1 if axis == "row" else 2
    for rc, fobj in sheet.formulas.items():
        if fobj.positional and (rc[ai - 1]) >= idx:
            dirty.add((sheet, rc[0], rc[1]))
    if n < 0 and book._damaged:
        damaged = book._damaged
        for key in changed_post:
            s, r, c = key
            if s.ws._cells.get((r, c)) in damaged:
                dirty.add(key)
    return dirty


def _apply(book, sheet, axis: str, idx: int, n: int, table_fill: bool = True, copy_formats: bool = True):
    """Shared implementation of insert (n > 0) and delete (n < 0)."""
    with paused_gc():
        return _apply_inner(book, sheet, axis, idx, n, table_fill, copy_formats)


def _apply_inner(book, sheet, axis: str, idx: int, n: int, table_fill: bool = True, copy_formats: bool = True):
    f = _shift_fn(axis, idx, n)
    has_tables = bool(book.table_index().by_name)
    spill_anchors = book.take_spills(sheet)  # recalculated below, once cells have moved
    changed_pre, names_changed = _rewrite_all(book, sheet, axis, idx, f)
    deleted = _move_cells(sheet, axis, idx, n)
    removed_dims = _shift_dims(sheet, axis, idx, n)
    if n > 0 and idx > 1 and copy_formats:
        _copy_formats(sheet, axis, idx - 1, idx, n)
    _shift_merged(sheet, axis, idx, n)
    _shift_misc(sheet, axis, idx, n)
    _shift_sheet_settings(book, sheet, axis, f)
    new_formulas = tables.shift(book, sheet, axis, idx, n, fill=table_fill)
    changed_post = _reindex(book, sheet, axis, idx, n, changed_pre, names_changed)
    book._pending_tokens = {}
    sheet.recompute_bounds()
    dirty = _dirty_after(book, sheet, axis, idx, n, changed_post)
    ai = 0 if axis == "row" else 1
    for rc in spill_anchors:
        if n < 0 and idx <= rc[ai] <= idx - n - 1:
            continue  # deleted
        key = _post_key(sheet, axis, idx, n, (sheet, rc[0], rc[1]))
        if (key[1], key[2]) in sheet.formulas:
            dirty.add(key)
    if has_tables:
        from .workbook import _formula_text

        for r, c in new_formulas:
            cell = sheet.ws._cells.get((r, c))
            if cell is not None and cell.data_type == "f":
                book._register(sheet, r, c, _formula_text(cell._value))
                dirty.add((sheet, r, c))
        dirty |= book.reregister_structured()
    damaged_pre = book._damaged_pre
    book._damaged = set()
    book._damaged_pre = {}
    if dirty:
        dirty |= book.graph.all_dependents(dirty)
        book._run(dirty)
    book.modified = True
    return damaged_pre, deleted, removed_dims


def insert(book, sheet, axis: str, idx: int, n: int = 1, record: bool = True, table_fill: bool = True,
           copy_formats: bool = True):
    """Insert n empty rows (axis='row') or columns before index idx.

    New rows inside a table get its calculated-column formulas unless table_fill is False, and
    the formatting of the row/column before them unless copy_formats is False."""
    limit = MAX_ROW if axis == "row" else MAX_COL
    used = sheet.max_row if axis == "row" else sheet.max_col
    if used + n > limit:
        raise ValueError("That would push data off the end of the sheet")
    _apply(book, sheet, axis, idx, n, table_fill, copy_formats)
    if record:
        where = (sheet.name, Rect(idx, 1, idx + n - 1, MAX_COL) if axis == "row" else Rect(1, idx, MAX_ROW, idx + n - 1))
        book.undo.push(UndoEntry(f"Insert {axis}s", lambda: delete(book, sheet, axis, idx, n, record=False),
                                 lambda: insert(book, sheet, axis, idx, n, record=False), where))


def delete(book, sheet, axis: str, idx: int, n: int = 1, record: bool = True):
    """Delete n rows/columns starting at idx."""
    if record:
        tables.check_delete(book, sheet, axis, idx, n)
    old_tables = tables.snapshot(sheet)
    old_merged = [mr.coord for mr in sheet.ws.merged_cells.ranges]
    old_settings = _settings_state(book) if record else None
    ai = 0 if axis == "row" else 1
    old_cached = {rc: v for rc, v in sheet.cached.items() if idx <= rc[ai] < idx + n} if record else None
    # Only references that were cut (not merely shifted) need restoring on undo:
    # re-inserting the rows shifts everything else back exactly.
    old_texts, deleted, removed_dims = _apply(book, sheet, axis, idx, -n)
    if not record:
        return

    def undo():
        from .workbook import _formula_text

        insert(book, sheet, axis, idx, n, record=False, table_fill=False, copy_formats=False)
        tables.restore(book, sheet, old_tables)
        _restore_settings(book, old_settings)
        ws = sheet.ws
        restored = set()
        for (r, c), (value, dtype, style) in deleted.items():
            cell = ws.cell(row=r, column=c)
            cell._value, cell.data_type, cell._style = value, dtype, copy(style)
            restored.add((sheet, r, c))
        sheet.cached.update(old_cached)
        for (s, r, c), (text, cached) in old_texts.items():
            cell = s.ws._cells.get((r, c))
            if cell is not None:
                cell._value = text
            if cached is not None:
                s.cached[(r, c)] = cached
            restored.add((s, r, c))
        if axis == "row":
            for r, d in removed_dims.items():
                d.index = r
                ws.row_dimensions[r] = d
        else:
            mapping = _explode_col_dims(ws)
            mapping.update(removed_dims)
            _rebuild_col_dims(ws, mapping)
            sheet.invalidate_layout()
        current = {mr.coord for mr in ws.merged_cells.ranges}
        for coord in old_merged:
            if coord not in current:
                cr = CellRange(coord)
                for mr in list(ws.merged_cells.ranges):
                    if mr.coord != coord and not mr.isdisjoint(cr):
                        ws.merged_cells.remove(mr)
                try:
                    ws.merge_cells(coord)
                except Exception:
                    pass
        for s, r, c in restored:
            book._unregister(s, r, c)
            cell = s.ws._cells.get((r, c))
            if cell is not None and cell.data_type == "f" and cell._value is not None:
                book._register(s, r, c, _formula_text(cell._value), array=isinstance(cell._value, ArrayFormula))
        sheet.recompute_bounds()
        book.tables_changed(restored)

    where = (sheet.name, Rect(idx, 1, idx + n - 1, MAX_COL) if axis == "row" else Rect(1, idx, MAX_ROW, idx + n - 1))
    book.undo.push(UndoEntry(f"Delete {axis}s", undo, lambda: delete(book, sheet, axis, idx, n, record=False), where))


# --------------------------------------------------------------------------- sheets

_BAD_SHEET_CHARS = set('[]:*?/\\')


def validate_sheet_name(book, name: str, exclude=None) -> str | None:
    name = name.strip()
    if not name:
        return "Sheet name can't be empty"
    if len(name) > 31:
        return "Sheet names are limited to 31 characters"
    if any(ch in _BAD_SHEET_CHARS for ch in name):
        return "Sheet names can't contain [ ] : * ? / \\"
    if name.startswith("'") or name.endswith("'"):
        return "Sheet names can't start or end with an apostrophe"
    other = book.sheet_by_name(name)
    if other is not None and other is not exclude:
        return f"A sheet named '{name}' already exists"
    return None


def add_sheet(book, name: str | None = None, index: int | None = None, record: bool = True):
    from .workbook import Sheet

    if name is None:
        i = len(book.sheets) + 1
        while book.sheet_by_name(f"Sheet{i}") is not None:
            i += 1
        name = f"Sheet{i}"
    err = validate_sheet_name(book, name)
    if err:
        raise ValueError(err)
    pos = len(book.wb._sheets) if index is None else book.wb._sheets.index(book.sheets[index].ws) if index < len(
        book.sheets) else len(book.wb._sheets)
    ws = book.wb.create_sheet(name, pos)
    sheet = Sheet(book, ws)
    book.sheets = [s for s in [*book.sheets, sheet]]
    book.sheets.sort(key=lambda s: book.wb.worksheets.index(s.ws))
    # formulas that referenced this name may now resolve
    _reresolve(book)
    book.modified = True
    if record:
        book.undo.push(UndoEntry("Add sheet", lambda: remove_sheet(book, sheet, record=False),
                                 lambda: _readd(book, sheet), (sheet.name, Rect(1, 1, 1, 1))))
    return sheet


def _readd(book, sheet):
    book.wb._sheets.insert(len(book.wb._sheets), sheet.ws)
    book.sheets.append(sheet)
    book.sheets.sort(key=lambda s: book.wb.worksheets.index(s.ws))
    _reresolve(book)


def remove_sheet(book, sheet, record: bool = True):
    """Delete a sheet. References to it become #REF! (not undoable when record=True)."""
    if len(book.sheets) <= 1:
        raise ValueError("A workbook must keep at least one sheet")
    if record:
        def fn_factory(s):
            def fn(info: RefInfo):
                target = s if info.sheet is None else book.sheet_by_name(info.sheet)
                return None if target is sheet else info
            return fn

        for s in book.sheets:
            if s is sheet:
                continue
            for (r, c), cell in list(s.ws._cells.items()):
                text = _formula_of(cell)
                if text is not None:
                    new = transform_refs(to_display(text), fn_factory(s))
                    if new != to_display(text):
                        cell._value = _with_text(cell._value, to_storage(new))
    book.graph.remove_sheet(sheet)
    book.wb.remove(sheet.ws)
    book.sheets.remove(sheet)
    if record:
        book.undo.clear()
    _reresolve(book)
    book.modified = True


def rename_sheet(book, sheet, new_name: str, record: bool = True):
    err = validate_sheet_name(book, new_name, exclude=sheet)
    if err:
        raise ValueError(err)
    old = sheet.name
    if old == new_name:
        return

    def fn_factory(s):
        def fn(info: RefInfo):
            if info.sheet is not None and info.sheet.lower() == old.lower():
                new = RefInfo(**info.__dict__)
                new.sheet = new_name
                new.sheet_raw = quote_sheet(new_name) + "!"
                return new
            return info
        return fn

    for s in book.sheets:
        for (r, c), cell in list(s.ws._cells.items()):
            text = _formula_of(cell)
            if text is not None and old.lower() in text.lower():
                disp = to_display(text)
                new = transform_refs(disp, fn_factory(s))
                if new != disp:
                    cell._value = _with_text(cell._value, to_storage(new))
    for dn in _defined_names(book):
        txt = "=" + (dn.attr_text or "")
        new = transform_refs(txt, fn_factory(None))
        if new != txt:
            dn.attr_text = new[1:]
    sheet.ws.title = new_name
    _reresolve(book)
    book.modified = True
    if record:
        book.undo.push(UndoEntry("Rename sheet", lambda: rename_sheet(book, sheet, old, record=False),
                                 lambda: rename_sheet(book, sheet, new_name, record=False),
                                 (new_name, Rect(1, 1, 1, 1))))


def _reresolve(book):
    """Re-parse all formulas (sheet names changed) and recompute those whose refs changed."""
    before = {k: list(v) for k, v in book.graph.precedents.items()}
    book._load_names()
    book.register_all()
    dirty = set()
    for k, refs in book.graph.precedents.items():
        old = before.get(k)
        if old is None or [(id(r[0]),) + tuple(r[1:]) for r in old] != [(id(r[0]),) + tuple(r[1:]) for r in refs]:
            dirty.add(k)
    for s in book.sheets:
        for rc in s.formulas:
            if (s, rc[0], rc[1]) not in before:
                dirty.add((s, rc[0], rc[1]))
    if dirty:
        dirty |= book.graph.all_dependents(dirty)
        book._run(dirty)
