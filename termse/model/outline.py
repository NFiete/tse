"""Outlines: grouped rows and columns that collapse and expand (Excel's Data > Group).

Excel keeps outlines in the row/column dimensions: each row or column has an outline
level (0-7). A group is a maximal run of rows/columns whose level is >= L. Its summary
row/column (the one carrying the +/- button) sits just below/right of the group, or
above/left of it when the sheet's outlinePr says summaryBelow/summaryRight="0".
Collapsing hides the group's members and sets `collapsed` on the summary; that flag is
how nested groups remember being collapsed while an outer group is hidden.
"""
from __future__ import annotations

from dataclasses import dataclass

from .refs import MAX_COL, MAX_ROW, Rect, col_letter
from .undo import UndoEntry

MAX_LEVEL = 7
MAX_SPAN = 100_000  # most rows/columns grouped or ungrouped in one go


@dataclass(frozen=True)
class Group:
    axis: str  # "row" | "col"
    level: int
    start: int
    end: int
    summary: int | None  # row/column with the +/- button (None at the sheet's edge)


# --------------------------------------------------------------------------- reading


def _limit(axis: str) -> int:
    return MAX_ROW if axis == "row" else MAX_COL


def summary_after(sheet, axis: str) -> bool:
    """True if summary rows are below groups (columns right of groups), Excel's default."""
    pr = getattr(sheet.ws.sheet_properties, "outlinePr", None)
    if pr is None:
        return True
    v = pr.summaryBelow if axis == "row" else pr.summaryRight
    return v is None or bool(v)


def symbols_shown(sheet) -> bool:
    pr = getattr(sheet.ws.sheet_properties, "outlinePr", None)
    return pr is None or pr.showOutlineSymbols is None or bool(pr.showOutlineSymbols)


def level(sheet, axis: str, i: int) -> int:
    if axis == "col":
        return sheet.col_outline(i)[0]
    d = sheet.ws.row_dimensions.get(i)
    return (d.outlineLevel or 0) if d is not None else 0


def collapsed_flag(sheet, axis: str, i: int) -> bool:
    if axis == "col":
        return sheet.col_outline(i)[1]
    d = sheet.ws.row_dimensions.get(i)
    return bool(d is not None and d.collapsed)


def is_hidden(sheet, axis: str, i: int) -> bool:
    return sheet.col_width(i)[1] if axis == "col" else sheet.row_hidden(i)


def outlined(sheet, axis: str) -> list[int]:
    """Sorted rows/columns with an outline level above 0."""
    if axis == "col":
        sheet.col_outline(1)
        return sorted(c for c, (lv, _) in sheet._colo.items() if lv > 0)
    return sorted(r for r, d in sheet.ws.row_dimensions.items() if d.outlineLevel)


def max_level(sheet, axis: str) -> int:
    if axis == "col":
        sheet.col_outline(1)
        return max((lv for lv, _ in sheet._colo.values()), default=0)
    dims = sheet.ws.row_dimensions
    key = (len(dims), sheet.outline_gen)
    cache = sheet._row_outline_cache
    if cache is None or cache[0] != key:
        cache = (key, max((d.outlineLevel or 0 for d in dims.values()), default=0))
        sheet._row_outline_cache = cache
    return cache[1]


def run(sheet, axis: str, i: int, lvl: int) -> tuple[int, int]:
    """The longest stretch of rows/columns around i whose level is at least lvl."""
    lo = hi = i
    while lo > 1 and level(sheet, axis, lo - 1) >= lvl:
        lo -= 1
    lim = _limit(axis)
    while hi < lim and level(sheet, axis, hi + 1) >= lvl:
        hi += 1
    return lo, hi


def make_group(sheet, axis: str, lvl: int, start: int, end: int) -> Group:
    s = end + 1 if summary_after(sheet, axis) else start - 1
    return Group(axis, lvl, start, end, s if 1 <= s <= _limit(axis) else None)


def group_at(sheet, axis: str, i: int, lvl: int) -> Group:
    return make_group(sheet, axis, lvl, *run(sheet, axis, i, lvl))


def groups_at(sheet, axis: str, i: int) -> list[Group]:
    """Groups that i belongs to or is the summary of, outermost first."""
    lv = level(sheet, axis, i)
    out = [group_at(sheet, axis, i, k) for k in range(1, lv + 1)]
    nb = i - 1 if summary_after(sheet, axis) else i + 1
    if 1 <= nb <= _limit(axis):
        out += [group_at(sheet, axis, nb, k) for k in range(lv + 1, level(sheet, axis, nb) + 1)]
    return out


def is_collapsed(sheet, g: Group) -> bool:
    """A group is collapsed when none of its members are showing."""
    return all(is_hidden(sheet, g.axis, i) for i in range(g.start, g.end + 1))


# --------------------------------------------------------------------------- changing


def _snapshot(sheet, axis: str, idxs) -> dict:
    return {i: (level(sheet, axis, i), is_hidden(sheet, axis, i), collapsed_flag(sheet, axis, i)) for i in idxs}


def _spans(idxs) -> list[tuple[int, int]]:
    out: list[list[int]] = []
    for i in sorted(idxs):
        if out and out[-1][1] == i - 1:
            out[-1][1] = i
        else:
            out.append([i, i])
    return [tuple(s) for s in out]


def _apply(book, sheet, axis: str, state: dict):
    ws = sheet.ws
    if axis == "col":
        from .structure import isolate_cols

        for a, b in _spans(state):
            isolate_cols(ws, a, b)
        for c, (lv, hidden, flag) in state.items():
            d = ws.column_dimensions[col_letter(c)]
            d.outlineLevel, d.hidden, d.collapsed = lv, hidden, flag
    else:
        for r, (lv, hidden, flag) in state.items():
            d = ws.row_dimensions[r]
            d.outlineLevel, d.hidden, d.collapsed = lv, hidden, flag
    sheet.invalidate_layout()
    book.modified = True


def _change(book, sheet, axis: str, new: dict, label: str, where: Rect | None) -> bool:
    old = _snapshot(sheet, axis, new)
    new = {i: v for i, v in new.items() if old[i] != v}
    if not new:
        return False
    old = {i: old[i] for i in new}
    _apply(book, sheet, axis, new)
    book.undo.push(UndoEntry(label, lambda: _apply(book, sheet, axis, old), lambda: _apply(book, sheet, axis, new),
                             (sheet.name, where) if where is not None else None))
    return True


def _rect(axis: str, a: int, b: int) -> Rect:
    return Rect(a, 1, b, MAX_COL) if axis == "row" else Rect(1, a, MAX_ROW, b)


def _check_span(a: int, b: int):
    if b - a + 1 > MAX_SPAN:
        raise ValueError(f"Can't group more than {MAX_SPAN:,} rows or columns at once")


def group(book, sheet, axis: str, a: int, b: int) -> bool:
    """Group rows/columns a..b (one outline level deeper)."""
    _check_span(a, b)
    snap = _snapshot(sheet, axis, range(a, b + 1))
    if all(lv >= MAX_LEVEL for lv, _, _ in snap.values()):
        raise ValueError(f"Outlines can't be more than {MAX_LEVEL} levels deep")
    new = {i: (min(lv + 1, MAX_LEVEL), h, f) for i, (lv, h, f) in snap.items()}
    return _change(book, sheet, axis, new, "Group " + ("rows" if axis == "row" else "columns"), _rect(axis, a, b))


def ungroup(book, sheet, axis: str, a: int, b: int) -> bool:
    """Take rows/columns a..b out of their innermost group. Members hidden by a collapsed group
    are shown again, since there would be no button left to expand them."""
    _check_span(a, b)
    snap = _snapshot(sheet, axis, range(a, b + 1))
    new = {}
    clear_flags = set()
    seen: dict[tuple, bool] = {}
    for i, (lv, h, f) in snap.items():
        if lv == 0:
            continue
        g = group_at(sheet, axis, i, lv)
        key = (lv, g.start)
        if key not in seen:
            seen[key] = is_collapsed(sheet, g)
            if g.start >= a and g.end <= b and g.summary is not None:
                clear_flags.add(g.summary)  # the whole group goes away
        new[i] = (lv - 1, h and not seen[key], f)
    if not new:
        return False
    for s in clear_flags:
        lv, h, _ = new[s] if s in new else _snapshot(sheet, axis, [s])[s]
        new[s] = (lv, h, False)
    return _change(book, sheet, axis, new, "Ungroup " + ("rows" if axis == "row" else "columns"), _rect(axis, a, b))


def _nested_collapsed(sheet, g: Group) -> set[int]:
    """Members of g that belong to a nested group which is itself flagged collapsed."""
    keep: set[int] = set()
    for lvl in range(g.level + 1, MAX_LEVEL + 1):
        found = False
        i = g.start
        while i <= g.end:
            if level(sheet, g.axis, i) < lvl:
                i += 1
                continue
            j = i
            while j < g.end and level(sheet, g.axis, j + 1) >= lvl:
                j += 1
            found = True
            sub = make_group(sheet, g.axis, lvl, i, j)
            if sub.summary is not None and sub.summary != g.summary and collapsed_flag(sheet, g.axis, sub.summary):
                keep.update(range(i, j + 1))
            i = j + 1
        if not found:
            break
    return keep


def set_collapsed(book, sheet, g: Group, collapse: bool) -> bool:
    axis = g.axis
    snap = _snapshot(sheet, axis, range(g.start, g.end + 1))
    if collapse:
        new = {i: (lv, True, f) for i, (lv, _, f) in snap.items()}
    else:
        keep = _nested_collapsed(sheet, g)
        new = {i: (lv, i in keep, f) for i, (lv, _, f) in snap.items()}
    if g.summary is not None:
        slv, sh, _ = _snapshot(sheet, axis, [g.summary])[g.summary]
        new[g.summary] = (slv, sh, collapse)
    return _change(book, sheet, axis, new, "Collapse group" if collapse else "Expand group", None)


def show_level(book, sheet, axis: str, n: int) -> bool:
    """Like Excel's outline level buttons: show levels below n, collapse groups at level n and deeper."""
    idxs = outlined(sheet, axis)
    if not idxs:
        return False
    snap = _snapshot(sheet, axis, idxs)
    new = {i: (lv, lv >= n, f) for i, (lv, _, f) in snap.items()}
    top = max(lv for lv, _, _ in snap.values())
    for lvl in range(1, top + 1):
        for a, b in _spans(i for i in idxs if snap[i][0] >= lvl):
            s = make_group(sheet, axis, lvl, a, b).summary
            if s is None:
                continue
            if s not in new:
                slv, sh, _ = _snapshot(sheet, axis, [s])[s]
                new[s] = (slv, sh, False)
            lv, h, _ = new[s]
            if lv < lvl:  # only the group's own summary, not a member of a deeper group
                new[s] = (lv, h, lvl >= n)
    return _change(book, sheet, axis, new, f"Show outline level {n}", None)
