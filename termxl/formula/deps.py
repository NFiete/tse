"""Dependency graph between formula cells and the cells/ranges they reference.

Keys are (sheet, row, col) tuples where `sheet` is a Sheet object.
"""
from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import defaultdict

_NARROW = 32  # ranges at most this many columns wide are indexed per column


class DepGraph:
    def __init__(self):
        self.cell_deps: dict[tuple, set] = defaultdict(set)
        self.col_ranges: dict[tuple, set] = defaultdict(set)  # (sheet, col) -> {(r1, r2, fkey)}
        self.wide_ranges: dict[object, set] = defaultdict(set)  # sheet -> {(r1, c1, r2, c2, fkey)}
        self.entries: dict[tuple, list] = {}  # fkey -> registered entries (for removal)
        self.precedents: dict[tuple, list] = {}  # fkey -> [(sheet, r1, c1, r2, c2)]
        self.volatile: set[tuple] = set()

    def clear(self):
        self.__init__()

    def add(self, fkey: tuple, refs: list[tuple], volatile: bool):
        self.remove(fkey)
        ents = []
        for ref in refs:
            sheet, r1, c1, r2, c2 = ref
            if r1 == r2 and c1 == c2:
                k = (sheet, r1, c1)
                self.cell_deps[k].add(fkey)
                ents.append(("c", k, fkey))
            elif c2 - c1 < _NARROW:
                e = (r1, r2, fkey)
                for c in range(c1, c2 + 1):
                    self.col_ranges[(sheet, c)].add(e)
                    ents.append(("r", (sheet, c), e))
            else:
                e = (r1, c1, r2, c2, fkey)
                self.wide_ranges[sheet].add(e)
                ents.append(("w", sheet, e))
        self.entries[fkey] = ents
        self.precedents[fkey] = refs
        if volatile:
            self.volatile.add(fkey)

    def remove(self, fkey: tuple):
        ents = self.entries.pop(fkey, None)
        self.precedents.pop(fkey, None)
        self.volatile.discard(fkey)
        if not ents:
            return
        for kind, k, e in ents:
            if kind == "c":
                s = self.cell_deps.get(k)
            elif kind == "r":
                s = self.col_ranges.get(k)
            else:
                s = self.wide_ranges.get(k)
            if s is not None:
                s.discard(e)
                if not s:
                    if kind == "c":
                        del self.cell_deps[k]
                    elif kind == "r":
                        del self.col_ranges[k]
                    else:
                        del self.wide_ranges[k]

    def remove_sheet(self, sheet):
        for fkey in [k for k in self.entries if k[0] is sheet]:
            self.remove(fkey)

    # ------------------------------------------------------------------ queries

    def dependents(self, cells) -> set:
        """Formula keys that directly reference any of the given cells."""
        out: set = set()
        by_col: dict[tuple, list] = defaultdict(list)
        by_sheet: dict[object, list] = defaultdict(list)
        cell_deps = self.cell_deps
        for key in cells:
            d = cell_deps.get(key)
            if d:
                out |= d
            s, r, c = key
            by_col[(s, c)].append(r)
            by_sheet[s].append((r, c))
        for sc, rows in by_col.items():
            ents = self.col_ranges.get(sc)
            if not ents:
                continue
            rows.sort()
            n = len(rows)
            for r1, r2, f in ents:
                if f in out:
                    continue
                i = bisect_left(rows, r1)
                if i < n and rows[i] <= r2:
                    out.add(f)
        for s, pts in by_sheet.items():
            wide = self.wide_ranges.get(s)
            if not wide:
                continue
            for r1, c1, r2, c2, f in wide:
                if f in out:
                    continue
                for r, c in pts:
                    if r1 <= r <= r2 and c1 <= c <= c2:
                        out.add(f)
                        break
        return out

    def dependents_rect(self, sheet, r1: int, c1: int, r2: int, c2: int) -> set:
        """Formula keys that directly reference any cell of a rectangle."""
        out: set = set()
        cell_deps = self.cell_deps
        if (r2 - r1 + 1) * (c2 - c1 + 1) <= len(cell_deps):
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    d = cell_deps.get((sheet, r, c))
                    if d:
                        out |= d
        else:
            for (s, r, c), d in cell_deps.items():
                if s is sheet and r1 <= r <= r2 and c1 <= c <= c2:
                    out |= d
        col_ranges = self.col_ranges
        for c in range(c1, c2 + 1):
            ents = col_ranges.get((sheet, c))
            if ents:
                for a, b, f in ents:
                    if a <= r2 and b >= r1:
                        out.add(f)
        for wr1, wc1, wr2, wc2, f in self.wide_ranges.get(sheet, ()):
            if wr1 <= r2 and wr2 >= r1 and wc1 <= c2 and wc2 >= c1:
                out.add(f)
        return out

    def all_dependents(self, cells) -> set:
        result: set = set()
        frontier = set(cells)
        while frontier:
            d = self.dependents(frontier)
            d -= result
            result |= d
            frontier = d
        return result

    def topo_order(self, dirty: set) -> list:
        """Order dirty formula keys so precedents come first (cycles are broken arbitrarily)."""
        cols_by_sheet: dict[object, dict[int, list]] = defaultdict(lambda: defaultdict(list))
        for s, r, c in dirty:
            cols_by_sheet[s][c].append(r)
        sorted_cols: dict[object, list] = {}
        for s, cols in cols_by_sheet.items():
            for rows in cols.values():
                rows.sort()
            sorted_cols[s] = sorted(cols)

        def dirty_precs(key):
            for sheet, r1, c1, r2, c2 in self.precedents.get(key, ()):
                if r1 == r2 and c1 == c2:
                    k = (sheet, r1, c1)
                    if k in dirty:
                        yield k
                    continue
                cols = cols_by_sheet.get(sheet)
                if not cols:
                    continue
                sc = sorted_cols[sheet]
                for c in sc[bisect_left(sc, c1):bisect_right(sc, c2)]:
                    rows = cols[c]
                    for r in rows[bisect_left(rows, r1):bisect_right(rows, r2)]:
                        yield (sheet, r, c)

        order: list = []
        state: dict = {}
        for start in dirty:
            if start in state:
                continue
            state[start] = 1
            stack = [(start, dirty_precs(start))]
            while stack:
                node, it = stack[-1]
                for p in it:
                    if p not in state:
                        state[p] = 1
                        stack.append((p, dirty_precs(p)))
                        break
                else:
                    stack.pop()
                    state[node] = 2
                    order.append(node)
        return order
