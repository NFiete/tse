"""Spilled array results (dynamic arrays).

A formula whose result is an array "spills" into the cells below and to the right of it.
Each sheet keeps a SpillIndex: the spilled values of every anchor (the formula cell), a
per-column interval index for finding which anchor covers a cell, and the anchors that
couldn't spill because something was in the way (they show #SPILL!).
"""
from __future__ import annotations

from bisect import bisect_left, bisect_right


class SpillIndex:
    __slots__ = ("anchors", "cols", "blocked")

    def __init__(self):
        self.anchors: dict[tuple, list] = {}  # (r, c) -> rows of values, anchor's own value at [0][0]
        self.cols: dict[int, tuple[list, list]] = {}  # c -> ([r1, ...], [(r1, r2, anchor), ...]) sorted
        self.blocked: dict[tuple, tuple] = {}  # (r, c) -> (nrows, ncols) it wanted to spill into

    def __bool__(self):
        return bool(self.anchors) or bool(self.blocked)

    def clear(self):
        self.anchors.clear()
        self.cols.clear()
        self.blocked.clear()

    def owner(self, r: int, c: int):
        """The anchor whose spill covers (r, c), or None."""
        col = self.cols.get(c)
        if col is None:
            return None
        starts, ents = col
        i = bisect_right(starts, r) - 1
        if i >= 0:
            e = ents[i]
            if r <= e[1]:
                return e[2]
        return None

    def value(self, anchor, r: int, c: int):
        return self.anchors[anchor][r - anchor[0]][c - anchor[1]]

    def rect(self, anchor) -> tuple[int, int, int, int]:
        rows = self.anchors[anchor]
        return anchor[0], anchor[1], anchor[0] + len(rows) - 1, anchor[1] + len(rows[0]) - 1

    def overlaps(self, r1: int, c1: int, r2: int, c2: int, exclude=None) -> bool:
        """Whether another anchor's spill covers any cell of the rectangle."""
        for c in range(c1, c2 + 1):
            col = self.cols.get(c)
            if col is None:
                continue
            starts, ents = col
            i = max(0, bisect_right(starts, r1) - 1)
            for e in ents[i:]:
                if e[0] > r2:
                    break
                if e[1] >= r1 and e[2] != exclude:
                    return True
        return False

    def add(self, anchor, rows: list):
        r1, c1 = anchor
        r2 = r1 + len(rows) - 1
        for c in range(c1, c1 + len(rows[0])):
            col = self.cols.get(c)
            if col is None:
                col = self.cols[c] = ([], [])
            starts, ents = col
            i = bisect_left(starts, r1)
            starts.insert(i, r1)
            ents.insert(i, (r1, r2, anchor))
        self.anchors[anchor] = rows

    def remove(self, anchor):
        """Forget an anchor's spill; returns its rows (or None)."""
        rows = self.anchors.pop(anchor, None)
        if rows is None:
            return None
        r1, c1 = anchor
        for c in range(c1, c1 + len(rows[0])):
            col = self.cols.get(c)
            if col is None:
                continue
            starts, ents = col
            i = bisect_left(starts, r1)
            while i < len(ents) and ents[i][0] == r1:
                if ents[i][2] == anchor:
                    del starts[i]
                    del ents[i]
                    break
                i += 1
            if not starts:
                del self.cols[c]
        return rows

    def cells_in(self, r1: int, c1: int, r2: int, c2: int):
        """Spilled (non-anchor) cells inside a rectangle, as (r, c, anchor)."""
        for anchor, rows in self.anchors.items():
            ar, ac = anchor
            br, bc = ar + len(rows) - 1, ac + len(rows[0]) - 1
            if br < r1 or ar > r2 or bc < c1 or ac > c2:
                continue
            for r in range(max(r1, ar), min(r2, br) + 1):
                for c in range(max(c1, ac), min(c2, bc) + 1):
                    if (r, c) != anchor:
                        yield r, c, anchor

    def blocked_in(self, r1: int, c1: int, r2: int, c2: int):
        """Blocked anchors whose intended spill area overlaps the rectangle."""
        for (ar, ac), (nr, nc) in self.blocked.items():
            if ar <= r2 and ar + nr - 1 >= r1 and ac <= c2 and ac + nc - 1 >= c1:
                yield ar, ac

    def extent(self) -> tuple[int, int]:
        """(max row, max col) covered by spills."""
        mr = mc = 0
        for anchor, rows in self.anchors.items():
            mr = max(mr, anchor[0] + len(rows) - 1)
            mc = max(mc, anchor[1] + len(rows[0]) - 1)
        return mr, mc
