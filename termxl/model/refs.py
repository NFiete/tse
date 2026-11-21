"""Cell address helpers: A1 <-> (row, col), ranges and sheet-name quoting."""
from __future__ import annotations

import re
from dataclasses import dataclass

MAX_ROW = 1_048_576
MAX_COL = 16_384

_COL_CACHE: dict[int, str] = {}


def col_letter(col: int) -> str:
    s = _COL_CACHE.get(col)
    if s is None:
        n, out = col, ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            out = chr(65 + rem) + out
        _COL_CACHE[col] = s = out
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def addr(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def parse_addr(text: str) -> tuple[int, int] | None:
    m = _CELL_RE.match(text.strip())
    if not m:
        return None
    r, c = int(m.group(2)), col_index(m.group(1))
    if not (1 <= r <= MAX_ROW and 1 <= c <= MAX_COL):
        return None
    return r, c


@dataclass(frozen=True)
class Rect:
    r1: int
    c1: int
    r2: int
    c2: int

    @staticmethod
    def of(r1: int, c1: int, r2: int, c2: int) -> "Rect":
        return Rect(min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))

    def contains(self, r: int, c: int) -> bool:
        return self.r1 <= r <= self.r2 and self.c1 <= c <= self.c2

    @property
    def rows(self) -> int:
        return self.r2 - self.r1 + 1

    @property
    def cols(self) -> int:
        return self.c2 - self.c1 + 1

    def cells(self):
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                yield r, c

    def a1(self) -> str:
        if self.r1 == self.r2 and self.c1 == self.c2:
            return addr(self.r1, self.c1)
        return f"{addr(self.r1, self.c1)}:{addr(self.r2, self.c2)}"


def parse_range(text: str) -> Rect | None:
    """Parse 'A1', 'A1:B5' (no sheet prefix)."""
    parts = text.strip().split(":")
    if len(parts) == 1:
        p = parse_addr(parts[0])
        return Rect(p[0], p[1], p[0], p[1]) if p else None
    if len(parts) == 2:
        a, b = parse_addr(parts[0]), parse_addr(parts[1])
        if a and b:
            return Rect.of(a[0], a[1], b[0], b[1])
    return None


_PLAIN_SHEET = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")


def quote_sheet(name: str) -> str:
    """Return a sheet name as it must appear in a formula (quoted when needed)."""
    if _PLAIN_SHEET.match(name) and not _CELL_RE.match(name) and not re.match(
        r"^[A-Za-z]{1,3}\d+$", name
    ) and name.upper() not in ("TRUE", "FALSE"):
        return name
    return "'" + name.replace("'", "''") + "'"
