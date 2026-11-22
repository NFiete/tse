"""Runtime value types and Excel coercion rules."""
from __future__ import annotations

import math
import re

from .errors import DIV0, NA, NUM, VALUE, ExcelError, XLError


class _Missing:
    """An omitted function argument, e.g. the middle arg of IF(A1,,1)."""

    __slots__ = ()

    def __repr__(self):
        return "MISSING"


MISSING = _Missing()


class RefValue:
    """A rectangular reference to cells on a sheet (the result of evaluating A1:B5)."""

    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet, self.r1, self.c1, self.r2, self.c2 = sheet, r1, c1, r2, c2

    @property
    def is_cell(self) -> bool:
        return self.r1 == self.r2 and self.c1 == self.c2

    @property
    def nrows(self) -> int:
        return self.r2 - self.r1 + 1

    @property
    def ncols(self) -> int:
        return self.c2 - self.c1 + 1

    def __repr__(self):
        return f"RefValue({getattr(self.sheet, 'name', self.sheet)}!R{self.r1}C{self.c1}:R{self.r2}C{self.c2})"


class Array:
    """An in-memory 2D array of scalar values."""

    __slots__ = ("rows",)

    def __init__(self, rows: list[list]):
        self.rows = rows

    @property
    def nrows(self) -> int:
        return len(self.rows)

    @property
    def ncols(self) -> int:
        return len(self.rows[0]) if self.rows else 0

    def flat(self):
        for row in self.rows:
            yield from row

    def __repr__(self):
        return f"Array({self.rows!r})"


class FuncValue:
    """A function passed as a value, such as SUM in GROUPBY(A1:A9,B1:B9,SUM) (an eta-reduced
    LAMBDA; stored in files as _xleta.SUM)."""

    __slots__ = ("name", "spec")

    def __init__(self, name: str, spec):
        self.name, self.spec = name, spec

    def __repr__(self):
        return f"FuncValue({self.name})"


class LambdaValue:
    """A LAMBDA: parameter keys (upper-case names), the body, and the LET names and parameters
    visible where it was created."""

    __slots__ = ("params", "body", "env")
    name = "LAMBDA"

    def __init__(self, params: tuple, body, env: dict):
        self.params, self.body, self.env = params, body, env

    def __repr__(self):
        return f"LambdaValue({', '.join(self.params)})"


def is_number(v) -> bool:
    t = type(v)
    return t is int or t is float


def fix_number(x):
    """Normalize a numeric result: errors for nan/inf, ints for large values -> float."""
    if type(x) is float:
        if math.isnan(x) or math.isinf(x):
            raise XLError(NUM)
        return x
    if type(x) is int and (x > 2**53 or x < -(2**53)):
        return float(x)
    return x


_NUM_RE = re.compile(r"^\s*([+-]?)\s*\$?\s*((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?\s*(%?)\s*$")


def parse_number_text(s: str):
    """Parse text the way Excel coerces strings to numbers. Returns None if not numeric."""
    m = _NUM_RE.match(s)
    if m:
        sign, body, exp, pct = m.groups()
        v = float(body.replace(",", "") + ("e" + exp if exp else ""))
        if math.isinf(v):
            return None
        if sign == "-":
            v = -v
        if pct:
            v /= 100
        if v.is_integer() and abs(v) < 2**53 and not exp and "." not in body and not pct:
            v = int(v)
        return v
    from ..model.inputparse import parse_date_text

    d = parse_date_text(s)
    if d is not None:
        return d[0]
    return None


def to_number(v):
    t = type(v)
    if t is int or t is float:
        return v
    if t is bool:
        return int(v)
    if v is None or v is MISSING:
        return 0
    if t is str:
        n = parse_number_text(v)
        if n is None:
            raise XLError(VALUE)
        return n
    if isinstance(v, ExcelError):
        raise XLError(v)
    raise XLError(VALUE)


def to_int(v) -> int:
    n = to_number(v)
    return int(math.floor(n)) if n >= 0 else -int(math.floor(-n))  # truncate toward zero


def format_general_number(x) -> str:
    """Number -> text as Excel does when coercing (up to 15 significant digits)."""
    if type(x) is int:
        return str(x)
    if x == 0:
        return "0"
    if x.is_integer() and abs(x) < 1e15:
        return str(int(x))
    s = f"{x:.15g}"
    if "e" in s:
        mant, exp = s.split("e")
        e = int(exp)
        if -10 < e < 0:
            return f"{x:.{15 - e - 1}f}".rstrip("0").rstrip(".")
        return f"{mant}E{'+' if e >= 0 else '-'}{abs(e):02d}"
    return s


def to_text(v) -> str:
    t = type(v)
    if t is str:
        return v
    if t is bool:
        return "TRUE" if v else "FALSE"
    if t is int or t is float:
        return format_general_number(v)
    if v is None or v is MISSING:
        return ""
    if isinstance(v, ExcelError):
        raise XLError(v)
    raise XLError(VALUE)


def to_bool(v) -> bool:
    t = type(v)
    if t is bool:
        return v
    if t is int or t is float:
        return v != 0
    if v is None or v is MISSING:
        return False
    if t is str:
        u = v.strip().upper()
        if u == "TRUE":
            return True
        if u == "FALSE":
            return False
        raise XLError(VALUE)
    if isinstance(v, ExcelError):
        raise XLError(v)
    raise XLError(VALUE)


def check_error(v):
    if isinstance(v, ExcelError):
        raise XLError(v)
    return v


# ------------------------------------------------------------------ comparisons


def _round15(x: float) -> float:
    if x == 0 or not math.isfinite(x):
        return x
    return float(f"{x:.15g}")


def type_rank(v) -> int:
    t = type(v)
    if t is bool:
        return 2
    if t is str:
        return 1
    return 0


def compare(a, b) -> int:
    """Excel comparison: returns -1/0/1. Blank acts as 0 / "" / FALSE depending on the other side."""
    if isinstance(a, ExcelError):
        raise XLError(a)
    if isinstance(b, ExcelError):
        raise XLError(b)
    if a is None or a is MISSING:
        a = _blank_like(b)
    if b is None or b is MISSING:
        b = _blank_like(a)
    ra, rb = type_rank(a), type_rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = a.lower(), b.lower()
    elif ra == 0:
        a, b = _round15(a), _round15(b)
    return (a > b) - (a < b)


def _blank_like(other):
    t = type(other)
    if t is str:
        return ""
    if t is bool:
        return False
    return 0


def values_equal(a, b) -> bool:
    try:
        return compare(a, b) == 0
    except XLError:
        return False


__all__ = [
    "MISSING", "RefValue", "Array", "FuncValue", "LambdaValue", "to_number", "to_int", "to_text", "to_bool", "compare",
    "fix_number", "is_number", "check_error", "parse_number_text", "DIV0", "NA",
]
