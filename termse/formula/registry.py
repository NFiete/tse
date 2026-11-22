"""Function registry and shared helpers for function implementations."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable

from .errors import NA, VALUE, ExcelError, XLError
from .values import MISSING, Array, RefValue, compare, is_number, parse_number_text, to_number


@dataclass
class FnSpec:
    name: str
    impl: Callable
    kind: str  # 'scalar' (lifted elementwise) | 'raw' | 'lazy'
    min_args: int
    max_args: int | None
    volatile: bool
    errors_ok: bool  # scalar: pass error values through instead of short-circuiting


FUNCTIONS: dict[str, FnSpec] = {}


def fn(*names: str, kind: str = "scalar", min_args: int = 0, max_args: int | None = -1,
       volatile: bool = False, errors_ok: bool = False):
    """Register a function implementation.

    scalar: impl(*scalars); arrays/ranges in any argument are mapped elementwise.
    raw:    impl(ctx, *values); values may be RefValue / Array / scalars / MISSING.
    lazy:   impl(ctx, *thunks); each thunk() evaluates that argument.
    max_args=-1 means "infer from impl signature" for scalar functions.
    """

    def deco(impl):
        mx = max_args
        if mx == -1:
            import inspect

            params = list(inspect.signature(impl).parameters.values())
            if kind != "scalar":
                params = params[1:]
            if any(p.kind == p.VAR_POSITIONAL for p in params):
                mx = None
            else:
                mx = len(params)
        mn = min_args
        if mn == 0 and kind == "scalar":
            import inspect

            mn = sum(1 for p in inspect.signature(impl).parameters.values()
                     if p.default is p.empty and p.kind == p.POSITIONAL_OR_KEYWORD)
        for name in names:
            FUNCTIONS[name] = FnSpec(name, impl, kind, mn, mx, volatile, errors_ok)
        return impl

    return deco


# ------------------------------------------------------------------ helpers for raw functions


def iter_ref(ctx, rv: RefValue, include_blank: bool = False):
    """Values in a reference, clipped to the sheet's used area."""
    eng = ctx.engine
    mr, mc = eng.bounds(rv.sheet)
    r2, c2 = min(rv.r2, mr), min(rv.c2, mc)
    if r2 < rv.r1 or c2 < rv.c1:
        return []
    if include_blank:
        return [v for row in eng.range_rows(rv.sheet, rv.r1, rv.c1, r2, c2) for v in row]
    return eng.range_values(rv.sheet, rv.r1, rv.c1, r2, c2)


def flatten(ctx, args):
    """Yield (value, direct) pairs. direct=False for values that came from ranges/arrays."""
    for a in args:
        if isinstance(a, RefValue):
            for v in iter_ref(ctx, a):
                yield v, False
        elif isinstance(a, Array):
            for v in a.flat():
                yield v, False
        elif a is MISSING:
            continue
        else:
            yield a, True


def numbers(ctx, args, *, bools_in_refs=False):
    """Numbers for SUM/AVERAGE/MIN/MAX-style aggregation (errors propagate)."""
    out = []
    ap = out.append
    for a in args:
        if isinstance(a, (RefValue, Array)):
            vals = iter_ref(ctx, a) if isinstance(a, RefValue) else a.flat()
            for v in vals:
                t = type(v)
                if t is int or t is float:
                    ap(v)
                elif t is bool:
                    if bools_in_refs:
                        ap(int(v))
                elif isinstance(v, ExcelError):
                    raise XLError(v)
            continue
        if a is MISSING:
            continue
        t = type(a)
        if t is int or t is float:
            ap(a)
        elif t is bool:
            ap(int(a))
        elif a is None:
            ap(0)
        elif isinstance(a, ExcelError):
            raise XLError(a)
        else:
            ap(to_number(a))
    return out


def to_array(ctx, v) -> Array:
    """Convert any value to an Array (references are clipped to the used area)."""
    if isinstance(v, Array):
        return v
    if isinstance(v, RefValue):
        eng = ctx.engine
        mr, mc = eng.bounds(v.sheet)
        r2, c2 = min(v.r2, max(mr, v.r1)), min(v.c2, max(mc, v.c1))
        return Array(eng.range_rows(v.sheet, v.r1, v.c1, r2, c2))
    if v is MISSING:
        v = None
    return Array([[v]])


def shaped_array(ctx, v, trim_edges: bool = True) -> Array:
    """Like to_array, for results whose shape matters (operators mapped over ranges, SORT,
    FILTER...): a reference keeps its full size, including blank rows and columns past the
    sheet's used area (as shared blank rows, so they're cheap). With trim_edges, references
    running to the last row or column of the sheet (A:A, 1:1, A2:A1048576) stay trimmed to the
    used area, so whole-column formulas don't work through a million blank rows."""
    from ..model.refs import MAX_COL, MAX_ROW

    a = to_array(ctx, v)
    if not isinstance(v, RefValue):
        return a
    nr = a.nrows if trim_edges and v.r2 >= MAX_ROW else v.nrows
    nc = a.ncols if trim_edges and v.c2 >= MAX_COL else v.ncols
    rows = a.rows
    if a.ncols < nc:
        extra = [None] * (nc - a.ncols)
        rows = [r + extra for r in rows]
    if len(rows) < nr:
        rows = rows + [[None] * nc] * (nr - len(rows))
    return Array(rows) if rows is not a.rows else a


def to_scalar(ctx, v):
    """Reduce a reference/array to one value (top-left, or implicit intersection)."""
    if isinstance(v, RefValue):
        if v.is_cell:
            return ctx.engine.get_value(v.sheet, v.r1, v.c1)
        if v.c1 == v.c2 and v.r1 <= ctx.row <= v.r2 and v.sheet is ctx.sheet:
            return ctx.engine.get_value(v.sheet, ctx.row, v.c1)
        if v.r1 == v.r2 and v.c1 <= ctx.col <= v.c2 and v.sheet is ctx.sheet:
            return ctx.engine.get_value(v.sheet, v.r1, ctx.col)
        return ctx.engine.get_value(v.sheet, v.r1, v.c1)
    if isinstance(v, Array):
        return v.rows[0][0] if v.rows and v.rows[0] else None
    return v


def scalar_arg(ctx, v):
    """A scalar argument with errors raised."""
    v = to_scalar(ctx, v)
    if isinstance(v, ExcelError):
        raise XLError(v)
    return None if v is MISSING else v


def vector(ctx, v) -> list:
    """Flatten a 1-D range/array to a list (row-major)."""
    return list(to_array(ctx, v).flat())


# ------------------------------------------------------------------ criteria (COUNTIF etc.)

_WILD_CACHE: dict[str, re.Pattern] = {}


def wildcard_regex(pattern: str) -> re.Pattern:
    rx = _WILD_CACHE.get(pattern)
    if rx is None:
        out, i = [], 0
        while i < len(pattern):
            ch = pattern[i]
            if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
                out.append(re.escape(pattern[i + 1]))
                i += 2
                continue
            if ch == "*":
                out.append(".*")
            elif ch == "?":
                out.append(".")
            else:
                out.append(re.escape(ch))
            i += 1
        rx = re.compile("".join(out), re.I | re.S)
        if len(_WILD_CACHE) > 2000:
            _WILD_CACHE.clear()
        _WILD_CACHE[pattern] = rx
    return rx


def has_wildcards(s: str) -> bool:
    return "*" in s or "?" in s


def make_criterion(crit) -> Callable[[object], bool]:
    """Build a predicate from an Excel criteria value such as 5, ">5", "<>x", "a*"."""
    if isinstance(crit, ExcelError):
        return lambda v: v is crit
    if crit is None or crit is MISSING:
        crit = ""
    if type(crit) is bool:
        return lambda v: type(v) is bool and v == crit
    if is_number(crit):
        return lambda v: is_number(v) and compare(v, crit) == 0
    s = str(crit)
    op = "="
    for o in ("<>", "<=", ">=", "=", "<", ">"):
        if s.startswith(o):
            op, s = o, s[len(o):]
            break
    num = parse_number_text(s) if s.strip() else None
    if num is not None and not (s.strip().upper() in ("TRUE", "FALSE")):
        def num_pred(v, op=op, n=num):
            if not is_number(v):
                if type(v) is str:
                    pv = parse_number_text(v)
                    if pv is None or op != "=":
                        return op == "<>"
                    return compare(pv, n) == 0
                return op == "<>"
            c = compare(v, n)
            return _apply_op(op, c)

        return num_pred
    up = s.upper()
    if up in ("TRUE", "FALSE"):
        b = up == "TRUE"
        return lambda v, op=op: _apply_op(op, 0 if (type(v) is bool and v == b) else 1) if op in ("=", "<>") \
            else False
    if s == "":
        if op == "=":
            return lambda v: v is None or v == ""
        if op == "<>":
            return lambda v: not (v is None or v == "")
        return lambda v: False
    if op in ("=", "<>"):
        if has_wildcards(s):
            rx = wildcard_regex(s)
            if op == "=":
                return lambda v: type(v) is str and rx.fullmatch(v) is not None
            return lambda v: not (type(v) is str and rx.fullmatch(v) is not None)
        low = s.lower()
        if op == "=":
            return lambda v: type(v) is str and v.lower() == low
        return lambda v: not (type(v) is str and v.lower() == low)
    low = s.lower()
    return lambda v, op=op: type(v) is str and _apply_op(op, (v.lower() > low) - (v.lower() < low))


def _apply_op(op: str, c: int) -> bool:
    if op == "=":
        return c == 0
    if op == "<>":
        return c != 0
    if op == "<":
        return c < 0
    if op == ">":
        return c > 0
    if op == "<=":
        return c <= 0
    return c >= 0


def lookup_equal(needle, v) -> bool:
    """Exact-match semantics for MATCH/VLOOKUP/XLOOKUP (case-insensitive text)."""
    if isinstance(needle, str) and isinstance(v, str):
        return needle.lower() == v.lower()
    if type(needle) is bool or type(v) is bool:
        return type(needle) is type(v) and needle == v
    if is_number(needle) and is_number(v):
        return compare(needle, v) == 0
    return False


def wild_matcher(needle):
    if isinstance(needle, str) and has_wildcards(needle):
        rx = wildcard_regex(needle)
        return lambda v: isinstance(v, str) and rx.fullmatch(v) is not None
    return lambda v: lookup_equal(needle, v)


def raise_na():
    raise XLError(NA)


def raise_value():
    raise XLError(VALUE)
