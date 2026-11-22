"""LET, LAMBDA and the LAMBDA helpers: MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, ISOMITTED.

LET and LAMBDA themselves are evaluated by the evaluator (they bind names); they're registered
here so formulas using them count as supported."""
from __future__ import annotations

from ..errors import CALC, NA, VALUE, ExcelError, XLError
from ..registry import fn, scalar_arg
from ..values import MISSING, Array, RefValue, to_int


def _syntax_only(ctx, *args):
    raise XLError(VALUE)  # only reachable as a function value, e.g. MAP(A1:A3,LET)


fn("LET", kind="lazy", min_args=3, max_args=None)(_syntax_only)
fn("LAMBDA", kind="lazy", min_args=1, max_args=None)(_syntax_only)


@fn("ISOMITTED", kind="raw", min_args=1, max_args=1)
def ISOMITTED(ctx, v):
    return v is MISSING


# ------------------------------------------------------------------ helpers


def _func(f, nargs: int):
    """The LAMBDA (or function value) argument of a helper, checked for its number of parameters."""
    from ..evaluator import arity, is_callable

    if isinstance(f, ExcelError):
        raise XLError(f)
    if not is_callable(f):
        raise XLError(VALUE)
    lo, hi = arity(f)
    if nargs < lo or (hi is not None and nargs > hi):
        raise XLError(VALUE)
    return f


def _call(ctx, f, args):
    from ..evaluator import invoke

    return invoke(ctx, f, args)


def _clip(ctx, v: RefValue) -> RefValue:
    """A reference clipped to the sheet's used area (so A:A doesn't mean a million rows)."""
    mr, mc = ctx.engine.bounds(v.sheet)
    return RefValue(v.sheet, v.r1, v.c1, min(v.r2, max(mr, v.r1)), min(v.c2, max(mc, v.c1)))


def _cells(ctx, v) -> list[list]:
    """The elements a helper passes to its LAMBDA, as rows: single-cell references for a range,
    values for an array."""
    if isinstance(v, RefValue):
        v = _clip(ctx, v)
        s = v.sheet
        return [[RefValue(s, r, c, r, c) for c in range(v.c1, v.c2 + 1)] for r in range(v.r1, v.r2 + 1)]
    if isinstance(v, Array):
        return v.rows
    return [[None if v is MISSING else v]]


def _lines(ctx, v, by_col: bool) -> list:
    """Rows (or columns) of a range or array, as references or 1-row (1-column) arrays."""
    if isinstance(v, RefValue):
        v = _clip(ctx, v)
        s = v.sheet
        if by_col:
            return [RefValue(s, v.r1, c, v.r2, c) for c in range(v.c1, v.c2 + 1)]
        return [RefValue(s, r, v.c1, r, v.c2) for r in range(v.r1, v.r2 + 1)]
    if not isinstance(v, Array):
        v = Array([[None if v is MISSING else v]])
    if by_col:
        return [Array([[row[j]] for row in v.rows]) for j in range(v.ncols)]
    return [Array([row]) for row in v.rows]


def _value(ctx, r):
    """One cell of a helper's result: a LAMBDA must return a single value, not an array."""
    if isinstance(r, RefValue):
        if not r.is_cell:
            return CALC
        r = ctx.engine.get_value(r.sheet, r.r1, r.c1)
    elif isinstance(r, Array):
        if r.nrows != 1 or r.ncols != 1:
            return CALC
        r = r.rows[0][0]
    if r is None or r is MISSING:
        return 0
    return r


# ------------------------------------------------------------------ functions


@fn("MAP", kind="raw", min_args=2, max_args=None)
def MAP(ctx, *args):
    *arrays, f = args
    f = _func(f, len(arrays))
    grids = [_cells(ctx, a) for a in arrays]
    nr = max(len(g) for g in grids)
    nc = max(len(g[0]) if g else 0 for g in grids)
    out = []
    for i in range(nr):
        row = []
        for j in range(nc):
            vals = []
            for g in grids:
                if i < len(g) and j < len(g[i]):
                    vals.append(g[i][j])
                elif len(g) == 1 and len(g[0]) == 1:
                    vals.append(g[0][0])  # a single value goes with every element
                else:
                    vals.append(NA)
            row.append(_value(ctx, _call(ctx, f, vals)))
        out.append(row)
    return Array(out)


@fn("REDUCE", kind="raw", min_args=3, max_args=3)
def REDUCE(ctx, initial, array, f):
    f = _func(f, 2)
    acc = initial
    for row in _cells(ctx, array):
        for v in row:
            acc = _call(ctx, f, [acc, v])
    if acc is MISSING:
        return 0
    return acc


@fn("SCAN", kind="raw", min_args=3, max_args=3)
def SCAN(ctx, initial, array, f):
    f = _func(f, 2)
    acc = initial
    out = []
    for row in _cells(ctx, array):
        line = []
        for v in row:
            acc = _call(ctx, f, [acc, v])
            line.append(_value(ctx, acc))
        out.append(line)
    return Array(out)


@fn("BYROW", kind="raw", min_args=2, max_args=2)
def BYROW(ctx, array, f):
    f = _func(f, 1)
    return Array([[_value(ctx, _call(ctx, f, [line]))] for line in _lines(ctx, array, False)])


@fn("BYCOL", kind="raw", min_args=2, max_args=2)
def BYCOL(ctx, array, f):
    f = _func(f, 1)
    return Array([[_value(ctx, _call(ctx, f, [line])) for line in _lines(ctx, array, True)]])


@fn("MAKEARRAY", kind="raw", min_args=3, max_args=3)
def MAKEARRAY(ctx, rows, cols, f):
    nr, nc = to_int(scalar_arg(ctx, rows)), to_int(scalar_arg(ctx, cols))
    if nr < 1 or nc < 1 or nr * nc > 16_384 * 1_048_576:
        raise XLError(VALUE)
    f = _func(f, 2)
    return Array([[_value(ctx, _call(ctx, f, [i, j])) for j in range(1, nc + 1)] for i in range(1, nr + 1)])
