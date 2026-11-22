"""Logical and information functions."""
from __future__ import annotations

from ..errors import ERROR_TYPE_NUM, NA, VALUE, ExcelError, XLError
from ..registry import flatten, fn, lookup_equal, to_scalar
from ..values import MISSING, Array, RefValue, is_number, to_bool, to_number


def _is_multi(v):
    return isinstance(v, Array) or (isinstance(v, RefValue) and not v.is_cell)


def _unmissing(v, default=0):
    return default if v is MISSING else v


@fn("IF", kind="lazy", min_args=1, max_args=3)
def IF(ctx, cond, when_true=None, when_false=None):
    c = cond()
    if _is_multi(c):
        from ..evaluator import lift

        tv = _unmissing(when_true()) if when_true else True
        fv = _unmissing(when_false()) if when_false else False

        def pick(cv, a, b):
            if isinstance(cv, ExcelError):
                return cv
            return a if to_bool(cv) else b

        return lift(ctx, pick, [c, tv, fv], errors_ok=True)
    c = to_scalar(ctx, c)
    if isinstance(c, ExcelError):
        return c
    if to_bool(c):
        return _unmissing(when_true()) if when_true else True
    return _unmissing(when_false()) if when_false else False


@fn("IFS", kind="lazy", min_args=2)
def IFS(ctx, *args):
    if len(args) % 2:
        raise XLError(VALUE)
    for i in range(0, len(args), 2):
        c = to_scalar(ctx, args[i]())
        if isinstance(c, ExcelError):
            return c
        if to_bool(c):
            return _unmissing(args[i + 1]())
    return NA


@fn("IFERROR", kind="lazy", min_args=2, max_args=2)
def IFERROR(ctx, value, alt):
    v = value()
    if _is_multi(v):
        from ..evaluator import lift

        a = _unmissing(alt())
        return lift(ctx, lambda x, y: y if isinstance(x, ExcelError) else x, [v, a], errors_ok=True)
    v = to_scalar(ctx, v)
    if isinstance(v, ExcelError):
        return _unmissing(alt())
    return _unmissing(v)


@fn("IFNA", kind="lazy", min_args=2, max_args=2)
def IFNA(ctx, value, alt):
    v = value()
    if _is_multi(v):
        from ..evaluator import lift

        a = _unmissing(alt())
        return lift(ctx, lambda x, y: y if x is NA else x, [v, a], errors_ok=True)
    v = to_scalar(ctx, v)
    if v is NA:
        return _unmissing(alt())
    return _unmissing(v)


@fn("SWITCH", kind="lazy", min_args=3)
def SWITCH(ctx, expr, *args):
    v = to_scalar(ctx, expr())
    if isinstance(v, ExcelError):
        return v
    pairs = len(args) // 2
    for i in range(pairs):
        cand = to_scalar(ctx, args[2 * i]())
        if isinstance(cand, ExcelError):
            return cand
        if lookup_equal(v, cand):
            return _unmissing(args[2 * i + 1]())
    if len(args) % 2:
        return _unmissing(args[-1]())
    return NA


@fn("CHOOSE", kind="lazy", min_args=2)
def CHOOSE(ctx, index, *choices):
    i = to_scalar(ctx, index())
    if isinstance(i, ExcelError):
        return i
    n = int(to_number(i))
    if n < 1 or n > len(choices):
        return VALUE
    return _unmissing(choices[n - 1]())


def _logicals(ctx, args):
    out = []
    for v, direct in flatten(ctx, args):
        if isinstance(v, ExcelError):
            raise XLError(v)
        t = type(v)
        if t is bool:
            out.append(v)
        elif t is int or t is float:
            out.append(v != 0)
        elif direct and t is str:
            out.append(to_bool(v))
    if not out:
        raise XLError(VALUE)
    return out


@fn("AND", kind="raw", min_args=1)
def AND(ctx, *args):
    return all(_logicals(ctx, args))


@fn("OR", kind="raw", min_args=1)
def OR(ctx, *args):
    return any(_logicals(ctx, args))


@fn("XOR", kind="raw", min_args=1)
def XOR(ctx, *args):
    return sum(_logicals(ctx, args)) % 2 == 1


@fn("NOT")
def NOT(v):
    return not to_bool(v)


@fn("TRUE")
def TRUE():
    return True


@fn("FALSE")
def FALSE():
    return False


# ------------------------------------------------------------------ information


@fn("ISBLANK", errors_ok=True)
def ISBLANK(v):
    return v is None


@fn("ISNUMBER", errors_ok=True)
def ISNUMBER(v):
    return is_number(v)


@fn("ISTEXT", errors_ok=True)
def ISTEXT(v):
    return type(v) is str


@fn("ISNONTEXT", errors_ok=True)
def ISNONTEXT(v):
    return type(v) is not str


@fn("ISLOGICAL", errors_ok=True)
def ISLOGICAL(v):
    return type(v) is bool


@fn("ISERROR", errors_ok=True)
def ISERROR(v):
    return isinstance(v, ExcelError)


@fn("ISERR", errors_ok=True)
def ISERR(v):
    return isinstance(v, ExcelError) and v is not NA


@fn("ISNA", errors_ok=True)
def ISNA(v):
    return v is NA


@fn("ISEVEN")
def ISEVEN(v):
    return int(to_number(v)) % 2 == 0


@fn("ISODD")
def ISODD(v):
    return int(to_number(v)) % 2 == 1


@fn("ISREF", kind="raw", min_args=1, max_args=1)
def ISREF(ctx, v):
    return isinstance(v, RefValue)


@fn("ISFORMULA", kind="raw", min_args=1, max_args=1)
def ISFORMULA(ctx, v):
    if not isinstance(v, RefValue):
        raise XLError(NA)
    return ctx.engine.is_formula(v.sheet, v.r1, v.c1)


@fn("NA")
def NA_():
    return NA


@fn("ERROR.TYPE", errors_ok=True)
def ERROR_TYPE(v):
    if isinstance(v, ExcelError) and v in ERROR_TYPE_NUM:
        return ERROR_TYPE_NUM[v]
    return NA


@fn("N")
def N(v):
    if is_number(v):
        return v
    if type(v) is bool:
        return int(v)
    return 0


@fn("T")
def T(v):
    return v if type(v) is str else ""


@fn("TYPE", kind="raw", min_args=1, max_args=1)
def TYPE(ctx, v):
    if _is_multi(v):
        return 64
    v = to_scalar(ctx, v)
    if isinstance(v, ExcelError):
        return 16
    if type(v) is bool:
        return 4
    if type(v) is str:
        return 2
    return 1
