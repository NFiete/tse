"""Math and trigonometry functions."""
from __future__ import annotations

import math
import random
from decimal import ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_UP, ROUND_UP, Decimal

from ..errors import DIV0, NUM, VALUE, ExcelError, XLError
from ..registry import flatten, fn, make_criterion, numbers, to_array
from ..values import MISSING, Array, RefValue, is_number, to_int, to_number


def _dec(x) -> Decimal:
    return Decimal(f"{float(x):.15g}") if type(x) is float else Decimal(x)


def _num(d: Decimal):
    f = float(d)
    return int(f) if f.is_integer() and abs(f) < 2**53 else f


def _round(x, digits, mode):
    x = to_number(x)
    digits = to_int(digits)
    q = Decimal(1).scaleb(-digits)
    return _num(_dec(x).quantize(q, rounding=mode)) if digits >= 0 else _num(
        (_dec(x) / Decimal(10) ** (-digits)).quantize(Decimal(1), rounding=mode) * Decimal(10) ** (-digits))


@fn("ROUND")
def ROUND(x, digits):
    return _round(x, digits, ROUND_HALF_UP)


@fn("ROUNDUP")
def ROUNDUP(x, digits):
    return _round(x, digits, ROUND_UP)


@fn("ROUNDDOWN")
def ROUNDDOWN(x, digits):
    return _round(x, digits, ROUND_DOWN)


@fn("TRUNC")
def TRUNC(x, digits=0):
    return _round(x, digits, ROUND_DOWN)


@fn("INT")
def INT(x):
    return math.floor(to_number(x))


@fn("MROUND")
def MROUND(x, m):
    x, m = to_number(x), to_number(m)
    if m == 0:
        return 0
    if (x > 0 > m) or (x < 0 < m):
        raise XLError(NUM)
    return _num((_dec(x) / _dec(m)).quantize(Decimal(1), rounding=ROUND_HALF_UP) * _dec(m))


def _ceil_floor(x, sig, up: bool):
    x = to_number(x)
    sig = to_number(sig)
    if sig == 0:
        return 0
    if x > 0 and sig < 0:
        raise XLError(NUM)
    q = _dec(x) / _dec(sig)
    r = q.quantize(Decimal(1), rounding=ROUND_UP if up else ROUND_DOWN) if x < 0 and sig < 0 else \
        q.quantize(Decimal(1), rounding=ROUND_CEILING if up else ROUND_FLOOR)
    return _num(r * _dec(sig))


@fn("CEILING")
def CEILING(x, sig=1):
    return _ceil_floor(x, sig, True)


@fn("FLOOR")
def FLOOR(x, sig=1):
    return _ceil_floor(x, sig, False)


@fn("CEILING.MATH", "CEILING.PRECISE", "ISO.CEILING")
def CEILING_MATH(x, sig=None, mode=None):
    x = to_number(x)
    sig = abs(to_number(sig)) if sig is not None else 1
    if sig == 0:
        return 0
    q = _dec(x) / _dec(sig)
    rounding = ROUND_CEILING
    if x < 0 and mode is not None and to_number(mode) != 0:
        rounding = ROUND_FLOOR
    return _num(q.quantize(Decimal(1), rounding=rounding) * _dec(sig))


@fn("FLOOR.MATH", "FLOOR.PRECISE")
def FLOOR_MATH(x, sig=None, mode=None):
    x = to_number(x)
    sig = abs(to_number(sig)) if sig is not None else 1
    if sig == 0:
        return 0
    q = _dec(x) / _dec(sig)
    rounding = ROUND_FLOOR
    if x < 0 and mode is not None and to_number(mode) != 0:
        rounding = ROUND_CEILING
    return _num(q.quantize(Decimal(1), rounding=rounding) * _dec(sig))


@fn("EVEN")
def EVEN(x):
    x = to_number(x)
    n = math.ceil(abs(x))
    if n % 2:
        n += 1
    return n if x >= 0 else -n


@fn("ODD")
def ODD(x):
    x = to_number(x)
    n = math.ceil(abs(x))
    if n % 2 == 0:
        n += 1
    return n if x >= 0 else -n


@fn("MOD")
def MOD(n, d):
    n, d = to_number(n), to_number(d)
    if d == 0:
        raise XLError(DIV0)
    r = n - d * math.floor(n / d)
    if type(r) is float:
        r = float(f"{r:.15g}")
    return r


@fn("QUOTIENT")
def QUOTIENT(n, d):
    n, d = to_number(n), to_number(d)
    if d == 0:
        raise XLError(DIV0)
    return int(n / d)


@fn("ABS")
def ABS(x):
    return abs(to_number(x))


@fn("SIGN")
def SIGN(x):
    x = to_number(x)
    return (x > 0) - (x < 0)


@fn("POWER")
def POWER(x, y):
    from ..evaluator import _pow

    return _pow(x, y)


@fn("SQRT")
def SQRT(x):
    x = to_number(x)
    if x < 0:
        raise XLError(NUM)
    r = math.sqrt(x)
    return int(r) if r.is_integer() else r


@fn("PI")
def PI():
    return math.pi


@fn("EXP")
def EXP(x):
    return math.exp(to_number(x))


@fn("LN")
def LN(x):
    x = to_number(x)
    if x <= 0:
        raise XLError(NUM)
    return math.log(x)


@fn("LOG")
def LOG(x, base=10):
    x, base = to_number(x), to_number(base)
    if x <= 0 or base <= 0:
        raise XLError(NUM)
    if base == 1:
        raise XLError(DIV0)
    r = math.log(x, base)
    rr = round(r)
    return rr if abs(r - rr) < 1e-12 else r


@fn("LOG10")
def LOG10(x):
    x = to_number(x)
    if x <= 0:
        raise XLError(NUM)
    r = math.log10(x)
    return int(r) if r.is_integer() else r


@fn("FACT")
def FACT(x):
    x = math.floor(to_number(x))
    if x < 0:
        raise XLError(NUM)
    if x > 170:
        raise XLError(NUM)
    return math.factorial(x)


@fn("COMBIN")
def COMBIN(n, k):
    n, k = to_int(n), to_int(k)
    if n < 0 or k < 0 or k > n:
        raise XLError(NUM)
    return math.comb(n, k)


@fn("PERMUT")
def PERMUT(n, k):
    n, k = to_int(n), to_int(k)
    if n < 0 or k < 0 or k > n:
        raise XLError(NUM)
    return math.perm(n, k)


@fn("GCD", kind="raw", min_args=1)
def GCD(ctx, *args):
    vals = [int(v) for v in numbers(ctx, args)]
    if any(v < 0 for v in vals):
        raise XLError(NUM)
    return math.gcd(*vals) if vals else 0


@fn("LCM", kind="raw", min_args=1)
def LCM(ctx, *args):
    vals = [int(v) for v in numbers(ctx, args)]
    if any(v < 0 for v in vals):
        raise XLError(NUM)
    return math.lcm(*vals) if vals else 0


for _name, _f in {
    "SIN": math.sin, "COS": math.cos, "TAN": math.tan, "ATAN": math.atan,
    "SINH": math.sinh, "COSH": math.cosh, "TANH": math.tanh, "DEGREES": math.degrees, "RADIANS": math.radians,
}.items():
    fn(_name)(lambda x, _f=_f: _f(to_number(x)))


@fn("ASIN")
def ASIN(x):
    x = to_number(x)
    if not -1 <= x <= 1:
        raise XLError(NUM)
    return math.asin(x)


@fn("ACOS")
def ACOS(x):
    x = to_number(x)
    if not -1 <= x <= 1:
        raise XLError(NUM)
    return math.acos(x)


@fn("ATAN2")
def ATAN2(x, y):
    x, y = to_number(x), to_number(y)
    if x == 0 and y == 0:
        raise XLError(DIV0)
    return math.atan2(y, x)


@fn("RAND", volatile=True)
def RAND():
    return random.random()


@fn("RANDBETWEEN", volatile=True)
def RANDBETWEEN(lo, hi):
    lo, hi = math.ceil(to_number(lo)), math.floor(to_number(hi))
    if lo > hi:
        raise XLError(NUM)
    return random.randint(lo, hi)


# ------------------------------------------------------------------ aggregation


def _sum(vals):
    s = math.fsum(vals)
    return int(s) if s.is_integer() and abs(s) < 2**53 and all(type(v) is int for v in vals) else s


@fn("SUM", kind="raw")
def SUM(ctx, *args):
    return _sum(numbers(ctx, args))


@fn("PRODUCT", kind="raw", min_args=1)
def PRODUCT(ctx, *args):
    vals = numbers(ctx, args)
    if not vals:
        return 0
    p = 1
    for v in vals:
        p *= v
    return p


@fn("SUMSQ", kind="raw", min_args=1)
def SUMSQ(ctx, *args):
    return _sum([v * v for v in numbers(ctx, args)])


@fn("SUMPRODUCT", kind="raw", min_args=1)
def SUMPRODUCT(ctx, *args):
    arrs = [to_array(ctx, a) for a in args]
    nr, nc = arrs[0].nrows, arrs[0].ncols
    for a in arrs:
        if a.nrows != nr or a.ncols != nc:
            # clipped whole-column refs may differ: pad to the largest shape with blanks
            nr, nc = max(nr, a.nrows), max(nc, a.ncols)
    total = 0.0
    all_int = True
    for i in range(nr):
        for j in range(nc):
            p = 1
            for a in arrs:
                v = a.rows[i][j] if i < a.nrows and j < a.ncols else None
                if isinstance(v, ExcelError):
                    raise XLError(v)
                if not is_number(v):
                    p = 0
                    break
                if type(v) is float:
                    all_int = False
                p *= v
            total += p
    return int(total) if all_int and total.is_integer() else total


def criteria_pairs(ctx, pairs):
    """[(range, criteria), ...] -> list of (Array, predicate) with equal shapes."""
    out = []
    for rng, crit in pairs:
        if isinstance(crit, (RefValue, Array)):
            from ..registry import to_scalar

            crit = to_scalar(ctx, crit)
        out.append((to_array(ctx, rng) if isinstance(rng, (RefValue, Array)) else Array([[rng]]),
                     make_criterion(crit)))
    return out


def _cells_matching(ctx, ranges, shape_ref=None):
    """Yield (i, j) positions where every (Array, predicate) matches."""
    arrs = [a for a, _ in ranges]
    nr = max(a.nrows for a in arrs)
    nc = max(a.ncols for a in arrs)
    if shape_ref is not None:
        nr, nc = max(nr, shape_ref.nrows), max(nc, shape_ref.ncols)
    for i in range(nr):
        for j in range(nc):
            ok = True
            for a, pred in ranges:
                v = a.rows[i][j] if i < a.nrows and j < a.ncols else None
                if not pred(v):
                    ok = False
                    break
            if ok:
                yield i, j


def _value_at(arr: Array, i, j):
    return arr.rows[i][j] if i < arr.nrows and j < arr.ncols else None


def _resize_like(ctx, target, like):
    """SUMIF semantics: sum_range is resized to the criteria range's shape from its top-left."""
    if isinstance(target, RefValue) and isinstance(like, RefValue):
        return to_array(ctx, RefValue(target.sheet, target.r1, target.c1,
                                      target.r1 + like.nrows - 1, target.c1 + like.ncols - 1))
    return to_array(ctx, target)


def _matching_values(ctx, value_range, pairs):
    first = pairs[0][0]
    target = _resize_like(ctx, value_range, first) if pairs else to_array(ctx, value_range)
    crit = criteria_pairs(ctx, pairs)
    out = []
    for i, j in _cells_matching(ctx, crit, target):
        v = _value_at(target, i, j)
        if isinstance(v, ExcelError):
            raise XLError(v)
        out.append(v)
    return out


@fn("SUMIF", kind="raw", min_args=2, max_args=3)
def SUMIF(ctx, rng, crit, sum_range=MISSING):
    vals = _matching_values(ctx, rng if sum_range is MISSING else sum_range, [(rng, crit)])
    return _sum([v for v in vals if is_number(v)])


def _pairs(args):
    if len(args) % 2:
        raise XLError(VALUE)
    return [(args[i], args[i + 1]) for i in range(0, len(args), 2)]


@fn("SUMIFS", kind="raw", min_args=3)
def SUMIFS(ctx, sum_range, *args):
    vals = _matching_values(ctx, sum_range, _pairs(args))
    return _sum([v for v in vals if is_number(v)])


@fn("SUBTOTAL", kind="raw", min_args=2)
def SUBTOTAL(ctx, code, *refs):
    from . import stats_fns as st
    from ..registry import scalar_arg

    n = to_int(scalar_arg(ctx, code)) % 100
    table = {1: st.AVERAGE, 2: st.COUNT, 3: st.COUNTA, 4: st.MAX, 5: st.MIN, 6: PRODUCT,
             7: st.STDEV, 8: st.STDEVP, 9: SUM, 10: st.VAR, 11: st.VARP}
    f = table.get(n)
    if f is None:
        raise XLError(VALUE)
    return f(ctx, *refs)


@fn("AGGREGATE", kind="raw", min_args=3)
def AGGREGATE(ctx, code, options, *refs):
    from . import stats_fns as st
    from ..registry import scalar_arg

    n = to_int(scalar_arg(ctx, code))
    simple = {1: st.AVERAGE, 2: st.COUNT, 3: st.COUNTA, 4: st.MAX, 5: st.MIN, 6: PRODUCT,
              7: st.STDEV, 8: st.STDEVP, 9: SUM, 10: st.VAR, 11: st.VARP, 12: st.MEDIAN}
    opt = to_int(scalar_arg(ctx, options))
    ignore_errors = opt in (2, 3, 6, 7)
    if n in simple:
        if ignore_errors:
            vals = [v for v, _ in flatten(ctx, refs) if is_number(v)]
            return simple[n](ctx, Array([vals or [None]]))
        return simple[n](ctx, *refs)
    if n in (14, 15, 16, 17, 18, 19):
        data = [v for v, _ in flatten(ctx, refs[:1]) if is_number(v) or not ignore_errors]
        k = refs[1] if len(refs) > 1 else MISSING
        arr = Array([data or [None]])
        f = {14: st.LARGE, 15: st.SMALL, 16: st.PERCENTILE, 17: st.QUARTILE,
             18: st.PERCENTILE_EXC, 19: st.QUARTILE_EXC}[n]
        return f(ctx, arr, k)
    raise XLError(VALUE)
