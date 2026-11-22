"""Statistical and counting functions."""
from __future__ import annotations

import bisect
import math
from collections import Counter

from .. import special as S
from ..errors import DIV0, NA, NUM, ExcelError, XLError
from ..registry import flatten, fn, numbers, scalar_arg, to_array
from ..values import (
    MISSING, Array, RefValue, compare, is_number, parse_number_text, to_bool, to_int, to_number,
)
from .math_fns import _cells_matching, _matching_values, _pairs, criteria_pairs, lifts_criteria


@fn("AVERAGE", kind="raw", min_args=1)
def AVERAGE(ctx, *args):
    vals = numbers(ctx, args)
    if not vals:
        raise XLError(DIV0)
    return math.fsum(vals) / len(vals)


@fn("AVERAGEA", kind="raw", min_args=1)
def AVERAGEA(ctx, *args):
    vals = _a_values(ctx, args)
    if not vals:
        raise XLError(DIV0)
    return math.fsum(vals) / len(vals)


@fn("COUNT", kind="raw")
def COUNT(ctx, *args):
    n = 0
    for v, direct in flatten(ctx, args):
        if is_number(v):
            n += 1
        elif direct and (type(v) is bool or (type(v) is str and parse_number_text(v) is not None)):
            n += 1
    return n


@fn("COUNTA", kind="raw")
def COUNTA(ctx, *args):
    n = 0
    for a in args:
        if a is MISSING:
            n += 1
            continue
        for v, _ in flatten(ctx, [a]):
            if v is not None:
                n += 1
    return n


@fn("COUNTBLANK", kind="raw", min_args=1, max_args=1)
def COUNTBLANK(ctx, rng):
    if isinstance(rng, RefValue):
        area = rng.nrows * rng.ncols
        filled = 0
        for v, _ in flatten(ctx, [rng]):
            if v is not None and v != "":
                filled += 1
        return area - filled
    arr = to_array(ctx, rng)
    return sum(1 for v in arr.flat() if v is None or v == "")


@fn("COUNTIF", kind="raw", min_args=2, max_args=2)
@lifts_criteria(1)
def COUNTIF(ctx, rng, crit):
    return COUNTIFS(ctx, rng, crit)


@fn("COUNTIFS", kind="raw", min_args=2)
@lifts_criteria(1)
def COUNTIFS(ctx, *args):
    pairs = _pairs(args)
    crit = criteria_pairs(ctx, pairs)
    n = sum(1 for _ in _cells_matching(ctx, crit))
    # blanks beyond the used area of whole-column references
    first = pairs[0][0]
    if isinstance(first, RefValue):
        arr = crit[0][0]
        extra = first.nrows * first.ncols - arr.nrows * arr.ncols
        if extra > 0 and all(p(None) for _, p in crit):
            n += extra
    return n


@fn("AVERAGEIF", kind="raw", min_args=2, max_args=3)
@lifts_criteria(1)
def AVERAGEIF(ctx, rng, crit, avg_range=MISSING):
    vals = [v for v in _matching_values(ctx, rng if avg_range is MISSING else avg_range, [(rng, crit)])
            if is_number(v)]
    if not vals:
        raise XLError(DIV0)
    return math.fsum(vals) / len(vals)


@fn("AVERAGEIFS", kind="raw", min_args=3)
@lifts_criteria(2)
def AVERAGEIFS(ctx, avg_range, *args):
    vals = [v for v in _matching_values(ctx, avg_range, _pairs(args)) if is_number(v)]
    if not vals:
        raise XLError(DIV0)
    return math.fsum(vals) / len(vals)


@fn("MAX", kind="raw", min_args=1)
def MAX(ctx, *args):
    vals = numbers(ctx, args)
    return max(vals) if vals else 0


@fn("MIN", kind="raw", min_args=1)
def MIN(ctx, *args):
    vals = numbers(ctx, args)
    return min(vals) if vals else 0


@fn("MAXA", kind="raw", min_args=1)
def MAXA(ctx, *args):
    vals = _a_values(ctx, args)
    return max(vals) if vals else 0


@fn("MINA", kind="raw", min_args=1)
def MINA(ctx, *args):
    vals = _a_values(ctx, args)
    return min(vals) if vals else 0


@fn("MAXIFS", kind="raw", min_args=3)
@lifts_criteria(2)
def MAXIFS(ctx, rng, *args):
    vals = [v for v in _matching_values(ctx, rng, _pairs(args)) if is_number(v)]
    return max(vals) if vals else 0


@fn("MINIFS", kind="raw", min_args=3)
@lifts_criteria(2)
def MINIFS(ctx, rng, *args):
    vals = [v for v in _matching_values(ctx, rng, _pairs(args)) if is_number(v)]
    return min(vals) if vals else 0


@fn("MEDIAN", kind="raw", min_args=1)
def MEDIAN(ctx, *args):
    vals = sorted(numbers(ctx, args))
    if not vals:
        raise XLError(NUM)
    n = len(vals)
    mid = n // 2
    return vals[mid] if n % 2 else (vals[mid - 1] + vals[mid]) / 2


@fn("MODE", "MODE.SNGL", kind="raw", min_args=1)
def MODE(ctx, *args):
    vals = numbers(ctx, args)
    if not vals:
        raise XLError(NA)
    counts = Counter(vals)
    best = max(counts.values())
    if best < 2:
        raise XLError(NA)
    for v in vals:
        if counts[v] == best:
            return v


def _var(vals, sample: bool):
    n = len(vals)
    if n < (2 if sample else 1):
        raise XLError(DIV0)
    mean = math.fsum(vals) / n
    return math.fsum((v - mean) ** 2 for v in vals) / (n - 1 if sample else n)


@fn("VAR", "VAR.S", kind="raw", min_args=1)
def VAR(ctx, *args):
    return _var(numbers(ctx, args), True)


@fn("VARP", "VAR.P", kind="raw", min_args=1)
def VARP(ctx, *args):
    return _var(numbers(ctx, args), False)


@fn("STDEV", "STDEV.S", kind="raw", min_args=1)
def STDEV(ctx, *args):
    return math.sqrt(_var(numbers(ctx, args), True))


@fn("STDEVP", "STDEV.P", kind="raw", min_args=1)
def STDEVP(ctx, *args):
    return math.sqrt(_var(numbers(ctx, args), False))


@fn("LARGE", kind="raw", min_args=2, max_args=2)
def LARGE(ctx, arr, k):
    vals = sorted(numbers(ctx, [arr]), reverse=True)
    k = math.ceil(to_number(scalar_arg(ctx, k)))
    if k < 1 or k > len(vals):
        raise XLError(NUM)
    return vals[k - 1]


@fn("SMALL", kind="raw", min_args=2, max_args=2)
def SMALL(ctx, arr, k):
    vals = sorted(numbers(ctx, [arr]))
    k = math.ceil(to_number(scalar_arg(ctx, k)))
    if k < 1 or k > len(vals):
        raise XLError(NUM)
    return vals[k - 1]


@fn("RANK", "RANK.EQ", kind="raw", min_args=2, max_args=3)
def RANK(ctx, x, ref, order=MISSING):
    x = to_number(scalar_arg(ctx, x))
    vals = numbers(ctx, [ref])
    asc = order is not MISSING and to_number(scalar_arg(ctx, order)) != 0
    if not any(compare(v, x) == 0 for v in vals):
        raise XLError(NA)
    if asc:
        return 1 + sum(1 for v in vals if compare(v, x) < 0)
    return 1 + sum(1 for v in vals if compare(v, x) > 0)


@fn("RANK.AVG", kind="raw", min_args=2, max_args=3)
def RANK_AVG(ctx, x, ref, order=MISSING):
    first = RANK(ctx, x, ref, order)
    xv = to_number(scalar_arg(ctx, x))
    ties = sum(1 for v in numbers(ctx, [ref]) if compare(v, xv) == 0)
    return first + (ties - 1) / 2


def _percentile_inc(vals, p):
    if not vals or p < 0 or p > 1:
        raise XLError(NUM)
    vals = sorted(vals)
    h = (len(vals) - 1) * p
    lo = math.floor(h)
    hi = min(lo + 1, len(vals) - 1)
    return vals[lo] + (h - lo) * (vals[hi] - vals[lo])


def _percentile_exc(vals, p):
    vals = sorted(vals)
    n = len(vals)
    if not vals or p <= 0 or p >= 1:
        raise XLError(NUM)
    h = (n + 1) * p
    if h < 1 or h > n:
        raise XLError(NUM)
    lo = math.floor(h)
    if lo >= n:
        return vals[-1]
    return vals[lo - 1] + (h - lo) * (vals[lo] - vals[lo - 1])


@fn("PERCENTILE", "PERCENTILE.INC", kind="raw", min_args=2, max_args=2)
def PERCENTILE(ctx, arr, p):
    return _percentile_inc(numbers(ctx, [arr]), to_number(scalar_arg(ctx, p)))


@fn("PERCENTILE.EXC", kind="raw", min_args=2, max_args=2)
def PERCENTILE_EXC(ctx, arr, p):
    return _percentile_exc(numbers(ctx, [arr]), to_number(scalar_arg(ctx, p)))


@fn("QUARTILE", "QUARTILE.INC", kind="raw", min_args=2, max_args=2)
def QUARTILE(ctx, arr, q):
    q = to_int(scalar_arg(ctx, q))
    if q < 0 or q > 4:
        raise XLError(NUM)
    return _percentile_inc(numbers(ctx, [arr]), q / 4)


@fn("QUARTILE.EXC", kind="raw", min_args=2, max_args=2)
def QUARTILE_EXC(ctx, arr, q):
    q = to_int(scalar_arg(ctx, q))
    if q < 1 or q > 3:
        raise XLError(NUM)
    return _percentile_exc(numbers(ctx, [arr]), q / 4)


@fn("CORREL", "PEARSON", kind="raw", min_args=2, max_args=2)
def CORREL(ctx, a, b):
    xs, ys = _paired(ctx, a, b)
    if len(xs) < 2:
        raise XLError(DIV0)
    mx, my = math.fsum(xs) / len(xs), math.fsum(ys) / len(ys)
    sxy = math.fsum((x - mx) * (y - my) for x, y in zip(xs, ys))
    sxx = math.fsum((x - mx) ** 2 for x in xs)
    syy = math.fsum((y - my) ** 2 for y in ys)
    if sxx == 0 or syy == 0:
        raise XLError(DIV0)
    return sxy / math.sqrt(sxx * syy)


def _paired(ctx, a, b):
    va, vb = list(to_array(ctx, a).flat()), list(to_array(ctx, b).flat())
    if len(va) != len(vb):
        raise XLError(NA)
    xs, ys = [], []
    for x, y in zip(va, vb):
        if is_number(x) and is_number(y):
            xs.append(x)
            ys.append(y)
    return xs, ys


@fn("SLOPE", kind="raw", min_args=2, max_args=2)
def SLOPE(ctx, ys, xs):
    x, y = _paired(ctx, xs, ys)
    if len(x) < 2:
        raise XLError(DIV0)
    mx, my = math.fsum(x) / len(x), math.fsum(y) / len(y)
    sxx = math.fsum((a - mx) ** 2 for a in x)
    if sxx == 0:
        raise XLError(DIV0)
    return math.fsum((a - mx) * (b - my) for a, b in zip(x, y)) / sxx


@fn("INTERCEPT", kind="raw", min_args=2, max_args=2)
def INTERCEPT(ctx, ys, xs):
    x, y = _paired(ctx, xs, ys)
    m = SLOPE(ctx, ys, xs)
    return math.fsum(y) / len(y) - m * math.fsum(x) / len(x)


@fn("GEOMEAN", kind="raw", min_args=1)
def GEOMEAN(ctx, *args):
    vals = numbers(ctx, args)
    if not vals or any(v <= 0 for v in vals):
        raise XLError(NUM)
    return math.exp(math.fsum(math.log(v) for v in vals) / len(vals))


@fn("BINOM.DIST", "BINOMDIST", min_args=4, max_args=4)
def BINOM_DIST(number_s, trials, probability_s, cumulative):
    k, n = to_int(number_s), to_int(trials)
    p = to_number(probability_s)
    if n < 0 or k < 0 or k > n or not 0.0 <= p <= 1.0:
        raise XLError(NUM)
    return S.binom_cdf(k, n, p) if to_bool(cumulative) else S.binom_pmf(k, n, p)


# ------------------------------------------------------------------ dispersion and shape


def _numbers_or(ctx, args, err):
    vals = numbers(ctx, args)
    if not vals:
        raise XLError(err)
    return vals


@fn("AVEDEV", kind="raw", min_args=1)
def AVEDEV(ctx, *args):
    vals = _numbers_or(ctx, args, NUM)
    m = math.fsum(vals) / len(vals)
    return math.fsum(abs(v - m) for v in vals) / len(vals)


@fn("DEVSQ", kind="raw", min_args=1)
def DEVSQ(ctx, *args):
    vals = _numbers_or(ctx, args, NUM)
    m = math.fsum(vals) / len(vals)
    return math.fsum((v - m) ** 2 for v in vals)


def _std_moments(vals, sample: bool):
    """Standardized deviations (x - mean) / sd, with the sample or population sd."""
    n = len(vals)
    sd = math.sqrt(_var(vals, sample))
    if sd == 0:
        raise XLError(DIV0)
    m = math.fsum(vals) / n
    return [(v - m) / sd for v in vals]


@fn("SKEW", kind="raw", min_args=1)
def SKEW(ctx, *args):
    vals = numbers(ctx, args)
    n = len(vals)
    if n < 3:
        raise XLError(DIV0)
    z = _std_moments(vals, True)
    return n / ((n - 1) * (n - 2)) * math.fsum(v**3 for v in z)


@fn("SKEW.P", kind="raw", min_args=1)
def SKEW_P(ctx, *args):
    vals = numbers(ctx, args)
    if len(vals) < 3:
        raise XLError(DIV0)
    z = _std_moments(vals, False)
    return math.fsum(v**3 for v in z) / len(vals)


@fn("KURT", kind="raw", min_args=1)
def KURT(ctx, *args):
    vals = numbers(ctx, args)
    n = len(vals)
    if n < 4:
        raise XLError(DIV0)
    z = _std_moments(vals, True)
    return (n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * math.fsum(v**4 for v in z)
            - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3)))


@fn("HARMEAN", kind="raw", min_args=1)
def HARMEAN(ctx, *args):
    vals = _numbers_or(ctx, args, NUM)
    if any(v <= 0 for v in vals):
        raise XLError(NUM)
    return len(vals) / math.fsum(1 / v for v in vals)


@fn("TRIMMEAN", kind="raw", min_args=2, max_args=2)
def TRIMMEAN(ctx, arr, percent):
    vals = sorted(numbers(ctx, [arr]))
    p = to_number(scalar_arg(ctx, percent))
    if not vals or p < 0 or p >= 1:
        raise XLError(NUM)
    k = int(len(vals) * p / 2 + 1e-12)  # points dropped from each end
    kept = vals[k:len(vals) - k]
    return math.fsum(kept) / len(kept)


@fn("FISHER", min_args=1, max_args=1)
def FISHER(x):
    x = to_number(x)
    if x <= -1 or x >= 1:
        raise XLError(NUM)
    return math.atanh(x)


@fn("FISHERINV", min_args=1, max_args=1)
def FISHERINV(y):
    return math.tanh(to_number(y))


@fn("PERMUTATIONA", min_args=2, max_args=2)
def PERMUTATIONA(number, chosen):
    n, k = to_int(number), to_int(chosen)
    if n < 0 or k < 0:
        raise XLError(NUM)
    return n**k


# ------------------------------------------------------------------ "A" variants (text and logicals count)


def _a_values(ctx, args):
    """Values for AVERAGEA/STDEVA-style functions: TRUE=1, FALSE=0, text in ranges = 0."""
    vals = []
    for v, direct in flatten(ctx, args):
        if isinstance(v, ExcelError):
            raise XLError(v)
        if is_number(v):
            vals.append(v)
        elif type(v) is bool:
            vals.append(int(v))
        elif type(v) is str:
            vals.append(to_number(v) if direct else 0)
    return vals


@fn("VARA", kind="raw", min_args=1)
def VARA(ctx, *args):
    return _var(_a_values(ctx, args), True)


@fn("VARPA", kind="raw", min_args=1)
def VARPA(ctx, *args):
    return _var(_a_values(ctx, args), False)


@fn("STDEVA", kind="raw", min_args=1)
def STDEVA(ctx, *args):
    return math.sqrt(_var(_a_values(ctx, args), True))


@fn("STDEVPA", kind="raw", min_args=1)
def STDEVPA(ctx, *args):
    return math.sqrt(_var(_a_values(ctx, args), False))


# ------------------------------------------------------------------ ranks, modes, frequencies


def _percentrank(ctx, arr, x, sig, exclusive: bool):
    vals = sorted(numbers(ctx, [arr]))
    x = to_number(scalar_arg(ctx, x))
    digits = 3 if sig is MISSING else to_int(scalar_arg(ctx, sig))
    if not vals or digits < 1:
        raise XLError(NUM)
    n = len(vals)
    if x < vals[0] or x > vals[-1]:
        raise XLError(NA)
    below = sum(1 for v in vals if v < x)
    if vals[below] == x:
        pos = float(below)
    else:  # interpolate between the neighbours vals[below-1] < x < vals[below]
        lo, hi = vals[below - 1], vals[below]
        pos = below - 1 + (x - lo) / (hi - lo)
    if exclusive:
        r = (pos + 1) / (n + 1)
    elif n == 1:
        r = 1.0
    else:
        r = pos / (n - 1)
    scale = 10.0**digits
    return math.floor(round(r * scale, 9)) / scale  # truncated, not rounded


@fn("PERCENTRANK", "PERCENTRANK.INC", kind="raw", min_args=2, max_args=3)
def PERCENTRANK(ctx, arr, x, significance=MISSING):
    return _percentrank(ctx, arr, x, significance, False)


@fn("PERCENTRANK.EXC", kind="raw", min_args=2, max_args=3)
def PERCENTRANK_EXC(ctx, arr, x, significance=MISSING):
    return _percentrank(ctx, arr, x, significance, True)


@fn("MODE.MULT", kind="raw", min_args=1)
def MODE_MULT(ctx, *args):
    vals = numbers(ctx, args)
    counts = Counter(vals)
    best = max(counts.values(), default=0)
    if best < 2:
        raise XLError(NA)
    out, seen = [], set()
    for v in vals:
        if counts[v] == best and v not in seen:
            seen.add(v)
            out.append([v])
    return Array(out)


@fn("FREQUENCY", kind="raw", min_args=2, max_args=2)
def FREQUENCY(ctx, data, bins):
    vals = numbers(ctx, [data])
    edges = numbers(ctx, [bins])
    order = sorted(range(len(edges)), key=lambda i: edges[i])
    counts = [0] * (len(edges) + 1)
    sorted_edges = [edges[i] for i in order]
    for v in vals:
        j = bisect.bisect_left(sorted_edges, v)
        counts[order[j] if j < len(order) else len(edges)] += 1
    return Array([[c] for c in counts])


@fn("PROB", kind="raw", min_args=3, max_args=4)
def PROB(ctx, x_range, prob_range, lower, upper=MISSING):
    xs, ps = list(to_array(ctx, x_range).flat()), list(to_array(ctx, prob_range).flat())
    if len(xs) != len(ps):
        raise XLError(NA)
    lo = to_number(scalar_arg(ctx, lower))
    hi = lo if upper is MISSING else to_number(scalar_arg(ctx, upper))
    total, hit = [], []
    for x, p in zip(xs, ps):
        if isinstance(p, ExcelError):
            raise XLError(p)
        if not is_number(p):
            continue
        if p < 0 or p > 1:
            raise XLError(NUM)
        total.append(p)
        if is_number(x) and lo <= x <= hi:
            hit.append(p)
    if abs(math.fsum(total) - 1) > 1e-7:
        raise XLError(NUM)
    return math.fsum(hit)


# ------------------------------------------------------------------ covariance and simple regression


def _centered(ctx, a, b, minimum: int):
    xs, ys = _paired(ctx, a, b)
    n = len(xs)
    if n < minimum:
        raise XLError(DIV0)
    mx, my = math.fsum(xs) / n, math.fsum(ys) / n
    return [x - mx for x in xs], [y - my for y in ys], mx, my


@fn("COVAR", "COVARIANCE.P", kind="raw", min_args=2, max_args=2)
def COVAR(ctx, a, b):
    dx, dy, _, _ = _centered(ctx, a, b, 1)
    return math.fsum(x * y for x, y in zip(dx, dy)) / len(dx)


@fn("COVARIANCE.S", kind="raw", min_args=2, max_args=2)
def COVARIANCE_S(ctx, a, b):
    dx, dy, _, _ = _centered(ctx, a, b, 2)
    return math.fsum(x * y for x, y in zip(dx, dy)) / (len(dx) - 1)


@fn("RSQ", kind="raw", min_args=2, max_args=2)
def RSQ(ctx, ys, xs):
    return CORREL(ctx, ys, xs) ** 2


@fn("STEYX", kind="raw", min_args=2, max_args=2)
def STEYX(ctx, ys, xs):
    dx, dy, _, _ = _centered(ctx, xs, ys, 3)
    sxx = math.fsum(x * x for x in dx)
    if sxx == 0:
        raise XLError(DIV0)
    sxy = math.fsum(x * y for x, y in zip(dx, dy))
    ss = math.fsum(y * y for y in dy) - sxy * sxy / sxx
    return math.sqrt(max(ss, 0.0) / (len(dx) - 2))


@fn("FORECAST", "FORECAST.LINEAR", kind="raw", min_args=3, max_args=3)
def FORECAST(ctx, x, ys, xs):
    xv = to_number(scalar_arg(ctx, x))
    dx, dy, mx, my = _centered(ctx, xs, ys, 1)
    sxx = math.fsum(v * v for v in dx)
    if sxx == 0:
        raise XLError(DIV0)
    slope = math.fsum(a * b for a, b in zip(dx, dy)) / sxx
    return my + slope * (xv - mx)
