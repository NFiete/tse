"""Special functions and probability distributions used by the statistical functions.

Pure Python (math only): regularized incomplete gamma and beta functions, the CDFs and
survival functions built on them, and a bracketing root finder for their inverses.
Survival functions are computed directly (not as 1 - cdf) so right-tail probabilities
keep full relative precision.
"""
from __future__ import annotations

import math
from statistics import NormalDist

_EPS = 1e-16
_TINY = 1e-300
_MAXIT = 10000
_STD = NormalDist()
SQRT2 = math.sqrt(2.0)


# ------------------------------------------------------------------ normal


def norm_cdf(z: float) -> float:
    return 0.5 * math.erfc(-z / SQRT2)


def norm_sf(z: float) -> float:
    return 0.5 * math.erfc(z / SQRT2)


def norm_pdf(z: float) -> float:
    return math.exp(-0.5 * z * z) / math.sqrt(2.0 * math.pi)


def norm_ppf(p: float) -> float:
    """Inverse standard normal CDF for 0 < p < 1."""
    return _STD.inv_cdf(p)


# ------------------------------------------------------------------ incomplete gamma


def _gamma_series(a: float, x: float) -> float:
    """Regularized lower incomplete gamma P(a, x) by its power series (x < a + 1)."""
    term = total = 1.0 / a
    ap = a
    for _ in range(_MAXIT):
        ap += 1.0
        term *= x / ap
        total += term
        if abs(term) < abs(total) * _EPS:
            break
    return total * math.exp(-x + a * math.log(x) - math.lgamma(a))


def _gamma_cf(a: float, x: float) -> float:
    """Regularized upper incomplete gamma Q(a, x) by continued fraction (x >= a + 1)."""
    b = x + 1.0 - a
    c = 1.0 / _TINY
    d = 1.0 / b
    h = d
    for i in range(1, _MAXIT):
        an = -i * (i - a)
        b += 2.0
        d = an * d + b
        if abs(d) < _TINY:
            d = _TINY
        c = b + an / c
        if abs(c) < _TINY:
            c = _TINY
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < _EPS:
            break
    return math.exp(-x + a * math.log(x) - math.lgamma(a)) * h


def gammainc(a: float, x: float) -> float:
    """Regularized lower incomplete gamma P(a, x)."""
    if x <= 0.0:
        return 0.0
    if math.isinf(x):
        return 1.0
    if x < a + 1.0:
        return _gamma_series(a, x)
    return 1.0 - _gamma_cf(a, x)


def gammaincc(a: float, x: float) -> float:
    """Regularized upper incomplete gamma Q(a, x) = 1 - P(a, x)."""
    if x <= 0.0:
        return 1.0
    if math.isinf(x):
        return 0.0
    if x < a + 1.0:
        return 1.0 - _gamma_series(a, x)
    return _gamma_cf(a, x)


# ------------------------------------------------------------------ incomplete beta


def _beta_cf(a: float, b: float, x: float) -> float:
    """Continued fraction for the incomplete beta function (modified Lentz)."""
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < _TINY:
        d = _TINY
    d = 1.0 / d
    h = d
    for m in range(1, _MAXIT):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < _TINY:
            d = _TINY
        c = 1.0 + aa / c
        if abs(c) < _TINY:
            c = _TINY
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < _TINY:
            d = _TINY
        c = 1.0 + aa / c
        if abs(c) < _TINY:
            c = _TINY
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < _EPS:
            break
    return h


def lbeta(a: float, b: float) -> float:
    return math.lgamma(a) + math.lgamma(b) - math.lgamma(a + b)


def betainc(a: float, b: float, x: float, y: float | None = None) -> float:
    """Regularized incomplete beta I_x(a, b). Pass y = 1 - x when it is known more
    precisely than 1 - x would compute it."""
    if y is None:
        y = 1.0 - x
    if x <= 0.0:
        return 0.0
    if y <= 0.0:
        return 1.0
    front = math.exp(a * math.log(x) + b * math.log(y) - lbeta(a, b))
    if x < (a + 1.0) / (a + b + 2.0):
        return front * _beta_cf(a, b, x) / a
    return 1.0 - front * _beta_cf(b, a, y) / b


def betaincc(a: float, b: float, x: float, y: float | None = None) -> float:
    """1 - I_x(a, b), computed without cancellation."""
    if y is None:
        y = 1.0 - x
    return betainc(b, a, y, x)


# ------------------------------------------------------------------ root finding


def brent(f, lo: float, hi: float, flo: float | None = None, fhi: float | None = None) -> float:
    """Root of f in [lo, hi], where f(lo) and f(hi) have opposite signs (Brent's method)."""
    a, b = lo, hi
    fa = f(a) if flo is None else flo
    fb = f(b) if fhi is None else fhi
    if fa == 0.0:
        return a
    if fb == 0.0:
        return b
    if (fa > 0) == (fb > 0):
        raise ValueError("root not bracketed")
    c, fc = a, fa
    d = e = b - a
    for _ in range(500):
        if (fb > 0) == (fc > 0):
            c, fc = a, fa
            d = e = b - a
        if abs(fc) < abs(fb):
            a, b, c = b, c, b
            fa, fb, fc = fb, fc, fb
        tol = 2.0 * 2.2e-16 * abs(b) + 1e-300
        m = 0.5 * (c - b)
        if abs(m) <= tol or fb == 0.0:
            return b
        if abs(e) >= tol and abs(fa) > abs(fb):
            s = fb / fa
            if a == c:
                p, q = 2.0 * m * s, 1.0 - s
            else:
                q, r = fa / fc, fb / fc
                p = s * (2.0 * m * q * (q - r) - (b - a) * (r - 1.0))
                q = (q - 1.0) * (r - 1.0) * (s - 1.0)
            if p > 0:
                q = -q
            else:
                p = -p
            if 2.0 * p < min(3.0 * m * q - abs(tol * q), abs(e * q)):
                e, d = d, p / q
            else:
                d = e = m
        else:
            d = e = m
        a, fa = b, fb
        b += d if abs(d) > tol else (tol if m > 0 else -tol)
        fb = f(b)
    return b


def invert(fn, target: float, lo: float = 0.0, hi: float | None = None, increasing: bool = True) -> float:
    """Solve fn(x) = target for a monotone fn on [lo, inf) (or [lo, hi])."""
    sign = 1.0 if increasing else -1.0

    def g(x):
        return sign * (fn(x) - target)

    glo = g(lo)
    if hi is None:
        hi = max(1.0, 2.0 * abs(lo))
        ghi = g(hi)
        while ghi < 0:
            lo, glo = hi, ghi
            hi *= 2.0
            if hi > 1e300:
                raise OverflowError
            ghi = g(hi)
    else:
        ghi = g(hi)
    return brent(g, lo, hi, glo, ghi)


# ------------------------------------------------------------------ distributions


def t_sf2(t: float, df: float) -> float:
    """Two-tailed Student t probability P(|T| > |t|)."""
    t2 = t * t
    return betainc(df / 2.0, 0.5, df / (df + t2), t2 / (df + t2))


def t_cdf(t: float, df: float) -> float:
    tail = 0.5 * t_sf2(t, df)
    return 1.0 - tail if t > 0 else tail


def t_pdf(t: float, df: float) -> float:
    return math.exp(math.lgamma((df + 1) / 2) - math.lgamma(df / 2)
                    - 0.5 * math.log(df * math.pi) - (df + 1) / 2 * math.log1p(t * t / df))


def t_ppf2(p: float, df: float) -> float:
    """t >= 0 with two-tailed probability p (0 < p <= 1)."""
    if p >= 1.0:
        return 0.0
    if df == 1.0:
        return 1.0 / math.tan(math.pi * p / 2.0)
    if df == 2.0:
        return math.sqrt(2.0 / (p * (2.0 - p)) - 2.0)
    return invert(lambda t: t_sf2(t, df), p, 0.0, None, increasing=False)


def t_ppf(p: float, df: float) -> float:
    if p == 0.5:
        return 0.0
    t = t_ppf2(2.0 * min(p, 1.0 - p), df)
    return t if p > 0.5 else -t


def chi2_cdf(x: float, df: float) -> float:
    return gammainc(df / 2.0, x / 2.0)


def chi2_sf(x: float, df: float) -> float:
    return gammaincc(df / 2.0, x / 2.0)


def gamma_pdf(x: float, a: float, scale: float) -> float:
    if x == 0.0:
        if a < 1.0:
            raise ZeroDivisionError
        return 1.0 / scale if a == 1.0 else 0.0
    return math.exp((a - 1.0) * math.log(x) - x / scale - a * math.log(scale) - math.lgamma(a))


def gamma_ppf(p: float, a: float, scale: float = 1.0, upper: bool = False) -> float:
    if upper:
        return scale * invert(lambda x: gammaincc(a, x), p, 0.0, None, increasing=False)
    return scale * invert(lambda x: gammainc(a, x), p, 0.0, None)


def f_cdf(x: float, d1: float, d2: float) -> float:
    if x <= 0.0:
        return 0.0
    return betainc(d1 / 2.0, d2 / 2.0, d1 * x / (d1 * x + d2), d2 / (d1 * x + d2))


def f_sf(x: float, d1: float, d2: float) -> float:
    if x <= 0.0:
        return 1.0
    return betaincc(d1 / 2.0, d2 / 2.0, d1 * x / (d1 * x + d2), d2 / (d1 * x + d2))


def f_pdf(x: float, d1: float, d2: float) -> float:
    if x == 0.0:
        if d1 < 2.0:
            raise ZeroDivisionError
        return 1.0 if d1 == 2.0 else 0.0
    return math.exp(0.5 * (d1 * math.log(d1 * x) + d2 * math.log(d2) - (d1 + d2) * math.log(d1 * x + d2))
                    - math.log(x) - lbeta(d1 / 2.0, d2 / 2.0))


def beta_pdf(x: float, a: float, b: float) -> float:
    if x == 0.0 or x == 1.0:
        e = a if x == 0.0 else b
        if e < 1.0:
            raise ZeroDivisionError
        if e > 1.0:
            return 0.0
        other = b if x == 0.0 else a
        return other  # a == 1 (or b == 1): density is b (or a) at that end
    return math.exp((a - 1.0) * math.log(x) + (b - 1.0) * math.log1p(-x) - lbeta(a, b))


def beta_ppf(p: float, a: float, b: float) -> float:
    if p <= 0.0:
        return 0.0
    if p >= 1.0:
        return 1.0
    return invert(lambda x: betainc(a, b, x), p, 0.0, 1.0)


def binom_pmf(k: int, n: int, p: float) -> float:
    if k < 0 or k > n:
        return 0.0
    if p == 0.0 or p == 1.0:
        return 1.0 if k == (0 if p == 0.0 else n) else 0.0
    try:
        return float(math.comb(n, k)) * p**k * (1.0 - p) ** (n - k)
    except OverflowError:
        lg = math.lgamma(n + 1) - math.lgamma(k + 1) - math.lgamma(n - k + 1)
        return math.exp(lg + k * math.log(p) + (n - k) * math.log1p(-p))


def binom_cdf(k: int, n: int, p: float) -> float:
    if k < 0:
        return 0.0
    if k >= n:
        return 1.0
    if n <= 1000:
        return min(1.0, math.fsum(binom_pmf(i, n, p) for i in range(k + 1)))
    if p == 0.0:
        return 1.0
    if p == 1.0:
        return 0.0
    return betainc(n - k, k + 1, 1.0 - p, p)


def poisson_pmf(k: int, lam: float) -> float:
    if lam == 0.0:
        return 1.0 if k == 0 else 0.0
    return math.exp(k * math.log(lam) - lam - math.lgamma(k + 1))


def poisson_cdf(k: int, lam: float) -> float:
    if lam == 0.0:
        return 1.0
    if k < 200:
        return min(1.0, math.fsum(poisson_pmf(i, lam) for i in range(k + 1)))
    return gammaincc(k + 1, lam)
