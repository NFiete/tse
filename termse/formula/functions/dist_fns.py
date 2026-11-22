"""Probability distributions, their inverses and hypothesis tests (NORM.DIST, T.TEST, ...).

Each modern function is registered together with its pre-2010 compatibility name where the
two have the same meaning; legacy names with different signatures are separate functions.
"""
from __future__ import annotations

import math

from .. import special as S
from ..errors import DIV0, NA, NUM, XLError
from ..registry import fn, numbers, scalar_arg, to_array
from ..values import MISSING, is_number, to_bool, to_int, to_number


def _num(v) -> float:
    return to_number(v)


def _bad(cond: bool):
    if cond:
        raise XLError(NUM)


def _prob(p, *, zero_ok=False, one_ok=False) -> float:
    """A probability argument; out-of-range values are #NUM!."""
    p = _num(p)
    _bad(p < 0 or p > 1 or (p == 0 and not zero_ok) or (p == 1 and not one_ok))
    return p


def _dof(v, minimum: int = 1) -> int:
    """Degrees of freedom: truncated to an integer, at least `minimum`, at most 10^10."""
    n = to_int(v)
    _bad(n < minimum or n > 10**10)
    return n


# ------------------------------------------------------------------ normal


def _norm(x, mean, sd, cum):
    x, mean, sd = _num(x), _num(mean), _num(sd)
    _bad(sd <= 0)
    z = (x - mean) / sd
    return S.norm_cdf(z) if to_bool(cum) else S.norm_pdf(z) / sd


@fn("NORM.DIST", min_args=4, max_args=4)
def NORM_DIST(x, mean, sd, cumulative):
    return _norm(x, mean, sd, cumulative)


@fn("NORMDIST", min_args=4, max_args=4)
def NORMDIST(x, mean, sd, cumulative):
    return _norm(x, mean, sd, cumulative)


@fn("NORM.INV", "NORMINV", min_args=3, max_args=3)
def NORM_INV(p, mean, sd):
    p, mean, sd = _prob(p), _num(mean), _num(sd)
    _bad(sd <= 0)
    return mean + sd * S.norm_ppf(p)


@fn("NORM.S.DIST", min_args=2, max_args=2)
def NORM_S_DIST(z, cumulative):
    z = _num(z)
    return S.norm_cdf(z) if to_bool(cumulative) else S.norm_pdf(z)


@fn("NORMSDIST", min_args=1, max_args=1)
def NORMSDIST(z):
    return S.norm_cdf(_num(z))


@fn("NORM.S.INV", "NORMSINV", min_args=1, max_args=1)
def NORM_S_INV(p):
    return S.norm_ppf(_prob(p))


@fn("PHI", min_args=1, max_args=1)
def PHI(x):
    return S.norm_pdf(_num(x))


@fn("GAUSS", min_args=1, max_args=1)
def GAUSS(z):
    z = _num(z)
    # P(0 <= Z <= z), keeping precision for large |z|
    return 0.5 - S.norm_sf(z) if z >= 0 else S.norm_cdf(z) - 0.5


@fn("STANDARDIZE", min_args=3, max_args=3)
def STANDARDIZE(x, mean, sd):
    x, mean, sd = _num(x), _num(mean), _num(sd)
    _bad(sd <= 0)
    return (x - mean) / sd


# ------------------------------------------------------------------ lognormal


@fn("LOGNORM.DIST", min_args=4, max_args=4)
def LOGNORM_DIST(x, mean, sd, cumulative):
    x, mean, sd = _num(x), _num(mean), _num(sd)
    _bad(x <= 0 or sd <= 0)
    z = (math.log(x) - mean) / sd
    return S.norm_cdf(z) if to_bool(cumulative) else S.norm_pdf(z) / (x * sd)


@fn("LOGNORMDIST", min_args=3, max_args=3)
def LOGNORMDIST(x, mean, sd):
    return LOGNORM_DIST(x, mean, sd, True)


@fn("LOGNORM.INV", "LOGINV", min_args=3, max_args=3)
def LOGNORM_INV(p, mean, sd):
    p, mean, sd = _prob(p), _num(mean), _num(sd)
    _bad(sd <= 0)
    return math.exp(mean + sd * S.norm_ppf(p))


# ------------------------------------------------------------------ exponential, Weibull


@fn("EXPON.DIST", "EXPONDIST", min_args=3, max_args=3)
def EXPON_DIST(x, lam, cumulative):
    x, lam = _num(x), _num(lam)
    _bad(x < 0 or lam <= 0)
    return -math.expm1(-lam * x) if to_bool(cumulative) else lam * math.exp(-lam * x)


@fn("WEIBULL.DIST", "WEIBULL", min_args=4, max_args=4)
def WEIBULL_DIST(x, alpha, beta, cumulative):
    x, a, b = _num(x), _num(alpha), _num(beta)
    _bad(x < 0 or a <= 0 or b <= 0)
    if to_bool(cumulative):
        return -math.expm1(-((x / b) ** a))
    if x == 0:
        if a < 1:
            raise XLError(NUM)
        return 1 / b if a == 1 else 0.0
    return a / b**a * x ** (a - 1) * math.exp(-((x / b) ** a))


# ------------------------------------------------------------------ gamma


@fn("GAMMA", min_args=1, max_args=1)
def GAMMA(x):
    x = _num(x)
    _bad(x <= 0 and x == math.floor(x))
    return math.gamma(x)


@fn("GAMMALN", "GAMMALN.PRECISE", min_args=1, max_args=1)
def GAMMALN(x):
    x = _num(x)
    _bad(x <= 0)
    return math.lgamma(x)


@fn("GAMMA.DIST", "GAMMADIST", min_args=4, max_args=4)
def GAMMA_DIST(x, alpha, beta, cumulative):
    x, a, b = _num(x), _num(alpha), _num(beta)
    _bad(x < 0 or a <= 0 or b <= 0)
    if to_bool(cumulative):
        return S.gammainc(a, x / b)
    return S.gamma_pdf(x, a, b)


@fn("GAMMA.INV", "GAMMAINV", min_args=3, max_args=3)
def GAMMA_INV(p, alpha, beta):
    p, a, b = _prob(p, zero_ok=True), _num(alpha), _num(beta)
    _bad(a <= 0 or b <= 0)
    if p == 0:
        return 0.0
    return S.gamma_ppf(p, a, b)


# ------------------------------------------------------------------ beta


def _beta_bounds(lo, hi):
    lo = 0.0 if lo is None else _num(lo)
    hi = 1.0 if hi is None else _num(hi)
    _bad(lo == hi)
    return lo, hi


@fn("BETA.DIST", min_args=4, max_args=6)
def BETA_DIST(x, alpha, beta, cumulative, a=None, b=None):
    x, al, be = _num(x), _num(alpha), _num(beta)
    lo, hi = _beta_bounds(a, b)
    _bad(al <= 0 or be <= 0 or x < lo or x > hi)
    u = (x - lo) / (hi - lo)
    if to_bool(cumulative):
        return S.betainc(al, be, u)
    return S.beta_pdf(u, al, be) / (hi - lo)


@fn("BETADIST", min_args=3, max_args=5)
def BETADIST(x, alpha, beta, a=None, b=None):
    x, al, be = _num(x), _num(alpha), _num(beta)
    lo, hi = _beta_bounds(a, b)
    _bad(al <= 0 or be <= 0 or x < lo or x > hi)
    return S.betainc(al, be, (x - lo) / (hi - lo))


@fn("BETA.INV", "BETAINV", min_args=3, max_args=5)
def BETA_INV(p, alpha, beta, a=None, b=None):
    p, al, be = _prob(p, one_ok=True), _num(alpha), _num(beta)
    lo, hi = _beta_bounds(a, b)
    _bad(al <= 0 or be <= 0)
    return lo + (hi - lo) * S.beta_ppf(p, al, be)


# ------------------------------------------------------------------ chi-squared


@fn("CHISQ.DIST", min_args=3, max_args=3)
def CHISQ_DIST(x, df, cumulative):
    x, k = _num(x), _dof(df)
    _bad(x < 0)
    if to_bool(cumulative):
        return S.chi2_cdf(x, k)
    return S.gamma_pdf(x, k / 2, 2.0)


@fn("CHISQ.DIST.RT", "CHIDIST", min_args=2, max_args=2)
def CHISQ_DIST_RT(x, df):
    x, k = _num(x), _dof(df)
    _bad(x < 0)
    return S.chi2_sf(x, k)


@fn("CHISQ.INV", min_args=2, max_args=2)
def CHISQ_INV(p, df):
    p, k = _prob(p, zero_ok=True), _dof(df)
    if p == 0:
        return 0.0
    return S.gamma_ppf(p, k / 2, 2.0)


@fn("CHISQ.INV.RT", "CHIINV", min_args=2, max_args=2)
def CHISQ_INV_RT(p, df):
    p, k = _prob(p, one_ok=True), _dof(df)
    if p == 1:
        return 0.0
    return S.gamma_ppf(p, k / 2, 2.0, upper=True)


@fn("CHISQ.TEST", "CHITEST", kind="raw", min_args=2, max_args=2)
def CHISQ_TEST(ctx, actual, expected):
    a, e = to_array(ctx, actual), to_array(ctx, expected)
    if (a.nrows, a.ncols) != (e.nrows, e.ncols):
        raise XLError(NA)
    r, c = a.nrows, a.ncols
    if r == 1 and c == 1:
        raise XLError(NA)
    df = (r - 1) * (c - 1) if r > 1 and c > 1 else max(r, c) - 1
    stat = []
    for x, m in zip(a.flat(), e.flat()):
        if is_number(x) and is_number(m):
            if m == 0:
                raise XLError(DIV0)
            stat.append((x - m) ** 2 / m)
    return S.chi2_sf(math.fsum(stat), df)


# ------------------------------------------------------------------ Student t


@fn("T.DIST", min_args=3, max_args=3)
def T_DIST(x, df, cumulative):
    x, k = _num(x), _dof(df)
    return S.t_cdf(x, k) if to_bool(cumulative) else S.t_pdf(x, k)


@fn("T.DIST.RT", min_args=2, max_args=2)
def T_DIST_RT(x, df):
    x, k = _num(x), _dof(df)
    return S.t_cdf(-x, k)


@fn("T.DIST.2T", min_args=2, max_args=2)
def T_DIST_2T(x, df):
    x, k = _num(x), _dof(df)
    _bad(x < 0)
    return S.t_sf2(x, k)


@fn("TDIST", min_args=3, max_args=3)
def TDIST(x, df, tails):
    x, k, tails = _num(x), _dof(df), to_int(tails)
    _bad(x < 0 or tails not in (1, 2))
    return S.t_sf2(x, k) / (3 - tails)


@fn("T.INV", min_args=2, max_args=2)
def T_INV(p, df):
    p, k = _prob(p), _dof(df)
    return S.t_ppf(p, k)


@fn("T.INV.2T", "TINV", min_args=2, max_args=2)
def T_INV_2T(p, df):
    p, k = _prob(p, one_ok=True), _dof(df)
    return S.t_ppf2(p, k)


def _mean_var(xs):
    n = len(xs)
    m = math.fsum(xs) / n
    return m, math.fsum((x - m) ** 2 for x in xs) / (n - 1)


@fn("T.TEST", "TTEST", kind="raw", min_args=4, max_args=4)
def T_TEST(ctx, array1, array2, tails, kind):
    tails = to_int(scalar_arg(ctx, tails))
    kind = to_int(scalar_arg(ctx, kind))
    _bad(tails not in (1, 2) or kind not in (1, 2, 3))
    if kind == 1:
        a, b = list(to_array(ctx, array1).flat()), list(to_array(ctx, array2).flat())
        if len(a) != len(b):
            raise XLError(NA)
        d = [x - y for x, y in zip(a, b) if is_number(x) and is_number(y)]
        if len(d) < 2:
            raise XLError(DIV0)
        m, v = _mean_var(d)
        if v == 0:
            raise XLError(DIV0)
        t, df = m / math.sqrt(v / len(d)), len(d) - 1.0
    else:
        a, b = numbers(ctx, [array1]), numbers(ctx, [array2])
        n1, n2 = len(a), len(b)
        if n1 < 2 or n2 < 2:
            raise XLError(DIV0)
        m1, v1 = _mean_var(a)
        m2, v2 = _mean_var(b)
        if kind == 2:
            df = n1 + n2 - 2.0
            pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / df
            se = math.sqrt(pooled * (1 / n1 + 1 / n2))
        else:
            s1, s2 = v1 / n1, v2 / n2
            se = math.sqrt(s1 + s2)
            if se == 0:
                raise XLError(DIV0)
            df = (s1 + s2) ** 2 / (s1**2 / (n1 - 1) + s2**2 / (n2 - 1))
        if se == 0:
            raise XLError(DIV0)
        t = (m1 - m2) / se
    return S.t_sf2(abs(t), df) / (3 - tails)


# ------------------------------------------------------------------ F


@fn("F.DIST", min_args=4, max_args=4)
def F_DIST(x, df1, df2, cumulative):
    x, d1, d2 = _num(x), _dof(df1), _dof(df2)
    _bad(x < 0)
    return S.f_cdf(x, d1, d2) if to_bool(cumulative) else S.f_pdf(x, d1, d2)


@fn("F.DIST.RT", "FDIST", min_args=3, max_args=3)
def F_DIST_RT(x, df1, df2):
    x, d1, d2 = _num(x), _dof(df1), _dof(df2)
    _bad(x < 0)
    return S.f_sf(x, d1, d2)


@fn("F.INV", min_args=3, max_args=3)
def F_INV(p, df1, df2):
    p, d1, d2 = _prob(p, zero_ok=True), _dof(df1), _dof(df2)
    if p == 0:
        return 0.0
    return _f_ppf(p, d1, d2, upper=False)


@fn("F.INV.RT", "FINV", min_args=3, max_args=3)
def F_INV_RT(p, df1, df2):
    p, d1, d2 = _prob(p, one_ok=True), _dof(df1), _dof(df2)
    if p == 1:
        return 0.0
    return _f_ppf(p, d1, d2, upper=True)


def _f_ppf(p, d1, d2, upper):
    # Invert through the beta distribution: X = d2*u / (d1*(1-u)) with u ~ Beta(d1/2, d2/2).
    a, b = d1 / 2.0, d2 / 2.0
    if upper:  # 1 - I_u(a, b) = p  <=>  I_{1-u}(b, a) = p
        w = S.beta_ppf(p, b, a)  # w = 1 - u
        return d2 * (1.0 - w) / (d1 * w)
    u = S.beta_ppf(p, a, b)
    return d2 * u / (d1 * (1.0 - u))


@fn("F.TEST", "FTEST", kind="raw", min_args=2, max_args=2)
def F_TEST(ctx, array1, array2):
    a, b = numbers(ctx, [array1]), numbers(ctx, [array2])
    if len(a) < 2 or len(b) < 2:
        raise XLError(DIV0)
    v1, v2 = _mean_var(a)[1], _mean_var(b)[1]
    if v1 == 0 or v2 == 0:
        raise XLError(DIV0)
    f = v1 / v2
    d1, d2 = len(a) - 1, len(b) - 1
    return min(1.0, 2.0 * min(S.f_cdf(f, d1, d2), S.f_sf(f, d1, d2)))


# ------------------------------------------------------------------ z


@fn("Z.TEST", "ZTEST", kind="raw", min_args=2, max_args=3)
def Z_TEST(ctx, array, x, sigma=MISSING):
    vals = numbers(ctx, [array])
    if not vals:
        raise XLError(NA)
    x = to_number(scalar_arg(ctx, x))
    n = len(vals)
    mean = math.fsum(vals) / n
    if sigma is MISSING:
        if n < 2:
            raise XLError(DIV0)
        sd = math.sqrt(_mean_var(vals)[1])
    else:
        sd = to_number(scalar_arg(ctx, sigma))
        _bad(sd <= 0)
    if sd == 0:
        raise XLError(DIV0)
    return S.norm_sf((mean - x) / (sd / math.sqrt(n)))


# ------------------------------------------------------------------ confidence intervals


@fn("CONFIDENCE.NORM", "CONFIDENCE", min_args=3, max_args=3)
def CONFIDENCE_NORM(alpha, sd, size):
    alpha, sd, n = _num(alpha), _num(sd), to_int(size)
    _bad(alpha <= 0 or alpha >= 1 or sd <= 0 or n < 1)
    return -S.norm_ppf(alpha / 2) * sd / math.sqrt(n)


@fn("CONFIDENCE.T", min_args=3, max_args=3)
def CONFIDENCE_T(alpha, sd, size):
    alpha, sd, n = _num(alpha), _num(sd), to_int(size)
    _bad(alpha <= 0 or alpha >= 1 or sd <= 0 or n < 1)
    if n == 1:
        raise XLError(DIV0)
    return S.t_ppf2(alpha, n - 1) * sd / math.sqrt(n)


# ------------------------------------------------------------------ discrete distributions


@fn("BINOM.DIST.RANGE", min_args=3, max_args=4)
def BINOM_DIST_RANGE(trials, probability_s, number_s, number_s2=None):
    n, p, k1 = to_int(trials), _num(probability_s), to_int(number_s)
    k2 = k1 if number_s2 is None else to_int(number_s2)
    _bad(n < 0 or not 0 <= p <= 1 or k1 < 0 or k1 > n or k2 < k1 or k2 > n)
    if k1 == 0:
        return S.binom_cdf(k2, n, p)
    if k2 - k1 <= 1000:
        return min(1.0, math.fsum(S.binom_pmf(i, n, p) for i in range(k1, k2 + 1)))
    return S.binom_cdf(k2, n, p) - S.binom_cdf(k1 - 1, n, p)


@fn("BINOM.INV", "CRITBINOM", min_args=3, max_args=3)
def BINOM_INV(trials, probability_s, alpha):
    n, p, alpha = to_int(trials), _num(probability_s), _num(alpha)
    _bad(n < 0 or not 0 <= p <= 1 or not 0 <= alpha <= 1)
    if n <= 1000:
        total = 0.0
        for k in range(n + 1):
            total += S.binom_pmf(k, n, p)
            if total >= alpha * (1 - 1e-15):
                return k
        return n
    lo, hi = 0, n  # smallest k with cdf(k) >= alpha
    while lo < hi:
        mid = (lo + hi) // 2
        if S.binom_cdf(mid, n, p) >= alpha:
            hi = mid
        else:
            lo = mid + 1
    return lo


@fn("NEGBINOM.DIST", min_args=4, max_args=4)
def NEGBINOM_DIST(number_f, number_s, probability_s, cumulative):
    f, s, p = to_int(number_f), to_int(number_s), _num(probability_s)
    _bad(f < 0 or s < 1 or p < 0 or p > 1)
    if to_bool(cumulative):
        return S.betainc(s, f + 1, p)
    return _negbinom_pmf(f, s, p)


@fn("NEGBINOMDIST", min_args=3, max_args=3)
def NEGBINOMDIST(number_f, number_s, probability_s):
    return NEGBINOM_DIST(number_f, number_s, probability_s, False)


def _negbinom_pmf(f, s, p):
    if p == 0:
        return 0.0
    if p == 1:
        return 1.0 if f == 0 else 0.0
    try:
        return float(math.comb(f + s - 1, s - 1)) * p**s * (1 - p) ** f
    except OverflowError:
        lg = math.lgamma(f + s) - math.lgamma(s) - math.lgamma(f + 1)
        return math.exp(lg + s * math.log(p) + f * math.log1p(-p))


@fn("POISSON.DIST", "POISSON", min_args=3, max_args=3)
def POISSON_DIST(x, mean, cumulative):
    k, lam = to_int(x), _num(mean)
    _bad(k < 0 or lam < 0)
    return S.poisson_cdf(k, lam) if to_bool(cumulative) else S.poisson_pmf(k, lam)


def _hypgeom(k, n, K, N, cum):
    k, n, K, N = to_int(k), to_int(n), to_int(K), to_int(N)
    _bad(N <= 0 or n <= 0 or n > N or K <= 0 or K > N)
    _bad(k < 0 or k > min(n, K) or k < max(0, n - N + K))
    lo = max(0, n - N + K)
    ks = range(lo, k + 1) if cum else (k,)
    if N <= 100000:
        return sum(math.comb(K, i) * math.comb(N - K, n - i) for i in ks) / math.comb(N, n)

    def lcomb(a, b):
        return math.lgamma(a + 1) - math.lgamma(b + 1) - math.lgamma(a - b + 1)

    total = lcomb(N, n)
    return min(1.0, math.fsum(math.exp(lcomb(K, i) + lcomb(N - K, n - i) - total) for i in ks))


@fn("HYPGEOM.DIST", min_args=5, max_args=5)
def HYPGEOM_DIST(sample_s, number_sample, population_s, number_pop, cumulative):
    return _hypgeom(sample_s, number_sample, population_s, number_pop, to_bool(cumulative))


@fn("HYPGEOMDIST", min_args=4, max_args=4)
def HYPGEOMDIST(sample_s, number_sample, population_s, number_pop):
    return _hypgeom(sample_s, number_sample, population_s, number_pop, False)
