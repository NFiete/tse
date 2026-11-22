"""Least-squares regression returning arrays: LINEST, LOGEST, TREND, GROWTH."""
from __future__ import annotations

import math

from ..errors import NA, NUM, REF, VALUE, XLError
from ..registry import fn, scalar_arg, to_array
from ..values import MISSING, Array, is_number, to_bool


def _numeric(arr: Array) -> list[list[float]]:
    rows = []
    for row in arr.rows:
        out = []
        for v in row:
            if not is_number(v):
                raise XLError(VALUE)
            out.append(float(v))
        rows.append(out)
    return rows


def _flag(ctx, v, default: bool) -> bool:
    if v is MISSING:
        return default
    v = scalar_arg(ctx, v)
    return default if v is None else to_bool(v)


class _Data:
    """Observations ys[i] with regressors xs[i][0..k-1], and how the variables are laid out."""

    def __init__(self, ys, xs, by_row: bool, single: bool):
        self.ys, self.xs = ys, xs
        self.by_row = by_row  # variables are rows of known_x (known_y is a row)
        self.single = single  # one variable whose values have known_y's shape
        self.k = len(xs[0]) if xs else 0


def _data(ctx, known_y, known_x) -> _Data:
    y = _numeric(to_array(ctx, known_y))
    ny_r, ny_c = len(y), len(y[0])
    ys = [v for row in y for v in row]
    if known_x is MISSING:
        return _Data(ys, [[float(i)] for i in range(1, len(ys) + 1)], ny_r == 1 and ny_c > 1, True)
    x = _numeric(to_array(ctx, known_x))
    nx_r, nx_c = len(x), len(x[0])
    if (nx_r, nx_c) == (ny_r, ny_c):
        return _Data(ys, [[v] for row in x for v in row], ny_r == 1 and ny_c > 1, True)
    if ny_c == 1 and nx_r == ny_r:
        return _Data(ys, [list(row) for row in x], False, False)
    if ny_r == 1 and nx_c == ny_c:
        return _Data(ys, [[x[r][i] for r in range(nx_r)] for i in range(nx_c)], True, False)
    raise XLError(REF)


class _Fit:
    """Ordinary least squares via modified Gram-Schmidt, dropping collinear regressors."""

    def __init__(self, d: _Data, const: bool):
        n, k = len(d.ys), d.k
        if const and n < 2:
            raise XLError(VALUE)
        self.n, self.k, self.const = n, k, const
        xm = [math.fsum(row[j] for row in d.xs) / n for j in range(k)] if const else [0.0] * k
        ym = math.fsum(d.ys) / n if const else 0.0
        cols = [[row[j] - xm[j] for row in d.xs] for j in range(k)]
        yc = [v - ym for v in d.ys]
        q: list[list[float]] = []
        r: list[list[float]] = []  # r[i] = coefficients of kept column i on q[0..i]
        kept: list[int] = []
        for j, col in enumerate(cols):
            v = list(col)
            norm0 = math.sqrt(math.fsum(t * t for t in v))
            coeffs = [0.0] * len(q)
            for _ in range(2):  # re-orthogonalize once for stability
                for i, qi in enumerate(q):
                    c = math.fsum(a * b for a, b in zip(qi, v))
                    coeffs[i] += c
                    v = [a - c * b for a, b in zip(v, qi)]
            nv = math.sqrt(math.fsum(t * t for t in v))
            if norm0 == 0 or nv <= 1e-10 * norm0:
                continue
            q.append([t / nv for t in v])
            r.append(coeffs + [nv])
            kept.append(j)
        p = len(kept)
        # R is upper triangular with column i = r[i]; solve R b = Q'y and invert R.
        qty = [math.fsum(a * b for a, b in zip(qi, yc)) for qi in q]
        rinv = [[0.0] * p for _ in range(p)]
        for i in range(p - 1, -1, -1):
            rinv[i][i] = 1.0 / r[i][i]
            for j in range(i + 1, p):
                rinv[i][j] = -math.fsum(r[m][i] * rinv[m][j] for m in range(i + 1, j + 1)) / r[i][i]
        bk = [math.fsum(rinv[i][j] * qty[j] for j in range(i, p)) for i in range(p)]
        self.coef = [0.0] * k
        for i, j in enumerate(kept):
            self.coef[j] = bk[i]
        self.intercept = ym - math.fsum(self.coef[j] * xm[j] for j in range(k)) if const else 0.0
        resid = [yc[i] - math.fsum(self.coef[j] * cols[j][i] for j in range(k)) for i in range(n)]
        self.sstotal = math.fsum(v * v for v in yc)
        self.ssresid = math.fsum(e * e for e in resid)
        if self.ssresid <= 1e-24 * self.sstotal:  # exact fit up to rounding
            self.ssresid = 0.0
        self.ssreg = self.sstotal - self.ssresid
        self.p = p
        self.df = n - p - (1 if const else 0)
        # diag of (X'X)^-1 for kept columns = row sums of squares of R^-1
        self._cov = {kept[i]: math.fsum(rinv[i][j] ** 2 for j in range(i, p)) for i in range(p)}
        self._rinv, self._kept, self._xm = rinv, kept, xm

    def predict(self, row: list[float]) -> float:
        return self.intercept + math.fsum(c * x for c, x in zip(self.coef, row))

    def stats(self):
        """[[se_m_k..se_m_1, se_b], [r2, sey], [F, df], [ssreg, ssresid]] with errors as values."""
        k, df = self.k, self.df
        if df < 0:
            raise XLError(NUM)
        s2 = self.ssresid / df if df > 0 else 0.0
        se = [math.sqrt(s2 * self._cov[j]) if j in self._cov else 0.0 for j in range(k)]
        if self.const:
            # Var(b) = s2 * (1/n + xm' (Xc'Xc)^-1 xm); (Xc'Xc)^-1 = Rinv Rinv'
            w = [math.fsum(self._rinv[i][j] * self._xm[self._kept[i]] for i in range(j + 1))
                 for j in range(self.p)]
            se_b = math.sqrt(s2 * (1.0 / self.n + math.fsum(t * t for t in w)))
        else:
            se_b = NA
        r2 = self.ssreg / self.sstotal if self.sstotal != 0 else 1.0
        sey = math.sqrt(s2)
        if df == 0 or self.p == 0 or self.ssresid == 0:
            f = NA
        else:
            f = (self.ssreg / self.p) / s2
        return se[::-1] + [se_b], [r2, sey], [f, float(df)], [self.ssreg, self.ssresid]


def _regression_table(fit: _Fit, top: list, stats: bool) -> Array:
    width = fit.k + 1
    if not stats:
        return Array([top])
    se, (r2, sey), (f, df), (ssreg, ssresid) = fit.stats()
    pad = [NA] * (width - 2)
    return Array([top, se, [r2, sey] + pad, [f, df] + pad, [ssreg, ssresid] + pad])


def _clean(v):
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return NUM
    return v


def _finish(arr: Array) -> Array:
    return Array([[_clean(v) for v in row] for row in arr.rows])


@fn("LINEST", kind="raw", min_args=1, max_args=4)
def LINEST(ctx, known_y, known_x=MISSING, const=MISSING, stats=MISSING):
    d = _data(ctx, known_y, known_x)
    fit = _Fit(d, _flag(ctx, const, True))
    top = fit.coef[::-1] + [fit.intercept]
    return _finish(_regression_table(fit, top, _flag(ctx, stats, False)))


def _log_data(d: _Data) -> _Data:
    if any(v <= 0 for v in d.ys):
        raise XLError(NUM)
    d.ys = [math.log(v) for v in d.ys]
    return d


@fn("LOGEST", kind="raw", min_args=1, max_args=4)
def LOGEST(ctx, known_y, known_x=MISSING, const=MISSING, stats=MISSING):
    d = _log_data(_data(ctx, known_y, known_x))
    fit = _Fit(d, _flag(ctx, const, True))
    top = [math.exp(c) for c in fit.coef[::-1]] + [math.exp(fit.intercept)]
    return _finish(_regression_table(fit, top, _flag(ctx, stats, False)))


def _predict(ctx, d: _Data, fit: _Fit, known_x, new_x, transform) -> Array:
    if new_x is MISSING:
        if known_x is MISSING:
            n = len(d.ys)
            new = [[float(i) for i in range(1, n + 1)]] if d.by_row else [[float(i)] for i in range(1, n + 1)]
        else:
            new = _numeric(to_array(ctx, known_x))
    else:
        new = _numeric(to_array(ctx, new_x))
    if d.k == 1:
        return Array([[transform(fit.predict([v])) for v in row] for row in new])
    if d.by_row:
        if len(new) != d.k:
            raise XLError(REF)
        return Array([[transform(fit.predict([new[j][i] for j in range(d.k)])) for i in range(len(new[0]))]])
    if len(new[0]) != d.k:
        raise XLError(REF)
    return Array([[transform(fit.predict(row))] for row in new])


@fn("TREND", kind="raw", min_args=1, max_args=4)
def TREND(ctx, known_y, known_x=MISSING, new_x=MISSING, const=MISSING):
    d = _data(ctx, known_y, known_x)
    fit = _Fit(d, _flag(ctx, const, True))
    return _finish(_predict(ctx, d, fit, known_x, new_x, lambda v: v))


@fn("GROWTH", kind="raw", min_args=1, max_args=4)
def GROWTH(ctx, known_y, known_x=MISSING, new_x=MISSING, const=MISSING):
    d = _log_data(_data(ctx, known_y, known_x))
    fit = _Fit(d, _flag(ctx, const, True))
    return _finish(_predict(ctx, d, fit, known_x, new_x, math.exp))
