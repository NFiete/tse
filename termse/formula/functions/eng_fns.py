"""Engineering functions: base conversions, bit operations, CONVERT, ERF, Bessel functions and
complex numbers (the IM* functions)."""
from __future__ import annotations

import cmath
import math
import re

from ..errors import DIV0, NA, NUM, VALUE, XLError
from ..registry import fn, iter_ref
from ..values import MISSING, Array, RefValue, format_general_number, to_number, to_text


def _num(v):
    """A numeric argument. Unlike ordinary arithmetic, TRUE/FALSE are #VALUE!."""
    if type(v) is bool:
        raise XLError(VALUE)
    return to_number(v)


def _trunc(v) -> int:
    return math.trunc(_num(v))


# ------------------------------------------------------------------ base conversions

# base -> (number of bits held in 10 digits, digit regex)
_BASES = {2: (10, re.compile(r"[01]{0,10}")), 8: (30, re.compile(r"[0-7]{0,10}")),
          16: (40, re.compile(r"[0-9A-Fa-f]{0,10}"))}
_DIGITS = {2: "b", 8: "o", 16: "X"}


def _parse_base(v, base: int) -> int:
    """Signed value of a 10-digit two's-complement number written in `base`."""
    if type(v) is bool:
        raise XLError(VALUE)
    s = "" if v is None else to_text(v)
    bits, rx = _BASES[base]
    if not rx.fullmatch(s):
        raise XLError(NUM)
    if not s:
        return 0
    n = int(s, base)
    return n - (1 << bits) if n >> (bits - 1) else n


def _format_base(n: int, base: int, places) -> str:
    bits = _BASES[base][0]
    if places is not None:
        places = _trunc(places)
        if not 0 <= places <= 10:
            raise XLError(NUM)
    if not -(1 << (bits - 1)) <= n < (1 << (bits - 1)):
        raise XLError(NUM)
    if n < 0:  # ten digits of two's complement; places is ignored
        return format(n + (1 << bits), _DIGITS[base])
    s = format(n, _DIGITS[base])
    if places is not None:
        if len(s) > places:
            raise XLError(NUM)
        s = s.zfill(places)
    return s


def _register_base(src: str, sbase: int | None, dst: str, dbase: int | None):
    if dbase is None:
        def impl(number):
            return _parse_base(number, sbase)
    elif sbase is None:
        def impl(number, places=None):
            return _format_base(_trunc(number), dbase, places)
    else:
        def impl(number, places=None):
            return _format_base(_parse_base(number, sbase), dbase, places)
    fn(f"{src}2{dst}")(impl)


_BASE_NAMES = {"BIN": 2, "OCT": 8, "DEC": None, "HEX": 16}
for _s, _sb in _BASE_NAMES.items():
    for _d, _db in _BASE_NAMES.items():
        if _s != _d:
            _register_base(_s, _sb, _d, _db)


# ------------------------------------------------------------------ bit operations

_BIT_LIMIT = 1 << 48


def _bit_arg(v) -> int:
    x = _num(v)
    if x < 0 or x != math.floor(x) or x >= _BIT_LIMIT:
        raise XLError(NUM)
    return int(x)


@fn("BITAND")
def BITAND(a, b):
    return _bit_arg(a) & _bit_arg(b)


@fn("BITOR")
def BITOR(a, b):
    return _bit_arg(a) | _bit_arg(b)


@fn("BITXOR")
def BITXOR(a, b):
    return _bit_arg(a) ^ _bit_arg(b)


def _shift(number, amount, left: bool):
    n = _bit_arg(number)
    k = _trunc(amount)
    if abs(k) > 53:
        raise XLError(NUM)
    if not left:
        k = -k
    return n << k if k >= 0 else n >> -k


@fn("BITLSHIFT")
def BITLSHIFT(number, shift_amount):
    return _shift(number, shift_amount, True)


@fn("BITRSHIFT")
def BITRSHIFT(number, shift_amount):
    return _shift(number, shift_amount, False)


# ------------------------------------------------------------------ DELTA, GESTEP, ERF


@fn("DELTA")
def DELTA(a, b=None):
    return 1 if _num(a) == (0 if b is None else _num(b)) else 0


@fn("GESTEP")
def GESTEP(number, step=None):
    return 1 if _num(number) >= (0 if step is None else _num(step)) else 0


@fn("ERF")
def ERF(lower, upper=MISSING):
    lo = _num(lower)
    if upper is MISSING:  # omitted (a blank cell counts as 0)
        return math.erf(lo)
    return math.erf(_num(upper)) - math.erf(lo)


@fn("ERF.PRECISE")
def ERF_PRECISE(x):
    return math.erf(_num(x))


@fn("ERFC", "ERFC.PRECISE")
def ERFC(x):
    return math.erfc(_num(x))


# ------------------------------------------------------------------ CONVERT

# unit -> (category, factor to the category's base unit, accepts metric prefixes)
_UNITS: dict[str, tuple[str, float, bool]] = {}


def _units(cat: str, factor: float, prefixed: bool, *names: str):
    for n in names:
        _UNITS[n] = (cat, factor, prefixed)


_IN = 0.0254
_FT = 0.3048
_MI = 1609.344
_NMI = 1852.0
_LY = 9.4607304725808e15
_PICAPT = _IN / 72
_LBF = 4.4482216152605
_HP = 550 * _FT * _LBF
_GAL = 231 * _IN ** 3
_UK_GAL = 4.54609e-3

# Weight and mass (grams)
_units("mass", 1.0, True, "g")
_units("mass", _LBF / _FT * 1000, False, "sg")
_units("mass", 453.59237, False, "lbm")
_units("mass", 1.660538782e-24, True, "u")
_units("mass", 453.59237 / 16, False, "ozm")
_units("mass", 0.06479891, False, "grain")
_units("mass", 45359.237, False, "cwt", "shweight")
_units("mass", 50802.34544, False, "uk_cwt", "lcwt", "hweight")
_units("mass", 6350.29318, False, "stone")
_units("mass", 907184.74, False, "ton")
_units("mass", 1016046.9088, False, "uk_ton", "LTON", "brton")
# Distance (metres)
_units("dist", 1.0, True, "m")
_units("dist", _MI, False, "mi")
_units("dist", _NMI, False, "Nmi")
_units("dist", _IN, False, "in")
_units("dist", _FT, False, "ft")
_units("dist", 0.9144, False, "yd")
_units("dist", 1e-10, True, "ang")
_units("dist", 1.143, False, "ell")
_units("dist", _LY, True, "ly")
_units("dist", 3.08567758128155e16, True, "parsec", "pc")
_units("dist", _PICAPT, False, "Picapt", "Pica")
_units("dist", _IN / 6, False, "pica")
_units("dist", 5280 * 1200 / 3937, False, "survey_mi")
# Time (seconds)
_units("time", 365.25 * 86400, False, "yr")
_units("time", 86400.0, False, "day", "d")
_units("time", 3600.0, False, "hr")
_units("time", 60.0, False, "mn", "min")
_units("time", 1.0, True, "sec", "s")
# Pressure (pascals)
_units("pressure", 1.0, True, "Pa", "p")
_units("pressure", 101325.0, True, "atm", "at")
_units("pressure", 101325 / 760, True, "mmHg")
_units("pressure", _LBF / _IN ** 2, False, "psi")
_units("pressure", 101325 / 760, False, "Torr")
# Force (newtons)
_units("force", 1.0, True, "N")
_units("force", 1e-5, True, "dyn", "dy")
_units("force", _LBF, False, "lbf")
_units("force", 9.80665e-3, True, "pond")
# Energy (joules)
_units("energy", 1.0, True, "J")
_units("energy", 1e-7, True, "e")
_units("energy", 4.184, True, "c")
_units("energy", 4.1868, True, "cal")
_units("energy", 1.602176487e-19, True, "eV", "ev")
_units("energy", _HP * 3600, False, "HPh", "hh")
_units("energy", 3600.0, True, "Wh", "wh")
_units("energy", _FT * _LBF, False, "flb")
_units("energy", 1055.05585262, False, "BTU", "btu")
# Power (watts)
_units("power", _HP, False, "HP", "h")
_units("power", 735.49875, False, "PS")
_units("power", 1.0, True, "W", "w")
# Magnetism (teslas)
_units("magnetism", 1.0, True, "T")
_units("magnetism", 1e-4, True, "ga")
# Temperature: handled specially, factor unused
_units("temp", 1.0, False, "C", "cel", "F", "fah", "Rank", "Reau")
_units("temp", 1.0, True, "K", "kel")
# Volume (cubic metres)
_units("volume", _GAL / 768, False, "tsp")
_units("volume", 5e-6, False, "tspm")
_units("volume", _GAL / 256, False, "tbs")
_units("volume", _GAL / 128, False, "oz")
_units("volume", _GAL / 16, False, "cup")
_units("volume", _GAL / 8, False, "pt", "us_pt")
_units("volume", _UK_GAL / 8, False, "uk_pt")
_units("volume", _GAL / 4, False, "qt")
_units("volume", _UK_GAL / 4, False, "uk_qt")
_units("volume", _GAL, False, "gal")
_units("volume", _UK_GAL, False, "uk_gal")
_units("volume", 1e-3, True, "l", "L", "lt")
_units("volume", 1e-30, True, "ang3", "ang^3")
_units("volume", 42 * _GAL, False, "barrel")
_units("volume", 2150.42 * _IN ** 3, False, "bushel")
_units("volume", _FT ** 3, False, "ft3", "ft^3")
_units("volume", _IN ** 3, False, "in3", "in^3")
_units("volume", _LY ** 3, False, "ly3", "ly^3")
_units("volume", 1.0, True, "m3", "m^3")
_units("volume", _MI ** 3, False, "mi3", "mi^3")
_units("volume", 0.9144 ** 3, False, "yd3", "yd^3")
_units("volume", _NMI ** 3, False, "Nmi3", "Nmi^3")
_units("volume", _PICAPT ** 3, False, "Pica3", "Pica^3", "Picapt3", "Picapt^3")
_units("volume", 100 * _FT ** 3, False, "GRT", "regton")
_units("volume", 40 * _FT ** 3, False, "MTON")
# Area (square metres)
_units("area", 4046.8564224, False, "uk_acre")
_units("area", 4046.8564224 * (1200 / 3937 / _FT) ** 2, False, "us_acre")
_units("area", 1e-20, True, "ang2", "ang^2")
_units("area", 100.0, True, "ar")
_units("area", _FT ** 2, False, "ft2", "ft^2")
_units("area", 1e4, False, "ha")
_units("area", _IN ** 2, False, "in2", "in^2")
_units("area", _LY ** 2, False, "ly2", "ly^2")
_units("area", 1.0, True, "m2", "m^2")
_units("area", 2500.0, False, "Morgen")
_units("area", _MI ** 2, False, "mi2", "mi^2")
_units("area", _NMI ** 2, False, "Nmi2", "Nmi^2")
_units("area", _PICAPT ** 2, False, "Pica2", "Pica^2", "Picapt2", "Picapt^2")
_units("area", 0.9144 ** 2, False, "yd2", "yd^2")
# Information (bits)
_units("info", 1.0, True, "bit")
_units("info", 8.0, True, "byte")
# Speed (metres per second)
_units("speed", 6080 * _FT / 3600, False, "admkn")
_units("speed", _NMI / 3600, False, "kn")
_units("speed", 1 / 3600, True, "m/h", "m/hr")
_units("speed", 1.0, True, "m/s", "m/sec")
_units("speed", _MI / 3600, False, "mph")

_PREFIXES = {
    "Y": 1e24, "Z": 1e21, "E": 1e18, "P": 1e15, "T": 1e12, "G": 1e9, "M": 1e6, "k": 1e3, "h": 1e2,
    "da": 1e1, "e": 1e1, "d": 1e-1, "c": 1e-2, "m": 1e-3, "u": 1e-6, "n": 1e-9, "p": 1e-12,
    "f": 1e-15, "a": 1e-18, "z": 1e-21, "y": 1e-24,
}
_BIN_PREFIXES = {"ki": 2.0 ** 10, "Mi": 2.0 ** 20, "Gi": 2.0 ** 30, "Ti": 2.0 ** 40, "Pi": 2.0 ** 50,
                 "Ei": 2.0 ** 60, "Zi": 2.0 ** 70, "Yi": 2.0 ** 80}


def _unit(u) -> tuple[str, str, float]:
    """(category, base unit name, factor) for a unit with an optional prefix."""
    if not isinstance(u, str):
        raise XLError(NA)
    hit = _UNITS.get(u)
    if hit is not None:
        return hit[0], u, hit[1]
    for p in (u[:2], u[:1]):
        rest = u[len(p):]
        hit = _UNITS.get(rest)
        if not rest or hit is None or not hit[2]:
            continue
        cat, factor, _ = hit
        if p in _BIN_PREFIXES and cat == "info":
            return cat, rest, factor * _BIN_PREFIXES[p]
        if p in _PREFIXES:
            # in "km2" the prefix is squared; in "kar" or "mL" it isn't
            power = int(rest[-1]) if rest[-1] in "23" else 1
            return cat, rest, factor * _PREFIXES[p] ** power
    raise XLError(NA)


def _to_kelvin(x: float, unit: str) -> float:
    if unit in ("C", "cel"):
        return x + 273.15
    if unit in ("F", "fah"):
        return (x - 32) * 5 / 9 + 273.15
    if unit == "Rank":
        return x * 5 / 9
    if unit == "Reau":
        return x * 5 / 4 + 273.15
    return x


def _from_kelvin(k: float, unit: str) -> float:
    if unit in ("C", "cel"):
        return k - 273.15
    if unit in ("F", "fah"):
        return (k - 273.15) * 9 / 5 + 32
    if unit == "Rank":
        return k * 9 / 5
    if unit == "Reau":
        return (k - 273.15) * 4 / 5
    return k


@fn("CONVERT")
def CONVERT(number, from_unit, to_unit):
    x = _num(number)
    cat1, u1, f1 = _unit(from_unit)
    cat2, u2, f2 = _unit(to_unit)
    if cat1 != cat2:
        raise XLError(NA)
    if from_unit == to_unit:
        return x
    if cat1 == "temp":
        r = _from_kelvin(_to_kelvin(x, u1) * f1, u2) / f2
    else:
        r = x * f1 / f2
    # drop the floating-point noise of the factors (CONVERT(6,"in","ft") is 0.5, not 0.49999...)
    return float(f"{r:.15g}")


# ------------------------------------------------------------------ Bessel functions

_EULER = 0.5772156649015329


def _j_series(x: float, n: int) -> float:
    h = x / 2
    term = math.exp(n * math.log(h) - math.lgamma(n + 1))
    total = term
    k = 0
    while abs(term) > abs(total) * 1e-17:
        k += 1
        term *= -h * h / (k * (n + k))
        total += term
    return total


def _miller_j(x: float, n: int) -> list[float]:
    """J_0(x) .. J_m(x) (m >= n, with the tail negligible) for x > 0."""
    top = max(n, int(x))
    if top > _BESSEL_LIMIT:
        raise XLError(NUM)
    m = 2 * ((top + 20 + int(math.sqrt(60 * top))) // 2)
    if x < 1:  # the power series has no cancellation here (and the recurrence would overflow)
        return [_j_series(x, k) for k in range(m + 1)]
    out = [0.0] * (m + 2)
    out[m] = 1e-300
    nxt, cur = 0.0, 1e-300
    norm = 0.0
    for k in range(m, 0, -1):
        prev = 2 * k / x * cur - nxt
        nxt, cur = cur, prev
        out[k - 1] = cur
        if k - 1 > 0 and (k - 1) % 2 == 0:
            norm += 2 * cur
        if abs(cur) > 1e250:  # rescale to avoid overflow
            scale = 1e-250
            nxt *= scale
            cur *= scale
            norm *= scale
            for i in range(k - 1, m + 1):
                out[i] *= scale
    norm += out[0]
    return [v / norm for v in out[:m + 1]]


def _hankel(x: float, nu: int) -> tuple[float, float]:
    """(J_nu(x), Y_nu(x)) for nu in (0, 1) and large x, by the asymptotic expansion."""
    mu = 4.0 * nu * nu
    p = q = 0.0
    term = 1.0
    k = 0
    while True:
        if k % 2 == 0:
            p += term if k % 4 == 0 else -term
        else:
            q += term if k % 4 == 1 else -term
        k += 1
        new = term * (mu - (2 * k - 1) ** 2) / (k * 8 * x)
        if abs(new) >= abs(term) or abs(new) < 1e-17:
            break
        term = new
    s, c = math.sin(x), math.cos(x)
    r2 = math.sqrt(0.5)
    if nu == 0:  # chi = x - pi/4
        cc, sc = (c + s) * r2, (s - c) * r2
    else:        # chi = x - 3pi/4
        cc, sc = (s - c) * r2, -(s + c) * r2
    a = math.sqrt(2 / (math.pi * x))
    return a * (p * cc - q * sc), a * (p * sc + q * cc)


# BESSELJ is Numerical Recipes' bessj (rational approximations for J0 and J1, and Miller's
# algorithm with ACC=40 for higher orders), which is accurate to about 1e-8 rather than exact.
# It reproduces Microsoft's example BESSELJ(1.9,2) = 0.329925829 (exactly, it's 0.3299257277).


def _nr_j0(ax: float) -> float:
    if ax < 8:
        y = ax * ax
        a1 = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (
            77392.33017 + y * -184.9052456))))
        a2 = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (
            267.8532712 + y))))
        return a1 / a2
    z = 8 / ax
    y = z * z
    xx = ax - 0.785398164
    a1 = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)))
    a2 = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (
        0.7621095161e-6 - y * 0.934935152e-7)))
    return math.sqrt(0.636619772 / ax) * (math.cos(xx) * a1 - z * math.sin(xx) * a2)


def _nr_j1(ax: float) -> float:
    if ax < 8:
        y = ax * ax
        a1 = ax * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (
            15704.48260 + y * -30.16036606)))))
        a2 = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (
            376.9991397 + y))))
        return a1 / a2
    z = 8 / ax
    y = z * z
    xx = ax - 2.356194491
    a1 = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)))
    a2 = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)))
    return math.sqrt(0.636619772 / ax) * (math.cos(xx) * a1 - z * math.sin(xx) * a2)


def _bessel_j(x: float, n: int) -> float:
    sign = -1 if x < 0 and n % 2 else 1
    ax = abs(x)
    if n == 0:
        return _nr_j0(ax)
    if n == 1:
        return sign * _nr_j1(ax)
    if ax == 0:
        return 0.0
    tox = 2 / ax
    if ax > n:  # upward recurrence
        bjm, bj = _nr_j0(ax), _nr_j1(ax)
        for j in range(1, n):
            bjm, bj = bj, j * tox * bj - bjm
        return sign * bj
    # Miller's downward recurrence, normalised by J0 + 2(J2 + J4 + ...) = 1
    m = 2 * ((n + int(math.sqrt(40 * n))) // 2)
    even = False
    bjp = ans = total = 0.0
    bj = 1.0
    for j in range(m, 0, -1):
        bjm = j * tox * bj - bjp
        bjp, bj = bj, bjm
        if abs(bj) > 1e10:
            bj *= 1e-10
            bjp *= 1e-10
            ans *= 1e-10
            total *= 1e-10
        if even:
            total += bj
        even = not even
        if j == n:
            ans = bjp
    return sign * ans / (2 * total - bj)


def _bessel_y(x: float, n: int) -> float:
    if x >= 30:
        _, y0 = _hankel(x, 0)
        _, y1 = _hankel(x, 1)
    else:
        j = _miller_j(x, 1)
        lg = math.log(x / 2) + _EULER
        s0 = math.fsum((-1) ** k * j[2 * k] / k for k in range(1, len(j) // 2))
        s1 = math.fsum((-1) ** k * (2 * k + 1) * j[2 * k + 1] / (k * (k + 1))
                       for k in range(1, (len(j) - 1) // 2))
        y0 = 2 / math.pi * (lg * j[0] - 2 * s0)
        y1 = 2 / math.pi * (-j[0] / x + (lg - 1) * j[1] - s1)
    if n == 0:
        return y0
    for k in range(1, n):
        y0, y1 = y1, 2 * k / x * y1 - y0
        if math.isinf(y1):
            raise XLError(NUM)
    return y1


def _bessel_i(x: float, n: int) -> float:
    sign = -1 if x < 0 and n % 2 else 1
    h = abs(x) / 2
    if h == 0:
        return 1.0 if n == 0 else 0.0
    term = math.exp(n * math.log(h) - math.lgamma(n + 1))
    total = term
    q = h * h
    k = 0
    while True:
        k += 1
        term *= q / (k * (n + k))
        total += term
        if term <= total * 1e-17:
            break
        if math.isinf(total):
            raise XLError(NUM)
    return sign * total


def _bessel_k01(x: float) -> tuple[float, float]:
    """(K_0(x), K_1(x)) for x > 0.

    These are the polynomial approximations of Abramowitz and Stegun 9.8.5-9.8.8 (relative error
    around 1e-7) rather than exact values: LibreOffice uses them too, and they reproduce the
    digits of Microsoft's example BESSELK(1.5,1) = 0.277387804 (exactly, it's 0.2773878005).
    """
    if x <= 2:
        h = x / 2
        y = h * h
        lg = math.log(h)
        k0 = -lg * _bessel_i(x, 0) + (-0.57721566 + y * (0.42278420 + y * (0.23069756 + y * (
            0.3488590e-1 + y * (0.262698e-2 + y * (0.10750e-3 + y * 0.74e-5))))))
        k1 = lg * _bessel_i(x, 1) + (1 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (
            -0.1919402e-1 + y * (-0.110404e-2 + y * -0.4686e-4)))))) / x
        return k0, k1
    y = 2 / x
    a = math.exp(-x) / math.sqrt(x)
    k0 = a * (1.25331414 + y * (-0.7832358e-1 + y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (
        0.587872e-2 + y * (-0.251540e-2 + y * 0.53208e-3))))))
    k1 = a * (1.25331414 + y * (0.23498619 + y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (
        -0.780353e-2 + y * (0.325614e-2 + y * -0.68245e-3))))))
    return k0, k1


def _bessel_k(x: float, n: int) -> float:
    k0, k1 = _bessel_k01(x)
    if n == 0:
        return k0
    for k in range(1, n):
        k0, k1 = k1, 2 * k / x * k1 + k0
        if math.isinf(k1):
            raise XLError(NUM)
    return k1


# Larger orders or arguments would take too long (the recurrences run over every order up to n,
# and Miller's algorithm over every order up to about max(n, x)).
_BESSEL_LIMIT = 1_000_000


def _bessel_args(x, n) -> tuple[float, int]:
    x = _num(x)
    n = _trunc(n)
    if n < 0 or n > _BESSEL_LIMIT:
        raise XLError(NUM)
    return x, n


@fn("BESSELJ")
def BESSELJ(x, n):
    return _bessel_j(*_bessel_args(x, n))


@fn("BESSELY")
def BESSELY(x, n):
    x, n = _bessel_args(x, n)
    if x <= 0:
        raise XLError(NUM)
    return _bessel_y(x, n)


@fn("BESSELI")
def BESSELI(x, n):
    return _bessel_i(*_bessel_args(x, n))


@fn("BESSELK")
def BESSELK(x, n):
    x, n = _bessel_args(x, n)
    if x <= 0:
        raise XLError(NUM)
    return _bessel_k(x, n)


# ------------------------------------------------------------------ complex numbers

_NUM_PART = r"(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?"
_COMPLEX_RE = re.compile(
    rf"(?P<re>[+-]?{_NUM_PART})?"
    rf"(?:(?P<isign>[+-])?(?P<im>{_NUM_PART})?(?P<suffix>[ij]))?")


def _parse_complex(v) -> tuple[complex, str | None]:
    """A complex argument and its suffix ('i', 'j', or None if it had no imaginary part)."""
    if type(v) is bool:
        raise XLError(VALUE)
    if v is None or v is MISSING:
        return 0j, None
    if isinstance(v, (int, float)):
        return complex(v, 0), None
    s = to_text(v)
    if s == "":
        return 0j, None
    m = _COMPLEX_RE.fullmatch(s)
    if m is None or not s:
        raise XLError(NUM)
    re_s, isign, im_s, suffix = m.group("re", "isign", "im", "suffix")
    if suffix is None:
        return complex(float(re_s), 0), None
    if re_s is not None and isign is None and im_s is None:
        # "3i" / "-2.5j": the number was the imaginary part
        return complex(0, float(re_s)), suffix
    if re_s is not None and isign is None:
        raise XLError(NUM)  # "3 4i" has no operator between the parts
    im = float(im_s) if im_s is not None else 1.0
    if isign == "-":
        im = -im
    return complex(float(re_s) if re_s is not None else 0.0, im), suffix


def _fmt_part(x: float) -> str:
    if x == 0:
        return "0"
    return format_general_number(float(f"{x:.15g}"))


def _format_complex(z: complex, suffix: str | None = "i") -> str:
    suffix = suffix or "i"
    re_, im = z.real, z.imag
    if not (math.isfinite(re_) and math.isfinite(im)):
        raise XLError(NUM)
    re_ = float(f"{re_:.15g}")
    im = float(f"{im:.15g}")
    if im == 0:
        return _fmt_part(re_)
    if im == 1:
        ims = suffix
    elif im == -1:
        ims = "-" + suffix
    else:
        ims = _fmt_part(im) + suffix
    if re_ == 0:
        return ims
    return _fmt_part(re_) + ("" if ims.startswith("-") else "+") + ims


def _suffix(*suffixes) -> str:
    used = {s for s in suffixes if s}
    if len(used) > 1:
        raise XLError(VALUE)
    return used.pop() if used else "i"


@fn("COMPLEX")
def COMPLEX(real_num, i_num, suffix=None):
    re_, im = _num(real_num), _num(i_num)
    if suffix is None:
        suffix = "i"
    elif type(suffix) is not str or suffix not in ("i", "j", ""):
        raise XLError(VALUE)
    return _format_complex(complex(re_, im), suffix or "i")


@fn("IMREAL")
def IMREAL(z):
    return _parse_complex(z)[0].real


@fn("IMAGINARY")
def IMAGINARY(z):
    return _parse_complex(z)[0].imag


@fn("IMABS")
def IMABS(z):
    return abs(_parse_complex(z)[0])


@fn("IMARGUMENT")
def IMARGUMENT(z):
    c = _parse_complex(z)[0]
    if c == 0:
        raise XLError(DIV0)
    return math.atan2(c.imag, c.real)


@fn("IMCONJUGATE")
def IMCONJUGATE(z):
    c, s = _parse_complex(z)
    return _format_complex(c.conjugate(), s)


def _unary(name: str, f, nonzero: bool = False):
    def impl(z):
        c, s = _parse_complex(z)
        if nonzero and c == 0:
            raise XLError(NUM)
        try:
            r = f(c)
        except (ZeroDivisionError, ValueError, OverflowError):
            raise XLError(NUM)
        return _format_complex(r, s)

    fn(name)(impl)


def _polar_pow(c: complex, p: float) -> complex:
    r = abs(c)
    if r == 0:
        if p > 0:
            return 0j
        raise XLError(NUM)
    t = math.atan2(c.imag, c.real) * p
    rp = r ** p
    return complex(rp * math.cos(t), rp * math.sin(t))


def _ln(c: complex) -> complex:
    return complex(math.log(abs(c)), math.atan2(c.imag, c.real))


_unary("IMEXP", cmath.exp)
_unary("IMLN", _ln, nonzero=True)
_unary("IMLOG10", lambda c: _ln(c) / math.log(10), nonzero=True)
_unary("IMLOG2", lambda c: _ln(c) / math.log(2), nonzero=True)
_unary("IMSQRT", lambda c: _polar_pow(c, 0.5))
# The quotients use the closed forms below rather than complex division: they reproduce the
# last digit of Microsoft's documented examples.


def _tan(c: complex) -> complex:
    x, y = c.real, c.imag
    d = math.cos(2 * x) + math.cosh(2 * y)
    return complex(math.sin(2 * x) / d, math.sinh(2 * y) / d)


def _cot(c: complex) -> complex:
    x, y = c.real, c.imag
    d = math.cosh(2 * y) - math.cos(2 * x)
    return complex(math.sin(2 * x) / d, -math.sinh(2 * y) / d)


def _sec(c: complex) -> complex:
    x, y = c.real, c.imag
    d = math.cos(2 * x) + math.cosh(2 * y)
    return complex(2 * math.cos(x) * math.cosh(y) / d, 2 * math.sin(x) * math.sinh(y) / d)


def _csc(c: complex) -> complex:
    x, y = c.real, c.imag
    d = math.cosh(2 * y) - math.cos(2 * x)
    return complex(2 * math.sin(x) * math.cosh(y) / d, -2 * math.cos(x) * math.sinh(y) / d)


def _sech(c: complex) -> complex:
    x, y = c.real, c.imag
    d = math.cosh(2 * x) + math.cos(2 * y)
    return complex(2 * math.cosh(x) * math.cos(y) / d, -2 * math.sinh(x) * math.sin(y) / d)


def _csch(c: complex) -> complex:
    x, y = c.real, c.imag
    d = math.cosh(2 * x) - math.cos(2 * y)
    return complex(2 * math.sinh(x) * math.cos(y) / d, -2 * math.cosh(x) * math.sin(y) / d)


_unary("IMSIN", cmath.sin)
_unary("IMCOS", cmath.cos)
_unary("IMTAN", _tan)
_unary("IMSINH", cmath.sinh)
_unary("IMCOSH", cmath.cosh)
_unary("IMSEC", _sec)
_unary("IMCSC", _csc, nonzero=True)
_unary("IMCOT", _cot, nonzero=True)
_unary("IMSECH", _sech)
_unary("IMCSCH", _csch, nonzero=True)


@fn("IMPOWER")
def IMPOWER(z, number):
    c, s = _parse_complex(z)
    return _format_complex(_polar_pow(c, _num(number)), s)


@fn("IMDIV")
def IMDIV(z1, z2):
    a, s1 = _parse_complex(z1)
    b, s2 = _parse_complex(z2)
    if b == 0:
        raise XLError(NUM)
    return _format_complex(a / b, _suffix(s1, s2))


@fn("IMSUB")
def IMSUB(z1, z2):
    a, s1 = _parse_complex(z1)
    b, s2 = _parse_complex(z2)
    return _format_complex(a - b, _suffix(s1, s2))


def _complex_args(ctx, args):
    for a in args:
        if isinstance(a, RefValue):
            vals = iter_ref(ctx, a, include_blank=True)
        elif isinstance(a, Array):
            vals = a.flat()
        else:
            vals = [a]
        for v in vals:
            yield _parse_complex(v)


@fn("IMSUM", kind="raw", min_args=1)
def IMSUM(ctx, *args):
    total, suffixes = 0j, []
    for c, s in _complex_args(ctx, args):
        total += c
        suffixes.append(s)
    return _format_complex(total, _suffix(*suffixes))


@fn("IMPRODUCT", kind="raw", min_args=1)
def IMPRODUCT(ctx, *args):
    total, suffixes = 1 + 0j, []
    for c, s in _complex_args(ctx, args):
        total *= c
        suffixes.append(s)
    return _format_complex(total, _suffix(*suffixes))
