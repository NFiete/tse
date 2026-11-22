"""Excel number-format rendering.

Supports General, digit placeholders (0 # ?), thousands separators and scaling,
percent, scientific notation, fractions, literal text, [Color] tags, conditions,
up to four ;-separated sections, text sections (@), and date/time codes including
elapsed time ([h]:mm).
"""
from __future__ import annotations

import math
import re
from decimal import ROUND_HALF_UP, Decimal
from fractions import Fraction
from functools import lru_cache

from ..formula.errors import ExcelError
from ..model.dates import from_serial, serial_date_parts

COLOR_NAMES = {
    "black": "000000", "blue": "0000FF", "cyan": "00FFFF", "green": "00FF00",
    "magenta": "FF00FF", "red": "FF0000", "white": "FFFFFF", "yellow": "FFFF00",
}

_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
           "October", "November", "December"]
_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


# --------------------------------------------------------------------------- parsing

class Section:
    __slots__ = ("tokens", "color", "cond", "kind", "percent", "thousands", "scale", "has_ampm",
                 "frac_digits_sec", "raw")

    def __init__(self):
        self.tokens: list[tuple[str, str]] = []
        self.color: str | None = None
        self.cond: tuple[str, float] | None = None
        self.kind = "number"  # number | date | text | general | fraction
        self.percent = 0
        self.thousands = False
        self.scale = 0
        self.has_ampm = False
        self.frac_digits_sec = 0
        self.raw = ""


def _split_sections(fmt: str) -> list[str]:
    out, cur, i, n = [], [], 0, len(fmt)
    while i < n:
        ch = fmt[i]
        if ch == '"':
            j = fmt.find('"', i + 1)
            j = n - 1 if j < 0 else j
            cur.append(fmt[i:j + 1])
            i = j + 1
            continue
        if ch == "\\" and i + 1 < n:
            cur.append(fmt[i:i + 2])
            i += 2
            continue
        if ch == "[":
            j = fmt.find("]", i)
            j = n - 1 if j < 0 else j
            cur.append(fmt[i:j + 1])
            i = j + 1
            continue
        if ch == ";":
            out.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


_COND_RE = re.compile(r"^\[(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)\]$")
_DATE_CHARS = set("ymdhsYMDHS")


def _parse_section(text: str) -> Section:
    sec = Section()
    sec.raw = text
    toks = sec.tokens
    i, n = 0, len(text)
    has_digit_ph = False
    has_date = False
    while i < n:
        ch = text[i]
        if ch == '"':
            j = text.find('"', i + 1)
            if j < 0:
                j = n
            toks.append(("lit", text[i + 1:j]))
            i = j + 1
            continue
        if ch == "\\":
            if i + 1 < n:
                toks.append(("lit", text[i + 1]))
            i += 2
            continue
        if ch == "_":
            toks.append(("lit", " "))
            i += 2
            continue
        if ch == "*":
            i += 2  # repeat-fill character: ignored
            continue
        if ch == "[":
            j = text.find("]", i)
            if j < 0:
                j = n
            body = text[i + 1:j]
            low = body.lower()
            if low in COLOR_NAMES:
                sec.color = COLOR_NAMES[low]
            elif low.startswith("color") and low[5:].isdigit():
                from openpyxl.styles.colors import COLOR_INDEX

                idx = int(low[5:]) + 7
                if 0 <= idx < len(COLOR_INDEX):
                    sec.color = COLOR_INDEX[idx][-6:]
            elif _COND_RE.match("[" + body + "]"):
                m = _COND_RE.match("[" + body + "]")
                sec.cond = (m.group(1), float(m.group(2)))
            elif body.startswith("$"):
                sym = body[1:].split("-")[0]
                if sym:
                    toks.append(("lit", sym))
            elif re.fullmatch(r"[hH]+|[mM]+|[sS]+", body):
                toks.append(("elapsed", low))
                has_date = True
            i = j + 1
            continue
        if text[i:i + 7].lower() == "general":
            toks.append(("general", ""))
            i += 7
            continue
        if text[i:i + 5].upper() == "AM/PM":
            toks.append(("ampm", text[i:i + 5]))
            sec.has_ampm = True
            i += 5
            continue
        if text[i:i + 3].upper() == "A/P":
            toks.append(("ap", text[i:i + 3]))
            sec.has_ampm = True
            i += 3
            continue
        if ch in "0#?":
            toks.append(("ph", ch))
            has_digit_ph = True
            i += 1
            continue
        if ch == ".":
            toks.append(("dot", "."))
            i += 1
            continue
        if ch == ",":
            toks.append(("comma", ","))
            i += 1
            continue
        if ch == "%":
            toks.append(("pct", "%"))
            sec.percent += 1
            i += 1
            continue
        if ch in "eE" and i + 1 < n and text[i + 1] in "+-":
            toks.append(("exp", text[i + 1]))
            i += 2
            continue
        if ch == "@":
            toks.append(("text", "@"))
            i += 1
            continue
        if ch == "/":
            toks.append(("slash", "/"))
            i += 1
            continue
        if ch in _DATE_CHARS or ch in "eg" and False:
            j = i
            while j < n and text[j].lower() == ch.lower():
                j += 1
            toks.append(("date", text[i:j].lower()))
            has_date = True
            i = j
            continue
        if ch.isdigit() and toks and toks[-1][0] == "slash":
            j = i
            while j < n and text[j].isdigit():
                j += 1
            toks.append(("den", text[i:j]))
            i = j
            continue
        toks.append(("lit", ch))
        i += 1
    if any(t[0] == "general" for t in toks):
        sec.kind = "general"
    elif has_date and not (has_digit_ph and not any(t[0] == "date" and t[1][0] in "s" for t in toks)):
        sec.kind = "date"
    elif any(t[0] == "text" for t in toks) and not has_digit_ph:
        sec.kind = "text"
    elif any(t[0] == "slash" for t in toks) and has_digit_ph:
        sec.kind = "fraction"
    else:
        sec.kind = "number"
    if sec.kind == "number":
        _analyze_commas(sec)
    if sec.kind == "date":
        _resolve_minutes(sec)
    return sec


def _analyze_commas(sec: Section):
    """A comma between digit placeholders = thousands; trailing commas = scale by 1000."""
    toks = sec.tokens
    ph_idx = [i for i, t in enumerate(toks) if t[0] == "ph"]
    if not ph_idx:
        return
    first, last = ph_idx[0], ph_idx[-1]
    dot = next((i for i, t in enumerate(toks) if t[0] == "dot"), None)
    new = []
    for i, t in enumerate(toks):
        if t[0] == "comma":
            if first < i < last and (dot is None or i < dot):
                sec.thousands = True
                continue
            if i > last or (dot is not None and i == dot - 1 and i > first):
                # trailing scaling comma (after last placeholder or right before the decimal point)
                j = i + 1
                if all(tt[0] in ("comma", "lit", "pct", "text") for tt in toks[j:]) or (
                        dot is not None and i < dot):
                    sec.scale += 1
                    continue
            new.append(("lit", ","))
            continue
        new.append(t)
    sec.tokens = new


def _resolve_minutes(sec: Section):
    toks = sec.tokens
    date_idx = [i for i, t in enumerate(toks) if t[0] in ("date", "elapsed")]
    for pos, i in enumerate(date_idx):
        code = toks[i][1]
        if toks[i][0] != "date" or code[0] != "m" or len(code) > 2:
            continue
        prev = toks[date_idx[pos - 1]] if pos > 0 else None
        nxt = toks[date_idx[pos + 1]] if pos + 1 < len(date_idx) else None
        if (prev and prev[1][0] in "h" or (prev and prev[0] == "elapsed" and prev[1][0] == "h")) or (
                nxt and (nxt[1][0] == "s" or (nxt[0] == "elapsed" and nxt[1][0] == "s"))):
            toks[i] = ("min", code)
    # fractional seconds: dot followed by zeros after a seconds token
    out = []
    i = 0
    while i < len(toks):
        t = toks[i]
        if t[0] == "dot" and out and out[-1][0] in ("date", "elapsed") and out[-1][1][0] == "s":
            j = i + 1
            k = 0
            while j < len(toks) and toks[j] == ("ph", "0"):
                k += 1
                j += 1
            if k:
                out.append(("fsec", str(k)))
                sec.frac_digits_sec = max(sec.frac_digits_sec, k)
                i = j
                continue
        if t[0] == "ph":
            out.append(("lit", t[1]))
        elif t[0] == "dot":
            out.append(("lit", "."))
        elif t[0] == "comma":
            out.append(("lit", ","))
        elif t[0] == "slash":
            out.append(("lit", "/"))
        else:
            out.append(t)
        i += 1
    sec.tokens = out


@lru_cache(maxsize=1024)
def parse_format(fmt: str) -> tuple[Section, ...]:
    if not fmt:
        fmt = "General"
    return tuple(_parse_section(s) for s in _split_sections(fmt))


@lru_cache(maxsize=1024)
def is_date_format(fmt: str | None) -> bool:
    if not fmt or fmt == "General":
        return False
    secs = parse_format(fmt)
    return secs[0].kind == "date"


def is_percent_format(fmt: str | None) -> bool:
    return bool(fmt) and parse_format(fmt)[0].percent > 0


# --------------------------------------------------------------------------- rendering


def general(x, maxlen: int = 11) -> str:
    """Excel 'General' rendering of a number within maxlen characters (best effort)."""
    if type(x) is bool:
        return "TRUE" if x else "FALSE"
    if x == 0:
        return "0"
    ax = abs(x)
    neg = 1 if x < 0 else 0
    if type(x) is int and len(str(x)) <= maxlen:
        return str(x)
    if 1e-9 <= ax < 1e11 or (type(x) is int and ax < 1e11):
        intpart = int(ax)
        intlen = len(str(intpart)) + neg
        if intlen <= maxlen:
            sig_int = len(str(intpart)) if intpart else 0
            dec = min(maxlen - intlen - 1, 15 - sig_int if sig_int else 15 + _leading_zeros(ax))
            if dec > 0:
                s = f"{x:.{dec}f}".rstrip("0").rstrip(".")
            else:
                s = str(round(x))
            if s not in ("0", "-0") and len(s) <= maxlen:
                return s
    # scientific
    for p in range(min(maxlen, 14), -1, -1):
        m, e = f"{x:.{p}E}".split("E")
        if "." in m:
            m = m.rstrip("0").rstrip(".")
        ei = int(e)
        s = f"{m}E{'+' if ei >= 0 else '-'}{abs(ei):02d}"
        if len(s) <= maxlen:
            return s
    return "#" * maxlen


def _leading_zeros(ax: float) -> int:
    return max(0, -int(math.floor(math.log10(ax))) - 1) if 0 < ax < 1 else 0


def _round_half_up(x: float, nd: int) -> Decimal:
    d = Decimal(f"{x:.15g}") if isinstance(x, float) else Decimal(x)
    return d.quantize(Decimal(1).scaleb(-nd), rounding=ROUND_HALF_UP)


def _render_number(x: float, sec: Section) -> str:
    """x is already sign-adjusted (abs value for negative sections)."""
    toks = sec.tokens
    v = x * (100 ** sec.percent) / (1000 ** sec.scale)
    exp_i = next((i for i, t in enumerate(toks) if t[0] == "exp"), None)
    if exp_i is not None:
        return _render_sci(v, toks, exp_i)
    dot = next((i for i, t in enumerate(toks) if t[0] == "dot"), None)
    ph_idx = [i for i, t in enumerate(toks) if t[0] == "ph"]
    if not ph_idx:
        # no placeholders: just literals (e.g. "Total")
        return "".join(t[1] for t in toks if t[0] in ("lit", "pct"))
    first_ph = ph_idx[0]
    int_end = dot if dot is not None else (ph_idx[-1] + 1)
    frac_ph = [i for i in ph_idx if dot is not None and i > dot]
    nd = len(frac_ph)
    d = _round_half_up(abs(v), nd)
    s = f"{d:.{nd}f}"
    int_str, _, frac_str = s.partition(".")
    neg = v < 0 and d != 0
    if int_str == "0":
        int_str = ""
    int_ph = [i for i in ph_idx if i < int_end]
    out: list[str] = []
    # prefix literals before the first placeholder
    for t in toks[:first_ph]:
        out.append(t[1] if t[0] in ("lit", "pct") else "")
    # integer part
    if sec.thousands:
        zeros = sum(1 for i in int_ph if toks[i][1] == "0")
        digits = int_str.rjust(zeros, "0") if len(int_str) < zeros else int_str
        grouped = f"{int(digits):,}" if digits else ""
        if digits and len(digits) > len(str(int(digits))):
            grouped = digits[: len(digits) - len(str(int(digits)))] + grouped if int(digits) else digits
        # pad with spaces for '?' placeholders
        qs = sum(1 for i in int_ph if toks[i][1] == "?")
        if qs and len(int_str) < len(int_ph):
            grouped = " " * min(qs, len(int_ph) - len(int_str)) + grouped
        out.append(grouped)
        for t in toks[first_ph:int_end]:
            if t[0] == "lit":
                out.append(t[1])
    else:
        region = toks[first_ph:int_end]
        digits = list(int_str)
        rev: list[str] = []
        for t in reversed(region):
            if t[0] == "ph":
                if digits:
                    rev.append(digits.pop())
                elif t[1] == "0":
                    rev.append("0")
                elif t[1] == "?":
                    rev.append(" ")
            elif t[0] in ("lit", "pct"):
                rev.append(t[1])
        res = "".join(reversed(rev))
        if digits:
            # extra digits go before the first placeholder
            res = "".join(digits) + res
        out.append(res)
    # decimal part
    if dot is not None:
        out.append(".")
        frac_region = toks[dot + 1:]
        sig = len(frac_str.rstrip("0"))
        k = 0
        for t in frac_region:
            if t[0] == "ph":
                ch = frac_str[k] if k < len(frac_str) else "0"
                if k >= sig:
                    ch = {"0": "0", "#": "", "?": " "}[t[1]]
                out.append(ch)
                k += 1
            elif t[0] in ("lit", "pct", "text"):
                out.append(t[1])
    else:
        for t in toks[int_end:]:
            if t[0] in ("lit", "pct", "text"):
                out.append(t[1])
    res = "".join(out)
    return ("-" + res) if neg else res


def _render_sci(v: float, toks, exp_i: int) -> str:
    mant_toks = toks[:exp_i]
    exp_toks = toks[exp_i + 1:]
    sign_char = toks[exp_i][1]
    dot = next((i for i, t in enumerate(mant_toks) if t[0] == "dot"), None)
    int_ph = [t for i, t in enumerate(mant_toks) if t[0] == "ph" and (dot is None or i < dot)]
    frac_ph = [t for i, t in enumerate(mant_toks) if t[0] == "ph" and dot is not None and i > dot]
    nd = len(frac_ph)
    k = max(1, len(int_ph))
    if v == 0:
        e = 0
        mant = 0.0
    else:
        e = math.floor(math.log10(abs(v)))
        if k > 1 and any(t[1] == "#" for t in int_ph):
            e = math.floor(e / k) * k
        else:
            e = e - (k - 1) if k > 1 else e
        mant = v / (10 ** e)
        if float(_round_half_up(abs(mant), nd)) >= 10 ** (k if any(t[1] == "#" for t in int_ph) and k > 1 else k):
            e += k if any(t[1] == "#" for t in int_ph) and k > 1 else 1
            mant = v / (10 ** e)
    d = _round_half_up(abs(mant), nd)
    s = f"{d:.{nd}f}"
    ip, _, fp = s.partition(".")
    prefix = "".join(t[1] for t in mant_toks if t[0] == "lit" and mant_toks.index(t) < (
        next((i for i, tt in enumerate(mant_toks) if tt[0] == "ph"), 0)))
    out = prefix + ("-" if v < 0 else "") + ip + (("." + fp) if dot is not None else "")
    e_digits = sum(1 for t in exp_toks if t[0] == "ph")
    es = str(abs(e)).rjust(max(e_digits, 1), "0")
    sign = "-" if e < 0 else ("+" if sign_char == "+" else "")
    suffix = "".join(t[1] for t in exp_toks if t[0] == "lit")
    return f"{out}E{sign}{es}{suffix}"


_FRAC_RE = re.compile(r"^(?P<pre>[^#0?]*?)(?:(?P<whole>[#0?,]+)(?P<sp>\s+))?(?P<num>[#0?]+)\s*/\s*(?P<den>[#0?]+|\d+)(?P<post>.*)$")


def _render_fraction(x: float, sec: Section) -> str:
    raw = "".join(t[1] if t[0] != "den" else t[1] for t in sec.tokens)
    m = _FRAC_RE.match(raw)
    if not m:
        return general(x)
    neg = x < 0
    ax = abs(x)
    whole = int(ax) if m.group("whole") else 0
    frac = ax - whole
    den_spec = m.group("den")
    if den_spec.isdigit():
        den = int(den_spec)
        num = round(frac * den)
    else:
        maxden = 10 ** len(den_spec) - 1
        f = Fraction(frac).limit_denominator(maxden)
        num, den = f.numerator, f.denominator
    if m.group("whole") and num == den and den:
        whole += 1
        num = 0
    pre, post = m.group("pre"), m.group("post")
    width = len(den_spec)
    if m.group("whole"):
        if num == 0:
            body = f"{whole}" + " " * (len(m.group("sp")) + len(m.group("num")) + 1 + width)
            if whole == 0:
                body = "0" + body[1:]
        else:
            ws = f"{whole}" if whole else ""
            body = f"{ws}{m.group('sp') if ws else ' ' * len(m.group('sp'))}{str(num).rjust(len(m.group('num')))}/{str(den).ljust(width)}"
            if not ws:
                body = body.lstrip() if not m.group("whole").strip("#") else body
    else:
        body = f"{str(num).rjust(len(m.group('num')))}/{str(den).ljust(width)}"
    return ("-" if neg else "") + pre + body.rstrip() + post


def _render_date(serial: float, sec: Section, epoch1904: bool) -> str:
    if serial < 0 or serial > 2958465.99999:
        return None  # caller shows ####
    nd = sec.frac_digits_sec
    # round to displayed precision
    unit = 86400 * (10 ** nd)
    total = round(serial * unit)
    days, rem = divmod(total, unit)
    secs_total = rem / (10 ** nd)
    y, mo, d = serial_date_parts(days, epoch1904)
    wd = None
    if not (y == 1900 and mo in (1, 2) and not epoch1904 and (d == 0 or (mo == 2 and d == 29))):
        wd = from_serial(days, epoch1904).weekday()
    elif d == 0:
        wd = 5  # Excel shows Saturday for 1900-01-00
    else:
        wd = 2
    hh = int(secs_total // 3600)
    mi = int(secs_total // 60 % 60)
    ss_f = secs_total - hh * 3600 - mi * 60
    ss = int(ss_f)
    out = []
    for kind, code in sec.tokens:
        if kind == "lit":
            out.append(code)
        elif kind == "date":
            c0 = code[0]
            ln = len(code)
            if c0 == "y":
                out.append(f"{y % 100:02d}" if ln <= 2 else f"{y:04d}")
            elif c0 == "m":
                if ln == 1:
                    out.append(str(mo))
                elif ln == 2:
                    out.append(f"{mo:02d}")
                elif ln == 3:
                    out.append(_MONTHS[mo - 1][:3])
                elif ln == 5:
                    out.append(_MONTHS[mo - 1][0])
                else:
                    out.append(_MONTHS[mo - 1])
            elif c0 == "d":
                if ln == 1:
                    out.append(str(d))
                elif ln == 2:
                    out.append(f"{d:02d}")
                elif ln == 3:
                    out.append(_DAYS[wd][:3])
                else:
                    out.append(_DAYS[wd])
            elif c0 == "h":
                h = hh
                if sec.has_ampm:
                    h = hh % 12 or 12
                out.append(str(h) if ln == 1 else f"{h:02d}")
            elif c0 == "s":
                out.append(str(ss) if ln == 1 else f"{ss:02d}")
            else:
                out.append(code)
        elif kind == "min":
            out.append(str(mi) if len(code) == 1 else f"{mi:02d}")
        elif kind == "fsec":
            k = int(code)
            frac = ss_f - ss
            out.append("." + f"{round(frac * 10 ** k):0{k}d}"[:k])
        elif kind == "elapsed":
            c0 = code[0]
            whole_secs = days * 86400 + secs_total
            if c0 == "h":
                v = int(whole_secs // 3600)
            elif c0 == "m":
                v = int(whole_secs // 60)
            else:
                v = int(whole_secs)
            out.append(str(v).rjust(len(code), "0"))
        elif kind == "ampm":
            pm = hh >= 12
            txt = "PM" if pm else "AM"
            out.append(txt if code[0].isupper() else txt.lower())
        elif kind == "ap":
            pm = hh >= 12
            txt = "P" if pm else "A"
            out.append(txt if code[0].isupper() else txt.lower())
        elif kind == "text":
            out.append("")
    return "".join(out)


def _cond_match(cond, x) -> bool:
    op, v = cond
    return {"<": x < v, ">": x > v, "=": x == v, "<=": x <= v, ">=": x >= v, "<>": x != v}[op]


def _pick_section(secs, x):
    """Return (section, magnitude_to_render, show_minus)."""
    if any(s.cond for s in secs[:2]):
        if secs[0].cond and _cond_match(secs[0].cond, x):
            return secs[0], abs(x), x < 0
        if len(secs) > 1:
            if secs[1].cond is None or _cond_match(secs[1].cond, x):
                return secs[1], abs(x), False
            if len(secs) > 2:
                return secs[2], abs(x), x < 0
        return secs[0], abs(x), x < 0
    k = min(len(secs), 3)
    if k == 1:
        return secs[0], abs(x), x < 0
    if x > 0 or (x == 0 and k == 2):
        return secs[0], x, False
    if x < 0:
        return secs[1], -x, False
    return secs[2], x, False


def format_value(value, fmt: str | None, epoch1904: bool = False) -> tuple[str, str | None]:
    """Render a cell value with an Excel number format. Returns (text, color hex or None)."""
    if value is None:
        return "", None
    if isinstance(value, ExcelError):
        return value.code, None
    if type(value) is bool:
        return ("TRUE" if value else "FALSE"), None
    secs = parse_format(fmt or "General")
    if isinstance(value, str):
        tsec = secs[3] if len(secs) >= 4 else (secs[0] if secs[0].kind == "text" else None)
        if tsec is None or not any(t[0] == "text" for t in tsec.tokens):
            return value, (tsec.color if tsec is not None else None)
        return "".join(value if t[0] == "text" else (t[1] if t[0] == "lit" else "") for t in tsec.tokens), tsec.color
    x = value
    if isinstance(x, float) and (math.isnan(x) or math.isinf(x)):
        return "#NUM!", None
    sec, v, minus = _pick_section(secs, x)
    sign = "-" if minus else ""
    if sec.kind == "general":
        before, after, seen = [], [], False
        for t in sec.tokens:
            if t[0] == "general":
                seen = True
            elif t[0] == "lit":
                (after if seen else before).append(t[1])
        return sign + "".join(before) + general(v) + "".join(after), sec.color
    if sec.kind == "date":
        txt = None if minus else _render_date(v, sec, epoch1904)
        if txt is None:
            return "#" * 8, None
        return txt, sec.color
    if sec.kind == "text":
        return "".join(t[1] for t in sec.tokens if t[0] == "lit"), sec.color
    txt = _render_fraction(v, sec) if sec.kind == "fraction" else _render_number(v, sec)
    return sign + txt, sec.color


def fit_general(x, width: int) -> str:
    """General-format number sized for a column width (Excel shrinks precision to fit)."""
    if width <= 0:
        return ""
    s = general(x, max(1, min(width, 11)))
    if len(s) > width:
        return "#" * width
    return s


def edit_text_for_date(serial: float, fmt: str, epoch1904: bool = False) -> str:
    """Text shown in the formula bar for a date/time cell."""
    has_time = serial != math.floor(serial)
    if 0 <= serial < 1 and has_time:
        return format_value(serial, "h:mm:ss AM/PM", epoch1904)[0]
    secs = parse_format(fmt)
    time_codes = any(t[0] in ("min", "elapsed", "ampm") or (t[0] == "date" and t[1][0] in "hs") for t in secs[0].tokens)
    if has_time or time_codes:
        return format_value(serial, "m/d/yyyy h:mm:ss AM/PM", epoch1904)[0]
    return format_value(serial, "m/d/yyyy", epoch1904)[0]
