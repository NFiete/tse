"""Turn text typed into a cell into a typed value, the way Excel does."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass

from .dates import to_serial

_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
_MONTHS.update({m: i for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july", "august", "september",
     "october", "november", "december"], 1)})
_MONTHS["sept"] = 9

_TIME = r"(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp][Mm]?)?"
_RE_TIME = re.compile(rf"^{_TIME}$")
_RE_ISO = re.compile(rf"^(\d{{4}})-(\d{{1,2}})-(\d{{1,2}})(?:[ T]{_TIME})?$")
_RE_MDY = re.compile(rf"^(\d{{1,2}})/(\d{{1,2}})(?:/(\d{{2}}|\d{{4}}))?(?:\s+{_TIME})?$")
_RE_DMONY = re.compile(r"^(\d{1,2})[- ]([A-Za-z]{3,9})(?:[- ,]+(\d{2}|\d{4}))?$")
_RE_MONDY = re.compile(r"^([A-Za-z]{3,9})[- ](\d{1,2})(?:,?[- ](\d{4}))?$")


def _year(y: str | None) -> int:
    if y is None:
        return dt.date.today().year
    n = int(y)
    if len(y) == 2:
        n += 2000 if n < 30 else 1900
    return n


def _time_parts(h, m, s, ampm):
    h, m = int(h), int(m)
    sec = float(s) if s else 0.0
    if ampm:
        if h > 12:
            return None
        pm = ampm[0].lower() == "p"
        h = (h % 12) + (12 if pm else 0)
    if h > 23 or m > 59 or sec >= 60:
        return None
    return h, m, sec


def _time_fmt(s, ampm) -> str:
    base = "h:mm:ss" if s else "h:mm"
    return base + " AM/PM" if ampm else base


def parse_date_text(text: str, epoch1904: bool = False):
    """Return (serial, number_format) for date/time-looking text, else None."""
    t = text.strip()
    if not t or len(t) > 40:
        return None
    try:
        m = _RE_TIME.match(t)
        if m:
            tp = _time_parts(*m.groups())
            if tp is None:
                return None
            h, mi, sec = tp
            return (h * 3600 + mi * 60 + sec) / 86400, _time_fmt(m.group(3), m.group(4))
        m = _RE_ISO.match(t)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            date = dt.datetime(y, mo, d)
            fmt = "yyyy-mm-dd"
            if m.group(4):
                tp = _time_parts(*m.groups()[3:7])
                if tp is None:
                    return None
                date = date.replace(hour=tp[0], minute=tp[1]) + dt.timedelta(seconds=tp[2])
                fmt = "yyyy-mm-dd " + _time_fmt(m.group(6), m.group(7))
            return to_serial(date, epoch1904), fmt
        m = _RE_MDY.match(t)
        if m:
            mo, d = int(m.group(1)), int(m.group(2))
            date = dt.datetime(_year(m.group(3)), mo, d)
            fmt = "m/d/yyyy" if m.group(3) else "d-mmm"
            if m.group(4):
                tp = _time_parts(*m.groups()[3:7])
                if tp is None:
                    return None
                date = date.replace(hour=tp[0], minute=tp[1]) + dt.timedelta(seconds=tp[2])
                fmt = "m/d/yyyy h:mm"
            return to_serial(date, epoch1904), fmt
        m = _RE_DMONY.match(t)
        if m and m.group(2).lower() in _MONTHS:
            date = dt.datetime(_year(m.group(3)), _MONTHS[m.group(2).lower()], int(m.group(1)))
            return to_serial(date, epoch1904), "d-mmm-yy" if m.group(3) else "d-mmm"
        m = _RE_MONDY.match(t)
        if m and m.group(1).lower() in _MONTHS:
            date = dt.datetime(_year(m.group(3)), _MONTHS[m.group(1).lower()], int(m.group(2)))
            return to_serial(date, epoch1904), "mmm d, yyyy" if m.group(3) else "d-mmm"
    except ValueError:
        return None
    return None


_RE_NUM = re.compile(
    r"^(?P<neg>\()?(?P<sign>[+-])?(?P<cur>[$€£¥])?(?P<sign2>[+-])?"
    r"(?P<body>(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE](?P<exp>[+-]?\d+))?"
    r"(?P<pct>%)?(?P<neg2>\))?$"
)
_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


@dataclass
class Parsed:
    kind: str  # 'empty' | 'formula' | 'number' | 'text' | 'bool' | 'error'
    value: object
    number_format: str | None = None


def parse_input(text: str, epoch1904: bool = False) -> Parsed:
    if text == "":
        return Parsed("empty", None)
    if text.startswith("'"):
        return Parsed("text", text[1:])
    if text.startswith("=") and len(text) > 1:
        return Parsed("formula", text)
    s = text.strip()
    if s.startswith("+") and len(s) > 1 and not _RE_NUM.match(s):
        # Excel treats "+A1" like a formula
        return Parsed("formula", "=" + s[1:])
    u = s.upper()
    if u in ("TRUE", "FALSE"):
        return Parsed("bool", u == "TRUE")
    if u in _ERRORS:
        return Parsed("error", u)
    m = _RE_NUM.match(s)
    if m and bool(m.group("neg")) == bool(m.group("neg2")):
        body = m.group("body")
        v = float(body.replace(",", ""))
        if m.group("exp"):
            v *= 10 ** int(m.group("exp"))
        negative = bool(m.group("neg")) or m.group("sign") == "-" or m.group("sign2") == "-"
        if negative:
            v = -v
        fmt = None
        decimals = len(body.split(".")[1]) if "." in body else 0
        if m.group("pct"):
            v /= 100
            fmt = "0%" if decimals == 0 else "0." + "0" * decimals + "%"
        elif m.group("cur"):
            sym = m.group("cur")
            q = '"' + sym + '"' if sym != "$" else "$"
            fmt = f"{q}#,##0" + (".00" if decimals else "") + f"_);({q}#,##0" + (".00" if decimals else "") + ")"
        elif "," in body:
            fmt = "#,##0" + ("." + "0" * decimals if decimals else "")
        elif m.group("exp"):
            fmt = "0.00E+00"
        if v.is_integer() and abs(v) < 2**53 and not m.group("pct"):
            v = int(v)
        return Parsed("number", v, fmt)
    d = parse_date_text(s, epoch1904)
    if d is not None:
        return Parsed("number", d[0], d[1])
    return Parsed("text", text)
