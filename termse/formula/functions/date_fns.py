"""Date and time functions."""
from __future__ import annotations

import calendar
import datetime as dt
import math

from ...model import dates
from ...model.inputparse import parse_date_text
from ..errors import NUM, VALUE, XLError
from ..registry import fn, flatten, scalar_arg
from ..values import MISSING, is_number, to_int, to_number, to_text

STATE = {"epoch1904": False}


def _e():
    return STATE["epoch1904"]


def _serial(v) -> float:
    if type(v) is str:
        d = parse_date_text(v, _e())
        if d is None:
            n = to_number(v)
            return n
        return d[0]
    n = to_number(v)
    if n < 0:
        raise XLError(NUM)
    return n


def _date(v) -> dt.datetime:
    n = _serial(v)
    if n > 2958465:
        raise XLError(NUM)
    return dates.from_serial(math.floor(n), _e())


def _to_serial(d) -> float:
    return dates.to_serial(d, _e())


@fn("TODAY", volatile=True)
def TODAY():
    return _to_serial(dt.date.today())


@fn("NOW", volatile=True)
def NOW():
    return _to_serial(dt.datetime.now())


@fn("DATE")
def DATE(y, m, d):
    y, m, d = to_int(y), to_int(m), to_int(d)
    if 0 <= y < 1900:
        y += 1900
    if y < 0 or y > 9999:
        raise XLError(NUM)
    total = y * 12 + (m - 1)
    yy, mm = divmod(total, 12)
    if yy < 1 or yy > 9999:
        raise XLError(NUM)
    base = dt.datetime(yy, mm + 1, 1)
    s = _to_serial(base) + (d - 1)
    if s < 0:
        raise XLError(NUM)
    return s


@fn("TIME")
def TIME(h, m, s):
    total = to_int(h) * 3600 + to_int(m) * 60 + to_int(s)
    if total < 0:
        raise XLError(NUM)
    return (total % 86400) / 86400


def _parts(v):
    return dates.serial_date_parts(_serial(v), _e())


@fn("YEAR")
def YEAR(v):
    return _parts(v)[0]


@fn("MONTH")
def MONTH(v):
    return _parts(v)[1]


@fn("DAY")
def DAY(v):
    return _parts(v)[2]


def _secs(v) -> int:
    n = _serial(v)
    return round((n - math.floor(n)) * 86400) % 86400


@fn("HOUR")
def HOUR(v):
    return _secs(v) // 3600


@fn("MINUTE")
def MINUTE(v):
    return _secs(v) // 60 % 60


@fn("SECOND")
def SECOND(v):
    return _secs(v) % 60


@fn("WEEKDAY")
def WEEKDAY(v, kind=1):
    n = math.floor(_serial(v))
    kind = to_int(kind if kind is not None else 1)
    # serial 1 (1900-01-01) is a Sunday in Excel's calendar (because of the 1900 leap-year bug)
    sun0 = (n + 5) % 7 if _e() else (n - 1) % 7  # 0 = Sunday
    if kind == 1:
        return sun0 + 1
    if kind in (2, 11):
        return (sun0 - 1) % 7 + 1
    if kind == 3:
        return (sun0 - 1) % 7
    if 12 <= kind <= 17:
        start = kind - 10  # 12 -> Tuesday (2)
        return (sun0 - start) % 7 + 1
    raise XLError(NUM)


def _add_months(d: dt.datetime, months: int) -> dt.datetime:
    total = d.year * 12 + d.month - 1 + months
    y, m = divmod(total, 12)
    if y < 1900 or y > 9999:
        raise XLError(NUM)
    last = calendar.monthrange(y, m + 1)[1]
    return dt.datetime(y, m + 1, min(d.day, last))


@fn("EDATE")
def EDATE(start, months):
    return _to_serial(_add_months(_date(start), to_int(months)))


@fn("EOMONTH")
def EOMONTH(start, months):
    d = _add_months(_date(start).replace(day=1), to_int(months))
    return _to_serial(d.replace(day=calendar.monthrange(d.year, d.month)[1]))


@fn("DAYS")
def DAYS(end, start):
    return math.floor(_serial(end)) - math.floor(_serial(start))


@fn("DATEDIF")
def DATEDIF(start, end, unit):
    s, e = _date(start), _date(end)
    u = to_text(unit).upper()
    if s > e:
        raise XLError(NUM)
    if u == "D":
        return (e - s).days
    months = (e.year - s.year) * 12 + e.month - s.month - (1 if e.day < s.day else 0)
    if u == "M":
        return months
    if u == "Y":
        return months // 12
    if u == "YM":
        return months % 12
    if u == "MD":
        if e.day >= s.day:
            return e.day - s.day
        pm = e.replace(day=1) - dt.timedelta(days=1)
        anchor = dt.datetime(pm.year, pm.month, min(s.day, pm.day))
        return (e - anchor).days
    if u == "YD":
        def same_day_in(year):  # Feb 29 rolls over to Mar 1, like DATE()
            return dt.datetime(year, s.month, 1) + dt.timedelta(days=s.day - 1)

        ss = same_day_in(e.year)
        if ss > e:
            ss = same_day_in(e.year - 1)
        return (e - ss).days
    raise XLError(NUM)


@fn("DATEVALUE")
def DATEVALUE(text):
    d = parse_date_text(to_text(text), _e())
    if d is None:
        raise XLError(VALUE)
    return math.floor(d[0])


@fn("TIMEVALUE")
def TIMEVALUE(text):
    d = parse_date_text(to_text(text), _e())
    if d is None:
        raise XLError(VALUE)
    return d[0] - math.floor(d[0])


@fn("WEEKNUM")
def WEEKNUM(v, kind=1):
    d = _date(v)
    kind = to_int(kind if kind is not None else 1)
    if kind == 21:
        return d.isocalendar()[1]
    start = {1: 6, 17: 6, 2: 0, 11: 0, 12: 1, 13: 2, 14: 3, 15: 4, 16: 5}.get(kind)
    if start is None:
        raise XLError(NUM)
    jan1 = dt.datetime(d.year, 1, 1)
    offset = (jan1.weekday() - start) % 7
    return (d - jan1).days // 7 + (1 if ((d - jan1).days % 7) + offset < 7 else 2)


@fn("ISOWEEKNUM")
def ISOWEEKNUM(v):
    return _date(v).isocalendar()[1]


def _holidays(ctx, h):
    if h is MISSING:
        return set()
    out = set()
    for v, _ in flatten(ctx, [h]):
        if is_number(v):
            out.add(math.floor(v))
    return out


def _is_workday(serial: int, holidays: set) -> bool:
    if serial in holidays:
        return False
    wd = WEEKDAY(serial, 2)
    return wd <= 5


@fn("NETWORKDAYS", kind="raw", min_args=2, max_args=3)
def NETWORKDAYS(ctx, start, end, holidays=MISSING):
    s = math.floor(_serial(scalar_arg(ctx, start)))
    e = math.floor(_serial(scalar_arg(ctx, end)))
    hol = _holidays(ctx, holidays)
    sign = 1
    if s > e:
        s, e, sign = e, s, -1
    n = sum(1 for d in range(s, e + 1) if _is_workday(d, hol))
    return sign * n


@fn("WORKDAY", kind="raw", min_args=2, max_args=3)
def WORKDAY(ctx, start, days, holidays=MISSING):
    s = math.floor(_serial(scalar_arg(ctx, start)))
    n = to_int(scalar_arg(ctx, days))
    hol = _holidays(ctx, holidays)
    step = 1 if n >= 0 else -1
    d = s
    remaining = abs(n)
    while remaining:
        d += step
        if _is_workday(d, hol):
            remaining -= 1
    return d


@fn("YEARFRAC")
def YEARFRAC(start, end, basis=0):
    s, e = _date(start), _date(end)
    if s > e:
        s, e = e, s
    b = to_int(basis if basis is not None else 0)
    if b in (0, 4):
        d1, d2 = s.day, e.day
        if b == 0:
            if d1 == 31 or (s.month == 2 and d1 == calendar.monthrange(s.year, 2)[1]):
                d1 = 30
            if d2 == 31 and d1 == 30:
                d2 = 30
        else:
            d1, d2 = min(d1, 30), min(d2, 30)
        return ((e.year - s.year) * 360 + (e.month - s.month) * 30 + (d2 - d1)) / 360
    days = (e - s).days
    if b == 1:
        if s.year == e.year:
            ylen = 366 if calendar.isleap(s.year) else 365
        else:
            yrs = range(s.year, e.year + 1)
            ylen = sum(366 if calendar.isleap(y) else 365 for y in yrs) / len(yrs)
        return days / ylen
    if b == 2:
        return days / 360
    if b == 3:
        return days / 365
    raise XLError(NUM)
