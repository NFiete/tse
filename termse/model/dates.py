"""Excel date serial numbers (1900 and 1904 date systems)."""
from __future__ import annotations

import datetime as dt

_EPOCH_1900 = dt.datetime(1899, 12, 30)
_EPOCH_1904 = dt.datetime(1904, 1, 1)


def to_serial(value, epoch1904: bool = False) -> float:
    """datetime/date/time/timedelta -> Excel serial number."""
    if isinstance(value, dt.timedelta):
        return value.total_seconds() / 86400
    if isinstance(value, dt.time):
        return (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6) / 86400
    if isinstance(value, dt.datetime):
        d = value.replace(tzinfo=None)
    elif isinstance(value, dt.date):
        d = dt.datetime(value.year, value.month, value.day)
    else:
        raise TypeError(value)
    if epoch1904:
        delta = d - _EPOCH_1904
        serial = delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
    else:
        delta = d - _EPOCH_1900
        serial = delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
        if serial < 61:
            serial -= 1  # Excel's fictitious 29 Feb 1900
    if serial == int(serial):
        return int(serial)
    return serial


def from_serial(serial: float, epoch1904: bool = False) -> dt.datetime:
    """Excel serial -> datetime (serial 60 in the 1900 system maps to 28 Feb 1900)."""
    ms = round(serial * 86_400_000)
    days, ms_in_day = divmod(ms, 86_400_000)
    if epoch1904:
        base = _EPOCH_1904 + dt.timedelta(days=days)
    elif days >= 61:
        base = _EPOCH_1900 + dt.timedelta(days=days)
    elif days == 60:
        base = dt.datetime(1900, 2, 28)
    else:
        base = dt.datetime(1899, 12, 31) + dt.timedelta(days=days)
    return base + dt.timedelta(milliseconds=ms_in_day)


def serial_date_parts(serial: float, epoch1904: bool = False) -> tuple[int, int, int]:
    """(year, month, day) honoring the fake 1900-02-29 and day 0 (1900-01-00)."""
    days = int(serial // 1)
    if not epoch1904:
        if days == 60:
            return 1900, 2, 29
        if days == 0:
            return 1900, 1, 0
    d = from_serial(days, epoch1904)
    return d.year, d.month, d.day
