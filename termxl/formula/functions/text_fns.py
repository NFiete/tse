"""Text functions."""
from __future__ import annotations

from ..errors import NA, VALUE, XLError
from ..registry import flatten, fn, scalar_arg, wildcard_regex
from ..values import MISSING, parse_number_text, to_bool, to_int, to_number, to_text


@fn("CONCATENATE", min_args=1, max_args=None)
def CONCATENATE(*args):
    return "".join(to_text(a) for a in args)


@fn("CONCAT", kind="raw", min_args=1)
def CONCAT(ctx, *args):
    out = []
    for v, _ in flatten(ctx, args):
        out.append(to_text(v))
    return _limit("".join(out))


@fn("TEXTJOIN", kind="raw", min_args=3)
def TEXTJOIN(ctx, delim, ignore_empty, *args):
    d = to_text(scalar_arg(ctx, delim))
    skip = to_bool(scalar_arg(ctx, ignore_empty))
    parts = []
    for a in args:
        if a is MISSING:
            if not skip:
                parts.append("")
            continue
        from ..values import RefValue

        if isinstance(a, RefValue):
            from ..registry import iter_ref

            vals = iter_ref(ctx, a, include_blank=True)
        else:
            vals = (v for v, _ in flatten(ctx, [a]))
        for v in vals:
            s = to_text(v)
            if s == "" and skip:
                continue
            parts.append(s)
    return _limit(d.join(parts))


def _limit(s: str) -> str:
    if len(s) > 32767:
        raise XLError(VALUE)
    return s


@fn("LEFT")
def LEFT(text, n=1):
    n = to_int(n)
    if n < 0:
        raise XLError(VALUE)
    return to_text(text)[:n]


@fn("RIGHT")
def RIGHT(text, n=1):
    n = to_int(n)
    if n < 0:
        raise XLError(VALUE)
    s = to_text(text)
    return s[len(s) - n:] if n else ""


@fn("MID")
def MID(text, start, n):
    start, n = to_int(start), to_int(n)
    if start < 1 or n < 0:
        raise XLError(VALUE)
    return to_text(text)[start - 1:start - 1 + n]


@fn("LEN")
def LEN(text):
    return len(to_text(text))


@fn("UPPER")
def UPPER(text):
    return to_text(text).upper()


@fn("LOWER")
def LOWER(text):
    return to_text(text).lower()


@fn("PROPER")
def PROPER(text):
    out, prev_alpha = [], False
    for ch in to_text(text):
        out.append(ch.lower() if prev_alpha else ch.upper())
        prev_alpha = ch.isalpha()
    return "".join(out)


@fn("TRIM")
def TRIM(text):
    return " ".join(p for p in to_text(text).split(" ") if p)


@fn("CLEAN")
def CLEAN(text):
    return "".join(ch for ch in to_text(text) if ord(ch) >= 32)


@fn("SUBSTITUTE")
def SUBSTITUTE(text, old, new, instance=None):
    s, old, new = to_text(text), to_text(old), to_text(new)
    if old == "":
        return s
    if instance is None:
        return s.replace(old, new)
    n = to_int(instance)
    if n < 1:
        raise XLError(VALUE)
    idx = -1
    for _ in range(n):
        idx = s.find(old, idx + 1)
        if idx < 0:
            return s
    return s[:idx] + new + s[idx + len(old):]


@fn("REPLACE")
def REPLACE(text, start, n, new):
    s = to_text(text)
    start, n = to_int(start), to_int(n)
    if start < 1 or n < 0:
        raise XLError(VALUE)
    return s[:start - 1] + to_text(new) + s[start - 1 + n:]


@fn("FIND")
def FIND(needle, hay, start=1):
    needle, hay = to_text(needle), to_text(hay)
    start = to_int(start if start is not None else 1)
    if start < 1 or start > len(hay) + 1:
        raise XLError(VALUE)
    i = hay.find(needle, start - 1)
    if i < 0:
        raise XLError(VALUE)
    return i + 1


@fn("SEARCH")
def SEARCH(needle, hay, start=1):
    needle, hay = to_text(needle), to_text(hay)
    start = to_int(start if start is not None else 1)
    if start < 1 or start > len(hay) + 1:
        raise XLError(VALUE)
    if any(c in needle for c in "*?~"):
        rx = wildcard_regex(needle)
        m = rx.search(hay, start - 1)
        if not m:
            raise XLError(VALUE)
        return m.start() + 1
    i = hay.lower().find(needle.lower(), start - 1)
    if i < 0:
        raise XLError(VALUE)
    return i + 1


@fn("TEXT")
def TEXT(value, fmt):
    from ...format.numfmt import format_value

    fmt = to_text(fmt)
    if type(value) is str:
        n = parse_number_text(value)
        if n is not None:
            value = n
    text, _ = format_value(value, fmt)
    return text


@fn("VALUE")
def VALUE_(text):
    if type(text) in (int, float):
        return text
    if text is None:
        return 0
    n = parse_number_text(to_text(text))
    if n is None:
        raise XLError(VALUE)
    return n


@fn("NUMBERVALUE")
def NUMBERVALUE(text, dec=".", grp=","):
    s = to_text(text).replace(to_text(grp) if grp is not None else ",", "")
    if dec is not None and to_text(dec) != ".":
        s = s.replace(to_text(dec), ".")
    n = parse_number_text(s.replace(" ", ""))
    if n is None:
        raise XLError(VALUE)
    return n


@fn("REPT")
def REPT(text, n):
    n = to_int(n)
    if n < 0:
        raise XLError(VALUE)
    return _limit(to_text(text) * n)


@fn("EXACT")
def EXACT(a, b):
    return to_text(a) == to_text(b)


@fn("CHAR")
def CHAR(n):
    n = to_int(n)
    if n < 1 or n > 255:
        raise XLError(VALUE)
    return bytes([n]).decode("cp1252", errors="replace")


@fn("UNICHAR")
def UNICHAR(n):
    n = to_int(n)
    if n < 1:
        raise XLError(VALUE)
    return chr(n)


@fn("CODE")
def CODE(text):
    s = to_text(text)
    if not s:
        raise XLError(VALUE)
    try:
        return s[0].encode("cp1252")[0]
    except UnicodeEncodeError:
        return 63


@fn("UNICODE")
def UNICODE(text):
    s = to_text(text)
    if not s:
        raise XLError(VALUE)
    return ord(s[0])


@fn("FIXED")
def FIXED(x, decimals=2, no_commas=False):
    from .math_fns import ROUND

    d = to_int(decimals if decimals is not None else 2)
    v = ROUND(to_number(x), d)
    commas = not to_bool(no_commas if no_commas is not None else False)
    fmt = f"{{:{',' if commas else ''}.{max(d, 0)}f}}"
    return fmt.format(v)


@fn("DOLLAR")
def DOLLAR(x, decimals=2):
    d = to_int(decimals if decimals is not None else 2)
    v = to_number(x)
    s = f"${abs(v):,.{max(d, 0)}f}"
    return f"({s})" if v < 0 else s


@fn("TEXTBEFORE")
def TEXTBEFORE(text, delim, instance=1):
    s, d = to_text(text), to_text(delim)
    n = to_int(instance if instance is not None else 1)
    parts = s.split(d)
    if n == 0 or abs(n) >= len(parts):
        raise XLError(NA)
    return d.join(parts[:n]) if n > 0 else d.join(parts[:n])


@fn("TEXTAFTER")
def TEXTAFTER(text, delim, instance=1):
    s, d = to_text(text), to_text(delim)
    n = to_int(instance if instance is not None else 1)
    parts = s.split(d)
    if n == 0 or abs(n) >= len(parts):
        raise XLError(NA)
    return d.join(parts[n:]) if n > 0 else d.join(parts[n:])

