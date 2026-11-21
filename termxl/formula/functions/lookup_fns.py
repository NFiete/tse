"""Lookup and reference functions."""
from __future__ import annotations

from ...model.refs import MAX_COL, MAX_ROW, col_letter, quote_sheet
from ..errors import CALC, NA, REF, VALUE, ExcelError, XLError
from ..registry import fn, lookup_equal, scalar_arg, to_array, to_scalar, wild_matcher
from ..values import MISSING, Array, RefValue, compare, to_bool, to_int, to_number, to_text, type_rank


def _approx_index(keys, x, descending=False):
    """Index of the last key <= x (or for descending, last key >= x) among comparable keys."""
    best = None
    rx = type_rank(x)
    for i, k in enumerate(keys):
        if k is None or isinstance(k, ExcelError) or type_rank(k) != rx:
            continue
        c = compare(k, x)
        if (c <= 0) if not descending else (c >= 0):
            best = i
        else:
            break
    return best


def _exact_index(keys, x, reverse=False):
    match = wild_matcher(x)
    rng = range(len(keys) - 1, -1, -1) if reverse else range(len(keys))
    for i in rng:
        if match(keys[i]):
            return i
    return None


def _lookup_value(ctx, x):
    x = scalar_arg(ctx, x)
    if x is None:
        x = 0
    return x


def _table_shape(table, arr: Array):
    if isinstance(table, RefValue):
        return table.nrows, table.ncols
    return arr.nrows, arr.ncols


def _arr_get(arr: Array, i, j):
    if i < arr.nrows and j < arr.ncols:
        return arr.rows[i][j]
    return None


@fn("VLOOKUP", kind="raw", min_args=3, max_args=4)
def VLOOKUP(ctx, x, table, col, approx=MISSING):
    x = _lookup_value(ctx, x)
    arr = to_array(ctx, table)
    _, ncols = _table_shape(table, arr)
    col = to_int(scalar_arg(ctx, col))
    if col < 1:
        raise XLError(VALUE)
    if col > ncols:
        raise XLError(REF)
    approx = True if approx is MISSING else to_bool(scalar_arg(ctx, approx) or False)
    keys = [r[0] if r else None for r in arr.rows]
    i = _approx_index(keys, x) if approx else _exact_index(keys, x)
    if i is None:
        raise XLError(NA)
    v = _arr_get(arr, i, col - 1)
    return 0 if v is None else v


@fn("HLOOKUP", kind="raw", min_args=3, max_args=4)
def HLOOKUP(ctx, x, table, row, approx=MISSING):
    x = _lookup_value(ctx, x)
    arr = to_array(ctx, table)
    nrows, _ = _table_shape(table, arr)
    row = to_int(scalar_arg(ctx, row))
    if row < 1:
        raise XLError(VALUE)
    if row > nrows:
        raise XLError(REF)
    approx = True if approx is MISSING else to_bool(scalar_arg(ctx, approx) or False)
    keys = arr.rows[0] if arr.rows else []
    i = _approx_index(keys, x) if approx else _exact_index(keys, x)
    if i is None:
        raise XLError(NA)
    v = _arr_get(arr, row - 1, i)
    return 0 if v is None else v


def _vector(ctx, v) -> list:
    arr = to_array(ctx, v)
    if arr.nrows == 1:
        return list(arr.rows[0])
    if arr.ncols == 1:
        return [r[0] for r in arr.rows]
    raise XLError(NA)


@fn("MATCH", kind="raw", min_args=2, max_args=3)
def MATCH(ctx, x, arr, mtype=MISSING):
    x = _lookup_value(ctx, x)
    keys = _vector(ctx, arr)
    mt = 1 if mtype is MISSING else to_int(scalar_arg(ctx, mtype) or 0)
    if mt == 0:
        i = _exact_index(keys, x)
    elif mt > 0:
        i = _approx_index(keys, x)
    else:
        i = _approx_index(keys, x, descending=True)
    if i is None:
        raise XLError(NA)
    return i + 1


def _xmatch_index(keys, x, match_mode, search_mode):
    reverse = search_mode in (-1, -2)
    order = range(len(keys) - 1, -1, -1) if reverse else range(len(keys))
    if match_mode == 2:
        return _exact_index(keys, x, reverse)
    eq = None
    best, best_v = None, None
    for i in order:
        k = keys[i]
        if k is None or isinstance(k, ExcelError):
            continue
        if lookup_equal(x, k):
            eq = i
            break
        if match_mode in (-1, 1) and type_rank(k) == type_rank(x):
            c = compare(k, x)
            if match_mode == -1 and c < 0 and (best is None or compare(k, best_v) > 0):
                best, best_v = i, k
            elif match_mode == 1 and c > 0 and (best is None or compare(k, best_v) < 0):
                best, best_v = i, k
    if eq is not None:
        return eq
    return best if match_mode in (-1, 1) else None


@fn("XLOOKUP", kind="raw", min_args=3, max_args=6)
def XLOOKUP(ctx, x, look, ret, not_found=MISSING, match_mode=MISSING, search_mode=MISSING):
    x = _lookup_value(ctx, x)
    larr = to_array(ctx, look)
    vertical = larr.ncols == 1 and not (isinstance(look, RefValue) and look.nrows == 1 and look.ncols > 1)
    keys = [r[0] for r in larr.rows] if vertical else (larr.rows[0] if larr.rows else [])
    mm = 0 if match_mode is MISSING else to_int(scalar_arg(ctx, match_mode) or 0)
    sm = 1 if search_mode is MISSING else to_int(scalar_arg(ctx, search_mode) or 1)
    i = _xmatch_index(keys, x, mm, sm)
    if i is None:
        if not_found is not MISSING:
            return not_found
        raise XLError(NA)
    if isinstance(ret, RefValue):
        if vertical:
            return RefValue(ret.sheet, ret.r1 + i, ret.c1, ret.r1 + i, ret.c2)
        return RefValue(ret.sheet, ret.r1, ret.c1 + i, ret.r2, ret.c1 + i)
    rarr = to_array(ctx, ret)
    if vertical:
        if i >= rarr.nrows:
            raise XLError(VALUE)
        return Array([rarr.rows[i]]) if rarr.ncols > 1 else rarr.rows[i][0]
    if i >= rarr.ncols:
        raise XLError(VALUE)
    return Array([[r[i]] for r in rarr.rows]) if rarr.nrows > 1 else rarr.rows[0][i]


@fn("XMATCH", kind="raw", min_args=2, max_args=4)
def XMATCH(ctx, x, arr, match_mode=MISSING, search_mode=MISSING):
    x = _lookup_value(ctx, x)
    keys = _vector(ctx, arr)
    mm = 0 if match_mode is MISSING else to_int(scalar_arg(ctx, match_mode) or 0)
    sm = 1 if search_mode is MISSING else to_int(scalar_arg(ctx, search_mode) or 1)
    i = _xmatch_index(keys, x, mm, sm)
    if i is None:
        raise XLError(NA)
    return i + 1


@fn("LOOKUP", kind="raw", min_args=2, max_args=3)
def LOOKUP(ctx, x, look, result=MISSING):
    x = _lookup_value(ctx, x)
    larr = to_array(ctx, look)
    if result is MISSING:
        if larr.ncols > larr.nrows:
            keys, vals = larr.rows[0], larr.rows[-1]
        else:
            keys, vals = [r[0] for r in larr.rows], [r[-1] for r in larr.rows]
    else:
        keys = _vector(ctx, look)
        vals = _vector(ctx, result)
    i = _approx_index(keys, x)
    if i is None or i >= len(vals):
        raise XLError(NA)
    return 0 if vals[i] is None else vals[i]


@fn("INDEX", kind="raw", min_args=2, max_args=4)
def INDEX(ctx, ref, row, col=MISSING, area=MISSING):
    r = to_int(scalar_arg(ctx, row) or 0)
    c = None if col is MISSING else to_int(scalar_arg(ctx, col) or 0)
    if isinstance(ref, RefValue):
        nr, nc = ref.nrows, ref.ncols
    else:
        ref = to_array(ctx, ref)
        nr, nc = ref.nrows, ref.ncols
    if c is None:
        if nr == 1 and nc > 1:
            r, c = 1, r
        else:
            c = 1 if nc == 1 else 0
    if r < 0 or c < 0 or r > nr or c > nc:
        raise XLError(REF)
    if isinstance(ref, RefValue):
        r1, r2 = (ref.r1, ref.r2) if r == 0 else (ref.r1 + r - 1,) * 2
        c1, c2 = (ref.c1, ref.c2) if c == 0 else (ref.c1 + c - 1,) * 2
        return RefValue(ref.sheet, r1, c1, r2, c2)
    if r == 0 and c == 0:
        return ref
    if r == 0:
        return Array([[row_[c - 1]] for row_ in ref.rows])
    if c == 0:
        return Array([ref.rows[r - 1]])
    return ref.rows[r - 1][c - 1]


@fn("CHOOSECOLS", kind="raw", min_args=2)
def CHOOSECOLS(ctx, arr, *cols):
    a = to_array(ctx, arr)
    idx = [to_int(scalar_arg(ctx, c)) for c in cols]
    out = []
    for row in a.rows:
        out.append([row[i - 1] if i > 0 else row[i] for i in idx])
    return Array(out)


@fn("ROW", kind="raw", max_args=1)
def ROW(ctx, ref=MISSING):
    if ref is MISSING:
        return ctx.row
    if isinstance(ref, RefValue):
        return ref.r1
    raise XLError(VALUE)


@fn("COLUMN", kind="raw", max_args=1)
def COLUMN(ctx, ref=MISSING):
    if ref is MISSING:
        return ctx.col
    if isinstance(ref, RefValue):
        return ref.c1
    raise XLError(VALUE)


@fn("ROWS", kind="raw", min_args=1, max_args=1)
def ROWS(ctx, ref):
    if isinstance(ref, (RefValue, Array)):
        return ref.nrows
    return 1


@fn("COLUMNS", kind="raw", min_args=1, max_args=1)
def COLUMNS(ctx, ref):
    if isinstance(ref, (RefValue, Array)):
        return ref.ncols
    return 1


@fn("OFFSET", kind="raw", min_args=3, max_args=5, volatile=True)
def OFFSET(ctx, ref, rows, cols, height=MISSING, width=MISSING):
    if not isinstance(ref, RefValue):
        raise XLError(VALUE)
    dr = to_int(scalar_arg(ctx, rows) or 0)
    dc = to_int(scalar_arg(ctx, cols) or 0)
    h = ref.nrows if height is MISSING else to_int(scalar_arg(ctx, height))
    w = ref.ncols if width is MISSING else to_int(scalar_arg(ctx, width))
    if h == 0 or w == 0:
        raise XLError(REF)
    r1, c1 = ref.r1 + dr, ref.c1 + dc
    r2, c2 = r1 + h - 1 if h > 0 else r1, c1 + w - 1 if w > 0 else c1
    if h < 0:
        r1, r2 = r1 + h + 1, r1
    if w < 0:
        c1, c2 = c1 + w + 1, c1
    if r1 < 1 or c1 < 1 or r2 > MAX_ROW or c2 > MAX_COL:
        raise XLError(REF)
    return RefValue(ref.sheet, r1, c1, r2, c2)


@fn("INDIRECT", kind="raw", min_args=1, max_args=2, volatile=True)
def INDIRECT(ctx, text, a1=MISSING):
    from ..parser import FormulaSyntaxError, RefNode, Unsupported, parse
    from ..evaluator import evaluate

    s = to_text(scalar_arg(ctx, text)).strip()
    if not s:
        raise XLError(REF)
    use_a1 = True if a1 is MISSING else to_bool(scalar_arg(ctx, a1))
    if not use_a1:
        s = _r1c1_to_a1(s)
    try:
        node = parse(s)
    except (FormulaSyntaxError, Unsupported):
        raise XLError(REF)
    from ..parser import Binary, NameNode

    if not isinstance(node, (RefNode, NameNode)) and not (isinstance(node, Binary) and node.op == ":"):
        raise XLError(REF)
    v = evaluate(node, ctx)
    if not isinstance(v, RefValue):
        raise XLError(REF)
    return v


def _r1c1_to_a1(s: str) -> str:
    import re

    sheet = ""
    if "!" in s:
        sheet, s = s.rsplit("!", 1)
        sheet += "!"
    parts = []
    for p in s.split(":"):
        m = re.fullmatch(r"[Rr](\d+)[Cc](\d+)", p.strip())
        if not m:
            raise XLError(REF)
        parts.append(f"{col_letter(int(m.group(2)))}{int(m.group(1))}")
    return sheet + ":".join(parts)


@fn("ADDRESS", kind="raw", min_args=2, max_args=5)
def ADDRESS(ctx, row, col, abs_num=MISSING, a1=MISSING, sheet=MISSING):
    r = to_int(scalar_arg(ctx, row))
    c = to_int(scalar_arg(ctx, col))
    if r < 1 or c < 1 or r > MAX_ROW or c > MAX_COL:
        raise XLError(VALUE)
    a = 1 if abs_num is MISSING else to_int(scalar_arg(ctx, abs_num))
    use_a1 = True if a1 is MISSING else to_bool(scalar_arg(ctx, a1))
    prefix = ""
    if sheet is not MISSING:
        prefix = quote_sheet(to_text(scalar_arg(ctx, sheet))) + "!"
    if use_a1:
        cr = "$" if a in (1, 3) else ""
        rr = "$" if a in (1, 2) else ""
        return f"{prefix}{cr}{col_letter(c)}{rr}{r}"
    rs = f"R{r}" if a in (1, 2) else f"R[{r}]"
    cs = f"C{c}" if a in (1, 3) else f"C[{c}]"
    return prefix + rs + cs


@fn("TRANSPOSE", kind="raw", min_args=1, max_args=1)
def TRANSPOSE(ctx, arr):
    a = to_array(ctx, arr)
    return Array([list(col) for col in zip(*a.rows)])


# ------------------------------------------------------------------ dynamic-array helpers
# Spilling is not supported, but these work when nested inside aggregates (SUM(FILTER(...))).


@fn("FILTER", kind="raw", min_args=2, max_args=3)
def FILTER(ctx, arr, include, if_empty=MISSING):
    a = to_array(ctx, arr)
    inc = to_array(ctx, include)
    rows = []
    if inc.ncols == 1 and inc.nrows == a.nrows:
        for row, flag in zip(a.rows, inc.rows):
            f = flag[0]
            if isinstance(f, ExcelError):
                raise XLError(f)
            if to_bool(f if f is not None else False):
                rows.append(row)
    elif inc.nrows == 1 and inc.ncols == a.ncols:
        keep = [j for j, f in enumerate(inc.rows[0]) if to_bool(f if f is not None else False)]
        rows = [[row[j] for j in keep] for row in a.rows] if keep else []
    else:
        raise XLError(VALUE)
    if not rows or not rows[0]:
        if if_empty is not MISSING:
            return to_scalar(ctx, if_empty)
        raise XLError(CALC)
    return Array(rows)


@fn("UNIQUE", kind="raw", min_args=1, max_args=3)
def UNIQUE(ctx, arr, by_col=MISSING, exactly_once=MISSING):
    a = to_array(ctx, arr)
    once = False if exactly_once is MISSING else to_bool(scalar_arg(ctx, exactly_once))

    def key(row):
        return tuple((v.lower() if isinstance(v, str) else v) for v in row)

    counts: dict = {}
    order = []
    for row in a.rows:
        k = key(row)
        if k not in counts:
            counts[k] = 0
            order.append(row)
        counts[k] += 1
    rows = [r for r in order if not once or counts[key(r)] == 1]
    if not rows:
        raise XLError(CALC)
    return Array(rows)


@fn("SORT", kind="raw", min_args=1, max_args=4)
def SORT(ctx, arr, idx=MISSING, order=MISSING, by_col=MISSING):
    import functools

    a = to_array(ctx, arr)
    i = 1 if idx is MISSING else to_int(scalar_arg(ctx, idx))
    desc = False if order is MISSING else to_number(scalar_arg(ctx, order)) < 0

    def cmp(x, y):
        return compare(x[i - 1], y[i - 1])

    rows = sorted(a.rows, key=functools.cmp_to_key(cmp), reverse=desc)
    return Array(rows)


@fn("SEQUENCE", kind="raw", min_args=1, max_args=4)
def SEQUENCE(ctx, rows, cols=MISSING, start=MISSING, step=MISSING):
    r = to_int(scalar_arg(ctx, rows))
    c = 1 if cols is MISSING else to_int(scalar_arg(ctx, cols))
    s = 1 if start is MISSING else to_number(scalar_arg(ctx, start))
    st = 1 if step is MISSING else to_number(scalar_arg(ctx, step))
    if r < 1 or c < 1:
        raise XLError(CALC)
    return Array([[s + (i * c + j) * st for j in range(c)] for i in range(r)])
