"""Lookup and reference functions."""
from __future__ import annotations

from ...model.refs import MAX_COL, MAX_ROW, col_letter, quote_sheet
from ..errors import CALC, NA, NUM, REF, VALUE, ExcelError, XLError
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


# ------------------------------------------------------------------ array shaping


def full_array(ctx, v) -> Array:
    """Like to_array, but a reference keeps its full size: blank rows and columns beyond the
    sheet's used area are included (as shared blank rows, so whole columns stay cheap)."""
    a = to_array(ctx, v)
    if not isinstance(v, RefValue):
        return a
    rows = a.rows
    if a.ncols < v.ncols:
        extra = [None] * (v.ncols - a.ncols)
        rows = [r + extra for r in rows]
    if len(rows) < v.nrows:
        rows = rows + [[None] * v.ncols] * (v.nrows - len(rows))
    return Array(rows) if rows is not a.rows else a


def _opt_int(ctx, v):
    return None if v is MISSING else to_int(scalar_arg(ctx, v))


def _index_list(ctx, args) -> list[int]:
    """Row/column numbers given as scalars or arrays, e.g. CHOOSEROWS(A1:C9,1,{3,4})."""
    out = []
    for a in args:
        if isinstance(a, (RefValue, Array)):
            for v in to_array(ctx, a).flat():
                if isinstance(v, ExcelError):
                    raise XLError(v)
                out.append(to_int(v))
        else:
            out.append(to_int(scalar_arg(ctx, a)))
    return out


def _resolve_indexes(idx: list[int], size: int) -> list[int]:
    out = []
    for i in idx:
        if 0 < i <= size:
            out.append(i - 1)
        elif 0 < -i <= size:
            out.append(size + i)
        else:
            raise XLError(VALUE)
    return out


@fn("CHOOSEROWS", kind="raw", min_args=2)
def CHOOSEROWS(ctx, arr, *nums):
    a = full_array(ctx, arr)
    return Array([list(a.rows[i]) for i in _resolve_indexes(_index_list(ctx, nums), a.nrows)])


@fn("CHOOSECOLS", kind="raw", min_args=2)
def CHOOSECOLS(ctx, arr, *nums):
    a = full_array(ctx, arr)
    idx = _resolve_indexes(_index_list(ctx, nums), a.ncols)
    return Array([[row[j] for j in idx] for row in a.rows])


def _take_span(n, size):
    if n is None:
        return 0, size
    if n == 0:
        raise XLError(CALC)
    return (0, min(n, size)) if n > 0 else (max(size + n, 0), size)


def _drop_span(n, size):
    if not n:
        return 0, size
    if abs(n) >= size:
        raise XLError(CALC)
    return (n, size) if n > 0 else (0, size + n)


@fn("TAKE", kind="raw", min_args=2, max_args=3)
def TAKE(ctx, arr, rows, cols=MISSING):
    a = full_array(ctx, arr)
    r1, r2 = _take_span(_opt_int(ctx, rows), a.nrows)
    c1, c2 = _take_span(_opt_int(ctx, cols), a.ncols)
    return _slice(a, r1, r2, c1, c2)


def _slice(a: Array, r1, r2, c1, c2) -> Array:
    if c1 == 0 and c2 == a.ncols:
        return Array(a.rows[r1:r2])  # rows are never mutated, so they can be shared
    return Array([row[c1:c2] for row in a.rows[r1:r2]])


@fn("DROP", kind="raw", min_args=2, max_args=3)
def DROP(ctx, arr, rows, cols=MISSING):
    a = full_array(ctx, arr)
    r1, r2 = _drop_span(_opt_int(ctx, rows), a.nrows)
    c1, c2 = _drop_span(_opt_int(ctx, cols), a.ncols)
    return _slice(a, r1, r2, c1, c2)


@fn("EXPAND", kind="raw", min_args=2, max_args=4)
def EXPAND(ctx, arr, rows, cols=MISSING, pad=MISSING):
    a = full_array(ctx, arr)
    nr = _opt_int(ctx, rows)
    nc = _opt_int(ctx, cols)
    nr = a.nrows if nr is None else nr
    nc = a.ncols if nc is None else nc
    if nr < a.nrows or nc < a.ncols:
        raise XLError(VALUE)
    p = NA if pad is MISSING else to_scalar(ctx, pad)
    out = [list(row) + [p] * (nc - a.ncols) for row in a.rows]
    out += [[p] * nc for _ in range(nr - a.nrows)]
    return Array(out)


def _stack_parts(ctx, arrays):
    return [full_array(ctx, a) for a in arrays]


@fn("VSTACK", kind="raw", min_args=1)
def VSTACK(ctx, *arrays):
    parts = _stack_parts(ctx, arrays)
    width = max(p.ncols for p in parts)
    out = []
    for p in parts:
        if p.ncols == width:
            out.extend(p.rows)
        else:
            fill = [NA] * (width - p.ncols)
            out.extend(row + fill for row in p.rows)
    return Array(out)


@fn("HSTACK", kind="raw", min_args=1)
def HSTACK(ctx, *arrays):
    parts = _stack_parts(ctx, arrays)
    height = max(p.nrows for p in parts)
    out = [[] for _ in range(height)]
    for p in parts:
        for i in range(height):
            out[i].extend(p.rows[i] if i < p.nrows else [NA] * p.ncols)
    return Array(out)


def _flat_ignoring(ctx, arr, ignore, by_col) -> list:
    a = full_array(ctx, arr)
    ig = 0 if ignore is MISSING else to_int(scalar_arg(ctx, ignore))
    if ig not in (0, 1, 2, 3):
        raise XLError(VALUE)
    col = False if by_col is MISSING else to_bool(scalar_arg(ctx, by_col))
    vals = (v for c in zip(*a.rows) for v in c) if col else a.flat()
    out = [v for v in vals
           if not ((ig & 1 and v is None) or (ig & 2 and isinstance(v, ExcelError)))]
    if not out:
        raise XLError(CALC)
    return out


@fn("TOCOL", kind="raw", min_args=1, max_args=3)
def TOCOL(ctx, arr, ignore=MISSING, by_col=MISSING):
    return Array([[v] for v in _flat_ignoring(ctx, arr, ignore, by_col)])


@fn("TOROW", kind="raw", min_args=1, max_args=3)
def TOROW(ctx, arr, ignore=MISSING, by_col=MISSING):
    return Array([_flat_ignoring(ctx, arr, ignore, by_col)])


def _wrap(ctx, vec, count, pad) -> list[list]:
    a = full_array(ctx, vec)
    if a.nrows != 1 and a.ncols != 1:
        raise XLError(VALUE)
    n = to_int(scalar_arg(ctx, count))
    if n < 1:
        raise XLError(NUM)
    vals = list(a.flat())
    p = NA if pad is MISSING else to_scalar(ctx, pad)
    chunks = [vals[i:i + n] for i in range(0, len(vals), n)]
    chunks[-1] += [p] * (n - len(chunks[-1]))
    return chunks


@fn("WRAPROWS", kind="raw", min_args=2, max_args=3)
def WRAPROWS(ctx, vec, count, pad=MISSING):
    return Array(_wrap(ctx, vec, count, pad))


@fn("WRAPCOLS", kind="raw", min_args=2, max_args=3)
def WRAPCOLS(ctx, vec, count, pad=MISSING):
    return Array([list(r) for r in zip(*_wrap(ctx, vec, count, pad))])


def sort_cmp(a, b, desc=False) -> int:
    """Order for SORT/SORTBY keys: numbers, text, logicals, errors; blanks always last."""
    if a is None or b is None:
        return (a is None) - (b is None)
    ea, eb = isinstance(a, ExcelError), isinstance(b, ExcelError)
    if ea or eb:
        c = ea - eb
    else:
        c = compare(a, b)
    return -c if desc else c


@fn("SORTBY", kind="raw", min_args=2)
def SORTBY(ctx, arr, *pairs):
    import functools

    a = full_array(ctx, arr)
    keys, by_cols = [], set()
    for i in range(0, len(pairs), 2):
        raw = pairs[i]
        by = to_array(ctx, raw)  # clipped to the used area: positions past it are blank
        nr, nc = (raw.nrows, raw.ncols) if isinstance(raw, RefValue) else (by.nrows, by.ncols)
        order = pairs[i + 1] if i + 1 < len(pairs) else MISSING
        o = 1 if order is MISSING else to_number(scalar_arg(ctx, order))
        if o not in (1, -1):
            raise XLError(VALUE)
        if nc == 1 and nr == a.nrows:
            by_cols.add(False)
            keys.append(([r[0] for r in by.rows], o < 0))
        elif nr == 1 and nc == a.ncols:
            by_cols.add(True)
            keys.append((list(by.rows[0]), o < 0))
        else:
            raise XLError(VALUE)
    single = a.nrows == 1 and a.ncols == 1
    if len(by_cols) > 1 and not single:
        raise XLError(VALUE)
    by_col = False if single else by_cols.pop()
    n = a.ncols if by_col else a.nrows
    used = min(n, max(len(vals) for vals, _ in keys))
    for vals, _ in keys:
        vals.extend([None] * (used - len(vals)))

    def cmp(i, j):
        for vals, desc in keys:
            c = sort_cmp(vals[i], vals[j], desc)
            if c:
                return c
        return 0

    # Positions whose keys are all blank go last in their original order; not sorting them
    # keeps whole-column references (mostly blank) fast.
    blank = [i for i in range(used) if all(vals[i] is None for vals, _ in keys)]
    if blank:
        bset = set(blank)
        rest = [i for i in range(used) if i not in bset]
    else:
        rest = range(used)
    order = sorted(rest, key=functools.cmp_to_key(cmp)) + blank
    if by_col:
        return Array([[row[j] for j in order] + row[used:] for row in a.rows])
    return Array([a.rows[i] for i in order] + a.rows[used:])


# ------------------------------------------------------------------ reference information


@fn("FORMULATEXT", kind="raw", min_args=1, max_args=1)
def FORMULATEXT(ctx, ref):
    if isinstance(ref, ExcelError):
        raise XLError(ref)
    if not isinstance(ref, RefValue):
        raise XLError(VALUE)
    text = ctx.engine.formula_text(ref.sheet, ref.r1, ref.c1)
    if text is None:
        raise XLError(NA)
    return text


@fn("HYPERLINK", kind="raw", min_args=1, max_args=2)
def HYPERLINK(ctx, link, name=MISSING):
    v = to_scalar(ctx, link if name is MISSING else name)
    return None if v is MISSING else v


@fn("AREAS", kind="raw", min_args=1, max_args=1)
def AREAS(ctx, ref):
    if isinstance(ref, ExcelError):
        raise XLError(ref)
    if not isinstance(ref, RefValue):
        raise XLError(VALUE)
    return 1


def _trim_span(flags: list[bool], mode: int):
    """First and last index to keep, given which positions hold a value."""
    lo, hi = 0, len(flags) - 1
    if mode & 1:
        while lo <= hi and not flags[lo]:
            lo += 1
    if mode & 2:
        while hi >= lo and not flags[hi]:
            hi -= 1
    return lo, hi


@fn("TRIMRANGE", kind="raw", min_args=1, max_args=3)
def TRIMRANGE(ctx, rng, trim_rows=MISSING, trim_cols=MISSING):
    tr = 3 if trim_rows is MISSING else to_int(scalar_arg(ctx, trim_rows))
    tc = 3 if trim_cols is MISSING else to_int(scalar_arg(ctx, trim_cols))
    if tr not in (0, 1, 2, 3) or tc not in (0, 1, 2, 3):
        raise XLError(VALUE)
    if isinstance(rng, ExcelError):
        raise XLError(rng)
    if isinstance(rng, RefValue):
        # Rows and columns past the used area are blank; only scan the part inside it.
        mr, mc = ctx.engine.bounds(rng.sheet)
        r2, c2 = min(rng.r2, mr), min(rng.c2, mc)
        rows = ctx.engine.range_rows(rng.sheet, rng.r1, rng.c1, r2, c2) if r2 >= rng.r1 and c2 >= rng.c1 else []
        row_has = [any(v is not None for v in row) for row in rows]
        row_has += [False] * (rng.nrows - len(row_has))
        col_has = [any(row[j] is not None for row in rows) for j in range(c2 - rng.c1 + 1)] if rows else []
        col_has += [False] * (rng.ncols - len(col_has))
        if not any(row_has):
            return RefValue(rng.sheet, rng.r1, rng.c1, rng.r1, rng.c1)
        rlo, rhi = _trim_span(row_has, tr)
        clo, chi = _trim_span(col_has, tc)
        return RefValue(rng.sheet, rng.r1 + rlo, rng.c1 + clo, rng.r1 + rhi, rng.c1 + chi)
    a = to_array(ctx, rng)
    row_has = [any(v is not None for v in row) for row in a.rows]
    if not any(row_has):
        raise XLError(CALC)
    col_has = [any(row[j] is not None for row in a.rows) for j in range(a.ncols)]
    rlo, rhi = _trim_span(row_has, tr)
    clo, chi = _trim_span(col_has, tc)
    return Array([row[clo:chi + 1] for row in a.rows[rlo:rhi + 1]])


# ------------------------------------------------------------------ spilling


@fn("SINGLE", kind="raw", min_args=1, max_args=1)
def SINGLE(ctx, v):
    """The @ operator: implicit intersection of a range with the formula's row or column,
    or the first value of an array."""
    if isinstance(v, RefValue):
        if v.is_cell:
            return v
        from ..evaluator import intersect

        res = intersect(ctx, v)
        if isinstance(res, ExcelError):
            raise XLError(res)
        return res
    if isinstance(v, Array):
        if not v.rows or not v.rows[0]:
            raise XLError(VALUE)
        return v.rows[0][0]
    return v


@fn("ANCHORARRAY", kind="raw", min_args=1, max_args=1)
def ANCHORARRAY(ctx, v):
    """A1#: the whole range spilled by the formula in A1."""
    if isinstance(v, ExcelError):
        raise XLError(v)
    if not isinstance(v, RefValue):
        raise XLError(REF)
    return ctx.engine.spill_ref(v.sheet, v.r1, v.c1)
