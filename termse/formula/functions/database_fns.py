"""Database functions (DSUM, DCOUNT, DGET...): aggregate one field of the records that match a
criteria range, using the same criteria rules as Excel's Advanced Filter."""
from __future__ import annotations

import math

from ..errors import DIV0, NUM, VALUE, ExcelError, XLError
from ..registry import fn, make_criterion, scalar_arg, to_array
from ..values import MISSING, RefValue, is_number, parse_number_text, to_int
from .math_fns import _sum
from .stats_fns import _var


def _key(v):
    """Labels match case-insensitively; numeric labels match equal numbers."""
    if type(v) is str:
        return v.strip().lower()
    return v


def _criterion(v):
    """A criteria cell's predicate. Plain text without an operator matches text that starts with
    it ("Dav" finds "Davolio"); use ="=Dav" for an exact match."""
    if type(v) is str and v and v[0] not in "<>=" and parse_number_text(v) is None \
            and v.strip().upper() not in ("TRUE", "FALSE"):
        return make_criterion(v + "*")
    return make_criterion(v)


def _computed(ctx, crit: RefValue, r: int, c: int, first_row: int | None):
    """A computed criterion (a formula under a label that isn't a database field): a predicate
    on the record's row that re-evaluates the formula with its relative references moved from
    the first record to that row."""
    from ..evaluator import Ctx, evaluate
    from ..parser import parse, translate

    sheet = crit.sheet
    f = sheet.formulas.get((r, c)) if hasattr(sheet, "formulas") else None
    if f is None or first_row is None:
        v = ctx.engine.get_value(sheet, r, c)
        return lambda row: _truthy(v)

    def pred(row):
        try:
            node = parse(translate(f.text, row - first_row, 0))
            v = evaluate(node, Ctx(ctx.engine, sheet, r, c))
            if isinstance(v, RefValue):
                v = ctx.engine.get_value(v.sheet, v.r1, v.c1)
        except Exception:
            return False
        return _truthy(v)

    return pred


def _truthy(v) -> bool:
    return v is True or (is_number(v) and v != 0)


def _records(ctx, database, criteria):
    """(headers, matching records) for a database range and a criteria range."""
    db = to_array(ctx, database)
    if db.nrows < 1:
        raise XLError(VALUE)
    headers = [_key(h) for h in db.rows[0]]
    index = {}
    for j, h in enumerate(headers):
        if h is not None and h != "":
            index.setdefault(h, j)
    first_row = database.r1 + 1 if isinstance(database, RefValue) else None

    cr = to_array(ctx, criteria)
    if (criteria.nrows if isinstance(criteria, RefValue) else cr.nrows) < 2:
        raise XLError(VALUE)  # a criteria range needs labels and at least one row of conditions
    # Rows: list of [(field index, predicate on value) | (None, predicate on sheet row)].
    # A blank criteria row (including blank rows past the sheet's used area) matches everything.
    rows = []
    match_all = isinstance(criteria, RefValue) and criteria.r1 + cr.nrows - 1 < criteria.r2
    for i in range(1, cr.nrows):
        conds = []
        for j, v in enumerate(cr.rows[i]):
            if v is None or v == "":
                continue
            label = _key(cr.rows[0][j])
            col = index.get(label) if label is not None and label != "" else None
            if col is not None:
                conds.append((col, _criterion(v)))
            elif isinstance(criteria, RefValue):
                conds.append((None, _computed(ctx, criteria, criteria.r1 + i, criteria.c1 + j, first_row)))
            else:
                conds.append((None, lambda row, v=v: _truthy(v)))
        if not conds:
            match_all = True
            break
        rows.append(conds)

    out = []
    for i in range(1, db.nrows):
        rec = db.rows[i]
        row = first_row + i - 1 if first_row is not None else None
        if match_all or any(all(p(rec[col]) if col is not None else p(row) for col, p in conds)
                            for conds in rows):
            out.append(rec)
    return db.rows[0], headers, out


def _field(ctx, field, headers, optional=False):
    if field is MISSING:
        if optional:
            return None
        raise XLError(VALUE)
    f = scalar_arg(ctx, field)
    if f is None:
        if optional:
            return None
        raise XLError(VALUE)
    if type(f) is str:
        k = _key(f)
        for j, h in enumerate(headers):
            if h == k:
                return j
        raise XLError(VALUE)
    if type(f) is bool:
        raise XLError(VALUE)
    n = to_int(f)
    if not 1 <= n <= len(headers):
        raise XLError(VALUE)
    return n - 1


def _values(ctx, database, field, criteria, optional=False):
    """The field's values in matching records (None as the field means whole records)."""
    _, headers, recs = _records(ctx, database, criteria)
    col = _field(ctx, field, headers, optional)
    if col is None:
        return None, recs
    return [r[col] for r in recs], recs


def _numbers(vals):
    out = []
    for v in vals:
        if is_number(v):
            out.append(v)
        elif isinstance(v, ExcelError):
            raise XLError(v)
    return out


@fn("DSUM", kind="raw", min_args=3, max_args=3)
def DSUM(ctx, database, field, criteria):
    vals, _ = _values(ctx, database, field, criteria)
    return _sum(_numbers(vals))


@fn("DAVERAGE", kind="raw", min_args=3, max_args=3)
def DAVERAGE(ctx, database, field, criteria):
    vals, _ = _values(ctx, database, field, criteria)
    nums = _numbers(vals)
    if not nums:
        raise XLError(DIV0)
    return math.fsum(nums) / len(nums)


@fn("DCOUNT", kind="raw", min_args=3, max_args=3)
def DCOUNT(ctx, database, field, criteria):
    vals, recs = _values(ctx, database, field, criteria, optional=True)
    if vals is None:
        return len(recs)
    return sum(1 for v in vals if is_number(v))


@fn("DCOUNTA", kind="raw", min_args=3, max_args=3)
def DCOUNTA(ctx, database, field, criteria):
    vals, recs = _values(ctx, database, field, criteria, optional=True)
    if vals is None:
        return len(recs)
    return sum(1 for v in vals if v is not None and v != "")


@fn("DGET", kind="raw", min_args=3, max_args=3)
def DGET(ctx, database, field, criteria):
    vals, _ = _values(ctx, database, field, criteria)
    if not vals:
        raise XLError(VALUE)
    if len(vals) > 1:
        raise XLError(NUM)
    v = vals[0]
    return 0 if v is None else v


@fn("DMAX", kind="raw", min_args=3, max_args=3)
def DMAX(ctx, database, field, criteria):
    nums = _numbers(_values(ctx, database, field, criteria)[0])
    return max(nums) if nums else 0


@fn("DMIN", kind="raw", min_args=3, max_args=3)
def DMIN(ctx, database, field, criteria):
    nums = _numbers(_values(ctx, database, field, criteria)[0])
    return min(nums) if nums else 0


@fn("DPRODUCT", kind="raw", min_args=3, max_args=3)
def DPRODUCT(ctx, database, field, criteria):
    nums = _numbers(_values(ctx, database, field, criteria)[0])
    return math.prod(nums) if nums else 0


def _stat(sample: bool, root: bool):
    def impl(ctx, database, field, criteria):
        v = _var(_numbers(_values(ctx, database, field, criteria)[0]), sample)
        return math.sqrt(v) if root else v

    return impl


fn("DSTDEV", kind="raw", min_args=3, max_args=3)(_stat(True, True))
fn("DSTDEVP", kind="raw", min_args=3, max_args=3)(_stat(False, True))
fn("DVAR", kind="raw", min_args=3, max_args=3)(_stat(True, False))
fn("DVARP", kind="raw", min_args=3, max_args=3)(_stat(False, False))
