"""GROUPBY and PIVOTBY: aggregate values grouped by row (and column) fields."""
from __future__ import annotations

import functools

from ..errors import CALC, VALUE, ExcelError, XLError
from ..registry import fn, scalar_arg, to_array
from ..values import MISSING, Array, FuncValue, LambdaValue, RefValue, is_number, to_bool, to_int
from .lookup_fns import sort_cmp


class _Group:
    """One group of rows; children are the next field's groups (sorted into `kids`)."""

    __slots__ = ("key", "rows", "children", "kids", "parent", "level", "_set")

    def __init__(self, key, parent, level):
        self.key, self.parent, self.level = key, parent, level
        self.rows: list[int] = []
        self.children: dict = {}
        self.kids: list[_Group] = []
        self._set = None

    def rowset(self) -> set:
        if self._set is None:
            self._set = set(self.rows)
        return self._set

    def path(self) -> list:
        keys, g = [], self
        while g.parent is not None:
            keys.append(g.key)
            g = g.parent
        return keys[::-1]


def _norm(v):
    if isinstance(v, str):
        return ("s", v.lower())
    if type(v) is bool:
        return ("b", v)
    if is_number(v):
        return ("n", float(v))
    if isinstance(v, ExcelError):
        return ("e", v.code)
    return ("z",)


# ------------------------------------------------------------------ argument handling


def _functions(f) -> list:
    """The function argument: a function value or LAMBDA, or a row or column of them."""
    if isinstance(f, (FuncValue, LambdaValue)):
        return [f]
    if isinstance(f, ExcelError):
        raise XLError(f)
    if isinstance(f, Array) and (f.nrows == 1 or f.ncols == 1):
        items = list(f.flat())
        if items and all(isinstance(x, (FuncValue, LambdaValue)) for x in items):
            return items
    raise XLError(VALUE)


def _padded(arrays: list[Array]) -> list[Array]:
    """Give every array the same number of rows (references are clipped to the used area)."""
    n = max(a.nrows for a in arrays)
    return [a if a.nrows == n else Array(a.rows + [[None] * a.ncols] * (n - a.nrows)) for a in arrays]


def _headers(ctx, mode, values: Array, multi_level: bool) -> tuple[bool, bool]:
    """(data has a header row, show headers in the result)."""
    if mode is MISSING:
        has = values.nrows >= 2 and isinstance(values.rows[0][0], str) and is_number(values.rows[1][0])
        return has, has or multi_level
    m = to_int(scalar_arg(ctx, mode))
    if m not in (0, 1, 2, 3):
        raise XLError(VALUE)
    return m in (1, 3), m in (2, 3)


def _filter(ctx, flt, ndata: int, has_header: bool) -> list[int]:
    if flt is MISSING:
        return list(range(ndata))
    a = to_array(ctx, flt)
    if a.nrows != 1 and a.ncols != 1:
        raise XLError(VALUE)
    flags = list(a.flat())
    if has_header and len(flags) == ndata + 1:
        flags = flags[1:]
    if len(flags) > ndata:
        raise XLError(VALUE)
    keep = []
    for i, v in enumerate(flags):
        if isinstance(v, ExcelError):
            raise XLError(v)
        if v is not None and to_bool(v):
            keep.append(i)
    return keep


def _depth(ctx, v, nfields: int, table: bool) -> int:
    if v is MISSING:
        return 2 if nfields > 1 and not table else 1
    d = to_int(scalar_arg(ctx, v))
    limit = 1 if table else nfields
    return max(-limit, min(limit, d))


def _sort_specs(ctx, v, nfields: int, nvalues: int) -> list[int]:
    if v is MISSING:
        return []
    if isinstance(v, (RefValue, Array)):
        specs = []
        for x in to_array(ctx, v).flat():
            if isinstance(x, ExcelError):
                raise XLError(x)
            specs.append(to_int(x))
    else:
        specs = [to_int(scalar_arg(ctx, v))]
    for s in specs:
        if s == 0 or abs(s) > nfields + nvalues:
            raise XLError(VALUE)
    if len(specs) > 1 and any(abs(s) > nfields for s in specs):
        raise XLError(VALUE)  # a list of sort columns may only name fields
    return specs


# ------------------------------------------------------------------ grouping


def _tree(keys: list[list], idxs: list[int], table: bool) -> _Group:
    root = _Group(None, None, 0)
    root.rows = list(idxs)
    for i in idxs:
        path = [tuple(keys[i])] if table else keys[i]
        node = root
        for v in path:
            nk = tuple(_norm(x) for x in v) if table else _norm(v)
            ch = node.children.get(nk)
            if ch is None:
                ch = node.children[nk] = _Group(v, node, node.level + 1)
            ch.rows.append(i)
            node = ch
    return root


def _sort_tree(root: _Group, specs: list[int], nfields: int, table: bool, total):
    """Order every group's children. Field sorts apply at their own level; value sorts (columns
    after the fields) compare the groups' totals at every level."""

    def cmp(a: _Group, b: _Group, level: int) -> int:
        for s in specs:
            col, desc = abs(s), s < 0
            if col <= nfields:
                if table:
                    c = sort_cmp(a.key[col - 1], b.key[col - 1], desc)
                elif col == level:
                    c = sort_cmp(a.key, b.key, desc)
                else:
                    continue
            else:
                c = sort_cmp(total(a, col - nfields - 1), total(b, col - nfields - 1), desc)
            if c:
                return c
        if table:
            for x, y in zip(a.key, b.key):
                c = sort_cmp(x, y)
                if c:
                    return c
            return 0
        return sort_cmp(a.key, b.key)

    stack = [root]
    while stack:
        g = stack.pop()
        level = g.level + 1
        g.kids = sorted(g.children.values(), key=functools.cmp_to_key(lambda a, b, lv=level: cmp(a, b, lv)))
        stack.extend(g.kids)


def _entries(root: _Group, nfields: int, depth: int, table: bool) -> list[tuple[list, _Group]]:
    """Output lines in order: (labels for the field columns, group). Subtotal lines show the
    group's own labels with blanks for the deeper fields; the grand total line says Total."""
    out = []

    def labels(g: _Group) -> list:
        p = list(g.key) if table else g.path()
        return p + [None] * (nfields - len(p))

    def walk(g: _Group):
        for ch in g.kids:
            if table or ch.level == nfields:
                out.append((labels(ch), ch))
                continue
            sub = ch.level < abs(depth)
            if sub and depth < 0:
                out.append((labels(ch), ch))
            walk(ch)
            if sub and depth > 0:
                out.append((labels(ch), ch))

    walk(root)
    if depth:
        total = (["Total"] + [None] * (nfields - 1), root)
        if depth < 0:
            out.insert(0, total)
        else:
            out.append(total)
    return out


class _Aggregator:
    def __init__(self, ctx, values: Array, funcs: list):
        from ..evaluator import invoke

        self.ctx, self.funcs, self.call = ctx, funcs, invoke
        self.cols = [list(c) for c in zip(*values.rows)] if values.rows else []
        self.cache: dict = {}

    def __call__(self, rows: list[int], j: int, f, total_rows: list[int] | None):
        from ..evaluator import arity

        col = self.cols[j]
        args = [Array([[col[i]] for i in rows])]
        if arity(f)[0] >= 2:  # PERCENTOF, or a LAMBDA(values, all_values)
            args.append(Array([[col[i]] for i in total_rows] or [[None]]))
        r = self.call(self.ctx, f, args)
        if isinstance(r, Array):
            r = r.rows[0][0] if r.rows and r.rows[0] else CALC
        elif isinstance(r, RefValue):
            r = self.ctx.engine.get_value(r.sheet, r.r1, r.c1)
        return None if r is MISSING else r

    def total(self, rows_of_all: list[int]):
        """Sort key for a group: its first function applied to value column j."""

        def get(g: _Group, j: int):
            key = (id(g), j)
            if key not in self.cache:
                self.cache[key] = self(g.rows, j, self.funcs[0], rows_of_all)
            return self.cache[key]

        return get


def _labels(funcs: list, vheaders: list, nv: int) -> list:
    if len(funcs) == 1:
        return list(vheaders)
    if nv == 1:
        return [f.name for f in funcs]
    return [f"{h} {f.name}" if h is not None else f.name for h in vheaders for f in funcs]


def _split(a: Array, has_header: bool, generated: str):
    if has_header:
        return list(a.rows[0]), Array(a.rows[1:])
    names = [f"{generated} {j + 1}" if a.ncols > 1 else generated for j in range(a.ncols)]
    return names, a


@fn("GROUPBY", kind="raw", min_args=3, max_args=8)
def GROUPBY(ctx, row_fields, values, function, field_headers=MISSING, total_depth=MISSING,
            sort_order=MISSING, filter_array=MISSING, field_relationship=MISSING):
    funcs = _functions(function)
    rf, vals = _padded([to_array(ctx, row_fields), to_array(ctx, values)])
    nrf, nv = rf.ncols, vals.ncols
    rel = 0 if field_relationship is MISSING else to_int(scalar_arg(ctx, field_relationship))
    if rel not in (0, 1):
        raise XLError(VALUE)
    table = rel == 1 and nrf > 1
    has, show = _headers(ctx, field_headers, vals, nrf > 1)
    rheaders, rf = _split(rf, has, "Row Field")
    vheaders, vals = _split(vals, has, "Values")
    idxs = _filter(ctx, filter_array, rf.nrows, has)
    if not idxs:
        raise XLError(CALC)
    depth = _depth(ctx, total_depth, nrf, table)
    specs = _sort_specs(ctx, sort_order, nrf, nv)

    agg = _Aggregator(ctx, vals, funcs)
    root = _tree(rf.rows, idxs, table)
    _sort_tree(root, specs, nrf, table, agg.total(idxs))
    cols = [(j, f) for j in range(nv) for f in funcs]
    out = []
    if show:
        out.append(rheaders + _labels(funcs, vheaders, nv))
    for labels, g in _entries(root, nrf, depth, table):
        out.append(labels + [agg(g.rows, j, f, idxs) for j, f in cols])
    return Array(out)


@fn("PIVOTBY", kind="raw", min_args=4, max_args=11)
def PIVOTBY(ctx, row_fields, col_fields, values, function, field_headers=MISSING, row_total_depth=MISSING,
            row_sort_order=MISSING, col_total_depth=MISSING, col_sort_order=MISSING, filter_array=MISSING,
            relative_to=MISSING):
    funcs = _functions(function)
    rf, cf, vals = _padded([to_array(ctx, row_fields), to_array(ctx, col_fields), to_array(ctx, values)])
    nrf, ncf, nv = rf.ncols, cf.ncols, vals.ncols
    has, show = _headers(ctx, field_headers, vals, nrf > 1 or ncf > 1)
    rheaders, rf = _split(rf, has, "Row Field")
    cheaders, cf = _split(cf, has, "Column Field")
    vheaders, vals = _split(vals, has, "Values")
    idxs = _filter(ctx, filter_array, rf.nrows, has)
    if not idxs:
        raise XLError(CALC)
    rel = 0 if relative_to is MISSING else to_int(scalar_arg(ctx, relative_to))
    if rel not in (0, 1, 2, 3, 4):
        raise XLError(VALUE)

    agg = _Aggregator(ctx, vals, funcs)
    rroot = _tree(rf.rows, idxs, False)
    croot = _tree(cf.rows, idxs, False)
    _sort_tree(rroot, _sort_specs(ctx, row_sort_order, nrf, nv), nrf, False, agg.total(idxs))
    _sort_tree(croot, _sort_specs(ctx, col_sort_order, ncf, nv), ncf, False, agg.total(idxs))
    rentries = _entries(rroot, nrf, _depth(ctx, row_total_depth, nrf, False), False)
    centries = _entries(croot, ncf, _depth(ctx, col_total_depth, ncf, False), False)
    cols = [(j, f) for j in range(nv) for f in funcs]
    two_args = any(f.spec.min_args >= 2 for f in funcs)

    def total_rows(rg: _Group, cg: _Group) -> list[int]:
        if rel == 0:
            return cg.rows
        if rel == 1:
            return rg.rows
        if rel == 2:
            return idxs
        if rel == 3:
            pc = cg.parent or cg
            return [i for i in rg.rows if i in pc.rowset()]
        pr = rg.parent or rg
        return [i for i in cg.rows if i in pr.rowset()]

    out = []
    for level in range(ncf):
        row = [None] * nrf
        if show:
            row[-1] = cheaders[level]
        for labels, _ in centries:
            row.extend([labels[level]] * len(cols))
        out.append(row)
    if show or len(cols) > 1:
        row = list(rheaders) if show else [None] * nrf
        names = _labels(funcs, vheaders, nv)
        for _ in centries:
            row.extend(names if len(cols) > 1 or show else [None] * len(cols))
        out.append(row)
    for rlabels, rg in rentries:
        row = list(rlabels)
        for _, cg in centries:
            cset = cg.rowset()
            rows = [i for i in rg.rows if i in cset]
            tot = total_rows(rg, cg) if two_args else None
            for j, f in cols:
                row.append(agg(rows, j, f, tot) if rows else None)
        out.append(row)
    return Array(out)
