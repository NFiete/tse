"""AST evaluation with Excel semantics (coercion, error propagation, array lifting)."""
from __future__ import annotations

import math

from . import functions  # noqa: F401  (registers all functions)
from .errors import CALC, DIV0, NA, NAME, NUM, REF, VALUE, ExcelError, Unsupported, XLError
from .parser import (
    ArrayNode, Binary, Call, Const, Func, LocalNode, Missing, NameNode, Node, RefNode, StructRefNode, Unary,
)
from .registry import FUNCTIONS, to_array
from .values import MISSING, Array, FuncValue, LambdaValue, RefValue, compare, fix_number, to_number, to_text

_NO_LOCALS: dict = {}
MAX_LAMBDA_DEPTH = 1000  # nested LAMBDA calls (recursion through a defined name) before #NUM!


class Ctx:
    """Evaluation context: the engine plus the formula cell's position."""

    __slots__ = ("engine", "sheet", "row", "col", "depth", "arrays", "legacy", "env", "calls")

    def __init__(self, engine, sheet, row: int, col: int):
        self.engine, self.sheet, self.row, self.col = engine, sheet, row, col
        self.depth = 0
        self.env = _NO_LOCALS  # LET names and LAMBDA parameters in scope: key -> value
        self.calls = 0  # LAMBDA call depth
        # Set when evaluation worked on arrays (an operator or function mapped over a range, or a
        # function returned an array): the formula needs dynamic-array semantics.
        self.arrays = False
        # Pre-dynamic-array semantics: a range where a single value is expected gives the cell in
        # the formula's row or column (implicit intersection) instead of being mapped over.
        # Arguments of functions that take ranges or arrays (SUMPRODUCT, INDEX...) are still
        # evaluated as arrays.
        self.legacy = False


def _is_multi(v) -> bool:
    return isinstance(v, Array) or (isinstance(v, RefValue) and not v.is_cell)


def _deref(ctx: Ctx, v):
    if isinstance(v, RefValue):
        return ctx.engine.get_value(v.sheet, v.r1, v.c1)
    return v


def _safe(fn, *args):
    try:
        r = fn(*args)
        t = type(r)
        if t is int or t is float:
            return fix_number(r)
        return r
    except XLError as e:
        return e.err
    except ZeroDivisionError:
        return DIV0
    except (OverflowError, ValueError):
        return NUM


def lift(ctx: Ctx, fn, args: list, errors_ok: bool = False):
    """Apply a scalar function, mapping it over any array/range arguments."""
    if not any(_is_multi(a) for a in args):
        vals = [None if a is MISSING else _deref(ctx, a) for a in args]
        if not errors_ok:
            for v in vals:
                if isinstance(v, ExcelError):
                    return v
        return _safe(fn, *vals)
    if ctx.legacy and any(isinstance(a, RefValue) and not a.is_cell for a in args):
        args = [intersect(ctx, a) if isinstance(a, RefValue) and not a.is_cell else a for a in args]
        return lift(ctx, fn, args, errors_ok)
    ctx.arrays = True
    arrs = [to_array(ctx, a) if _is_multi(a) else None for a in args]
    nr = max(a.nrows for a in arrs if a is not None)
    nc = max(a.ncols for a in arrs if a is not None)
    scal = [None if (a is None and args[i] is MISSING) else (_deref(ctx, args[i]) if a is None else None)
            for i, a in enumerate(arrs)]
    rows = []
    for i in range(nr):
        row = []
        for j in range(nc):
            vals = []
            bad = None
            for k, a in enumerate(arrs):
                if a is None:
                    v = scal[k]
                else:
                    ii = 0 if a.nrows == 1 else i
                    jj = 0 if a.ncols == 1 else j
                    if ii >= a.nrows or jj >= a.ncols:
                        v = NA
                    else:
                        v = a.rows[ii][jj]
                if not errors_ok and isinstance(v, ExcelError) and bad is None:
                    bad = v
                vals.append(v)
            row.append(bad if bad is not None else _safe(fn, *vals))
        rows.append(row)
    return Array(rows)


# ------------------------------------------------------------------ operators


def _add(a, b):
    return to_number(a) + to_number(b)


def _sub(a, b):
    return to_number(a) - to_number(b)


def _mul(a, b):
    return to_number(a) * to_number(b)


def _div(a, b):
    a, b = to_number(a), to_number(b)
    if b == 0:
        raise XLError(DIV0)
    r = a / b
    return int(r) if r.is_integer() and abs(r) < 2**53 and type(a) is int and type(b) is int else r


def _pow(a, b):
    a, b = float(to_number(a)), float(to_number(b))
    if a == 0 and b == 0:
        raise XLError(NUM)
    if a == 0 and b < 0:
        raise XLError(DIV0)
    if a < 0 and not b.is_integer():
        raise XLError(NUM)
    r = a ** b
    if isinstance(r, complex):
        raise XLError(NUM)
    return int(r) if r.is_integer() and abs(r) < 2**53 else r


def _concat(a, b):
    return to_text(a) + to_text(b)


def _neg(a):
    return -to_number(a)


def _pct(a):
    return to_number(a) / 100


_BINOPS = {
    "+": _add, "-": _sub, "*": _mul, "/": _div, "^": _pow, "&": _concat,
    "=": lambda a, b: compare(a, b) == 0,
    "<>": lambda a, b: compare(a, b) != 0,
    "<": lambda a, b: compare(a, b) < 0,
    ">": lambda a, b: compare(a, b) > 0,
    "<=": lambda a, b: compare(a, b) <= 0,
    ">=": lambda a, b: compare(a, b) >= 0,
}


# ------------------------------------------------------------------ evaluation


def evaluate(node: Node, ctx: Ctx):
    t = type(node)
    if t is Const:
        return node.value
    if t is RefNode:
        if node.sheet is None:
            sheet = ctx.sheet
        else:
            sheet = ctx.engine.sheet_by_name(node.sheet)
            if sheet is None:
                return REF
        return RefValue(sheet, node.r1, node.c1, node.r2, node.c2)
    if t is Func:
        return _call(node, ctx)
    if t is LocalNode:
        v = ctx.env.get(node.key, NAME)
        return v
    if t is Call:
        f = evaluate(node.callee, ctx)
        return invoke(ctx, f, [evaluate(a, ctx) for a in node.args])
    if t is StructRefNode:
        return ctx.engine.resolve_sref(node.ref, ctx.sheet, ctx.row, ctx.col)
    if t is Binary:
        op = node.op
        left = evaluate(node.left, ctx)
        right = evaluate(node.right, ctx)
        if op == ":":
            if isinstance(left, RefValue) and isinstance(right, RefValue) and left.sheet is right.sheet:
                return RefValue(left.sheet, min(left.r1, right.r1), min(left.c1, right.c1),
                                max(left.r2, right.r2), max(left.c2, right.c2))
            if isinstance(left, ExcelError):
                return left
            if isinstance(right, ExcelError):
                return right
            return VALUE
        return lift(ctx, _BINOPS[op], [left, right])
    if t is Unary:
        v = evaluate(node.arg, ctx)
        if node.op == "+":
            return v
        return lift(ctx, _neg if node.op == "-" else _pct, [v])
    if t is NameNode:
        target = ctx.engine.resolve_name(node.name, ctx.sheet, node.sheet)
        if target is None:
            if node.sheet is None:
                fv = function_value(node.name)
                if fv is not None:
                    return fv
            return NAME
        return _eval_name(ctx, target)
    if t is ArrayNode:
        return Array([list(r) for r in node.rows])
    if t is Missing:
        return MISSING
    raise Unsupported(f"Cannot evaluate {node!r}")


def _eval_name(ctx: Ctx, target: Node):
    """Evaluate a defined name's formula (which can't see the caller's LET names)."""
    ctx.depth += 1
    env, ctx.env = ctx.env, _NO_LOCALS
    try:
        if ctx.depth > 50:
            return REF
        return evaluate(target, ctx)
    finally:
        ctx.depth -= 1
        ctx.env = env


def _let(node: Func, ctx: Ctx):
    args = node.args
    env = dict(ctx.env)
    saved = ctx.env
    try:
        for i in range(0, len(args) - 1, 2):
            ctx.env = env
            v = evaluate(args[i + 1], ctx)
            env = dict(env)
            env[args[i].key] = v
        ctx.env = env
        return evaluate(args[-1], ctx)
    finally:
        ctx.env = saved


def _lambda(node: Func, ctx: Ctx):
    return LambdaValue(tuple(p.key for p in node.args[:-1]), node.args[-1], ctx.env)


_SYNTAX = {"LET": _let, "LAMBDA": _lambda}


def arity(f) -> tuple[int, int | None]:
    """(minimum, maximum) number of arguments a function value takes."""
    if isinstance(f, LambdaValue):
        return len(f.params), len(f.params)
    return f.spec.min_args, f.spec.max_args


def is_callable(v) -> bool:
    return isinstance(v, (LambdaValue, FuncValue))


def invoke(ctx: Ctx, f, args: list):
    """Call a LAMBDA or function value with evaluated arguments; returns a value (errors as
    values). Anything else gives #VALUE! (or the error it is)."""
    if isinstance(f, LambdaValue):
        if len(args) != len(f.params):
            return VALUE
        if ctx.calls >= MAX_LAMBDA_DEPTH:
            return NUM
        env = dict(f.env)
        env.update(zip(f.params, args))
        saved = ctx.env
        ctx.env = env
        ctx.calls += 1
        try:
            return evaluate(f.body, ctx)
        except RecursionError:
            return NUM
        finally:
            ctx.calls -= 1
            ctx.env = saved
    if isinstance(f, FuncValue):
        return call_function(ctx, f, args)
    if isinstance(f, RefValue) and f.is_cell:
        f = ctx.engine.get_value(f.sheet, f.r1, f.c1)
    if isinstance(f, ExcelError):
        return f
    return VALUE


def function_value(name: str) -> FuncValue | None:
    """The function a bare name such as SUM or _xleta.SUM refers to, if any."""
    up = name.upper()
    if up.startswith("_XLETA."):
        up = up[7:]
    for p in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
        if up.startswith(p):
            up = up[len(p):]
            break
    spec = FUNCTIONS.get(up)
    return FuncValue(up, spec) if spec is not None else None


def call_function(ctx: Ctx, fv: FuncValue, args: list):
    """Call a function value with already-evaluated arguments; returns a value (errors as values)."""
    spec = fv.spec
    n = len(args)
    if n < spec.min_args or (spec.max_args is not None and n > spec.max_args):
        return VALUE
    if spec.kind == "lazy":
        try:
            return spec.impl(ctx, *[(lambda a=a: a) for a in args])
        except XLError as e:
            return e.err
    return _apply(spec, ctx, args)


def _call(node: Func, ctx: Ctx):
    special = _SYNTAX.get(node.name)
    if special is not None:
        return special(node, ctx)
    spec = FUNCTIONS.get(node.name)
    if spec is None:
        return _call_named(node, ctx)
    n = len(node.args)
    if n < spec.min_args or (spec.max_args is not None and n > spec.max_args):
        return VALUE
    if spec.kind == "lazy":
        thunks = [(lambda a=a: evaluate(a, ctx)) for a in node.args]
        try:
            return spec.impl(ctx, *thunks)
        except XLError as e:
            return e.err
    if ctx.legacy and spec.kind == "raw":
        ctx.legacy = False
        try:
            args = [evaluate(a, ctx) for a in node.args]
        finally:
            ctx.legacy = True
        return _apply(spec, ctx, args)
    return _apply(spec, ctx, [evaluate(a, ctx) for a in node.args])


def _call_named(node: Func, ctx: Ctx):
    """A call such as MYFUNC(1) where MYFUNC is a defined name holding a LAMBDA."""
    resolve = getattr(ctx.engine, "resolve_name", None)
    target = resolve(node.name, ctx.sheet, None, tables=False) if resolve is not None else None
    if target is None:
        raise Unsupported(node.name)
    f = _eval_name(ctx, target)
    if isinstance(f, ExcelError):
        return f
    if not is_callable(f):
        return NAME
    return invoke(ctx, f, [evaluate(a, ctx) for a in node.args])


def _apply(spec, ctx: Ctx, args: list):
    if spec.kind == "raw":
        try:
            r = spec.impl(ctx, *args)
        except XLError as e:
            return e.err
        except ZeroDivisionError:
            return DIV0
        except (OverflowError, ValueError):
            return NUM
        t = type(r)
        if t is int or t is float:
            try:
                return fix_number(r)
            except XLError as e:
                return e.err
        if t is Array:
            ctx.arrays = True
        return r
    return lift(ctx, spec.impl, args, spec.errors_ok)


def intersect(ctx: Ctx, v: RefValue):
    """Implicit intersection: the cell of a range in the formula's row or column, else #VALUE!."""
    r, c = ctx.row, ctx.col
    if v.c1 == v.c2 and v.r1 <= r <= v.r2:
        return RefValue(v.sheet, r, v.c1, r, v.c1)
    if v.r1 == v.r2 and v.c1 <= c <= v.c2:
        return RefValue(v.sheet, v.r1, c, v.r1, c)
    if v.r1 <= r <= v.r2 and v.c1 <= c <= v.c2:
        return RefValue(v.sheet, r, c, r, c)
    return VALUE


def finalize(ctx: Ctx, v, legacy: bool = False):
    """Turn an evaluation result into a cell value. Arrays and ranges give their first value,
    or with legacy=True (formulas from files without dynamic-array support) a range gives the
    cell in the formula's row or column (implicit intersection)."""
    if isinstance(v, RefValue):
        if legacy and not v.is_cell:
            v = intersect(ctx, v)
            if isinstance(v, ExcelError):
                return v
        if isinstance(v, RefValue):
            v = ctx.engine.get_value(v.sheet, v.r1, v.c1)
    elif isinstance(v, Array):
        v = v.rows[0][0] if v.rows and v.rows[0] else NA
    return _cell_value(v)


def _cell_value(v, blank=0):
    if v is None or v is MISSING:
        return blank
    if isinstance(v, FuncValue):
        return NAME
    if isinstance(v, LambdaValue):
        return CALC
    if type(v) is float and (math.isnan(v) or math.isinf(v)):
        return NUM
    return v


def is_multi(v) -> bool:
    """Whether a result is an array or multi-cell range (i.e. a candidate for spilling)."""
    return isinstance(v, Array) or (isinstance(v, RefValue) and not v.is_cell)


def spill_rows(ctx: Ctx, v, size: tuple | None = None) -> list | None:
    """The cell values an array result spills into, as a list of rows, or None if the result is
    a single value. With size=(rows, cols) (legacy array formulas) the result is fitted to that
    size: single rows/columns repeat, and positions past the result's edge are #N/A."""
    if isinstance(v, RefValue):
        if v.is_cell and size is None:
            return None
        src = ctx.engine.range_rows(v.sheet, v.r1, v.c1, v.r2, v.c2)
        rows = [[_cell_value(x) for x in row] for row in src]
    elif isinstance(v, Array):
        if not v.rows or not v.rows[0]:
            return None
        width = max(len(r) for r in v.rows)
        rows = [[_cell_value(x, None) for x in row] + [NA] * (width - len(row)) for row in v.rows]
    elif size is None:
        return None
    else:
        rows = [[_cell_value(_deref(ctx, v))]]
    if size is not None:
        nr, nc = size
        ar, ac = len(rows), len(rows[0])
        rows = [[(rows[0 if ar == 1 else i][0 if ac == 1 else j]
                  if (ar == 1 or i < ar) and (ac == 1 or j < ac) else NA) for j in range(nc)]
                for i in range(nr)]
    elif len(rows) == 1 and len(rows[0]) == 1:
        return None
    return rows


def walk(node: Node):
    yield node
    t = type(node)
    if t is Func:
        for a in node.args:
            yield from walk(a)
    elif t is Call:
        yield from walk(node.callee)
        for a in node.args:
            yield from walk(a)
    elif t is Binary:
        yield from walk(node.left)
        yield from walk(node.right)
    elif t is Unary:
        yield from walk(node.arg)


def is_volatile(node: Node) -> bool:
    for n in walk(node):
        if type(n) is Func:
            spec = FUNCTIONS.get(n.name)
            if spec is not None and spec.volatile:
                return True
    return False


def unsupported_functions(node: Node) -> list[str]:
    return sorted({n.name for n in walk(node) if type(n) is Func and n.name not in FUNCTIONS})


_STATIC_REFS = (RefNode, StructRefNode)


def analyze(node: Node):
    """One pass over an AST: (ref nodes, name nodes, function names, has dynamic range,
    structured reference nodes)."""
    refs, names, funcs, srefs = [], [], set(), []
    dynamic = False
    stack = [node]
    pop, push = stack.pop, stack.append
    while stack:
        n = pop()
        t = type(n)
        if t is RefNode:
            refs.append(n)
        elif t is Func:
            funcs.add(n.name)
            stack.extend(n.args)
        elif t is Call:
            push(n.callee)
            stack.extend(n.args)
        elif t is Binary:
            if n.op == ":" and not (type(n.left) in _STATIC_REFS and type(n.right) in _STATIC_REFS):
                dynamic = True
            push(n.left)
            push(n.right)
        elif t is Unary:
            push(n.arg)
        elif t is NameNode:
            names.append(n)
        elif t is StructRefNode:
            srefs.append(n)
    return refs, names, funcs, dynamic, srefs

