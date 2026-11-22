"""Excel formula lexer and Pratt parser.

The lexer keeps source spans for every token so references can be rewritten
in place (copy/paste translation, row/column insertion, sheet renames)
without disturbing the rest of the formula text.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from ..model.refs import MAX_COL, MAX_ROW, col_index, col_letter
from .errors import ERROR_CODES, REF, ExcelError, Unsupported


class FormulaSyntaxError(Exception):
    pass


# --------------------------------------------------------------------------- tokens

NUM, STR, BOOL, ERR, REFT, NAMET, FUNC, OP, LP, RP, COMMA, SEMI, LBRACE, RBRACE = (
    "NUM", "STR", "BOOL", "ERR", "REF", "NAME", "FUNC", "OP", "(", ")", ",", ";", "{", "}",
)


@dataclass
class RefInfo:
    sheet: str | None  # unquoted sheet name, None = same sheet
    kind: str  # 'cell' | 'range' | 'col' | 'row'
    r1: int
    c1: int
    r2: int
    c2: int
    ar1: bool = False
    ac1: bool = False
    ar2: bool = False
    ac2: bool = False
    sheet_raw: str = ""  # prefix exactly as written, including the '!'


@dataclass
class Token:
    kind: str
    text: str
    start: int
    end: int
    value: object = None
    ref: RefInfo | None = None
    sheet: str | None = None  # for sheet-qualified names


_ID_START = r"A-Za-z_\\¡-￿"
_ID_REST = r"A-Za-z0-9_.\\?¡-￿"
_NOTREF = r"(?![A-Za-z0-9_(.!\[])"
_PFX = r"'(?:[^']|'')+'|[A-Za-z0-9_.¡-￿]+"
_MASTER = re.compile(
    r"(?P<ws>[ \t\r\n]+)"
    r"|(?P<at>@)"
    r'|(?P<str>"(?:[^"]|"")*")'
    r"|(?P<err>#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|SPILL!|CALC!|GETTING_DATA))"
    r"|(?P<bracket>\[)"
    rf"|(?:(?P<pfx>{_PFX})!)?(?:"
    rf"(?P<xr>\$?[A-Za-z]{{1,3}}\$?\d+:\$?[A-Za-z]{{1,3}}\$?\d+){_NOTREF}"
    rf"|(?P<cell>\$?[A-Za-z]{{1,3}}\$?\d+){_NOTREF}"
    rf"|(?P<cr>\$?[A-Za-z]{{1,3}}:\$?[A-Za-z]{{1,3}}){_NOTREF}"
    rf"|(?P<rr>\$?\d+:\$?\d+){_NOTREF})"
    rf"|(?P<pfx2>{_PFX})!(?:(?P<perr>#REF!)|(?P<pname>[{_ID_START}][{_ID_REST}]*))"
    r"|(?P<num>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)"
    r"|(?P<bool>(?i:TRUE|FALSE))(?![A-Za-z0-9_.(])"
    rf"|(?P<func>[{_ID_START}][{_ID_REST}]*)\("
    rf"|(?P<name>[{_ID_START}][{_ID_REST}]*)(?P<sref>\[)?"
    r"|(?P<op><>|<=|>=|[-+*/^&=<>%:])"
    r"|(?P<punct>[(),;{}])"
)
_CELLPART = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")
_ROWPART = re.compile(r"(\$?)(\d+)")
_COLPART = re.compile(r"(\$?)([A-Za-z]{1,3})")


def _unquote_sheet(p: str) -> str:
    if p.startswith("'"):
        return p[1:-1].replace("''", "'")
    return p


def _make_ref(kind: str, body: str, sheet, sheet_raw: str) -> RefInfo | None:
    if kind == "cell":
        ac, cl, ar, rw = _CELLPART.fullmatch(body).groups()
        c, r = col_index(cl), int(rw)
        if not (1 <= r <= MAX_ROW and c <= MAX_COL):
            return None
        return RefInfo(sheet, "cell", r, c, r, c, bool(ar), bool(ac), bool(ar), bool(ac), sheet_raw)
    a, b = body.split(":")
    if kind == "xr":
        ac1, c1, ar1, r1 = _CELLPART.fullmatch(a).groups()
        ac2, c2, ar2, r2 = _CELLPART.fullmatch(b).groups()
        c1, c2, r1, r2 = col_index(c1), col_index(c2), int(r1), int(r2)
        if not (1 <= r1 <= MAX_ROW and 1 <= r2 <= MAX_ROW and c1 <= MAX_COL and c2 <= MAX_COL):
            return None
        return RefInfo(sheet, "range", r1, c1, r2, c2, bool(ar1), bool(ac1), bool(ar2), bool(ac2), sheet_raw)
    if kind == "cr":
        a1, c1 = _COLPART.fullmatch(a).groups()
        a2, c2 = _COLPART.fullmatch(b).groups()
        c1, c2 = col_index(c1), col_index(c2)
        if c1 > MAX_COL or c2 > MAX_COL:
            return None
        return RefInfo(sheet, "col", 1, c1, MAX_ROW, c2, True, bool(a1), True, bool(a2), sheet_raw)
    a1, r1 = _ROWPART.fullmatch(a).groups()
    a2, r2 = _ROWPART.fullmatch(b).groups()
    r1, r2 = int(r1), int(r2)
    if not (1 <= r1 <= MAX_ROW and 1 <= r2 <= MAX_ROW):
        return None
    return RefInfo(sheet, "row", r1, 1, r2, MAX_COL, bool(a1), True, bool(a2), True, sheet_raw)


def tokenize(src: str) -> list[Token]:
    """Tokenize a formula body (without the leading '=')."""
    toks: list[Token] = []
    append = toks.append
    i, n = 0, len(src)
    match = _MASTER.match
    while i < n:
        m = match(src, i)
        if m is None:
            ch = src[i]
            if ch == '"':
                raise FormulaSyntaxError("Unterminated string")
            if ch == "#":
                raise Unsupported("Unknown # token")
            raise FormulaSyntaxError(f"Unexpected character {ch!r}")
        kind = m.lastgroup
        j = m.end()
        text = m.group()
        if kind in ("xr", "cell", "cr", "rr"):
            pfx = m.group("pfx")
            sheet = _unquote_sheet(pfx) if pfx else None
            if sheet is not None and "[" in sheet:
                raise Unsupported("External workbook references are not supported")
            info = _make_ref(kind, m.group(kind), sheet, (pfx + "!") if pfx else "")
            if info is None:
                # e.g. XFE1 (past the last column): treat as a name
                if kind == "cell" and not pfx:
                    append(Token(NAMET, text, i, j, text))
                    i = j
                    continue
                raise FormulaSyntaxError(f"Invalid reference {text}")
            append(Token(REFT, text, i, j, ref=info))
        elif kind == "ws" or kind == "at":
            pass
        elif kind == "num":
            v = float(text)
            if "." not in text and "e" not in text and "E" not in text:
                v = int(text)
            append(Token(NUM, text, i, j, v))
        elif kind == "str":
            append(Token(STR, text, i, j, text[1:-1].replace('""', '"')))
        elif kind == "op":
            append(Token(OP, text, i, j))
        elif kind == "punct":
            append(Token(text, text, i, j))
        elif kind == "func":
            append(Token(FUNC, text, i, j, text[:-1]))
        elif kind == "name" or kind == "sref":
            if m.group("sref"):
                raise Unsupported("Structured references are not supported")
            append(Token(NAMET, text, i, j, text))
        elif kind == "bool":
            append(Token(BOOL, text, i, j, text.upper() == "TRUE"))
        elif kind == "err":
            append(Token(ERR, text, i, j, ERROR_CODES[text.upper()]))
        elif kind == "perr":
            append(Token(ERR, text, i, j, ERROR_CODES["#REF!"]))
        elif kind == "pname":
            sheet = _unquote_sheet(m.group("pfx2"))
            if "[" in sheet:
                raise Unsupported("External workbook references are not supported")
            append(Token(NAMET, text, i, j, m.group("pname"), sheet=sheet))
        elif kind == "bracket":
            raise Unsupported("Structured or external references are not supported")
        else:
            raise FormulaSyntaxError(f"Unexpected {text!r}")
        i = j
    return toks


# --------------------------------------------------------------------------- AST


class Node:
    __slots__ = ()


class Const(Node):
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value

    def __repr__(self):
        return f"Const({self.value!r})"


class Missing(Node):
    __slots__ = ()

    def __repr__(self):
        return "Missing()"


class RefNode(Node):
    __slots__ = ("sheet", "r1", "c1", "r2", "c2", "kind")

    def __init__(self, info: RefInfo):
        self.sheet = info.sheet
        self.r1, self.c1, self.r2, self.c2 = info.r1, info.c1, info.r2, info.c2
        self.kind = info.kind
        if self.r1 > self.r2:
            self.r1, self.r2 = self.r2, self.r1
        if self.c1 > self.c2:
            self.c1, self.c2 = self.c2, self.c1

    def __repr__(self):
        return f"Ref({self.sheet}!{col_letter(self.c1)}{self.r1}:{col_letter(self.c2)}{self.r2})"


class NameNode(Node):
    __slots__ = ("name", "sheet")

    def __init__(self, name: str, sheet: str | None = None):
        self.name, self.sheet = name, sheet

    def __repr__(self):
        return f"Name({self.name})"


class Func(Node):
    __slots__ = ("name", "args", "raw_name")

    def __init__(self, name: str, args: list[Node], raw_name: str):
        self.name, self.args, self.raw_name = name, args, raw_name

    def __repr__(self):
        return f"{self.name}({', '.join(map(repr, self.args))})"


class Unary(Node):
    __slots__ = ("op", "arg")

    def __init__(self, op: str, arg: Node):
        self.op, self.arg = op, arg

    def __repr__(self):
        return f"({self.op}{self.arg!r})" if self.op != "%" else f"({self.arg!r}%)"


class Binary(Node):
    __slots__ = ("op", "left", "right")

    def __init__(self, op: str, left: Node, right: Node):
        self.op, self.left, self.right = op, left, right

    def __repr__(self):
        return f"({self.left!r} {self.op} {self.right!r})"


class ArrayNode(Node):
    __slots__ = ("rows",)

    def __init__(self, rows: list[list]):
        self.rows = rows

    def __repr__(self):
        return f"Array({self.rows!r})"


_INFIX_BP = {
    "=": 10, "<>": 10, "<": 10, ">": 10, "<=": 10, ">=": 10,
    "&": 20,
    "+": 30, "-": 30,
    "*": 40, "/": 40,
    "^": 50,
    ":": 80,
}
_POSTFIX_BP = 60
_PREFIX_BP = 70

_FUNC_PREFIXES = ("_XLFN._XLWS.", "_XLFN.", "_XLWS.")


def canonical_func_name(raw: str) -> str:
    name = raw.upper()
    for p in _FUNC_PREFIXES:
        if name.startswith(p):
            return name[len(p):]
    return name


class _Parser:
    def __init__(self, toks: list[Token]):
        self.toks = toks
        self.i = 0

    def peek(self) -> Token | None:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def next(self) -> Token:
        if self.i >= len(self.toks):
            raise FormulaSyntaxError("Unexpected end of formula")
        t = self.toks[self.i]
        self.i += 1
        return t

    def expect(self, kind: str) -> Token:
        t = self.next()
        if t.kind != kind:
            raise FormulaSyntaxError(f"Expected {kind!r} but found {t.text!r}")
        return t

    def parse(self) -> Node:
        if not self.toks:
            raise FormulaSyntaxError("Empty formula")
        node = self.expr(0)
        if self.i != len(self.toks):
            raise FormulaSyntaxError(f"Unexpected {self.toks[self.i].text!r}")
        return node

    def expr(self, min_bp: int) -> Node:
        left = self.prefix()
        while True:
            t = self.peek()
            if t is None or t.kind != OP:
                break
            if t.text == "%":
                if _POSTFIX_BP < min_bp:
                    break
                self.i += 1
                left = Unary("%", left)
                continue
            bp = _INFIX_BP.get(t.text)
            if bp is None or bp < min_bp:
                break
            self.i += 1
            right = self.expr(bp + 1)
            left = Binary(t.text, left, right)
        return left

    def prefix(self) -> Node:
        t = self.next()
        k = t.kind
        if k in (NUM, STR, BOOL, ERR):
            return Const(t.value)
        if k == REFT:
            return RefNode(t.ref)
        if k == NAMET:
            return NameNode(t.value, t.sheet)
        if k == OP and t.text in "+-":
            arg = self.expr(_PREFIX_BP)
            return Unary(t.text, arg)
        if k == LP:
            node = self.expr(0)
            self.expect(RP)
            return node
        if k == FUNC:
            return self.func(t)
        if k == LBRACE:
            return self.array()
        raise FormulaSyntaxError(f"Unexpected {t.text!r}")

    def func(self, t: Token) -> Node:
        args: list[Node] = []
        nt = self.peek()
        if nt is not None and nt.kind == RP:
            self.i += 1
            return Func(canonical_func_name(t.value), args, t.value)
        while True:
            nt = self.peek()
            if nt is None:
                raise FormulaSyntaxError("Missing ')'")
            if nt.kind in (COMMA, RP):
                args.append(Missing())
            else:
                args.append(self.expr(0))
            nt = self.next()
            if nt.kind == COMMA:
                continue
            if nt.kind == RP:
                break
            raise FormulaSyntaxError(f"Unexpected {nt.text!r} in function arguments")
        return Func(canonical_func_name(t.value), args, t.value)

    def array(self) -> Node:
        rows: list[list] = [[]]
        while True:
            t = self.next()
            sign = 1
            if t.kind == OP and t.text in "+-":
                sign = -1 if t.text == "-" else 1
                t = self.next()
            if t.kind == NUM:
                rows[-1].append(sign * t.value)
            elif t.kind in (STR, BOOL, ERR) and sign == 1:
                rows[-1].append(t.value)
            else:
                raise FormulaSyntaxError("Array constants may only contain values")
            t = self.next()
            if t.kind == COMMA:
                continue
            if t.kind == SEMI:
                rows.append([])
                continue
            if t.kind == RBRACE:
                break
            raise FormulaSyntaxError(f"Unexpected {t.text!r} in array constant")
        if len({len(r) for r in rows}) != 1:
            raise FormulaSyntaxError("Array constant rows must have the same length")
        return ArrayNode(rows)


def parse(formula: str) -> Node:
    """Parse a formula (with or without the leading '=')."""
    body = formula[1:] if formula.startswith("=") else formula
    return _Parser(tokenize(body)).parse()


# --------------------------------------------------------------------------- rewriting


def ref_text(info: RefInfo) -> str:
    """Render a RefInfo back to A1 text (keeping the original sheet prefix)."""
    p = info.sheet_raw
    if info.kind == "cell":
        return f"{p}{'$' if info.ac1 else ''}{col_letter(info.c1)}{'$' if info.ar1 else ''}{info.r1}"
    if info.kind == "range":
        return (f"{p}{'$' if info.ac1 else ''}{col_letter(info.c1)}{'$' if info.ar1 else ''}{info.r1}:"
                f"{'$' if info.ac2 else ''}{col_letter(info.c2)}{'$' if info.ar2 else ''}{info.r2}")
    if info.kind == "col":
        return f"{p}{'$' if info.ac1 else ''}{col_letter(info.c1)}:{'$' if info.ac2 else ''}{col_letter(info.c2)}"
    return f"{p}{'$' if info.ar1 else ''}{info.r1}:{'$' if info.ar2 else ''}{info.r2}"


def rewrite_tokens(formula: str, fn) -> str:
    """Rebuild formula text, replacing each token for which fn(token) returns a string."""
    if not formula.startswith("="):
        return formula
    body = formula[1:]
    try:
        toks = tokenize(body)
    except (FormulaSyntaxError, Unsupported):
        return formula
    out, pos, changed = [], 0, False
    for t in toks:
        new = fn(t)
        if new is not None and new != t.text:
            out.append(body[pos:t.start])
            out.append(new)
            pos = t.end
            changed = True
    if not changed:
        return formula
    out.append(body[pos:])
    return "=" + "".join(out)


def transform_refs(formula: str, fn) -> str:
    """fn(RefInfo) -> RefInfo | None(=#REF!) | ... ; return value identical object => unchanged."""

    def tok_fn(t: Token):
        if t.kind != REFT:
            return None
        res = fn(t.ref)
        if res is t.ref:
            return None
        if res is None:
            return (t.ref.sheet_raw or "") + "#REF!"
        return ref_text(res)

    return rewrite_tokens(formula, tok_fn)


def transform_refs_tokens(formula: str, fn):
    """Like transform_refs, but also returns the adjusted token list so the caller can
    build the new AST without tokenizing again. Returns (text, tokens or None)."""
    if not formula.startswith("="):
        return formula, None
    body = formula[1:]
    try:
        toks = tokenize(body)
    except (FormulaSyntaxError, Unsupported):
        return formula, None
    out, pos, changed = [], 0, False
    for t in toks:
        if t.kind != REFT:
            continue
        res = fn(t.ref)
        if res is t.ref:
            continue
        if res is None:
            new = (t.ref.sheet_raw or "") + "#REF!"
            t.kind, t.value, t.ref = ERR, REF, None
        else:
            new = ref_text(res)
            t.ref = res
        if new != t.text:
            out.append(body[pos:t.start])
            out.append(new)
            pos = t.end
            changed = True
    if not changed:
        return formula, toks
    out.append(body[pos:])
    return "=" + "".join(out), toks


def parse_tokens(toks: list[Token]) -> Node:
    return _Parser(toks).parse()


def translate(formula: str, dr: int, dc: int) -> str:
    """Shift relative references by (dr, dc), as when copying a formula."""
    if dr == 0 and dc == 0:
        return formula

    def fn(info: RefInfo):
        new = RefInfo(**info.__dict__)
        if info.kind in ("cell", "range", "row"):
            if not info.ar1:
                new.r1 += dr
            if not info.ar2:
                new.r2 += dr
        if info.kind in ("cell", "range", "col"):
            if not info.ac1:
                new.c1 += dc
            if not info.ac2:
                new.c2 += dc
        if info.kind == "cell":
            new.r2, new.c2 = new.r1, new.c1
        if min(new.r1, new.r2) < 1 or max(new.r1, new.r2) > MAX_ROW or min(new.c1, new.c2) < 1 or max(
            new.c1, new.c2
        ) > MAX_COL:
            return None
        return new

    return transform_refs(formula, fn)


# Functions that Excel stores with the _xlfn. prefix in the file format.
XLFN_FUNCS = {
    "XLOOKUP", "XMATCH", "IFS", "SWITCH", "CONCAT", "TEXTJOIN", "MAXIFS", "MINIFS", "IFNA", "XOR",
    "DAYS", "ISOWEEKNUM", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "CEILING.MATH", "FLOOR.MATH",
    "RANK.EQ", "RANK.AVG", "NETWORKDAYS.INTL", "WORKDAY.INTL", "AGGREGATE", "MODE.SNGL",
    "PERCENTILE.INC", "PERCENTILE.EXC", "QUARTILE.INC", "QUARTILE.EXC", "SEQUENCE", "UNIQUE",
    "SORT", "SORTBY", "FILTER", "LET", "LAMBDA", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
    "CEILING.PRECISE", "FLOOR.PRECISE", "BINOM.DIST",
}
_XLWS_FUNCS = {"FILTER", "SORT", "SORTBY"}


def normalize(formula: str) -> str:
    """Canonicalize user-typed formula text (uppercase functions and refs)."""

    def fn(t: Token):
        if t.kind == FUNC:
            return t.text.upper()
        if t.kind == REFT:
            return ref_text(t.ref)
        if t.kind == BOOL:
            return t.text.upper()
        if t.kind == ERR:
            return t.text.upper()
        return None

    return rewrite_tokens(formula, fn)


def to_storage(formula: str) -> str:
    """Add _xlfn. prefixes required by the xlsx format for newer functions."""

    def fn(t: Token):
        if t.kind != FUNC:
            return None
        name = t.text[:-1]
        canon = canonical_func_name(name)
        if canon in XLFN_FUNCS and canon == name.upper():
            prefix = "_xlfn._xlws." if canon in _XLWS_FUNCS else "_xlfn."
            return prefix + name.upper() + "("
        return None

    return rewrite_tokens(formula, fn)


def to_display(formula: str) -> str:
    """Strip _xlfn./_xlws. prefixes for display and editing."""
    if "_xl" not in formula.lower():
        return formula

    def fn(t: Token):
        if t.kind != FUNC:
            return None
        canon = canonical_func_name(t.text[:-1])
        return canon + "(" if canon != t.text[:-1].upper() else None

    return rewrite_tokens(formula, fn)


def is_error_literal(text: str) -> ExcelError | None:
    return ERROR_CODES.get(text.strip().upper())


__all__ = [
    "parse", "tokenize", "FormulaSyntaxError", "Unsupported", "translate", "transform_refs",
    "normalize", "to_storage", "to_display", "RefInfo", "ref_text", "REF",
]
