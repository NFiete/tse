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
SREFT = "SREF"  # structured (table) reference such as Table1[Amount] or [@Qty]


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


# Special item specifiers of structured references, by their lower-case spelling.
SREF_ITEMS = {"#all": "all", "#data": "data", "#headers": "headers", "#totals": "totals",
              "#this row": "thisrow"}
_ITEM_TEXT = {"all": "#All", "data": "#Data", "headers": "#Headers", "totals": "#Totals",
              "thisrow": "#This Row"}
_ITEM_ORDER = ("all", "headers", "data", "totals", "thisrow")


@dataclass(frozen=True)
class StructRef:
    """A parsed structured reference. table=None means "the table containing the formula"."""

    table: str | None
    items: frozenset = frozenset()  # subset of SREF_ITEMS values
    col1: str | None = None  # None = every column
    col2: str | None = None  # last column of a [A]:[B] span (None for a single column)


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
_ERR_AT = re.compile(r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|SPILL!|CALC!|GETTING_DATA)")
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


def _match_bracket(src: str, i: int) -> int:
    """Index just past the ']' matching the '[' at src[i] ("'" escapes the next character)."""
    depth, n = 0, len(src)
    while i < n:
        ch = src[i]
        if ch == "'":
            i += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    raise FormulaSyntaxError("Missing ']' in table reference")


def _unescape(name: str) -> str:
    out, i = [], 0
    while i < len(name):
        if name[i] == "'" and i + 1 < len(name):
            i += 1
        out.append(name[i])
        i += 1
    return "".join(out)


def _sref_parts(body: str) -> list[str]:
    """Split "[a],[b]:[c]" into ["[a]", ",", "[b]", ":", "[c]"] (whitespace dropped)."""
    parts, i, n = [], 0, len(body)
    while i < n:
        ch = body[i]
        if ch in " \t\r\n":
            i += 1
        elif ch == "[":
            j = _match_bracket(body, i)
            parts.append(body[i:j])
            i = j
        elif ch in ",:":
            parts.append(ch)
            i += 1
        else:
            raise FormulaSyntaxError("Bad table reference")
    return parts


def _sref_columns(parts: list[str]) -> tuple[str | None, str | None]:
    """Column part of a structured reference: ["[A]"] or ["[A]", ":", "[B]"]."""
    if not parts:
        return None, None
    if len(parts) == 1:
        return _unescape(parts[0][1:-1]), None
    if len(parts) == 3 and parts[1] == ":":
        return _unescape(parts[0][1:-1]), _unescape(parts[2][1:-1])
    raise FormulaSyntaxError("Bad column range in table reference")


def _scan_sref(src: str, i: int, table: str | None) -> tuple[int, StructRef]:
    """Parse the bracketed part of a structured reference starting at src[i] == '['."""
    j = _match_bracket(src, i)
    body = src[i + 1:j - 1].strip()
    items: set[str] = set()
    cols: tuple = (None, None)
    if not body:
        pass
    elif body.startswith("@"):
        items.add("thisrow")
        rest = body[1:].strip()
        if rest.startswith("["):
            cols = _sref_columns(_sref_parts(rest))
        elif rest:
            cols = (_unescape(rest), None)
    elif body.startswith("#"):
        item = SREF_ITEMS.get(" ".join(body.lower().split()))
        if item is None:
            raise FormulaSyntaxError(f"Unknown table item {body}")
        items.add(item)
    elif body.startswith("["):
        parts = _sref_parts(body)
        groups: list[list[str]] = [[]]
        for p in parts:
            if p == ",":
                groups.append([])
            else:
                groups[-1].append(p)
        col_groups = []
        for g in groups:
            if len(g) == 1 and g[0][1:].lstrip().startswith("#"):
                item = SREF_ITEMS.get(" ".join(g[0][1:-1].lower().split()))
                if item is None:
                    raise FormulaSyntaxError(f"Unknown table item {g[0]}")
                items.add(item)
            elif g:
                col_groups.append(g)
            else:
                raise FormulaSyntaxError("Bad table reference")
        if len(col_groups) > 1:
            raise FormulaSyntaxError("Bad table reference")
        if col_groups:
            cols = _sref_columns(col_groups[0])
    else:
        cols = (_unescape(body), None)
    if "thisrow" in items and len(items) > 1:
        raise FormulaSyntaxError("#This Row can't be combined with other items")
    if len(items) > 1 and items not in ({"headers", "data"}, {"data", "totals"}):
        raise FormulaSyntaxError("Invalid combination of table items")
    return j, StructRef(table, frozenset(items), cols[0], cols[1])


_SPECIAL_COL_CHARS = set("\t\n\r,:.[]#'\"{}$^&*+=-<>/")


def _escape(name: str) -> str:
    return "".join("'" + ch if ch in "[]#'" else ch for ch in name)


def _col_text(name: str, bracket: bool) -> str:
    return f"[{_escape(name)}]" if bracket else _escape(name)


def sref_text(sr: StructRef, storage: bool = False, table: bool = True) -> str:
    """Render a StructRef. storage=True spells "@" as [#This Row] (as xlsx files require);
    table=False leaves out the table name."""
    name = (sr.table or "") if table else ""
    cols = None
    if sr.col1 is not None:
        cols = _col_text(sr.col1, True) + (":" + _col_text(sr.col2, True) if sr.col2 is not None else "")
    items = [i for i in _ITEM_ORDER if i in sr.items]
    if items == ["thisrow"] and not storage:
        if cols is None:
            return name + "[@]"
        if sr.col2 is None and not any(ch in _SPECIAL_COL_CHARS or ch == " " for ch in sr.col1):
            return f"{name}[@{_escape(sr.col1)}]"
        return f"{name}[@{cols}]"
    if not items:
        if cols is None:
            return name + "[]"
        if sr.col2 is None and not any(ch in _SPECIAL_COL_CHARS for ch in sr.col1):
            return f"{name}[{_escape(sr.col1)}]"
        return f"{name}[{cols}]"
    parts = [f"[{_ITEM_TEXT[i]}]" for i in items]
    if cols is None and len(parts) == 1:
        return f"{name}{parts[0]}"
    if cols is not None:
        parts.append(cols)
    return f"{name}[{','.join(parts)}]"



def tokenize(src: str) -> list[Token]:
    """Tokenize a formula body (without the leading '=')."""
    toks: list[Token] = []
    append = toks.append
    i, n = 0, len(src)
    match = _MASTER.match
    while i < n:
        if src[i] == "#" and toks and toks[-1].end == i and toks[-1].kind in (REFT, NAMET) and (
                toks[-1].kind == NAMET or toks[-1].ref.kind == "cell") and not _ERR_AT.match(src, i):
            append(Token(OP, "#", i, i + 1))  # spill reference: A1#
            i += 1
            continue
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
        elif kind == "ws":
            pass
        elif kind == "at":
            append(Token(OP, text, i, j))
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
                j, sref = _scan_sref(src, j - 1, m.group("name"))
                append(Token(SREFT, src[i:j], i, j, sref))
            else:
                append(Token(NAMET, text, i, j, text))
        elif kind == "bool":
            append(Token(BOOL, text, i, j, text.upper() == "TRUE"))
        elif kind == "err":
            append(Token(ERR, text, i, j, ERROR_CODES[text.upper()]))
        elif kind == "perr":
            append(Token(ERR, text, i, j, ERROR_CODES["#REF!"]))
        elif kind == "pname":
            sheet = _unquote_sheet(m.group("pfx2"))
            if "[" in sheet or src.startswith("[", j):
                raise Unsupported("External workbook references are not supported")
            append(Token(NAMET, text, i, j, m.group("pname"), sheet=sheet))
        elif kind == "bracket":
            j, sref = _scan_sref(src, i, None)
            if j < n and (src[j].isalnum() or src[j] in "_'\\"):
                raise Unsupported("External workbook references are not supported")  # [1]Sheet1!A1
            append(Token(SREFT, src[i:j], i, j, sref))
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


class StructRefNode(Node):
    __slots__ = ("ref",)

    def __init__(self, ref: StructRef):
        self.ref = ref

    def __repr__(self):
        return f"SRef({sref_text(self.ref)})"


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
_SPILL_BP = 90  # A1# binds tighter than the range operator

_FUNC_PREFIXES = ("_XLFN._XLWS.", "_XLFN.", "_XLWS.")


def canonical_func_name(raw: str) -> str:
    name = raw.upper()
    for p in _FUNC_PREFIXES:
        if name.startswith(p):
            return name[len(p):]
    return name


class LocalNode(Node):
    """A LET name or LAMBDA parameter (declaration or use). key is the upper-case name."""

    __slots__ = ("name", "key")

    def __init__(self, name: str):
        self.name, self.key = name, local_key(name)

    def __repr__(self):
        return f"Local({self.name})"


class Call(Node):
    """Calling a function value: LAMBDA(x,x+1)(2), or f(2) where f is a LET name."""

    __slots__ = ("callee", "args")

    def __init__(self, callee: Node, args: list[Node]):
        self.callee, self.args = callee, args

    def __repr__(self):
        return f"Call({self.callee!r}, {', '.join(map(repr, self.args))})"


def local_key(name: str) -> str:
    """Upper-case name without the _xlpm. prefix files use for LET names and LAMBDA parameters."""
    up = name.upper()
    return up[6:] if up.startswith("_XLPM.") else up


_MAX_PARAMS = 253


def _bind(node: Node, scope: frozenset) -> Node:
    """Resolve LET names and LAMBDA parameters: their declarations and uses become LocalNodes,
    calls of them become Call nodes. Raises FormulaSyntaxError for malformed LET/LAMBDA."""
    t = type(node)
    if t is NameNode:
        if node.sheet is None and local_key(node.name) in scope:
            return LocalNode(node.name)
        return node
    if t is Func:
        name = node.name
        if name == "LET":
            args = node.args
            if len(args) < 3 or len(args) % 2 == 0:
                raise FormulaSyntaxError("LET needs name, value pairs followed by a calculation")
            out = []
            for i in range(0, len(args) - 1, 2):
                decl = _declared(args[i])
                out.append(decl)
                out.append(_bind(args[i + 1], scope))
                scope = scope | {decl.key}
            out.append(_bind(args[-1], scope))
            return Func(name, out, node.raw_name)
        if name == "LAMBDA":
            args = node.args
            if not args or len(args) - 1 > _MAX_PARAMS:
                raise FormulaSyntaxError("LAMBDA needs a calculation as its last argument")
            params = [_declared(a) for a in args[:-1]]
            keys = {p.key for p in params}
            if len(keys) != len(params):
                raise FormulaSyntaxError("LAMBDA parameter names must be different")
            return Func(name, params + [_bind(args[-1], scope | keys)], node.raw_name)
        args = [_bind(a, scope) for a in node.args]
        if scope and local_key(node.raw_name) in scope and not node.raw_name.upper().startswith(("_XLFN.", "_XLWS.")):
            return Call(LocalNode(node.raw_name), args)
        node.args = args
        return node
    if t is Call:
        node.callee = _bind(node.callee, scope)
        node.args = [_bind(a, scope) for a in node.args]
        return node
    if t is Binary:
        node.left = _bind(node.left, scope)
        node.right = _bind(node.right, scope)
        return node
    if t is Unary:
        node.arg = _bind(node.arg, scope)
        return node
    return node


def _declared(node: Node) -> LocalNode:
    if type(node) is NameNode and node.sheet is None and "." not in local_key(node.name):
        return LocalNode(node.name)
    if type(node) is LocalNode:  # shadowing an outer name
        return node
    raise FormulaSyntaxError("LET names and LAMBDA parameters must be names such as x or rate")


class _Parser:
    def __init__(self, toks: list[Token]):
        self.toks = toks
        self.i = 0
        self.binders = False  # saw LET, LAMBDA or a call: names need resolving

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
        if self.binders:
            node = _bind(node, frozenset())
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
            if t.text == "#":
                if _SPILL_BP < min_bp:
                    break
                self.i += 1
                left = Func("ANCHORARRAY", [left], "#")
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
            name = t.value
            if t.sheet is None and name[:6].upper() == "_XLPM.":
                name = name[6:]
            return NameNode(name, t.sheet)
        if k == SREFT:
            return StructRefNode(t.value)
        if k == OP and t.text in "+-":
            arg = self.expr(_PREFIX_BP)
            return Unary(t.text, arg)
        if k == OP and t.text == "@":
            # implicit intersection; binds tighter than everything but ':' and '#'
            return Func("SINGLE", [self.expr(_INFIX_BP[":"])], "@")
        if k == LP:
            node = self.expr(0)
            self.expect(RP)
            return self.calls(node)
        if k == FUNC:
            return self.calls(self.func(t))
        if k == LBRACE:
            return self.array()
        raise FormulaSyntaxError(f"Unexpected {t.text!r}")

    def calls(self, node: Node) -> Node:
        """Postfix call syntax, as in LAMBDA(x,x+1)(2)."""
        while (nt := self.peek()) is not None and nt.kind == LP:
            self.i += 1
            self.binders = True
            node = Call(node, self.arguments())
        return node

    def func(self, t: Token) -> Node:
        raw = t.value
        if raw[:6].upper() == "_XLPM.":
            raw = raw[6:]
        name = canonical_func_name(raw)
        if name == "LET" or name == "LAMBDA":
            self.binders = True
        return Func(name, self.arguments(), raw)

    def arguments(self) -> list[Node]:
        """Arguments up to and including the closing ')' (the '(' is already consumed)."""
        args: list[Node] = []
        nt = self.peek()
        if nt is not None and nt.kind == RP:
            self.i += 1
            return args
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
        return args

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


def local_tokens(toks: list[Token]) -> set[int]:
    """Indexes of the name and function tokens that declare or use LET names and LAMBDA
    parameters (the ones files spell with _xlpm.)."""
    found: set[int] = set()
    if not any(t.kind == FUNC and canonical_func_name(t.value) in ("LET", "LAMBDA") for t in toks):
        return found
    # One frame per open bracket: [function name or None, argument index, declared names]
    stack: list[list] = []
    visible: set[str] = set()
    n = len(toks)
    for i, t in enumerate(toks):
        k = t.kind
        if k == FUNC:
            if t.value[:6].upper() != "_XLPM." and local_key(t.value) not in visible:
                stack.append([canonical_func_name(t.value), 0, set()])
                continue
            found.add(i)
            stack.append([None, 0, set()])
        elif k in (LP, LBRACE):
            stack.append([None, 0, set()])
        elif k in (RP, RBRACE):
            if stack:
                stack.pop()
                visible = set().union(*(f[2] for f in stack))
        elif k == COMMA:
            if stack:
                stack[-1][1] += 1
        elif k == NAMET and t.sheet is None:
            key = local_key(t.value)
            frame = stack[-1] if stack else None
            whole_arg = (i > 0 and toks[i - 1].kind in (FUNC, COMMA) and i + 1 < n and toks[i + 1].kind == COMMA)
            if frame is not None and whole_arg and (
                    (frame[0] == "LET" and frame[1] % 2 == 0) or frame[0] == "LAMBDA"):
                frame[2].add(key)
                visible.add(key)
                found.add(i)
            elif key in visible or t.value[:6].upper() == "_XLPM.":
                found.add(i)
    return found


def rewrite_tokens(formula: str, fn, skip_locals: bool = False) -> str:
    """Rebuild formula text, replacing each token for which fn(token) returns a string.
    skip_locals=True leaves LET names and LAMBDA parameters alone."""
    if not formula.startswith("="):
        return formula
    body = formula[1:]
    try:
        toks = tokenize(body)
    except (FormulaSyntaxError, Unsupported):
        return formula
    skip = local_tokens(toks) if skip_locals else ()
    out, pos, changed = [], 0, False
    for i, t in enumerate(toks):
        if i in skip:
            continue
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
    "CEILING.PRECISE", "FLOOR.PRECISE", "BINOM.DIST", "ANCHORARRAY", "SINGLE",
    # dynamic-array and lookup functions
    "TAKE", "DROP", "VSTACK", "HSTACK", "TOCOL", "TOROW", "CHOOSEROWS", "CHOOSECOLS", "EXPAND", "WRAPROWS",
    "WRAPCOLS", "FORMULATEXT", "GROUPBY", "PIVOTBY", "TRIMRANGE", "PERCENTOF",
    # LAMBDA helpers
    "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY", "ISOMITTED",
    # statistical functions added in Excel 2010-2016
    "BETA.DIST", "BETA.INV", "BINOM.DIST.RANGE", "BINOM.INV", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV",
    "CHISQ.INV.RT", "CHISQ.TEST", "CONFIDENCE.NORM", "CONFIDENCE.T", "COVARIANCE.P", "COVARIANCE.S",
    "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FORECAST.LINEAR", "GAMMA",
    "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE", "GAUSS", "HYPGEOM.DIST", "LOGNORM.DIST", "LOGNORM.INV",
    "MODE.MULT", "NEGBINOM.DIST", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "PERCENTRANK.EXC",
    "PERCENTRANK.INC", "PERMUTATIONA", "PHI", "POISSON.DIST", "SKEW.P", "T.DIST", "T.DIST.2T", "T.DIST.RT",
    "T.INV", "T.INV.2T", "T.TEST", "WEIBULL.DIST", "Z.TEST",
}
_XLWS_FUNCS = {"FILTER", "SORT"}
# Functions whose arguments may be functions (eta-reduced LAMBDAs such as SUM in
# GROUPBY(A1:A9,B1:B9,SUM)); files spell those arguments _xleta.SUM.
ETA_ARG_FUNCS = {"GROUPBY", "PIVOTBY", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY"}


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

    return rewrite_tokens(formula, fn, skip_locals=True)


def _operand_end(toks: list[Token], i: int) -> int:
    """Index of the last token of the operand starting at toks[i]: a reference, name, function
    call or bracketed expression, with any following ':' and '#' operators."""
    n = len(toks)
    j = i
    while j < n and toks[j].kind == OP and toks[j].text == "@":
        j += 1
    if j >= n:
        return n - 1
    if toks[j].kind in (FUNC, LP, LBRACE):
        close = _match_close(toks, j)
        j = n - 1 if close is None else close
        while j + 1 < n and toks[j + 1].kind == LP:  # LAMBDA(x,x)(1)
            close = _match_close(toks, j + 1)
            j = n - 1 if close is None else close
    if j + 1 < n and toks[j + 1].kind == OP and toks[j + 1].text == "#":
        j += 1
    if j + 2 < n and toks[j + 1].kind == OP and toks[j + 1].text == ":":
        return _operand_end(toks, j + 2)
    return j


def _match_close(toks: list[Token], i: int) -> int | None:
    """Index of the ')' or '}' closing the bracket opened by toks[i]."""
    depth = 0
    for j in range(i, len(toks)):
        k = toks[j].kind
        if k in (FUNC, LP, LBRACE):
            depth += 1
        elif k in (RP, RBRACE):
            depth -= 1
            if depth == 0:
                return j
    return None


def _rebuild(body: str, toks: list[Token], repl: dict) -> str:
    out, pos = [], 0
    for i in sorted(repl):
        t = toks[i]
        out.append(body[pos:t.start])
        out.append(repl[i])
        pos = t.end
    out.append(body[pos:])
    return "".join(out)


def to_storage(formula: str) -> str:
    """Add _xlfn. prefixes required by the xlsx format for newer functions, _xleta. to functions
    passed as arguments, spell [@Col] table references as [[#This Row],[Col]], A1# as
    _xlfn.ANCHORARRAY(A1) and @x as _xlfn.SINGLE(x)."""
    if not formula.startswith("="):
        return formula
    body = formula[1:]
    try:
        toks = tokenize(body)
    except (FormulaSyntaxError, Unsupported):
        return formula
    stack: list = []  # enclosing function names (None for plain parentheses and array constants)
    repl: dict[int, str] = {}
    local = local_tokens(toks)
    for i, t in enumerate(toks):
        new = None
        if i in local:
            if t.kind == FUNC:
                stack.append(None)
            if t.text[:6].upper() != "_XLPM.":
                new = "_xlpm." + t.text
        elif t.kind == SREFT:
            new = sref_text(t.value, storage=True) if "thisrow" in t.value.items else None
        elif t.kind in (LP, LBRACE):
            stack.append(None)
        elif t.kind in (RP, RBRACE):
            if stack:
                stack.pop()
        elif t.kind == NAMET:
            if any(f in ETA_ARG_FUNCS for f in stack) and "!" not in t.text and _is_function_name(t.text):
                new = "_xleta." + canonical_func_name(t.text)
        elif t.kind == FUNC:
            name = t.text[:-1]
            canon = canonical_func_name(name)
            stack.append(canon)
            if canon in XLFN_FUNCS and canon == name.upper():
                prefix = "_xlfn._xlws." if canon in _XLWS_FUNCS else "_xlfn."
                new = prefix + name.upper() + "("
        if new is not None and new != t.text:
            repl[i] = new
    for i, t in enumerate(toks):
        if t.kind == OP and t.text == "#" and i > 0:
            repl[i - 1] = "_xlfn.ANCHORARRAY(" + repl.get(i - 1, toks[i - 1].text)
            repl[i] = ")"
    for i in range(len(toks) - 1, -1, -1):
        t = toks[i]
        if t.kind == OP and t.text == "@" and i + 1 < len(toks):
            j = _operand_end(toks, i + 1)
            repl[i] = "_xlfn.SINGLE("
            repl[j] = repl.get(j, toks[j].text) + ")"
    if not repl:
        return formula
    return "=" + _rebuild(body, toks, repl)


def _is_function_name(name: str) -> bool:
    from .functions import registered

    up = name.upper()
    return not up.startswith(("_XLETA.", "_XLFN.", "_XLWS.")) and up in registered()


def to_display(formula: str) -> str:
    """Strip _xlfn./_xlws. prefixes for display and editing, show [#This Row] as @,
    _xlfn.ANCHORARRAY(A1) as A1# and _xlfn.SINGLE(x) as @x."""
    low = formula.lower()
    if "_xl" not in low and "#this row" not in low:
        return formula
    if not formula.startswith("="):
        return formula
    body = formula[1:]
    try:
        toks = tokenize(body)
    except (FormulaSyntaxError, Unsupported):
        return formula
    repl: dict[int, str] = {}
    n = len(toks)
    for i, t in enumerate(toks):
        new = None
        if t.kind == SREFT:
            new = sref_text(t.value) if "thisrow" in t.value.items else None
        elif t.kind == NAMET:
            pfx = t.text[:7].upper()
            new = t.text[7:] if pfx == "_XLETA." else t.text[6:] if pfx[:6] == "_XLPM." else None
        elif t.kind == FUNC and t.text[:6].upper() == "_XLPM.":
            new = t.text[6:]
        elif t.kind == FUNC:
            canon = canonical_func_name(t.text[:-1])
            new = canon + "(" if canon != t.text[:-1].upper() else None
        if new is not None:
            repl[i] = new
    for i, t in enumerate(toks):
        if t.kind != FUNC:
            continue
        canon = canonical_func_name(t.text[:-1])
        if canon == "ANCHORARRAY" and i + 2 < n and toks[i + 1].kind in (REFT, NAMET) \
                and toks[i + 2].kind == RP and (toks[i + 1].kind == NAMET or toks[i + 1].ref.kind == "cell"):
            repl[i] = ""
            repl[i + 2] = "#"
        elif canon == "SINGLE" and t.text[:-1].upper() != "SINGLE":
            j = _match_close(toks, i)
            if j is None or j == i + 1:
                continue
            if _operand_end(toks, i + 1) == j - 1:
                repl[i] = "@"
                repl[j] = ""
            else:
                repl[i] = "@("
    if not repl:
        return formula
    return "=" + _rebuild(body, toks, repl)


def is_error_literal(text: str) -> ExcelError | None:
    return ERROR_CODES.get(text.strip().upper())


__all__ = [
    "parse", "tokenize", "LocalNode", "Call", "local_tokens", "StructRef", "StructRefNode", "sref_text", "SREFT", "FormulaSyntaxError", "Unsupported", "translate", "transform_refs",
    "normalize", "to_storage", "to_display", "RefInfo", "ref_text", "REF",
]
