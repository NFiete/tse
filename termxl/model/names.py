"""Defined names (named ranges): list, validate, create, edit and delete."""
from __future__ import annotations

import re
from copy import copy
from dataclasses import dataclass

from openpyxl.workbook.defined_name import DefinedName

from ..formula.errors import Unsupported
from ..formula.parser import (
    NAMET, FormulaSyntaxError, RefInfo, normalize, parse, rewrite_tokens, to_display, to_storage, transform_refs,
)
from .refs import parse_addr, quote_sheet
from .undo import UndoEntry

_NAME_RE = re.compile(r"^[^\W\d][\w.\\?]*$|^\\[\w.\\?]*$")
_R1C1_RE = re.compile(r"^[Rr]\d*[Cc]\d*$|^[RrCc]$")


@dataclass
class NameInfo:
    name: str
    scope: object  # Sheet, or None for the whole workbook
    refers_to: str  # display form, with leading "="

    @property
    def scope_label(self) -> str:
        return "Workbook" if self.scope is None else self.scope.name


def _store(book, scope):
    return book.wb.defined_names if scope is None else scope.ws.defined_names


def _visible(name: str, dn) -> bool:
    return not name.startswith("_xlnm.") and not dn.hidden


def list_names(book) -> list[NameInfo]:
    """User-visible names, workbook-level first, sorted by name."""
    out = []
    for name, dn in book.wb.defined_names.items():
        if _visible(name, dn):
            out.append(NameInfo(name, None, to_display("=" + (dn.attr_text or ""))))
    for s in book.sheets:
        for name, dn in s.ws.defined_names.items():
            if _visible(name, dn):
                out.append(NameInfo(name, s, to_display("=" + (dn.attr_text or ""))))
    out.sort(key=lambda n: (n.name.lower(), n.scope is not None, n.scope_label.lower()))
    return out


def find(book, name: str, scope) -> str | None:
    """Actual stored spelling of a name in a scope (names are case-insensitive)."""
    low = name.lower()
    for key in _store(book, scope):
        if key.lower() == low:
            return key
    return None


def validate_name(book, name: str, scope, exclude: tuple | None = None) -> str | None:
    """Excel's naming rules. exclude=(name, scope) skips the name being edited."""
    if not name:
        return "Enter a name"
    if len(name) > 255:
        return "Names can be at most 255 characters"
    if " " in name:
        return "Names can't contain spaces (use _ or . instead)"
    if not _NAME_RE.match(name):
        return "Names must start with a letter, _ or \\ and contain only letters, digits, _ . \\ ?"
    if parse_addr(name) is not None or _R1C1_RE.match(name) or re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
        return "A name can't look like a cell reference"
    if name.upper() in ("TRUE", "FALSE"):
        return "TRUE and FALSE can't be used as names"
    existing = find(book, name, scope)
    if existing is not None:
        if exclude is not None and exclude[1] is scope and exclude[0].lower() == name.lower():
            return None
        where = "this workbook" if scope is None else f"sheet {scope.name}"
        return f"'{existing}' is already defined in {where}"
    return None


def normalize_refers_to(book, text: str, sheet) -> str:
    """Validate 'Refers to' text and return it as stored (no '=').

    Unqualified references are anchored to `sheet` and made absolute, as Excel's
    Name Manager does. Raises ValueError with a readable message."""
    text = text.strip()
    if not text:
        raise ValueError("Enter what the name refers to, e.g. Sheet1!$A$1:$B$10 or 0.25")
    if not text.startswith("="):
        text = "=" + text
    try:
        parse(text)
    except FormulaSyntaxError as e:
        raise ValueError(f"There's a problem with this reference: {e}") from None
    except Unsupported as e:
        raise ValueError(str(e)) from None
    raw = quote_sheet(sheet.name) + "!"

    def anchor(info: RefInfo):
        new = RefInfo(**info.__dict__)
        if new.sheet is None:
            new.sheet, new.sheet_raw = sheet.name, raw
        new.ar1 = new.ar2 = new.kind != "col"
        new.ac1 = new.ac2 = new.kind != "row"
        return new

    text = to_storage(normalize(transform_refs(text, anchor)))
    for s in _sheet_names_in(text):
        if book.sheet_by_name(s) is None:
            raise ValueError(f"There's no sheet named '{s}'")
    return text[1:]


def _sheet_names_in(formula: str) -> set[str]:
    from ..formula.parser import REFT, tokenize

    try:
        toks = tokenize(formula[1:])
    except (FormulaSyntaxError, Unsupported):
        return set()
    return {t.ref.sheet for t in toks if t.kind == REFT and t.ref.sheet}


def selection_ref(sheet, rect) -> str:
    """Absolute, sheet-qualified reference text for a selection (Refers to default)."""
    from .refs import col_letter

    a = f"${col_letter(rect.c1)}${rect.r1}"
    if rect.r1 == rect.r2 and rect.c1 == rect.c2:
        return f"{quote_sheet(sheet.name)}!{a}"
    return f"{quote_sheet(sheet.name)}!{a}:${col_letter(rect.c2)}${rect.r2}"


# --------------------------------------------------------------------------- changes


def _get(book, name, scope):
    key = find(book, name, scope)
    return copy(_store(book, scope)[key]) if key is not None else None


def _put(book, name, scope, dn):
    store = _store(book, scope)
    key = find(book, name, scope)
    if key is not None:
        del store[key]
    if dn is not None:
        store[dn.name] = copy(dn)


def _rename_in_formulas(book, old: str, new: str, scope) -> dict:
    """Point formulas that use `old` at `new`. Returns {(sheet, r, c): old_text}."""
    low = old.lower()
    changed = {}
    for s in book.sheets:
        local = find(book, old, s) is not None  # a sheet-level name shadows a workbook one
        for (r, c), cell in list(s.ws._cells.items()):
            if cell.data_type != "f" or not isinstance(cell._value, str) or low not in cell._value.lower():
                continue

            def fn(t, s=s, local=local):
                if t.kind != NAMET or t.value.lower() != low:
                    return None
                if t.sheet is not None:
                    target = book.sheet_by_name(t.sheet)
                    ok = target is scope if scope is not None else False
                else:
                    ok = (s is scope) if scope is not None else not local
                if not ok:
                    return None
                return t.text[: len(t.text) - len(t.value)] + new

            text = rewrite_tokens(cell._value, fn)
            if text != cell._value:
                changed[(s, r, c)] = cell._value
                cell._value = text
    return changed


def _refresh(book, names: set[str], extra_keys=()):
    """Reload names and recalculate formulas that mention any of them."""
    from .workbook import _formula_text

    book._load_names()
    lows = {n.lower() for n in names}
    dirty = set(extra_keys)
    for s in book.sheets:
        for rc, f in list(s.formulas.items()):
            text = f.text.lower()
            if (s, rc[0], rc[1]) in dirty or any(n in text for n in lows):
                dirty.add((s, rc[0], rc[1]))
    for s, r, c in list(dirty):
        cell = s.ws._cells.get((r, c))
        book._unregister(s, r, c)
        if cell is not None and cell.data_type == "f" and cell._value is not None:
            book._register(s, r, c, _formula_text(cell._value))
    if dirty:
        dirty |= book.graph.all_dependents(dirty)
        book._run(dirty)
    book.modified = True


def set_name(book, name: str, refers_to: str, scope, sheet, old: tuple | None = None) -> NameInfo:
    """Create a name, or edit one (old=(old_name, old_scope)); renaming updates formulas.

    `sheet` anchors unqualified references in refers_to. Raises ValueError."""
    name = name.strip()
    err = validate_name(book, name, scope, exclude=old)
    if err:
        raise ValueError(err)
    attr = normalize_refers_to(book, refers_to, sheet)
    keys = [(name, scope)]
    if old is not None:
        keys.append(old)
    before = {(n.lower(), id(sc)): (n, sc, _get(book, n, sc)) for n, sc in keys}

    renamed = {}
    if old is not None:
        _put(book, old[0], old[1], None)
        if old[1] is scope and old[0] != name:
            renamed = _rename_in_formulas(book, old[0], name, scope)
    dn = DefinedName(name, attr_text=attr)
    _put(book, name, scope, dn)
    after = {(n.lower(), id(sc)): (n, sc, _get(book, n, sc)) for n, sc in keys}
    involved = {name} | ({old[0]} if old else set())
    _refresh(book, involved, renamed.keys())
    new_texts = {k: k[0].ws._cells[(k[1], k[2])]._value for k in renamed}

    def apply(state, texts):
        for n, sc, d in state.values():
            _put(book, n, sc, d)
        for (s, r, c), text in texts.items():
            cell = s.ws._cells.get((r, c))
            if cell is not None:
                cell._value = text
        _refresh(book, involved, texts.keys())

    label = "Edit name" if old else "Define name"
    book.undo.push(UndoEntry(label, lambda: apply(before, renamed), lambda: apply(after, new_texts)))
    return NameInfo(name, scope, to_display("=" + attr))


def delete_name(book, name: str, scope):
    """Delete a name. Formulas that use it show #NAME? (as in Excel)."""
    before = _get(book, name, scope)
    if before is None:
        raise ValueError(f"No name '{name}'")
    _put(book, name, scope, None)
    _refresh(book, {name})

    def apply(dn):
        _put(book, name, scope, dn)
        _refresh(book, {name})

    book.undo.push(UndoEntry("Delete name", lambda: apply(before), lambda: apply(None)))


def resolve_rect(book, name: str, sheet):
    """(sheet, Rect) a name points at, if it's a plain reference; else None."""
    from ..formula.parser import RefNode
    from .refs import MAX_ROW, Rect

    ast = book.resolve_name(name, sheet)
    if isinstance(ast, RefNode):
        s = sheet if ast.sheet is None else book.sheet_by_name(ast.sheet)
        if s is not None:
            return s, Rect(ast.r1, ast.c1, min(ast.r2, MAX_ROW), ast.c2)
    return None
