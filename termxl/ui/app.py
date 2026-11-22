"""The term_xl Textual application."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass

from rich.cells import cell_len
from textual import events, on, work
from textual.app import App, ComposeResult, SystemCommand
from textual.binding import Binding
from textual.containers import Horizontal
from textual.widgets import Footer, Input, Static

from ..format.numfmt import format_value
from ..formula.errors import ExcelError
from ..formula.parser import REFT, SREFT, FormulaSyntaxError, Unsupported, tokenize
from ..formula.registry import wildcard_regex
from ..io import xlsx
from ..model import names, outline, structure, tables
from ..model.refs import MAX_COL, MAX_ROW, Rect, addr, col_letter, quote_sheet
from .dialogs import (
    ChoiceDialog, FindDialog, HelpScreen, InputDialog, NameEditDialog, NameManagerDialog, TableDialog,
)
from .grid import Grid
from .theme import DARK, LIGHT
from .widgets import FormulaHighlighter, FormulaInput, NameBox, SheetTabs, StatusBar, cycle_absolute

_ARROWS = {"up": (-1, 0), "down": (1, 0), "left": (0, -1), "right": (0, 1)}
_POINTABLE_BEFORE = set("=(,+-*/^&<>:;")

NUMBER_FORMATS = [
    ("General", "General"),
    ("Number (1,234.56)", "#,##0.00"),
    ("Integer (1,235)", "#,##0"),
    ("Currency ($1,234.56)", '"$"#,##0.00'),
    ("Accounting", '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'),
    ("Percent (12%)", "0%"),
    ("Percent (12.34%)", "0.00%"),
    ("Scientific (1.23E+04)", "0.00E+00"),
    ("Fraction (1 3/4)", "# ?/?"),
    ("Short date (11/21/2026)", "m/d/yyyy"),
    ("ISO date (2026-11-21)", "yyyy-mm-dd"),
    ("Long date (Saturday, November 21, 2026)", "dddd, mmmm d, yyyy"),
    ("Time (1:30:00 PM)", "h:mm:ss AM/PM"),
    ("Date and time", "m/d/yyyy h:mm"),
    ("Text (@)", "@"),
]

COLORS = {
    "none": None, "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "00B050",
    "blue": "0070C0", "yellow": "FFFF00", "orange": "FFC000", "purple": "7030A0", "gray": "808080",
    "light gray": "D9D9D9", "light blue": "DDEBF7", "light green": "E2EFDA", "light yellow": "FFF2CC",
    "light red": "FCE4D6", "dark blue": "1F4E78", "dark green": "375623",
}


@dataclass
class EditState:
    sheet: object
    r: int
    c: int
    mode: str  # 'enter' | 'edit' | 'point'
    original: str


class TermXLApp(App):
    TITLE = "term_xl"
    CSS = """
    Screen { layout: vertical; }
    #bar { height: 1; width: 1fr; }
    NameBox { width: 16; height: 1; border: none; padding: 0 1; background: $panel; color: $text;
              text-style: bold; }
    NameBox:focus { border: none; background: $boost; }
    #fx { width: 4; height: 1; color: $text-muted; padding: 0 1; }
    FormulaInput { height: 1; border: none; padding: 0 1; width: 1fr; background: $surface; }
    FormulaInput:focus { border: none; background: $boost; }
    Footer { height: 1; }
    """
    ENABLE_COMMAND_PALETTE = True

    BINDINGS = [
        Binding("f1", "help", "Help"),
        Binding("ctrl+s", "save", "Save"),
        Binding("f12", "save_as", "Save as", show=False),
        Binding("ctrl+o", "open", "Open"),
        Binding("ctrl+n", "new", "New", show=False),
        Binding("ctrl+z", "undo", "Undo"),
        Binding("ctrl+y", "redo", "Redo", show=False),
        Binding("ctrl+c", "copy", "Copy", show=False),
        Binding("ctrl+x", "cut", "Cut", show=False),
        Binding("ctrl+v", "paste", "Paste", show=False),
        Binding("ctrl+b,alt+b", "font('bold')", "Bold", show=False),
        Binding("ctrl+i,alt+i", "font('italic')", "Italic", show=False),
        Binding("ctrl+u,alt+u", "font('underline')", "Underline", show=False),
        Binding("ctrl+d", "fill('down')", "Fill down", show=False),
        Binding("ctrl+r", "fill('right')", "Fill right", show=False),
        Binding("ctrl+f", "find", "Find"),
        Binding("ctrl+h", "replace", "Replace", show=False),
        Binding("ctrl+g", "goto", "Go to"),
        Binding("ctrl+pageup", "sheet(-1)", show=False),
        Binding("ctrl+pagedown", "sheet(1)", show=False),
        Binding("f9", "recalc", show=False),
        Binding("ctrl+f3", "name_manager", "Names", show=False),
        Binding("ctrl+t", "table", "Table", show=False),
        Binding("alt+equals_sign,alt+plus,ctrl+plus,ctrl+equals_sign", "insert", "Insert", show=False),
        Binding("alt+minus,ctrl+minus", "delete", "Delete", show=False),
        Binding("alt+left", "col_width(-1)", show=False),
        Binding("alt+right", "col_width(1)", show=False),
        Binding("alt+shift+right", "group", show=False),
        Binding("alt+shift+left", "ungroup", show=False),
        Binding("alt+shift+up", "outline(True)", show=False),
        Binding("alt+shift+down", "outline(False)", show=False),
        Binding("escape", "escape", show=False),
        Binding("ctrl+q", "quit_app", "Quit", priority=True),
    ]

    def __init__(self, path: str | None = None, sheet: str | None = None):
        super().__init__()
        self.path_arg = path
        self.sheet_arg = sheet
        self.book = None
        self.editing: EditState | None = None
        self.point: dict | None = None
        self.clip: dict | None = None
        self.clip_cut = False
        self.clip_text: str | None = None
        self.message = ""
        self.ack_warnings: set = set()
        self._find_state = None

    # ------------------------------------------------------------------ layout

    def compose(self) -> ComposeResult:
        with Horizontal(id="bar"):
            yield NameBox(value="A1", id="namebox", select_on_focus=True)
            yield Static("fx", id="fx")
            yield FormulaInput(id="formula", highlighter=FormulaHighlighter(self), select_on_focus=False)
        yield Grid(id="grid")
        yield SheetTabs(id="tabs")
        yield StatusBar(id="status")
        yield Footer(compact=True, show_command_palette=True)

    @property
    def grid(self) -> Grid:
        return self.query_one("#grid", Grid)

    @property
    def formula(self) -> FormulaInput:
        return self.query_one("#formula", FormulaInput)

    def on_mount(self):
        self.theme = "textual-dark"
        path = self.path_arg
        if path and os.path.exists(path):
            self.message = f"Opening {os.path.basename(path)}…"
            self._set_book(xlsx.new_workbook(), quiet=True)
            self.load_file(path)
        else:
            book = xlsx.new_workbook()
            if path:
                book.path = os.path.abspath(path)
                self.message = f"New file — Ctrl+S will create {os.path.basename(path)}"
            self._set_book(book)

    @work(thread=True, exclusive=True)
    def load_file(self, path: str):
        try:
            book = xlsx.load(path)
        except Exception as e:  # noqa: BLE001
            self.call_from_thread(self._load_failed, path, e)
            return
        self.call_from_thread(self._loaded, book)

    def _load_failed(self, path, e):
        self.message = f"Could not open {path}"
        self.notify(f"Could not open {path}:\n{e}", severity="error", timeout=10)
        self.refresh_status()

    def _loaded(self, book):
        sheet = None
        if self.sheet_arg:
            sheet = book.sheet_by_name(self.sheet_arg)
        if sheet is None:
            try:
                active_ws = book.wb.active
                sheet = next((s for s in book.sheets if s.ws is active_ws), None)
            except Exception:
                sheet = None
        self._set_book(book, sheet)
        n = sum(len(s.formulas) for s in book.sheets)
        self.message = f"Opened {os.path.basename(book.path)}" + (f" — {n:,} formulas" if n else "")
        unsupported = sorted({name for s in book.sheets for f in s.formulas.values() for name in f.unsupported})
        if unsupported:
            self.notify("Not supported, so their cells keep the values saved by Excel: "
                        + ", ".join(unsupported[:12]) + ("…" if len(unsupported) > 12 else ""),
                        severity="warning", timeout=8)
        if book.warnings:
            self.notify("This file contains things term_xl can't save back: " + "; ".join(book.warnings)
                        + ".\nYou'll be asked before overwriting it.", severity="warning", timeout=10)
        self.refresh_status()

    def _set_book(self, book, sheet=None, quiet=False):
        self.book = book
        self.clip = None
        self.grid.clip_rect = None
        self.grid.set_book(book, sheet)
        self.query_one("#tabs", SheetTabs).refresh()
        self.grid.focus()
        self.grid_changed()
        name = os.path.basename(book.path) if book.path else "untitled"
        self.sub_title = name

    # ------------------------------------------------------------------ status

    def grid_changed(self):
        if self.book is None:
            return
        g = self.grid
        sel = g.selection
        nb = self.query_one("#namebox", NameBox)
        if not nb.has_focus:
            if sel.rows * sel.cols > 1 and g._dragging:
                label = f"{min(sel.rows, MAX_ROW)}R × {sel.cols}C"
            else:
                label = self._name_for_selection() or addr(*g.active)
            with nb.prevent(Input.Changed):
                nb.value = label
                nb.cursor_position = 0
        if self.editing is None:
            inp = self.formula
            text = g.sheet.input_text(*g.active)
            with inp.prevent(Input.Changed):
                inp.value = text
                inp.cursor_position = 0
        self.refresh_status()

    def _selection_stats(self) -> str:
        g = self.grid
        sel = g.selection
        if sel.rows * sel.cols <= 1:
            return ""
        sheet = g.sheet
        book = self.book
        rect = book.clip(sheet, sel)
        cells = sheet.ws._cells
        if rect.rows * rect.cols > len(cells):
            keys = [k for k in cells if rect.contains(*k)]
            keys += [(r, c) for r, c, _ in sheet.spill.cells_in(rect.r1, rect.c1, rect.r2, rect.c2)]
        else:
            keys = list(rect.cells())
        if len(keys) > 500_000:
            return "Large selection"
        nums = []
        count = 0
        fmt = None
        for r, c in keys:
            v = book.get_value(sheet, r, c)
            if v is None or v == "":
                continue
            count += 1
            if type(v) in (int, float):
                nums.append(v)
                if fmt is None:
                    fmt = sheet.number_format(r, c)
        if count == 0:
            return ""
        parts = []
        if nums:
            fmt = fmt if fmt and fmt != "General" and "%" not in fmt else "General"

            def f(x):
                return format_value(x, fmt, book.epoch1904)[0].strip() if fmt != "General" else \
                    format_value(round(x, 10) if isinstance(x, float) else x, "General")[0]

            parts.append(f"Average: {f(sum(nums) / len(nums))}")
        parts.append(f"Count: {count}")
        if nums:
            parts.append(f"Sum: {f(sum(nums))}")
        return "  ".join(parts)

    def refresh_status(self):
        if self.book is None:
            return
        g = self.grid
        mode = "READY"
        if self.editing is not None:
            mode = "POINT" if self.point else self.editing.mode.upper()
        msg = self.message
        if self.editing is None and not msg:
            r, c = g.active
            sheet = g.sheet
            f = sheet.formulas.get((r, c))
            cell = sheet.ws._cells.get((r, c))
            spill_from = sheet.spill_anchor(r, c)
            if (r, c) in sheet.stale and f is not None:
                what = ", ".join(f.unsupported) if f.unsupported else "this formula"
                msg = f"Showing the value saved by Excel ({what} not supported here)"
            elif f is not None and (r, c) in sheet.spill.blocked:
                nr, nc = sheet.spill.blocked[(r, c)]
                where = Rect(r, c, r + nr - 1, c + nc - 1)
                msg = (f"Can't spill: the result ({nr}×{nc}) would go past the edge of the sheet"
                       if where.r2 > MAX_ROW or where.c2 > MAX_COL else
                       f"Can't spill: {where.a1()} isn't empty")
            elif spill_from is not None:
                anchor = sheet.formulas.get(spill_from)
                msg = f"Spilled from {addr(*spill_from)}: {anchor.text if anchor else ''}"
            elif f is not None and f.error:
                msg = f"Formula problem: {f.error}"
            elif cell is not None and cell.comment is not None:
                msg = "Note: " + cell.comment.text.replace("\n", " ")[:200]
            elif cell is not None and cell.hyperlink is not None:
                msg = "Link: " + (cell.hyperlink.target or cell.hyperlink.location or "")
            else:
                hit = tables.table_at(self.book, sheet, r, c)
                if hit is not None:
                    names_ = hit.col_names()
                    i = c - hit.g.c1
                    part = ("header" if hit.g.hdr and r == hit.g.r1 else
                            "totals row" if hit.g.tot and r == hit.g.r2 else "")
                    msg = f"Table {hit.name}, column {names_[i] if i < len(names_) else '?'}" + \
                          (f" ({part})" if part else "") + "  (Ctrl+T: table settings)"
        book = self.book
        name = os.path.basename(book.path) if book.path else "untitled"
        label = ("● " if book.modified else "") + name
        self.query_one("#status", StatusBar).show(mode, msg, self._selection_stats(), label)

    def set_message(self, msg: str):
        self.message = msg
        self.refresh_status()

    def after_change(self, layout: bool = False):
        g = self.grid
        if layout:
            g.invalidate()
        g.refresh()
        self.query_one("#tabs", SheetTabs).refresh()
        self.grid_changed()

    # ------------------------------------------------------------------ editing

    def start_edit(self, initial: str | None = None, mode: str = "enter", focus: bool = True):
        if self.editing is not None or self.book is None:
            return
        g = self.grid
        sheet = g.sheet
        r, c = g.active
        original = sheet.input_text(r, c)
        text = original if initial is None else initial
        self.editing = EditState(sheet, r, c, mode, original)
        self.point = None
        self.message = ""
        inp = self.formula
        with inp.prevent(Input.Changed):
            inp.value = text
        inp.cursor_position = len(text)
        g.edit_cell = (sheet, r, c)
        g.edit_text = text
        self._update_ref_highlights(text)
        if focus:
            inp.focus()
        g.refresh()
        self.refresh_status()

    @on(Input.Changed, "#formula")
    def _formula_changed(self, event: Input.Changed):
        if self.editing is None:
            return
        g = self.grid
        g.edit_text = event.value
        self._update_ref_highlights(event.value)
        g.refresh()

    def _update_ref_highlights(self, text: str):
        g = self.grid
        hl = []
        if text.startswith("=") and self.editing is not None:
            try:
                toks = tokenize(text[1:])
            except (FormulaSyntaxError, Unsupported):
                toks = []
            colors = g.palette.ref_colors
            i = 0
            e = self.editing
            for t in toks:
                if t.kind == REFT:
                    info = t.ref
                    s = e.sheet if info.sheet is None else self.book.sheet_by_name(info.sheet)
                    if s is not None:
                        r2 = min(info.r2, MAX_ROW)
                        hl.append((s, Rect.of(info.r1, info.c1, r2, info.c2), colors[i % len(colors)]))
                    i += 1
                elif t.kind == SREFT:
                    res = tables.resolve(self.book, t.value, e.sheet, e.r, e.c)
                    if isinstance(res, tuple):
                        s, r1, c1, r2, c2 = res
                        hl.append((s, Rect.of(r1, c1, r2, c2), colors[i % len(colors)]))
                    i += 1
        g.ref_highlights = hl

    def can_point(self) -> bool:
        e = self.editing
        if e is None:
            return False
        inp = self.formula
        text = inp.value
        if not text.startswith("="):
            return False
        if self.point is not None:
            return True
        if e.mode == "edit":
            return False
        before = text[:inp.cursor_position].rstrip()
        return bool(before) and before[-1] in _POINTABLE_BEFORE

    def end_point(self):
        if self.point is not None:
            self.point = None
            self.refresh_status()

    def _apply_point(self):
        p = self.point
        inp = self.formula
        e = self.editing
        rect = Rect.of(*p["anchor"], *p["corner"])
        ref = rect.a1()
        if p["sheet"] is not e.sheet:
            ref = quote_sheet(p["sheet"].name) + "!" + ref
        text = inp.value
        new = text[:p["start"]] + ref + text[p["end"]:]
        p["end"] = p["start"] + len(ref)
        inp.value = new
        inp.cursor_position = p["end"]
        g = self.grid
        g.ensure_visible(*p["corner"])
        g.refresh()
        self.refresh_status()

    def _start_point(self, base):
        inp = self.formula
        pos = inp.cursor_position
        self.point = {"start": pos, "end": pos, "anchor": base, "corner": base, "sheet": self.grid.sheet}

    def point_move(self, dr: int, dc: int, extend: bool):
        e = self.editing
        g = self.grid
        if self.point is None:
            base = (e.r, e.c) if g.sheet is e.sheet else g.active
            self._start_point(base)
            if not extend:
                self.point["anchor"] = self.point["corner"] = g._step(*base, dr, dc)
            else:
                self.point["corner"] = g._step(*base, dr, dc)
        else:
            p = self.point
            if extend:
                p["corner"] = g._step(*p["corner"], dr, dc)
            else:
                p["anchor"] = p["corner"] = g._step(*p["corner"], dr, dc)
        self._apply_point()

    def point_at(self, r: int, c: int, extend: bool = False):
        if self.point is None:
            self._start_point((r, c))
        p = self.point
        p["sheet"] = self.grid.sheet
        if extend:
            p["corner"] = (r, c)
        else:
            p["anchor"] = p["corner"] = (r, c)
        self._apply_point()
        self.formula.focus()

    def edit_key(self, event: events.Key) -> bool:
        """Handle a key while editing. Returns True if consumed."""
        e = self.editing
        if e is None:
            return False
        key = event.key
        if key in ("enter", "ctrl+j", "ctrl+m"):
            self.commit_edit((1, 0))
            return True
        if key == "shift+enter":
            self.commit_edit((-1, 0))
            return True
        if key == "tab":
            self.commit_edit((0, 1))
            return True
        if key == "shift+tab":
            self.commit_edit((0, -1))
            return True
        if key == "escape":
            self.cancel_edit()
            return True
        if key == "f2":
            e.mode = "edit" if e.mode != "edit" else "enter"
            self.point = None
            self.refresh_status()
            return True
        if key == "f4":
            inp = self.formula
            text, pos = cycle_absolute(inp.value, inp.cursor_position)
            inp.value = text
            inp.cursor_position = pos
            return True
        if key in ("ctrl+pageup", "ctrl+pagedown") and inp_is_formula(self.formula.value):
            d = -1 if key == "ctrl+pageup" else 1
            self._switch_sheet_relative(d)
            if self.point is not None:
                self.point["sheet"] = self.grid.sheet
            return True
        base = key.replace("shift+", "")
        if base in _ARROWS:
            dr, dc = _ARROWS[base]
            extend = key.startswith("shift+")
            if self.can_point():
                self.point_move(dr, dc, extend)
                return True
            if e.mode == "enter":
                self.commit_edit((dr, dc), arrow=True)
                return True
            if base in ("up", "down"):
                return True
        return False

    def commit_edit(self, move=(1, 0), arrow: bool = False) -> bool:
        e = self.editing
        if e is None:
            return True
        text = self.formula.value
        if text.startswith("="):
            text = _autoclose(text)
        g = self.grid
        if text != e.original:
            try:
                self.book.set_input(e.sheet, e.r, e.c, text)
            except FormulaSyntaxError as err:
                self.notify(f"There's a problem with this formula: {err}", severity="error", timeout=6,
                            markup=False)
                return False
            except ValueError as err:
                self.notify(str(err), severity="error", timeout=6, markup=False)
                return False
        self._end_edit()
        if g.sheet is not e.sheet:
            g.show_sheet(e.sheet)
        g.active = g.anchor = g.corner = (e.r, e.c)
        mr = e.sheet.merged_at(e.r, e.c)
        if mr is not None:
            g.corner = (mr.max_row, mr.max_col)
        if move:
            dr, dc = move
            if arrow:
                g.action_move(dr, dc)
            elif dr:
                g.action_enter_move(dr)
            else:
                g.action_tab_move(dc)
        self.after_change()
        return True

    def cancel_edit(self):
        e = self.editing
        if e is None:
            return
        self._end_edit()
        g = self.grid
        if g.sheet is not e.sheet:
            g.show_sheet(e.sheet)
        self.after_change()

    def _end_edit(self):
        g = self.grid
        self.editing = None
        self.point = None
        g.edit_cell = None
        g.ref_highlights = []
        g.focus()

    # ------------------------------------------------------------------ actions: clipboard

    def _need_grid(self) -> bool:
        """Actions that operate on cells are ignored while typing in the editor."""
        return self.editing is None and self.book is not None

    def action_copy(self, cut: bool = False):
        if self.editing is not None:
            self.formula.action_copy() if not cut else self.formula.action_cut()
            return
        g = self.grid
        sel = g.selection
        self.clip = self.book.capture(g.sheet, sel)
        self.clip_cut = cut
        g.clip_rect = (g.sheet, self.clip["rect"])
        rect = self.clip["rect"]
        if rect.rows * rect.cols <= 200_000:
            self.clip_text = self.book.text_of(g.sheet, rect)
            try:
                self.copy_to_clipboard(self.clip_text)
            except Exception:
                pass
        n = rect.rows * rect.cols
        self.set_message(f"{'Cut' if cut else 'Copied'} {rect.a1()} ({n} cell{'s' if n != 1 else ''}) — "
                         f"select a destination and press Ctrl+V")
        g.refresh()

    def action_cut(self):
        self.action_copy(cut=True)

    def action_paste(self):
        if self.editing is not None:
            self.formula.action_paste()
            return
        g = self.grid
        if self.clip is None:
            self.notify("Nothing copied yet. To paste text from other programs use your terminal's paste "
                        "(often Ctrl+Shift+V).", timeout=6)
            return
        if self.clip_cut and self.clip["sheet"].book is not self.book:
            self.clip = None
            return
        r, c = g.selection.r1, g.selection.c1
        try:
            dest = self.book.paste(self.clip, g.sheet, r, c, cut=self.clip_cut, tile_to=g.selection)
        except ValueError as e:
            self.notify(str(e), severity="error")
            return
        if self.clip_cut:
            self.clip = None
            g.clip_rect = None
        self._select(dest)
        self.set_message(f"Pasted into {dest.a1()}")
        self.after_change()

    def on_paste(self, event: events.Paste):
        """Text pasted by the terminal (e.g. Ctrl+Shift+V) while the grid has focus."""
        if self.editing is not None or self.book is None:
            return
        text = event.text
        if not text:
            return
        if self.clip is not None and self.clip_text is not None and text.strip("\n") == self.clip_text.strip("\n"):
            self.action_paste()
            return
        g = self.grid
        dest = self.book.paste_text(g.sheet, *g.active, text)
        self._select(dest)
        self.set_message(f"Pasted text into {dest.a1()}")
        self.after_change()

    def _select(self, rect: Rect):
        g = self.grid
        g.active = (rect.r1, rect.c1)
        g.anchor = (rect.r1, rect.c1)
        g.corner = (rect.r2, rect.c2)
        g.ensure_visible(rect.r1, rect.c1)

    def action_escape(self):
        if self.editing is not None:
            self.cancel_edit()
            return
        if self.clip is not None or self.grid.clip_rect is not None:
            self.clip = None
            self.grid.clip_rect = None
            self.grid.refresh()
        self.set_message("")

    # ------------------------------------------------------------------ actions: editing

    def action_clear(self):
        if not self._need_grid():
            return
        g = self.grid
        self.book.clear(g.sheet, g.selection)
        self.after_change()

    def action_undo(self):
        if self.editing is not None:
            return
        e = self.book.do_undo()
        if e is None:
            self.set_message("Nothing to undo")
            return
        self._goto_where(e.where)
        self.set_message(f"Undid: {e.label}")
        self.after_change(layout=True)

    def action_redo(self):
        if self.editing is not None:
            return
        e = self.book.do_redo()
        if e is None:
            self.set_message("Nothing to redo")
            return
        self._goto_where(e.where)
        self.set_message(f"Redid: {e.label}")
        self.after_change(layout=True)

    def _goto_where(self, where):
        if not where:
            return
        name, rect = where
        s = self.book.sheet_by_name(name)
        g = self.grid
        if s is not None and s is not g.sheet:
            g.show_sheet(s)
        if s is not None:
            r2 = rect.r2 if rect.r2 < MAX_ROW else MAX_ROW
            self._select(Rect(rect.r1, rect.c1, r2, rect.c2))

    def action_font(self, attr: str):
        if not self._need_grid():
            return
        g = self.grid
        self.book.toggle_font(g.sheet, g.selection, attr)
        self.grid.resolver.invalidate()
        self.after_change()

    def action_fill(self, direction: str):
        if not self._need_grid():
            return
        g = self.grid
        self.book.fill(g.sheet, g.selection, direction)
        self.after_change()

    def action_recalc(self):
        if not self._need_grid():
            return
        self.book.recalc_all()
        self.set_message("Recalculated all formulas")
        self.after_change()

    def action_col_width(self, delta: int):
        if not self._need_grid():
            return
        g = self.grid
        sel = g.selection
        for c in range(sel.c1, min(sel.c2, sel.c1 + 200) + 1):
            w, _ = g.sheet.col_width(c)
            self.book.set_col_width(g.sheet, c, max(1.0, round(w) + delta))
        self.after_change(layout=True)

    def set_width_dialog(self):
        g = self.grid
        w, _ = g.sheet.col_width(g.active[1])

        def done(v):
            if v is None:
                return
            sel = g.selection
            for c in range(sel.c1, min(sel.c2, sel.c1 + 500) + 1):
                self.book.set_col_width(g.sheet, c, float(v))
            self.after_change(layout=True)

        self.push_screen(InputDialog("Column width", "Width in characters:", f"{w:g}",
                                     validate=_validate_float(0, 255)), done)

    def autofit(self):
        g = self.grid
        sheet = g.sheet
        sel = g.selection
        for c in range(sel.c1, min(sel.c2, sel.c1 + 200) + 1):
            best = 0
            for (r, cc), cell in sheet.ws._cells.items():
                if cc != c or cell._value is None:
                    continue
                v = self.book.get_value(sheet, r, c)
                txt = format_value(v, cell.number_format, self.book.epoch1904)[0]
                best = max(best, cell_len(txt))
            self.book.set_col_width(sheet, c, float(min(max(best + 1, 3), 80)))
        self.after_change(layout=True)

    def hide(self, axis: str, hidden: bool):
        g = self.grid
        sheet = g.sheet
        sel = g.selection
        if axis == "col":
            for c in range(sel.c1, min(sel.c2, sel.c1 + 500) + 1):
                self.book.set_col_width(sheet, c, None, hidden=hidden)
        else:
            from ..model.undo import UndoEntry

            rows = list(range(sel.r1, min(sel.r2, sel.r1 + 5000) + 1))
            prev = {r: sheet.row_hidden(r) for r in rows}

            def apply(state):
                for r, h in state.items():
                    sheet.ws.row_dimensions[r].hidden = h
                self.book.modified = True

            new = {r: hidden for r in rows}
            apply(new)
            self.book.undo.push(UndoEntry("Hide rows" if hidden else "Unhide rows",
                                          lambda: apply(prev), lambda: apply(new), (sheet.name, sel)))
        if hidden:
            # move off hidden cells
            if axis == "col":
                g.action_move(0, 1)
            else:
                g.action_move(1, 0)
        self.after_change(layout=True)

    # ------------------------------------------------------------------ outlines (grouped rows/columns)

    def _selection_axis(self) -> str | None:
        """'row' or 'col' when whole rows or whole columns are selected, else None."""
        sel = self.grid.selection
        full_rows = sel.c1 == 1 and sel.c2 >= MAX_COL
        full_cols = sel.r1 == 1 and sel.r2 >= MAX_ROW
        if full_rows != full_cols:
            return "row" if full_rows else "col"
        return None

    def _with_axis(self, title: str, then):
        axis = self._selection_axis()
        if axis is not None:
            then(axis)
            return
        self.push_screen(ChoiceDialog(title, f"{title} rows or columns?",
                                      [("row", "&Rows"), ("col", "&Columns"), ("cancel", "Cancel")], default="row"),
                         lambda v: then(v) if v in ("row", "col") else None)

    def group_selection(self, axis: str, ungroup: bool = False):
        if self.editing is not None and not self.commit_edit(move=None):
            return
        g = self.grid
        sel = g.selection
        a, b = (sel.r1, sel.r2) if axis == "row" else (sel.c1, sel.c2)
        what = "rows" if axis == "row" else "columns"
        try:
            fn = outline.ungroup if ungroup else outline.group
            changed = fn(self.book, g.sheet, axis, a, b)
        except ValueError as e:
            self.notify(str(e), severity="error")
            return
        if ungroup and not changed:
            self.set_message(f"Those {what} aren't grouped")
        else:
            self.set_message(f"{'Ungrouped' if ungroup else 'Grouped'} {what}")
        self.after_change(layout=True)

    def action_group(self):
        if self._need_grid():
            self._with_axis("Group", lambda axis: self.group_selection(axis))

    def action_ungroup(self):
        if self._need_grid():
            self._with_axis("Ungroup", lambda axis: self.group_selection(axis, ungroup=True))

    def _cursor_to_visible(self):
        """After hiding rows/columns, move the cursor to the nearest cell still showing."""
        g = self.grid
        r, c = g.active

        def nearest(i, hidden, limit):
            for d in range(0, limit):
                for j in (i + d, i - d):
                    if 1 <= j <= limit and not hidden(j):
                        return j
                if i + d > limit and i - d < 1:
                    break
            return i

        nr = nearest(r, g.sheet.row_hidden, MAX_ROW) if g.sheet.row_hidden(r) else r
        nc = nearest(c, lambda j: g.col_px(j) == 0, MAX_COL) if g.col_px(c) == 0 else c
        if (nr, nc) != (r, c):
            g.set_cursor(nr, nc)

    def set_group_collapsed(self, grp, collapse: bool):
        g = self.grid
        outline.set_collapsed(self.book, g.sheet, grp, collapse)
        g.invalidate()
        if collapse:
            self._cursor_to_visible()
        n = grp.end - grp.start + 1
        what = ("row" if grp.axis == "row" else "column") + ("s" if n != 1 else "")
        self.set_message(f"{'Collapsed' if collapse else 'Expanded'} {n} {what}")
        self.after_change(layout=True)

    def toggle_group(self, grp):
        self.set_group_collapsed(grp, not outline.is_collapsed(self.grid.sheet, grp))

    def _outline_target(self, collapse: bool):
        """The group the cursor is in (or on the button of) for collapse/expand from the keyboard:
        the innermost expanded group to collapse, or the outermost collapsed one to expand."""
        g = self.grid
        sheet = g.sheet
        r, c = g.active
        axis = self._selection_axis()
        for ax in ([axis] if axis else ["row", "col"]):
            i = r if ax == "row" else c
            groups = outline.groups_at(sheet, ax, i)
            if collapse:
                cands = [x for x in groups if not outline.is_collapsed(sheet, x)]
                if cands:
                    return cands[-1]
            else:
                cands = [x for x in groups if x.summary == i and outline.is_collapsed(sheet, x)]
                if cands:
                    return cands[0]
        return None

    def action_outline(self, collapse: bool):
        if not self._need_grid():
            return
        if self.editing is not None and not self.commit_edit(move=None):
            return
        grp = self._outline_target(collapse)
        if grp is None:
            self.set_message("No group to collapse here" if collapse else
                             "No collapsed group here — put the cursor on the row or column with the [+]")
            return
        self.set_group_collapsed(grp, collapse)

    def show_outline_level(self, axis: str, n: int):
        g = self.grid
        if outline.show_level(self.book, g.sheet, axis, n):
            g.invalidate()
            self._cursor_to_visible()
        what = "rows" if axis == "row" else "columns"
        self.set_message(f"Showing outline level {n} of the {what}")
        self.after_change(layout=True)

    def outline_all(self, collapse: bool):
        g = self.grid
        changed = False
        for axis in ("row", "col"):
            top = outline.max_level(g.sheet, axis)
            if top:
                changed |= outline.show_level(self.book, g.sheet, axis, 1 if collapse else top + 1)
        if not outline.max_level(g.sheet, "row") and not outline.max_level(g.sheet, "col"):
            self.set_message("This sheet has no grouped rows or columns")
            return
        g.invalidate()
        if collapse:
            self._cursor_to_visible()
        self.set_message(("Collapsed" if collapse else "Expanded") + " all groups" + ("" if changed else " (no change)"))
        self.after_change(layout=True)

    def toggle_outline_symbols(self):
        from openpyxl.worksheet.properties import Outline

        ws = self.grid.sheet.ws
        pr = ws.sheet_properties.outlinePr
        if pr is None:
            pr = ws.sheet_properties.outlinePr = Outline(summaryBelow=True, summaryRight=True)
        pr.showOutlineSymbols = not outline.symbols_shown(self.grid.sheet)
        self.book.modified = True
        self.after_change(layout=True)

    def number_format_dialog(self):
        g = self.grid
        cur = g.sheet.number_format(*g.active)
        presets = "\n".join(f"  {i + 1:>2}. {name}" for i, (name, _) in enumerate(NUMBER_FORMATS))

        def done(v):
            if v is None or v == "":
                return
            v = v.strip()
            if v.isdigit() and 1 <= int(v) <= len(NUMBER_FORMATS):
                v = NUMBER_FORMATS[int(v) - 1][1]
            self.book.set_number_format(g.sheet, g.selection, v)
            self.after_change()

        self.push_screen(InputDialog("Number format", "Pick a number from the list or type an Excel format code:",
                                     cur, help_text=presets), done)

    def apply_number_format(self, fmt: str):
        g = self.grid
        self.book.set_number_format(g.sheet, g.selection, fmt)
        self.after_change()

    def change_decimals(self, delta: int):
        g = self.grid
        cur = g.sheet.number_format(*g.active)
        self.book.set_number_format(g.sheet, g.selection, _adjust_decimals(cur, delta, self.book.get_value(
            g.sheet, *g.active)))
        self.after_change()

    def color_dialog(self, which: str):
        g = self.grid
        names = ", ".join(COLORS)

        def done(v):
            if v is None:
                return
            key = v.strip().lower().lstrip("#")
            rgb = COLORS.get(key, key.upper() if re.fullmatch(r"[0-9a-fA-F]{6}", key) else "INVALID")
            if rgb == "INVALID":
                self.notify("Unknown color", severity="error")
                return
            self.book.set_color(g.sheet, g.selection, which, rgb)
            g.resolver.invalidate()
            self.after_change()

        self.push_screen(InputDialog("Fill color" if which == "fill" else "Font color",
                                     "Color name or hex code (e.g. FFC000):", "",
                                     help_text=names), done)

    def align(self, how: str | None):
        g = self.grid
        self.book.set_alignment(g.sheet, g.selection, how)
        g.resolver.invalidate()
        self.after_change()

    def clear_formats(self):
        g = self.grid
        self.book.clear(g.sheet, g.selection, formats=True)
        g.resolver.invalidate()
        self.after_change()

    def freeze(self, how: str):
        g = self.grid
        ws = g.sheet.ws
        old = ws.freeze_panes
        if how == "cursor":
            r, c = g.active
            new = addr(r, c) if (r, c) != (1, 1) else None
        elif how == "row":
            new = "A2"
        elif how == "col":
            new = "B1"
        else:
            new = None
        from ..model.undo import UndoEntry

        def apply(v):
            ws.freeze_panes = v
            self.book.modified = True

        apply(new)
        self.book.undo.push(UndoEntry("Freeze panes", lambda: apply(old), lambda: apply(new)))
        fr, fc = g.frozen()
        g.top = max(g.top, fr + 1)
        g.left = max(g.left, fc + 1)
        self.after_change(layout=True)

    def sort_selection(self, descending: bool):
        g = self.grid
        sheet = g.sheet
        sel = self.book.clip(sheet, g.selection)
        if sel.rows * sel.cols == 1:
            sel = self._current_region(*g.active)
        key_col = g.active[1] if sel.c1 <= g.active[1] <= sel.c2 else sel.c1
        header = False
        first = [self.book.get_value(sheet, sel.r1, c) for c in range(sel.c1, sel.c2 + 1)]
        second = self.book.get_value(sheet, sel.r1 + 1, key_col) if sel.rows > 1 else None
        if sel.rows > 1 and all(isinstance(v, str) for v in first if v is not None) and any(first) and \
                not isinstance(second, str):
            header = True
        body = Rect(sel.r1 + (1 if header else 0), sel.c1, sel.r2, sel.c2)
        if body.rows < 2:
            return
        from ..model.workbook import _formula_text
        from ..formula.parser import translate
        from ..formula.values import compare

        rows = []
        for r in range(body.r1, body.r2 + 1):
            k = self.book.get_value(sheet, r, key_col)
            snaps = []
            for c in range(body.c1, body.c2 + 1):
                cell = sheet.ws._cells.get((r, c))
                snaps.append(None if cell is None else (cell._value, cell.data_type, cell._style))
            rows.append((k, r, snaps))
        import functools

        def cmp(a, b):
            ka, kb = a[0], b[0]
            if ka is None or ka == "":
                return 0 if (kb is None or kb == "") else 1
            if kb is None or kb == "":
                return -1
            if isinstance(ka, ExcelError) or isinstance(kb, ExcelError):
                return (isinstance(ka, ExcelError)) - (isinstance(kb, ExcelError))
            c = compare(ka, kb)
            return -c if descending else c

        ordered = sorted(rows, key=functools.cmp_to_key(cmp))
        from copy import copy

        with self.book.transaction("Sort", (sheet.name, sel)):
            for i, (_, src_r, snaps) in enumerate(ordered):
                r = body.r1 + i
                for j, snap in enumerate(snaps):
                    c = body.c1 + j
                    if snap is None:
                        self.book.set_raw(sheet, r, c, None)
                        continue
                    value, dtype, style = snap
                    if dtype == "f" and value is not None:
                        value = translate(_formula_text(value), r - src_r, 0)
                        self.book.set_raw(sheet, r, c, value)
                    else:
                        self.book.set_raw(sheet, r, c, value, dtype if dtype in ("s", "e") else None)
                    self.book.set_style(sheet, r, c, _style=copy(style))
        self._select(sel)
        self.set_message(f"Sorted {body.a1()} by column {col_letter(key_col)}"
                         + (" (kept header row)" if header else ""))
        self.after_change()

    def _current_region(self, r: int, c: int) -> Rect:
        """The contiguous block of non-empty cells around (r, c) (like Excel's Ctrl+A first press)."""
        sheet = self.grid.sheet
        book = self.book

        def filled(rr, cc):
            if rr < 1 or cc < 1:
                return False
            v = book.get_value(sheet, rr, cc)
            return v is not None and v != ""

        r1 = r2 = r
        c1 = c2 = c
        changed = True
        while changed:
            changed = False
            if r1 > 1 and any(filled(r1 - 1, cc) for cc in range(c1 - 1, c2 + 2)):
                r1 -= 1
                changed = True
            if r2 < sheet.max_row and any(filled(r2 + 1, cc) for cc in range(c1 - 1, c2 + 2)):
                r2 += 1
                changed = True
            if c1 > 1 and any(filled(rr, c1 - 1) for rr in range(r1 - 1, r2 + 2)):
                c1 -= 1
                changed = True
            if c2 < sheet.max_col and any(filled(rr, c2 + 1) for rr in range(r1 - 1, r2 + 2)):
                c2 += 1
                changed = True
        return Rect(r1, c1, r2, c2)

    # ------------------------------------------------------------------ rows / columns

    def action_insert(self):
        if not self._need_grid():
            return
        self._row_col_op("insert")

    def action_delete(self):
        if not self._need_grid():
            return
        self._row_col_op("delete")

    def _row_col_op(self, op: str):
        g = self.grid
        sel = g.selection
        if sel.c1 == 1 and sel.c2 == MAX_COL:
            self._do_row_col(op, "row")
        elif sel.r1 == 1 and sel.r2 == MAX_ROW:
            self._do_row_col(op, "col")
        else:
            def done(choice):
                if choice in ("row", "col"):
                    self._do_row_col(op, choice)

            n_r, n_c = sel.rows, sel.cols
            self.push_screen(ChoiceDialog(
                "Insert" if op == "insert" else "Delete",
                f"{op.capitalize()} entire rows or columns?",
                [("row", f"&Rows ({n_r})"), ("col", f"&Columns ({n_c})"), ("cancel", "Cancel")],
                default="row"), done)

    def _do_row_col(self, op: str, axis: str):
        g = self.grid
        sel = g.selection
        idx = sel.r1 if axis == "row" else sel.c1
        n = (sel.rows if axis == "row" else sel.cols)
        if (axis == "row" and sel.r2 == MAX_ROW) or (axis == "col" and sel.c2 == MAX_COL):
            n = 1
        try:
            if op == "insert":
                structure.insert(self.book, g.sheet, axis, idx, n)
            else:
                structure.delete(self.book, g.sheet, axis, idx, n)
        except ValueError as e:
            self.notify(str(e), severity="error")
            return
        what = "row" if axis == "row" else "column"
        self.set_message(f"{'Inserted' if op == 'insert' else 'Deleted'} {n} {what}{'s' if n != 1 else ''}")
        self.after_change(layout=True)

    # ------------------------------------------------------------------ sheets

    def switch_sheet(self, idx: int):
        if self.editing is not None and not inp_is_formula(self.formula.value):
            if not self.commit_edit(move=None):
                return
        s = self.book.sheets[idx]
        self.grid.show_sheet(s)
        self.after_change(layout=True)

    def _switch_sheet_relative(self, d: int):
        sheets = [s for s in self.book.sheets if s.ws.sheet_state == "visible"] or self.book.sheets
        cur = self.grid.sheet
        i = sheets.index(cur) if cur in sheets else 0
        nxt = sheets[(i + d) % len(sheets)]
        self.grid.show_sheet(nxt)
        self.after_change(layout=True)

    def action_sheet(self, d: int):
        if self.editing is not None:
            if not self.commit_edit(move=None):
                return
        self._switch_sheet_relative(d)

    def action_add_sheet(self):
        if self.editing is not None and not self.commit_edit(move=None):
            return
        idx = self.book.sheets.index(self.grid.sheet) + 1
        s = structure.add_sheet(self.book, None, idx)
        self.grid.show_sheet(s)
        self.after_change(layout=True)

    def action_rename_sheet(self):
        g = self.grid
        sheet = g.sheet

        def done(v):
            if v is None:
                return
            try:
                structure.rename_sheet(self.book, sheet, v.strip())
            except ValueError as e:
                self.notify(str(e), severity="error")
                return
            self.after_change()

        self.push_screen(InputDialog("Rename sheet", "New name:", sheet.name,
                                     validate=lambda v: structure.validate_sheet_name(self.book, v, exclude=sheet)),
                         done)

    def action_delete_sheet(self):
        g = self.grid
        sheet = g.sheet
        if len(self.book.sheets) == 1:
            self.notify("A workbook needs at least one sheet", severity="error")
            return

        def done(choice):
            if choice != "delete":
                return
            idx = self.book.sheets.index(sheet)
            structure.remove_sheet(self.book, sheet)
            g.show_sheet(self.book.sheets[max(0, idx - 1)])
            self.set_message(f"Deleted sheet {sheet.name} (this can't be undone)")
            self.after_change(layout=True)

        self.push_screen(ChoiceDialog("Delete sheet", f"Delete '{sheet.name}'? This can't be undone, and "
                                                      "formulas that refer to it will show #REF!.",
                                      [("delete", "&Delete"), ("cancel", "Cancel")], default="cancel"), done)

    def move_sheet(self, d: int):
        g = self.grid
        wb = self.book.wb
        ws = g.sheet.ws
        i = wb._sheets.index(ws)
        j = i + d
        if not (0 <= j < len(wb._sheets)):
            return
        wb._sheets[i], wb._sheets[j] = wb._sheets[j], wb._sheets[i]
        self.book.sheets.sort(key=lambda s: wb.worksheets.index(s.ws))
        self.book.modified = True
        self.after_change()

    def unhide_sheets(self):
        n = 0
        for s in self.book.sheets:
            if s.ws.sheet_state != "visible":
                s.ws.sheet_state = "visible"
                n += 1
        self.book.modified = self.book.modified or n > 0
        self.set_message(f"Unhid {n} sheet{'s' if n != 1 else ''}")
        self.after_change()

    def hide_sheet(self):
        g = self.grid
        visible = [s for s in self.book.sheets if s.ws.sheet_state == "visible"]
        if len(visible) <= 1:
            self.notify("At least one sheet must stay visible", severity="error")
            return
        g.sheet.ws.sheet_state = "hidden"
        self.book.modified = True
        self._switch_sheet_relative(1)

    # ------------------------------------------------------------------ named ranges

    def _name_for_selection(self) -> str | None:
        """Excel shows a range's name in the address box when the selection matches it."""
        g = self.grid
        sel = g.selection
        if sel.rows * sel.cols == 1 and not g.sheet.merged_at(*g.active):
            return None
        for info in names.list_names(self.book):
            if info.scope is not None and info.scope is not g.sheet:
                continue
            target = names.resolve_rect(self.book, info.name, g.sheet)
            if target is not None and target[0] is g.sheet and target[1] == sel:
                return info.name
        for hit in self.book.table_index().by_sheet.get(g.sheet, ()):
            if sel == hit.g.data_rect or sel == hit.g.rect:
                return hit.name
        return None

    @on(Input.Submitted, "#namebox")
    def _namebox_submitted(self, event: Input.Submitted):
        text = event.value.strip()
        if not text:
            self.namebox_done()
            return
        target = self._resolve_goto(text)
        if target is not None:
            sheet, rect = target
            g = self.grid
            if sheet is not g.sheet:
                g.show_sheet(sheet)
            self._select(rect)
            self.namebox_done()
            self.after_change(layout=True)
            return
        g = self.grid
        if self.book.resolve_name(text, g.sheet) is not None:
            self.notify(f"'{text}' isn't a range (see Name Manager, Ctrl+F3)", severity="warning")
            self.namebox_done()
            return
        try:
            names.set_name(self.book, text, names.selection_ref(g.sheet, self.book.clip(g.sheet, g.selection)),
                           None, g.sheet)
        except ValueError as e:
            self.notify(str(e), severity="error")
            return
        self.set_message(f"Defined name {text} = {names.selection_ref(g.sheet, self.book.clip(g.sheet, g.selection))}")
        self.namebox_done()
        self.after_change()

    def namebox_done(self):
        self.grid.focus()
        self.grid_changed()

    def name_list(self):
        """[(NameInfo, current value as text)] for the Name Manager."""
        from ..formula.evaluator import Ctx, evaluate
        from ..formula.values import Array, RefValue

        out = []
        g = self.grid
        for info in names.list_names(self.book):
            sheet = info.scope or g.sheet
            value = ""
            ast = self.book.resolve_name(info.name, sheet, info.scope.name if info.scope else None) \
                if info.scope else self.book.names.get(info.name.lower())
            if ast is not None:
                try:
                    v = evaluate(ast, Ctx(self.book, sheet, 1, 1))
                    if isinstance(v, RefValue):
                        v = self.book.get_value(v.sheet, v.r1, v.c1) if v.is_cell else f"{{{v.nrows}×{v.ncols} range}}"
                    elif isinstance(v, Array):
                        v = f"{{{v.nrows}×{v.ncols} array}}"
                    value = v if isinstance(v, str) else format_value(v, "General")[0]
                except Exception:  # noqa: BLE001
                    value = "?"
            out.append((info, value))
        return out

    def action_name_manager(self):
        if self.editing is not None or self.book is None:
            return
        self.push_screen(NameManagerDialog(self))

    def action_define_name(self):
        if self.editing is not None or self.book is None:
            return
        self.open_name_editor(None)

    def open_name_editor(self, info, done=None):
        g = self.grid
        book = self.book
        if info is None:
            refers = "=" + names.selection_ref(g.sheet, book.clip(g.sheet, g.selection))
            old = None
            dlg_args = dict(name="", refers_to=refers, scope=None, editing=False)
        else:
            old = (info.name, info.scope)
            dlg_args = dict(name=info.name, refers_to=info.refers_to, scope=info.scope, editing=True)

        def submit(name, refers, scope):
            try:
                names.set_name(book, name, refers, scope, g.sheet, old=old)
            except ValueError as e:
                return str(e)
            self.after_change()
            return None

        def closed(name):
            if name:
                self.set_message(("Updated" if old else "Defined") + f" name {name}")
            if done is not None:
                done(name)

        self.push_screen(NameEditDialog(submit, list(book.sheets), **dlg_args), closed)

    def remove_name(self, info):
        names.delete_name(self.book, info.name, info.scope)
        self.set_message(f"Deleted name {info.name}")
        self.after_change()

    def goto_name(self, info) -> bool:
        target = names.resolve_rect(self.book, info.name, info.scope or self.grid.sheet)
        if target is None:
            self.notify(f"{info.name} doesn't refer to a range", severity="warning")
            return False
        sheet, rect = target
        g = self.grid
        if sheet is not g.sheet:
            g.show_sheet(sheet)
        self._select(rect)
        self.after_change(layout=True)
        return True

    # ------------------------------------------------------------------ tables

    def _table_here(self):
        g = self.grid
        return tables.table_at(self.book, g.sheet, *g.active)

    def action_table(self):
        """Ctrl+T: edit the table at the cursor, or make the selection (or current region) a table."""
        if not self._need_grid():
            return
        if self._table_here() is not None:
            self.edit_table_dialog()
        else:
            self.create_table_dialog()

    def _table_range(self, text: str):
        target = self._resolve_goto(text.strip())
        if target is None:
            return None, "Enter a range such as A1:D20"
        sheet, rect = target
        if sheet is not self.grid.sheet:
            return None, "The range has to be on this sheet"
        return self.book.clip(sheet, rect), None

    def create_table_dialog(self):
        if not self._need_grid():
            return
        g = self.grid
        sel = self.book.clip(g.sheet, g.selection)
        if sel.rows * sel.cols == 1:
            sel = self._current_region(*g.active)
        name = tables.default_name(self.book)

        def submit(name, ref, headers, totals):
            rect, err = self._table_range(ref)
            if err:
                return err
            try:
                t = tables.create(self.book, g.sheet, rect, headers, name)
            except ValueError as e:
                return str(e)
            self._select(tables.geom(t).rect)
            self.set_message(f"Created table {t.displayName} ({t.ref}). Use it in formulas as "
                             f"{t.displayName}[{t.tableColumns[0].name}]")
            self.after_change()
            return None

        self.push_screen(TableDialog(submit, name, sel.a1()))

    def edit_table_dialog(self):
        if not self._need_grid():
            return
        hit = self._table_here()
        if hit is None:
            self.notify("The cursor isn't in a table (Ctrl+T on a range creates one)", severity="warning")
            return
        name = hit.name

        def submit(new_name, ref, headers, totals):
            cur = tables.find_table(self.book, name)
            if cur is None:
                return "The table no longer exists"
            rect, err = self._table_range(ref)
            if err:
                return err
            u = self.book.undo
            top = u.entries[u.pos - 1] if u.pos else None
            try:
                with self.book.transaction("Table settings", (cur.sheet.name, cur.g.rect)):
                    if new_name != cur.name:
                        tables.rename(self.book, cur, new_name)
                        cur = tables.find_table(self.book, new_name)
                    if rect != cur.g.rect:
                        tables.resize(self.book, cur, rect)
                        cur = tables.find_table(self.book, cur.name)
                    if totals != bool(cur.g.tot):
                        tables.set_totals(self.book, cur, totals)
            except ValueError as e:
                if u.pos and u.entries[u.pos - 1] is not top:  # roll back the part that was done
                    self.book.do_undo()
                    del u.entries[u.pos:]
                self.after_change()
                return str(e)
            self.set_message(f"Updated table {new_name}")
            self.after_change()
            return None

        def done(result):
            cur = tables.find_table(self.book, name)
            if result == "convert" and cur is not None:
                tables.convert_to_range(self.book, cur)
                self.set_message(f"Converted table {name} to a normal range (Ctrl+Z undoes this)")
                self.after_change()
            elif result == "delete" and cur is not None:
                tables.delete_table(self.book, cur)
                self.set_message(f"Deleted table {name} and its contents (Ctrl+Z undoes this)")
                self.after_change()

        self.push_screen(TableDialog(submit, name, hit.g.rect.a1(), editing=True, totals=bool(hit.g.tot)), done)

    def table_command(self, what: str):
        if not self._need_grid():
            return
        hit = self._table_here()
        if hit is None:
            self.notify("The cursor isn't in a table", severity="warning")
            return
        try:
            if what == "totals":
                tables.set_totals(self.book, hit, not hit.g.tot)
                self.set_message(f"{'Removed' if hit.g.tot else 'Added'} the totals row of {hit.name}")
            elif what == "convert":
                tables.convert_to_range(self.book, hit)
                self.set_message(f"Converted table {hit.name} to a normal range")
            elif what == "delete":
                tables.delete_table(self.book, hit)
                self.set_message(f"Deleted table {hit.name}")
            elif what == "select":
                self._select(hit.g.data_rect)
            elif what == "column":
                c = self.grid.active[1]
                self._select(Rect(hit.g.data_r1, c, hit.g.data_r2, c))
        except ValueError as e:
            self.notify(str(e), severity="error")
            return
        self.after_change()

    # ------------------------------------------------------------------ navigation dialogs

    def action_goto(self):
        if self.editing is not None:
            return

        def done(v):
            if not v:
                return
            target = self._resolve_goto(v.strip())
            if target is None:
                self.notify(f"Can't find {v}", severity="error", markup=False)
                return
            sheet, rect = target
            g = self.grid
            if sheet is not g.sheet:
                g.show_sheet(sheet)
            self._select(rect)
            self.after_change(layout=True)

        self.push_screen(InputDialog("Go to", "Cell, range or name (e.g. B12, Sheet2!A1:C9):", ""), done)

    def _resolve_goto(self, text: str):
        sheet = self.grid.sheet
        if "!" in text:
            name, _, ref = text.rpartition("!")
            name = name.strip("'").replace("''", "'")
            sheet = self.book.sheet_by_name(name)
            if sheet is None:
                return None
            text = ref
        from ..model.refs import parse_range

        rect = parse_range(text.replace("$", ""))
        if rect is not None:
            return sheet, rect
        ast = self.book.resolve_name(text, sheet)
        from ..formula.parser import RefNode, StructRefNode, parse

        if ast is None and "[" in text:
            try:
                ast = parse("=" + text)
            except (FormulaSyntaxError, Unsupported):
                ast = None
        if isinstance(ast, RefNode):
            s = sheet if ast.sheet is None else self.book.sheet_by_name(ast.sheet)
            if s is not None:
                return s, Rect(ast.r1, ast.c1, min(ast.r2, MAX_ROW), ast.c2)
        if isinstance(ast, StructRefNode):
            r, c = self.grid.active
            res = tables.resolve(self.book, ast.ref, sheet, r, c)
            if isinstance(res, tuple):
                s, r1, c1, r2, c2 = res
                return s, Rect(r1, c1, r2, c2)
        return None

    def action_find(self, replace: bool = False):
        if self.editing is not None:
            return
        self._find_state = None
        self.push_screen(FindDialog(replace=replace))

    def action_replace(self):
        self.action_find(replace=True)

    def _cell_search_text(self, sheet, r, c, formulas: bool) -> str:
        if formulas:
            t = sheet.input_text(r, c)
            if t:
                return t
        v = self.book.get_value(sheet, r, c)
        if v is None:
            return ""
        return format_value(v, sheet.number_format(r, c), self.book.epoch1904)[0]

    def _matcher(self, text: str, case: bool, whole: bool):
        if any(ch in text for ch in "*?~"):
            rx = wildcard_regex(text)
            if case:
                rx = re.compile(rx.pattern, re.S)
            return (lambda s: rx.fullmatch(s) is not None) if whole else (lambda s: rx.search(s) is not None)
        if not case:
            low = text.lower()
            return (lambda s: s.lower() == low) if whole else (lambda s: low in s.lower())
        return (lambda s: s == text) if whole else (lambda s: text in s)

    def find_next(self, text: str, case: bool, whole: bool, formulas: bool, direction: int = 1):
        g = self.grid
        sheet = g.sheet
        match = self._matcher(text, case, whole)
        keys = sorted(k for k, cell in sheet.ws._cells.items() if cell._value is not None)
        if not keys:
            return None
        cur = g.active
        order = keys if direction > 0 else list(reversed(keys))
        after = [k for k in order if (k > cur if direction > 0 else k < cur)]
        before = [k for k in order if (k <= cur if direction > 0 else k >= cur)]
        for r, c in after + before:
            if match(self._cell_search_text(sheet, r, c, formulas)):
                g.set_cursor(r, c)
                return f"Found in {addr(r, c)}"
        return None

    def replace_current(self, replacement: str, text: str, case: bool, whole: bool, formulas: bool) -> str:
        g = self.grid
        sheet = g.sheet
        r, c = g.active
        src = sheet.input_text(r, c)
        match = self._matcher(text, case, whole)
        if src and match(src if formulas or not sheet.is_formula(r, c) else ""):
            new = _replace_text(src, text, replacement, case, whole)
            try:
                self.book.set_input(sheet, r, c, new)
            except FormulaSyntaxError as e:
                return f"Can't replace in {addr(r, c)}: {e}"
            self.after_change()
        found = self.find_next(text, case, whole, formulas)
        return found or "No more matches"

    def replace_all(self, replacement: str, text: str, case: bool, whole: bool, formulas: bool) -> int:
        g = self.grid
        sheet = g.sheet
        sel = g.selection
        limit = sel if sel.rows * sel.cols > 1 else None
        match = self._matcher(text, case, whole)
        count = 0
        with self.book.transaction("Replace all", (sheet.name, sel)):
            for (r, c), cell in sorted(sheet.ws._cells.items()):
                if cell._value is None or (limit is not None and not limit.contains(r, c)):
                    continue
                if sheet.is_formula(r, c) and not formulas:
                    continue
                src = sheet.input_text(r, c)
                if src and match(src):
                    try:
                        self.book.set_input(sheet, r, c, _replace_text(src, text, replacement, case, whole))
                        count += 1
                    except FormulaSyntaxError:
                        pass
        self.after_change()
        return count

    # ------------------------------------------------------------------ files

    def action_save(self):
        if self.editing is not None and not self.commit_edit(move=None):
            return
        book = self.book
        if book.path is None:
            self.action_save_as()
            return
        self._save_to(book.path)

    def _save_to(self, path: str, confirmed: bool = False):
        book = self.book
        ext = os.path.splitext(path)[1].lower()
        same_file = book.path is not None and os.path.abspath(path) == os.path.abspath(book.path)
        if not confirmed and same_file and book.warnings and path not in self.ack_warnings and os.path.exists(path):
            def done(choice):
                if choice == "overwrite":
                    self.ack_warnings.add(path)
                    self._save_to(path, confirmed=True)
                elif choice == "copy":
                    base, e = os.path.splitext(path)
                    self.action_save_as(default=base + " (term_xl)" + (e if e in xlsx.EXCEL_EXTS else ".xlsx"))

            self.push_screen(ChoiceDialog(
                "Save", "This file contains things term_xl can't write back:\n  • " + "\n  • ".join(book.warnings)
                + "\n\nOverwriting will remove them. Save a copy instead?",
                [("copy", "Save &copy…"), ("overwrite", "&Overwrite"), ("cancel", "Cancel")], default="copy"), done)
            return
        if not confirmed and ext in xlsx.TEXT_EXTS and (len(book.sheets) > 1 or any(s.formulas for s in book.sheets)):
            def done(choice):
                if choice == "ok":
                    self._save_to(path, confirmed=True)

            self.push_screen(ChoiceDialog(
                "Save as CSV", "CSV keeps only the current sheet's displayed values — formulas, formatting and "
                               "other sheets are not saved. Continue?",
                [("ok", "&Save CSV"), ("cancel", "Cancel")], default="ok"), done)
            return
        try:
            xlsx.save(book, path, active_sheet=self.grid.sheet, cursor=self.grid.active)
        except Exception as e:  # noqa: BLE001
            self.notify(f"Save failed: {e}", severity="error", timeout=10)
            return
        if ext in xlsx.TEXT_EXTS:
            book.path = path
        self.sub_title = os.path.basename(path)
        self.set_message(f"Saved {path}")
        self.after_change()

    def action_save_as(self, default: str | None = None):
        if self.editing is not None and not self.commit_edit(move=None):
            return
        book = self.book
        if default is None:
            default = book.path or os.path.join(os.getcwd(), "Book1.xlsx")

        def done(v):
            if not v:
                return
            path = os.path.abspath(os.path.expanduser(v.strip()))
            if not os.path.splitext(path)[1]:
                path += ".xlsx"
            if os.path.isdir(path):
                self.notify("That's a folder", severity="error")
                return
            if os.path.exists(path) and path != book.path:
                def ok(choice):
                    if choice == "yes":
                        self._save_to(path, confirmed=True)
                        book.warnings = []

                self.push_screen(ChoiceDialog("Replace file?", f"{path} already exists. Replace it?",
                                              [("yes", "&Replace"), ("no", "Cancel")], default="no"), ok)
                return
            self._save_to(path, confirmed=True)
            if path != book.path:
                book.warnings = []
            book.path = path

        self.push_screen(InputDialog("Save as", "File name (.xlsx, .xlsm or .csv) — Tab completes paths:",
                                     default, path=True), done)

    def _confirm_discard(self, then):
        if self.book is None or not self.book.modified:
            then()
            return

        def done(choice):
            if choice == "save":
                self.action_save()
                if not self.book.modified:
                    then()
            elif choice == "discard":
                then()

        name = os.path.basename(self.book.path) if self.book.path else "this workbook"
        self.push_screen(ChoiceDialog("Unsaved changes", f"Save changes to {name}?",
                                      [("save", "&Save"), ("discard", "&Don't save"), ("cancel", "Cancel")],
                                      default="save"), done)

    def action_open(self):
        if self.editing is not None:
            return

        def open_path(v):
            if not v:
                return
            path = os.path.abspath(os.path.expanduser(v.strip()))
            if not os.path.exists(path):
                self.notify(f"No such file: {path}", severity="error")
                return
            self._confirm_discard(lambda: self._open(path))

        start = os.path.dirname(self.book.path) + os.sep if self.book and self.book.path else os.getcwd() + os.sep
        self.push_screen(InputDialog("Open", "File (.xlsx, .xlsm, .csv) — Tab completes paths:", start, path=True),
                         open_path)

    def _open(self, path: str):
        self.set_message(f"Opening {os.path.basename(path)}…")
        self.load_file(path)

    def action_new(self):
        if self.editing is not None:
            return
        self._confirm_discard(lambda: (self._set_book(xlsx.new_workbook()), self.set_message("New workbook")))

    def action_quit_app(self):
        if self.editing is not None:
            self.cancel_edit()
        self._confirm_discard(self.exit)

    def action_help(self):
        self.push_screen(HelpScreen())

    def action_toggle_paper(self):
        g = self.grid
        g.set_palette(LIGHT if g.palette is DARK else DARK)

    # ------------------------------------------------------------------ command palette

    def get_system_commands(self, screen):
        if self.book is None or self.editing is not None:
            yield from super().get_system_commands(screen)
            return
        C = SystemCommand
        yield C("Save", "Save the workbook (Ctrl+S)", self.action_save)
        yield C("Save as…", "Save under a new name or format (F12)", self.action_save_as)
        yield C("Open…", "Open a workbook (Ctrl+O)", self.action_open)
        yield C("New workbook", "Start a blank workbook (Ctrl+N)", self.action_new)
        yield C("Insert rows", "Insert rows above the selection", lambda: self._do_row_col("insert", "row"))
        yield C("Insert columns", "Insert columns left of the selection", lambda: self._do_row_col("insert", "col"))
        yield C("Delete rows", "Delete the selected rows", lambda: self._do_row_col("delete", "row"))
        yield C("Delete columns", "Delete the selected columns", lambda: self._do_row_col("delete", "col"))
        yield C("Column width…", "Set the width of the selected columns", self.set_width_dialog)
        yield C("Autofit column width", "Fit selected columns to their contents", self.autofit)
        yield C("Hide columns", "Hide the selected columns", lambda: self.hide("col", True))
        yield C("Unhide columns", "Unhide columns within the selection", lambda: self.hide("col", False))
        yield C("Hide rows", "Hide the selected rows", lambda: self.hide("row", True))
        yield C("Unhide rows", "Unhide rows within the selection", lambda: self.hide("row", False))
        yield C("Group rows", "Group the selected rows so they can be collapsed (Alt+Shift+Right)",
                lambda: self.group_selection("row"))
        yield C("Group columns", "Group the selected columns so they can be collapsed (Alt+Shift+Right)",
                lambda: self.group_selection("col"))
        yield C("Ungroup rows", "Remove the selected rows from their group (Alt+Shift+Left)",
                lambda: self.group_selection("row", ungroup=True))
        yield C("Ungroup columns", "Remove the selected columns from their group (Alt+Shift+Left)",
                lambda: self.group_selection("col", ungroup=True))
        yield C("Collapse group", "Hide the detail of the group at the cursor (Alt+Shift+Up)",
                lambda: self.action_outline(True))
        yield C("Expand group", "Show the detail of the collapsed group at the cursor (Alt+Shift+Down)",
                lambda: self.action_outline(False))
        yield C("Collapse all groups", "Collapse every grouped row and column on this sheet",
                lambda: self.outline_all(True))
        yield C("Expand all groups", "Expand every grouped row and column on this sheet",
                lambda: self.outline_all(False))
        yield C("Show or hide outline symbols", "Toggle the group bars and [+]/[-] buttons",
                self.toggle_outline_symbols)
        yield C("Number format…", "Choose or type a number format", self.number_format_dialog)
        for name, fmt in NUMBER_FORMATS:
            yield C(f"Format as {name}", fmt, lambda f=fmt: self.apply_number_format(f))
        yield C("Increase decimals", "Show one more decimal place", lambda: self.change_decimals(1))
        yield C("Decrease decimals", "Show one less decimal place", lambda: self.change_decimals(-1))
        yield C("Bold", "Toggle bold (Ctrl+B)", lambda: self.action_font("bold"))
        yield C("Italic", "Toggle italic (Ctrl+I / Alt+I)", lambda: self.action_font("italic"))
        yield C("Underline", "Toggle underline (Ctrl+U)", lambda: self.action_font("underline"))
        yield C("Strikethrough", "Toggle strikethrough", lambda: self.action_font("strike"))
        yield C("Align left", "Left-align cells", lambda: self.align("left"))
        yield C("Align center", "Center cells", lambda: self.align("center"))
        yield C("Align right", "Right-align cells", lambda: self.align("right"))
        yield C("Align general", "Default alignment (numbers right, text left)", lambda: self.align(None))
        yield C("Fill color…", "Background color of the selected cells", lambda: self.color_dialog("fill"))
        yield C("Font color…", "Text color of the selected cells", lambda: self.color_dialog("font"))
        yield C("Clear formats", "Remove formatting and contents of the selection", self.clear_formats)
        yield C("Sort ascending", "Sort the selection (or current region) by the active column, A→Z",
                lambda: self.sort_selection(False))
        yield C("Sort descending", "Sort the selection (or current region) by the active column, Z→A",
                lambda: self.sort_selection(True))
        yield C("Freeze panes at cursor", "Keep rows above and columns left of the cursor visible",
                lambda: self.freeze("cursor"))
        yield C("Freeze top row", "Keep row 1 visible", lambda: self.freeze("row"))
        yield C("Freeze first column", "Keep column A visible", lambda: self.freeze("col"))
        yield C("Unfreeze panes", "Remove frozen panes", lambda: self.freeze(None))
        yield C("Add sheet", "Insert a new sheet after this one", self.action_add_sheet)
        yield C("Rename sheet…", "Rename the current sheet", self.action_rename_sheet)
        yield C("Delete sheet", "Delete the current sheet", self.action_delete_sheet)
        yield C("Move sheet left", "Move the current sheet one place left", lambda: self.move_sheet(-1))
        yield C("Move sheet right", "Move the current sheet one place right", lambda: self.move_sheet(1))
        yield C("Hide sheet", "Hide the current sheet", self.hide_sheet)
        yield C("Unhide all sheets", "Show hidden sheets", self.unhide_sheets)
        yield C("Recalculate all", "Recompute every formula (F9)", self.action_recalc)
        yield C("Name manager…", "List, create, edit and delete named ranges (Ctrl+F3)", self.action_name_manager)
        yield C("Define name…", "Name the current selection (or type a name in the address box)",
                self.action_define_name)
        yield C("Create table…", "Make the selection (or the data around the cursor) a table (Ctrl+T)",
                self.create_table_dialog)
        if self._table_here() is not None:
            yield C("Table settings…", "Rename or resize the table at the cursor, totals row (Ctrl+T)",
                    self.edit_table_dialog)
            yield C("Table: toggle totals row", "Show or hide the totals row below the table",
                    lambda: self.table_command("totals"))
            yield C("Table: select data", "Select the data rows of the table at the cursor",
                    lambda: self.table_command("select"))
            yield C("Table: select column", "Select the data of the table column at the cursor",
                    lambda: self.table_command("column"))
            yield C("Table: convert to range", "Remove the table but keep its cells; table references "
                    "in formulas become normal references", lambda: self.table_command("convert"))
            yield C("Table: delete", "Delete the table at the cursor and its contents",
                    lambda: self.table_command("delete"))
        yield C("Go to…", "Jump to a cell, range or name (Ctrl+G)", self.action_goto)
        yield C("Find…", "Find text (Ctrl+F)", self.action_find)
        yield C("Replace…", "Find and replace (Ctrl+H)", self.action_replace)
        yield C("Toggle light or dark grid", "Switch between a dark grid and a white paper-like grid",
                self.action_toggle_paper)
        yield C("Keyboard help", "Show keyboard shortcuts (F1)", self.action_help)
        yield from super().get_system_commands(screen)


# ---------------------------------------------------------------------- helpers


def inp_is_formula(text: str) -> bool:
    return text.startswith("=")


def _autoclose(text: str) -> str:
    depth = 0
    in_str = False
    for ch in text:
        if ch == '"':
            in_str = not in_str
        elif not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
    if in_str:
        text += '"'
    if depth > 0:
        text += ")" * depth
    return text


def _replace_text(src: str, find: str, repl: str, case: bool, whole: bool) -> str:
    if whole:
        return repl
    if any(ch in find for ch in "*?~"):
        rx = wildcard_regex(find)
        if case:
            rx = re.compile(rx.pattern, re.S)
        return rx.sub(lambda m: repl, src)
    flags = 0 if case else re.I
    return re.sub(re.escape(find), lambda m: repl, src, flags=flags)


def _validate_float(lo: float, hi: float):
    def v(text):
        try:
            x = float(text)
        except ValueError:
            return "Enter a number"
        if not lo <= x <= hi:
            return f"Enter a number between {lo:g} and {hi:g}"
        return None

    return v


def _adjust_decimals(fmt: str, delta: int, value) -> str:
    if fmt in ("General", "", None):
        if isinstance(value, float) and not value.is_integer():
            s = f"{value:.10g}"
            cur = len(s.split(".")[1]) if "." in s else 0
        else:
            cur = 0
        n = max(0, cur + delta)
        return "0" + ("." + "0" * n if n else "")
    sections = fmt.split(";")
    out = []
    for sec in sections:
        m = re.search(r"0(\.0*)?(?=[^0]*$)|0(\.0+)?", sec)
        if "." in sec:
            i = sec.index(".")
            j = i + 1
            while j < len(sec) and sec[j] == "0":
                j += 1
            n = (j - i - 1) + delta
            sec = sec[:i] + ("." + "0" * n if n > 0 else "") + sec[j:]
        elif delta > 0 and m:
            k = sec.rfind("0")
            sec = sec[:k + 1] + "." + "0" * delta + sec[k + 1:]
        out.append(sec)
    return ";".join(out)
