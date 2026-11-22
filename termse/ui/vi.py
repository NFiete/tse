"""Vi-style keys for the grid.

Turned on with `tse --vi`, "Use vi keys" in the command palette, or `:set vi`; the choice is
remembered in the settings file. The grid passes every key to `ViKeys.handle` first; keys it
doesn't claim (Enter, Tab, F2, Delete, Shift/Ctrl+arrows, Alt+…, most Ctrl+… shortcuts) keep their
Excel meaning. Editing a cell is "insert mode": Esc there confirms the edit instead of cancelling it.
"""
from __future__ import annotations

import os
from contextlib import contextmanager

from textual import events

from ..model import structure
from ..model.refs import MAX_COL, MAX_ROW, Rect, addr

# One-key motions: name -> (kind, args). See ViKeys._target.
_STEPS = {
    "h": (0, -1), "left": (0, -1), "backspace": (0, -1),
    "l": (0, 1), "right": (0, 1), " ": (0, 1),
    "j": (1, 0), "down": (1, 0),
    "k": (-1, 0), "up": (-1, 0),
}
_JUMPS = {"w": (0, 1), "b": (0, -1), "}": (1, 0), "{": (-1, 0)}
_MOTIONS = set(_STEPS) | set(_JUMPS) | {
    "0", "^", "home", "$", "|", "G", "gg", "H", "M", "L",
    "ctrl+f", "pagedown", "ctrl+b", "pageup", "ctrl+d", "ctrl+u", "ctrl+e", "ctrl+y",
}
_PREFIXES = set("gzdycZm'`")


class ViKeys:
    def __init__(self, app):
        self.app = app
        self.count = ""
        self.pending = ""
        self.visual: str | None = None  # None | "char" | "line"
        self.line_col = 1  # the cursor column while in line-visual mode
        self.reg: tuple[str, int] | None = None  # ("row"|"col", n) when the last yank was whole rows/cols
        self.marks: dict[str, tuple] = {}
        self.last_search: tuple[str, int] | None = None  # (text, direction)

    @property
    def showcmd(self) -> str:
        return self.count + self.pending

    @property
    def mode_label(self) -> str:
        return {"char": "VISUAL", "line": "V-LINE"}.get(self.visual, "NORMAL")

    # ------------------------------------------------------------------ dispatch

    def handle(self, event: events.Key) -> bool:
        """Handle a key pressed on the grid. Returns True if it was consumed."""
        key = event.key
        if key.startswith(("alt+", "super+", "meta+")) or key == "shift+space":
            return False
        ch = event.character
        tok = ch if (event.is_printable and ch and not key.startswith("ctrl+")) else key
        if tok == "escape":
            if self.visual or self.showcmd:
                self.exit_visual()
                self._reset()
                return True
            return False
        if self.pending:
            prefix = self.pending
            self.pending = ""
            n, has = self._take_count()
            self._prefixed(prefix, tok, n, has)
            self._status()
            return True
        if tok.isdigit() and len(tok) == 1 and (tok != "0" or self.count):
            self.count += tok
            self._status()
            return True
        if tok in _PREFIXES and not (self.visual and tok in "dyc"):
            self.pending = tok
            self._status()
            return True
        n, has = self._take_count()
        handled = self._single(tok, n, has)
        self._status()
        return handled or (event.is_printable and not key.startswith("ctrl+"))

    def _take_count(self) -> tuple[int, bool]:
        has = bool(self.count)
        n = min(int(self.count), MAX_ROW) if has else 1
        self.count = ""
        return n, has

    def _reset(self):
        self.count = ""
        self.pending = ""
        self._status()

    def _status(self):
        self.app.refresh_status()

    @property
    def g(self):
        return self.app.grid

    def _single(self, tok: str, n: int, has: bool) -> bool:
        app = self.app
        g = self.g
        if tok in _MOTIONS:
            self.motion(tok, n, has)
        elif tok in ("v", "ctrl+v"):
            if self.visual == "char":
                self.exit_visual()
            else:
                self.visual = "char"
        elif tok == "V":
            if self.visual == "line":
                self.exit_visual()
            else:
                r, c = self._base()
                g.anchor = (g.anchor[0] if self.visual else r, 1)
                g.corner = (r, MAX_COL)
                self.visual = "line"
                self.line_col = c
                self._moved(r, c)
        elif self.visual and tok in ("y", "d", "x", "c", "s", "p", "P"):
            self._visual_op(tok)
        elif tok in ("i", "a", "A", "I"):
            self._edit(None)
            if tok == "I" and app.editing is not None:
                app.formula.cursor_position = 0
        elif tok in ("s", "S", "C"):
            self._edit("")
        elif tok == "=":
            self._edit("=")
        elif tok == "x":
            self.delete_cells(self._extend_rows(g.selection, n))
        elif tok == "Y":
            self.yank(self._extend_rows(g.selection, n))
        elif tok in ("p", "P"):
            self.put(tok == "P", n)
        elif tok in ("o", "O"):
            r = g.active[0] + (0 if tok == "O" else 1)
            if self._structure("insert", "row", r, n):
                g.set_cursor(r, g.active[1])
                self._edit("")
        elif tok in ("u", "ctrl+r"):
            self.visual = None
            for _ in range(n):
                app.action_undo() if tok == "u" else app.action_redo()
            g.set_cursor(*g.active)  # cursor at the change, without selecting it
        elif tok in ("/", "?"):
            self.search_prompt(1 if tok == "/" else -1)
        elif tok in ("n", "N"):
            self.search_again(1 if tok == "n" else -1)
        elif tok == ":":
            self.command_prompt()
        else:
            return False
        return True

    def _prefixed(self, prefix: str, tok: str, n: int, has: bool):
        app = self.app
        g = self.g
        cmd = prefix + tok
        if prefix == "m":
            if tok.isalpha() and len(tok) == 1:
                self.marks[tok] = (g.sheet, *g.active)
                app.set_message(f"Mark {tok} set at {addr(*g.active)}")
            return
        if prefix in "'`":
            mark = self.marks.get(tok)
            if mark is None or mark[0] not in app.book.sheets:
                app.set_message(f"Mark {tok} not set")
                return
            sheet, r, c = mark
            if sheet is not g.sheet:
                g.show_sheet(sheet)
                app.after_change(layout=True)
            self._go(r, c)
            return
        if cmd == "gg":
            self.motion("gg", n, has)
        elif cmd in ("gt", "gT"):
            if cmd == "gt" and has:
                visible = [s for s in app.book.sheets if s.ws.sheet_state == "visible"]
                if 1 <= n <= len(visible):
                    app.switch_sheet(app.book.sheets.index(visible[n - 1]))
            else:
                app._switch_sheet_relative(n if cmd == "gt" else -n)
            self.exit_visual(collapse=False)
        elif cmd in ("g-", "g+"):
            node = app.book.undo.step_time(-n if cmd == "g-" else n)
            if node is None:
                app.set_message("Already at the oldest change" if cmd == "g-" else "Already at the newest change")
            else:
                self.visual = None
                app.undo_jump(node, announce=True)
                g.set_cursor(*g.active)
        elif cmd in ("zz", "zt", "zb"):
            self.scroll_to_cursor(cmd)
        elif cmd in ("zc", "zo"):
            app.action_outline(cmd == "zc")
        elif cmd in ("zM", "zR"):
            app.outline_all(cmd == "zM")
        elif cmd in ("dd", "yy", "cc"):
            rect = self._extend_rows(g.selection, n)
            if prefix == "y":
                self.yank(rect)
            else:
                self.delete_cells(rect)
                if prefix == "c":
                    self._edit("")
        elif cmd in ("dr", "yr", "dc", "yc"):
            axis = "row" if tok == "r" else "col"
            sel = g.selection
            a = sel.r1 if axis == "row" else sel.c1
            b = max(sel.r2 if axis == "row" else sel.c2, a + n - 1)
            if prefix == "y":
                self.yank_lines(axis, a, b)
            else:
                self.delete_lines(axis, a, b)
        elif cmd == "ZZ":
            self.ex("x")
        elif cmd == "ZQ":
            self.ex("q!")
        else:
            app.bell()

    # ------------------------------------------------------------------ motions

    def _base(self) -> tuple[int, int]:
        g = self.g
        if self.visual == "char":
            return g.corner
        if self.visual == "line":
            return g.corner[0], self.line_col
        return g.active

    def _target(self, tok: str, n: int, has: bool) -> tuple[int, int]:
        g = self.g
        sheet = g.sheet
        r, c = self._base()
        if tok in _STEPS:
            dr, dc = _STEPS[tok]
            for _ in range(min(n, MAX_COL if dc else MAX_ROW)):
                nr, nc = g._step(r, c, dr, dc)
                if (nr, nc) == (r, c):
                    break
                r, c = nr, nc
            return r, c
        if tok in _JUMPS:
            dr, dc = _JUMPS[tok]
            for _ in range(min(n, 1000)):
                nr, nc = g.jump_target(r, c, dr, dc)
                if (nr, nc) == (r, c):
                    break
                r, c = nr, nc
            return r, c
        if tok in ("0", "^", "home"):
            return r, 1
        if tok == "|":
            return r, min(n, MAX_COL)
        if tok == "$":
            for cc in range(max(1, sheet.max_col), 0, -1):
                if not g._is_blank(r, cc):
                    return r, cc
            return r, c
        if tok == "G":
            return (n if has else max(1, sheet.max_row)), c
        if tok == "gg":
            return (n if has else 1), c
        if tok in ("H", "M", "L"):
            L = g.layout()
            rows = L.rows[L.n_frozen_rows_visible:] or L.rows
            if not rows:
                return r, c
            i = {"H": min(n, len(rows)) - 1, "M": (len(rows) - 1) // 2,
                 "L": max(0, len(rows) - n)}[tok]
            return rows[i], c
        page = g.visible_rows_count()
        step = {"ctrl+f": page, "pagedown": page, "ctrl+b": -page, "pageup": -page,
                "ctrl+d": max(1, page // 2), "ctrl+u": -max(1, page // 2), "ctrl+e": 1, "ctrl+y": -1}[tok] * n
        top0 = max(g.frozen()[0] + 1, g.top)
        g.top = max(g.frozen()[0] + 1, min(MAX_ROW, g.top + step))
        g._layout_key = None
        g.refresh()
        moved = g.top - top0
        if tok in ("ctrl+e", "ctrl+y"):
            # scroll the view; the cursor only moves if it would leave the screen
            rows = g.layout().rows
            body = rows[g.layout().n_frozen_rows_visible:]
            if body and r < body[0]:
                return body[0], c
            if body and r > body[-1]:
                return body[-1], c
            return r, c
        return max(1, min(MAX_ROW, r + (moved if moved else step))), c

    def motion(self, tok: str, n: int = 1, has: bool = False):
        self._go(*self._target(tok, n, has))

    def _go(self, r: int, c: int):
        g = self.g
        if self.visual == "char":
            g.set_cursor(r, c, extend=True)
        elif self.visual == "line":
            g.anchor = (g.anchor[0], 1)
            g.corner = (r, MAX_COL)
            self.line_col = c
            self._moved(r, c)
        else:
            g.set_cursor(r, c)

    def _moved(self, r: int, c: int):
        g = self.g
        g.ensure_visible(r, c)
        g.refresh()
        self.app.grid_changed()

    def scroll_to_cursor(self, where: str):
        g = self.g
        r = self._base()[0]
        n = g.visible_rows_count()
        top = {"zt": r, "zz": r - n // 2, "zb": r - n + 1}[where]
        g.top = max(g.frozen()[0] + 1, min(MAX_ROW, top))
        g._layout_key = None
        g.refresh()

    def exit_visual(self, collapse: bool = True):
        if not self.visual:
            return
        r, c = self._base()
        self.visual = None
        if collapse:
            self.g.set_cursor(r, c)
        self._status()

    # ------------------------------------------------------------------ editing

    def _edit(self, initial: str | None):
        app = self.app
        if self.visual:
            self.exit_visual(collapse=False)
        app.start_edit(initial=initial, mode="enter" if initial is not None else "edit")

    @staticmethod
    def _extend_rows(rect: Rect, n: int) -> Rect:
        return Rect(rect.r1, rect.c1, min(MAX_ROW, max(rect.r2, rect.r1 + n - 1)), rect.c2)

    def _after_op(self, r: int, c: int):
        self.visual = None
        self.g.set_cursor(r, c)

    def yank(self, rect: Rect, kind: tuple[str, int] | None = None, quiet: bool = False):
        app = self.app
        g = self.g
        app.clip = app.book.capture(g.sheet, rect)
        app.clip_cut = False
        g.clip_rect = None
        self.reg = kind
        clipped = app.clip["rect"]
        if clipped.rows * clipped.cols <= 200_000:
            app.clip_text = app.book.text_of(g.sheet, clipped)
            try:
                app.copy_to_clipboard(app.clip_text)
            except Exception:  # noqa: BLE001
                pass
        if not quiet:
            if kind is not None:
                what = f"{kind[1]} {kind[0] if kind[0] == 'row' else 'column'}{'s' if kind[1] != 1 else ''}"
            else:
                what = rect.a1()
            app.set_message(f"Yanked {what}")
        g.refresh()

    def delete_cells(self, rect: Rect):
        app = self.app
        g = self.g
        self.yank(rect, quiet=True)
        app.book.clear(g.sheet, rect)
        app.set_message(f"Deleted {rect.a1()} (p puts it back)")
        app.after_change()

    def _visual_op(self, tok: str):
        app = self.app
        g = self.g
        sel = g.selection
        line = self.visual == "line"
        top = (sel.r1, self.line_col if line else sel.c1)
        if tok in ("p", "P"):
            self.visual = None
            app.action_paste()
            g.set_cursor(*g.active)
            return
        if line and tok in ("y", "d", "x"):
            if tok == "y":
                self.yank_lines("row", sel.r1, sel.r2)
            else:
                self.delete_lines("row", sel.r1, sel.r2)
            self._after_op(*top)
            return
        if tok == "y":
            self.yank(sel)
            self._after_op(*top)
            return
        self.delete_cells(sel)
        if tok in ("c", "s"):
            self.visual = None
            self._edit("")
        else:
            self._after_op(*top)

    def yank_lines(self, axis: str, a: int, b: int, quiet: bool = False):
        rect = Rect(a, 1, b, MAX_COL) if axis == "row" else Rect(1, a, MAX_ROW, b)
        self.yank(rect, kind=(axis, b - a + 1), quiet=quiet)

    def delete_lines(self, axis: str, a: int, b: int):
        self.yank_lines(axis, a, b, quiet=True)
        if self._structure("delete", axis, a, b - a + 1):
            g = self.g
            r, c = g.active
            g.set_cursor(*((a, c) if axis == "row" else (r, a)))

    def _structure(self, op: str, axis: str, idx: int, n: int) -> bool:
        app = self.app
        try:
            (structure.insert if op == "insert" else structure.delete)(app.book, self.g.sheet, axis, idx, n)
        except ValueError as e:
            app.notify(str(e), severity="error")
            return False
        what = "row" if axis == "row" else "column"
        app.set_message(f"{'Inserted' if op == 'insert' else 'Deleted'} {n} {what}{'s' if n != 1 else ''}")
        app.after_change(layout=True)
        return True

    @contextmanager
    def one_undo(self, label: str):
        """Merge the undo steps pushed inside the block into a single step."""
        tree = self.app.book.undo
        mark = tree.current
        try:
            yield
        finally:
            tree.squash(mark, label)

    def put(self, before: bool, n: int = 1):
        """p / P. Whole rows or columns (from yr/dr/yc/dc/V) are put as new rows below (above) the
        cursor, or new columns right (left) of it; anything else is pasted at the cursor."""
        app = self.app
        g = self.g
        if app.clip is None or self.reg is None or app.clip["sheet"].book is not app.book:
            for _ in range(n):
                app.action_paste()
            g.set_cursor(*g.active)  # don't leave the pasted range selected
            return
        axis, count = self.reg
        r, c = g.active
        at = (r if axis == "row" else c) + (0 if before else 1)
        total = count * n
        src = app.clip["rect"]
        with self.one_undo("Put " + ("rows" if axis == "row" else "columns")):
            if not self._structure("insert", axis, at, total):
                return
            for i in range(n):
                try:
                    if axis == "row":
                        app.book.paste(app.clip, g.sheet, at + i * count, src.c1)
                    else:
                        app.book.paste(app.clip, g.sheet, src.r1, at + i * count)
                except ValueError as e:
                    app.notify(str(e), severity="error")
                    break
        g.set_cursor(at, c) if axis == "row" else g.set_cursor(r, at)
        what = "row" if axis == "row" else "column"
        app.set_message(f"Put {total} {what}{'s' if total != 1 else ''}")
        app.after_change(layout=True)

    # ------------------------------------------------------------------ search and : commands

    def search_prompt(self, direction: int):
        from .dialogs import CommandLine

        def done(text):
            if text:
                self.last_search = (text, direction)
                self._find(text, direction)

        self.app.push_screen(CommandLine("/" if direction > 0 else "?"), done)

    def search_again(self, d: int):
        if self.last_search is None:
            self.app.set_message("No previous search")
            return
        text, direction = self.last_search
        self._find(text, direction * d)

    def _find(self, text: str, direction: int):
        app = self.app
        self.exit_visual(collapse=False)
        found = app.find_next(text, case=False, whole=False, formulas=True, direction=direction)
        app.set_message(found or f"Not found: {text}")

    def command_prompt(self):
        from .dialogs import CommandLine

        self.app.push_screen(CommandLine(":"), lambda text: self.ex(text) if text else None)

    def ex(self, text: str):
        """Run a : command."""
        app = self.app
        book = app.book
        text = text.strip().lstrip(":").strip()
        cmd, _, arg = text.partition(" ")
        arg = arg.strip()
        bang = cmd.endswith("!")
        cmd = cmd.rstrip("!")
        if not cmd:
            return
        if cmd.isdigit():
            self._go(max(1, min(MAX_ROW, int(cmd))), self._base()[1])
            return
        if cmd in ("w", "write", "wq", "x", "xit", "exit"):
            quit_after = cmd != "w" and cmd != "write"
            if cmd in ("x", "xit", "exit") and not book.modified and not arg:
                app.exit()
                return
            if arg:
                path = os.path.abspath(os.path.expanduser(arg))
                if not os.path.splitext(path)[1]:
                    path += ".xlsx"
                if os.path.exists(path) and path != book.path and not bang:
                    app.notify(f"{arg} exists (add ! to overwrite)", severity="error", markup=False)
                    return
                app._save_to(path, confirmed=bang)
                if book.path is None:
                    book.path = path
            elif book.path is None:
                app.action_save_as()
                return
            else:
                app.action_save()
            if quit_after and not book.modified:
                app.exit()
        elif cmd in ("q", "quit", "qa", "qall", "quitall"):
            if bang or not book.modified:
                app.exit()
            else:
                app.notify("No write since last change (add ! to override)", severity="error")
        elif cmd in ("e", "edit"):
            if not arg:
                if bang and book.path and os.path.exists(book.path):
                    app._open(book.path)
                else:
                    app.action_open()
                return
            path = os.path.abspath(os.path.expanduser(arg))
            if not os.path.exists(path):
                app.notify(f"No such file: {path}", severity="error", markup=False)
            elif bang:
                app._open(path)
            else:
                app._confirm_discard(lambda: app._open(path))
        elif cmd in ("sav", "saveas"):
            if arg:
                app.action_save_as(default=os.path.abspath(os.path.expanduser(arg)))
            else:
                app.action_save_as()
        elif cmd == "set":
            opt = arg.replace(" ", "")
            if opt in ("vi", "novi"):
                app.set_vi(opt == "vi")
            else:
                app.notify(f"Unknown option: {arg}  (try :set novi)", severity="error", markup=False)
        elif cmd in ("tabn", "tabnext", "bn", "bnext"):
            app._switch_sheet_relative(1)
        elif cmd in ("tabp", "tabprevious", "tabN", "tabNext", "bp", "bprevious"):
            app._switch_sheet_relative(-1)
        elif cmd in ("h", "help"):
            app.action_help()
        elif cmd in ("undol", "undolist", "undotree", "undohistory"):
            app.action_undo_tree()
        elif cmd == "sort":
            app.sort_selection(descending=bang)
        elif cmd in ("noh", "nohlsearch"):
            app.set_message("")
        else:
            target = app._resolve_goto(text)
            if target is None:
                app.notify(f"Not a command, cell or name: {text}", severity="error", markup=False)
                return
            sheet, rect = target
            if sheet is not self.g.sheet:
                self.g.show_sheet(sheet)
            self.visual = None
            app._select(rect)
            app.after_change(layout=True)

