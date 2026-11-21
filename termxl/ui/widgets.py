"""Formula bar, sheet tabs and status bar."""
from __future__ import annotations

import re

from rich.highlighter import Highlighter
from rich.text import Text
from textual import events
from textual.widget import Widget
from textual.widgets import Input, Static

from ..formula.parser import REFT, FormulaSyntaxError, Unsupported, tokenize


class FormulaHighlighter(Highlighter):
    """Colors references in a formula with the same colors the grid uses to outline them."""

    def __init__(self, app):
        super().__init__()
        self._app = app

    def highlight(self, text: Text) -> None:
        s = text.plain
        if not s.startswith("="):
            return
        try:
            toks = tokenize(s[1:])
        except (FormulaSyntaxError, Unsupported):
            return
        colors = self._app.grid.palette.ref_colors
        i = 0
        for t in toks:
            if t.kind == REFT:
                text.stylize(f"bold #{colors[i % len(colors)]}", t.start + 1, t.end + 1)
                i += 1
            elif t.kind == "STR":
                text.stylize("#CE9178", t.start + 1, t.end + 1)
            elif t.kind == "FUNC":
                text.stylize("#DCDCAA", t.start + 1, t.end)


class FormulaInput(Input):
    """The formula bar editor. Key handling is delegated to the app's edit logic."""

    async def _on_key(self, event: events.Key) -> None:
        app = self.app
        if app.editing is None:
            app.start_edit(mode="edit", focus=False)
        if app.edit_key(event):
            event.stop()
            event.prevent_default()
            return
        if event.is_printable or event.key in ("backspace", "delete"):
            app.end_point()
        await super()._on_key(event)

    def _on_paste(self, event: events.Paste) -> None:
        if self.app.editing is None:
            self.app.start_edit(mode="edit", focus=False)
        super()._on_paste(event)

    def on_focus(self, event: events.Focus) -> None:
        if self.app.editing is None:
            self.app.start_edit(mode="edit", focus=False)


class SheetTabs(Widget):
    DEFAULT_CSS = """
    SheetTabs { height: 1; width: 1fr; background: $panel; }
    """

    def __init__(self, **kw):
        super().__init__(**kw)
        self._spans: list[tuple[int, int, int]] = []  # (x0, x1, sheet index)
        self.offset_idx = 0

    def render(self) -> Text:
        app = self.app
        book = app.book
        t = Text()
        self._spans = []
        if book is None:
            return t
        x = 0
        current = app.grid.sheet
        visible = [(i, s) for i, s in enumerate(book.sheets) if s.ws.sheet_state == "visible" or s is current]
        for i, s in visible:
            label = f" {s.name} "
            on = s is current
            style = "bold reverse" if on else "dim"
            if s.ws.sheet_properties.tabColor is not None and getattr(s.ws.sheet_properties.tabColor, "rgb", None):
                rgb = s.ws.sheet_properties.tabColor.rgb
                if isinstance(rgb, str) and len(rgb) >= 6:
                    style += f" underline #{rgb[-6:]}"
            t.append(label, style=style)
            t.append("│", style="dim")
            self._spans.append((x, x + len(label), i))
            x += len(label) + 1
        hidden = sum(1 for s in book.sheets if s.ws.sheet_state != "visible")
        t.append(" + ", style="bold")
        self._spans.append((x, x + 3, -1))
        if hidden:
            t.append(f"  ({hidden} hidden)", style="dim")
        return t

    def on_click(self, event: events.Click) -> None:
        for x0, x1, idx in self._spans:
            if x0 <= event.x < x1:
                if idx == -1:
                    self.app.action_add_sheet()
                else:
                    self.app.switch_sheet(idx)
                    if event.chain == 2:
                        self.app.action_rename_sheet()
                return


class StatusBar(Static):
    DEFAULT_CSS = """
    StatusBar { height: 1; width: 1fr; background: $boost; color: $text-muted; }
    """

    def show(self, mode: str, message: str, stats: str, file_label: str):
        w = self.size.width or 80
        left = Text()
        left.append(f" {mode} ", style="bold reverse")
        left.append(" ")
        left.append(message)
        right = Text()
        if stats:
            right.append(stats + "   ")
        right.append(file_label, style="bold")
        right.append(" ")
        pad = max(1, w - left.cell_len - right.cell_len)
        if left.cell_len + right.cell_len + 1 > w:
            left.truncate(max(0, w - right.cell_len - 1), overflow="ellipsis")
            pad = max(1, w - left.cell_len - right.cell_len)
        self.update(Text.assemble(left, " " * pad, right))


_REF_AT = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")


def cycle_absolute(text: str, cursor: int) -> tuple[str, int]:
    """F4: cycle A1 -> $A$1 -> A$1 -> $A1 -> A1 for the reference at/just before the cursor."""
    body = text
    # find the reference token touching the cursor
    m = None
    for mm in re.finditer(r"\$?[A-Za-z]{1,3}\$?\d+", body):
        if mm.start() <= cursor <= mm.end():
            m = mm
    if m is None:
        return text, cursor
    ref = m.group()
    p = _REF_AT.match(ref)
    if not p:
        return text, cursor
    ac, col, ar, row = p.groups()
    state = (bool(ac), bool(ar))
    nxt = {(False, False): (True, True), (True, True): (False, True), (False, True): (True, False),
           (True, False): (False, False)}[state]
    new = f"{'$' if nxt[0] else ''}{col}{'$' if nxt[1] else ''}{row}"
    out = body[:m.start()] + new + body[m.end():]
    return out, m.start() + len(new)


class NameBox(Input):
    """The cell-address box. Type an address or name to jump there, or a new name to
    define it for the current selection (like Excel's Name Box)."""

    async def _on_key(self, event: events.Key) -> None:
        if event.key == "escape":
            event.stop()
            event.prevent_default()
            self.app.namebox_done()
            return
        await super()._on_key(event)

    _select_on_mouse_up = False

    def on_focus(self, event: events.Focus) -> None:
        if self.app.editing is not None and not self.app.commit_edit(move=None):
            self.app.formula.focus()
            return
        self.select_all()
        self._select_on_mouse_up = True

    async def _on_mouse_up(self, event: events.MouseUp) -> None:
        await super()._on_mouse_up(event)
        if self._select_on_mouse_up:
            # the click that focused the box selects everything, as in Excel
            self._select_on_mouse_up = False
            self.select_all()

    def on_blur(self, event: events.Blur) -> None:
        self._select_on_mouse_up = False
        self.app.grid_changed()
