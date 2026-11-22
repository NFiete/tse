"""Formula bar, sheet tabs and status bar."""
from __future__ import annotations

import re

from rich.cells import cell_len
from rich.highlighter import Highlighter
from rich.text import Text
from textual import events
from textual.widget import Widget
from textual.widgets import Input, Static

from ..formula.parser import REFT, SREFT, FormulaSyntaxError, Unsupported, tokenize


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
            if t.kind == REFT or t.kind == SREFT:
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
    """The sheet tab strip. When the tabs don't fit it scrolls: ◀/▶ markers show that tabs are cut off
    (click them or use the mouse wheel to scroll) and switching sheets brings the active tab into view."""

    DEFAULT_CSS = """
    SheetTabs { height: 1; width: 1fr; background: $panel; }
    """
    LEFT = "◀ "
    RIGHT = " ▶"
    ADD = " + "

    def __init__(self, **kw):
        super().__init__(**kw)
        self._spans: list[tuple[int, int, int]] = []  # (x0, x1, sheet index); -1 add, -2/-3 scroll left/right
        self.offset_idx = 0  # index into the visible tabs of the first one shown
        self._follow_key = None
        self._max_offset = 0

    def _tabs(self):
        book = self.app.book
        current = self.app.grid.sheet
        return [(i, s) for i, s in enumerate(book.sheets) if s.ws.sheet_state == "visible" or s is current]

    def _fit(self, widths: list[int], o: int, area: int) -> int:
        """How many tabs starting at `o` fit in `area` cells (at least one), allowing for the arrows."""
        left = cell_len(self.LEFT) if o > 0 else 0
        rest = widths[o:]
        if left + sum(rest) <= area:
            return len(rest)
        room = area - left - cell_len(self.RIGHT)
        n = used = 0
        for w in rest:
            if used + w > room:
                break
            used += w
            n += 1
        return max(1, n)

    def scroll_tabs(self, d: int) -> None:
        o = max(0, min(self._max_offset, self.offset_idx + d))
        if o != self.offset_idx:
            self.offset_idx = o
            self.refresh()

    def render(self) -> Text:
        app = self.app
        book = app.book
        t = Text()
        self._spans = []
        if book is None:
            return t
        current = app.grid.sheet
        tabs = self._tabs()
        labels = [f" {s.name} " for _, s in tabs]
        widths = [cell_len(lb) + 1 for lb in labels]  # + separator
        hidden = sum(1 for s in book.sheets if s.ws.sheet_state != "visible")
        hidden_txt = f"  ({hidden} hidden)" if hidden else ""
        width = self.size.width or 10**6
        area = width - cell_len(self.ADD) - cell_len(hidden_txt)
        if area < 20 and hidden_txt:
            hidden_txt = ""
            area = width - cell_len(self.ADD)
        n = len(tabs)
        # The furthest-right offset that still leaves no empty space after the last tab.
        mx = n - 1 if n else 0
        while mx > 0 and (mx - 1) + self._fit(widths, mx - 1, area) >= n:
            mx -= 1
        self._max_offset = mx
        o = max(0, min(self.offset_idx, mx))
        key = (current, tuple(lb for lb in labels), width)
        if key != self._follow_key:
            # The active sheet, the tab names or the width changed: bring the active tab into view.
            self._follow_key = key
            ci = next((k for k, (_, s) in enumerate(tabs) if s is current), 0)
            if ci < o:
                o = ci
            while o < ci and ci >= o + self._fit(widths, o, area):
                o += 1
            o = min(o, mx)
        self.offset_idx = o
        shown = self._fit(widths, o, area) if n else 0
        x = 0
        if o > 0:
            t.append(self.LEFT, style="bold")
            self._spans.append((x, x + cell_len(self.LEFT), -2))
            x += cell_len(self.LEFT)
        for k in range(o, o + shown):
            i, s = tabs[k]
            label = labels[k]
            on = s is current
            style = "bold reverse" if on else "dim"
            if s.ws.sheet_properties.tabColor is not None and getattr(s.ws.sheet_properties.tabColor, "rgb", None):
                rgb = s.ws.sheet_properties.tabColor.rgb
                if isinstance(rgb, str) and len(rgb) >= 6:
                    style += f" underline #{rgb[-6:]}"
            t.append(label, style=style)
            t.append("│", style="dim")
            self._spans.append((x, x + cell_len(label), i))
            x += widths[k]
        if o + shown < n:
            t.append(self.RIGHT, style="bold")
            self._spans.append((x, x + cell_len(self.RIGHT), -3))
            x += cell_len(self.RIGHT)
        t.append(self.ADD, style="bold")
        self._spans.append((x, x + cell_len(self.ADD), -1))
        if hidden_txt:
            t.append(hidden_txt, style="dim")
        return t

    def on_click(self, event: events.Click) -> None:
        for x0, x1, idx in self._spans:
            if x0 <= event.x < x1:
                if idx == -1:
                    self.app.action_add_sheet()
                elif idx == -2:
                    self.scroll_tabs(-1)
                elif idx == -3:
                    self.scroll_tabs(1)
                else:
                    self.app.switch_sheet(idx)
                    if event.chain == 2:
                        self.app.action_rename_sheet()
                return

    def on_mouse_scroll_down(self, event: events.MouseScrollDown) -> None:
        event.stop()
        self.scroll_tabs(1)

    def on_mouse_scroll_up(self, event: events.MouseScrollUp) -> None:
        event.stop()
        self.scroll_tabs(-1)

    def on_mouse_scroll_right(self, event: events.MouseScrollRight) -> None:
        event.stop()
        self.scroll_tabs(1)

    def on_mouse_scroll_left(self, event: events.MouseScrollLeft) -> None:
        event.stop()
        self.scroll_tabs(-1)


class StatusBar(Static):
    DEFAULT_CSS = """
    StatusBar { height: 1; width: 1fr; background: $boost; color: $text-muted; }
    """

    def show(self, mode: str, message: str, stats: str, file_label: str, keys: str = ""):
        w = self.size.width or 80
        left = Text()
        left.append(f" {mode} ", style="bold reverse")
        left.append(" ")
        left.append(message)
        right = Text()
        if keys:
            right.append(keys + "   ", style="bold")
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
