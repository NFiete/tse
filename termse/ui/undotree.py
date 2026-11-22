"""Undo history panel: browse the undo tree and jump to any earlier state."""
from __future__ import annotations

import time

from rich.text import Text
from textual import on
from textual.app import ComposeResult
from textual.binding import Binding
from textual.containers import Vertical
from textual.screen import ModalScreen
from textual.widgets import Label, OptionList, Static
from textual.widgets.option_list import Option

from ..model.refs import MAX_COL, MAX_ROW, col_letter
from ..model.undo import UndoNode, graph
from .dialogs import DIALOG_CSS


def _when(t: float) -> str:
    if time.time() - t < 20 * 3600:
        return time.strftime("%H:%M:%S", time.localtime(t))
    return time.strftime("%b %d %H:%M", time.localtime(t))


def where_text(where, with_sheet: bool) -> str:
    """'B3', 'B:D', '3:5' or 'Sheet2!A1:C4' for an UndoEntry's where."""
    if not where:
        return ""
    name, r = where
    if r.r1 <= 1 and r.r2 >= MAX_ROW and not (r.c1 <= 1 and r.c2 >= MAX_COL):
        s = f"{col_letter(r.c1)}:{col_letter(r.c2)}"
    elif r.c1 <= 1 and r.c2 >= MAX_COL and not (r.r1 <= 1 and r.r2 >= MAX_ROW):
        s = f"{r.r1}:{r.r2}"
    else:
        s = r.a1()
    if with_sheet:
        s = (f"'{name}'" if not name.replace("_", "").isalnum() else name) + "!" + s
    return s


class UndoTreeScreen(ModalScreen):
    """A panel on the right listing every state of the workbook, newest first. Moving the highlight
    previews that state in the grid; Enter keeps it, Esc goes back to where you were."""

    DEFAULT_CSS = DIALOG_CSS + """
    UndoTreeScreen { align: right top; background: $background 0%; }
    UndoTreeScreen #dialog { width: 60; max-width: 70%; height: 100%; max-height: 100%; padding: 0 1; }
    UndoTreeScreen #dialog Label.title { margin-bottom: 0; }
    UndoTreeScreen OptionList { height: 1fr; border: none; padding: 0; background: $surface; }
    UndoTreeScreen OptionList:focus { border: none; }
    """
    BINDINGS = [
        Binding("escape,q", "cancel", "Cancel", show=False),
        Binding("j", "move(1)", show=False),
        Binding("k", "move(-1)", show=False),
    ]

    def __init__(self, app_ref):
        super().__init__()
        self._app = app_ref
        self._tree = app_ref.book.undo
        self._rows: list[tuple[str, UndoNode | None]] = []

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label("Undo history", classes="title")
            yield OptionList(id="states")
            yield Static("↑↓ preview  Enter keep  Esc cancel", classes="msg")

    def on_mount(self):
        ol = self.query_one(OptionList)
        self._rows = graph(self._tree)
        ol.add_options([Option(self._prompt(p, n), id=str(n.seq) if n is not None else None, disabled=n is None)
                        for p, n in self._rows])
        ol.highlighted = self._index(self._tree.current)
        ol.focus()

    def _index(self, node: UndoNode) -> int:
        return next(i for i, (_, n) in enumerate(self._rows) if n is node)

    def _prompt(self, prefix: str, node: UndoNode | None) -> Text:
        if node is None:
            return Text(prefix, style="dim")
        cur = node is self._tree.current
        t = Text()
        t.append(prefix, style="bold" if cur else "dim")
        t.append(f"{node.seq:>3} ", style="dim")
        t.append(node.label, style="bold" if cur else "")
        if node.entry is not None and node.entry.where:
            multi = len(self._app.book.sheets) > 1
            t.append(" " + where_text(node.entry.where, multi), style="cyan")
        if node is self._tree.saved:
            t.append(" saved", style="green")
        t.append(f"  {_when(node.time)}", style="dim")
        return t

    def _redraw(self, nodes):
        ol = self.query_one(OptionList)
        for n in nodes:
            i = self._index(n)
            prefix = self._rows[i][0]
            # the ● marker is the only thing in the graph that depends on the current state
            prefix = prefix.replace("●", "○") if n is not self._tree.current else prefix.replace("○", "●", 1)
            self._rows[i] = (prefix, n)
            ol.replace_option_prompt_at_index(i, self._prompt(prefix, n))

    def _jump(self, node: UndoNode):
        before = self._tree.current
        if node is before:
            return
        self._app.undo_jump(node)
        self._redraw([before, self._tree.current])

    @on(OptionList.OptionHighlighted)
    def _highlighted(self, event: OptionList.OptionHighlighted):
        node = self._rows[event.option_index][1]
        if node is not None:
            self._jump(node)

    @on(OptionList.OptionSelected)
    def _selected(self, event: OptionList.OptionSelected):
        node = self._rows[event.option_index][1]
        if node is not None:
            self._jump(node)
        self.dismiss(True)

    def action_move(self, d: int):
        ol = self.query_one(OptionList)
        ol.action_cursor_down() if d > 0 else ol.action_cursor_up()

    def action_cancel(self):
        self.dismiss(False)
