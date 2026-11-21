"""Modal dialogs: text prompts, choices, find/replace and help."""
from __future__ import annotations

import glob
import os

from textual import events, on
from textual.app import ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal, Vertical, VerticalScroll
from textual.screen import ModalScreen
from textual.widgets import Button, Checkbox, Input, Label, Static

DIALOG_CSS = """
ModalScreen { align: center middle; background: $background 60%; }
#dialog {
    width: auto; min-width: 50; max-width: 90%; height: auto; max-height: 90%;
    border: round $accent; background: $surface; padding: 1 2;
}
#dialog Label.title { text-style: bold; margin-bottom: 1; }
#dialog .error { color: $error; height: auto; }
#dialog .msg { color: $text-muted; height: auto; }
#dialog Horizontal.buttons { height: auto; margin-top: 1; align: right middle; }
#dialog Horizontal.buttons Button { margin-left: 1; min-width: 10; }
#dialog Input { width: 60; }
#dialog Horizontal.opts { height: auto; }
"""


class PathInput(Input):
    """Input with shell-like Tab completion of file paths."""

    async def _on_key(self, event: events.Key) -> None:
        if event.key == "tab":
            event.stop()
            event.prevent_default()
            self._complete()
            return
        await super()._on_key(event)

    def _complete(self):
        text = os.path.expanduser(self.value)
        matches = sorted(glob.glob(glob.escape(text) + "*"))
        if not matches:
            self.app.bell()
            return
        common = os.path.commonprefix(matches)
        if len(matches) == 1 and os.path.isdir(common):
            common += os.sep
        if self.value.startswith("~"):
            home = os.path.expanduser("~")
            if common.startswith(home):
                common = "~" + common[len(home):]
        self.value = common
        self.cursor_position = len(common)
        if len(matches) > 1:
            names = [os.path.basename(m.rstrip(os.sep)) + ("/" if os.path.isdir(m) else "") for m in matches[:12]]
            self.notify("  ".join(names) + ("  …" if len(matches) > 12 else ""), timeout=4)


class InputDialog(ModalScreen):
    DEFAULT_CSS = DIALOG_CSS
    BINDINGS = [Binding("escape", "cancel", show=False)]

    def __init__(self, title: str, prompt: str = "", value: str = "", validate=None, path: bool = False,
                 placeholder: str = "", help_text: str = ""):
        super().__init__()
        self._title, self._prompt, self._value = title, prompt, value
        self._validate = validate
        self._path = path
        self._placeholder = placeholder
        self._help = help_text

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label(self._title, classes="title")
            if self._prompt:
                yield Label(self._prompt)
            cls = PathInput if self._path else Input
            yield cls(value=self._value, placeholder=self._placeholder, id="value", select_on_focus=not self._path)
            if self._help:
                yield Static(self._help, classes="msg")
            yield Static("", classes="error", id="error")
            with Horizontal(classes="buttons"):
                yield Button("Cancel", id="cancel")
                yield Button("OK", variant="primary", id="ok")

    def on_mount(self):
        inp = self.query_one("#value", Input)
        inp.focus()
        inp.cursor_position = len(inp.value)

    @on(Input.Submitted)
    @on(Button.Pressed, "#ok")
    def submit(self):
        v = self.query_one("#value", Input).value
        if self._validate is not None:
            err = self._validate(v)
            if err:
                self.query_one("#error", Static).update(err)
                return
        self.dismiss(v)

    @on(Button.Pressed, "#cancel")
    def action_cancel(self):
        self.dismiss(None)


class ChoiceDialog(ModalScreen):
    DEFAULT_CSS = DIALOG_CSS
    BINDINGS = [Binding("escape", "cancel", show=False),
                Binding("left", "app.focus_previous", show=False),
                Binding("right", "app.focus_next", show=False)]

    def __init__(self, title: str, message: str, choices: list[tuple[str, str]], default: str | None = None):
        super().__init__()
        self._title, self._message, self._choices = title, message, choices
        self._default = default or choices[-1][0]

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label(self._title, classes="title")
            yield Static(self._message)
            with Horizontal(classes="buttons"):
                for key, label in self._choices:
                    yield Button(label, id=f"c_{key}", variant="primary" if key == self._default else "default")

    def on_mount(self):
        self.query_one(f"#c_{self._default}", Button).focus()

    def on_key(self, event: events.Key):
        if event.character:
            ch = event.character.lower()
            for key, label in self._choices:
                if label.lower().lstrip("&").startswith(ch):
                    event.stop()
                    self.dismiss(key)
                    return

    @on(Button.Pressed)
    def pressed(self, event: Button.Pressed):
        self.dismiss(event.button.id[2:])

    def action_cancel(self):
        self.dismiss(None)


class FindDialog(ModalScreen):
    """Find / replace. Talks to the app through find_next / replace_current / replace_all."""

    DEFAULT_CSS = DIALOG_CSS + """
    FindDialog { align: right top; background: transparent; }
    FindDialog #dialog { margin: 3 2; }
    """
    BINDINGS = [Binding("escape", "close", show=False),
                Binding("shift+enter", "find(-1)", show=False)]

    def __init__(self, replace: bool = False, text: str = ""):
        super().__init__()
        self._replace = replace
        self._text = text

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label("Find and Replace" if self._replace else "Find", classes="title")
            yield Input(value=self._text, placeholder="Find what (wildcards * ? allowed)", id="find")
            if self._replace:
                yield Input(placeholder="Replace with", id="replace")
            with Horizontal(classes="opts"):
                yield Checkbox("Match case", id="case")
                yield Checkbox("Entire cell", id="whole")
                yield Checkbox("Look in formulas", value=True, id="formulas")
            yield Static("Enter: next   Shift+Enter: previous   Esc: close", classes="msg", id="msg")
            with Horizontal(classes="buttons"):
                if self._replace:
                    yield Button("Replace All", id="all")
                    yield Button("Replace", id="one")
                yield Button("Find Next", variant="primary", id="next")
                yield Button("Close", id="close")

    def on_mount(self):
        self.query_one("#find", Input).focus()

    def _opts(self):
        return dict(text=self.query_one("#find", Input).value,
                    case=self.query_one("#case", Checkbox).value,
                    whole=self.query_one("#whole", Checkbox).value,
                    formulas=self.query_one("#formulas", Checkbox).value)

    def _msg(self, text: str):
        self.query_one("#msg", Static).update(text)

    @on(Input.Submitted)
    @on(Button.Pressed, "#next")
    def find_next(self):
        self.action_find(1)

    def action_find(self, direction: int = 1):
        o = self._opts()
        if not o["text"]:
            return
        found = self.app.find_next(direction=direction, **o)
        self._msg(found or "Not found")

    @on(Button.Pressed, "#one")
    def replace_one(self):
        o = self._opts()
        msg = self.app.replace_current(self.query_one("#replace", Input).value, **o)
        self._msg(msg)

    @on(Button.Pressed, "#all")
    def replace_all(self):
        o = self._opts()
        n = self.app.replace_all(self.query_one("#replace", Input).value, **o)
        self._msg(f"Replaced {n} occurrence{'s' if n != 1 else ''}")

    @on(Button.Pressed, "#close")
    def action_close(self):
        self.dismiss(None)


HELP_TEXT = """\
[b]Moving[/b]
  Arrows  PgUp/PgDn  Home       move
  Ctrl+Arrows                   jump to the edge of the data
  Shift+Arrows  or click-drag   select
  Ctrl+Home / Ctrl+End          first / last used cell
  Ctrl+Space / Shift+Space      select column / row
  Ctrl+A                        select all
  Ctrl+G                        go to a cell, range or name
  Ctrl+PgUp / Ctrl+PgDn         previous / next sheet

[b]Editing[/b]
  type                          replace the cell's contents
  F2  or double-click           edit the cell
  Enter / Tab                   confirm and move   (Esc cancels)
  arrows while typing =...      point at cells to insert references (Shift extends)
  F4                            cycle $ absolute / relative on the reference
  Delete / Backspace            clear / clear and edit
  Ctrl+C  Ctrl+X  Ctrl+V        copy, cut, paste (formulas adjust like Excel)
  Ctrl+D / Ctrl+R               fill down / right
  Ctrl+Z / Ctrl+Y               undo / redo
  Ctrl+B  Ctrl+I  Ctrl+U        bold, italic, underline  (Alt+B / Alt+I / Alt+U too)
  Alt+=  /  Alt+-               insert / delete rows or columns
  Alt+Left / Alt+Right          narrower / wider column
  Ctrl+F / Ctrl+H               find / replace
  F9                            recalculate everything

[b]Files[/b]
  Ctrl+S / F12                  save / save as  (.xlsx, .xlsm, .csv)
  Ctrl+O / Ctrl+N               open / new
  Ctrl+Q                        quit

[b]Everything else[/b]
  Ctrl+P opens the command palette: number formats, alignment, colors, column widths,
  hide/unhide, freeze panes, sorting, sheet management, light/dark grid and more.
"""


class HelpScreen(ModalScreen):
    DEFAULT_CSS = DIALOG_CSS + """
    HelpScreen #dialog { height: 90%; }
    HelpScreen #helpbody { height: 1fr; }
    """
    BINDINGS = [Binding("escape,f1,q", "close", show=False)]

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label("term_xl — keyboard reference", classes="title")
            with VerticalScroll(id="helpbody"):
                yield Static(HELP_TEXT)
            with Horizontal(classes="buttons"):
                yield Button("Close", variant="primary", id="close")

    def on_mount(self):
        self.query_one("#helpbody").focus()

    @on(Button.Pressed)
    def action_close(self):
        self.dismiss(None)
