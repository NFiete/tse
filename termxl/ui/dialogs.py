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


class _CommandInput(Input):
    async def _on_key(self, event: events.Key) -> None:
        if event.key == "backspace" and not self.value:
            event.stop()
            event.prevent_default()
            self.screen.dismiss(None)
            return
        await super()._on_key(event)


class CommandLine(ModalScreen):
    """A vi-style one-line prompt at the bottom of the screen (for : commands and / searches)."""

    DEFAULT_CSS = """
    CommandLine { align: left bottom; background: transparent; }
    CommandLine #cmdrow { height: 1; width: 100%; background: $surface; }
    CommandLine #prompt { width: 1; height: 1; text-style: bold; }
    CommandLine Input { height: 1; width: 1fr; border: none; padding: 0; background: $surface; }
    CommandLine Input:focus { border: none; }
    """
    BINDINGS = [Binding("escape", "cancel", show=False)]

    def __init__(self, prompt: str = ":"):
        super().__init__()
        self._prompt = prompt

    def compose(self) -> ComposeResult:
        with Horizontal(id="cmdrow"):
            yield Static(self._prompt, id="prompt")
            yield _CommandInput(id="cmd")

    def on_mount(self):
        self.query_one("#cmd", Input).focus()

    @on(Input.Submitted)
    def submit(self, event: Input.Submitted):
        self.dismiss(event.value)

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
  address box (top left)        type a cell or name to go there,
                                or a new name to name the selection
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
  Alt+Shift+Right / Left        group / ungroup rows or columns
  Alt+Shift+Up / Down           collapse / expand the group at the cursor
                                (or click the [-] / [+] buttons and level numbers)
  Ctrl+F / Ctrl+H               find / replace
  F9                            recalculate everything
  Ctrl+F3                       name manager: create, edit, delete named ranges
  Ctrl+T                        make the selection a table, or edit the table at the cursor
                                (name, range, totals row, convert to range); formulas can use
                                Table1\\[Column], Table1\\[#Totals] and \\[@Column] inside the table

[b]Files[/b]
  Ctrl+S / F12                  save / save as  (.xlsx, .xlsm, .csv)
  Ctrl+O / Ctrl+N               open / new
  Ctrl+Q                        quit

[b]Everything else[/b]
  Ctrl+P opens the command palette: number formats, alignment, colors, column widths,
  hide/unhide, grouping, freeze panes, sorting, sheet management, light/dark grid and more.
"""


VI_HELP_TEXT = """\
[b]Vi keys[/b]  (on; turn off with :set novi or "Use Excel keys" in Ctrl+P)
  h j k l  arrows  Space        move (a count repeats: 5j)
  w b / { }                     jump to the edge of the data right/left / up/down
  0 ^ / $                       first column / last used column in the row
  gg / G  (5G or :5)            first / last used row (or row 5);  3| goes to column C
  H M L                         top / middle / bottom of the screen
  Ctrl+F Ctrl+B / Ctrl+D Ctrl+U page down/up / half page;  Ctrl+E Ctrl+Y scroll
  zz zt zb                      scroll the cursor to the middle / top / bottom
  v  V  Ctrl+V                  select cells / whole rows / cells (Esc ends)
  i a A I                       edit the cell (I puts the cursor at the start)
  s S C cc  =                   replace the cell's contents / start a formula
  Esc while editing             confirm the edit (Esc again after an error discards it)
  o / O                         insert a row below / above and start typing in it
  x  dd                         delete the cells' contents (p puts them back)
  dr dc  /  yr yc               delete / yank whole rows or columns (3dr: three rows)
  yy  Y                         yank the cell or selection
  p  P                          put; yanked rows/columns are inserted below/above (right/left)
  u  Ctrl+R                     undo / redo
  /text  ?text  n  N            search down / up, next / previous match
  gt gT  (3gt)                  next / previous sheet (sheet 3)
  zc zo / zM zR                 collapse / expand the group at the cursor / all groups
  ma  'a                        set mark a / jump to it
  :w \\[file]  :q  :wq  :x  :q!   save, quit (ZZ = :x, ZQ = :q!);  :e file, :sav file
  :B12  :Sheet2!A1  :Rate       go to a cell, range or name;  :sort  :sort! sort the selection
  Ctrl+… shortcuts not listed here keep their Excel meaning (Ctrl+Z, Ctrl+S, Ctrl+C, Ctrl+G…)

"""


class HelpScreen(ModalScreen):
    DEFAULT_CSS = DIALOG_CSS + """
    HelpScreen #dialog { height: 90%; }
    HelpScreen #helpbody { height: 1fr; }
    """
    BINDINGS = [Binding("escape,f1,q", "close", show=False)]

    def __init__(self, vi: bool = False):
        super().__init__()
        self._vi = vi

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label("term_xl — keyboard reference", classes="title")
            with VerticalScroll(id="helpbody"):
                yield Static((VI_HELP_TEXT if self._vi else "") + HELP_TEXT)
            with Horizontal(classes="buttons"):
                yield Button("Close", variant="primary", id="close")

    def on_mount(self):
        self.query_one("#helpbody").focus()

    @on(Button.Pressed)
    def action_close(self):
        self.dismiss(None)


class NameEditDialog(ModalScreen):
    """Create or edit a defined name. on_submit(name, refers_to, scope) -> error text or None."""

    DEFAULT_CSS = DIALOG_CSS + """
    NameEditDialog Select { width: 60; }
    NameEditDialog .field { margin-top: 1; }
    """
    BINDINGS = [Binding("escape", "cancel", show=False)]

    def __init__(self, on_submit, sheets, name: str = "", refers_to: str = "", scope=None, editing: bool = False):
        super().__init__()
        self._on_submit = on_submit
        self._sheets = sheets
        self._name, self._refers, self._scope = name, refers_to, scope
        self._editing = editing

    def compose(self) -> ComposeResult:
        from textual.widgets import Select

        with Vertical(id="dialog"):
            yield Label("Edit name" if self._editing else "New name", classes="title")
            yield Label("Name:")
            yield Input(value=self._name, placeholder="e.g. TaxRate or Sales_2026", id="name")
            yield Label("Refers to:", classes="field")
            yield Input(value=self._refers, placeholder="e.g. Sheet1!$A$1:$B$10, =0.25 or =SUM(Data!B:B)",
                        id="refers")
            yield Label("Scope:", classes="field")
            options = [("Workbook", -1)] + [(s.name, i) for i, s in enumerate(self._sheets)]
            current = -1 if self._scope is None else self._sheets.index(self._scope)
            yield Select(options, value=current, allow_blank=False, id="scope", disabled=self._editing)
            yield Static("References without a sheet name refer to the current sheet.", classes="msg")
            yield Static("", classes="error", id="error")
            with Horizontal(classes="buttons"):
                yield Button("Cancel", id="cancel")
                yield Button("OK", variant="primary", id="ok")

    def on_mount(self):
        target = self.query_one("#refers" if self._editing else "#name", Input)
        target.focus()
        target.cursor_position = len(target.value)

    @on(Input.Submitted)
    @on(Button.Pressed, "#ok")
    def submit(self):
        from textual.widgets import Select

        name = self.query_one("#name", Input).value.strip()
        refers = self.query_one("#refers", Input).value
        idx = self.query_one("#scope", Select).value
        scope = None if idx == -1 else self._sheets[idx]
        err = self._on_submit(name, refers, scope)
        if err:
            self.query_one("#error", Static).update(err)
            return
        self.dismiss(name)

    @on(Button.Pressed, "#cancel")
    def action_cancel(self):
        self.dismiss(None)


class TableDialog(ModalScreen):
    """Create a table, or edit one (name, range, totals row; convert or delete it).

    on_submit(name, range_text, headers, totals) -> error text or None.
    Dismisses with "ok", "convert", "delete" or None."""

    DEFAULT_CSS = DIALOG_CSS + """
    TableDialog .field { margin-top: 1; }
    """
    BINDINGS = [Binding("escape", "cancel", show=False)]

    def __init__(self, on_submit, name: str, ref: str, editing: bool = False, totals: bool = False,
                 headers: bool = True):
        super().__init__()
        self._on_submit = on_submit
        self._name, self._ref = name, ref
        self._editing, self._totals, self._headers = editing, totals, headers

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label(f"Table {self._name}" if self._editing else "Create table", classes="title")
            yield Label("Table name:")
            yield Input(value=self._name, placeholder="e.g. Sales", id="name")
            yield Label("Range (including the header row):" if self._editing or self._headers else "Range:",
                        classes="field")
            yield Input(value=self._ref, placeholder="e.g. A1:D20", id="range")
            with Horizontal(classes="opts"):
                if self._editing:
                    yield Checkbox("Totals row", value=self._totals, id="totals")
                else:
                    yield Checkbox("My table has headers", value=self._headers, id="headers")
            if self._editing:
                yield Static("Formulas can use the table as e.g. =SUM(" + self._name + "[Column]), or [@Column] "
                             "for the same row inside it.", classes="msg", markup=False)
            else:
                yield Static("Without headers, the row above the range gets Column1, Column2, …", classes="msg")
            yield Static("", classes="error", id="error", markup=False)
            with Horizontal(classes="buttons"):
                if self._editing:
                    yield Button("Delete table", id="delete")
                    yield Button("Convert to range", id="convert")
                yield Button("Cancel", id="cancel")
                yield Button("OK", variant="primary", id="ok")

    def on_mount(self):
        target = self.query_one("#name" if self._editing else "#range", Input)
        target.focus()
        target.cursor_position = len(target.value)

    @on(Input.Submitted)
    @on(Button.Pressed, "#ok")
    def submit(self):
        name = self.query_one("#name", Input).value.strip()
        ref = self.query_one("#range", Input).value.strip()
        if self._editing:
            headers, totals = True, self.query_one("#totals", Checkbox).value
        else:
            headers, totals = self.query_one("#headers", Checkbox).value, False
        err = self._on_submit(name, ref, headers, totals)
        if err:
            self.query_one("#error", Static).update(err)
            return
        self.dismiss("ok")

    @on(Button.Pressed, "#convert")
    def convert(self):
        self.dismiss("convert")

    @on(Button.Pressed, "#delete")
    def delete(self):
        self.dismiss("delete")

    @on(Button.Pressed, "#cancel")
    def action_cancel(self):
        self.dismiss(None)


class NameManagerDialog(ModalScreen):
    """List defined names with New / Edit / Delete / Go to."""

    DEFAULT_CSS = DIALOG_CSS + """
    NameManagerDialog #dialog { width: 90%; height: 80%; }
    NameManagerDialog DataTable { height: 1fr; }
    """
    BINDINGS = [
        Binding("escape", "close", "Close", show=False),
        Binding("n", "new", "New", show=False),
        Binding("e", "edit", "Edit", show=False),
        Binding("delete,d", "delete", "Delete", show=False),
        Binding("g", "goto", "Go to", show=False),
    ]

    def __init__(self, app_ref):
        super().__init__()
        self._app = app_ref
        self._items = []

    def compose(self) -> ComposeResult:
        from textual.widgets import DataTable

        with Vertical(id="dialog"):
            yield Label("Name Manager", classes="title")
            yield DataTable(cursor_type="row", zebra_stripes=True, id="table")
            yield Static("", classes="msg", id="msg")
            with Horizontal(classes="buttons"):
                yield Button("New… (n)", id="new")
                yield Button("Edit… (e)", id="edit")
                yield Button("Delete (d)", id="delete")
                yield Button("Go to (g)", id="goto")
                yield Button("Close", variant="primary", id="close")

    def on_mount(self):
        from textual.widgets import DataTable

        table = self.query_one("#table", DataTable)
        table.add_columns("Name", "Refers to", "Value", "Scope")
        self.reload()
        table.focus()

    def reload(self, select: str | None = None):
        from textual.widgets import DataTable

        table = self.query_one("#table", DataTable)
        table.clear()
        self._items = self._app.name_list()
        for i, (info, value) in enumerate(self._items):
            table.add_row(info.name, info.refers_to, value, info.scope_label, key=str(i))
        msg = self.query_one("#msg", Static)
        msg.update("No names yet. Press n to define one for the current selection."
                   if not self._items else f"{len(self._items)} name{'s' if len(self._items) != 1 else ''}")
        if select is not None:
            for i, (info, _) in enumerate(self._items):
                if info.name.lower() == select.lower():
                    table.move_cursor(row=i)
                    break

    def _current(self):
        from textual.widgets import DataTable

        table = self.query_one("#table", DataTable)
        if not self._items or table.cursor_row is None or table.cursor_row >= len(self._items):
            return None
        return self._items[table.cursor_row][0]

    def on_data_table_row_selected(self, event):
        self.action_edit()

    @on(Button.Pressed, "#new")
    def action_new(self):
        self._app.open_name_editor(None, done=self.reload)

    @on(Button.Pressed, "#edit")
    def action_edit(self):
        info = self._current()
        if info is not None:
            self._app.open_name_editor(info, done=self.reload)

    @on(Button.Pressed, "#delete")
    def action_delete(self):
        info = self._current()
        if info is None:
            return

        def done(choice):
            if choice == "delete":
                self._app.remove_name(info)
                self.reload()

        self.app.push_screen(ChoiceDialog("Delete name", f"Delete the name '{info.name}'? Formulas that use it "
                                                         "will show #NAME?. (Ctrl+Z undoes this.)",
                                          [("delete", "&Delete"), ("cancel", "Cancel")], default="cancel"), done)

    @on(Button.Pressed, "#goto")
    def action_goto(self):
        info = self._current()
        if info is not None and self._app.goto_name(info):
            self.dismiss(None)

    @on(Button.Pressed, "#close")
    def action_close(self):
        self.dismiss(None)
