"""Scripted UI sessions using Textual's Pilot."""
import openpyxl
import pytest

from termxl.model.refs import parse_addr
from termxl.ui.app import TermXLApp
from termxl.ui.dialogs import ChoiceDialog, InputDialog


async def _ready(app, pilot):
    for _ in range(100):
        await pilot.pause(0.02)
        if app.book is not None and (app.path_arg is None or app.book.path):
            if not app.message.startswith("Opening"):
                break
    await pilot.pause(0.05)


def val(app, a1, sheet=None):
    return app.book.get_value(sheet or app.grid.sheet, *parse_addr(a1))


def screen_text(app) -> str:
    g = app.grid
    return "\n".join(g.render_line(y).text for y in range(g.size.height))


async def type_text(pilot, text):
    keys = []
    for ch in text:
        keys.append({" ": "space", "=": "equals_sign", "(": "left_parenthesis", ")": "right_parenthesis",
                     ":": "colon", "+": "plus", "*": "asterisk", "-": "minus", ",": "comma", ".": "full_stop",
                     '"': "quotation_mark", "/": "slash", "$": "dollar_sign", "%": "percent_sign",
                     "&": "ampersand", "<": "less_than_sign", ">": "greater_than_sign",
                     "'": "apostrophe"}.get(ch, ch))
    await pilot.press(*keys)


@pytest.mark.asyncio
async def test_type_values_and_formulas(tmp_path):
    path = str(tmp_path / "new.xlsx")
    app = TermXLApp(path)
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        await type_text(pilot, "10")
        await pilot.press("enter")
        await type_text(pilot, "32")
        await pilot.press("enter")
        await type_text(pilot, "=A1+A2")
        await pilot.press("enter")
        assert val(app, "A3") == 42
        assert app.grid.active == (4, 1)
        # editing A1 updates A3
        await pilot.press("ctrl+home")
        await type_text(pilot, "20")
        await pilot.press("enter")
        assert val(app, "A3") == 52
        assert "52" in screen_text(app)
        # undo
        await pilot.press("ctrl+z")
        assert val(app, "A3") == 42
        await pilot.press("ctrl+y")
        assert val(app, "A3") == 52
        # save to the path given on the command line
        await pilot.press("ctrl+s")
        await pilot.pause(0.1)
    wb = openpyxl.load_workbook(path)
    assert wb.active["A3"].value == "=A1+A2"
    assert openpyxl.load_workbook(path, data_only=True).active["A3"].value == 52


@pytest.mark.asyncio
async def test_point_mode_builds_references():
    app = TermXLApp()
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        for v in ("1", "2", "3"):
            await type_text(pilot, v)
            await pilot.press("enter")
        # in A4: =SUM( then arrow up 3 times and shift+down twice
        await type_text(pilot, "=SUM(")
        await pilot.press("up", "up", "up", "shift+down", "shift+down")
        assert app.formula.value == "=SUM(A1:A3"
        assert app.point is not None
        await pilot.press("enter")  # auto-closes the parenthesis
        assert app.book.sheets[0].input_text(4, 1) == "=SUM(A1:A3)"
        assert val(app, "A4") == 6


@pytest.mark.asyncio
async def test_arrow_commits_in_enter_mode_and_f2_edits():
    app = TermXLApp()
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        await type_text(pilot, "hello")
        await pilot.press("right")
        assert val(app, "A1") == "hello"
        assert app.grid.active == (1, 2)
        await pilot.press("left", "f2")
        assert app.editing is not None and app.editing.mode == "edit"
        await pilot.press("left", "left")
        await type_text(pilot, "X")
        await pilot.press("enter")
        assert val(app, "A1") == "helXlo"
        # escape cancels
        await pilot.press("up")
        await type_text(pilot, "zzz")
        await pilot.press("escape")
        assert val(app, "A1") == "helXlo"


@pytest.mark.asyncio
async def test_copy_paste_and_fill():
    app = TermXLApp()
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        for v in ("1", "2", "3"):
            await type_text(pilot, v)
            await pilot.press("enter")
        await pilot.press("ctrl+home", "right")
        await type_text(pilot, "=A1*2")
        await pilot.press("enter")
        await pilot.press("up", "ctrl+c", "down", "shift+down", "ctrl+v")
        sheet = app.grid.sheet
        assert sheet.input_text(2, 2) == "=A2*2"
        assert val(app, "B3") == 6
        # fill right
        await pilot.press("ctrl+home", "shift+right", "shift+right", "ctrl+r")
        assert val(app, "C1") == 1


@pytest.mark.asyncio
async def test_sample_file_navigation_and_insert_row(sample_path):
    app = TermXLApp(sample_path)
    async with app.run_test(size=(110, 30)) as pilot:
        await _ready(app, pilot)
        text = screen_text(app)
        assert "Household Budget 2026" in text
        assert "$7,238.75" in text
        assert app.grid.sheet.name == "Budget"
        # Ctrl+Down jumps through the block in column A
        await pilot.press("ctrl+home")
        assert app.grid.active == (4, 2)  # first unfrozen cell (freeze at B4)
        await pilot.press("ctrl+down")
        assert app.grid.active == (9, 2)
        # insert a row above 6 and check totals keep working
        await pilot.press("up", "up", "up", "shift+space")
        await pilot.press("alt+equals_sign")
        await pilot.pause(0.05)
        s = app.grid.sheet
        assert s.input_text(10, 5) == "=SUM(E4:E9)"
        await pilot.press("ctrl+pagedown")
        assert app.grid.sheet.name == "Data"
        assert "Region" in screen_text(app)


@pytest.mark.asyncio
async def test_goto_and_find(sample_path):
    app = TermXLApp(sample_path)
    async with app.run_test(size=(110, 30)) as pilot:
        await _ready(app, pilot)
        await pilot.press("ctrl+g")
        await pilot.pause(0.05)
        assert isinstance(app.screen, InputDialog)
        await type_text(pilot, "Data!I3")
        await pilot.press("enter")
        await pilot.pause(0.05)
        assert app.grid.sheet.name == "Data" and app.grid.active == (3, 9)
        msg = app.find_next("gizmo", case=False, whole=False, formulas=False)
        assert msg == "Found in B6"
        n = app.replace_all("Gadget-X", "Gadget", case=False, whole=True, formulas=False)
        assert n == 2


@pytest.mark.asyncio
async def test_quit_prompts_when_modified():
    app = TermXLApp()
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        await type_text(pilot, "1")
        await pilot.press("enter", "ctrl+q")
        await pilot.pause(0.05)
        assert isinstance(app.screen, ChoiceDialog)
        await pilot.press("escape")
        await pilot.pause(0.05)
        assert app.is_running


@pytest.mark.asyncio
async def test_formula_error_keeps_editing():
    app = TermXLApp()
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        await type_text(pilot, "=1+*")
        await pilot.press("enter")
        assert app.editing is not None  # bad formula is not committed
        await pilot.press("escape")
        assert app.editing is None


@pytest.mark.asyncio
async def test_name_box_defines_and_jumps():
    from termxl.model import names

    app = TermXLApp()
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        for v in ("1", "2", "3"):
            await type_text(pilot, v)
            await pilot.press("enter")
        await pilot.press("ctrl+home", "shift+down", "shift+down")
        await pilot.click("#namebox")
        await type_text(pilot, "Stuff")
        await pilot.press("enter")
        assert [n.name for n in names.list_names(app.book)] == ["Stuff"]
        assert app.query_one("#namebox").value == "Stuff"  # selection matches the name
        await pilot.press("right")
        await type_text(pilot, "=SUM(Stuff)")
        await pilot.press("enter")
        assert val(app, "B1") == 6
        # jump back via the name box
        await pilot.click("#namebox")
        await type_text(pilot, "stuff")
        await pilot.press("enter")
        assert app.grid.selection.a1() == "A1:A3"
        assert app.grid.has_focus


@pytest.mark.asyncio
async def test_name_manager_new_edit_delete():
    from termxl.ui.dialogs import NameEditDialog, NameManagerDialog

    app = TermXLApp()
    async with app.run_test(size=(110, 32)) as pilot:
        await _ready(app, pilot)
        await type_text(pilot, "=A3*Rate")
        await pilot.press("enter")
        assert str(val(app, "A1")) == "#NAME?"
        await pilot.press("ctrl+f3")
        await pilot.pause(0.1)
        assert isinstance(app.screen, NameManagerDialog)
        await pilot.press("n")
        await pilot.pause(0.1)
        assert isinstance(app.screen, NameEditDialog)
        await type_text(pilot, "Rate")
        await pilot.press("tab")
        await pilot.press("ctrl+a", "ctrl+k")  # clear "Refers to"
        await type_text(pilot, "=0.25")
        await pilot.press("enter")
        await pilot.pause(0.1)
        assert isinstance(app.screen, NameManagerDialog)
        assert val(app, "A1") == 0
        app.book.set_input(app.grid.sheet, 3, 1, "8")
        assert val(app, "A1") == 2
        # edit: rename Rate -> TaxRate, formulas follow
        await pilot.press("e")
        await pilot.pause(0.1)
        await pilot.press("shift+tab")
        await pilot.press("ctrl+a", "ctrl+k")
        await type_text(pilot, "TaxRate")
        await pilot.press("enter")
        await pilot.pause(0.1)
        assert app.grid.sheet.input_text(1, 1) == "=A3*TaxRate"
        # delete -> #NAME?
        await pilot.press("d")
        await pilot.pause(0.1)
        await pilot.press("d")  # "Delete" in the confirmation
        await pilot.pause(0.1)
        assert str(val(app, "A1")) == "#NAME?"
        await pilot.press("escape")
        await pilot.pause(0.05)
        await pilot.press("ctrl+z")
        assert val(app, "A1") == 2
