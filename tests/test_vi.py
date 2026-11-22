"""Vi-style keys (TermXLApp(vi=True))."""
import openpyxl
import pytest

from termxl import settings
from termxl.model.refs import parse_addr
from termxl.ui.app import TermXLApp
from tests.test_ui import _ready, type_text, val


def put(app, a1, text):
    app.book.set_input(app.grid.sheet, *parse_addr(a1), text)
    app.after_change()


def at(app):
    return app.grid.active


@pytest.mark.asyncio
async def test_move_edit_and_esc_confirms():
    app = TermXLApp(vi=True)
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        assert "NORMAL" in str(app.query_one("#status").render())
        await pilot.press("i")
        await type_text(pilot, "10")
        await pilot.press("escape")
        assert app.editing is None and val(app, "A1") == 10 and at(app) == (1, 1)
        await pilot.press("j", "equals_sign")
        await type_text(pilot, "A1*2")
        await pilot.press("escape")
        assert val(app, "A2") == 20 and at(app) == (2, 1)
        # letters don't start editing in normal mode
        await pilot.press("q", "t", "r")
        assert app.editing is None and val(app, "A3") is None
        await pilot.press("5", "j")
        assert at(app) == (7, 1)
        await pilot.press("g", "g")
        assert at(app) == (1, 1)
        await pilot.press("G")
        assert at(app) == (2, 1)
        await pilot.press("1", "2", "G", "3", "vertical_line")
        assert at(app) == (12, 3)
        await pilot.press("0", "l", "l", "k", "h")
        assert at(app) == (11, 2)
        put(app, "D11", "x")
        await pilot.press("dollar_sign")
        assert at(app) == (11, 4)
        await pilot.press("u")
        assert val(app, "D11") is None
        # Enter still moves down; arrows take counts
        await pilot.press("enter", "2", "up")
        assert at(app) == (10, 4)


@pytest.mark.asyncio
async def test_esc_after_formula_error_then_discard():
    app = TermXLApp(vi=True)
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        await pilot.press("equals_sign")
        await type_text(pilot, "1+*")
        await pilot.press("escape")
        assert app.editing is not None
        await pilot.press("escape")
        assert app.editing is None and val(app, "A1") is None


@pytest.mark.asyncio
async def test_yank_delete_put_cells():
    app = TermXLApp(vi=True)
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        put(app, "A1", "1")
        put(app, "B1", "=A1+1")
        await pilot.press("v", "l", "y")
        assert app.vi.visual is None and at(app) == (1, 1)
        await pilot.press("3", "j", "p")
        assert val(app, "A4") == 1 and app.grid.sheet.input_text(4, 2) == "=A4+1"
        await pilot.press("d", "d")
        assert val(app, "A4") is None and val(app, "B4") == 1
        await pilot.press("x")  # nothing left in A4; x on B4 clears it
        await pilot.press("l", "x")
        assert val(app, "B4") is None
        await pilot.press("u", "u")
        assert val(app, "A4") == 1 and val(app, "B4") == 2
        await pilot.press("ctrl+r")
        assert val(app, "A4") is None


@pytest.mark.asyncio
async def test_rows_delete_put_and_open():
    app = TermXLApp(vi=True)
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        for i, v in enumerate(["a", "b", "c", "d"], 1):
            put(app, f"A{i}", v)
        put(app, "B4", "=A1")
        await pilot.press("d", "r")  # delete row 1: rows are now b c d
        assert [val(app, f"A{i}") for i in range(1, 4)] == ["b", "c", "d"]
        assert app.grid.sheet.input_text(3, 2) == "=#REF!"
        await pilot.press("p")  # put row a below the cursor
        assert [val(app, f"A{i}") for i in range(1, 5)] == ["b", "a", "c", "d"]
        assert at(app) == (2, 1)
        await pilot.press("u")  # one undo step for insert + paste
        assert [val(app, f"A{i}") for i in range(1, 4)] == ["b", "c", "d"]
        assert at(app) == (2, 1) and app.grid.selection.rows == 1
        await pilot.press("k")
        # V selects whole rows; 2 rows yanked then put above row 1
        await pilot.press("V", "j", "y", "P")
        assert [val(app, f"A{i}") for i in range(1, 6)] == ["b", "c", "b", "c", "d"]
        await pilot.press("V", "j", "d")
        assert [val(app, f"A{i}") for i in range(1, 4)] == ["b", "c", "d"]
        # o inserts a row below and starts typing in it
        await pilot.press("o")
        assert app.editing is not None
        await type_text(pilot, "new")
        await pilot.press("escape")
        assert [val(app, f"A{i}") for i in range(1, 5)] == ["b", "new", "c", "d"]


@pytest.mark.asyncio
async def test_command_line_and_search(tmp_path):
    path = str(tmp_path / "v.xlsx")
    app = TermXLApp(path, vi=True)
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        put(app, "C7", "apple")
        put(app, "A9", "Pineapple")
        await pilot.press("colon")
        await type_text(pilot, "B5")
        await pilot.press("enter")
        assert at(app) == (5, 2)
        await pilot.press("slash")
        await type_text(pilot, "APPLE")
        await pilot.press("enter")
        assert at(app) == (7, 3)
        await pilot.press("n")
        assert at(app) == (9, 1)
        await pilot.press("N")
        assert at(app) == (7, 3)
        await pilot.press("colon")
        await type_text(pilot, "w")
        await pilot.press("enter")
        await pilot.pause(0.05)
        assert not app.book.modified
        await pilot.press("colon")
        await type_text(pilot, "set novi")
        await pilot.press("enter")
        assert app.vi is None
        await type_text(pilot, "x")  # Excel keys again: typing edits
        assert app.editing is not None
    assert openpyxl.load_workbook(path).active["C7"].value == "apple"


@pytest.mark.asyncio
async def test_preference_is_remembered():
    app = TermXLApp()
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        assert app.vi is None
        app.set_vi(True)
    assert settings.load()["keys"] == "vi"
    assert TermXLApp().vi is not None
    assert TermXLApp(vi=False).vi is None
