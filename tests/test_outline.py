"""Grouped (collapsible) rows and columns."""
import openpyxl
import pytest
from openpyxl.worksheet.properties import Outline

from termxl.io import xlsx
from termxl.io.xlsx import new_workbook
from termxl.model import outline
from termxl.model.refs import col_index
from termxl.ui.app import TermXLApp

from .test_ui import _ready, screen_text


def hidden_rows(sheet, a, b):
    return [r for r in range(a, b + 1) if sheet.row_hidden(r)]


def test_group_collapse_expand_rows():
    book = new_workbook()
    s = book.sheets[0]
    outline.group(book, s, "row", 2, 4)
    assert [outline.level(s, "row", r) for r in range(1, 6)] == [0, 1, 1, 1, 0]
    g = outline.groups_at(s, "row", 3)[0]
    assert (g.level, g.start, g.end, g.summary) == (1, 2, 4, 5)
    assert outline.groups_at(s, "row", 5) == [g]  # the summary row carries the button

    outline.set_collapsed(book, s, g, True)
    assert hidden_rows(s, 1, 6) == [2, 3, 4]
    assert outline.is_collapsed(s, g) and outline.collapsed_flag(s, "row", 5)

    outline.set_collapsed(book, s, g, False)
    assert hidden_rows(s, 1, 6) == []
    assert not outline.collapsed_flag(s, "row", 5)

    book.do_undo()  # expand
    assert hidden_rows(s, 1, 6) == [2, 3, 4]
    book.do_undo()  # collapse
    book.do_undo()  # group
    assert outline.max_level(s, "row") == 0
    book.do_redo()
    assert outline.level(s, "row", 3) == 1


def test_nested_groups_remember_collapsed_state():
    book = new_workbook()
    s = book.sheets[0]
    outline.group(book, s, "row", 2, 10)
    outline.group(book, s, "row", 4, 6)
    inner = outline.group_at(s, "row", 5, 2)
    outer = outline.group_at(s, "row", 5, 1)
    assert (inner.summary, outer.summary) == (7, 11)

    outline.set_collapsed(book, s, inner, True)
    outline.set_collapsed(book, s, outer, True)
    assert hidden_rows(s, 1, 12) == list(range(2, 11))
    outline.set_collapsed(book, s, outer, False)
    assert hidden_rows(s, 1, 12) == [4, 5, 6]  # the inner group stays collapsed


def test_ungroup_shows_collapsed_members():
    book = new_workbook()
    s = book.sheets[0]
    outline.group(book, s, "row", 2, 4)
    outline.set_collapsed(book, s, outline.group_at(s, "row", 2, 1), True)
    assert outline.ungroup(book, s, "row", 2, 4)
    assert hidden_rows(s, 1, 6) == []
    assert outline.max_level(s, "row") == 0
    assert not outline.collapsed_flag(s, "row", 5)
    assert not outline.ungroup(book, s, "row", 2, 4)  # nothing left to ungroup


def test_show_level():
    book = new_workbook()
    s = book.sheets[0]
    outline.group(book, s, "row", 2, 10)
    outline.group(book, s, "row", 4, 6)
    outline.show_level(book, s, "row", 2)
    assert hidden_rows(s, 1, 12) == [4, 5, 6]
    outline.show_level(book, s, "row", 1)
    assert hidden_rows(s, 1, 12) == list(range(2, 11))
    outline.show_level(book, s, "row", 3)
    assert hidden_rows(s, 1, 12) == []


def test_group_columns_splits_wide_dimension():
    book = new_workbook()
    s = book.sheets[0]
    ws = s.ws
    ws.column_dimensions.group("B", "XFD", outline_level=0)
    ws.column_dimensions["B"].width = 20
    s.invalidate_layout()
    outline.group(book, s, "col", 5, 7)
    assert len(ws.column_dimensions) < 10  # B:XFD split into B:D, E, F, G, H:XFD
    assert [outline.level(s, "col", c) for c in range(4, 9)] == [0, 1, 1, 1, 0]
    assert s.col_width(6)[0] == 20 and s.col_width(9)[0] == 20


def test_summary_left_and_above():
    book = new_workbook()
    s = book.sheets[0]
    s.ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False)
    outline.group(book, s, "col", 3, 5)
    g = outline.group_at(s, "col", 4, 1)
    assert g.summary == 2
    assert outline.groups_at(s, "col", 2) == [g]


def test_save_roundtrip(tmp_path):
    book = new_workbook()
    s = book.sheets[0]
    outline.group(book, s, "row", 3, 5)
    outline.group(book, s, "col", 2, 3)
    outline.set_collapsed(book, s, outline.group_at(s, "row", 3, 1), True)
    path = str(tmp_path / "o.xlsx")
    xlsx.save(book, path)

    ws = openpyxl.load_workbook(path).active
    assert ws.sheet_format.outlineLevelRow == 1
    assert ws.sheet_format.outlineLevelCol == 1
    assert [ws.row_dimensions[r].outlineLevel for r in (3, 4, 5)] == [1, 1, 1]
    assert all(ws.row_dimensions[r].hidden for r in (3, 4, 5))
    assert ws.row_dimensions[6].collapsed

    again = xlsx.load(path).sheets[0]
    assert outline.max_level(again, "row") == 1 and outline.max_level(again, "col") == 1
    assert outline.is_collapsed(again, outline.group_at(again, "row", 4, 1))


def _exam_like(path):
    """Like the practice-exam workbooks: solution columns L:S grouped and collapsed."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["K1"] = "Solution >"
    ws["L1"] = "answer"
    ws.column_dimensions.group("L", "S", outline_level=1, hidden=True)
    wb.save(path)


@pytest.mark.asyncio
async def test_ui_click_buttons(tmp_path):
    path = str(tmp_path / "exam.xlsx")
    _exam_like(path)
    app = TermXLApp(path)
    async with app.run_test(size=(160, 12)) as pilot:
        await _ready(app, pilot)
        g = app.grid
        lines = screen_text(app).splitlines()
        assert lines[0].startswith("1 2")
        x = lines[0].index("[+]")
        assert "answer" not in screen_text(app)
        await pilot.click("#grid", offset=(x, 0))
        await pilot.pause(0.05)
        assert not g.sheet.col_width(col_index("L"))[1]
        assert "answer" in screen_text(app)
        assert "─" in screen_text(app).splitlines()[0]
        # level button "1" collapses everything again
        await pilot.click("#grid", offset=(0, 0))
        await pilot.pause(0.05)
        assert g.sheet.col_width(col_index("L"))[1]
        assert "answer" not in screen_text(app)


@pytest.mark.asyncio
async def test_ui_group_keys(tmp_path):
    app = TermXLApp(str(tmp_path / "new.xlsx"))
    async with app.run_test(size=(100, 20)) as pilot:
        await _ready(app, pilot)
        g = app.grid
        g.set_cursor(2, 1)
        g.set_cursor(4, 1, extend=True)
        await pilot.press("shift+space", "alt+shift+right")  # whole rows: no rows/columns question
        await pilot.pause(0.05)
        assert outline.max_level(g.sheet, "row") == 1
        assert screen_text(app).splitlines()[0].startswith("1 2")  # row level buttons
        g.set_cursor(3, 1)
        await pilot.press("alt+shift+up")
        await pilot.pause(0.05)
        assert hidden_rows(g.sheet, 1, 6) == [2, 3, 4]
        assert g.active == (5, 1)  # moved to the summary row
        await pilot.press("alt+shift+down")
        await pilot.pause(0.05)
        assert hidden_rows(g.sheet, 1, 6) == []
        g.set_cursor(2, 1)
        g.set_cursor(3, 2, extend=True)
        await pilot.press("alt+shift+left", "r")  # not whole rows/columns: asks, R picks rows
        await pilot.pause(0.05)
        assert [outline.level(g.sheet, "row", r) for r in range(1, 6)] == [0, 0, 0, 1, 0]
