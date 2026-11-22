"""Spilling: array results fill the cells below/right of the formula (dynamic arrays)."""
import zipfile

import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

from termxl.formula.errors import ExcelError
from termxl.formula.parser import parse, to_display, to_storage
from termxl.io import xlsx
from termxl.model import structure
from termxl.model.refs import Rect, parse_addr


def col(ws, a1, n):
    """n values going down from a1."""
    r, c = parse_addr(a1)
    out = []
    for i in range(n):
        v = ws.book.get_value(ws.sheet, r + i, c)
        out.append(v.code if isinstance(v, ExcelError) else v)
    return out


def block(ws, a1, nr, nc):
    r, c = parse_addr(a1)
    out = []
    for i in range(nr):
        row = []
        for j in range(nc):
            v = ws.book.get_value(ws.sheet, r + i, c + j)
            row.append(v.code if isinstance(v, ExcelError) else v)
        out.append(row)
    return out


@pytest.fixture
def nums(ws):
    for i in range(1, 6):
        ws[f"A{i}"] = str(i * 10)
    return ws


def test_spill_basic(nums):
    ws = nums
    ws["C1"] = "=A1:A5*2"
    assert col(ws, "C1", 6) == [20, 40, 60, 80, 100, None]
    assert ws.text("C2") == ""  # spilled cells hold no formula
    assert ws.sheet.spill_anchor(3, 3) == (1, 3)
    assert ws.sheet.max_row == 5
    ws["E1"] = "=SEQUENCE(2,3)"
    assert block(ws, "E1", 2, 3) == [[1, 2, 3], [4, 5, 6]]
    ws["E1"] = "=SEQUENCE(1,2)"  # shrinking clears the rest
    assert block(ws, "E1", 2, 3) == [[1, 2, None], [None, None, None]]
    ws["E1"] = "5"
    assert block(ws, "E1", 1, 2) == [[5, None]]


def test_one_by_one_and_empty_cells(ws):
    ws["A1"] = "=SEQUENCE(1)"
    assert ws["A1"] == 1 and not ws.sheet.spill.anchors
    ws["B2"] = "x"
    ws["D1"] = "=A1:B2"  # blanks in a spilled range show as 0
    assert block(ws, "D1", 2, 2) == [[1, 0], [0, "x"]]


def test_blocked_spill_and_recovery(nums):
    ws = nums
    ws["C1"] = "=A1:A5"
    ws["C3"] = "in the way"
    assert ws["C1"].code == "#SPILL!"
    assert ws["C2"] is None
    assert ws.sheet.spill.blocked == {(1, 3): (5, 1)}
    ws["C3"] = ""
    assert col(ws, "C1", 5) == [10, 20, 30, 40, 50]
    ws["C4"] = "x"  # typing into a spilled cell blocks it
    assert col(ws, "C1", 5) == ["#SPILL!", None, None, "x", None]
    ws.book.do_undo()
    assert col(ws, "C1", 5) == [10, 20, 30, 40, 50]


def test_spills_block_each_other(ws):
    ws["B1"] = "=SEQUENCE(3)"  # B1:B3
    ws["A2"] = "=SEQUENCE(1,3)"  # A2:C2 crosses B2
    assert ws["A2"].code == "#SPILL!"
    ws["B1"] = "=SEQUENCE(1)"
    assert block(ws, "A2", 1, 3) == [[1, 2, 3]]


def test_merged_cells_and_sheet_edge_block(ws):
    ws.sheet.ws.merge_cells("A3:B3")
    ws["A1"] = "=SEQUENCE(3)"
    assert ws["A1"].code == "#SPILL!"
    ws["Z1048575"] = "=SEQUENCE(3)"
    assert ws["Z1048575"].code == "#SPILL!"


def test_dependents_follow_the_spill(nums):
    ws = nums
    ws["C1"] = "=A1:A5+1"
    ws["E1"] = "=C3*2"
    ws["E2"] = "=SUM(C1:C10)"
    ws["E3"] = "=COUNT(C:C)"
    assert (ws["E1"], ws["E2"], ws["E3"]) == (62, 155, 5)
    ws["A3"] = "100"
    assert (ws["E1"], ws["E2"], ws["E3"]) == (202, 225, 5)
    ws["C1"] = "=A1:A2+1"  # shrinks: C3 is now empty
    assert (ws["E1"], ws["E2"], ws["E3"]) == (0, 32, 2)
    ws["C1"] = ""
    assert (ws["E1"], ws["E2"], ws["E3"]) == (0, 0, 0)
    ws["C1"] = "=SORT(A1:A5,,-1)"  # grows again
    assert (ws["E1"], ws["E2"], ws["E3"]) == (80, 220, 5)


def test_spill_reference(nums):
    ws = nums
    ws["C1"] = "=FILTER(A1:A5,A1:A5>20)"
    ws["D1"] = "=SUM(C1#)"
    ws["D2"] = "=ROWS(C1#)"
    ws["D3"] = "=C1#*10"
    assert (ws["D1"], ws["D2"]) == (120, 3)
    assert col(ws, "D3", 3) == [300, 400, 500]
    ws["A1"] = "99"
    assert (ws["D1"], ws["D2"]) == (219, 4)
    assert col(ws, "D3", 4) == [990, 300, 400, 500]
    ws["F1"] = "=A1#"  # not a formula
    assert ws["F1"].code == "#REF!"
    ws["F1"] = "=D1#"  # a formula that doesn't spill is a one-cell range
    assert ws["F1"] == 219
    ws["C2"] = "x"  # a blocked spill can't be referenced
    assert ws["D1"].code == "#REF!"


def test_self_reference_is_circular(ws):
    ws["A1"] = "=SEQUENCE(3)"
    ws["A1"] = "=SEQUENCE(3)+SUM(A2:A3)"
    assert ws["A1"].code == "#CIRC!"


def test_implicit_intersection_operator(nums):
    ws = nums
    ws["C3"] = "=@A1:A5"
    assert ws["C3"] == 30
    assert ws.text("C3") == "=@A1:A5"
    ws["C7"] = "=@A1:A5"
    assert ws["C7"].code == "#VALUE!"
    ws["D1"] = "=@SEQUENCE(3)"
    assert ws["D1"] == 1 and ws["D2"] is None


def test_storage_forms():
    cases = {
        "=A1#": "=_xlfn.ANCHORARRAY(A1)",
        "=SUM(Sheet2!B3#)": "=SUM(_xlfn.ANCHORARRAY(Sheet2!B3))",
        "=@A1:A10*2": "=_xlfn.SINGLE(A1:A10)*2",
        "=@INDEX(A1:B5,0,1)": "=_xlfn.SINGLE(INDEX(A1:B5,0,1))",
        "=@A1#": "=_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))",
        "=Table1[@Qty]": "=Table1[[#This Row],[Qty]]",
    }
    for display, storage in cases.items():
        assert to_storage(display) == storage
        assert to_display(storage) == display
    assert to_display("=_xlfn.SINGLE(A1:A3+1)") == "=@(A1:A3+1)"
    assert repr(parse("=A1#:B5")) == "(ANCHORARRAY(Ref(None!A1:A1)) : Ref(None!B5:B5))"


def test_insert_and_delete_rows(nums):
    ws = nums
    ws["C1"] = "=A1:A5"
    ws["D1"] = "=SUM(C1#)"
    structure.insert(ws.book, ws.sheet, "row", 1, 2)
    assert col(ws, "C1", 8) == [None, None, 10, 20, 30, 40, 50, None]
    assert ws["D3"] == 150
    structure.delete(ws.book, ws.sheet, "row", 4, 1)  # deletes A4 (20)
    assert col(ws, "C3", 5) == [10, 30, 40, 50, None]
    assert ws["D3"] == 130
    ws.book.do_undo()
    ws.book.do_undo()
    assert col(ws, "C1", 5) == [10, 20, 30, 40, 50]


def test_copy_paste_spilled_cells(nums):
    ws = nums
    book, sheet = ws.book, ws.sheet
    ws["C1"] = "=A1:A5*2"
    clip = book.capture(sheet, Rect(2, 3, 3, 3))  # C2:C3, without the formula: values
    book.paste(clip, sheet, 1, 6)
    assert col(ws, "F1", 3) == [40, 60, None]
    assert ws.text("F1") == "40"
    clip = book.capture(sheet, Rect(1, 3, 5, 3))  # the formula and its spill: it spills again
    book.paste(clip, sheet, 1, 8)
    assert ws.text("H1") == "=F1:F5*2"
    assert col(ws, "H1", 3) == [80, 120, 0]
    book.paste(clip, sheet, 1, 10, values_only=True)
    assert col(ws, "J1", 5) == [20, 40, 60, 80, 100]


def test_spilled_cells_use_the_formula_number_format(ws):
    ws["A1"] = "=SEQUENCE(3,1,DATE(2026,1,1))"
    ws.book.set_number_format(ws.sheet, Rect(1, 1, 1, 1), "m/d/yyyy")
    ws["A3"] = ""
    assert ws.fmt("A3") == "m/d/yyyy"
    ws.book.set_number_format(ws.sheet, Rect(3, 1, 3, 1), "0.0")  # its own format wins
    assert ws.fmt("A3") == "0.0"


def test_save_and_load(nums, tmp_path):
    ws = nums
    ws["C1"] = "=A1:A5*2"
    ws["D1"] = "=SUM(C1#)"
    ws["E1"] = '=FILTER(A1:A5,A1:A5>100,"none")'
    ws["F1"] = "=SEQUENCE(2)"
    ws["F2"] = "x"
    ws["G1"] = "=SUM(A1:A5)"
    p = str(tmp_path / "s.xlsx")
    xlsx.save(ws.book, p)
    assert ws.sheet.ws["C2"].value is None  # the saved values were only put there while saving

    saved = openpyxl.load_workbook(p).active
    assert isinstance(saved["C1"].value, ArrayFormula) and saved["C1"].value.ref == "C1:C5"
    assert saved["C3"].value == 60
    assert saved["D1"].value == "=SUM(_xlfn.ANCHORARRAY(C1))"
    assert saved["E1"].value.ref == "E1"  # array evaluation even though the result is one value
    assert saved["F1"].value.ref == "F1"
    assert saved["G1"].value == "=SUM(A1:A5)"
    with zipfile.ZipFile(p) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode()
        assert '<c r="C1" cm="1">' in xml and '<c r="G1">' in xml
        assert "XLDAPR" in z.read("xl/metadata.xml").decode()
        assert "sheetMetadata" in z.read("xl/_rels/workbook.xml.rels").decode()
        assert "/xl/metadata.xml" in z.read("[Content_Types].xml").decode()

    book = xlsx.load(p)
    s = book.sheets[0]
    assert s.formulas[(1, 3)].text == "=A1:A5*2" and not s.formulas[(1, 3)].legacy
    assert s.formulas[(1, 4)].text == "=SUM(C1#)"
    assert s.ws["C3"].value is None
    assert [book.get_value(s, r, 3) for r in range(1, 6)] == [20, 40, 60, 80, 100]
    assert book.get_value(s, 1, 6).code == "#SPILL!"
    book.set_input(s, 2, 6, "")  # the blocked one spills once its range is clear
    assert [book.get_value(s, r, 6) for r in range(1, 3)] == [1, 2]
    book.set_input(s, 1, 1, "1")
    assert [book.get_value(s, r, 3) for r in range(1, 3)] == [2, 40]
    assert book.get_value(s, 1, 4) == 282


def test_load_keeps_saved_spill_of_unsupported_formula(tmp_path):
    """A spilled formula we can't calculate keeps the whole result Excel saved."""
    ws_book = xlsx.new_workbook()
    sheet = ws_book.sheets[0]
    ws_book.set_input(sheet, 1, 1, "=SEQUENCE(3)")
    p = str(tmp_path / "u.xlsx")
    xlsx.save(ws_book, p)
    q = str(tmp_path / "u2.xlsx")
    with zipfile.ZipFile(p) as zin, zipfile.ZipFile(q, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"_xlfn.SEQUENCE(3)", b"_xlfn.NOTAFUNCTION(3)")
            zout.writestr(item, data)
    book = xlsx.load(q)
    s = book.sheets[0]
    assert [book.get_value(s, r, 1) for r in range(1, 4)] == [1, 2, 3]
    assert (1, 1) in s.stale
    book.set_input(s, 3, 2, "5")  # unrelated edit: the saved spill stays
    assert [book.get_value(s, r, 1) for r in range(1, 4)] == [1, 2, 3]


def test_legacy_formulas_use_implicit_intersection(tmp_path):
    """Ordinary formulas in a file (not dynamic arrays) don't spill: Excel shows them as =@..."""
    wb = openpyxl.Workbook()
    w = wb.active
    for i in range(1, 6):
        w.cell(i, 1, i * 10)
    w["B3"] = "=A1:A5*2"
    w["C1"] = "=A1:A5"
    w["D8"] = "=A1:A5"
    w["E3"] = "=SUMPRODUCT((A1:A5>20)*A1:A5)+A1:A5"  # array arguments stay arrays
    p = str(tmp_path / "legacy.xlsx")
    wb.save(p)
    book = xlsx.load(p)
    s = book.sheets[0]
    assert book.get_value(s, 3, 2) == 60
    assert book.get_value(s, 1, 3) == 10
    assert book.get_value(s, 8, 4).code == "#VALUE!"
    book.recalc_all()
    assert book.get_value(s, 3, 5) == 150
    assert not s.spill.anchors
    book.set_input(s, 1, 1, "1")  # recalculated, still no spill
    assert book.get_value(s, 1, 3) == 1 and book.get_value(s, 2, 3) is None
    book.set_input(s, 1, 3, "=A1:A5")  # retyped: now it spills
    assert book.get_value(s, 2, 3) == 20
    xlsx.save(book, p)
    saved = openpyxl.load_workbook(p).active
    assert saved["B3"].value == "=A1:A5*2"
    assert saved["C1"].value.ref == "C1:C5"


def test_legacy_array_formula(tmp_path):
    """Ctrl+Shift+Enter array formulas fill their fixed range and can't be changed in part."""
    wb = openpyxl.Workbook()
    w = wb.active
    for i in range(1, 4):
        w.cell(i, 1, i)
    w["C1"] = ArrayFormula("C1:C4", "=A1:A3*10")
    w["D1"] = ArrayFormula("D1:E2", "=A1:A2")
    p = str(tmp_path / "cse.xlsx")
    wb.save(p)
    book = xlsx.load(p)
    s = book.sheets[0]
    vals = [book.get_value(s, r, 3) for r in range(1, 5)]
    assert vals[:3] == [10, 20, 30] and vals[3].code == "#N/A"
    assert [[book.get_value(s, r, c) for c in (4, 5)] for r in (1, 2)] == [[1, 1], [2, 2]]
    with pytest.raises(ValueError):
        book.set_input(s, 2, 3, "5")
    book.set_input(s, 1, 1, "7")
    assert book.get_value(s, 1, 3) == 70
    xlsx.save(book, p)
    saved = openpyxl.load_workbook(p).active
    assert saved["C1"].value.ref == "C1:C4"
    assert saved["C2"].value == 20
    with zipfile.ZipFile(p) as z:
        assert "cm=" not in z.read("xl/worksheets/sheet1.xml").decode()


@pytest.mark.asyncio
async def test_ui_spilled_range():
    from termxl.ui.app import TermXLApp
    from tests.test_ui import _ready, screen_text, type_text

    app = TermXLApp()
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        await type_text(pilot, "=SEQUENCE(4)*7")
        await pilot.press("enter")
        await pilot.pause()
        g = app.grid
        assert g.active == (2, 1)
        assert g._active_spill() == (1, 1, 4, 1)
        assert "28" in screen_text(app)
        status = str(app.query_one("#status").render())
        assert "Spilled from A1: =SEQUENCE(4)*7" in status
        await type_text(pilot, "x")
        await pilot.press("enter")
        await pilot.press("ctrl+home")
        await pilot.pause()
        assert "#SPILL!" in screen_text(app)
        assert "Can't spill: A1:A4 isn't empty" in str(app.query_one("#status").render())
