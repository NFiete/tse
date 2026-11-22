from termse.formula.errors import CIRC, REF
from termse.model import structure
from termse.model.refs import Rect


def test_recalc_chain(ws):
    ws["A1"] = "1"
    for i in range(2, 2001):
        ws[f"A{i}"] = f"=A{i - 1}+1"
    assert ws["A2000"] == 2000
    ws["A1"] = "10"
    assert ws["A2000"] == 2009


def test_range_dependents(ws):
    ws["B1"] = "=SUM(A:A)"
    ws["C1"] = "=COUNT(A1:A100)"
    ws["A50"] = "5"
    assert ws["B1"] == 5 and ws["C1"] == 1
    ws["A51"] = "7"
    assert ws["B1"] == 12 and ws["C1"] == 2
    ws["A51"] = ""
    assert ws["B1"] == 5


def test_cross_sheet(ws):
    b = ws.book
    s2 = structure.add_sheet(b, "Other")
    b.set_input(s2, 1, 1, "21")
    ws["A1"] = "=Other!A1*2"
    assert ws["A1"] == 42
    b.set_input(s2, 1, 1, "50")
    assert ws["A1"] == 100
    structure.rename_sheet(b, s2, "My Data")
    assert ws.text("A1") == "='My Data'!A1*2"
    assert ws["A1"] == 100


def test_circular(ws):
    ws["A1"] = "=B1"
    ws["B1"] = "=A1"
    assert CIRC in (ws["A1"], ws["B1"])


def test_undo_redo(ws):
    b = ws.book
    ws["A1"] = "5"
    ws["A2"] = "=A1*2"
    ws["A1"] = "7"
    assert ws["A2"] == 14
    b.do_undo()
    assert ws["A1"] == 5 and ws["A2"] == 10
    b.do_undo()
    assert ws["A2"] is None
    b.do_redo()
    b.do_redo()
    assert ws["A2"] == 14


def test_format_inference(ws):
    ws["A1"] = "2026-01-15"
    assert ws.fmt("A1") == "yyyy-mm-dd"
    ws["A2"] = "=A1+7"
    assert ws.fmt("A2") == "yyyy-mm-dd"
    ws["A3"] = "=TODAY()"
    assert ws.fmt("A3") == "m/d/yyyy"
    ws["B1"] = "15%"
    assert ws["B1"] == 0.15 and ws.fmt("B1") == "0%"
    ws["B1"] = "20"  # keeps the percent format, like Excel
    assert ws.fmt("B1") == "0%"


def test_input_text_roundtrip(ws):
    ws["A1"] = "'=not a formula"
    assert ws["A1"] == "=not a formula"
    assert ws.text("A1") == "'=not a formula"
    ws["A2"] = "'123"
    assert ws["A2"] == "123"
    assert ws.text("A2") == "'123"
    ws["A3"] = "=sum(1,2)"
    assert ws.text("A3") == "=SUM(1,2)"
    ws["A4"] = "=xlookup(1,{1},{2})"
    assert ws["A4"] == 2
    assert ws.sheet.ws["A4"].value == "=_xlfn.XLOOKUP(1,{1},{2})"


def test_copy_paste_translates(ws):
    b = ws.book
    ws["A1"] = "1"
    ws["A2"] = "2"
    ws["B1"] = "=A1*10"
    clip = b.capture(ws.sheet, Rect(1, 2, 1, 2))
    b.paste(clip, ws.sheet, 2, 2)
    assert ws.text("B2") == "=A2*10" and ws["B2"] == 20


def test_cut_paste_moves_references(ws):
    b = ws.book
    ws["A1"] = "3"
    ws["C1"] = "=A1+1"
    clip = b.capture(ws.sheet, Rect(1, 1, 1, 1))
    b.paste(clip, ws.sheet, 5, 1, cut=True)
    assert ws["A1"] is None
    assert ws.text("C1") == "=A5+1" and ws["C1"] == 4


def test_paste_values_onto_overlapping_range(ws):
    ws["A1"] = "=1"
    ws["A2"] = "=2"
    ws["A3"] = "=3"
    book, sheet = ws.book, ws.sheet
    book.paste(book.capture(sheet, Rect(1, 1, 3, 1)), sheet, 2, 1, values_only=True)
    assert [ws[f"A{r}"] for r in range(1, 5)] == [1, 1, 2, 3]
    assert ws.text("A3") == "2"


def test_fill_down(ws):
    b = ws.book
    for i in range(1, 5):
        ws[f"A{i}"] = str(i)
    ws["B1"] = "=A1*A1"
    b.fill(ws.sheet, Rect(1, 2, 4, 2), "down")
    assert [ws[f"B{i}"] for i in range(1, 5)] == [1, 4, 9, 16]


def test_paste_text(ws):
    b = ws.book
    b.paste_text(ws.sheet, 1, 1, "a\tb\n1\t2.5\n")
    assert ws["A1"] == "a" and ws["B2"] == 2.5


def test_insert_delete_rows(ws):
    b = ws.book
    for i in range(1, 6):
        ws[f"A{i}"] = str(i)
    ws["B1"] = "=SUM(A1:A5)"
    ws["B2"] = "=A5"
    structure.insert(b, ws.sheet, "row", 3, 2)
    assert ws.text("B1") == "=SUM(A1:A7)"
    assert ws.text("B2") == "=A7"
    assert ws["A3"] is None and ws["A5"] == 3
    assert ws["B1"] == 15
    structure.delete(b, ws.sheet, "row", 3, 2)
    assert ws.text("B1") == "=SUM(A1:A5)" and ws["B1"] == 15
    structure.delete(b, ws.sheet, "row", 5, 1)
    assert ws.text("B2") == "=#REF!"
    assert ws.text("B1") == "=SUM(A1:A4)" and ws["B1"] == 10
    b.do_undo()
    assert ws.text("B2") == "=A5" and ws["B2"] == 5
    assert ws["B1"] == 15


def test_insert_delete_columns(ws):
    b = ws.book
    ws["A1"] = "1"
    ws["B1"] = "2"
    ws["C1"] = "=A1+B1"
    structure.insert(b, ws.sheet, "col", 2, 1)
    assert ws.text("D1") == "=A1+C1" and ws["D1"] == 3
    structure.delete(b, ws.sheet, "col", 1, 1)
    assert ws.text("C1") == "=#REF!+B1"
    assert ws["C1"] == REF


def test_merged_cells_shift(ws):
    b = ws.book
    ws["A2"] = "title"
    ws.sheet.ws.merge_cells("A2:C3")
    structure.insert(b, ws.sheet, "row", 1, 1)
    assert [m.coord for m in ws.sheet.ws.merged_cells.ranges] == ["A3:C4"]
    structure.delete(b, ws.sheet, "row", 1, 1)
    assert [m.coord for m in ws.sheet.ws.merged_cells.ranges] == ["A2:C3"]


def test_defined_names(ws):
    from openpyxl.workbook.defined_name import DefinedName

    b = ws.book
    ws["A1"] = "4"
    ws["A2"] = "6"
    b.wb.defined_names["Stuff"] = DefinedName("Stuff", attr_text="Sheet!$A$1:$A$2")
    b._load_names()
    ws["B1"] = "=SUM(Stuff)"
    assert ws["B1"] == 10
    ws["A1"] = "5"
    assert ws["B1"] == 11


def test_bold_toggle_and_undo(ws):
    b = ws.book
    ws["A1"] = "x"
    b.toggle_font(ws.sheet, Rect(1, 1, 1, 1), "bold")
    assert ws.sheet.ws["A1"].font.b
    b.do_undo()
    assert not ws.sheet.ws["A1"].font.b


def test_structural_edits_recalc_what_changes(ws):
    b = ws.book
    other = structure.add_sheet(b, "Other")
    for i in range(1, 11):
        ws[f"A{i}"] = str(i * 10)
    ws["B1"] = "=INDEX(A1:A10,5)"      # range spans the insert -> value changes
    ws["B2"] = "=ROWS(A1:A10)"
    ws["C8"] = "=ROW()"                # moves -> value changes
    ws["C9"] = "=A9*2"                 # moves, value doesn't change
    b.set_input(other, 1, 1, "=Sheet!A9+Sheet!A2")
    structure.insert(b, ws.sheet, "row", 3, 2)
    assert ws["B1"] == 30 and ws.text("B1") == "=INDEX(A1:A12,5)"
    assert ws["B2"] == 12
    assert ws["C10"] == 10
    assert ws["C11"] == 180 and ws.text("C11") == "=A11*2"
    assert b.sheets[1].input_text(1, 1) == "=Sheet!A11+Sheet!A2"
    assert b.get_value(other, 1, 1) == 110
    # deleting a referenced row makes #REF!, and undo restores everything
    structure.delete(b, ws.sheet, "row", 11, 1)
    assert b.sheets[1].input_text(1, 1) == "=Sheet!#REF!+Sheet!A2"
    assert ws["B2"] == 11
    b.do_undo()
    assert b.sheets[1].input_text(1, 1) == "=Sheet!A11+Sheet!A2"
    assert b.get_value(other, 1, 1) == 110
    assert ws["C11"] == 180 and ws["B2"] == 12
    b.do_undo()  # undo the insert
    assert ws.text("B1") == "=INDEX(A1:A10,5)" and ws["B1"] == 50
    assert ws["C8"] == 8
