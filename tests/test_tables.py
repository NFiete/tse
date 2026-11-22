"""Excel tables: structured references, creating/editing tables, and keeping them in step with edits."""
import math
import os
import shutil
import subprocess

import openpyxl
import pytest
from openpyxl.worksheet.table import Table, TableColumn, TableFormula

from termxl.formula.errors import ExcelError
from termxl.formula.parser import FormulaSyntaxError, StructRef, parse, sref_text, to_display, to_storage
from termxl.io import xlsx
from termxl.model import names, structure, tables
from termxl.model.refs import Rect, parse_addr

from .conftest import Sheetish

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def v(x):
    return x.code if isinstance(x, ExcelError) else x


def _table(book, name="Table1"):
    return tables.find_table(book, name)


def _cols(book, name="Table1"):
    return _table(book, name).col_names()


@pytest.fixture
def sales(ws):
    """Item / Qty / Price in A1:C4 made into Table1."""
    data = [["Item", "Qty", "Price"], ["apple", 2, 3], ["pear", 4, 5], ["fig", 6, 7]]
    for r, row in enumerate(data, 1):
        for c, x in enumerate(row):
            ws[f"{'ABC'[c]}{r}"] = x
    tables.create(ws.book, ws.sheet, Rect(1, 1, 4, 3))
    return ws


# --------------------------------------------------------------------------- parsing


@pytest.mark.parametrize("text,expected", [
    ("Table1[Qty]", StructRef("Table1", frozenset(), "Qty")),
    ("Table1[]", StructRef("Table1")),
    ("Table1[#All]", StructRef("Table1", frozenset({"all"}))),
    ("[@Qty]", StructRef(None, frozenset({"thisrow"}), "Qty")),
    ("[@[Unit Price]]", StructRef(None, frozenset({"thisrow"}), "Unit Price")),
    ("Table1[@]", StructRef("Table1", frozenset({"thisrow"}))),
    ("Table1[[#This Row],[Qty]]", StructRef("Table1", frozenset({"thisrow"}), "Qty")),
    ("T[[#Headers],[#Data],[A]:[B]]", StructRef("T", frozenset({"headers", "data"}), "A", "B")),
    ("T[ [#Totals] , [Sales Amount] ]", StructRef("T", frozenset({"totals"}), "Sales Amount")),
    ("T[[Qty]:[Price]]", StructRef("T", frozenset(), "Qty", "Price")),
    ("T[Sales Amount]", StructRef("T", frozenset(), "Sales Amount")),
    ("T[a'[1']]", StructRef("T", frozenset(), "a[1]")),
    ("T[[it''s]]", StructRef("T", frozenset(), "it's")),
])
def test_parse_structured_refs(text, expected):
    node = parse("=" + text)
    assert node.ref == expected
    # rendering parses back to the same reference
    assert parse("=" + sref_text(node.ref)).ref == expected
    assert parse("=" + sref_text(node.ref, storage=True)).ref == expected


@pytest.mark.parametrize("bad", ["=T[#Foo]", "=T[[#This Row],[#Data]]", "=T[A", "=T[[#Headers],[#Totals]]"])
def test_bad_structured_refs(bad):
    with pytest.raises(FormulaSyntaxError):
        parse(bad)


def test_storage_and_display_forms():
    assert to_storage("=[@Qty]*Table1[@[Unit Price]]") == \
        "=[[#This Row],[Qty]]*Table1[[#This Row],[Unit Price]]"
    assert to_display("=Table1[[#This Row],[Qty]]*2") == "=Table1[@Qty]*2"
    assert to_display("=SUM(Table1[Qty])") == "=SUM(Table1[Qty])"


# --------------------------------------------------------------------------- evaluation


def test_evaluate_structured_refs(sales):
    ws = sales
    ws["F1"] = "=SUM(Table1[Qty])"
    ws["F2"] = "=COUNTA(Table1[#All])"
    ws["F3"] = "=Table1[[#Headers],[Price]]"
    ws["F4"] = "=SUM(Table1[[Qty]:[Price]])"
    ws["F5"] = "=ROWS(Table1)"
    ws["F6"] = "=ROWS(Table1[[#Headers],[#Data]])"
    ws["F7"] = "=Table1[@Price]"  # below the table
    ws["F8"] = "=Table1[[#This Row],[Price]]"
    ws["F9"] = "=SUM(Table1[Nope])"
    ws["F10"] = "=SUM(Nope[Qty])"
    ws["F11"] = "=Table1[#Totals]"
    ws["F12"] = "=SUM(table1[QTY])"  # case-insensitive
    assert [v(ws[f"F{i}"]) for i in range(1, 13)] == [
        12, 12, "Price", 27, 3, 4, "#VALUE!", "#VALUE!", "#REF!", "#REF!", "#REF!", 12]
    ws["G3"] = "=Table1[@Price]*10"  # same row as the data, outside the table: implicit intersection
    assert ws["G3"] == 50


def test_this_row_inside_table_and_recalc(sales):
    ws = sales
    ws["D1"] = "Total"  # typing next to the table adds a column
    assert _cols(ws.book) == ["Item", "Qty", "Price", "Total"]
    ws["D2"] = "=[@Qty]*[@Price]"
    # a formula in an empty column fills the column (calculated column)
    assert [ws["D2"], ws["D3"], ws["D4"]] == [6, 20, 42]
    assert ws.text("D3") == "=[@Qty]*[@Price]"
    cell = ws.sheet.ws._cells[(3, 4)]
    assert cell._value == "=Table1[[#This Row],[Qty]]*Table1[[#This Row],[Price]]"
    tc = _table(ws.book).table.tableColumns[3]
    assert tc.calculatedColumnFormula.attr_text == "Table1[[#This Row],[Qty]]*Table1[[#This Row],[Price]]"
    ws["H1"] = "=SUM(Table1[Total])"
    assert ws["H1"] == 68
    ws["B3"] = 10
    assert ws["D3"] == 50 and ws["H1"] == 98


def test_unqualified_reference_outside_table_is_rejected(sales):
    with pytest.raises(FormulaSyntaxError):
        sales["H1"] = "=[@Qty]"


def test_bare_table_name_and_names(sales):
    ws = sales
    ws["F1"] = "=SUM(Table1)"
    assert ws["F1"] == 2 + 3 + 4 + 5 + 6 + 7
    names.set_name(ws.book, "Qtys", "=Table1[Qty]", None, ws.sheet)
    ws["F2"] = "=SUM(Qtys)"
    assert ws["F2"] == 12
    ws["B2"] = 20
    assert ws["F2"] == 30
    assert names.validate_name(ws.book, "table1", None) is not None  # clashes with the table


# --------------------------------------------------------------------------- loading files


def _make_file(path, cached: bool = False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    rows = [["Region", "Units", "Unit Price", "Revenue"], ["East", 10, 2.5, None], ["West", 4, 10, None],
            ["North", 7, 3, None]]
    for r, row in enumerate(rows, 1):
        for c, x in enumerate(row, 1):
            ws.cell(row=r, column=c, value=x)
    for r in range(2, 5):
        ws.cell(row=r, column=4, value="=Sales[[#This Row],[Units]]*Sales[[#This Row],[Unit Price]]")
    ws.cell(row=5, column=1, value="Total")
    ws.cell(row=5, column=4, value="=SUBTOTAL(109,Sales[Revenue])")
    cols = [TableColumn(id=i, name=n) for i, n in enumerate(rows[0], 1)]
    cols[0].totalsRowLabel = "Total"
    cols[3].totalsRowFunction = "sum"
    cols[3].calculatedColumnFormula = TableFormula(attr_text="Sales[[#This Row],[Units]]*Sales[[#This Row],[Unit Price]]")
    t = Table(displayName="Sales", ref="A1:D5", tableColumns=cols, totalsRowCount=1)
    ws.add_table(t)
    other = wb.create_sheet("Report")
    formulas = ["=SUM(Sales[Revenue])", "=Sales[[#Totals],[Revenue]]", '=COUNTIF(Sales[Region],"*t")',
                '=SUMIFS(Sales[Revenue],Sales[Region],"West")', "=AVERAGE(Sales[Unit Price])",
                "=ROWS(Sales[#All])", "=INDEX(Sales[Region],2)", '=MATCH("North",Sales[Region],0)',
                "=MAX(Sales[[Units]:[Unit Price]])", "=SUM(Sales[[#Headers],[#Data],[Units]])"]
    for i, f in enumerate(formulas, 1):
        other.cell(row=i, column=1, value=f)
    wb.save(path)
    return formulas


EXPECTED = [86.0, 86.0, 2, 40, 5.166666666666667, 5, "West", 3, 10, 21]


def test_load_file_with_table(tmp_path):
    p = str(tmp_path / "t.xlsx")
    _make_file(p)
    book = xlsx.load(p)
    rep = book.sheet_by_name("Report")
    data = book.sheet_by_name("Data")
    got = [book.get_value(rep, i, 1) for i in range(1, 11)]
    assert all(math.isclose(a, b) if isinstance(b, float) else a == b for a, b in zip(got, EXPECTED)), got
    assert not data.stale and not rep.stale
    assert data.input_text(2, 4) == "=[@Units]*[@[Unit Price]]"
    # editing data flows through
    book.set_input(data, 2, 2, "20")
    assert book.get_value(data, 2, 4) == 50
    assert book.get_value(rep, 1, 1) == 111


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_structured_refs_against_libreoffice(tmp_path):
    p = tmp_path / "t.xlsx"
    _make_file(str(p))
    outdir = tmp_path / "lo"
    outdir.mkdir()
    env = dict(os.environ, HOME=str(tmp_path))
    subprocess.run([SOFFICE, "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", str(outdir), str(p)],
                   check=True, capture_output=True, timeout=180, env=env)
    lo = openpyxl.load_workbook(outdir / "t.xlsx", data_only=True)["Report"]
    theirs = [lo.cell(row=i, column=1).value for i in range(1, 11)]
    book = xlsx.load(str(p))
    rep = book.sheet_by_name("Report")
    ours = [book.get_value(rep, i, 1) for i in range(1, 11)]
    for a, b in zip(ours, theirs):
        assert (math.isclose(a, b) if isinstance(a, (int, float)) else a == b), (ours, theirs)


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_created_table_opens_in_libreoffice(sales, tmp_path):
    ws = sales
    ws["D1"] = "Total"
    ws["D2"] = "=[@Qty]*[@Price]"
    tables.set_totals(ws.book, _table(ws.book), True)
    ws["H1"] = "=SUM(Table1[Total])*2"
    out = tmp_path / "made.xlsx"
    xlsx.save(ws.book, str(out))
    # strip our cached values so LibreOffice must compute
    wb = openpyxl.load_workbook(out)
    wb.save(out)
    outdir = tmp_path / "lo"
    outdir.mkdir()
    env = dict(os.environ, HOME=str(tmp_path))
    subprocess.run([SOFFICE, "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", str(outdir), str(out)],
                   check=True, capture_output=True, timeout=180, env=env)
    lo = openpyxl.load_workbook(outdir / "made.xlsx", data_only=True).active
    assert [lo[a].value for a in ("D2", "D3", "D4", "D5", "H1")] == [6, 20, 42, 68, 136]


def test_save_round_trip(sales, tmp_path):
    ws = sales
    ws["D1"] = "Total"
    ws["D2"] = "=[@Qty]*[@Price]"
    tables.set_totals(ws.book, _table(ws.book), True)
    out = str(tmp_path / "rt.xlsx")
    xlsx.save(ws.book, out)
    wb = openpyxl.load_workbook(out)
    t = wb.active.tables["Table1"]
    assert t.ref == "A1:D5" and t.totalsRowCount == 1 and t.autoFilter.ref == "A1:D4"
    assert [c.name for c in t.tableColumns] == ["Item", "Qty", "Price", "Total"]
    assert t.tableColumns[3].totalsRowFunction == "sum"
    assert t.tableColumns[0].totalsRowLabel == "Total"
    assert wb.active["D5"].value == "=SUBTOTAL(109,Table1[Total])"
    assert wb.active["D3"].value == "=Table1[[#This Row],[Qty]]*Table1[[#This Row],[Price]]"
    book = xlsx.load(out)
    assert book.get_value(book.sheets[0], 5, 4) == 68
    assert _cols(book) == ["Item", "Qty", "Price", "Total"]


# --------------------------------------------------------------------------- creating


def test_create_table_fixes_headers(ws):
    ws["A1"] = "Name"
    ws["B1"] = "2024"
    ws["C1"] = "Name"
    ws["A2"] = "x"
    tables.create(ws.book, ws.sheet, Rect(1, 1, 3, 4))
    assert _cols(ws.book) == ["Name", "2024", "Name2", "Column1"]
    assert ws["B1"] == "2024" and ws["C1"] == "Name2" and ws["D1"] == "Column1"
    ws.book.do_undo()
    assert ws.sheet.ws.tables == {} and ws["B1"] == 2024 and ws["C1"] == "Name"
    ws.book.do_redo()
    assert _cols(ws.book) == ["Name", "2024", "Name2", "Column1"]


def test_create_without_headers_uses_row_above(ws):
    ws["A3"] = 1
    ws["A4"] = 2
    with pytest.raises(ValueError):
        tables.create(ws.book, ws.sheet, Rect(1, 1, 2, 1), has_headers=False)
    tables.create(ws.book, ws.sheet, Rect(3, 1, 4, 2), has_headers=False, name="Nums")
    assert _table(ws.book, "Nums").g.rect == Rect(2, 1, 4, 2)
    assert ws["A2"] == "Column1" and ws["B2"] == "Column2"
    ws["D1"] = "=SUM(Nums[Column1])"
    assert ws["D1"] == 3


def test_create_validation(sales):
    book, sheet = sales.book, sales.sheet
    with pytest.raises(ValueError, match="overlaps"):
        tables.create(book, sheet, Rect(3, 3, 6, 5))
    with pytest.raises(ValueError, match="already"):
        tables.create(book, sheet, Rect(10, 1, 12, 2), name="table1")
    for bad in ("A1", "my table", "1abc", "R1C1"):
        assert tables.validate_table_name(book, bad) is not None
    assert tables.default_name(book) == "Table2"


# --------------------------------------------------------------------------- editing


def test_rename_column_by_editing_header(sales):
    ws = sales
    ws["F1"] = "=SUM(Table1[Qty])"
    ws["E2"] = "=Table1[@Qty]*2"
    ws["B1"] = "Units"
    assert _cols(ws.book)[:3] == ["Item", "Units", "Price"]
    assert ws.text("F1") == "=SUM(Table1[Units])"
    assert ws.text("E2") == "=Table1[@Units]*2"
    assert ws["F1"] == 12
    ws.book.do_undo()
    assert ws.text("F1") == "=SUM(Table1[Qty])" and _cols(ws.book)[1] == "Qty" and ws["F1"] == 12
    ws.book.do_redo()
    assert ws.text("F1") == "=SUM(Table1[Units])" and _cols(ws.book)[1] == "Units"


def test_clearing_a_header_gives_a_default_name(sales):
    ws = sales
    ws["B1"] = ""
    assert ws["B1"] == "Column1" and _cols(ws.book)[1] == "Column1"
    ws["C1"] = "Item"  # duplicate
    assert ws["C1"] == "Item2"


def test_rename_table(sales):
    ws = sales
    ws["F1"] = "=SUM(Table1[Qty])+ROWS(Table1)"
    names.set_name(ws.book, "Prices", "=Table1[Price]", None, ws.sheet)
    tables.rename(ws.book, _table(ws.book), "Fruit")
    assert ws.text("F1") == "=SUM(Fruit[Qty])+ROWS(Fruit)" and ws["F1"] == 15
    assert ws.book.wb.defined_names["Prices"].attr_text == "Fruit[Price]"
    ws.book.do_undo()
    assert ws.text("F1") == "=SUM(Table1[Qty])+ROWS(Table1)" and ws["F1"] == 15
    assert ws.book.wb.defined_names["Prices"].attr_text == "Table1[Price]"
    assert _table(ws.book) is not None and _table(ws.book, "Fruit") is None


def test_auto_expand_below_and_right(sales):
    ws = sales
    ws["D1"] = "Total"
    ws["D2"] = "=[@Qty]*[@Price]"
    ws["H1"] = "=SUM(Table1[Qty])"
    ws["A5"] = "kiwi"
    assert _table(ws.book).g.rect == Rect(1, 1, 5, 4)
    assert ws.text("D5") == "=[@Qty]*[@Price]"  # calculated column continues
    ws["B5"] = 8
    ws["C5"] = 1
    assert ws["D5"] == 8 and ws["H1"] == 20
    ws.book.do_undo()
    ws.book.do_undo()
    ws.book.do_undo()
    assert _table(ws.book).g.rect == Rect(1, 1, 4, 4) and ws["A5"] is None and ws.text("D5") == ""


def test_totals_row(sales):
    ws = sales
    hit = _table(ws.book)
    tables.set_totals(ws.book, hit, True)
    assert ws["A5"] == "Total" and ws.text("C5") == "=SUBTOTAL(109,Table1[Price])" and ws["C5"] == 15
    ws["F1"] = "=Table1[[#Totals],[Price]]"
    assert ws["F1"] == 15
    ws["C5"] = "=SUBTOTAL(101,Table1[Price])"
    tc = _table(ws.book).table.tableColumns[2]
    assert tc.totalsRowFunction == "average" and ws["F1"] == 5
    ws["B5"] = "=MAX(Table1[Qty])*2"
    tc = _table(ws.book).table.tableColumns[1]
    assert tc.totalsRowFunction == "custom" and tc.totalsRowFormula.attr_text == "MAX(Table1[Qty])*2"
    # typing below a table with a totals row doesn't grow it
    ws["A6"] = "x"
    assert _table(ws.book).g.rect == Rect(1, 1, 5, 3)
    tables.set_totals(ws.book, _table(ws.book), False)
    assert ws["A5"] is None and _table(ws.book).g.rect == Rect(1, 1, 4, 3) and v(ws["F1"]) == "#REF!"
    # turning it back on restores the previous functions
    ws["A6"] = ""
    tables.set_totals(ws.book, _table(ws.book), True)
    assert ws.text("C5") == "=SUBTOTAL(101,Table1[Price])" and ws.text("B5") == "=MAX(Table1[Qty])*2"


def test_totals_needs_empty_row(sales):
    sales["B5"] = "x"
    # typing right below extends the table, so put the blocker below a totals-capable table
    hit = _table(sales.book)
    assert hit.g.r2 == 5
    sales["B6"] = "y"
    with pytest.raises(ValueError, match="empty"):
        tables.set_totals(sales.book, hit, True)


def test_resize(sales):
    ws = sales
    ws["D1"] = "Extra"  # becomes a column (auto-expand)
    tables.resize(ws.book, _table(ws.book), Rect(1, 1, 3, 2))
    assert _cols(ws.book) == ["Item", "Qty"] and _table(ws.book).g.rect == Rect(1, 1, 3, 2)
    ws["F1"] = "=SUM(Table1[Qty])"
    assert ws["F1"] == 6
    tables.resize(ws.book, _table(ws.book), Rect(1, 1, 4, 4))
    assert _cols(ws.book) == ["Item", "Qty", "Price", "Extra"] and ws["F1"] == 12
    with pytest.raises(ValueError, match="header"):
        tables.resize(ws.book, _table(ws.book), Rect(2, 1, 4, 4))
    tables.set_totals(ws.book, _table(ws.book), True)
    with pytest.raises(ValueError, match="Row 4 must be empty"):  # the totals row would land on data
        tables.resize(ws.book, _table(ws.book), Rect(1, 1, 4, 3))
    tables.resize(ws.book, _table(ws.book), Rect(1, 1, 6, 3))  # the range includes the totals row
    hit = _table(ws.book)
    assert hit.g.rect == Rect(1, 1, 6, 3) and hit.g.tot and ws["A6"] == "Total" and ws["A5"] is None
    assert _cols(ws.book) == ["Item", "Qty", "Price"]


def test_convert_to_range(sales):
    ws = sales
    ws["D1"] = "Total"
    ws["D2"] = "=[@Qty]*[@Price]"
    ws["F1"] = "=SUM(Table1[Total])"
    ws["F2"] = "=ROWS(Table1)"
    other = ws.book.sheets[0]
    structure.add_sheet(ws.book, "Other")
    o = Sheetish(ws.book, ws.book.sheet_by_name("Other"))
    o["A1"] = "=Table1[[#Headers],[Qty]]"
    tables.convert_to_range(ws.book, _table(ws.book))
    assert other.ws.tables == {}
    assert ws.text("D3") == "=$B3*$C3"
    assert ws.text("F1") == "=SUM($D$2:$D$4)" and ws["F1"] == 68
    assert ws.text("F2") == "=ROWS($A$2:$D$4)"
    assert o.text("A1") == "=Sheet!$B$1" and o["A1"] == "Qty"
    ws.book.do_undo()
    assert ws.text("F1") == "=SUM(Table1[Total])" and _table(ws.book) is not None and ws["F1"] == 68


def test_delete_table(sales):
    ws = sales
    ws["F1"] = "=COUNTA(Table1[Item])"
    tables.delete_table(ws.book, _table(ws.book))
    assert ws.sheet.ws.tables == {} and ws["A2"] is None and ws.text("F1") == "=COUNTA($A$2:$A$4)"
    ws.book.do_undo()
    assert ws["A2"] == "apple" and ws["F1"] == 3


# --------------------------------------------------------------------------- rows and columns


def test_insert_rows_inside_table(sales):
    ws = sales
    ws["D1"] = "Total"
    ws["D2"] = "=[@Qty]*[@Price]"
    ws["H1"] = "=SUM(Table1[Qty])"
    structure.insert(ws.book, ws.sheet, "row", 3, 2)
    hit = _table(ws.book)
    assert hit.g.rect == Rect(1, 1, 6, 4)
    assert ws.text("D3") == "=[@Qty]*[@Price]" and ws["D3"] == 0
    ws["B4"] = 100
    assert ws["H1"] == 112
    ws.book.do_undo()  # the B4 edit
    ws.book.do_undo()  # the insert
    assert _table(ws.book).g.rect == Rect(1, 1, 4, 4) and ws["H1"] == 12 and ws["D3"] == 20
    structure.insert(ws.book, ws.sheet, "row", 1, 1)  # above: the table moves
    assert _table(ws.book).g.rect == Rect(2, 1, 5, 4) and ws["H2"] == 12


def test_insert_and_delete_columns(sales):
    ws = sales
    ws["F1"] = "=SUM(Table1[Price])"
    structure.insert(ws.book, ws.sheet, "col", 2, 1)
    assert _cols(ws.book) == ["Item", "Column1", "Qty", "Price"] and ws["B1"] == "Column1"
    assert ws["G1"] == 15
    ws.book.do_undo()
    assert _cols(ws.book) == ["Item", "Qty", "Price"]
    structure.delete(ws.book, ws.sheet, "col", 2, 1)
    assert _cols(ws.book) == ["Item", "Price"] and _table(ws.book).g.rect == Rect(1, 1, 4, 2)
    assert ws["E1"] == 15
    ws.book.do_undo()
    assert _cols(ws.book) == ["Item", "Qty", "Price"] and ws["B2"] == 2 and ws["F1"] == 15
    structure.delete(ws.book, ws.sheet, "col", 1, 3)  # the whole table
    assert ws.sheet.ws.tables == {}
    ws.book.do_undo()
    assert _cols(ws.book) == ["Item", "Qty", "Price"]


def test_delete_rows(sales):
    ws = sales
    ws["F1"] = "=SUM(Table1[Qty])"
    structure.delete(ws.book, ws.sheet, "row", 3, 1)
    assert _table(ws.book).g.rect == Rect(1, 1, 3, 3) and ws["F1"] == 8
    ws.book.do_undo()
    assert _table(ws.book).g.rect == Rect(1, 1, 4, 3) and ws["F1"] == 12
    with pytest.raises(ValueError, match="header"):
        structure.delete(ws.book, ws.sheet, "row", 1, 2)
    with pytest.raises(ValueError, match="at least one row"):
        structure.delete(ws.book, ws.sheet, "row", 2, 3)
    structure.delete(ws.book, ws.sheet, "row", 1, 4)
    assert ws.sheet.ws.tables == {}


def test_parse_storage_render_matches_excel_spelling():
    assert sref_text(StructRef("T", frozenset({"thisrow"}), "Qty"), storage=True) == "T[[#This Row],[Qty]]"
    assert sref_text(StructRef("T", frozenset({"thisrow"}), "Unit Price")) == "T[@[Unit Price]]"
    assert sref_text(StructRef("T", frozenset({"totals"}))) == "T[#Totals]"
    assert sref_text(StructRef("T", frozenset(), "A", "B")) == "T[[A]:[B]]"
    assert sref_text(StructRef("T", frozenset({"headers", "data"}), "A")) == "T[[#Headers],[#Data],[A]]"
    assert parse_addr("A1") == (1, 1)


# --------------------------------------------------------------------------- UI


@pytest.mark.asyncio
async def test_ui_create_and_edit_table():
    from termxl.ui.app import TermXLApp
    from termxl.ui.dialogs import TableDialog

    from .test_ui import _ready, type_text

    app = TermXLApp()
    async with app.run_test(size=(110, 32)) as pilot:
        await _ready(app, pilot)
        book, sheet = app.book, app.grid.sheet
        for a1, x in {"A1": "Item", "B1": "Qty", "A2": "a", "B2": "3", "A3": "b", "B3": "4"}.items():
            book.set_input(sheet, *parse_addr(a1), x)
        app.after_change()
        await pilot.press("ctrl+t")
        await pilot.pause(0.1)
        assert isinstance(app.screen, TableDialog)
        await pilot.press("enter")  # current region A1:B3, with headers
        await pilot.pause(0.1)
        hit = tables.find_table(book, "Table1")
        assert hit is not None and hit.g.rect == Rect(1, 1, 3, 2)
        assert "Table1" in app.query_one("#namebox").value
        # a formula inside the table using [@Column]
        app._select(Rect(2, 3, 2, 3))
        await pilot.press("equals_sign", "left_square_bracket", "at", "Q", "t", "y", "right_square_bracket",
                          "asterisk", "2", "enter")
        assert book.get_value(sheet, 2, 3) == 6 and book.get_value(sheet, 3, 3) == 8
        assert tables.find_table(book, "Table1").col_names() == ["Item", "Qty", "Column1"]
        # Ctrl+T in the table: rename it and add a totals row
        await pilot.press("ctrl+t")
        await pilot.pause(0.1)
        assert isinstance(app.screen, TableDialog)
        await pilot.press("ctrl+a", "ctrl+k")
        await type_text(pilot, "Stock")
        app.screen.query_one("#totals").value = True
        await pilot.press("enter")
        await pilot.pause(0.1)
        hit = tables.find_table(book, "Stock")
        assert hit is not None and hit.g.tot and book.get_value(sheet, 4, 3) == 14
        app._select(Rect(2, 1, 2, 1))
        app.message = ""
        app.after_change()
        assert "Table Stock, column Item" in str(app.query_one("#status").render())
        # Go to a structured reference
        target = app._resolve_goto("Stock[Qty]")
        assert target == (sheet, Rect(2, 2, 3, 2))
        await pilot.press("ctrl+z")
        assert tables.find_table(book, "Table1") is not None and not tables.find_table(book, "Table1").g.tot
