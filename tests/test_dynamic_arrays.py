"""Dynamic-array and lookup functions: GROUPBY, PIVOTBY, TRIMRANGE, SORTBY, TAKE/DROP and friends.
Most simple cases are also cross-checked against LibreOffice in test_libreoffice.py."""
import pytest

from termse.formula.errors import ExcelError
from termse.formula.evaluator import Ctx, evaluate
from termse.formula.parser import to_display, to_storage
from termse.formula.values import Array, RefValue
from termse.io import xlsx
from termse.model.refs import parse_addr

SALES = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["West", "B", 5],
    ["East", "B", 7],
    ["West", "A", 3],
    ["East", "A", 1],
]


@pytest.fixture
def sales(ws):
    for r, row in enumerate(SALES, 1):
        for c, v in enumerate(row):
            ws["ABC"[c] + str(r)] = str(v)
    return ws


def grid(ws, formula):
    """Evaluate a formula in Z1 and return its whole result as a list of rows."""
    ws["Z1"] = formula
    r, c = parse_addr("Z1")
    f = ws.sheet.formulas[(r, c)]
    v = evaluate(f.ast, Ctx(ws.book, ws.sheet, r, c))
    if isinstance(v, ExcelError):
        return v.code
    if isinstance(v, RefValue):
        v = Array(ws.book.range_rows(v.sheet, v.r1, v.c1, v.r2, v.c2))
    assert isinstance(v, Array), v
    return [[x.code if isinstance(x, ExcelError) else x for x in row] for row in v.rows]


def test_groupby_basic(sales):
    assert grid(sales, "=GROUPBY(A1:A6,C1:C6,SUM)") == [
        ["Region", "Sales"], ["East", 18], ["West", 8], ["Total", 26]]
    assert sales["Z1"] == "Region"
    # no header row: nothing is detected or shown
    assert grid(sales, "=GROUPBY(A2:A6,C2:C6,SUM)") == [["East", 18], ["West", 8], ["Total", 26]]
    assert grid(sales, "=GROUPBY(A2:A6,C2:C6,AVERAGE,0,0)") == [["East", 6], ["West", 4]]


def test_groupby_subtotals_and_sorting(sales):
    assert grid(sales, "=GROUPBY(A2:B6,C2:C6,SUM,0)") == [
        ["East", "A", 11], ["East", "B", 7], ["East", None, 18],
        ["West", "A", 3], ["West", "B", 5], ["West", None, 8],
        ["Total", None, 26]]
    # grand total and subtotals on top, groups sorted by sales, descending
    assert grid(sales, "=GROUPBY(A2:B6,C2:C6,SUM,0,-2,-3)") == [
        ["Total", None, 26],
        ["East", None, 18], ["East", "A", 11], ["East", "B", 7],
        ["West", None, 8], ["West", "B", 5], ["West", "A", 3]]
    assert grid(sales, "=GROUPBY(A2:A6,C2:C6,SUM,0,1,-1)") == [["West", 8], ["East", 18], ["Total", 26]]
    # table relationship: no subtotals
    assert grid(sales, "=GROUPBY(A2:B6,C2:C6,SUM,0,2,{2,1},,1)") == [
        ["East", "A", 11], ["West", "A", 3], ["East", "B", 7], ["West", "B", 5], ["Total", None, 26]]
    assert grid(sales, "=GROUPBY(A2:A6,C2:C6,SUM,0,1,5)") == "#VALUE!"


def test_groupby_filter_and_functions(sales):
    assert grid(sales, "=GROUPBY(A2:A6,C2:C6,SUM,0,1,,C2:C6>4)") == [["East", 17], ["West", 5], ["Total", 22]]
    assert grid(sales, "=GROUPBY(A2:A6,C2:C6,SUM,0,1,,C2:C6>100)") == "#CALC!"
    assert grid(sales, "=GROUPBY(A1:A6,C1:C6,HSTACK(SUM,COUNT),3,0)") == [
        ["Region", "SUM", "COUNT"], ["East", 18, 3], ["West", 8, 2]]
    pct = grid(sales, "=GROUPBY(A2:A6,C2:C6,PERCENTOF)")
    assert [row[0] for row in pct] == ["East", "West", "Total"]
    assert pct[0][1] == pytest.approx(18 / 26) and pct[2][1] == 1
    assert grid(sales, "=GROUPBY(A2:A6,C2:C6,42)") == "#VALUE!"


def test_pivotby(sales):
    assert grid(sales, "=PIVOTBY(A2:A6,B2:B6,C2:C6,SUM)") == [
        [None, "A", "B", "Total"], ["East", 11, 7, 18], ["West", 3, 5, 8], ["Total", 14, 12, 26]]
    assert grid(sales, "=PIVOTBY(A2:A6,B2:B6,C2:C6,SUM,0,0,,0)") == [[None, "A", "B"], ["East", 11, 7], ["West", 3, 5]]
    # percent of the grand total
    pct = grid(sales, "=PIVOTBY(A2:A6,B2:B6,C2:C6,PERCENTOF,0,0,,0,,,2)")
    assert pct[1][1:] == [pytest.approx(11 / 26), pytest.approx(7 / 26)]
    # percent of the column total (the default)
    pct = grid(sales, "=PIVOTBY(A2:A6,B2:B6,C2:C6,PERCENTOF,0,0,,0)")
    assert pct[1][1:] == [pytest.approx(11 / 14), pytest.approx(7 / 12)]
    # an empty combination is blank
    sales["C3"] = ""
    sales["B3"] = "C"
    assert grid(sales, "=PIVOTBY(A2:A6,B2:B6,C2:C6,SUM,0,0,,0)") == [
        [None, "A", "B", "C"], ["East", 11, 7, None], ["West", 3, None, 0]]


def test_function_names_are_stored_with_xleta():
    stored = to_storage("=GROUPBY(A1:A6,C1:C6,HSTACK(sum,AVERAGE))")
    assert stored == "=_xlfn.GROUPBY(A1:A6,C1:C6,_xlfn.HSTACK(_xleta.SUM,_xleta.AVERAGE))"
    assert to_display(stored) == "=GROUPBY(A1:A6,C1:C6,HSTACK(SUM,AVERAGE))"
    # only inside functions that take functions
    assert to_storage("=SUM(A1)+DAYS") == "=SUM(A1)+DAYS"
    assert to_storage("=SORTBY(A1:A3,B1:B3)") == "=_xlfn.SORTBY(A1:A3,B1:B3)"


def test_bare_function_name_in_a_cell(ev):
    assert ev("=SUM").code == "#NAME?"
    assert ev("=NOSUCHNAME").code == "#NAME?"


def test_groupby_round_trip(sales, tmp_path):
    sales["E1"] = "=GROUPBY(A1:A6,C1:C6,SUM)"
    sales["H1"] = "=INDEX(GROUPBY(A2:A6,C2:C6,SUM),2,2)"
    p = tmp_path / "g.xlsx"
    xlsx.save(sales.book, str(p))
    import openpyxl

    saved = openpyxl.load_workbook(p).active
    assert saved["E1"].value.text == "=_xlfn.GROUPBY(A1:A6,C1:C6,_xleta.SUM)"
    assert saved["E1"].value.ref == "E1:F4"
    # needs array evaluation, so it's saved as a one-cell dynamic array
    assert saved["H1"].value.text == "=INDEX(_xlfn.GROUPBY(A2:A6,C2:C6,_xleta.SUM),2,2)"
    assert saved["H1"].value.ref == "H1"
    book = xlsx.load(str(p))
    s = book.sheets[0]
    assert s.formulas[(1, 5)].text == "=GROUPBY(A1:A6,C1:C6,SUM)"
    assert book.get_value(s, 3, 6) == 8
    assert book.get_value(s, 1, 8) == 8
    assert not s.stale


def test_trimrange(ws):
    ws["B3"] = "1"
    ws["D5"] = "2"
    assert grid(ws, "=TRIMRANGE(A1:F10)") == [[1, None, None], [None, None, None], [None, None, 2]]
    ws["H1"] = "=ROWS(TRIMRANGE(A1:F10))&\",\"&COLUMNS(TRIMRANGE(A1:F10))"
    assert ws["H1"] == "3,3"
    ws["H1"] = "=ROWS(TRIMRANGE(A1:F10,2,1))&\",\"&COLUMNS(TRIMRANGE(A1:F10,2,1))"
    assert ws["H1"] == "5,5"  # A1:F10 -> B1:F5: trailing rows, leading columns
    ws["H1"] = "=ROWS(TRIMRANGE(A1:F10,0,0))"
    assert ws["H1"] == 10
    ws["H1"] = "=SUM(TRIMRANGE(A:E))"
    assert ws["H1"] == 3
    ws["H1"] = "=ROW(TRIMRANGE(A1:F10))"
    assert ws["H1"] == 3
    ws["H1"] = "=ROWS(TRIMRANGE(J1:K5))"  # entirely blank: the top-left cell
    assert ws["H1"] == 1
    assert grid(ws, "=TRIMRANGE({\"\",1;2,3})") == [["", 1], [2, 3]]
    assert grid(ws, "=TRIMRANGE(A1:F10,4)") == "#VALUE!"


def test_trim_references(ws):
    ws["B3"] = "1"
    ws["D5"] = "2"
    assert grid(ws, "=A1:.F10") == grid(ws, "=TRIMRANGE(A1:F10,2,2)")
    assert grid(ws, "=A1.:F10") == grid(ws, "=TRIMRANGE(A1:F10,1,1)")
    assert grid(ws, "=A1.:.F10") == [[1, None, None], [None, None, None], [None, None, 2]]
    ws["Z1"] = ""
    ws["H1"] = "=ROWS(A1:.F10)&\",\"&COLUMNS(A1:.F10)"
    assert ws["H1"] == "5,4"
    ws["H1"] = "=ROWS(A1.:F10)&\",\"&COLUMNS(A1.:F10)"
    assert ws["H1"] == "8,5"
    ws["H1"] = "=SUM(B:.B)+COUNTA(5:.5)"
    assert ws["H1"] == 2
    ws["H2"] = "=ROWS(B:.B)"
    assert ws["H2"] == 3
    ws["B9"] = "7"  # the trimmed range grows with the data
    assert ws["H2"] == 9
    assert ws["H1"] == 9
    # the operator form, between things that aren't plain references
    ws["H1"] = "=ROWS(B1:.INDEX(B:B,20))"
    assert ws["H1"] == 9
    ws["H1"] = "=sum(b:.b)"
    assert ws.text("H1") == "=SUM(B:.B)"


def test_trim_reference_storage():
    pairs = [
        ("=SUM(A:.A)", "=SUM(_xlfn._TRO_TRAILING(A:A))"),
        ("=A1.:E10", "=_xlfn._TRO_LEADING(A1:E10)"),
        ("=$A$1.:.$E10", "=_xlfn._TRO_ALL($A$1:$E10)"),
        ("=Sheet2!1:.3", "=_xlfn._TRO_TRAILING(Sheet2!1:3)"),
        ("=XLOOKUP(1,A:.A,B:.B)", "=_xlfn.XLOOKUP(1,_xlfn._TRO_TRAILING(A:A),_xlfn._TRO_TRAILING(B:B))"),
        ("=@A1.:B2", "=_xlfn.SINGLE(_xlfn._TRO_LEADING(A1:B2))"),
        ("=LEN(A1:.INDEX(B:B,5))", "=LEN(_xlfn._TRO_TRAILING(A1:INDEX(B:B,5)))"),
        ("=A1.:B2:.C3", "=_xlfn._TRO_TRAILING(_xlfn._TRO_LEADING(A1:B2):C3)"),
        ("=A1:.B2:C3", "=_xlfn._TRO_TRAILING(A1:B2):C3"),
    ]
    for display, stored in pairs:
        assert to_storage(display) == stored
        assert to_display(stored) == display
    # calls that can't be written as an operator stay calls
    assert to_display("=_xlfn._TRO_ALL(A1)") == "=_TRO_ALL(A1)"
    assert to_display("=_xlfn._TRO_ALL(INDEX(A:A,1))") == "=_TRO_ALL(INDEX(A:A,1))"
    assert to_storage("=_TRO_ALL(A1)") == "=_xlfn._TRO_ALL(A1)"


def test_trim_references_follow_edits(ws):
    from termse.formula.parser import translate
    from termse.model import structure

    assert translate("=SUM(A1:.B5)+SUM(C:.C)", 2, 1) == "=SUM(B3:.C7)+SUM(D:.D)"
    ws["A1"] = "1"
    ws["A2"] = "2"
    ws["C1"] = "=SUM(A1.:.A9)"
    structure.insert(ws.book, ws.sheet, "row", 1, 1)
    assert ws.text("C2") == "=SUM(A2.:.A10)"
    assert ws["C2"] == 3


def test_trim_reference_round_trip(ws, tmp_path):
    ws["A1"] = "1"
    ws["A2"] = "2"
    ws["B1"] = "=SUM(A:.A)"
    ws["C1"] = "=A1.:.A5*10"
    p = tmp_path / "t.xlsx"
    xlsx.save(ws.book, str(p))
    import openpyxl

    saved = openpyxl.load_workbook(p).active
    assert saved["B1"].value == "=SUM(_xlfn._TRO_TRAILING(A:A))"
    assert saved["C1"].value.text == "=_xlfn._TRO_ALL(A1:A5)*10"
    book = xlsx.load(str(p))
    s = book.sheets[0]
    assert s.formulas[(1, 2)].text == "=SUM(A:.A)"
    assert book.get_value(s, 1, 2) == 3
    assert book.get_value(s, 2, 3) == 20
    assert not s.stale


def test_shapes_keep_full_reference_size(ws):
    ws["A1"] = "1"
    ws["A2"] = "2"
    ws["H1"] = "=ROWS(TAKE(A1:A10,-3))"
    assert ws["H1"] == 3
    ws["H1"] = "=TAKE(A1:A10,-1)"
    assert ws["H1"] == 0  # A10 is blank
    ws["H1"] = "=ROWS(DROP(A:A,1))"
    assert ws["H1"] == 1048575
    ws["H1"] = "=ROWS(VSTACK(A1:A5,A1:A2))"
    assert ws["H1"] == 7
    ws["H1"] = "=COUNTA(TOCOL(A1:B10,1))"
    assert ws["H1"] == 2
    ws["H1"] = "=INDEX(SORTBY(A:A,A:A,-1),1)"
    assert ws["H1"] == 2


def test_shape_errors(ev):
    assert ev("=TAKE({1,2;3,4},0)").code == "#CALC!"
    assert ev("=DROP({1,2;3,4},-2)").code == "#CALC!"
    assert ev("=EXPAND({1,2;3,4},1)").code == "#VALUE!"
    assert ev("=WRAPROWS({1,2;3,4},2)").code == "#VALUE!"
    assert ev("=WRAPCOLS({1,2,3},0)").code == "#NUM!"
    assert ev("=CHOOSEROWS({1;2},3)").code == "#VALUE!"
    assert ev("=CHOOSECOLS({1,2},-3)").code == "#VALUE!"
    assert ev("=TOCOL({\"\",1},3)") == ""
    assert ev("=SORTBY({1;2;3},{1;2})").code == "#VALUE!"
    assert ev("=SORTBY({1;2;3},{3;2;1},2)").code == "#VALUE!"


def test_shapes(ev):
    assert ev("=INDEX(CHOOSEROWS({1;2;3},{3,1}),2)") == 1
    assert ev("=INDEX(VSTACK({1,2},3),2,2)").code == "#N/A"
    assert ev("=INDEX(HSTACK({1;2},{3}),2,2)").code == "#N/A"
    assert ev("=INDEX(TOROW({1,#N/A;2,3},2),1,2)") == 2
    assert ev("=INDEX(WRAPCOLS({1,2,3,4,5},2),2,3)").code == "#N/A"
    assert ev("=INDEX(SORTBY({\"a\",\"b\",\"c\"},{2,3,1}),1,1)") == "c"
    assert ev("=INDEX(SORTBY({1;2;3;4},{\"x\";\"y\";\"x\";\"y\"},-1,{1;2;3;4},1),2)") == 4


def test_formulatext_follows_edits(ws):
    ws["A1"] = "=1+1"
    ws["B1"] = "=FORMULATEXT(A1)"
    assert ws["B1"] == "=1+1"
    ws["A1"] = "=2*3"
    assert ws["B1"] == "=2*3"
    ws["A1"] = "5"
    assert ws["B1"].code == "#N/A"
    ws["B1"] = "=FORMULATEXT(\"A1\")"
    assert ws["B1"].code == "#VALUE!"


def test_hyperlink_and_areas(ev):
    assert ev('=HYPERLINK("https://example.com","Example")') == "Example"
    assert ev('=HYPERLINK("https://example.com")') == "https://example.com"
    assert ev("=AREAS(B2:D4)") == 1
    assert ev("=AREAS(5)").code == "#VALUE!"


def test_row_and_column_of_ranges_spill(ws):
    assert grid(ws, "=ROW(1:10)") == [[r] for r in range(1, 11)]
    assert grid(ws, "=ROW(B3:D5)") == [[3], [4], [5]]
    assert grid(ws, "=COLUMN(B:D)") == [[2, 3, 4]]
    ws["H1"] = "=SUM(ROW(A1:A10))"
    assert ws["H1"] == 55
    ws["H1"] = "=ROW(C7)+COLUMN(C7)"
    assert ws["H1"] == 10
