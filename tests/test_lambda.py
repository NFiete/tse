"""LET, LAMBDA and the LAMBDA helpers (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, ISOMITTED).
LET is also cross-checked against LibreOffice in test_libreoffice.py; LibreOffice has no LAMBDA."""
import openpyxl
import pytest

from termxl.formula.errors import ExcelError
from termxl.formula.parser import FormulaSyntaxError, normalize, parse, to_display, to_storage
from termxl.io import xlsx
from termxl.model import names

from .test_dynamic_arrays import grid


def code(v):
    return v.code if isinstance(v, ExcelError) else v


@pytest.fixture
def data(ws):
    for i, (a, b) in enumerate([(1, 10), (2, 20), (3, 30)], 1):
        ws[f"A{i}"] = a
        ws[f"B{i}"] = b
    return ws


# ------------------------------------------------------------------ syntax


@pytest.mark.parametrize("display, storage", [
    ("=LET(x,1,x+1)", "=_xlfn.LET(_xlpm.x,1,_xlpm.x+1)"),
    ("=LET(f,LAMBDA(x,x+1),f(2))", "=_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.f(2))"),
    ("=LAMBDA(x,x+1)(2)", "=_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)"),
    ("=MAP(A1:A3,LAMBDA(v,v*2))", "=_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))"),
    ("=BYROW(A1:B3,SUM)", "=_xlfn.BYROW(A1:B3,_xleta.SUM)"),
    ("=LET(x,B3,LET(x,x+1,x)*2)", "=_xlfn.LET(_xlpm.x,B3,_xlfn.LET(_xlpm.x,_xlpm.x+1,_xlpm.x)*2)"),
    ("=LET(rate,0.1,rate)+rate", "=_xlfn.LET(_xlpm.rate,0.1,_xlpm.rate)+rate"),  # the last one is a defined name
])
def test_storage_form(display, storage):
    assert to_storage(display) == storage
    assert to_display(storage) == display


def test_local_names_keep_their_case():
    assert normalize("=let(Total,sum(a1:a3),f,lambda(x,x*Total),f(2))") == \
        "=LET(Total,SUM(A1:A3),f,LAMBDA(x,x*Total),f(2))"


@pytest.mark.parametrize("formula", [
    "=LET(x,1)", "=LET(x,1,y)(2)+LET(x,1,y,2)", "=LET(1,2,3)", "=LET(A1,2,A1)", "=LAMBDA(x,x,x)",
    "=LAMBDA(a.b,1)", "=LET(Sheet1!x,1,2)",
])
def test_malformed(formula):
    with pytest.raises(FormulaSyntaxError):
        parse(formula)


# ------------------------------------------------------------------ LET and LAMBDA


@pytest.mark.parametrize("formula, expected", [
    ("=LET(x,1,x+1)", 2),
    ("=LET(x,A2,y,x*3,x+y)", 8),
    ("=LET(r,A1:A3,SUM(r)/ROWS(r))", 2),
    ("=LET(x,1/0,5)", 5),
    ("=LET(x,1/0,x)", "#DIV/0!"),
    ("=LET(x,B3,LET(x,x+1,x)*2)", 62),
    ("=LET(f,LAMBDA(x,x+1),f(2))", 3),
    ("=LAMBDA(x,y,x*y)(A3,B3)", 90),
    ("=LAMBDA(42)()", 42),
    ("=LAMBDA(x,x+1)", "#CALC!"),
    ("=LET(f,LAMBDA(x,x),f)", "#CALC!"),
    ("=LAMBDA(x,y,x+y)(1)", "#VALUE!"),
    ("=LAMBDA(x,x)(1,2)", "#VALUE!"),
    ('=LAMBDA(x,y,IF(ISOMITTED(y),"none",x+y))(1,)', "none"),
    ('=LAMBDA(x,y,IF(ISOMITTED(y),"none",x+y))(1,2)', 3),
    ("=LET(a,5,f,LAMBDA(x,x+a),f(1))", 6),  # a LAMBDA sees the LET names around it
    ("=LAMBDA(x,LET(y,x,LAMBDA(z,y*z)))(3)(4)", 12),
    ("=LET(f,SUM,f(1,2,3))", 6),  # a function name is a function value
    ("=LET(x,5,f,LAMBDA(x,x*2),f(1)+x)", 7),  # a parameter shadows an outer name
    ("=LET(n,2,n)(1)", "#VALUE!"),  # calling something that isn't a function
    ("=LET(x,A1:A3,x*2)", 2),  # spills; Z1 shows the first value
    ("=REDUCE(0,A1:B3,LAMBDA(a,v,a+v))", 66),
    ("=REDUCE(1,A1:A3,LAMBDA(a,v,a*v))", 6),
    ("=REDUCE(,A1:A3,LAMBDA(a,v,IF(ISOMITTED(a),v,a+v)))", 6),
])
def test_let_and_lambda(data, formula, expected):
    data["Z1"] = formula
    assert code(data["Z1"]) == expected


def test_let_result_spills(data):
    assert grid(data, "=LET(x,A1:A3,x*2)") == [[2], [4], [6]]
    data["D1"] = "=LET(r,A1:B3,r)"
    assert [data[a] for a in ("D1", "E1", "D3", "E3")] == [1, 10, 3, 30]


def test_let_recalculates(data):
    data["D1"] = "=LET(x,A1,f,LAMBDA(v,v+B1),f(x))"
    assert data["D1"] == 11
    data["A1"] = 5
    data["B1"] = 100
    assert data["D1"] == 105


# ------------------------------------------------------------------ helpers


@pytest.mark.parametrize("formula, expected", [
    ("=MAP(A1:A3,LAMBDA(v,v*2))", [[2], [4], [6]]),
    ("=MAP(A1:A3,B1:B3,LAMBDA(a,b,a+b))", [[11], [22], [33]]),
    ("=MAP(A1:A3,10,LAMBDA(a,b,a*b))", [[10], [20], [30]]),
    ("=MAP(A1:A3,B1:B2,LAMBDA(a,b,a+b))", [[11], [22], ["#N/A"]]),
    ("=MAP(A1:A3,LAMBDA(c,ROW(c)))", [[1], [2], [3]]),  # ranges pass cell references
    ("=MAP({1,2},LAMBDA(v,{1,2}))", [["#CALC!", "#CALC!"]]),
    ("=MAP(A1:A3,LAMBDA(a,b,a))", "#VALUE!"),
    ("=MAP(A1:A3,5)", "#VALUE!"),
    ('=REDUCE("",A1:A3,LAMBDA(a,v,IF(v>1,VSTACK(a,v),a)))', [[""], [2], [3]]),
    ("=SCAN(0,A1:A3,LAMBDA(a,v,a+v))", [[1], [3], [6]]),
    ('=SCAN("",{"a","b";"c","d"},LAMBDA(a,v,a&v))', [["a", "ab"], ["abc", "abcd"]]),
    ("=BYROW(A1:B3,LAMBDA(r,SUM(r)))", [[11], [22], [33]]),
    ("=BYROW(A1:B3,MAX)", [[10], [20], [30]]),
    ("=BYROW(A1:B3,LAMBDA(r,r))", [["#CALC!"], ["#CALC!"], ["#CALC!"]]),
    ("=BYCOL(A1:B3,LAMBDA(c,SUM(c)))", [[6, 60]]),
    ("=BYCOL({1,2;3,4},LAMBDA(c,PRODUCT(c)))", [[3, 8]]),
    ("=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))", [[11, 12, 13], [21, 22, 23]]),
    ("=MAKEARRAY(0,3,LAMBDA(r,c,r))", "#VALUE!"),
    ("=MAKEARRAY(2,2,LAMBDA(r,r))", "#VALUE!"),
])
def test_helpers(data, formula, expected):
    res = grid(data, formula)
    if isinstance(res, list):
        res = [[code(v) for v in row] for row in res]
    assert res == expected


def test_isomitted_outside_lambda(ev):
    assert ev("=ISOMITTED(1)") is False


def test_groupby_with_lambda(ws):
    for i, (k, v) in enumerate([("a", 1), ("b", 2), ("a", 5)], 1):
        ws[f"A{i}"] = k
        ws[f"B{i}"] = v
    assert grid(ws, "=GROUPBY(A1:A3,B1:B3,LAMBDA(x,MAX(x)-MIN(x)),0,0)") == [["a", 4], ["b", 0]]
    assert grid(ws, "=GROUPBY(A1:A3,B1:B3,LAMBDA(x,all,SUM(x)/SUM(all)),0,0)") == [["a", 0.75], ["b", 0.25]]


# ------------------------------------------------------------------ LAMBDAs in defined names


def test_named_lambda(data):
    book, sheet = data.book, data.sheet
    names.set_name(book, "ADDB", "=LAMBDA(x,x+$B$1)", None, sheet)
    data["D1"] = "=ADDB(A3)"
    assert data["D1"] == 13
    data["B1"] = 100  # the name's references are dependencies of the cell
    assert data["D1"] == 103
    assert data.sheet.formulas[(1, 4)].unsupported == []


def test_named_lambda_defined_later(data):
    data["D1"] = "=TRIPLE(2)"
    assert code(data["D1"]) == "#NAME?"
    names.set_name(data.book, "TRIPLE", "=LAMBDA(x,x*3)", None, data.sheet)
    assert data["D1"] == 6


def test_recursive_named_lambda(data):
    names.set_name(data.book, "FACT2", "=LAMBDA(n,IF(n<=1,1,n*FACT2(n-1)))", None, data.sheet)
    data["D1"] = "=FACT2(10)"
    assert data["D1"] == 3628800
    data["D2"] = "=FACT2(100000)"  # runaway recursion
    assert code(data["D2"]) == "#NUM!"


def test_named_lambda_doesnt_see_callers_locals(data):
    names.set_name(data.book, "GETX", "=LAMBDA(v,v+x)", None, data.sheet)
    data["D1"] = "=LET(x,5,GETX(1))"
    assert code(data["D1"]) == "#NAME?"


def test_rename_named_lambda(data):
    book, sheet = data.book, data.sheet
    names.set_name(book, "DBL", "=LAMBDA(x,x*2)", None, sheet)
    names.set_name(book, "rate", "=0.5", None, sheet)
    data["D1"] = "=DBL(3)+LET(rate,1,rate)+rate"
    assert data["D1"] == 7.5
    names.set_name(book, "TWICE", "=LAMBDA(x,x*2)", None, sheet, old=("DBL", None))
    names.set_name(book, "pct", "=0.5", None, sheet, old=("rate", None))
    assert data.text("D1") == "=TWICE(3)+LET(rate,1,rate)+pct"
    assert data["D1"] == 7.5


# ------------------------------------------------------------------ files


def test_save_and_reload(data, tmp_path):
    book = data.book
    names.set_name(book, "ADDB", "=LAMBDA(x,x+$B$1)", None, data.sheet)
    data["D1"] = "=LET(f,LAMBDA(x,x+1),f(A1))"
    data["E1"] = "=SCAN(0,A1:A3,LAMBDA(a,v,a+v))"
    data["F1"] = "=ADDB(1)"
    out = tmp_path / "lambda.xlsx"
    xlsx.save(book, str(out))

    raw = openpyxl.load_workbook(out)
    assert raw.active["D1"].value == "=_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.f(A1))"
    assert raw.defined_names["ADDB"].attr_text == "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+Sheet!$B$1)"

    book2 = xlsx.load(str(out))
    s = book2.sheets[0]
    assert s.input_text(1, 4) == "=LET(f,LAMBDA(x,x+1),f(A1))"
    assert [book2.get_value(s, r, 5) for r in (1, 2, 3)] == [1, 3, 6]
    assert book2.get_value(s, 1, 6) == 11
    book2.set_input(s, 1, 1, "5")
    assert book2.get_value(s, 1, 4) == 6
    assert book2.get_value(s, 3, 5) == 10
