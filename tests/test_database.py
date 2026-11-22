"""Database functions (DSUM, DCOUNT, DGET...), using Microsoft's documented example."""
import pytest

from termse.formula.errors import DIV0, NUM, VALUE

CRITERIA = [
    ["Tree", "Height", "Age", "Yield", "Profit", "Height"],
    ['="=Apple"', ">10", None, None, None, "<16"],
    ['="=Pear"'],
]
DATA = [
    ["Tree", "Height", "Age", "Yield", "Profit"],
    ["Apple", 18, 20, 14, 105],
    ["Pear", 12, 12, 10, 96],
    ["Cherry", 13, 14, 9, 105],
    ["Apple", 14, 15, 10, 75],
    ["Pear", 9, 8, 8, 76.8],
    ["Apple", 8, 9, 6, 45],
]


@pytest.fixture
def db(ws):
    for r, row in enumerate(CRITERIA + [[]] + DATA, start=1):
        for c, v in enumerate(row):
            if v is not None:
                ws["ABCDEF"[c] + str(r)] = v
    return ws


def run(ws, formula):
    ws["H1"] = formula
    return ws["H1"]


@pytest.mark.parametrize("formula, expected", [
    ('=DCOUNT(A5:E11,"Age",A1:F2)', 1),
    ('=DCOUNTA(A5:E11,"Profit",A1:F2)', 1),
    ('=DMAX(A5:E11,"Profit",A1:A3)', 105),
    ('=DMIN(A5:E11,"Profit",A1:B2)', 75),
    ('=DSUM(A5:E11,"Profit",A1:A2)', 225),
    ('=DSUM(A5:E11,"Profit",A1:F2)', 75),
    ('=DPRODUCT(A5:E11,"Yield",A1:F2)', 10),
    ('=DAVERAGE(A5:E11,"Yield",A1:B2)', 12),
    ('=DAVERAGE(A5:E11,3,A5:E11)', 13),
    ('=DSTDEV(A5:E11,"Yield",A1:A3)', pytest.approx(2.966479395)),
    ('=DSTDEVP(A5:E11,"Yield",A1:A3)', pytest.approx(2.653299832)),
    ('=DVAR(A5:E11,"Yield",A1:A3)', pytest.approx(8.8)),
    ('=DVARP(A5:E11,"Yield",A1:A3)', pytest.approx(7.04)),
    ('=DGET(A5:E11,"Yield",A1:A3)', NUM),
    # field by index, by reference to a header cell, case-insensitive
    ('=DSUM(A5:E11,5,A1:A2)', 225),
    ('=DSUM(A5:E11,E5,A1:A2)', 225),
    ('=DSUM(A5:E11,"pROFIT",A1:A2)', 225),
    # omitted field counts matching records
    ("=DCOUNT(A5:E11,,A1:A3)", 5),
    ("=DCOUNTA(A5:E11,,A1:F2)", 1),
])
def test_microsoft_examples(db, formula, expected):
    assert run(db, formula) == expected


def test_bad_field(db):
    assert run(db, '=DSUM(A5:E11,"Nope",A1:A2)') == VALUE
    assert run(db, "=DSUM(A5:E11,6,A1:A2)") == VALUE
    assert run(db, "=DSUM(A5:E11,0,A1:A2)") == VALUE
    assert run(db, "=DSUM(A5:E11,,A1:A2)") == VALUE


def test_text_criterion_is_a_prefix(db):
    db["A2"] = "App"
    assert run(db, '=DSUM(A5:E11,"Profit",A1:A2)') == 225
    db["A2"] = "=\"=App\""
    assert run(db, '=DSUM(A5:E11,"Profit",A1:A2)') == 0
    db["A2"] = "*rr*"
    assert run(db, '=DGET(A5:E11,"Profit",A1:A2)') == 105
    db["A2"] = "<>Apple"
    assert run(db, '=DCOUNT(A5:E11,"Profit",A1:A2)') == 3


def test_rows_are_or_columns_are_and(db):
    # Pear, or Apple with 10 < height < 16
    db["A3"] = '="=Pear"'
    assert run(db, '=DSUM(A5:E11,"Profit",A1:F3)') == 75 + 96 + 76.8


def test_blank_criteria_row_matches_everything(db):
    assert run(db, '=DCOUNT(A5:E11,"Profit",A1:A4)') == 6
    assert run(db, '=DCOUNT(A5:E11,"Profit",C1:C2)') == 6  # Age has no condition
    assert run(db, '=DCOUNT(A5:E11,"Profit",A1:A100)') == 6  # blank rows past the used area


def test_no_matches(db):
    db["A2"] = '="=Plum"'
    assert run(db, '=DSUM(A5:E11,"Profit",A1:A2)') == 0
    assert run(db, '=DMAX(A5:E11,"Profit",A1:A2)') == 0
    assert run(db, '=DPRODUCT(A5:E11,"Profit",A1:A2)') == 0
    assert run(db, '=DAVERAGE(A5:E11,"Profit",A1:A2)') == DIV0
    assert run(db, '=DGET(A5:E11,"Profit",A1:A2)') == VALUE
    assert run(db, '=DSTDEV(A5:E11,"Profit",A1:A2)') == DIV0


def test_dget_single_match(db):
    db["A2"] = "Cherry"
    assert run(db, '=DGET(A5:E11,"Height",A1:A2)') == 13
    assert run(db, '=DGET(A5:E11,"Tree",A1:A2)') == "Cherry"


def test_criteria_needs_a_condition_row(db):
    assert run(db, '=DSUM(A5:E11,"Profit",A1)') == VALUE


def test_errors(db):
    db["E9"] = "=1/0"
    assert run(db, '=DSUM(A5:E11,"Profit",A1:A2)') == DIV0  # Apple rows include E9
    db["A2"] = '="=Pear"'
    assert run(db, '=DSUM(A5:E11,"Profit",A1:A2)') == 96 + 76.8  # errors in other records are ignored
    assert run(db, '=DCOUNTA(A5:E11,"Profit",A5:A6)') == 3  # Apple rows (E9 counts)


def test_computed_criterion(db):
    # a formula under a label that isn't a field, written for the first record (row 6)
    db["G1"] = "Above average"
    db["G2"] = "=C6>AVERAGE($C$6:$C$11)"
    assert run(db, '=DCOUNT(A5:E11,"Age",G1:G2)') == 3  # ages 20, 14, 15 > 13
    db["G1"] = ""
    assert run(db, '=DSUM(A5:E11,"Age",G1:G2)') == 49
    db["G2"] = "=AND(A6=\"Apple\",B6>10)"
    assert run(db, '=DSUM(A5:E11,"Profit",G1:G2)') == 180


def test_recalculates(db):
    assert run(db, '=DSUM(A5:E11,"Profit",A1:A2)') == 225
    db["E11"] = 100
    assert db["H1"] == 280
    db["A2"] = "Pear"
    assert db["H1"] == 96 + 76.8
