import math

import pytest

from termse.formula.errors import DIV0, NA, NAME, NUM, REF, VALUE


@pytest.mark.parametrize("formula,expected", [
    ("=1+2*3", 7), ("=(1+2)*3", 9), ("=-2^2", 4), ("=2^-1", 0.5), ("=10/4", 2.5), ("=7/0", DIV0),
    ("=50%", 0.5), ('="3"+4', 7), ('="a"+1', VALUE), ("=TRUE+1", 2), ('="x"&1.5', "x1.5"),
    ("=0.1+0.2=0.3", True), ('="ABC"="abc"', True), ('="a"<"b"', True), ("=1<\"a\"", True),
    ("=ROUND(2.675,2)", 2.68), ("=ROUND(-2.5,0)", -3), ("=ROUND(1234.5,-2)", 1200), ("=ROUNDUP(1.21,1)", 1.3),
    ("=ROUNDDOWN(-1.29,1)", -1.2), ("=INT(-1.5)", -2), ("=TRUNC(-1.5)", -1), ("=MOD(-3,2)", 1), ("=MOD(3,-2)", -1),
    ("=ABS(-4)", 4), ("=SQRT(16)", 4), ("=SQRT(-1)", NUM), ("=POWER(2,10)", 1024), ("=CEILING(4.2,2)", 6),
    ("=FLOOR(-2.5,2)", -4), ("=CEILING(-2.5,2)", -2), ("=EVEN(3)", 4), ("=ODD(2)", 3), ("=LN(0)", NUM),
    ("=LOG(8,2)", 3), ("=FACT(5)", 120), ("=PRODUCT(2,3,4)", 24), ("=SIGN(-3)", -1), ("=MROUND(10,3)", 9),
    ("=QUOTIENT(7,2)", 3), ("=GCD(12,18)", 6), ("=LCM(4,6)", 12), ("=COMBIN(5,2)", 10),
])
def test_math(ev, formula, expected):
    v = ev(formula)
    if isinstance(expected, float):
        assert v == pytest.approx(expected)
    else:
        assert v == expected


def test_aggregates_ignore_text_in_ranges(ev):
    cells = dict(A1="1", A2="2", A3="hello", A4="TRUE", A5="'5")
    assert ev("=SUM(A1:A5)", **cells) == 3
    assert ev("=COUNT(A1:A5)") == 2
    assert ev("=COUNTA(A1:A5)") == 5
    assert ev("=AVERAGE(A1:A5)") == 1.5
    assert ev("=SUM(A1:A5,TRUE,\"3\")") == 7  # direct args are coerced
    assert ev("=MAX(A1:A5)") == 2
    assert ev("=MIN(B1:B5)") == 0
    assert ev("=COUNTBLANK(A1:A6)") == 1


def test_yearfrac_actual_actual(ev):
    # Within a year: 366 days if a 29 February falls in the period, else 365.
    assert ev("=YEARFRAC(DATE(2023,12,1),DATE(2024,3,1),1)") == pytest.approx(91 / 366)
    assert ev("=YEARFRAC(DATE(2023,3,1),DATE(2024,3,1),1)") == 1
    assert ev("=YEARFRAC(DATE(2024,3,1),DATE(2025,2,1),1)") == pytest.approx(337 / 365)
    assert ev("=YEARFRAC(DATE(2024,1,1),DATE(2024,2,1),1)") == pytest.approx(31 / 366)
    # Longer: the average length of the years involved
    assert ev("=YEARFRAC(DATE(2023,1,1),DATE(2024,6,1),1)") == pytest.approx(517 / 365.5)


def test_dollar_rounds_half_away_from_zero(ev):
    assert ev("=DOLLAR(0.125)") == "$0.13"
    assert ev("=DOLLAR(2.5,0)") == "$3"
    assert ev("=DOLLAR(1234.567,-2)") == "$1,200"
    assert ev("=DOLLAR(-1234.567,1)") == "($1,234.6)"


def test_huge_number_text(ev):
    assert ev('=VALUE("1e400")') == VALUE
    assert ev('="1e400"+1') == VALUE
    assert ev('=VALUE("1.5e2")') == 150


def test_error_propagation(ev):
    assert ev("=SUM(A1:A2)", A1="=1/0", A2="1") == DIV0
    assert ev("=IFERROR(A1,0)") == 0
    assert ev("=ISERROR(A1)") is True
    assert ev("=IFNA(NA(),\"x\")") == "x"
    assert ev("=ISNA(A1)") is False
    assert ev("=ERROR.TYPE(A1)") == 2


def test_criteria_functions(ev):
    cells = {f"A{i}": v for i, v in enumerate(["apple", "banana", "apricot", "cherry", "Apple"], 1)}
    cells.update({f"B{i}": v for i, v in enumerate([1, 2, 3, 4, 5], 1)})
    assert ev('=COUNTIF(A1:A5,"apple")', **cells) == 2
    assert ev('=COUNTIF(A1:A5,"ap*")') == 3
    assert ev('=COUNTIF(A1:A5,"?????")') == 2
    assert ev('=COUNTIF(B1:B5,">2")') == 3
    assert ev('=COUNTIF(B1:B5,"<>3")') == 4
    assert ev('=SUMIF(A1:A5,"ap*",B1:B5)') == 9
    assert ev('=SUMIFS(B1:B5,A1:A5,"a*",B1:B5,">1")') == 8
    assert ev('=AVERAGEIF(B1:B5,">=4")') == 4.5
    assert ev('=COUNTIFS(A1:A5,"*an*",B1:B5,2)') == 1
    assert ev('=MAXIFS(B1:B5,A1:A5,"a*")') == 5
    assert ev('=MINIFS(B1:B5,A1:A5,"c*")') == 4
    assert ev('=COUNTIF(C1:C5,"")') == 5


def test_criteria_functions_lift_over_array_criteria(ws):
    """Array or multi-cell criteria give one result per criteria element, like Excel."""
    for i, (region, amount) in enumerate([("East", 10), ("West", 5), ("East", 7), ("North", 3),
                                          ("West", 1)], 1):
        ws[f"A{i}"] = region
        ws[f"B{i}"] = amount
    ws["D1"] = '=SUMIFS(B1:B5,A1:A5,{"East";"West";"South"})'
    assert [ws[f"D{r}"] for r in (1, 2, 3)] == [17, 6, 0]
    ws["E1"] = "=SUMIF(A1:A5,UNIQUE(A1:A5),B1:B5)"
    assert [ws[f"E{r}"] for r in (1, 2, 3, 4)] == [17, 6, 3, None]
    ws["F1"] = "=COUNTIF(A1:A5,A1:A5)"  # a multi-cell criteria range spills too
    assert [ws[f"F{r}"] for r in range(1, 6)] == [2, 2, 2, 1, 2]
    ws["G1"] = '=COUNTIFS(A1:A5,"East",B1:B5,{">5",">8"})'  # row vector spills right
    assert (ws["G1"], ws["H1"]) == (2, 1)
    ws["G3"] = '=AVERAGEIFS(B1:B5,A1:A5,{"East","South"})'  # no match -> #DIV/0! for that element
    assert ws["G3"] == 8.5 and ws["H3"].code == "#DIV/0!"
    ws["G5"] = '=MAXIFS(B1:B5,A1:A5,{"East";"West"},B1:B5,{"<9","<9"})'  # broadcast to 2x2
    assert [[ws[f"{c}{r}"] for c in "GH"] for r in (5, 6)] == [[7, 7], [5, 5]]
    ws["J1"] = '=MINIFS(B1:B5,A1:A5,{"East";"West";"North"},B1:B5,{">1";">2"})'  # 3 vs 2 rows
    assert [ws[f"J{r}"] for r in (1, 2, 3)] == [7, 5, NA]
    ws["K1"] = '=SUM(COUNTIF(A1:A5,{"East","North"}))'
    assert ws["K1"] == 3 and ws["L1"] is None
    ws["K2"] = '=AVERAGEIF(A1:A5,{"West"},B1:B5)'  # 1x1 array is just a scalar
    assert ws["K2"] == 3


def test_lookups(ev):
    cells = {"A1": "id", "B1": "name", "A2": "10", "B2": "ten", "A3": "20", "B3": "twenty", "A4": "30",
             "B4": "thirty"}
    assert ev("=VLOOKUP(20,A2:B4,2,FALSE)", **cells) == "twenty"
    assert ev("=VLOOKUP(25,A2:B4,2)") == "twenty"   # approximate
    assert ev("=VLOOKUP(5,A2:B4,2)") == NA
    assert ev("=VLOOKUP(20,A2:B4,3,FALSE)") == REF
    assert ev('=VLOOKUP("tw*",B2:B4,1,FALSE)') == "twenty"
    assert ev("=HLOOKUP(\"name\",A1:B4,3,FALSE)") == "twenty"
    assert ev("=MATCH(30,A2:A4,0)") == 3
    assert ev("=MATCH(25,A2:A4)") == 2
    # An empty argument is FALSE / 0 (exact match), unlike an omitted one
    assert ev("=VLOOKUP(25,A2:B4,2,)") == NA
    assert ev("=VLOOKUP(20,A2:B4,2,)") == "twenty"
    assert ev('=HLOOKUP("nam",A1:B4,2,)') == NA
    assert ev("=MATCH(25,A2:A4,)") == NA
    assert ev("=MATCH(20,A2:A4,)") == 2
    assert ev("=INDEX(B2:B4,3)") == "thirty"
    assert ev("=INDEX(A1:B4,2,2)") == "ten"
    assert ev("=SUM(INDEX(A2:B4,0,1))") == 60
    assert ev("=XLOOKUP(20,A2:A4,B2:B4)") == "twenty"
    assert ev('=XLOOKUP(99,A2:A4,B2:B4,"none")') == "none"
    assert ev("=XLOOKUP(25,A2:A4,B2:B4,,-1)") == "twenty"
    assert ev("=XLOOKUP(25,A2:A4,B2:B4,,1)") == "thirty"
    assert ev("=XMATCH(30,A2:A4)") == 3
    assert ev("=LOOKUP(22,A2:A4,B2:B4)") == "twenty"
    assert ev("=SUM(A2:INDEX(A2:A4,2))") == 30
    assert ev("=SUM(OFFSET(A2,1,0,2,1))") == 50
    assert ev('=INDIRECT("B"&3)') == "twenty"
    assert ev("=ROWS(A1:B4)*COLUMNS(A1:B4)") == 8
    assert ev("=ROW(C7)+COLUMN(C7)") == 10
    assert ev("=CHOOSE(2,\"a\",\"b\",\"c\")") == "b"
    assert ev("=ADDRESS(3,2)") == "$B$3"


def test_logical(ev):
    assert ev("=IF(1>2,\"y\",\"n\")") == "n"
    assert ev("=IF(TRUE,,1)") == 0
    assert ev("=IF(FALSE,1)") is False
    assert ev("=AND(TRUE,1,2>1)") is True
    assert ev("=OR(FALSE,0)") is False
    assert ev("=XOR(TRUE,TRUE)") is False
    assert ev("=NOT(0)") is True
    assert ev("=IFS(1>2,\"a\",2>1,\"b\")") == "b"
    assert ev("=SWITCH(2,1,\"one\",2,\"two\",\"other\")") == "two"
    assert ev("=SWITCH(9,1,\"one\",\"other\")") == "other"
    assert ev("=IF(1,1,1/0)") == 1  # lazy: unused branch errors don't matter


def test_text(ev):
    assert ev('=LEFT("hello",2)&RIGHT("hello",2)&MID("hello",2,3)') == "heloell"
    assert ev('=LEN("héllo")') == 5
    assert ev('=UPPER("ab")&LOWER("CD")&PROPER("hello wORLD")') == "ABcdHello World"
    assert ev('=TRIM("  a   b  ")') == "a b"
    assert ev('=SUBSTITUTE("a-b-c","-","+")') == "a+b+c"
    assert ev('=SUBSTITUTE("a-b-c","-","+",2)') == "a-b+c"
    assert ev('=REPLACE("abcdef",2,3,"X")') == "aXef"
    assert ev('=FIND("b","abcb")') == 2
    assert ev('=FIND("B","abcb")') == VALUE
    assert ev('=SEARCH("B","abcb")') == 2
    assert ev('=SEARCH("c?","abcb")') == 3
    assert ev('=TEXT(1234.567,"#,##0.00")') == "1,234.57"
    assert ev('=TEXT(0.25,"0%")') == "25%"
    assert ev('=TEXT(46347,"yyyy-mm-dd")') == "2026-11-21"
    assert ev('=VALUE("1,234.5")') == 1234.5
    assert ev('=VALUE("abc")') == VALUE
    assert ev('=CONCAT("a",1,TRUE)') == "a1TRUE"
    assert ev('=TEXTJOIN(", ",TRUE,A1:A3)', A1="x", A3="z") == "x, z"
    assert ev('=REPT("ab",3)') == "ababab"
    assert ev('=EXACT("a","A")') is False
    assert ev('=CHAR(65)&CODE("A")') == "A65"
    assert ev('="Total: "&1/3') == "Total: 0.333333333333333"


def test_dates(ev):
    assert ev("=DATE(2026,11,21)") == 46347
    assert ev("=DATE(2026,13,1)") == ev("=DATE(2027,1,1)")
    assert ev("=DATE(2026,3,0)") == ev("=DATE(2026,2,28)")
    assert ev("=YEAR(46347)*10000+MONTH(46347)*100+DAY(46347)") == 20261121
    assert ev("=WEEKDAY(46347)") == 7  # Saturday
    assert ev("=WEEKDAY(46347,2)") == 6
    assert ev("=EDATE(DATE(2026,1,31),1)") == ev("=DATE(2026,2,28)")
    assert ev("=EOMONTH(DATE(2026,2,10),0)") == ev("=DATE(2026,2,28)")
    assert ev('=DATEDIF(DATE(2020,5,15),DATE(2026,11,21),"Y")') == 6
    assert ev('=DATEDIF(DATE(2020,5,15),DATE(2026,11,21),"YM")') == 6
    assert ev('=DATEDIF(DATE(2020,5,15),DATE(2026,11,21),"MD")') == 6
    assert ev("=NETWORKDAYS(DATE(2026,11,16),DATE(2026,11,22))") == 5
    assert ev("=WORKDAY(DATE(2026,11,20),1)") == ev("=DATE(2026,11,23)")
    assert ev("=HOUR(0.75)*100+MINUTE(0.75)") == 1800
    assert ev("=TIME(6,30,0)") == pytest.approx(0.2708333333)
    assert ev('=DATEVALUE("2026-11-21")') == 46347
    assert ev('="2026-11-21"+1') == 46348
    assert ev("=DAYS(DATE(2026,12,25),DATE(2026,11,21))") == 34
    assert ev("=ISOWEEKNUM(DATE(2026,1,1))") == 1


def test_stats(ev):
    cells = {f"A{i}": v for i, v in enumerate([2, 4, 4, 4, 5, 5, 7, 9], 1)}
    assert ev("=MEDIAN(A1:A8)", **cells) == 4.5
    assert ev("=MODE(A1:A8)") == 4
    assert ev("=STDEV.P(A1:A8)") == 2
    assert ev("=VAR(A1:A8)") == pytest.approx(4.571428571)
    assert ev("=LARGE(A1:A8,2)") == 7
    assert ev("=SMALL(A1:A8,3)") == 4
    assert ev("=RANK(5,A1:A8)") == 3
    assert ev("=RANK(5,A1:A8,1)") == 5
    assert ev("=PERCENTILE(A1:A8,0.5)") == 4.5
    assert ev("=QUARTILE(A1:A8,1)") == 4


def test_binom_dist(ev):
    assert ev("=BINOM.DIST(3,10,0.5,FALSE)") == pytest.approx(120 / 1024)
    assert ev("=BINOM.DIST(3,10,0.5,TRUE)") == pytest.approx(176 / 1024)
    assert ev("=BINOMDIST(3.9,10,0.5,0)") == pytest.approx(120 / 1024)  # number_s truncated
    assert ev("=BINOM.DIST(10,10,0.3,TRUE)") == 1
    assert ev("=BINOM.DIST(0,5,0,FALSE)") == 1
    assert ev("=BINOM.DIST(5,5,1,FALSE)") == 1
    assert ev("=BINOM.DIST(500,2000,0.25,FALSE)") == pytest.approx(0.02059757143332068, rel=1e-12)
    for bad in ("=BINOM.DIST(11,10,0.5,FALSE)", "=BINOM.DIST(-1,10,0.5,FALSE)", "=BINOM.DIST(1,10,1.5,FALSE)"):
        assert str(ev(bad)) == "#NUM!"


def test_binom_dist_storage_prefix():
    from termse.formula.parser import to_display, to_storage
    assert to_storage("=BINOM.DIST(1,2,0.5,FALSE)") == "=_xlfn.BINOM.DIST(1,2,0.5,FALSE)"
    assert to_storage("=BINOMDIST(1,2,0.5,FALSE)") == "=BINOMDIST(1,2,0.5,FALSE)"
    assert to_display("=_xlfn.BINOM.DIST(1,2,0.5,FALSE)") == "=BINOM.DIST(1,2,0.5,FALSE)"


def test_arrays(ev):
    cells = {"A1": "1", "A2": "2", "A3": "3", "B1": "4", "B2": "5", "B3": "6"}
    assert ev("=SUMPRODUCT(A1:A3,B1:B3)", **cells) == 32
    assert ev("=SUMPRODUCT((A1:A3>1)*(B1:B3))") == 11
    assert ev("=SUMPRODUCT(A1:A3,B1:B2)") == VALUE  # different sizes
    assert ev("=SUMPRODUCT(A1:A3,{1;2})") == VALUE
    assert ev("=SUMPRODUCT(A1:A4,B1:B4)") == 32
    assert ev("=SUMPRODUCT(A:A,B:B)") == 32
    assert ev("=SUM(A1:A3*B1:B3)") == 32
    assert ev("=SUM({1,2;3,4})") == 10
    assert ev("=SUM(FILTER(B1:B3,A1:A3>1))") == 11
    assert ev("=COUNTA(UNIQUE({1;1;2}))") == 2
    assert ev("=SUM(SEQUENCE(4))") == 10
    assert ev("=SUM(IF(A1:A3>1,B1:B3,0))") == 11


def test_unknown_function_is_name_error(ev):
    assert ev("=NOSUCHFUNC(1)") == NAME


def test_math_constants(ev):
    assert ev("=PI()") == pytest.approx(math.pi)
    v = ev("=RAND()")
    assert 0 <= v < 1
