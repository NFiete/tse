"""Cross-check our formula engine against LibreOffice's.

A workbook full of formulas is written with openpyxl, recalculated and re-saved by
LibreOffice headless, and LibreOffice's cached results are compared with ours.
Skipped when soffice isn't installed.
"""
import math
import os
import shutil
import subprocess

import openpyxl
import pytest

from termxl.formula.errors import ExcelError
from termxl.formula.parser import to_storage
from termxl.io import xlsx

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")

DATA = [
    ["Name", "Qty", "Price", "Date", "Cat"],
    ["apple", 3, 1.25, "=DATE(2026,1,15)", "fruit"],
    ["banana", 12, 0.5, "=DATE(2026,2,28)", "fruit"],
    ["carrot", 7, 0.8, "=DATE(2025,12,31)", "veg"],
    ["dates", 0, 4.75, "=DATE(2026,11,21)", "fruit"],
    ["eggplant", 2, 2.2, "=DATE(2024,2,29)", "veg"],
    ["fig", 9, 3.1, "=DATE(2026,6,30)", "fruit"],
]

FORMULAS = [
    "=SUM(B2:B7)", "=AVERAGE(C2:C7)", "=SUMPRODUCT(B2:B7,C2:C7)", "=MAX(B2:B7)-MIN(C2:C7)",
    '=SUMIF(E2:E7,"fruit",B2:B7)', '=SUMIFS(B2:B7,E2:E7,"veg",C2:C7,">1")', '=COUNTIF(A2:A7,"*an*")',
    '=COUNTIFS(E2:E7,"fruit",B2:B7,">2")', '=AVERAGEIF(E2:E7,"veg",C2:C7)', "=COUNTA(A1:E7)", "=COUNTBLANK(A1:F8)",
    '=VLOOKUP("carrot",A2:C7,3,FALSE)', "=INDEX(A2:A7,MATCH(MAX(B2:B7),B2:B7,0))", "=MATCH(7,B2:B7,0)",
    '=HLOOKUP("Price",A1:E7,4,FALSE)', "=ROUND(AVERAGE(C2:C7),3)", "=ROUNDUP(2.341,1)", "=ROUNDDOWN(-2.349,2)",
    "=MOD(-7,3)", "=INT(-3.2)", "=TRUNC(-3.7)", "=CEILING(7.3,2)", "=FLOOR(7.3,2)", "=MROUND(17,5)",
    "=POWER(2,0.5)", "=SQRT(2)^2", "=LN(EXP(2))", "=LOG(1000)", "=ABS(-2.5)*SIGN(-1)", "=FACT(10)",
    "=COMBIN(10,3)", "=GCD(48,36)", "=LCM(4,6,10)", "=PRODUCT(B2:B4)", "=SUMSQ(1,2,3)", "=QUOTIENT(-7,2)",
    "=EVEN(-3)", "=ODD(4.1)", "=PI()*2", "=DEGREES(PI())", "=ATAN2(1,1)", "=SIN(1)^2+COS(1)^2",
    "=MEDIAN(B2:B7)", "=STDEV(B2:B7)", "=STDEVP(C2:C7)", "=VAR(B2:B7)", "=VARP(C2:C7)", "=LARGE(C2:C7,2)",
    "=SMALL(B2:B7,2)", "=RANK(7,B2:B7)", "=PERCENTILE(C2:C7,0.3)", "=QUARTILE(B2:B7,3)", "=MODE(1,2,2,3)",
    "=CORREL(B2:B7,C2:C7)", "=SLOPE(C2:C7,B2:B7)", "=INTERCEPT(C2:C7,B2:B7)", "=GEOMEAN(C2:C7)",
    "=BINOM.DIST(3,10,0.35,FALSE)", "=BINOM.DIST(4,12,0.2,TRUE)", "=BINOMDIST(B4,B3,0.4,1)",
    '=IF(B2>2,"big","small")', '=IFERROR(1/0,"err")', "=AND(B2>1,C2<2)", "=OR(B2>100,C2<2)", "=NOT(FALSE)",
    '=CHOOSE(2,"a","b","c")', '=LEFT(A3,3)&RIGHT(A3,2)', '=MID("spreadsheet",3,4)', '=LEN(A6)', '=UPPER(A2)',
    '=PROPER("hello wORLD")', '=TRIM("  a  b  ")', '=SUBSTITUTE("a.b.c",".","-")', '=SUBSTITUTE("a.b.c",".","-",2)',
    '=REPLACE("abcdef",3,2,"XY")', '=FIND("e","eggplant",2)', '=SEARCH("P","eggplant")', '=REPT("-",5)',
    '=CONCATENATE(A2," x",B2)', '=A2&" costs "&C2', '=TEXT(1234.5678,"#,##0.00")', '=TEXT(0.075,"0.0%")',
    '=TEXT(D2,"yyyy-mm-dd")', '=TEXT(D3,"mmm d, yyyy")', '=TEXT(D4,"dddd")', '=VALUE("3.5")+1', '=EXACT("A","a")',
    "=YEAR(D5)", "=MONTH(D6)", "=DAY(D3)", "=WEEKDAY(D5)", "=WEEKDAY(D5,2)", "=EDATE(D2,1)", "=EOMONTH(D2,1)",
    "=EOMONTH(D6,-12)", '=DATEDIF(D6,D5,"Y")', '=DATEDIF(D6,D5,"M")', '=DATEDIF(D6,D5,"D")', '=DATEDIF(D6,D5,"MD")',
    '=DATEDIF(D6,D5,"YM")', '=DATEDIF(D6,D5,"YD")', "=NETWORKDAYS(D4,D2)", "=WORKDAY(D4,10)", "=D5-D2",
    "=DATE(2026,14,35)", "=TIME(25,61,0)", "=HOUR(0.7)", "=MINUTE(0.7)", '=DATEVALUE("2026-03-04")',
    "=WEEKNUM(D5)", "=_xlfn.ISOWEEKNUM(D2)", "=YEARFRAC(D2,D5,1)", "=YEARFRAC(D2,D5)",
    "=ISBLANK(F2)", "=ISNUMBER(B2)", "=ISTEXT(A2)", '=ISERROR(1/0)', '=ISNA(MATCH(99,B2:B7,0))',
    "=N(TRUE)", '=T("x")', "=ROWS(A1:E7)*COLUMNS(A1:E7)", "=ROW(C5)+COLUMN(C5)", "=SUM(OFFSET(B2,1,0,3,1))",
    '=SUM(INDIRECT("B2:B"&4))', "=ADDRESS(3,4,2)", "=1/3+1/3+1/3=1", "=0.1+0.2", "=2^-2", "=-3^2",
    '="10"*2', "=TRUE+TRUE", '=SUMPRODUCT((E2:E7="fruit")*(B2:B7>2))', "=SUM((B2:B7>2)*C2:C7)",
    "=LOOKUP(5,B2:B7,A2:A7)", "=SUM(B2:INDEX(B2:B7,3))", '=COUNTIF(B2:B7,"<>3")', '=COUNTIF(A2:A7,"<c")',
    "=MAXA(B2:B7)", '=AVERAGEA(B2:B7)', "=SUBTOTAL(9,B2:B7)", "=SUBTOTAL(1,C2:C7)",
    '=FIXED(1234.567,1)', '=DOLLAR(-1234.5)', '=CHAR(65)&CODE("a")', "=CLEAN(\"a\"&CHAR(7)&\"b\")",
    "=1E+300*1E+300", "=SQRT(-1)", '=VALUE("abc")', "=NA()", "=1/0",
]

# Where LibreOffice and Excel disagree, we follow Excel (Microsoft 365).
EXCEL_DIFFERS = {
    "=SUM((B2:B7>2)*C2:C7)": 5.65,        # dynamic-array evaluation; LibreOffice: #VALUE!
    "=SQRT(-1)": "#NUM!",                  # LibreOffice: #VALUE!
    "=DOLLAR(-1234.5)": "($1,234.50)",     # en-US currency format; LibreOffice: -$1,234.50
}


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_against_libreoffice(tmp_path):
    src = tmp_path / "corpus.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "T"
    for r, row in enumerate(DATA, 1):
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)
    for i, f in enumerate(FORMULAS, 1):
        ws.cell(row=i, column=8, value=to_storage(f))
    wb.save(src)

    outdir = tmp_path / "lo"
    outdir.mkdir()
    env = dict(os.environ, HOME=str(tmp_path))  # isolated LibreOffice profile
    subprocess.run([SOFFICE, "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", str(outdir), str(src)],
                   check=True, capture_output=True, timeout=180, env=env)
    lo = openpyxl.load_workbook(outdir / "corpus.xlsx", data_only=True)["T"]

    book = xlsx.load(str(src))
    sheet = book.sheets[0]
    mismatches = []
    for i, f in enumerate(FORMULAS, 1):
        ours = book.get_value(sheet, i, 8)
        theirs = lo.cell(row=i, column=8).value
        if isinstance(ours, ExcelError):
            ours = ours.code
        if hasattr(theirs, "hour") or hasattr(theirs, "toordinal"):  # LO marks some results as dates/times
            from termxl.model.dates import to_serial

            theirs = to_serial(theirs)
        if f in EXCEL_DIFFERS:
            theirs = EXCEL_DIFFERS[f]
        if not _same(ours, theirs):
            mismatches.append(f"{f}: ours={ours!r} libreoffice={theirs!r}")
    assert not mismatches, "\n".join(mismatches)


def _same(a, b) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b or (isinstance(b, (int, float)) and not isinstance(b, bool) and bool(b) == a and b in (0, 1))
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-9)
    return a == b


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_saved_file_opens_in_libreoffice(sample_path, tmp_path):
    """A workbook edited and saved by term_xl opens in LibreOffice with the right values."""
    from termxl.model import structure
    from termxl.model.refs import parse_addr

    book = xlsx.load(sample_path)
    budget = book.sheet_by_name("Budget")
    book.set_input(budget, *parse_addr("B4"), "1600")
    structure.insert(book, budget, "row", 6, 1)
    book.set_input(budget, *parse_addr("A6"), "Insurance")
    book.set_input(budget, *parse_addr("B6"), "99")
    out = tmp_path / "edited.xlsx"
    xlsx.save(book, str(out))

    outdir = tmp_path / "csv"
    outdir.mkdir()
    env = dict(os.environ, HOME=str(tmp_path))
    subprocess.run([SOFFICE, "--headless", "--norestore", "--convert-to", "csv", "--outdir", str(outdir), str(out)],
                   check=True, capture_output=True, timeout=180, env=env)
    rows = (outdir / "edited.csv").read_text().splitlines()
    total_row = next(r for r in rows if r.startswith("Total"))
    # Jan total: 1600 + 420.5 + 99 + 180 + 95 + 200
    assert total_row.split(",")[1] == "2594.5"
    assert any(r.startswith("Insurance") for r in rows)
