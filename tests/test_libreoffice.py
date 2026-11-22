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

from termse.formula.errors import ExcelError
from termse.formula.parser import to_storage
from termse.io import xlsx

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
    '=COUNTIFS(E2:E7,"fruit",B2:B7,">2")', '=AVERAGEIF(E2:E7,"veg",C2:C7)',
    '=SUMPRODUCT(SUMIF(E2:E7,{"fruit","veg"},B2:B7)*{1,10})',
    '=SUMPRODUCT(COUNTIFS(E2:E7,{"fruit";"veg"},B2:B7,{">2",">5"})*{1,10;100,1000})', "=COUNTA(A1:E7)", "=COUNTBLANK(A1:F8)",
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
    # statistical distributions and tests
    "=NORM.DIST(1.3,0.5,2,TRUE)", "=NORM.DIST(1.3,0.5,2,FALSE)", "=NORM.INV(0.975,10,3)", "=NORM.S.INV(0.05)",
    "=NORMSDIST(-4)", "=LOGNORM.DIST(4,3.5,1.2,TRUE)", "=LOGINV(0.5,1,2)", "=EXPON.DIST(0.2,10,FALSE)",
    "=WEIBULL(2,1.5,3,0)", "=GAMMA(-0.75)", "=GAMMALN.PRECISE(0.3)", "=GAMMADIST(3,0.5,2,TRUE)",
    "=GAMMAINV(0.99,2.5,1)", "=BETA.DIST(0.3,2,3,TRUE)", "=BETAINV(0.2,4,5)", "=CHISQ.DIST(2,3,FALSE)",
    "=CHIDIST(3.5,2)", "=CHISQ.INV(0.93,1)", "=CHIINV(0.5,4)", "=CHISQ.INV.RT(1E-12,3)", "=T.DIST(-1.2,7,TRUE)",
    "=TDIST(1.96,60,1)", "=T.INV(0.01,15)", "=TINV(0.05,10)", "=T.INV(0.3,5.9)", "=F.DIST(2,3,30,FALSE)",
    "=FDIST(2,3,30)", "=F.INV.RT(0.01,6,4)", "=FINV(0.5,10,2)", "=POISSON(250,240,1)", "=HYPGEOMDIST(2,10,5,30)",
    "=NEGBINOMDIST(3,2,0.6)", "=CRITBINOM(100,0.3,0.5)", "=BINOM.INV(5000,0.01,0.9)",
    "=BINOM.DIST(2000,5000,0.4,TRUE)", "=CONFIDENCE(0.1,1,10)", "=CONFIDENCE.T(0.05,1,50)",
    "=T.TEST(B2:B7,C2:C7,2,1)", "=T.TEST(B2:B7,C2:C7,1,2)", "=TTEST(B2:B7,C2:C7,2,3)", "=F.TEST(B2:B7,C2:C7)",
    "=Z.TEST(C2:C7,2)", "=ZTEST(B2:B7,4,3)", "=CHISQ.TEST(B2:B4,C2:C4)",
    # descriptive statistics and regression
    "=AVEDEV(C2:C7)", "=DEVSQ(B2:B7)", "=SKEW(C2:C7)", "=SKEW.P(B2:B7)", "=KURT(C2:C7)", "=HARMEAN(C2:C7)",
    "=TRIMMEAN(B2:B7,0.4)", "=FISHER(0.3)", "=FISHERINV(-0.4)", "=PERMUTATIONA(4,3)", "=STANDARDIZE(5,3,2)",
    "=STDEVA(A1:C7)", "=VARPA(A1:C7)", "=PERCENTRANK.INC(B2:B7,5)", "=PERCENTRANK(C2:C7,1)",
    "=PERCENTRANK.EXC(B2:B7,3,4)", "=INDEX(MODE.MULT(1,2,3,4,3,2,1,2,3),2)",
    "=INDEX(FREQUENCY(C2:C7,{1,3}),2)", "=PROB(B2:B4,{0.2;0.5;0.3},3,12)", "=COVAR(B2:B7,C2:C7)",
    "=COVARIANCE.S(B2:B7,C2:C7)", "=RSQ(C2:C7,B2:B7)", "=STEYX(C2:C7,B2:B7)", "=FORECAST(10,C2:C7,B2:B7)",
    "=INDEX(LINEST(C2:C7,B2:B7,TRUE,TRUE),2,1)", "=INDEX(LINEST(C2:C7,B2:B7,TRUE,TRUE),4,1)",
    "=INDEX(LINEST(D2:D7,B2:C7,FALSE,TRUE),5,2)", "=INDEX(LOGEST(C2:C7,B2:B7,TRUE,TRUE),1,1)",
    "=INDEX(TREND(C2:C7,B2:B7,{15;20}),2)", "=INDEX(GROWTH(C2:C7,B2:B7,{15;20}),1)",
    # dynamic arrays and lookup
    "=INDEX(SORTBY(A2:A7,B2:B7,-1),2)", "=INDEX(SORTBY(A2:C7,E2:E7,1,C2:C7,-1),4,1)", "=INDEX(SORTBY(A1:E1,{5,4,3,2,1}),1,2)",
    "=INDEX(TAKE(A2:C7,-2,2),1,1)", "=ROWS(TAKE(A2:C7,10))", "=INDEX(TAKE(A2:C7,,-1),3)", "=INDEX(DROP(A2:C7,2,1),1,2)",
    "=ROWS(DROP(A2:C7,-4))", "=DROP(A2:C7,6)", "=TAKE(A2:C7,0)", "=INDEX(VSTACK(A2:B3,C2:C3),3,2)",
    "=INDEX(VSTACK(A2:B3,C2:C3),4,1)", "=INDEX(HSTACK(A2:A4,B2:C3),3,3)", "=COLUMNS(HSTACK(A2:A4,B2:C3,1))",
    "=INDEX(TOCOL(A2:C3),4)", "=INDEX(TOCOL(A2:C3,0,TRUE),2)", "=COUNTA(TOCOL(A1:G8,1))", "=INDEX(TOROW(B2:C4,,TRUE),1,4)",
    "=INDEX(CHOOSEROWS(A2:C7,-1,2),1,1)", "=INDEX(CHOOSECOLS(A2:C7,3,-3),2,2)", "=CHOOSEROWS(A2:C7,7)", "=CHOOSECOLS(A2:C7,0)",
    '=INDEX(EXPAND(B2:C3,3,3,"-"),3,3)', "=INDEX(EXPAND(B2:C3,3),3,1)", "=EXPAND(B2:C3,1)", "=INDEX(WRAPROWS(B2:B7,4),2,2)",
    "=INDEX(WRAPROWS(B2:B7,4),2,3)", '=INDEX(WRAPCOLS(A2:A7,4,"x"),3,2)', "=WRAPROWS(A2:C7,2)", "=WRAPROWS(B2:B7,0)",
    "=SUM(WRAPCOLS(C2:C7,5,0))", "=FORMULATEXT(D3)", "=FORMULATEXT(A3)", '=HYPERLINK("https://example.com","Example")',
    '=HYPERLINK("https://example.com")', "=AREAS(B2:D4)",
    # database functions (criteria in J1:M3)
    '=DSUM(A1:E7,"Qty",J1:L3)', '=DCOUNT(A1:E7,"Price",J1:K2)', '=DCOUNTA(A1:E7,"Name",J1:L3)',
    '=DCOUNT(A1:E7,,J1:L3)', '=DAVERAGE(A1:E7,3,J1:L3)', '=DMAX(A1:E7,"Price",J1:J2)', '=DMIN(A1:E7,"price",J1:K2)',
    '=DPRODUCT(A1:E7,2,J1:K2)', '=DSTDEV(A1:E7,"Qty",J1:L3)', '=DSTDEVP(A1:E7,"Qty",J1:L3)',
    '=DVAR(A1:E7,"Price",J1:J2)', '=DVARP(A1:E7,"Price",J1:J2)', '=DGET(A1:E7,"Name",M1:M2)',
    '=DGET(A1:E7,"Name",J1:L3)', '=DGET(A1:E7,"Name",K1:K2)', '=DSUM(A1:E7,"Nope",J1:L3)', '=DSUM(A1:E7,9,J1:L3)',
    '=DSUM(A1:E7,"Qty",L1:L2)', '=DAVERAGE(A1:E7,"Qty",M1:M3)',
    # LET (LibreOffice doesn't have LAMBDA)
    "=LET(x,B2,y,x*2,x+y)", "=LET(r,B2:B7,SUM(r)/ROWS(r))", "=LET(x,1/0,5)",
    # engineering (LibreOffice doesn't read IMTAN, IMSEC... with the _xlfn. prefix Excel uses)
    "=BIN2DEC(1111111111)", "=BIN2HEX(11111011,4)", "=BIN2OCT(1100100)", "=DEC2BIN(-100)", "=DEC2BIN(9.9,6)",
    "=DEC2HEX(-54)", "=DEC2OCT(B3,4)", '=HEX2BIN("FFFFFFFE00")', '=HEX2DEC("3DA408B9")', '=HEX2OCT("FFFFFFFF00")',
    "=OCT2BIN(7777777000)", "=OCT2DEC(7777777533)", "=OCT2HEX(100,4)", "=DEC2BIN(512)", '=HEX2DEC("G")',
    "=BITAND(13,25)", "=BITOR(23,10)", "=BITXOR(B3,B4)", "=BITLSHIFT(4,-2)", "=BITRSHIFT(5,-2)", "=BITLSHIFT(1,48)",
    "=DELTA(B2,3)", "=GESTEP(-4,-5)", "=ERF(0.745)", "=ERF(0.2,1.5)", "=ERF.PRECISE(-0.5)", "=ERFC(1)",
    "=ERFC.PRECISE(2)", '=CONVERT(68,"F","C")', '=CONVERT(B3,"km","m")', '=CONVERT(1,"Gibyte","Mibyte")',
    '=CONVERT(1,"mm2","cm2")', '=CONVERT(2,"mL","l")', '=CONVERT(500,"Rank","Reau")', '=CONVERT(1,"kft","ft")',
    '=CONVERT(6,"gal","l")', '=CONVERT(1,"m","sec")', "=BESSELJ(1.9,2)", "=BESSELJ(-3,1)", "=BESSELJ(31,40)",
    "=BESSELY(2.5,1)", "=BESSELY(0.5,3)", "=BESSELY(31,2)", "=BESSELI(1.5,1)", "=BESSELI(20,3)",
    "=BESSELK(1.5,1)", "=BESSELK(0.01,0)", "=BESSELK(50,3)", "=COMPLEX(3,-4)", '=COMPLEX(0,-1,"j")',
    '=IMABS("5+12i")', '=IMAGINARY("0-j")', '=IMREAL("1e3+2e-1i")', '=IMARGUMENT("3+4i")',
    '=IMCONJUGATE("3+4i")', '=IMCOS("1+i")', '=IMSIN("4+3i")', '=IMEXP("1+i")', '=IMLN("3+4i")',
    '=IMLOG10("3+4i")', '=IMLOG2("3+4i")', '=IMPOWER("2+3i",3)', '=IMSQRT("1+i")', '=IMDIV("-238+240i","10+24i")',
    '=IMSUB("13+4i","5+3i")', '=IMSUM("3+4i","5-3i","i")', '=IMPRODUCT("3+4i","5-3i")', '=IMREAL("3+4")',
]

# Criteria for the database functions, in J1:M3 (text criteria use ="=x" because plain text is a
# prefix match in Excel but a whole-cell match in LibreOffice).
CRITERIA = [
    ["Cat", "Qty", "Price", "Qty"],
    ['="=fruit"', ">2", None, ">10"],
    ['="=veg"', None, "<1", None],
]

# Where LibreOffice and Excel disagree, we follow Excel (Microsoft 365).
EXCEL_DIFFERS = {
    # DGET with several matching records: Excel documents #NUM!
    '=DGET(A1:E7,"Name",J1:L3)': "#NUM!", '=DGET(A1:E7,"Name",K1:K2)': "#NUM!",
    # a blank criteria row matches every record in Excel; LibreOffice skips it
    '=DAVERAGE(A1:E7,"Qty",M1:M3)': 5.5,
    "=SUM((B2:B7>2)*C2:C7)": 5.65,        # dynamic-array evaluation; LibreOffice: #VALUE!
    "=SQRT(-1)": "#NUM!",                  # LibreOffice: #VALUE!
    "=DOLLAR(-1234.5)": "($1,234.50)",     # en-US currency format; LibreOffice: -$1,234.50
    "=PERCENTRANK(C2:C7,1)": 0.288,        # Excel truncates to 3 digits; LibreOffice rounds: 0.289
    "=PERCENTRANK.EXC(B2:B7,3,4)": 0.4285,  # likewise; LibreOffice: 0.4286
    "=DROP(A2:C7,6)": "#CALC!",            # empty array; LibreOffice: #VALUE!
    "=TAKE(A2:C7,0)": "#CALC!",            # empty array; LibreOffice: #N/A
    "=WRAPROWS(B2:B7,0)": "#NUM!",         # wrap_count < 1; LibreOffice: #VALUE!
    # Microsoft documents #NUM! for invalid numbers and #N/A for unknown or mismatched units;
    # LibreOffice gives #VALUE!
    "=DEC2BIN(512)": "#NUM!", '=HEX2DEC("G")': "#NUM!", '=IMREAL("3+4")': "#NUM!",
    '=CONVERT(1,"kft","ft")': "#N/A", '=CONVERT(1,"m","sec")': "#N/A",
    # Microsoft's example; Excel's BESSELJ is approximate, LibreOffice's exact (0.329925727692387)
    "=BESSELJ(1.9,2)": 0.3299258286697852,
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
    for r, row in enumerate(CRITERIA, 1):
        for c, v in enumerate(row, 10):
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
            from termse.model.dates import to_serial

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
    """A workbook edited and saved by tse opens in LibreOffice with the right values."""
    from termse.model import structure
    from termse.model.refs import parse_addr

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
