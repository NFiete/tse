"""Statistical functions: distributions, hypothesis tests, descriptive statistics, regression."""
import pytest

from termxl.formula.errors import DIV0, NA, NUM, REF, VALUE

# Worked examples from Microsoft's documentation for each function.
EXCEL_EXAMPLES = [
    ("=NORM.DIST(42,40,1.5,TRUE)", 0.9087888), ("=NORM.DIST(42,40,1.5,FALSE)", 0.10934005),
    ("=NORMDIST(42,40,1.5,TRUE)", 0.9087888), ("=NORM.INV(0.908789,40,1.5)", 42.000002),
    ("=NORM.S.DIST(1.333333,TRUE)", 0.908788726), ("=NORM.S.DIST(1.333333,FALSE)", 0.164010148),
    ("=NORMSDIST(1.333333)", 0.908788726), ("=NORM.S.INV(0.908789)", 1.3333347), ("=STANDARDIZE(42,40,1.5)", 1.333333),
    ("=PHI(0.75)", 0.301137432), ("=GAUSS(2)", 0.47724987),
    ("=LOGNORM.DIST(4,3.5,1.2,TRUE)", 0.0390836), ("=LOGNORM.DIST(4,3.5,1.2,FALSE)", 0.0176176),
    ("=LOGNORM.INV(0.039084,3.5,1.2)", 4.0000252), ("=EXPON.DIST(0.2,10,TRUE)", 0.86466472),
    ("=EXPON.DIST(0.2,10,FALSE)", 1.35335283), ("=WEIBULL.DIST(105,20,100,TRUE)", 0.929581),
    ("=WEIBULL.DIST(105,20,100,FALSE)", 0.035589), ("=GAMMA(2.5)", 1.32934039), ("=GAMMA(-0.75)", -4.83414614),
    ("=GAMMALN(4)", 1.7917595), ("=GAMMA.DIST(10.00001131,9,2,FALSE)", 0.032639),
    ("=GAMMA.DIST(10.00001131,9,2,TRUE)", 0.068094), ("=GAMMA.INV(0.068094,9,2)", 10.0000112),
    ("=BETA.DIST(2,8,10,TRUE,1,3)", 0.6854706), ("=BETA.DIST(2,8,10,FALSE,1,3)", 1.4837646),
    ("=BETADIST(2,8,10,1,3)", 0.6854706), ("=BETA.INV(0.685470581,8,10,1,3)", 2),
    ("=CHISQ.DIST(0.5,1,TRUE)", 0.52049988), ("=CHISQ.DIST(2,3,FALSE)", 0.20755375),
    ("=CHISQ.DIST.RT(18.307,10)", 0.0500006), ("=CHIDIST(18.307,10)", 0.0500006), ("=CHISQ.INV(0.93,1)", 3.283020287),
    ("=CHISQ.INV.RT(0.050001,10)", 18.306973),
    ("=CHISQ.TEST({58,35;11,25;10,23},{45.35,47.65;17.56,18.44;16.09,16.91})", 0.000308192),
    ("=T.DIST(60,1,TRUE)", 0.99469533), ("=T.DIST(8,3,FALSE)", 0.00073691),
    ("=T.DIST.2T(1.959999998,60)", 0.054645), ("=T.DIST.RT(1.959999998,60)", 0.0273224),
    ("=TDIST(1.959999998,60,2)", 0.054645), ("=T.INV(0.75,2)", 0.8164966), ("=T.INV.2T(0.546449,60)", 0.606533076),
    ("=TINV(0.546449,60)", 0.606533076),
    ("=T.TEST({3,4,5,8,9,1,2,4,5},{6,19,3,2,14,4,5,17,1},2,1)", 0.196016),
    ("=F.DIST(15.2069,6,4,TRUE)", 0.99), ("=F.DIST(15.2069,6,4,FALSE)", 0.0012238),
    ("=F.DIST.RT(15.2068649,6,4)", 0.01), ("=F.INV(0.01,6,4)", 0.10930991), ("=F.INV.RT(0.01,6,4)", 15.20686),
    ("=F.TEST({6,7,9,15,21},{20,28,31,38,40})", 0.648318),
    ("=Z.TEST({3,6,7,8,6,5,4,2,1,9},4)", 0.090574), ("=Z.TEST({3,6,7,8,6,5,4,2,1,9},6)", 0.863043),
    ("=CONFIDENCE.NORM(0.05,2.5,50)", 0.692952), ("=CONFIDENCE.T(0.05,1,50)", 0.284196855),
    ("=POISSON.DIST(2,5,TRUE)", 0.124652), ("=POISSON.DIST(2,5,FALSE)", 0.084224),
    ("=HYPGEOM.DIST(1,4,8,20,TRUE)", 0.4654283), ("=HYPGEOM.DIST(1,4,8,20,FALSE)", 0.3632611),
    ("=NEGBINOM.DIST(10,5,0.25,TRUE)", 0.3135141), ("=NEGBINOM.DIST(10,5,0.25,FALSE)", 0.0550487),
    ("=BINOM.INV(6,0.5,0.75)", 4), ("=BINOM.DIST.RANGE(60,0.75,48)", 0.083974967),
    ("=BINOM.DIST.RANGE(60,0.75,45,50)", 0.523629793),
    ("=AVEDEV(4,5,6,7,5,4,3)", 1.020408), ("=DEVSQ(4,5,8,7,11,4,3)", 48),
    ("=KURT(3,4,5,2,3,4,5,6,4,7)", -0.151799637), ("=SKEW(3,4,5,2,3,4,5,6,4,7)", 0.359543),
    ("=SKEW.P(3,4,5,2,3,4,5,6,4,7)", 0.303193), ("=HARMEAN(4,5,8,7,11,4,3)", 5.028376),
    ("=TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)", 3.777778), ("=FISHER(0.75)", 0.972955),
    ("=FISHERINV(0.972955)", 0.75), ("=PERMUTATIONA(3,2)", 9),
    ("=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},2)", 0.333), ("=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},4)", 0.555),
    ("=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},8)", 0.666), ("=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},5)", 0.583),
    ("=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},7)", 0.7), ("=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43)", 0.381),
    ("=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43,1)", 0.3),
    ("=PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", 0.1), ("=PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)", 0.8),
    ("=COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})", 5.2), ("=COVARIANCE.S({2,4,8},{5,11,12})", 9.666666667),
    ("=STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 3.305719), ("=RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 0.05795),
    ("=FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})", 10.607253),
    ("=FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})", 10.607253),
    ("=INDEX(LINEST({1,9,5,7},{0,4,2,3}),1)", 2), ("=INDEX(LINEST({1,9,5,7},{0,4,2,3}),2)", 1),
    ("=INDEX(LOGEST({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16}),1)", 1.463275628),
    ("=INDEX(LOGEST({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16}),2)", 495.3047702),
    ("=INDEX(GROWTH({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16},{17;18}),1)", 320196.7184),
    ("=INDEX(GROWTH({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16},{17;18}),2)", 468536.0542),
]


@pytest.mark.parametrize("formula,expected", EXCEL_EXAMPLES)
def test_excel_documentation_examples(ev, formula, expected):
    assert ev(formula) == pytest.approx(expected, rel=1e-5, abs=1e-9)


# Extreme tails keep full relative precision (reference values from SciPy).
TAILS = [
    ("=NORM.S.DIST(-30,TRUE)", 4.906713927147908e-198), ("=CHISQ.DIST.RT(100,3)", 1.5541594313896026e-21),
    ("=T.DIST.2T(50,3)", 1.761715204127197e-05), ("=F.DIST.RT(1000,5,7)", 7.916143027631841e-10),
    ("=NORM.S.INV(1E-300)", -37.0470962993612), ("=T.INV(1E-10,4)", -416.1751403296041),
    ("=BETA.INV(0.999,0.5,0.5)", 0.999997532600929), ("=GAMMA.INV(1E-12,0.5,1)", 7.853981633974466e-25),
    ("=POISSON.DIST(1000,900,TRUE)", 0.9995093672671425), ("=BINOM.DIST(40000,100000,0.4,TRUE)", 0.5013734150485321),
]


@pytest.mark.parametrize("formula,expected", TAILS)
def test_precision_in_tails(ev, formula, expected):
    assert ev(formula) == pytest.approx(expected, rel=1e-9)


@pytest.mark.parametrize("formula,expected", [
    ("=NORM.DIST(1,0,0,TRUE)", NUM), ("=NORM.INV(0,0,1)", NUM), ("=NORM.INV(1,0,1)", NUM), ("=NORM.S.INV(1.5)", NUM),
    ("=LOGNORM.DIST(0,1,1,TRUE)", NUM), ("=EXPON.DIST(-1,1,TRUE)", NUM), ("=WEIBULL.DIST(1,0,1,TRUE)", NUM),
    ("=GAMMA(0)", NUM), ("=GAMMA(-2)", NUM), ("=GAMMALN(0)", NUM), ("=GAMMA.DIST(-1,1,1,TRUE)", NUM),
    ("=BETA.DIST(4,8,10,TRUE,1,3)", NUM), ("=BETA.DIST(0.5,1,1,TRUE,1,1)", NUM), ("=BETA.INV(0,8,10)", NUM),
    ("=CHISQ.DIST(1,0,TRUE)", NUM), ("=CHISQ.DIST.RT(-1,2)", NUM), ("=CHISQ.INV(1.1,2)", NUM),
    ("=CHISQ.TEST({1,2},{1,2,3})", NA), ("=CHISQ.TEST({1,2},{0,2})", DIV0), ("=T.DIST(1,0,TRUE)", NUM),
    ("=T.DIST.2T(-1,5)", NUM), ("=TDIST(1,5,3)", NUM), ("=T.INV(0,5)", NUM), ("=T.INV.2T(0,5)", NUM),
    ("=T.TEST({1,2,3},{1,2},2,1)", NA), ("=T.TEST({1,2,3},{4,5,6},3,1)", NUM), ("=T.TEST({1},{4,5,6},2,2)", DIV0),
    ("=F.DIST(-1,2,2,TRUE)", NUM), ("=F.INV(2,2,2)", NUM), ("=F.TEST({1,1,1},{1,2,3})", DIV0),
    ("=Z.TEST({1,2,3},2,0)", NUM), ("=CONFIDENCE.NORM(0,1,10)", NUM), ("=CONFIDENCE.T(0.05,1,1)", DIV0),
    ("=POISSON.DIST(-1,2,TRUE)", NUM), ("=HYPGEOM.DIST(5,4,8,20,TRUE)", NUM), ("=NEGBINOM.DIST(1,0,0.5,TRUE)", NUM),
    ("=BINOM.INV(-1,0.5,0.5)", NUM), ("=BINOM.DIST.RANGE(10,0.5,6,5)", NUM),
    ("=SKEW(1,2)", DIV0), ("=KURT(1,2,3)", DIV0), ("=SKEW(2,2,2)", DIV0), ("=HARMEAN(1,0)", NUM),
    ("=TRIMMEAN({1,2,3},1)", NUM), ("=FISHER(1)", NUM), ("=PERCENTRANK({1,2,3},4)", NA),
    ("=PERCENTRANK({1,2,3},2,0)", NUM), ("=PROB({1,2},{0.5,0.4},1)", NUM), ("=PROB({1,2},{0.5},1)", NA),
    ("=MODE.MULT(1,2,3)", NA), ("=COVARIANCE.S({1},{2})", DIV0), ("=COVAR({1,2},{1,2,3})", NA),
    ("=STEYX({1,2},{3,4})", DIV0), ("=FORECAST(1,{1,2},{3,3})", DIV0), ("=FORECAST(1,{1,2,3},{1,2})", NA),
    ('=FORECAST("x",{1,2},{3,4})', VALUE), ("=LINEST({1,2,3},{1,2})", REF), ('=LINEST({1,"a",3})', VALUE),
    ("=LOGEST({1,0,3})", NUM), ("=LINEST({5},{1})", VALUE),
])
def test_error_cases(ev, formula, expected):
    assert ev(formula) is expected


def test_storage_prefixes():
    from termxl.formula.parser import to_display, to_storage
    assert to_storage("=NORM.S.DIST(1,TRUE)+T.DIST.2T(1,2)") == "=_xlfn.NORM.S.DIST(1,TRUE)+_xlfn.T.DIST.2T(1,2)"
    assert to_storage("=GAMMA(2)*GAMMALN(3)") == "=_xlfn.GAMMA(2)*GAMMALN(3)"
    assert to_storage("=NORMSDIST(1)+FORECAST(1,A1:A3,B1:B3)") == "=NORMSDIST(1)+FORECAST(1,A1:A3,B1:B3)"
    assert to_display("=_xlfn.FORECAST.LINEAR(1,A1:A3,B1:B3)") == "=FORECAST.LINEAR(1,A1:A3,B1:B3)"


def test_distributions_lift_over_arrays(ev):
    assert ev("=SUM(NORM.S.DIST({-1,0,1},TRUE))") == pytest.approx(1.5)
    assert ev("=SUMPRODUCT(POISSON.DIST({0,1,2},1,FALSE))") == pytest.approx(2.5 / 2.718281828459045)


def test_a_variants_count_text_and_logicals(ev):
    cells = {"A1": "1", "A2": "TRUE", "A3": "text", "A4": "3"}
    assert ev("=VARA(A1:A4)", **cells) == pytest.approx(1.583333333333)  # 1, 1, 0, 3
    assert ev("=VARPA(A1:A4)") == pytest.approx(1.1875)
    assert ev("=STDEVA(A1:A4)") == pytest.approx(1.583333333333**0.5)
    assert ev("=STDEVPA(A1:A4)") == pytest.approx(1.1875**0.5)
    assert ev('=VARA(1,TRUE,"3")') == pytest.approx(4 / 3)  # typed text numbers count


def test_array_results(ev):
    assert ev('=TEXTJOIN(",",,FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}))') == "1,2,4,2"
    assert ev('=TEXTJOIN(",",,FREQUENCY({1,5,9},{6,2}))') == "1,1,1"  # unsorted bins keep their order
    assert ev('=TEXTJOIN(",",,MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1}))') == "1,2,3"
    assert ev("=ROWS(FREQUENCY({1,2},{1,2,3}))") == 4
    assert ev("=ROWS(LINEST({1,9,5,7},{0,4,2,3},TRUE,TRUE))*10+COLUMNS(LINEST({1,9,5,7},{0,4,2,3},TRUE,TRUE))") == 52
    assert ev('=TEXTJOIN(",",,TREND({2,4,6},,{4,5}))') == "8,10"
    assert ev("=ROWS(TREND({1;2;3},{1;2;3}))") == 3


def test_linest_statistics(ws):
    # y depends on two regressors; reference values from LibreOffice and R's lm().
    xs1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    xs2 = [2.0, 1.5, 3.7, 2.2, 4.9, 3.3, 5.8, 4.1, 6.6, 5.0]
    ys = [3.1, 5.2, 6.8, 9.1, 10.9, 13.2, 14.8, 17.3, 18.7, 21.4]
    for i, (a, b, y) in enumerate(zip(xs1, xs2, ys), 1):
        ws[f"A{i}"], ws[f"B{i}"], ws[f"C{i}"] = a, b, y
    expect = {
        (2, 1): 0.036404423358503, (2, 2): 0.0203616407542848, (2, 3): 0.092198262701772,
        (3, 1): 0.999741802325317, (3, 2): 0.110778838460943, (4, 1): 13552.0055029131, (4, 2): 7,
        (5, 1): 332.619096342644, (5, 2): 0.08590365735529,
    }
    for (r, c), want in expect.items():
        ws["E1"] = f"=INDEX(LINEST(C1:C10,A1:B10,TRUE,TRUE),{r},{c})"
        assert ws["E1"] == pytest.approx(want, rel=1e-9), (r, c)
    ws["E1"] = "=INDEX(LINEST(C1:C10,A1:B10,TRUE,TRUE),3,3)"
    assert ws["E1"] is NA
    ws["E1"] = "=INDEX(LINEST(C1:C10,A1:B10,FALSE,TRUE),2,3)"
    assert ws["E1"] is NA  # no standard error for a forced-zero intercept


def test_linest_drops_collinear_columns(ws):
    for i in range(1, 7):
        ws[f"A{i}"], ws[f"B{i}"], ws[f"C{i}"] = i, 2 * i, 3 * i + (i % 2)
    ws["E1"] = "=INDEX(LINEST(C1:C6,A1:B6,TRUE,TRUE),1,1)"  # B is 2*A: removed, coefficient 0
    assert ws["E1"] == 0
    ws["E1"] = "=INDEX(LINEST(C1:C6,A1:B6,TRUE,TRUE),2,1)"
    assert ws["E1"] == 0
    ws["E1"] = "=INDEX(LINEST(C1:C6,A1:B6,TRUE,TRUE),1,2)"
    assert ws["E1"] == pytest.approx(51 / 17.5)
    ws["E1"] = "=INDEX(LINEST(C1:C6,A1:B6,TRUE,TRUE),4,2)"  # degrees of freedom count only kept columns
    assert ws["E1"] == 4


def test_linest_exact_fit(ev):
    assert ev("=INDEX(LINEST({2,4,6},{1,2,3},TRUE,TRUE),5,2)") == 0
    assert ev("=INDEX(LINEST({2,4,6},{1,2,3},TRUE,TRUE),4,1)") is NA
    assert ev("=INDEX(LINEST({2,4,6},{1,2,3},TRUE,TRUE),3,1)") == 1
