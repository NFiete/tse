"""Engineering functions. Expected values are Microsoft's documented examples unless noted."""
import math

import pytest

from termse.formula.errors import DIV0, NA, NUM, VALUE
from termse.formula.parser import to_storage


@pytest.mark.parametrize("formula, expected", [
    ("=BIN2DEC(1100100)", 100), ("=BIN2DEC(1111111111)", -1), ('=BIN2DEC("")', 0),
    ("=BIN2HEX(11111011,4)", "00FB"), ("=BIN2HEX(1110)", "E"), ("=BIN2HEX(1111111111)", "FFFFFFFFFF"),
    ("=BIN2OCT(1001,3)", "011"), ("=BIN2OCT(1100100)", "144"), ("=BIN2OCT(1111111111)", "7777777777"),
    ("=DEC2BIN(9,4)", "1001"), ("=DEC2BIN(-100)", "1110011100"), ("=DEC2BIN(9.9)", "1001"),
    ('=DEC2BIN("12")', "1100"), ("=DEC2BIN(511)", "111111111"), ("=DEC2BIN(-512)", "1000000000"),
    ("=DEC2HEX(100,4)", "0064"), ("=DEC2HEX(-54)", "FFFFFFFFCA"), ("=DEC2HEX(28)", "1C"),
    ("=DEC2HEX(549755813887)", "7FFFFFFFFF"), ("=DEC2OCT(58,3)", "072"), ("=DEC2OCT(-100)", "7777777634"),
    ('=HEX2BIN("F",8)', "00001111"), ('=HEX2BIN("B7")', "10110111"), ('=HEX2BIN("FFFFFFFFFF")', "1111111111"),
    ('=HEX2BIN("FFFFFFFE00")', "1000000000"),
    ('=HEX2DEC("A5")', 165), ('=HEX2DEC("FFFFFFFF5B")', -165), ('=HEX2DEC("3DA408B9")', 1034160313),
    ('=HEX2DEC("ff")', 255), ("=HEX2DEC(255)", 597),
    ('=HEX2OCT("F",3)', "017"), ('=HEX2OCT("3B4E")', "35516"), ('=HEX2OCT("FFFFFFFF00")', "7777777400"),
    ("=OCT2BIN(3,3)", "011"), ("=OCT2BIN(7777777000)", "1000000000"),
    ("=OCT2DEC(54)", 44), ("=OCT2DEC(7777777533)", -165),
    ("=OCT2HEX(100,4)", "0040"), ("=OCT2HEX(7777777533)", "FFFFFFFF5B"),
    # places is ignored for negative numbers
    ("=DEC2BIN(-1,2)", "1111111111"),
])
def test_base_conversions(ev, formula, expected):
    assert ev(formula) == expected


@pytest.mark.parametrize("formula, expected", [
    ("=BIN2DEC(12)", NUM), ('=BIN2DEC("11111111111")', NUM), ("=BIN2DEC(-1)", NUM), ("=BIN2DEC(1.5)", NUM),
    ('=HEX2DEC(" 1")', NUM), ('=HEX2DEC("0x1")', NUM), ('=HEX2DEC("1_0")', NUM),
    ('=HEX2BIN("200")', NUM), ('=HEX2BIN("FFFFFFFDFF")', NUM), ('=HEX2OCT("FFDFFFFFFF")', NUM),
    ("=DEC2BIN(512)", NUM), ("=DEC2BIN(-513)", NUM), ("=DEC2HEX(549755813888)", NUM),
    ("=DEC2OCT(-536870913)", NUM), ("=DEC2HEX(64,1)", NUM),
    ("=BIN2HEX(101,-1)", NUM), ("=BIN2HEX(101,11)", NUM), ('=BIN2HEX(101,"a")', VALUE),
    ('=DEC2BIN("x")', VALUE), ("=DEC2BIN(TRUE)", VALUE), ("=BIN2DEC(TRUE)", VALUE), ("=DEC2BIN(1/0)", DIV0),
])
def test_base_conversion_errors(ev, formula, expected):
    assert ev(formula) is expected


@pytest.mark.parametrize("formula, expected", [
    ("=BITAND(1,5)", 1), ("=BITAND(13,25)", 9), ("=BITOR(23,10)", 31), ("=BITXOR(5,3)", 6),
    ("=BITLSHIFT(4,2)", 16), ("=BITRSHIFT(13,2)", 3), ("=BITLSHIFT(4,-2)", 1), ("=BITRSHIFT(5,-2)", 20),
    ("=BITAND(2^48-1,1)", 1), ('=BITAND("3",1)', 1), ("=BITLSHIFT(1,48)", 2 ** 48),
    ("=BITAND(2^48,1)", NUM), ("=BITAND(-1,1)", NUM), ("=BITAND(1.5,1)", NUM), ('=BITOR("a",1)', VALUE),
    ("=BITLSHIFT(1,54)", NUM), ("=BITRSHIFT(1,-54)", NUM),
])
def test_bit_functions(ev, formula, expected):
    assert ev(formula) == expected


@pytest.mark.parametrize("formula, expected", [
    ("=DELTA(5,4)", 0), ("=DELTA(5,5)", 1), ("=DELTA(0.5,0)", 0), ("=DELTA(0)", 1), ('=DELTA("1",1)', 1),
    ("=GESTEP(5,4)", 1), ("=GESTEP(5,5)", 1), ("=GESTEP(-4,-5)", 1), ("=GESTEP(-1)", 0),
    ("=ERF(0.745)", pytest.approx(0.707928920)), ("=ERF(1)", pytest.approx(0.842700793)),
    ("=ERF(0,1)", pytest.approx(0.842700793)), ("=ERF(1,0)", pytest.approx(-0.842700793)),
    ("=ERF.PRECISE(0.745)", pytest.approx(0.707928920)), ("=ERF.PRECISE(1)", pytest.approx(0.842700793)),
    ("=ERFC(1)", pytest.approx(0.157299207)), ("=ERFC.PRECISE(1)", pytest.approx(0.157299207)),
    ("=ERFC(-1)", pytest.approx(1.842700793)),
])
def test_delta_gestep_erf(ev, formula, expected):
    assert ev(formula) == expected


@pytest.mark.parametrize("formula, expected", [
    ('=CONVERT(1,"lbm","kg")', 0.45359237), ('=CONVERT(68,"F","C")', 20), ('=CONVERT(6,"tsp","tbs")', 2),
    ('=CONVERT(6,"gal","l")', 22.712470704), ('=CONVERT(6,"mi","km")', 9.656064),
    ('=CONVERT(6,"km","mi")', pytest.approx(3.728227153)), ('=CONVERT(6,"in","ft")', 0.5),
    ('=CONVERT(6,"cm","in")', pytest.approx(2.362204724)), ('=CONVERT(CONVERT(100,"ft","m"),"ft","m")', 9.290304),
    ('=CONVERT(100,"C","K")', 373.15), ('=CONVERT(-40,"F","cel")', -40), ('=CONVERT(100,"C","Rank")', 671.67),
    ('=CONVERT(100,"C","Reau")', 80), ('=CONVERT(1,"mK","K")', 0.001),
    # prefixes: squared/cubed units square/cube the prefix; litres and ares don't
    ('=CONVERT(1,"km2","m2")', 1e6), ('=CONVERT(1,"km^3","m^3")', 1e9), ('=CONVERT(1,"mm2","cm2")', 0.01),
    ('=CONVERT(1,"mL","l")', 0.001), ('=CONVERT(1,"cm3","mL")', 1), ('=CONVERT(1,"kar","ar")', 1000),
    ('=CONVERT(1,"dam","m")', 10), ('=CONVERT(1,"kWh","J")', 3.6e6),
    ('=CONVERT(1,"km/h","m/s")', pytest.approx(1 / 3.6)),
    ('=CONVERT(1,"byte","bit")', 8), ('=CONVERT(1,"kibyte","byte")', 1024), ('=CONVERT(1,"Gibyte","Mibyte")', 1024),
    ('=CONVERT(1,"kbyte","byte")', 1000),
    ('=CONVERT(1,"ft","Pica")', 864), ('=CONVERT(1,"ft","pica")', 72), ('=CONVERT(1,"MTON","ft3")', 40),
    ('=CONVERT(1,"hr","sec")', 3600), ('=CONVERT(1,"yr","day")', 365.25), ('=CONVERT(1,"atm","mmHg")', 760),
    ('=CONVERT("2","m","cm")', 200),
    # errors: unknown units, units that take no prefix, different categories
    ('=CONVERT(2.5,"ft","sec")', NA), ('=CONVERT(1,"M","m")', NA), ('=CONVERT(1,"kft","ft")', NA),
    ('=CONVERT(1,"kim","m")', NA), ('=CONVERT(1,"m",3)', NA), ('=CONVERT("a","m","ft")', VALUE),
    ('=CONVERT(TRUE,"m","ft")', VALUE),
])
def test_convert(ev, formula, expected):
    assert ev(formula) == expected


@pytest.mark.parametrize("formula, expected", [
    ("=BESSELI(1.5,1)", 0.981666428), ("=BESSELJ(1.9,2)", 0.329925829), ("=BESSELK(1.5,1)", 0.277387804),
    ("=BESSELY(2.5,1)", 0.145918138),
    # reference values from Abramowitz and Stegun
    ("=BESSELJ(1,0)", 0.7651976866), ("=BESSELJ(10,1)", 0.0434727462), ("=BESSELJ(-3,1)", -0.3390589585),
    ("=BESSELJ(50,1)", -0.0975118281), ("=BESSELJ(1,10)", 2.630615124e-10), ("=BESSELJ(0,0)", 1),
    ("=BESSELY(1,0)", 0.0882569642), ("=BESSELY(10,1)", 0.2490154242), ("=BESSELY(50,0)", -0.0980649954),
    ("=BESSELY(0.5,3)", -42.059494304), ("=BESSELI(1,0)", 1.2660658778), ("=BESSELI(-2,1)", -1.5906368546),
    ("=BESSELI(10,2)", 2281.5189677),
])
def test_bessel(ev, formula, expected):
    assert ev(formula) == pytest.approx(expected, rel=1e-8)


@pytest.mark.parametrize("formula, expected", [
    ("=BESSELJ(1,-1)", NUM), ("=BESSELY(0,1)", NUM), ("=BESSELY(-1,1)", NUM), ("=BESSELK(0,1)", NUM),
    ('=BESSELJ("a",1)', VALUE), ("=BESSELY(1,500)", NUM),
])
def test_bessel_errors(ev, formula, expected):
    assert ev(formula) is expected


def test_besselj_matches_numerical_recipes(ev):
    # Excel's BESSELJ is accurate to about 1e-8: Microsoft's example is 0.329925829 where the
    # exact J_2(1.9) is 0.3299257277. Numerical Recipes' bessj gives the same digits.
    assert ev("=BESSELJ(1.9,2)") == pytest.approx(0.3299258286697852, rel=1e-12)


def test_besselk_matches_polynomial_approximation(ev):
    # Excel and LibreOffice use approximations accurate to about 1e-7; the exact K_1(1.5) is
    # 0.2773878004568, but Microsoft's example gives 0.277387804.
    assert ev("=BESSELK(1.5,1)") == pytest.approx(0.277387803602702, rel=1e-12)
    assert ev("=BESSELK(0.01,0)") == pytest.approx(4.72124473505925, rel=1e-12)


@pytest.mark.parametrize("formula, expected", [
    ("=COMPLEX(3,4)", "3+4i"), ('=COMPLEX(3,4,"j")', "3+4j"), ("=COMPLEX(0,1)", "i"), ("=COMPLEX(1,0)", "1"),
    ("=COMPLEX(0,0)", "0"), ('=COMPLEX(0,-1,"j")', "-j"), ("=COMPLEX(1.5,-2.25)", "1.5-2.25i"),
    ('=COMPLEX(3,4,"")', "3+4i"),
    ('=IMABS("5+12i")', 13), ('=IMAGINARY("3+4i")', 4), ('=IMAGINARY("0-j")', -1), ("=IMAGINARY(4)", 0),
    ('=IMAGINARY("4i")', 4), ('=IMAGINARY("-4.5j")', -4.5), ('=IMREAL("6-9i")', 6), ('=IMREAL("i")', 0),
    ('=IMREAL("1e3+2e-1i")', 1000), ('=IMREAL(".5+.5i")', 0.5), ('=IMREAL("+3-i")', 3),
    ('=IMARGUMENT("3+4i")', pytest.approx(0.92729522)), ('=IMARGUMENT("-i")', pytest.approx(-math.pi / 2)),
    ('=IMCONJUGATE("3+4i")', "3-4i"), ('=IMCONJUGATE("j")', "-j"),
    ('=IMCOS("1+i")', "0.833730025131149-0.988897705762865i"),
    ('=IMCOSH("4+3i")', "-27.0349456030742+3.85115333481178i"),
    ('=IMCOT("4+3i")', "0.00490118239430447-0.999266927805902i"),
    ('=IMCSC("4+3i")', "-0.0754898329158637+0.0648774713706355i"),
    ('=IMCSCH("4+3i")', "-0.036275889628626-0.0051744731840194i"),
    ('=IMDIV("-238+240i","10+24i")', "5+12i"), ('=IMDIV("3+4j","1-j")', "-0.5+3.5j"),
    ('=IMEXP("1+i")', "1.46869393991589+2.28735528717884i"),
    ('=IMLN("3+4i")', "1.6094379124341+0.927295218001612i"),
    ('=IMLOG10("3+4i")', "0.698970004336019+0.402719196273373i"),
    ('=IMLOG2("3+4i")', "2.32192809488736+1.33780421245098i"),
    ('=IMPOWER("2+3i",3)', "-46+9.00000000000001i"), ('=IMPOWER(2,10)', "1024"),
    ('=IMPOWER("i",2)', "-1+1.22464679914735E-16i"),
    ('=IMPRODUCT("3+4i","5-3i")', "27+11i"), ('=IMPRODUCT("1+2i",30)', "30+60i"),
    ('=IMSEC("4+3i")', "-0.0652940278579471-0.0752249603027732i"),
    ('=IMSECH("4+3i")', "-0.0362534969158689-0.00516434460775318i"),
    ('=IMSIN("4+3i")', "-7.61923172032141-6.548120040911i"),
    ('=IMSINH("4+3i")', "-27.0168132580039+3.85373803791938i"),
    ('=IMSQRT("1+i")', "1.09868411346781+0.455089860562227i"),
    ('=IMSQRT("-4")', "1.22464679914735E-16+2i"),
    ('=IMSUB("13+4i","5+3i")', "8+i"), ('=IMSUM("3+4i","5-3i")', "8+i"),
    ('=IMSUM({"1+i","2"},"3i")', "3+4i"), ('=IMPRODUCT({"1+i";"1-i"})', "2"), ('=IMSUM("i","-i")', "0"),
    ('=IMTAN("4+3i")', "0.00490825806749606+1.00070953606723i"),
    ('=INDEX(IMREAL({"1+i","2-3i"}),2)', 2),
])
def test_complex(ev, formula, expected):
    assert ev(formula) == expected


@pytest.mark.parametrize("formula, expected", [
    ('=COMPLEX(3,4,"I")', VALUE), ('=COMPLEX(3,4,"x")', VALUE), ("=COMPLEX(TRUE,1)", VALUE),
    ('=IMREAL("3+4")', NUM), ('=IMREAL("3+4I")', NUM), ('=IMAGINARY(" 4i")', NUM), ("=IMREAL(TRUE)", VALUE),
    ('=IMSUM("1+i","2+j")', VALUE), ('=IMDIV("1+i",0)', NUM), ("=IMARGUMENT(0)", DIV0), ("=IMLN(0)", NUM),
    ("=IMPOWER(0,0)", NUM), ("=IMPOWER(0,-1)", NUM), ("=IMCSC(0)", NUM), ("=IMCOT(0)", NUM),
    ("=IMCSCH(0)", NUM), ('=IMEXP("1000")', NUM), ("=IMSUM(1/0,1)", DIV0),
])
def test_complex_errors(ev, formula, expected):
    assert ev(formula) is expected


def test_complex_blank_cells_are_zero(ev):
    assert ev("=IMSUM(A1:A3)", A1="1+i", A3="2") == "3+i"
    assert ev("=IMREAL(B1)") == 0


def test_newer_functions_get_xlfn_prefix():
    assert to_storage("=BITAND(1,2)+ERF.PRECISE(1)") == "=_xlfn.BITAND(1,2)+_xlfn.ERF.PRECISE(1)"
    assert to_storage('=IMTAN("i")') == '=_xlfn.IMTAN("i")'
    assert to_storage('=IMSIN("i")&BIN2DEC(1)') == '=IMSIN("i")&BIN2DEC(1)'
