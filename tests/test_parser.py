import pytest

from termxl.formula.parser import (
    FormulaSyntaxError, Unsupported, normalize, parse, to_display, to_storage, translate,
)


@pytest.mark.parametrize("src,expected", [
    ("=1+2*3", "(Const(1) + (Const(2) * Const(3)))"),
    ("=-2^2", "((-Const(2)) ^ Const(2))"),           # Excel: negation binds tighter than ^
    ("=2^3^2", "((Const(2) ^ Const(3)) ^ Const(2))"),  # left associative
    ("=1-2-3", "((Const(1) - Const(2)) - Const(3))"),
    ("=50%*2", "((Const(50)%) * Const(2))"),
    ('="a"&"b"="ab"', "((Const('a') & Const('b')) = Const('ab'))"),
    ("=A1:B2", "Ref(None!A1:B2)"),
    ("='My Sheet'!$A$1", "Ref(My Sheet!A1:A1)"),
    ("=IF(A1,,1)", "IF(Ref(None!A1:A1), Missing(), Const(1))"),
    ("={1,2;3,4}", "Array([[1, 2], [3, 4]])"),
    ("=_xlfn.XLOOKUP(1,A:A,B:B)", "XLOOKUP(Const(1), Ref(None!A1:A1048576), Ref(None!B1:B1048576))"),
    ("=LOG10(100)", "LOG10(Const(100))"),
    ("=SUM(1:1)", "SUM(Ref(None!A1:XFD1))"),
])
def test_parse(src, expected):
    assert repr(parse(src)) == expected


@pytest.mark.parametrize("bad", ["=1+", "=SUM(1,2", "=(1", "=1 2", "=)"])
def test_syntax_errors(bad):
    with pytest.raises(FormulaSyntaxError):
        parse(bad)


def test_external_refs_unsupported():
    with pytest.raises(Unsupported):
        parse("=[1]Sheet1!A1")
    with pytest.raises(Unsupported):
        parse("='[Book.xlsx]Sheet1'!A1")


def test_translate():
    assert translate("=A1+$B$2+B$3+$C4", 2, 1) == "=B3+$B$2+C$3+$C6"
    assert translate("=SUM(A:A)+SUM(1:2)", 1, 1) == "=SUM(B:B)+SUM(2:3)"
    assert translate("=Sheet2!A1", 1, 0) == "=Sheet2!A2"
    assert translate("=A1", -1, 0) == "=#REF!"
    assert translate('="A1"&A1', 1, 0) == '="A1"&A2'  # strings untouched


def test_normalize_and_prefixes():
    assert normalize("=sum(a1:b2)+true") == "=SUM(A1:B2)+TRUE"
    assert to_storage("=xlookup(1,A:A,B:B)") == "=_xlfn.XLOOKUP(1,A:A,B:B)"
    assert to_storage("=FILTER(A1:A3,B1:B3)") == "=_xlfn._xlws.FILTER(A1:A3,B1:B3)"
    assert to_storage("=SUM(1)") == "=SUM(1)"
    assert to_display("=_xlfn.XLOOKUP(1,A:A,B:B)") == "=XLOOKUP(1,A:A,B:B)"
