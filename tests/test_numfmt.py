import pytest

from termxl.format.numfmt import format_value, general, is_date_format
from termxl.model.inputparse import parse_input


@pytest.mark.parametrize("value,fmt,expected", [
    (1234.567, "#,##0.00", "1,234.57"),
    (-1234.567, "#,##0.00;[Red](#,##0.00)", "(1,234.57)"),
    (0.256, "0.0%", "25.6%"),
    (1234567, "#,##0,", "1,235"),
    (12345678, '0.0,,"M"', "12.3M"),
    (46347, "mmm d, yyyy", "Nov 21, 2026"),
    (46347.75, "m/d/yyyy h:mm AM/PM", "11/21/2026 6:00 PM"),
    (46347, "dddd, mmmm dd, yyyy", "Saturday, November 21, 2026"),
    (1.5, "[h]:mm:ss", "36:00:00"),
    (0.5, "h:mm", "12:00"),
    (12345.678, "0.00E+00", "1.23E+04"),
    (0.000123, "0.00E+00", "1.23E-04"),
    (1.75, "# ?/?", "1 3/4"),
    (5, "0000", "0005"),
    (5551234567, "000-000-0000", "555-123-4567"),
    (-5, '0;(0);"zero"', "(5)"),
    (0, '0;(0);"zero"', "zero"),
    ("hi", '0;0;0;"["@"]"', "[hi]"),
    (1234.5, "$#,##0.00_);($#,##0.00)", "$1,234.50 "),
    (-1234.5, "$#,##0.00_);($#,##0.00)", "($1,234.50)"),
    (1 / 3, "General", "0.333333333"),
    (0.1 + 0.2, "General", "0.3"),
    (-0.5, "0", "-1"),
    (1, "#.##", "1."),
    (0, "#,###", ""),
    (60, "yyyy-mm-dd", "1900-02-29"),
    (2.675, "0.00", "2.68"),
    (1234.5, "[$€-407] #,##0.00", "€ 1,234.50"),
    (25, '[<50]"low";"high"', "low"),
    (75, '[<50]"low";"high"', "high"),
    (0.75, "h:mm AM/PM", "6:00 PM"),
    (True, "0.00", "TRUE"),
    (None, "0.00", ""),
    (45000.5, "yyyy-mm-dd hh:mm:ss", "2023-03-15 12:00:00"),
    (0.000011574, "h:mm:ss.00", "0:00:01.00"),
])
def test_format(value, fmt, expected):
    assert format_value(value, fmt)[0] == expected


def test_color():
    assert format_value(-1, "0;[Red]-0") == ("-1", "FF0000")


def test_general_width():
    assert general(123456789012) == "1.23457E+11"
    assert general(-1 / 3) == "-0.33333333"
    assert general(1e-12) == "1E-12"


def test_is_date_format():
    assert is_date_format("yyyy-mm-dd")
    assert is_date_format("[$-409]mmmm d, yyyy")
    assert is_date_format("h:mm AM/PM")
    assert not is_date_format('"$"#,##0.00')
    assert not is_date_format("0.00E+00")
    assert not is_date_format("General")


@pytest.mark.parametrize("text,kind,value,fmt", [
    ("42", "number", 42, None),
    ("-3.5", "number", -3.5, None),
    ("1,234", "number", 1234, "#,##0"),
    ("12%", "number", 0.12, "0%"),
    ("$5.50", "number", 5.5, '$#,##0.00_);($#,##0.00)'),
    ("(100)", "number", -100, None),
    ("2026-11-21", "number", 46347, "yyyy-mm-dd"),
    ("11/21/2026", "number", 46347, "m/d/yyyy"),
    ("TRUE", "bool", True, None),
    ("#N/A", "error", "#N/A", None),
    ("'007", "text", "007", None),
    ("=A1", "formula", "=A1", None),
    ("hello", "text", "hello", None),
    ("1e3", "number", 1000.0, "0.00E+00"),
])
def test_parse_input(text, kind, value, fmt):
    p = parse_input(text)
    assert p.kind == kind
    assert p.value == value
    assert p.number_format == fmt


def test_parse_time():
    p = parse_input("1:30 PM")
    assert p.kind == "number" and p.value == pytest.approx(13.5 / 24)
    assert p.number_format == "h:mm AM/PM"
