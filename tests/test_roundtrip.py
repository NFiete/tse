"""Load -> edit -> save -> reload: everything we didn't touch must survive."""
import zipfile

import openpyxl

from termse.io import xlsx
from termse.model.refs import parse_addr


def _snapshot(ws):
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None or c.has_style:
                cells[c.coordinate] = (c.value, c.number_format, c.font.b, c.font.color.rgb if c.font.color else None,
                                       c.fill.fgColor.rgb if c.fill and c.fill.fill_type else None,
                                       c.alignment.horizontal)
    return {
        "cells": cells,
        "merged": sorted(m.coord for m in ws.merged_cells.ranges),
        "widths": {k: v.width for k, v in ws.column_dimensions.items()},
        "hidden_rows": sorted(r for r, d in ws.row_dimensions.items() if d.hidden),
        "freeze": ws.freeze_panes,
        "dv": [str(dv.sqref) for dv in ws.data_validations.dataValidation],
        "cf": sorted(str(cf.sqref) for cf in ws.conditional_formatting),
        "comments": {c.coordinate: c.comment.text for row in ws.iter_rows() for c in row if c.comment},
        "links": {c.coordinate: c.hyperlink.target for row in ws.iter_rows() for c in row if c.hyperlink},
    }


def test_untouched_roundtrip_is_identical(sample_path, tmp_path):
    before = openpyxl.load_workbook(sample_path)
    book = xlsx.load(sample_path)
    out = str(tmp_path / "out.xlsx")
    xlsx.save(book, out)
    after = openpyxl.load_workbook(out)
    assert before.sheetnames == after.sheetnames
    for name in before.sheetnames:
        assert _snapshot(before[name]) == _snapshot(after[name]), name
    assert set(before.defined_names) == set(after.defined_names)
    assert after["Notes"].sheet_properties.tabColor.rgb == before["Notes"].sheet_properties.tabColor.rgb


def test_edit_and_save(sample_path, tmp_path):
    book = xlsx.load(sample_path)
    budget = book.sheet_by_name("Budget")
    r, c = parse_addr("B4")
    book.set_input(budget, r, c, "1600")
    assert book.get_value(budget, *parse_addr("E4")) == 4600
    out = str(tmp_path / "edited.xlsx")
    xlsx.save(book, out)

    wb = openpyxl.load_workbook(out)
    assert wb["Budget"]["B4"].value == 1600
    assert wb["Budget"]["B4"].number_format == '"$"#,##0.00'
    assert wb["Budget"]["E4"].value == "=SUM(B4:D4)"
    # cached values were written for other readers
    wbv = openpyxl.load_workbook(out, data_only=True)
    assert wbv["Budget"]["E4"].value == 4600
    assert wbv["Data"]["I3"].value == 75
    assert wbv["Data"]["I5"].value == (7238.75 + 100) * 2
    assert wbv["Notes"]["A5"].value == "#DIV/0!"
    # and everything else is untouched
    before = _snapshot(openpyxl.load_workbook(sample_path)["Data"])
    assert _snapshot(wb["Data"]) == before


def test_cached_values_used_on_load(sample_path, tmp_path):
    """Values Excel saved are shown as-is; formulas without cached values get computed."""
    book = xlsx.load(sample_path)  # openpyxl-built file: no cached values -> computed
    s = book.sheet_by_name("Budget")
    assert book.get_value(s, *parse_addr("E9")) == 7238.75
    out = str(tmp_path / "c.xlsx")
    xlsx.save(book, out)
    with zipfile.ZipFile(out) as z:
        assert b"<v>" in z.read("xl/worksheets/sheet1.xml")
    book2 = xlsx.load(out)
    s2 = book2.sheet_by_name("Budget")
    assert s2.cached[parse_addr("E9")][1] == 7238.75


def test_unsupported_function_keeps_cached_value(tmp_path):
    p = str(tmp_path / "u.xlsx")
    wb = openpyxl.Workbook()
    wb.active["A1"] = '=WEBSERVICE("https://example.com/n")'
    wb.save(p)
    # inject a cached value as Excel would have written it
    import shutil

    src = zipfile.ZipFile(p)
    p2 = str(tmp_path / "u2.xlsx")
    with zipfile.ZipFile(p2, "w") as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                assert b"<v></v>" in data or b"<v />" in data
                data = data.replace(b"<v></v>", b"<v>42</v>").replace(b"<v />", b"<v>42</v>")
            out.writestr(item, data)
    src.close()
    shutil.move(p2, p)
    book = xlsx.load(p)
    s = book.sheets[0]
    assert book.get_value(s, 1, 1) == 42
    assert (1, 1) in s.stale


def test_csv_roundtrip(tmp_path):
    p = tmp_path / "data.csv"
    p.write_text("name,qty,price\nfoo,2,1.50\nbar,3,=B3*2\n")
    book = xlsx.load(str(p))
    s = book.sheets[0]
    assert book.get_value(s, 2, 2) == 2
    assert book.get_value(s, 3, 3) == 6
    out = tmp_path / "out.csv"
    xlsx.save(book, str(out), active_sheet=s)
    assert out.read_text().splitlines() == ["name,qty,price", "foo,2,1.5", "bar,3,6"]


def test_new_file_save(tmp_path):
    book = xlsx.new_workbook()
    s = book.sheets[0]
    book.set_input(s, 1, 1, "=SUM(1,2,3)")
    out = str(tmp_path / "new.xlsx")
    xlsx.save(book, out)
    wb = openpyxl.load_workbook(out, data_only=True)
    assert wb.active["A1"].value == 6


def test_detect_unpreservable(tmp_path, sample_path):
    assert xlsx.detect_unpreservable(sample_path) == []
    p = str(tmp_path / "s.xlsx")
    with zipfile.ZipFile(sample_path) as zin, zipfile.ZipFile(p, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("xl/slicers/slicer1.xml", "<x/>")
    assert "slicers" in xlsx.detect_unpreservable(p)
