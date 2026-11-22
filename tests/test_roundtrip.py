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


def _save_with_saved_values(wb, p):
    """Save wb to p with 42 as the value Excel saved for each formula on the first sheet."""
    wb.save(p)
    src = zipfile.ZipFile(p)
    infos = src.infolist()
    data = {i.filename: src.read(i.filename) for i in infos}
    src.close()
    sheet1 = data["xl/worksheets/sheet1.xml"]
    data["xl/worksheets/sheet1.xml"] = sheet1.replace(b"<v></v>", b"<v>42</v>").replace(b"<v />", b"<v>42</v>")
    with zipfile.ZipFile(p, "w") as out:
        for item in infos:
            out.writestr(item, data[item.filename])


def test_cached_value_survives_reference_rewrites(tmp_path):
    """A formula we can't calculate keeps Excel's saved value when its references are shifted or
    renamed, and loses it only when the cell gets new content (undo brings it back)."""
    from termse.formula.errors import NAME
    from termse.model import structure

    p = str(tmp_path / "u.xlsx")
    wb = openpyxl.Workbook()
    wb.active["B1"] = 5
    wb.active["A1"] = "=WEBSERVICE(B1)"
    wb.create_sheet("Other")["A1"] = "=Sheet!A1"
    _save_with_saved_values(wb, p)
    book = xlsx.load(p)
    s, other = book.sheets
    structure.insert(book, s, "row", 1, 1)
    structure.rename_sheet(book, s, "Data")
    assert s.formulas[(2, 1)].text == "=WEBSERVICE(B2)"
    book.set_input(s, 2, 2, "6")  # recalculates A2
    assert book.get_value(s, 2, 1) == 42 and (2, 1) in s.stale
    assert book.get_value(other, 1, 1) == 42
    book.set_input(s, 2, 1, "=WEBSERVICE(B2)+1")
    assert book.get_value(s, 2, 1) == NAME
    book.do_undo()
    assert book.get_value(s, 2, 1) == 42
    book.do_redo()
    assert book.get_value(s, 2, 1) == NAME


def test_cached_value_dropped_when_reference_is_deleted(tmp_path):
    """Deleting a row a formula we can't calculate refers to leaves it with #REF!, so Excel's saved
    value no longer applies. Undo brings it back, also for formulas in the deleted rows."""
    from termse.formula.errors import NAME
    from termse.model import structure

    p = str(tmp_path / "u.xlsx")
    wb = openpyxl.Workbook()
    wb.active["B2"] = 5
    wb.active["A1"] = "=WEBSERVICE(B2)"
    wb.active["A2"] = "=WEBSERVICE(B1)"
    _save_with_saved_values(wb, p)
    book = xlsx.load(p)
    s = book.sheets[0]
    structure.delete(book, s, "row", 2, 1)
    assert s.formulas[(1, 1)].text == "=WEBSERVICE(#REF!)"
    assert book.get_value(s, 1, 1) == NAME and (1, 1) not in s.stale
    book.do_undo()
    assert s.formulas[(1, 1)].text == "=WEBSERVICE(B2)"
    assert book.get_value(s, 1, 1) == 42 and (1, 1) in s.stale
    assert book.get_value(s, 2, 1) == 42 and (2, 1) in s.stale
    book.do_redo()
    assert book.get_value(s, 1, 1) == NAME


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


def test_control_characters_save(tmp_path):
    """Text with control characters (formula results, pasted values, spills) can be saved;
    the file spells them _x0007_ as Excel does."""
    from termse.model.refs import Rect

    book = xlsx.new_workbook()
    s = book.sheets[0]
    book.set_input(s, 1, 1, "=CHAR(7)")
    book.set_input(s, 1, 2, '="a"&CHAR(1)&SEQUENCE(2)')
    book.paste(book.capture(s, Rect(1, 1, 1, 1)), s, 3, 1, values_only=True)
    book.paste_text(s, 4, 1, "x\x07y")
    assert book.get_value(s, 3, 1) == "\x07"
    assert book.get_value(s, 4, 1) == "x\x07y"
    path = str(tmp_path / "ctrl.xlsx")
    xlsx.save(book, path)
    assert book.get_value(s, 4, 1) == "x\x07y"  # the escaping is only in the file
    ws = openpyxl.load_workbook(path).active
    assert ws["A3"].value == "_x0007_"
    assert ws["A4"].value == "x_x0007_y"
    assert ws["B2"].value == "a_x0001_2"
    with zipfile.ZipFile(path) as z:
        assert b"<v>_x0007_</v>" in z.read("xl/worksheets/sheet1.xml")


def test_sheet_settings_follow_inserted_and_deleted_rows(tmp_path):
    """Auto-filter, validation and conditional-format formulas and print titles/area move with
    inserted rows/columns, and undoing a deletion brings them back."""
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation

    from termse.model import structure

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 6):
        ws.cell(r, 1, r)
    ws.auto_filter.ref = "A1:A5"
    dv = DataValidation(type="list", formula1="$C$1:$C$3")
    dv.add("B2:B5")
    ws.add_data_validation(dv)
    ws.conditional_formatting.add("A1:A5", FormulaRule(formula=["$A1>$D$1"]))
    ws.print_title_rows = "1:1"
    ws.print_title_cols = "A:A"
    ws.print_area = "A1:D10"
    other = wb.create_sheet("O")
    dv2 = DataValidation(type="list", formula1="S!$C$1:$C$3")
    dv2.add("A1")
    other.add_data_validation(dv2)
    p = str(tmp_path / "settings.xlsx")
    wb.save(p)

    book = xlsx.load(p)
    s, o = book.sheets
    w = s.ws

    def state():
        cf = [(str(c.sqref), r.formula) for c in w.conditional_formatting for r in c.rules]
        return (w.auto_filter.ref, w.data_validations.dataValidation[0].formula1, cf, w.print_title_rows,
                w.print_title_cols, w.print_area, o.ws.data_validations.dataValidation[0].formula1)

    before = state()
    structure.insert(book, s, "row", 1, 2)
    assert state() == ("A3:A7", "$C$3:$C$5", [("A3:A7", ["$A3>$D$3"])], "$3:$3", "$A:$A",
                       "'S'!$A$3:$D$12", "S!$C$3:$C$5")
    xlsx.save(book, p)
    saved = openpyxl.load_workbook(p)["S"]
    assert saved.auto_filter.ref == "A3:A7" and saved.print_title_rows == "$3:$3"
    book.do_undo()
    assert state() == before
    structure.insert(book, s, "col", 1, 1)
    assert state()[:2] == ("B1:B5", "$D$1:$D$3") and state()[3:5] == ("$1:$1", "$B:$B")
    book.do_undo()
    structure.delete(book, s, "row", 1, 3)
    assert state()[1] == "#REF!" and state()[3] is None
    book.do_undo()
    assert state() == before
