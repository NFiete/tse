import openpyxl
import pytest

from termse.formula.errors import NAME
from termse.io import xlsx
from termse.model import names
from termse.model.refs import Rect


@pytest.fixture
def setup(ws):
    for i, v in enumerate([5, 10, 20], 1):
        ws[f"A{i}"] = str(v)
    return ws


def test_define_from_selection_and_use(setup):
    ws = setup
    b, s = ws.book, ws.sheet
    ref = names.selection_ref(s, Rect(1, 1, 3, 1))
    assert ref == "Sheet!$A$1:$A$3"
    names.set_name(b, "Amounts", ref, None, s)
    ws["C1"] = "=SUM(Amounts)"
    assert ws["C1"] == 35
    ws["A2"] = "100"
    assert ws["C1"] == 125
    assert [(n.name, n.refers_to, n.scope_label) for n in names.list_names(b)] == [
        ("Amounts", "=Sheet!$A$1:$A$3", "Workbook")]


def test_unqualified_refs_are_anchored(setup):
    b, s = setup.book, setup.sheet
    info = names.set_name(b, "Top", "a1:a2", None, s)
    assert info.refers_to == "=Sheet!$A$1:$A$2"
    info = names.set_name(b, "Rate", "0.25", None, s)
    assert info.refers_to == "=0.25"


def test_edit_value_recalculates(setup):
    ws = setup
    b, s = ws.book, ws.sheet
    names.set_name(b, "Rate", "0.25", None, s)
    ws["C1"] = "=A3*Rate"
    assert ws["C1"] == 5
    names.set_name(b, "Rate", "0.5", None, s, old=("Rate", None))
    assert ws["C1"] == 10
    b.do_undo()
    assert ws["C1"] == 5
    b.do_redo()
    assert ws["C1"] == 10


def test_rename_updates_formulas_and_undo(setup):
    ws = setup
    b, s = ws.book, ws.sheet
    names.set_name(b, "Rate", "0.25", None, s)
    ws["C1"] = "=A3*rate+Rate"
    names.set_name(b, "TaxRate", "0.25", None, s, old=("Rate", None))
    assert ws.text("C1") == "=A3*TaxRate+TaxRate"
    assert ws["C1"] == 5.25
    b.do_undo()
    assert ws.text("C1") == "=A3*rate+Rate" and ws["C1"] == 5.25
    assert names.find(b, "Rate", None) == "Rate" and names.find(b, "TaxRate", None) is None


def test_delete_gives_name_error_and_undo(setup):
    ws = setup
    b, s = ws.book, ws.sheet
    names.set_name(b, "Rate", "0.25", None, s)
    ws["C1"] = "=A3*Rate"
    names.delete_name(b, "Rate", None)
    assert ws["C1"] == NAME
    b.do_undo()
    assert ws["C1"] == 5


def test_sheet_scope(setup):
    ws = setup
    b, s = ws.book, ws.sheet
    from termse.model import structure

    other = structure.add_sheet(b, "Other")
    names.set_name(b, "Rate", "0.1", None, s)
    names.set_name(b, "Rate", "0.5", other, other)
    ws["C1"] = "=A3*Rate"
    b.set_input(other, 1, 1, "=10*Rate")
    assert ws["C1"] == 2
    assert b.get_value(other, 1, 1) == 5
    scopes = sorted(n.scope_label for n in names.list_names(b))
    assert scopes == ["Other", "Workbook"]


@pytest.mark.parametrize("bad,msg", [
    ("", "Enter a name"), ("my name", "spaces"), ("1abc", "must start"), ("A1", "cell reference"),
    ("XFD100", "cell reference"), ("R1C1", "cell reference"), ("r", "cell reference"), ("TRUE", "TRUE"),
])
def test_invalid_names(setup, bad, msg):
    with pytest.raises(ValueError, match=msg):
        names.set_name(setup.book, bad, "1", None, setup.sheet)


def test_duplicate_and_bad_reference(setup):
    b, s = setup.book, setup.sheet
    names.set_name(b, "Rate", "1", None, s)
    with pytest.raises(ValueError, match="already defined"):
        names.set_name(b, "rate", "2", None, s)
    with pytest.raises(ValueError, match="problem"):
        names.set_name(b, "X", "=SUM(", None, s)
    with pytest.raises(ValueError, match="no sheet"):
        names.set_name(b, "X", "Nope!A1", None, s)


def test_names_saved_to_file(setup, tmp_path):
    b, s = setup.book, setup.sheet
    from termse.model import structure

    other = structure.add_sheet(b, "Other")
    names.set_name(b, "Amounts", "A1:A3", None, s)
    names.set_name(b, "Local", "0.5", other, other)
    out = str(tmp_path / "n.xlsx")
    xlsx.save(b, out)
    wb = openpyxl.load_workbook(out)
    assert wb.defined_names["Amounts"].attr_text == "Sheet!$A$1:$A$3"
    assert wb["Other"].defined_names["Local"].attr_text == "0.5"
    b2 = xlsx.load(out)
    assert {n.name for n in names.list_names(b2)} == {"Amounts", "Local"}


def test_sheet_scoped_names_follow_structural_edits(setup):
    ws = setup
    b, s = ws.book, ws.sheet
    from termse.model import structure

    other = structure.add_sheet(b, "Other")
    names.set_name(b, "Top", "Sheet!$A$3", s, s)
    names.set_name(b, "Far", "Sheet!$A$2", other, other)
    ws["C1"] = "=Top"
    b.set_input(other, 1, 1, "=Far")
    structure.insert(b, s, "row", 1, 2)
    assert s.ws.defined_names["Top"].attr_text == "Sheet!$A$5"
    assert other.ws.defined_names["Far"].attr_text == "Sheet!$A$4"
    assert ws["C3"] == 20
    assert b.get_value(other, 1, 1) == 10
    b.do_undo()
    assert s.ws.defined_names["Top"].attr_text == "Sheet!$A$3"
    assert ws["C1"] == 20
    structure.rename_sheet(b, s, "Data")
    assert s.ws.defined_names["Top"].attr_text == "Data!$A$3"
    assert other.ws.defined_names["Far"].attr_text == "Data!$A$2"
    assert ws["C1"] == 20
    assert b.get_value(other, 1, 1) == 10
