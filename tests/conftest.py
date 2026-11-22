import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from termse.io.xlsx import new_workbook  # noqa: E402
from termse.model.refs import parse_addr  # noqa: E402


class Sheetish:
    """Tiny helper to drive a workbook in tests: s["A1"] = "=1+1"; s["A1"] -> 2"""

    def __init__(self, book, sheet=None):
        self.book = book
        self.sheet = sheet or book.sheets[0]

    def __setitem__(self, a1, text):
        r, c = parse_addr(a1)
        self.book.set_input(self.sheet, r, c, str(text))

    def __getitem__(self, a1):
        r, c = parse_addr(a1)
        return self.book.get_value(self.sheet, r, c)

    def text(self, a1):
        r, c = parse_addr(a1)
        return self.sheet.input_text(r, c)

    def fmt(self, a1):
        r, c = parse_addr(a1)
        return self.sheet.number_format(r, c)


@pytest.fixture(autouse=True)
def _isolated_settings(tmp_path_factory, monkeypatch):
    """Never read or write the real ~/.config/tse/settings.json from tests."""
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path_factory.mktemp("config")))


@pytest.fixture
def ws():
    return Sheetish(new_workbook())


@pytest.fixture
def ev(ws):
    """Evaluate a single formula in a scratch cell (Z1000)."""

    def run(formula, **cells):
        for k, v in cells.items():
            ws[k] = v
        ws["Z1000"] = formula
        return ws["Z1000"]

    return run


@pytest.fixture
def sample_path(tmp_path):
    from tests.fixtures.make_sample import build

    p = tmp_path / "sample.xlsx"
    build(str(p))
    return str(p)
