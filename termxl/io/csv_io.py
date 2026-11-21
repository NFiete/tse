"""CSV/TSV import and export."""
from __future__ import annotations

import csv
import os

import openpyxl

from ..format.numfmt import format_value
from ..formula.parser import FormulaSyntaxError
from ..model.inputparse import parse_input
from ..model.workbook import Workbook


def _sniff(path: str, sample: str) -> csv.Dialect | type[csv.Dialect]:
    if path.lower().endswith(".tsv"):
        return csv.excel_tab
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        return csv.excel


def load_csv(path: str) -> Workbook:
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        sample = fh.read(64 * 1024)
        fh.seek(0)
        dialect = _sniff(path, sample)
        reader = csv.reader(fh, dialect)
        wb = openpyxl.Workbook()
        wb.active.title = os.path.splitext(os.path.basename(path))[0][:31] or "Sheet1"
        book = Workbook(wb, path)
        sheet = book.sheets[0]
        for r, row in enumerate(reader, 1):
            for c, text in enumerate(row, 1):
                if text == "":
                    continue
                p = parse_input(text)
                try:
                    book._apply_parsed(sheet, r, c, p)
                except FormulaSyntaxError:
                    book.set_raw(sheet, r, c, text, "s")
    sheet.recompute_bounds()
    book.register_all()
    book.recalc_all()
    return book


def save_csv(book: Workbook, sheet, path: str):
    """Write the sheet's displayed values (as Excel does for CSV)."""
    delim = "\t" if path.lower().endswith(".tsv") else ","
    tmp = path + ".~tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh, delimiter=delim)
        for r in range(1, sheet.max_row + 1):
            row = []
            for c in range(1, sheet.max_col + 1):
                v = book.get_value(sheet, r, c)
                row.append("" if v is None else format_value(v, sheet.number_format(r, c), book.epoch1904)[0].strip())
            w.writerow(row)
    os.replace(tmp, path)
    book.modified = False
