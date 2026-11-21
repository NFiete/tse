"""Loading and saving workbooks (.xlsx/.xlsm/.xltx/.xltm, plus CSV/TSV)."""
from __future__ import annotations

import gc
import os
import re
import tempfile
import zipfile

import openpyxl

from ..formula.errors import ERROR_CODES, ExcelError
from ..model.dates import to_serial
from ..model.workbook import Workbook

EXCEL_EXTS = (".xlsx", ".xlsm", ".xltx", ".xltm")
TEXT_EXTS = (".csv", ".tsv", ".txt")


class FileFormatError(Exception):
    pass


def new_workbook() -> Workbook:
    wb = openpyxl.Workbook()
    book = Workbook(wb, None)
    book.register_all()
    return book


def load(path: str) -> Workbook:
    """Load a workbook. The cyclic garbage collector is paused while the (large) object
    graph is built and the result is frozen afterwards, which roughly halves load time."""
    gc.unfreeze()
    enabled = gc.isenabled()
    gc.disable()
    try:
        return _load(path)
    finally:
        if enabled:
            gc.enable()
        gc.freeze()


def _load(path: str) -> Workbook:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls" or ext == ".xlsb":
        raise FileFormatError(
            f"{ext} files are not supported. Convert first, e.g.:  soffice --headless --convert-to xlsx {path}")
    if ext in TEXT_EXTS:
        from .csv_io import load_csv

        return load_csv(path)
    if ext not in EXCEL_EXTS:
        raise FileFormatError(f"Unsupported file type: {ext or '(none)'}")
    keep_vba = ext in (".xlsm", ".xltm")
    try:
        wb = openpyxl.load_workbook(path, keep_vba=keep_vba, keep_links=True)
    except zipfile.BadZipFile as e:
        raise FileFormatError(f"Not a valid Excel file: {e}") from e
    book = Workbook(wb, path)
    book.register_all()
    _load_cached_values(book, path)
    book.initial_calc()
    book.warnings = detect_unpreservable(path)
    return book


def _norm_cached(value, data_type, epoch1904):
    if value is None:
        return None
    if data_type == "e" or (isinstance(value, str) and value in ERROR_CODES and data_type != "s"):
        return ERROR_CODES.get(str(value).upper(), ERROR_CODES["#VALUE!"])
    if isinstance(value, (int, float, str, bool)):
        return value
    try:
        return to_serial(value, epoch1904)
    except TypeError:
        return str(value)


_FORMULA_CELL_RE = re.compile(
    rb'<(?:\w+:)?c\b([^>]*)>\s*<(?:\w+:)?f\b[^>]*(?:/>|>[^<]*</(?:\w+:)?f>)\s*<(?:\w+:)?v>([^<]*)</(?:\w+:)?v>')
_ATTR_R_RE = re.compile(rb'\br="([A-Z]+)(\d+)"')
_ATTR_T_RE = re.compile(rb'\bt="(\w+)"')


def _load_cached_values(book: Workbook, path: str):
    """Read Excel's last computed results for formula cells straight from the sheet XML."""
    if not any(s.formulas for s in book.sheets):
        return
    try:
        with zipfile.ZipFile(path) as z:
            parts = _sheet_parts(z)
            names = set(z.namelist())
            for s in book.sheets:
                part = parts.get(s.name)
                if not s.formulas or part not in names:
                    continue
                _scan_cached(book, s, z.read(part))
    except Exception:  # fall back to openpyxl's (slower) reader
        _load_cached_values_openpyxl(book, path)


def _scan_cached(book: Workbook, s, data: bytes):
    from html import unescape

    from ..model.refs import col_index

    formulas = s.formulas
    for m in _FORMULA_CELL_RE.finditer(data):
        attrs = m.group(1)
        rm = _ATTR_R_RE.search(attrs)
        if rm is None:
            continue
        rc = (int(rm.group(2)), col_index(rm.group(1).decode()))
        f = formulas.get(rc)
        if f is None:
            continue
        raw = m.group(2).decode("utf-8")
        tm = _ATTR_T_RE.search(attrs)
        t = tm.group(1) if tm else b"n"
        if t == b"str" or t == b"inlineStr":
            v = unescape(raw)
        elif t == b"b":
            v = raw.strip() == "1"
        elif t == b"e":
            v = ERROR_CODES.get(raw.strip().upper(), ERROR_CODES["#VALUE!"])
        else:
            if raw == "":
                continue
            try:
                fv = float(raw)
            except ValueError:
                continue
            v = int(fv) if fv.is_integer() and abs(fv) < 2**53 else fv
        s.cached[rc] = (f.text, v)


def _load_cached_values_openpyxl(book: Workbook, path: str):
    try:
        ro = openpyxl.load_workbook(path, data_only=True, read_only=True, keep_links=False)
    except Exception:
        return
    try:
        for s in book.sheets:
            if not s.formulas or s.name not in ro.sheetnames:
                continue
            ws = ro[s.name]
            rows = [r for r, _ in s.formulas]
            for row in ws.iter_rows(min_row=min(rows), max_row=max(rows)):
                for cell in row:
                    v = getattr(cell, "value", None)
                    if v is None:
                        continue
                    rc = (cell.row, cell.column)
                    f = s.formulas.get(rc)
                    if f is not None:
                        s.cached[rc] = (f.text, _norm_cached(v, getattr(cell, "data_type", None), book.epoch1904))
    finally:
        ro.close()


def detect_unpreservable(path: str) -> list[str]:
    """Features in the file that openpyxl drops when re-saving."""
    found: list[str] = []
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()

            def has(prefix):
                return any(n.startswith(prefix) for n in names)

            if has("xl/slicers/") or has("xl/slicerCaches/"):
                found.append("slicers")
            if has("xl/timelines/"):
                found.append("timelines")
            if has("xl/ctrlProps/"):
                found.append("form controls")
            if has("xl/activeX/"):
                found.append("ActiveX controls")
            if has("xl/threadedComments/"):
                found.append("threaded comments (kept as plain notes)")
            if has("xl/richData/"):
                found.append("pictures in cells / linked data types")
            if has("xl/model/"):
                found.append("Power Pivot data model")
            if "xl/connections.xml" in names or has("xl/queryTables/"):
                found.append("data connections / queries")
            shapes = sparklines = x14cf = False
            for n in names:
                if n.startswith("xl/drawings/drawing") and n.endswith(".xml"):
                    data = z.read(n)
                    if b"<xdr:sp>" in data or b"<xdr:sp " in data or b"<xdr:grpSp" in data or b"<xdr:cxnSp" in data:
                        shapes = True
                elif n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                    if z.getinfo(n).file_size > 200_000_000:
                        continue
                    data = z.read(n)
                    if b"sparklineGroup" in data:
                        sparklines = True
                    if b"x14:conditionalFormatting" in data or b"x14:dataValidation" in data:
                        x14cf = True
            if shapes:
                found.append("shapes / text boxes")
            if sparklines:
                found.append("sparklines")
            if x14cf:
                found.append("some newer conditional formats / validations")
            if any(n.startswith("xl/charts/chart") for n in names):
                found.append("charts (kept, but some chart formatting may change)")
    except (zipfile.BadZipFile, OSError):
        pass
    return found


def save(book: Workbook, path: str, active_sheet=None, cursor=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in TEXT_EXTS:
        from .csv_io import save_csv

        save_csv(book, active_sheet or book.sheets[0], path)
        return
    if ext not in EXCEL_EXTS:
        raise FileFormatError(f"Can't save as {ext or '(no extension)'}; use .xlsx, .xlsm or .csv")
    wb = book.wb
    if ext in (".xlsx", ".xltx") and getattr(wb, "vba_archive", None) is not None:
        wb.vba_archive = None  # macros can't live in an .xlsx
    wb.template = ext in (".xltx", ".xltm")
    try:
        wb.calculation.fullCalcOnLoad = True
    except AttributeError:
        pass
    from ..model.outline import max_level

    for s in book.sheets:
        # Excel uses these to draw the outline bars. openpyxl doesn't maintain outlineLevelRow, and
        # writes outlineLevelCol from column_dimensions.max_outline before it has computed it.
        s.ws.sheet_format.outlineLevelRow = max_level(s, "row")
        s.ws.column_dimensions.max_outline = max_level(s, "col")
    if active_sheet is not None:
        idx = wb.worksheets.index(active_sheet.ws)
        wb.active = wb._sheets.index(wb.worksheets[idx])
        for ws in wb.worksheets:
            ws.sheet_view.tabSelected = ws is active_sheet.ws
        if cursor is not None:
            from ..model.refs import addr

            a = addr(*cursor)
            sel = active_sheet.ws.sheet_view.selection
            if sel:
                sel[0].activeCell = a
                sel[0].sqref = a
    directory = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(prefix=".termxl-", suffix=ext, dir=directory)
    os.close(fd)
    try:
        wb.save(tmp)
        _inject_cached_values(tmp, book)
        if os.path.exists(path):
            try:
                os.chmod(tmp, os.stat(path).st_mode & 0o7777)
            except OSError:
                pass
        else:
            os.chmod(tmp, 0o666 & ~_umask())
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    book.path = path
    book.modified = False


def _umask() -> int:
    m = os.umask(0)
    os.umask(m)
    return m


# --------------------------------------------------------------------------- cached values

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_RNS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _sheet_parts(z: zipfile.ZipFile) -> dict[str, str]:
    from lxml import etree

    wbx = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {}
    for rel in rels.iter(f"{{{_PKG_RNS}}}Relationship"):
        t = rel.get("Target")
        if t.startswith("/"):
            t = t[1:]
        elif not t.startswith("xl/"):
            t = "xl/" + t
        targets[rel.get("Id")] = t
    out = {}
    for sh in wbx.iter(f"{{{_NS}}}sheet"):
        rid = sh.get(f"{{{_RNS}}}id")
        if rid in targets:
            out[sh.get("name")] = targets[rid]
    return out


_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _inject_cached_values(path: str, book: Workbook):
    """openpyxl writes formulas without results; add our computed values so other
    readers (LibreOffice, Quick Look, pandas...) show numbers immediately."""
    if not any(s.formulas for s in book.sheets):
        return
    from lxml import etree

    from ..model.refs import col_index

    with zipfile.ZipFile(path) as zin:
        parts = _sheet_parts(zin)
        replacements: dict[str, bytes] = {}
        for s in book.sheets:
            part = parts.get(s.name)
            if part is None or not s.formulas or part not in zin.namelist():
                continue
            root = etree.fromstring(zin.read(part))
            changed = False
            for c in root.iter(f"{{{_NS}}}c"):
                f = c.find(f"{{{_NS}}}f")
                if f is None:
                    continue
                m = _CELL_RE.match(c.get("r", ""))
                if not m:
                    continue
                rc = (int(m.group(2)), col_index(m.group(1)))
                if rc not in s.formulas:
                    continue
                val = s.values.get(rc)
                v = c.find(f"{{{_NS}}}v")
                if v is None:
                    v = etree.SubElement(c, f"{{{_NS}}}v")
                if isinstance(val, ExcelError):
                    if val.code not in ERROR_CODES:
                        c.remove(v)
                        continue
                    c.set("t", "e")
                    v.text = val.code
                elif isinstance(val, bool):
                    c.set("t", "b")
                    v.text = "1" if val else "0"
                elif isinstance(val, (int, float)):
                    if "t" in c.attrib:
                        del c.attrib["t"]
                    v.text = repr(float(val)) if isinstance(val, float) else str(val)
                elif isinstance(val, str):
                    c.set("t", "str")
                    v.text = val
                else:
                    c.remove(v)
                    continue
                changed = True
            if changed:
                replacements[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        if not replacements:
            return
        tmp = path + ".v"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = replacements.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
    os.replace(tmp, path)
