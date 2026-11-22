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
    book.mark_legacy()
    _load_cached_values(book, path)
    book.adopt_arrays(_dynamic_anchors(book, path))
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


_CM_CELL_RE = re.compile(rb'<(?:\w+:)?c\b([^>]*\bcm="\d+"[^>]*)>')


def _dynamic_anchors(book: Workbook, path: str) -> dict:
    """Array formulas that Excel marked as dynamic arrays (cm="..." on the cell, pointing at
    XLDAPR cell metadata): {sheet: {(r, c), ...}}. The others are legacy Ctrl+Shift+Enter arrays."""
    from openpyxl.worksheet.formula import ArrayFormula

    from ..model.refs import col_index

    wanted = {}
    for s in book.sheets:
        arrays = {rc for rc, cell in s.ws._cells.items() if isinstance(cell._value, ArrayFormula)}
        if arrays:
            wanted[s] = arrays
    if not wanted:
        return {}
    out: dict = {}
    try:
        with zipfile.ZipFile(path) as z:
            parts = _sheet_parts(z)
            for s, arrays in wanted.items():
                part = parts.get(s.name)
                if part is None:
                    continue
                for m in _CM_CELL_RE.finditer(z.read(part)):
                    rm = _ATTR_R_RE.search(m.group(1))
                    if rm is not None:
                        rc = (int(rm.group(2)), col_index(rm.group(1).decode()))
                        if rc in arrays:
                            out.setdefault(s, set()).add(rc)
    except Exception:
        pass
    return out


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
    fd, tmp = tempfile.mkstemp(prefix=".termse-", suffix=ext, dir=directory)
    os.close(fd)
    try:
        restore, dynamic = _materialize_arrays(book)
        try:
            wb.save(tmp)
        finally:
            _restore_cells(restore)
        _inject_cached_values(tmp, book, dynamic)
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
    book.undo.mark_saved()


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


def _set_plain(cell, v):
    if isinstance(v, ExcelError):
        cell._value, cell.data_type = v.code, "e"
    elif isinstance(v, bool):
        cell._value, cell.data_type = v, "b"
    elif isinstance(v, (int, float)):
        cell._value, cell.data_type = v, "n"
    else:
        cell._value, cell.data_type = str(v), "s"


def _materialize_arrays(book: Workbook):
    """Temporarily put spilling formulas and their results into the openpyxl cells the way Excel
    stores them: the formula cell holds an array formula over the spill range, and the other
    cells hold the values. Returns (undo list for _restore_cells, {sheet name: {anchor address}}
    of dynamic-array formulas, which _inject_cached_values marks with cm="1")."""
    from openpyxl.worksheet.formula import ArrayFormula

    from ..model.refs import addr

    restore: list = []
    dynamic: dict = {}
    for s in book.sheets:
        ws = s.ws
        cells = ws._cells
        sp = s.spill
        for (r, c), f in s.formulas.items():
            if f.legacy or f.array:
                continue
            if (r, c) in sp.anchors:
                r1, c1, r2, c2 = sp.rect((r, c))
            elif (r, c) in sp.blocked or f.arrayish:
                r2, c2 = r, c
            else:
                continue
            cell = cells.get((r, c))
            if cell is None or not isinstance(cell._value, str):
                continue
            ref = addr(r, c) if (r2, c2) == (r, c) else f"{addr(r, c)}:{addr(r2, c2)}"
            restore.append((ws, (r, c), cell, cell._value, cell.data_type))
            cell._value = ArrayFormula(ref, cell._value)
            dynamic.setdefault(s.name, set()).add(addr(r, c))
        for (ar, ac), rows in sp.anchors.items():
            for i, row in enumerate(rows):
                for j, v in enumerate(row):
                    if v is None or (i == 0 and j == 0):
                        continue
                    rc = (ar + i, ac + j)
                    cell = cells.get(rc)
                    if cell is not None and cell._value is not None:
                        continue
                    created = cell is None
                    if created:
                        cell = ws.cell(row=rc[0], column=rc[1])
                    restore.append((ws, rc, None if created else cell, None, cell.data_type))
                    _set_plain(cell, v)
    return restore, dynamic


def _restore_cells(restore: list):
    for ws, rc, cell, value, data_type in reversed(restore):
        if cell is None:
            ws._cells.pop(rc, None)
        else:
            cell._value, cell.data_type = value, data_type


_METADATA_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    b'<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" '
    b'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
    b'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    b'<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    b'<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    b'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
    b'<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>')
_METADATA_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
_METADATA_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _metadata_parts(zin: zipfile.ZipFile) -> dict[str, bytes]:
    """The parts to add or change so that cm="1" marks dynamic-array formulas (as Excel writes)."""
    from lxml import etree

    names = set(zin.namelist())
    if "xl/metadata.xml" in names:
        return {}
    out = {"xl/metadata.xml": _METADATA_XML}
    rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    ids = {rel.get("Id") for rel in rels}
    n = 1
    while f"rId{n}" in ids:
        n += 1
    rel = etree.SubElement(rels, f"{{{_PKG_RNS}}}Relationship")
    rel.set("Id", f"rId{n}")
    rel.set("Type", _METADATA_REL)
    rel.set("Target", "metadata.xml")
    out["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8",
                                                       standalone=True)
    ct = etree.fromstring(zin.read("[Content_Types].xml"))
    over = etree.SubElement(ct, f"{{{_CT_NS}}}Override")
    over.set("PartName", "/xl/metadata.xml")
    over.set("ContentType", _METADATA_CT)
    out["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    return out


def _inject_cached_values(path: str, book: Workbook, dynamic: dict | None = None):
    """openpyxl writes formulas without results; add our computed values so other
    readers (LibreOffice, Quick Look, pandas...) show numbers immediately. Also mark
    dynamic-array formulas ({sheet name: {address}}) with cm="1"."""
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
            dyn = (dynamic or {}).get(s.name, ())
            for c in root.iter(f"{{{_NS}}}c"):
                f = c.find(f"{{{_NS}}}f")
                if f is None:
                    continue
                if dyn and c.get("r") in dyn:
                    c.set("cm", "1")
                    changed = True
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
        if dynamic:
            replacements.update(_metadata_parts(zin))
        if not replacements:
            return
        tmp = path + ".v"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = replacements.pop(item.filename, None)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
            for name, data in replacements.items():  # new parts
                zout.writestr(name, data)
    os.replace(tmp, path)
