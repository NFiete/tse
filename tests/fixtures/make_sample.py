"""Build tests/fixtures/sample.xlsx: a workbook exercising many Excel features."""
from __future__ import annotations

import datetime as dt
import os
import sys

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation


def build(path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Budget"
    ws["A1"] = "Household Budget 2026"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
    ws.merge_cells("A1:E1")
    ws["A1"].alignment = Alignment(horizontal="center")
    headers = ["Category", "Jan", "Feb", "Mar", "Total"]
    head_fill = PatternFill("solid", fgColor="1F4E78")
    thin = Side(style="thin", color="999999")
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.border = Border(bottom=thin)
        c.alignment = Alignment(horizontal="center")
    rows = [("Rent", 1500, 1500, 1500), ("Groceries", 420.5, 388.25, 450.1), ("Utilities", 180, 165.4, 172),
            ("Transport", 95, 110, 87.5), ("Fun", 200, 150, 320)]
    for r, (name, *vals) in enumerate(rows, 4):
        ws.cell(row=r, column=1, value=name)
        for j, v in enumerate(vals, 2):
            c = ws.cell(row=r, column=j, value=v)
            c.number_format = '"$"#,##0.00'
        t = ws.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})")
        t.number_format = '"$"#,##0.00'
        t.font = Font(bold=True)
    tr = 4 + len(rows)
    ws.cell(row=tr, column=1, value="Total").font = Font(bold=True)
    for j in range(2, 6):
        L = openpyxl.utils.get_column_letter(j)
        c = ws.cell(row=tr, column=j, value=f"=SUM({L}4:{L}{tr - 1})")
        c.number_format = '"$"#,##0.00'
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFF2CC")
    ws.cell(row=tr + 2, column=1, value="Average month")
    ws.cell(row=tr + 2, column=2, value=f"=AVERAGE(B{tr}:D{tr})").number_format = '"$"#,##0.00'
    ws.cell(row=tr + 3, column=1, value="Biggest expense")
    ws.cell(row=tr + 3, column=2, value=f'=INDEX(A4:A{tr - 1},MATCH(MAX(E4:E{tr - 1}),E4:E{tr - 1},0))')
    ws.cell(row=tr + 4, column=1, value="Share of rent")
    ws.cell(row=tr + 4, column=2, value=f"=E4/E{tr}").number_format = "0.0%"
    ws.cell(row=tr + 5, column=1, value="Updated")
    ws.cell(row=tr + 5, column=2, value=dt.datetime(2026, 3, 31)).number_format = "mmm d, yyyy"
    ws.column_dimensions["A"].width = 18
    for L in "BCDE":
        ws.column_dimensions[L].width = 12
    ws.freeze_panes = "B4"
    ws["A4"].comment = Comment("Paid on the 1st", "tester")
    ws.conditional_formatting.add(f"E4:E{tr - 1}", CellIsRule(operator="greaterThan", formula=["1000"],
                                                              fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.conditional_formatting.add(f"B4:D{tr - 1}", ColorScaleRule(start_type="min", start_color="FFFFFF",
                                                                  end_type="max", end_color="63BE7B"))
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="10000")
    dv.add(f"B4:D{tr - 1}")
    ws.add_data_validation(dv)

    ws2 = wb.create_sheet("Data")
    ws2.append(["Region", "Product", "Units", "Price", "Revenue", "Date"])
    data = [
        ("North", "Widget", 10, 2.5, dt.date(2026, 1, 5)), ("South", "Gadget", 4, 10, dt.date(2026, 1, 9)),
        ("North", "Gadget", 7, 10, dt.date(2026, 2, 1)), ("East", "Widget", 12, 2.5, dt.date(2026, 2, 14)),
        ("West", "Gizmo", 3, 25, dt.date(2026, 3, 3)), ("South", "Widget", 8, 2.5, dt.date(2026, 3, 20)),
    ]
    for i, (reg, prod, units, price, d) in enumerate(data, 2):
        ws2.append([reg, prod, units, price, f"=C{i}*D{i}", d])
        ws2.cell(row=i, column=6).number_format = "yyyy-mm-dd"
    n = len(data) + 1
    ws2["H1"] = "North revenue"
    ws2["I1"] = f'=SUMIFS(E2:E{n},A2:A{n},"North")'
    ws2["H2"] = "Widget units"
    ws2["I2"] = f'=SUMIF(B2:B{n},"Widget",C2:C{n})'
    ws2["H3"] = "Lookup Gizmo"
    ws2["I3"] = f'=VLOOKUP("Gizmo",B2:E{n},4,FALSE)'
    ws2["H4"] = "Regions"
    ws2["I4"] = f"=COUNTA(A2:A{n})"
    ws2["H5"] = "Cross-sheet"
    ws2["I5"] = "=Budget!E9*2"
    ws2["H6"] = "Named total"
    ws2["I6"] = "=SUM(Revenue)"
    ws2.row_dimensions[n + 2].hidden = True
    wb.defined_names["Revenue"] = DefinedName("Revenue", attr_text=f"Data!$E$2:$E${n}")
    ws2.column_dimensions["H"].width = 16

    ws3 = wb.create_sheet("Notes")
    ws3["A1"] = "Long text that should overflow into the empty cells to its right in the grid"
    ws3["A3"] = "=TODAY()"
    ws3["A3"].number_format = "yyyy-mm-dd"
    ws3["A4"] = '=TEXT(1234.5,"#,##0.0")'
    ws3["A5"] = "=1/0"
    ws3["A6"] = "=IFERROR(A5,\"div by zero\")"
    ws3["A7"] = True
    ws3["A8"] = "https://example.com"
    ws3["A8"].hyperlink = "https://example.com"
    ws3.sheet_properties.tabColor = "70AD47"
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "sample.xlsx"))
