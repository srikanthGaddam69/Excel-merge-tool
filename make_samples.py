#!/usr/bin/env python3
"""Regenerate the demo workbooks in this folder (needs: pip install openpyxl).

original.xlsx — a styled tracker: coloured header, frozen pane, filter,
                status colours, a formula column and a second sheet.
shared.xlsx   — a contributor's update with different column names,
                one blank cell, one extra key, one duplicated key
                and keys in mixed order.
"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
head_fill = PatternFill("solid", fgColor="1F4E78")
status_fill = {
    "Open": PatternFill("solid", fgColor="FCE4D6"),
    "In Progress": PatternFill("solid", fgColor="FFF2CC"),
    "Closed": PatternFill("solid", fgColor="E2EFDA"),
}

rows = [
    ("REQ-0001", "System shall log all user actions", "Software", "Open", "Test plan TP-01"),
    ("REQ-0002", "Login shall lock after 5 failed attempts", "Security", "In Progress", "Test plan TP-02"),
    ("REQ-0003", "Reports shall export to PDF", "Software", "Open", ""),
    ("REQ-0004", "Backup shall run nightly", "Operations", "Closed", "Ops checklist OC-07"),
    ("REQ-0005", "UI shall support dark mode", "Software", "Open", ""),
    ("REQ-0006", "Data shall be encrypted at rest", "Security", "Open", "Design note DN-12"),
    ("REQ-0007", "API response time under 300 ms", "Performance", "In Progress", ""),
    ("REQ-0008", "Audit trail retained for 7 years", "Compliance", "Open", "Policy POL-3"),
]


def make_original() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Tracker"
    headers = ["Reqt_ID", "Description", "Area", "Status", "Evidence", "Overall"]
    ws.append(headers)
    for c in ws[1]:
        c.fill = head_fill
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = box
    for i, r in enumerate(rows, start=2):
        ws.append(list(r) + [f'=IF(D{i}="Closed","Done","Pending")'])
        for c in ws[i]:
            c.border = box
            c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(i, 4).fill = status_fill[r[3]]
    for col, width in zip("ABCDEF", (12, 42, 14, 14, 40, 12)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:F{len(rows) + 1}"
    dv = DataValidation(type="list", formula1='"Open,In Progress,Closed"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"D2:D{len(rows) + 1}")

    notes = wb.create_sheet("Notes")
    notes["A1"] = "This sheet must come through the merge untouched."
    notes["A1"].font = Font(italic=True, color="7F7F7F")
    wb.save(HERE / "original.xlsx")


def make_shared() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Update"
    ws.append(["Requirement ID", "Owner", "New Status", "Test Report", "Comment"])
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.append(["REQ-0003", "A. Lee", "Closed", "Report RPT-301", "Passed first run"])
    ws.append(["REQ-0001", "B. Kumar", "In Progress", "Report RPT-298", ""])
    ws.append(["REQ-0006", "C. Ortiz", "", "Report RPT-305", "Status unchanged"])
    ws.append(["REQ-0004", "B. Kumar", "Closed", "Report RPT-290", ""])
    ws.append(["REQ-0099", "D. Smith", "Closed", "Report RPT-400", "Key not in original"])
    # Duplicate ID: same status (repeat is dropped), a second report (combined)
    ws.append(["REQ-0001", "B. Kumar", "In Progress", "Report RPT-299", "Re-test"])
    ws.column_dimensions["D"].width = 20
    wb.save(HERE / "shared.xlsx")


if __name__ == "__main__":
    make_original()
    make_shared()
    print("Wrote original.xlsx and shared.xlsx")
