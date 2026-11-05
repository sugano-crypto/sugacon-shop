"""Build the free LITE edition: entry sheets with a single totals line. Usage: python3 build_lite.py OUT.xlsx [--sample]"""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Protection
from openpyxl.workbook.protection import WorkbookProtection
import os
from openpyxl.worksheet.datavalidation import DataValidation

out = sys.argv[1]
sample = "--sample" in sys.argv
N = 50
NAVY, INPUT_BG = "1F3A5F", "FFF8DC"
hf, hfill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=NAVY)
inp = PatternFill("solid", fgColor=INPUT_BG)
MONEY = "#,##0.00"
HDR = 3  # header row; data starts row 4

wb = Workbook()
ws = wb.active
ws.title = "Start Here"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 100
lines = [
    ("Freelancer Income & Expense Tracker - LITE (free)", Font(bold=True, size=18, color=NAVY)),
    ("A simple place to log what you earn and what you spend.", Font(italic=True)),
    ("", None),
    ("HOW TO USE", Font(bold=True, size=12)),
    ("1.  Income: add one row per payment received (yellow cells).", None),
    ("2.  Expenses: add one row per cost (yellow cells).", None),
    ("3.  Summary: see total income, total expenses and your profit (income after fees minus expenses).", None),
    ("", None),
    ("LITE LIMITS", Font(bold=True, size=12)),
    (f"•  {N} rows each for income and expenses, plus one overall summary.", None),
    ("•  No monthly breakdown, profit margin, average payment, tax set-aside estimate, goals or expense categories.", None),
    ("", None),
    ("FULL VERSION adds: a month-by-month summary for the whole year with profit margin, payment count and average payment,", Font(bold=True)),
    ("a tax set-aside amount from a percentage you choose, monthly profit goals, editable categories, 500 rows each, and reuse every year.", Font(bold=True)),
    ("", None),
    ("The sheets are locked (except the yellow input cells) to protect the formulas. Please do not copy or redistribute this file.", None),
    ("This is a bookkeeping aid, not tax, legal or accounting advice. Works in Excel, LibreOffice and Google Sheets. No macros.", Font(italic=True)),
]
for i, (t, f) in enumerate(lines, start=2):
    c = ws.cell(i, 2, t)
    if f: c.font = f
    c.alignment = Alignment(wrap_text=True, vertical="top")

def sheet(name, cols, widths, total_col_idx, total_label):
    s = wb.create_sheet(name)
    s["A1"] = name
    s["A1"].font = Font(bold=True, size=14, color=NAVY)
    s["A2"] = total_label
    s["A2"].font = Font(bold=True)
    for i, h in enumerate(cols, start=1):
        c = s.cell(HDR, i, h); c.font, c.fill = hf, hfill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for col, w in zip("ABCDEFG", widths):
        s.column_dimensions[col].width = w
    s.freeze_panes = f"A{HDR+1}"
    return s

inc = sheet("Income", ["Date", "Client / Source", "Description", "Amount received", "Fees deducted at source", "Income after fees"],
            (13, 24, 34, 16, 16, 14), 6, "Total income after fees:")
exp = sheet("Expenses", ["Date", "Vendor", "Description", "Amount"], (13, 24, 34, 14), 4, "Total expenses:")

last = HDR + N
inc["D2"] = f"=SUM(F{HDR+1}:F{last})"; inc["D2"].number_format = MONEY; inc["D2"].font = Font(bold=True)
exp["D2"] = f"=SUM(D{HDR+1}:D{last})"; exp["D2"].number_format = MONEY; exp["D2"].font = Font(bold=True)

dvd = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                     showErrorMessage=True, error="Enter a valid date.")
inc.add_data_validation(dvd); exp.add_data_validation(dvd)
for r in range(HDR + 1, last + 1):
    for col in (1, 2, 3, 4, 5):
        inc.cell(r, col).fill = inp
        inc.cell(r, col).protection = Protection(locked=False)
    inc.cell(r, 1).number_format = "yyyy-mm-dd"
    inc.cell(r, 4).number_format = MONEY; inc.cell(r, 5).number_format = MONEY
    inc.cell(r, 6, f'=IF(D{r}="","",D{r}-N(E{r}))').number_format = MONEY
    for col in (1, 2, 3, 4):
        exp.cell(r, col).fill = inp
        exp.cell(r, col).protection = Protection(locked=False)
    exp.cell(r, 1).number_format = "yyyy-mm-dd"; exp.cell(r, 4).number_format = MONEY
dvd.add(f"A{HDR+1}:A{last}")

if sample:
    y = 2026
    for i, row in enumerate([(date(y,1,10),"Client A","Website copy",1200,0),(date(y,1,25),"Client B","Logo design",800,24),
                             (date(y,2,8),"Client A","Retainer",1500,0)], start=HDR+1):
        for j, v in enumerate(row, start=1): inc.cell(i, j, v)
    for i, row in enumerate([(date(y,1,3),"Adobe","Creative suite",60),(date(y,2,14),"Cowork Space","Day passes",90)], start=HDR+1):
        for j, v in enumerate(row, start=1): exp.cell(i, j, v)

sm = wb.create_sheet("Summary", 1)
sm.sheet_view.showGridLines = False
sm.column_dimensions["A"].width = 34; sm.column_dimensions["B"].width = 18
sm["A1"] = "Summary"; sm["A1"].font = Font(bold=True, size=14, color=NAVY)
rows = [("Total income after fees", "=Income!D2"), ("Total expenses", "=Expenses!D2"), ("Profit", "=B3-B4")]
for i, (a, f) in enumerate(rows, start=3):
    sm.cell(i, 1, a); c = sm.cell(i, 2, f); c.number_format = MONEY
sm["A5"].font = Font(bold=True); sm["B5"].font = Font(bold=True)
sm["A7"] = "Month-by-month breakdown, profit margin, average payment and tax set-aside are in the full version."
sm["A7"].font = Font(italic=True, size=9)

pw = os.environ.get("LITE_SHEET_PASSWORD", "")
for sh in wb.worksheets:
    sh.protection.sheet = True
    sh.protection.selectLockedCells = True  # True = NOT allowed: formula cells cannot be selected or copied
    if pw: sh.protection.set_password(pw)
wb.security = WorkbookProtection(lockStructure=True)
if pw: wb.security.set_workbook_password(pw)
wb.save(out)
print("saved", out, wb.sheetnames)
