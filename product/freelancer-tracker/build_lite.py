"""Build the free LITE edition: entry sheets with a single totals line. Usage: python3 build_lite.py OUT.xlsx [--sample]"""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
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
    ("3.  The totals line at the top of each sheet adds everything up.", None),
    ("", None),
    ("LITE LIMITS", Font(bold=True, size=12)),
    (f"•  {N} rows each for income and expenses, and one total per sheet.", None),
    ("•  No monthly breakdown, no tax set-aside estimate, no goals, no expense categories.", None),
    ("", None),
    ("FULL VERSION adds: monthly summary for the whole year, a tax set-aside amount from a percentage you choose,", Font(bold=True)),
    ("monthly profit goals, editable expense categories, 500 rows each, and reuse every year.", Font(bold=True)),
    ("", None),
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

inc = sheet("Income", ["Date", "Client / Source", "Description", "Amount received", "Fees deducted at source", "Net income"],
            (13, 24, 34, 16, 16, 14), 6, "Total net income:")
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
    inc.cell(r, 1).number_format = "yyyy-mm-dd"
    inc.cell(r, 4).number_format = MONEY; inc.cell(r, 5).number_format = MONEY
    inc.cell(r, 6, f'=IF(D{r}="","",D{r}-N(E{r}))').number_format = MONEY
    for col in (1, 2, 3, 4):
        exp.cell(r, col).fill = inp
    exp.cell(r, 1).number_format = "yyyy-mm-dd"; exp.cell(r, 4).number_format = MONEY
dvd.add(f"A{HDR+1}:A{last}")

if sample:
    y = 2026
    for i, row in enumerate([(date(y,1,10),"Client A","Website copy",1200,0),(date(y,1,25),"Client B","Logo design",800,24),
                             (date(y,2,8),"Client A","Retainer",1500,0)], start=HDR+1):
        for j, v in enumerate(row, start=1): inc.cell(i, j, v)
    for i, row in enumerate([(date(y,1,3),"Adobe","Creative suite",60),(date(y,2,14),"Cowork Space","Day passes",90)], start=HDR+1):
        for j, v in enumerate(row, start=1): exp.cell(i, j, v)

wb.save(out)
print("saved", out, wb.sheetnames)
