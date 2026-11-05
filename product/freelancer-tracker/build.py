"""Build the Freelancer Income & Tax Set-Aside Tracker (.xlsx). Usage: python3 build.py OUT.xlsx [--sample]"""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.table import Table, TableStyleInfo

out = sys.argv[1]
sample = "--sample" in sys.argv
N = 500  # rows available for entries

NAVY, INPUT_BG, GREY = "1F3A5F", "FFF8DC", "F2F2F2"
hdr_font = Font(bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor=NAVY)
inp_fill = PatternFill("solid", fgColor=INPUT_BG)
thin = Side(style="thin", color="BBBBBB")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
MONEY = '#,##0.00'

wb = Workbook()

# ---------- Start Here ----------
ws = wb.active
ws.title = "Start Here"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 100
lines = [
    ("Freelancer Income & Tax Set-Aside Tracker", Font(bold=True, size=18, color=NAVY)),
    ("Track income and expenses, see monthly profit, and know how much to put aside for taxes.", Font(italic=True)),
    ("", None),
    ("HOW TO USE (3 steps)", Font(bold=True, size=12)),
    ("1.  Settings: enter your tax year, currency symbol and the % of profit you want to set aside.", None),
    ("2.  Income and Expenses: add one row per payment or cost. Yellow cells are inputs; everything else calculates itself.", None),
    ("3.  Monthly Summary: see income, expenses, profit and the amount to set aside, month by month and for the year.", None),
    ("", None),
    ("GOOD TO KNOW", Font(bold=True, size=12)),
    (f"•  Room for {N} income rows and {N} expense rows. Insert rows inside the table if you need more.", None),
    ("•  Dates outside your chosen tax year are ignored in the summary, so you can reuse the file each year.", None),
    ("•  Works in Excel, LibreOffice and Google Sheets (File > Import > Upload). No macros.", None),
    ("•  Delete the sample rows (if present) before you start; the formulas stay intact.", None),
    ("", None),
    ("IMPORTANT", Font(bold=True, size=12, color="C00000")),
    ("This is a bookkeeping aid, not tax, legal or accounting advice. The set-aside rate is a rough planning figure that YOU choose;", None),
    ("actual tax depends on your country, income, deductions and circumstances. Ask a qualified tax professional for advice.", None),
]
for i, (t, f) in enumerate(lines, start=2):
    c = ws.cell(row=i, column=2, value=t)
    if f: c.font = f
    c.alignment = Alignment(wrap_text=True, vertical="top")

# ---------- Settings ----------
st = wb.create_sheet("Settings")
st.column_dimensions["A"].width = 34
st.column_dimensions["B"].width = 16
st.column_dimensions["C"].width = 60
rows = [
    ("Tax year", 2026, "Four-digit year. Only entries dated in this year are summarised."),
    ("Currency symbol", "$", "Display label only; amounts are not converted."),
    ("Tax set-aside rate (% of profit)", 0.25, "Your own planning rate. Example: 25% means 25 of every 100 profit is set aside."),
    ("Monthly profit goal", 3000, "Optional. Months at or above goal are highlighted green."),
]
st["A1"], st["B1"], st["C1"] = "Setting", "Value", "Note"
for c in st[1]:
    c.font, c.fill = hdr_font, hdr_fill
for r, (a, b, n) in enumerate(rows, start=2):
    st.cell(r, 1, a)
    v = st.cell(r, 2, b); v.fill = inp_fill; v.border = box
    st.cell(r, 3, n).alignment = Alignment(wrap_text=True)
st["B4"].number_format = "0.0%"
st["B5"].number_format = MONEY
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                    showErrorMessage=True, errorTitle="Rate", error="Enter a value between 0% and 100%.")
st.add_data_validation(dv); dv.add("B4")
dvy = DataValidation(type="whole", operator="between", formula1="2000", formula2="2100",
                     showErrorMessage=True, error="Enter a 4-digit year.")
st.add_data_validation(dvy); dvy.add("B2")

# ---------- Income ----------
inc = wb.create_sheet("Income")
inc_cols = ["Date", "Client / Source", "Description", "Amount received", "Fees deducted at source", "Income after fees", "Month"]
for i, h in enumerate(inc_cols, start=1):
    c = inc.cell(1, i, h); c.font, c.fill = hdr_font, hdr_fill
    c.alignment = Alignment(wrap_text=True, vertical="center")
for col, w in zip("ABCDEFG", (13, 24, 34, 16, 16, 14, 9)):
    inc.column_dimensions[col].width = w
inc.freeze_panes = "A2"

# ---------- Expenses ----------
exp = wb.create_sheet("Expenses")
exp_cols = ["Date", "Vendor", "Category", "Description", "Amount", "Month"]
for i, h in enumerate(exp_cols, start=1):
    c = exp.cell(1, i, h); c.font, c.fill = hdr_font, hdr_fill
for col, w in zip("ABCDEF", (13, 24, 20, 34, 14, 9)):
    exp.column_dimensions[col].width = w
exp.freeze_panes = "A2"

CATS = ["Software & subscriptions", "Equipment", "Office & coworking", "Marketing", "Travel",
        "Professional services", "Education", "Fees & bank charges", "Other"]
cat = wb.create_sheet("Lists")
cat["A1"] = "Expense categories (edit freely)"; cat["A1"].font = Font(bold=True)
for i, c_ in enumerate(CATS, start=2):
    cat.cell(i, 1, c_)
cat.column_dimensions["A"].width = 32
dvc = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${len(CATS)+20}", allow_blank=True)
exp.add_data_validation(dvc)
dvd = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                     showErrorMessage=True, error="Enter a valid date.")
inc.add_data_validation(dvd); exp.add_data_validation(dvd)

for r in range(2, N + 2):
    for col in (1, 2, 3, 4, 5):
        inc.cell(r, col).fill = inp_fill
    inc.cell(r, 1).number_format = "yyyy-mm-dd"
    inc.cell(r, 4).number_format = MONEY
    inc.cell(r, 5).number_format = MONEY
    inc.cell(r, 6, f'=IF(D{r}="","",D{r}-N(E{r}))').number_format = MONEY
    inc.cell(r, 7, f'=IF(A{r}="","",IF(YEAR(A{r})=Settings!$B$2,MONTH(A{r}),""))')
    for col in (1, 2, 3, 4, 5):
        exp.cell(r, col).fill = inp_fill
    exp.cell(r, 1).number_format = "yyyy-mm-dd"
    exp.cell(r, 5).number_format = MONEY
    exp.cell(r, 6, f'=IF(A{r}="","",IF(YEAR(A{r})=Settings!$B$2,MONTH(A{r}),""))')
dvd.add(f"A2:A{N+1}"); dvc.add(f"C2:C{N+1}")

# ---------- Monthly Summary ----------
ms = wb.create_sheet("Monthly Summary")
ms.sheet_view.showGridLines = False
heads = ["Month", "Income after fees", "Expenses", "Profit", "Profit margin", "Payments (count)", "Avg. payment",
         "Set aside for tax", "Left after set-aside", "Goal met?"]
ms["A1"] = '="Monthly Summary — "&Settings!B2&"  ("&Settings!B3&")"'
ms["A1"].font = Font(bold=True, size=14, color=NAVY)
for i, h in enumerate(heads, start=1):
    c = ms.cell(3, i, h); c.font, c.fill = hdr_font, hdr_fill
    c.alignment = Alignment(horizontal="center", wrap_text=True)
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
for m in range(1, 13):
    r = 3 + m
    ms.cell(r, 1, months[m - 1])
    ms.cell(r, 2, f"=SUMIF(Income!$G$2:$G${N+1},{m},Income!$F$2:$F${N+1})")
    ms.cell(r, 3, f"=SUMIF(Expenses!$F$2:$F${N+1},{m},Expenses!$E$2:$E${N+1})")
    ms.cell(r, 4, f"=B{r}-C{r}")
    ms.cell(r, 5, f'=IF(B{r}=0,"",D{r}/B{r})')
    ms.cell(r, 6, f"=COUNTIF(Income!$G$2:$G${N+1},{m})")
    ms.cell(r, 7, f'=IF(F{r}=0,"",B{r}/F{r})')
    ms.cell(r, 8, f"=MAX(0,D{r})*Settings!$B$4")
    ms.cell(r, 9, f"=D{r}-H{r}")
    ms.cell(r, 10, f'=IF(Settings!$B$5<=0,"",IF(D{r}>=Settings!$B$5,"Yes","No"))')
    ms.cell(r, 10).alignment = Alignment(horizontal="center")
tr = 16
ms.cell(tr, 1, "Year total").font = Font(bold=True)
for col in "BCDFHI":
    ms[f"{col}{tr}"] = f"=SUM({col}4:{col}15)"
ms[f"E{tr}"] = f'=IF(B{tr}=0,"",D{tr}/B{tr})'
ms[f"G{tr}"] = f'=IF(F{tr}=0,"",B{tr}/F{tr})'
for col in "ABCDEFGHI":
    ms[f"{col}{tr}"].font = Font(bold=True)
for r in range(4, tr + 1):
    for col in (2, 3, 4, 7, 8, 9):
        ms.cell(r, col).number_format = MONEY
    ms.cell(r, 5).number_format = "0.0%"
    ms.cell(r, 6).number_format = "0"
    for col in range(1, 11):
        ms.cell(r, col).border = box
        if r == tr: ms.cell(r, col).fill = PatternFill("solid", fgColor=GREY)
ms.column_dimensions["A"].width = 14
for col in "BCDEFGHI": ms.column_dimensions[col].width = 16
ms.column_dimensions["J"].width = 12
ms.row_dimensions[3].height = 32
ms.conditional_formatting.add("J4:J15", CellIsRule(operator="equal", formula=['"Yes"'],
                              fill=PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")))
ms.conditional_formatting.add("D4:D16", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000")))
ms["A18"] = "Profit margin = profit / income after fees. Avg. payment = income after fees / number of payments. Set-aside is calculated only on positive monthly profit. Planning aid only, not tax advice."
ms["A18"].font = Font(italic=True, size=9)

# ---------- Sample data ----------
if sample:
    y = 2026
    income = [(date(y,1,10),"Client A","Website copy",1200,0),(date(y,1,25),"Client B","Logo design",800,24),
              (date(y,2,8),"Client A","Retainer",1500,0),(date(y,2,20),"Client C","Consulting",600,18),
              (date(y,3,5),"Client B","Brand kit",2000,60)]
    expen = [(date(y,1,3),"Adobe","Software & subscriptions","Creative suite",60),
             (date(y,2,14),"Cowork Space","Office & coworking","Day passes",90),
             (date(y,3,9),"Ads Co","Marketing","Portfolio promo",150)]
    for i, row in enumerate(income, start=2):
        for j, v in enumerate(row, start=1): inc.cell(i, j, v)
    for i, row in enumerate(expen, start=2):
        for j, v in enumerate(row, start=1): exp.cell(i, j, v)

wb.move_sheet("Monthly Summary", offset=-(len(wb.sheetnames) - 3))
wb.save(out)

# Suppress Excel's cosmetic "inconsistent formula" green triangles on the summary sheet.
import zipfile, shutil, re, os
tmp = out + ".tmp"
with zipfile.ZipFile(out) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    wbx = zin.read("xl/workbook.xml").decode()
    rels = zin.read("xl/_rels/workbook.xml.rels").decode()
    rid = re.search(r'<sheet [^>]*name="Monthly Summary"[^>]*r:id="(rId\d+)"', wbx).group(1)
    target = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    target = "xl/" + target.group(1).lstrip("/").replace("xl/", "")
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == target:
            x = data.decode()
            tag = '<ignoredErrors><ignoredError sqref="B4:J16" formula="1"/></ignoredErrors>'
            assert "</pageMargins>" in x or "<pageMargins" in x
            x = re.sub(r'(<pageMargins[^>]*/>)', r'\1' + tag, x, count=1)
            data = x.encode()
        zout.writestr(item, data)
shutil.move(tmp, out)
print("saved", out, wb.sheetnames)
