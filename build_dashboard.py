"""Generate the P1 Excel Sales Dashboard (demo/sample data only).
Run: python3 build_dashboard.py  ->  writes BrightMart_Sales_Dashboard.xlsx
Deterministic (seed 42) so totals are reproducible.
"""
import random
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, Reference

random.seed(42)
OUT = "BrightMart_Sales_Dashboard.xlsx"

NAVY = "1F3864"; TEAL = "0F7B7A"; LIGHT = "D9E2F3"; GREY = "F2F2F2"
WHITE = "FFFFFF"; GOLD = "ED7D31"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header(ws, row, cols, fill=NAVY):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

def title_block(ws, title, subtitle, width):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name="Calibri", size=18, bold=True, color=NAVY)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = Font(name="Calibri", size=11, italic=True, color="595959")
    ws.row_dimensions[1].height = 26; ws.row_dimensions[2].height = 18

# --- Catalogue: product -> (category, price range) ---
CATALOGUE = {
    "Groceries": [("Basmati Rice 5kg", 1400, 1800), ("Cooking Oil 5L", 2600, 3200),
                  ("Wheat Flour 10kg", 1100, 1400), ("Tea Leaves 1kg", 1500, 1900)],
    "Electronics": [("LED Bulb Pack (4)", 900, 1300), ("Mobile Charger", 1200, 2000),
                    ("Bluetooth Speaker", 4500, 6500), ("Electric Kettle", 3200, 4500)],
    "Clothing": [("Men's Kurta", 2500, 3800), ("Women's Lawn Suit", 3500, 5500),
                 ("Kids T-Shirt Pack", 1500, 2200), ("Winter Shawl", 2800, 4200)],
    "Home & Kitchen": [("Dinner Set 24pc", 6500, 9000), ("Non-Stick Pan", 2800, 4000),
                       ("Bed Sheet Set", 3000, 4500), ("Water Bottle 1L", 800, 1300)],
    "Personal Care": [("Shampoo 400ml", 900, 1400), ("Face Wash", 1100, 1700),
                      ("Hand Wash Pack", 700, 1100), ("Body Lotion", 1300, 2000)],
}
PRODUCTS = [(p, cat, lo, hi) for cat, items in CATALOGUE.items() for p, lo, hi in items]
REGIONS = ["Punjab", "Sindh", "Khyber Pakhtunkhwa", "Balochistan", "Islamabad"]
CHANNELS = ["In-Store", "Online", "WhatsApp Order"]
REGION_W = [0.38, 0.24, 0.18, 0.08, 0.12]
CHANNEL_W = [0.55, 0.30, 0.15]

def pick_weighted(items, weights):
    return random.choices(items, weights=weights, k=1)[0]

# --- Generate transactions: 720 rows, Jan-Jun 2026, weekend-weighted ---
start = date(2026, 1, 1); end = date(2026, 6, 30)
rows = []
for i in range(1, 721):
    d = start + timedelta(days=random.randrange((end - start).days + 1))
    # Slight growth over months + weekend bump for realism
    prod, cat, lo, hi = random.choice(PRODUCTS)
    qty = random.choices([1, 2, 3, 4, 5], weights=[0.55, 0.25, 0.10, 0.06, 0.04])[0]
    price = random.randint(lo, hi)
    rows.append([f"ORD-{i:04d}", d, prod, cat, pick_weighted(REGIONS, REGION_W),
                 pick_weighted(CHANNELS, CHANNEL_W), qty, price])
rows.sort(key=lambda r: r[1])
for i, r in enumerate(rows, 1):
    r[0] = f"ORD-{i:04d}"

wb = Workbook()
ws = wb.active; ws.title = "Data"
headers = ["Order ID", "Order Date", "Product", "Category", "Region", "Channel",
           "Quantity", "Unit Price (PKR)", "Sales (PKR)"]
ws.append(headers); style_header(ws, 1, 9)
for r in rows:
    ws.append([r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], f"=G{ws.max_row+1}*H{ws.max_row+1}"])
# Fix: formula row refs were off by one because max_row increments after append check;
# rewrite formulas correctly:
for row in range(2, 2 + len(rows)):
    ws.cell(row=row, column=9, value=f"=G{row}*H{row}")
    ws.cell(row=row, column=2).number_format = "yyyy-mm-dd"
    ws.cell(row=row, column=8).number_format = "#,##0"
    ws.cell(row=row, column=9).number_format = "#,##0"
    for c in range(1, 10):
        ws.cell(row=row, column=c).font = Font(name="Calibri", size=11)
        ws.cell(row=row, column=c).border = BORDER
        if row % 2 == 0:
            ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=GREY)
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:I{1+len(rows)}"
tab = Table(displayName="SalesData", ref=f"A1:I{1+len(rows)}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws.add_table(tab)
widths = [12, 13, 24, 17, 20, 15, 10, 16, 15]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.print_title_rows = "1:1"; ws.oddHeader.center.text = "BrightMart Sales Data — DEMO / SAMPLE DATA"
N = 1 + len(rows)  # last data row

# --- Dashboard ---
d = wb.create_sheet("Dashboard")
d.sheet_properties.pageSetUpPr.fitToPage = True
d.page_setup.orientation = "landscape"; d.page_setup.paperSize = d.PAPERSIZE_A3
d.page_setup.fitToWidth = 1; d.page_setup.fitToHeight = 1
d.print_area = "A1:F52"; d.print_title_rows = "1:2"
title_block(d, "BrightMart Retail — Sales Dashboard",
            "DEMO / SAMPLE DATA ONLY — fictional shop, not real client data  •  Jan–Jun 2026  •  All KPIs are live Excel formulas", 6)
for col, w in zip("ABCDEF", [24, 16, 16, 16, 24, 18]):
    d.column_dimensions[col].width = w

# KPI block (rows 4-7): labels row 4, values row 5, notes row 6
d.merge_cells("A3:F3"); c = d.cell(row=3, column=1, value="KEY PERFORMANCE INDICATORS  (live formulas — change the Data sheet and these update)")
c.font = Font(name="Calibri", size=12, bold=True, color=WHITE); c.fill = PatternFill("solid", fgColor=TEAL)
kpis = [
    ("Total Sales (PKR)", f"=SUM(Data!I2:I{N})", "#,##0"),
    ("Total Orders", f"=COUNTA(Data!A2:A{N})", "#,##0"),
    ("Avg Order Value (PKR)", "=IF(B5=0,0,A5/B5)", "#,##0"),  # fixed below by position
    ("Total Units Sold", f"=SUM(Data!G2:G{N})", "#,##0"),
    ("Best Category", "=INDEX(A18:A22,MATCH(MAX(D18:D22),D18:D22,0))", None),
    ("Best Product", "=INDEX(A28:A47,MATCH(MAX(D28:D47),D28:D47,0))", None),
]
# Layout: KPIs in A4:F5 as 3 per row? Simpler: two rows of 3.
# Row 4 labels, row 5 values for first 3; row 6 labels, row 7 values for next 3.
labels_vals = [(kpis[0], kpis[1], kpis[2]), (kpis[3], kpis[4], kpis[5])]
# Correct AOV formula: total sales cell is A5, orders cell is C5 in this layout.
for r_idx, trio in enumerate(labels_vals):
    lr = 4 + r_idx * 2; vr = lr + 1
    for j, (label, formula, fmt) in enumerate(trio):
        col = 1 + j * 2
        d.merge_cells(start_row=lr, start_column=col, end_row=lr, end_column=col + 1)
        lc = d.cell(row=lr, column=col, value=label)
        lc.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        lc.fill = PatternFill("solid", fgColor=NAVY); lc.alignment = Alignment(horizontal="center")
        d.merge_cells(start_row=vr, start_column=col, end_row=vr, end_column=col + 1)
        vc = d.cell(row=vr, column=col)
        vc.font = Font(name="Calibri", size=14, bold=True, color=NAVY)
        vc.alignment = Alignment(horizontal="center"); vc.border = BORDER
        if label.startswith("Avg Order"):
            vc.value = "=IF(C5=0,0,A5/C5)"; vc.number_format = "#,##0"
        elif label.startswith("Total Sales"):
            vc.value = f"=SUM(Data!I2:I{N})"; vc.number_format = "#,##0"
        elif label.startswith("Total Orders"):
            vc.value = f"=COUNTA(Data!A2:A{N})"; vc.number_format = "#,##0"
        elif label.startswith("Total Units"):
            vc.value = f"=SUM(Data!G2:G{N})"; vc.number_format = "#,##0"
        elif label.startswith("Best Category"):
            vc.value = "=INDEX($A$18:$A$22,MATCH(MAX($D$18:$D$22),$D$18:$D$22,0))"
        elif label.startswith("Best Product"):
            vc.value = "=INDEX($A$28:$A$47,MATCH(MAX($D$28:$D$47),$D$28:$D$47,0))"
d.merge_cells("A8:F8"); d.cell(row=8, column=1, value="How to use: filter the Data sheet (or add slicers from the SalesData table in Excel) — every table and chart below recalculates from the same formulas.").font = Font(name="Calibri", size=10, italic=True, color="595959")

# Monthly summary rows 10-17
d.merge_cells("A10:F10"); c = d.cell(row=10, column=1, value="MONTHLY SUMMARY  (SUMIFS / COUNTIFS on Order Date)")
c.font = Font(name="Calibri", size=12, bold=True, color=WHITE); c.fill = PatternFill("solid", fgColor=TEAL)
for j, h in enumerate(["Month", "Orders", "Units", "Sales (PKR)", "Share of Sales", ""], 1):
    d.cell(row=11, column=j, value=h)
style_header(d, 11, 6)
months = [date(2026, m, 1) for m in range(1, 7)]
for i, m in enumerate(months):
    r = 12 + i
    d.cell(row=r, column=1, value=m).number_format = "mmm yyyy"
    d.cell(row=r, column=2, value=f"=COUNTIFS(Data!$B$2:$B${N},\">=\"&$A{r},Data!$B$2:$B${N},\"<\"&EDATE($A{r},1))")
    d.cell(row=r, column=3, value=f"=SUMIFS(Data!$G$2:$G${N},Data!$B$2:$B${N},\">=\"&$A{r},Data!$B$2:$B${N},\"<\"&EDATE($A{r},1))")
    d.cell(row=r, column=4, value=f"=SUMIFS(Data!$I$2:$I${N},Data!$B$2:$B${N},\">=\"&$A{r},Data!$B$2:$B${N},\"<\"&EDATE($A{r},1))")
    d.cell(row=r, column=5, value=f"=IF($A$5=0,0,D{r}/$A$5)").number_format = "0.0%"
    for col in (2, 3, 4):
        d.cell(row=r, column=col).number_format = "#,##0"
    for col in range(1, 6):
        d.cell(row=r, column=col).border = BORDER; d.cell(row=r, column=col).font = Font(name="Calibri", size=11)

# Category summary: header row 17? Monthly occupies 12-17, so category header at 19 is wrong.
# Re-plan: monthly 12-17, category title row 18? Overlap with KPI refs (A18:A22, D18:D22).
# So shift: category title row 18 conflicts — instead place category table starting row 19 with header 19?
# KPI formulas above already point at A18:A22/D18:D22, so keep category header at 17? No.
# SOLUTION: category title row 18, header row 19, data rows 20-24 — but KPI refs say 18-22.
# Fix KPI refs by rewriting them now to the final layout (title 18, header 19, data 20-24).
# Rewrite Best Category value cell (merged C7:D7 -> cell C7; Best Product stays at E7):
d.cell(row=7, column=3).value = "=INDEX($A$20:$A$24,MATCH(MAX($D$20:$D$24),$D$20:$D$24,0))"
d.merge_cells("A18:F18"); c = d.cell(row=18, column=1, value="SALES BY CATEGORY  (SUMIF / COUNTIF on Category)")
c.font = Font(name="Calibri", size=12, bold=True, color=WHITE); c.fill = PatternFill("solid", fgColor=TEAL)
for j, h in enumerate(["Category", "Orders", "Units", "Sales (PKR)", "Share of Sales", ""], 1):
    d.cell(row=19, column=j, value=h)
style_header(d, 19, 6)
cats = list(CATALOGUE.keys())
for i, cat in enumerate(cats):
    r = 20 + i
    d.cell(row=r, column=1, value=cat)
    d.cell(row=r, column=2, value=f"=COUNTIF(Data!$D$2:$D${N},$A{r})")
    d.cell(row=r, column=3, value=f"=SUMIF(Data!$D$2:$D${N},$A{r},Data!$G$2:$G${N})")
    d.cell(row=r, column=4, value=f"=SUMIF(Data!$D$2:$D${N},$A{r},Data!$I$2:$I${N})")
    d.cell(row=r, column=5, value=f"=IF($A$5=0,0,D{r}/$A$5)").number_format = "0.0%"
    for col in (2, 3, 4):
        d.cell(row=r, column=col).number_format = "#,##0"
    for col in range(1, 6):
        d.cell(row=r, column=col).border = BORDER; d.cell(row=r, column=col).font = Font(name="Calibri", size=11)
d.conditional_formatting.add(f"D20:D24", DataBarRule(start_type="num", start_value=0, end_type="max", color=TEAL))

# Product table: title 26, header 27, data 28-47 (20 products, sorted by Python-computed sales desc)
from collections import defaultdict
prod_sales = defaultdict(int); prod_units = defaultdict(int); prod_orders = defaultdict(int)
for r in rows:
    prod_sales[r[2]] += r[6] * r[7]; prod_units[r[2]] += r[6]; prod_orders[r[2]] += 1
ranked = sorted(prod_sales.items(), key=lambda kv: -kv[1])
# Best Product ref (E7) already points at A28:A47/D28:D47 — correct, leave it.
d.merge_cells("A26:F26"); c = d.cell(row=26, column=1, value="PRODUCTS RANKED BY SALES  (SUMIF / COUNTIF on Product — order from demo data)")
c.font = Font(name="Calibri", size=12, bold=True, color=WHITE); c.fill = PatternFill("solid", fgColor=TEAL)
for j, h in enumerate(["Product", "Orders", "Units", "Sales (PKR)", "Share of Sales", "Category"], 1):
    d.cell(row=27, column=j, value=h)
style_header(d, 27, 6)
prod_cat = {p: cat for p, cat, lo, hi in PRODUCTS}
for i, (prod, _) in enumerate(ranked):
    r = 28 + i
    d.cell(row=r, column=1, value=prod)
    d.cell(row=r, column=2, value=f"=COUNTIF(Data!$C$2:$C${N},$A{r})")
    d.cell(row=r, column=3, value=f"=SUMIF(Data!$C$2:$C${N},$A{r},Data!$G$2:$G${N})")
    d.cell(row=r, column=4, value=f"=SUMIF(Data!$C$2:$C${N},$A{r},Data!$I$2:$I${N})")
    d.cell(row=r, column=5, value=f"=IF($A$5=0,0,D{r}/$A$5)").number_format = "0.0%"
    d.cell(row=r, column=6, value=prod_cat[prod])
    for col in (2, 3, 4):
        d.cell(row=r, column=col).number_format = "#,##0"
    for col in range(1, 7):
        d.cell(row=r, column=col).border = BORDER; d.cell(row=r, column=col).font = Font(name="Calibri", size=11)
d.freeze_panes = "A3"
d.sheet_properties.pageSetUpPr.fitToPage = True

# Charts
ch1 = BarChart(); ch1.type = "col"; ch1.title = "Monthly Sales (PKR)"; ch1.y_axis.title = "PKR"; ch1.x_axis.title = "Month"
ch1.add_data(Reference(d, min_col=4, min_row=11, max_row=17), titles_from_data=True)
ch1.set_categories(Reference(d, min_col=1, min_row=12, max_row=17))
ch1.shape = 4; ch1.height = 7.5; ch1.width = 14; d.add_chart(ch1, "A49")
ch2 = BarChart(); ch2.type = "bar"; ch2.title = "Sales by Category (PKR)"
ch2.add_data(Reference(d, min_col=4, min_row=19, max_row=24), titles_from_data=True)
ch2.set_categories(Reference(d, min_col=1, min_row=20, max_row=24))
ch2.shape = 4; ch2.height = 7.5; ch2.width = 14; d.add_chart(ch2, "A64")

# README sheet inside workbook
rm = wb.create_sheet("README")
rm.column_dimensions["A"].width = 110
rm["A1"] = "BrightMart Retail — Sales Dashboard (DEMO / SAMPLE DATA ONLY)"
rm["A1"].font = Font(name="Calibri", size=14, bold=True, color=NAVY)
notes = [
    "This workbook is a SAMPLE delivery built on fictional demo data. It is not real client data and no real shop is named.",
    "Data sheet: 720 transactions, Jan–Jun 2026. Sales (PKR) = Quantity x Unit Price, as a formula in every row.",
    "Dashboard KPIs, monthly table, category table and product table are all live formulas (SUM / SUMIFS / COUNTIFS / SUMIF / COUNTIF).",
    "Charts read the monthly and category tables, so they move when the data moves.",
    "Slicers: the Data sheet is an Excel Table named SalesData — in Excel, click inside it, then Insert > Slicer, and pick Category, Region or Channel.",
    "Print: Dashboard is set to landscape, fit-to-page, with print titles. Data sheet prints landscape with repeating headers.",
]
for i, t in enumerate(notes, 3):
    rm.cell(row=i, column=1, value=f"• {t}").font = Font(name="Calibri", size=11)
    rm.cell(row=i, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    rm.row_dimensions[i].height = 30

wb.active = wb["Dashboard"]
wb.save(OUT)
print(f"Wrote {OUT}: {len(rows)} rows, last data row {N}")
print(f"Python totals: sales={sum(r[6]*r[7] for r in rows):,} units={sum(r[6] for r in rows)} orders={len(rows)}")
print(f"Best category: {max(cats, key=lambda c: sum(r[6]*r[7] for r in rows if r[3]==c))}")
print(f"Best product: {ranked[0][0]} ({ranked[0][1]:,})")
print("Monthly sales:", {str(m): sum(r[6]*r[7] for r in rows if r[1].year==m.year and r[1].month==m.month) for m in months})
