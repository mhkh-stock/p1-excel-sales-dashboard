# P1 — Excel Sales Dashboard for a Retail Shop

**Shop:** BrightMart Retail (fictional) · **Period:** Jan–Jun 2026 · **Status: SAMPLE delivery on DEMO data**
**Tools:** Microsoft Excel (formulas, Excel Table, charts) · **File:** `BrightMart_Sales_Dashboard.xlsx`

> ⚠️ **Demo data only.** Every number in this workbook comes from clearly-labelled fictional
> sample transactions generated for this project. It is **not** real client data and no real
> shop is named. The skill and the template are real; the client history is not claimed.

## Business question it answers

> "Show me my whole business on one page: how much did we sell, what is selling best,
> and which months are strong or weak?"

A small retail owner does not need a 40-page report. They need one Dashboard sheet they
can open every morning: total sales, orders, average order value, best category and
best product, a monthly trend, and a ranked product list — all updating from the
transaction data, with no manual re-typing of totals.

## What is inside

| Sheet | What it does |
|---|---|
| **Dashboard** | KPIs, monthly summary, category summary, ranked product list, 2 charts. Print-ready (landscape, fit-to-page). |
| **Data** | 720 transactions as an Excel Table (`SalesData`) with filters, frozen panes, and a formula in every Sales cell (`Quantity × Unit Price`). Add slicers in Excel via Insert → Slicer on Category, Region or Channel. |
| **README** | Short in-workbook guide (also this file, in full). |

**Every total is a formula, not a typed number:** `SUM`, `SUMIFS`, `COUNTIFS`,
`SUMIF`, `COUNTIF`, plus `INDEX`/`MATCH` for the best category and best product.
Change a row on the Data sheet and the Dashboard moves with it.

## 3 key findings (from the demo data)

1. **Home & Kitchen drives the shop:** PKR 1,112,961 — **32.1% of all sales**.
   One item, the Dinner Set 24pc, alone is PKR 500,289 (**14.4%** of total sales).
   *So what:* keep that set in stock first; a stock-out there hurts more than anywhere else.
2. **March is the peak, April the dip:** March PKR 659,012 (130 orders) vs April
   PKR 521,107 — a **26.5% swing** month to month.
   *So what:* check what changed in April (stock, pricing, footfall) before assuming a trend.
3. **Small baskets, store-led sales:** average order value **PKR 4,819** on
   **1.75 units per order**, and **56.1% of sales are In-Store** (Online 30.5%,
   WhatsApp Order 13.4%). Punjab is the top region at 34.0%.
   *So what:* the fastest growth lever in this sample is basket size (bundles, add-ons),
   not just more orders — and WhatsApp ordering has room to grow from 13.4%.

## Recommendation

Protect the Dinner Set 24pc first (never stock out), then grow basket size with
bundles/add-ons — in this sample, one more unit per order is worth more than more
orders. WhatsApp ordering (13.4%) is the cheapest channel to grow next.

## Demo-data KPIs (verified, see below)

| KPI | Value |
|---|---|
| Total sales | PKR 3,469,346 |
| Total orders | 720 |
| Total units sold | 1,262 |
| Average order value | PKR 4,819 |
| Best category | Home & Kitchen |
| Best product | Dinner Set 24pc (PKR 500,289) |
| Best month | March 2026 (PKR 659,012) |

*Verified by recalculation: the build script prints the same totals from the
source rows, and LibreOffice recalculation of the workbook reproduces every KPI,
monthly, category and product total (all three tables sum to PKR 3,469,346;
monthly orders sum to 720; no formula errors).*

## How to use it (under 5 steps)

1. Open `BrightMart_Sales_Dashboard.xlsx` and start on the **Dashboard** sheet.
2. To explore: go to the **Data** sheet, filter by Category, Region or Channel
   (or Insert → Slicer on the `SalesData` table for one-click filtering).
3. To reuse it for a real shop: replace the rows on the Data sheet with real
   transactions, keeping the same columns — every Dashboard formula points at
   rows 2–721, so totals, tables and charts update by themselves for up to 720
   rows. For a longer file, extend the formula ranges (e.g. `I2:I721` → `I2:I1501`)
   on the Dashboard and in the `SalesData` table.
4. To print: Dashboard is set to landscape, fit-to-one-page. Data prints landscape
   with repeating header rows.
5. To rebuild from scratch: `pip install -r requirements.txt` (openpyxl),
   then run `python3 build_dashboard.py` — it regenerates the workbook with the
   same seeded demo data and prints the verified totals.

## Tools

- Microsoft Excel (any recent version — formulas used are standard, no macros, no add-ins)
- Python + openpyxl (build script only; the workbook itself needs no code to run)

## Data source note

**Demo/sample data, generated for this project** (seeded random generator in
`build_dashboard.py`: 720 rows, Jan–Jun 2026, 20 products across 5 categories,
5 regions, 3 channels). Prices are realistic PKR ranges per product, not real prices.
No customer names, phone numbers or payment details exist in this data — by design.

## What I learned / limits

- Six months of demo data show seasonality only roughly — a full 12 months would
  separate a real trend from a one-off dip like April.
- No sales targets exist in this sample, so there is no targets-vs-actual view
  (the Premium tier adds one when a buyer supplies targets).
- Excel formulas are the right tool at this size; past ~10k rows a database +
  Power BI would be the honest next step.

## Repo standard note

This paid project follows the portfolio repo standard (Master Plan §3.2), slimmed for
an Excel deliverable: README + data dictionary (`data/README.md`) + the deliverable
itself + a Fiverr-ready package note (`PACKAGE-NOTE.md`) + a reproducible build script.
No screenshots are bundled yet — they are added at publish time from the live workbook.
