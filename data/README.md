# Data dictionary — BrightMart Sales Dashboard (DEMO / SAMPLE DATA)

Source sheet: **Data** in `BrightMart_Sales_Dashboard.xlsx`, Excel Table `SalesData`.
720 rows, Jan–Jun 2026. All rows are fictional sample transactions — no real customers.

| Column | Type | Meaning | Example / notes |
|---|---|---|---|
| Order ID | Text | Unique order reference, `ORD-0001` … `ORD-0720`, in date order | In this sample each Order ID appears once — one row = one order (a real shop file with multi-line orders would repeat an ID, and `COUNTA` would then count lines, not orders) |
| Order Date | Date (`yyyy-mm-dd`) | Date of the transaction | Period Jan–Jun 2026; first transaction 2026-01-03, last 2026-06-30 |
| Product | Text | Product name sold | 20 products (see below) |
| Category | Text | Product category | Groceries, Electronics, Clothing, Home & Kitchen, Personal Care |
| Region | Text | Customer/shop region for the sale | Punjab, Sindh, Khyber Pakhtunkhwa, Balochistan, Islamabad |
| Channel | Text | How the order came in | In-Store, Online, WhatsApp Order |
| Quantity | Whole number | Units sold on this line | 1–5 (most lines are 1) |
| Unit Price (PKR) | Whole number (PKR) | Price per unit for this product on this sale | Realistic range per product, e.g. Dinner Set 24pc PKR 6,500–9,000 |
| Sales (PKR) | Formula: `= Quantity × Unit Price` | Line total in PKR | **Never typed** — every cell is a formula, so Dashboard totals always reconcile to the Data sheet |

## Products by category (demo catalogue)

- **Groceries:** Basmati Rice 5kg, Cooking Oil 5L, Wheat Flour 10kg, Tea Leaves 1kg
- **Electronics:** LED Bulb Pack (4), Mobile Charger, Bluetooth Speaker, Electric Kettle
- **Clothing:** Men's Kurta, Women's Lawn Suit, Kids T-Shirt Pack, Winter Shawl
- **Home & Kitchen:** Dinner Set 24pc, Non-Stick Pan, Bed Sheet Set, Water Bottle 1L
- **Personal Care:** Shampoo 400ml, Face Wash, Hand Wash Pack, Body Lotion

## Quality notes

- No blank cells, no duplicate Order IDs, no negative quantities or prices.
- Dashboard totals were verified by recalculation (LibreOffice headless):
  monthly, category and product tables each sum to the KPI total (PKR 3,469,346),
  and monthly orders sum to 720. The same totals are printed by
  `python3 build_dashboard.py` and listed in the KPI table in `README.md`.
- To reuse with real data: keep these exact column names and order; the Dashboard
  formulas cover rows 2–721, so extend the range if the real file is longer.
