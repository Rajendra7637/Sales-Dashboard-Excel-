# Global Retail Superstore — Sales Data Cleaning & Analysis

An end-to-end Excel project that takes a messy, multi-source sales export and turns it into a clean, analysis-ready dataset with PivotTables, charts, and a one-page KPI dashboard.

## 📌 Background

"Global Retail Superstore" is a multinational e-commerce and retail company. Its order-management system exports raw sales data from multiple regional systems (web store, mobile app, in-store POS, and phone orders) into a single spreadsheet feed — with no standardization applied. This project cleans that raw export and builds a trustworthy sales analysis for leadership.

## 📂 File

`Sales_using_Excel.xlsx` — a single workbook with five sheets:

| Sheet | Description |
|---|---|
| `Problem_Statement` | Full project brief: background, known data-quality issues, and deliverables |
| `Raw_Data` | ~20,400 rows × 20 columns of raw, unstandardized order-level data |
| `Cleaned_Data` | Cleaned dataset (24 columns, incl. calculated `TotalPrice`) ready for analysis |
| `Pivot_table` | PivotTables covering revenue, returns, top products, AOV, and trends |
| `SalesDashboard` | One-page KPI dashboard with summary charts |

## 🧾 Raw data columns

`OrderID`, `OrderDate`, `CustomerName`, `CustomerEmail`, `Country`, `Region`, `City`, `ProductCategory`, `ProductName`, `Quantity`, `UnitPrice`, `Discount`, `ShippingCost`, `PaymentMethod`, `OrderStatus`, `SalesChannel`, `CustomerAgeGroup`, `Rating`, `ReturnFlag`, `Notes`

## 🧹 Data quality issues addressed

- Exact duplicate rows and fully blank rows
- `OrderDate` exported in 6+ different formats, standardized to one
- Inconsistent text casing / stray whitespace (names, categories, payment method, status, channel)
- Multiple spellings of the same country (e.g. `USA` / `U.S.A.` / `United States` / `us`) mapped to one standard value
- Typos in `ProductCategory` (e.g. `Electroncis`, `Furnature`, `Beuaty`)
- `UnitPrice` stored as text with mixed formats (`$123.45`, `123.45 USD`) converted to clean numeric
- `Discount` inconsistently stored as a fraction vs. whole-number percentage, normalized to one convention
- Negative/zero `Quantity` values investigated and handled
- Missing values in `CustomerEmail`, `OrderDate`, `Quantity`, `UnitPrice`, `ShippingCost`, `Rating`, `CustomerAgeGroup`
- Out-of-range `Rating` values (6–8) corrected or removed
- `ReturnFlag` encoded inconsistently (`Yes/No`, `Y/N`, `1/0`, blank) standardized to one convention
- Malformed emails (`"at"` instead of `"@"`) corrected/flagged

## 📊 Analysis performed

- Total revenue by year and by region
- Top 10 products by revenue
- Average order value (AOV) by sales channel
- Return rate (%) by product category
- Monthly revenue trend
- Average customer rating by product category

### Headline numbers (from the workbook)

- **Total revenue:** ~$53.28M
- **Total orders:** ~19,790–20,000
- **Overall return rate:** ~49.6%
- **Best-performing region:** South (~$6.98M)
- **Top sales channel by revenue:** Online (~$16.19M), followed by Store (~$15.91M)

## 📈 PivotTables & Dashboard

The `Pivot_table` sheet includes PivotTables for:
- Revenue by Region
- Orders/Return status counts
- Top 10 products by revenue
- AOV by sales channel
- Monthly revenue trend and average rating

The `SalesDashboard` sheet consolidates key KPIs (Total Revenue, Total Orders, AOV, Return Rate) with supporting charts into a single view for leadership.

## 🛠️ Skills practiced

TEXT functions (`TRIM`, `PROPER`, `UPPER`, `LOWER`, `SUBSTITUTE`), Find & Replace, Text to Columns, Flash Fill, data validation, duplicate removal, `VLOOKUP`/`INDEX-MATCH` for standardizing categories against a lookup table, PivotTables, PivotCharts, and conditional formatting for outlier detection.
