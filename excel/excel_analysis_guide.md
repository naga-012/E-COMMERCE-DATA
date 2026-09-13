# Excel E-Commerce Master Analytics Guide & Architecture Reference

This guide provides step-by-step instructions for navigating, auditing, and extending the consolidated all-in-one financial model and database in:
- **`excel/ecommerce_analysis.xlsx`** (Master Workbook)
- **`excel/ecommerce_master_all_in_one.xlsx`** (Consolidated Master All-In-One)

---

## 1. Workbook Architecture (17 Consolidated Sheets)

The workbook consolidates all executive dashboards, analytical financial models, and full underlying databases into one unified Excel file:

| # | Sheet Name | Category | Record Count | Purpose & Primary Features |
|---|---|---|---|---|
| 1 | **Overview & Navigation** | Navigation | 16 Direct Links | Table of contents with clickable `=HYPERLINK()` jumps, KPI summary cards, and workbook index. |
| 2 | **Dashboard** | Executive | Dynamic Model | C-suite sales & profit dashboard with interactive Year & Month dropdowns, KPI cards, and category & regional tables. |
| 3 | **KPI Summary** | Governance | 12 Metrics | Master enterprise metric dictionary with benchmark targets, formulas, and operational audits. |
| 4 | **Sales Analysis** | Financial | 36 Months | 3-year monthly performance time-series (2022–2024), MoM growth %, and net margin %. |
| 5 | **Product Analysis** | Portfolio | 40 Key SKUs | Top and bottom SKU profitability matrix, return rates, and BCG portfolio classifications. |
| 6 | **Customer Analysis** | Customer Science | 7 Segments | RFM segmentation summary (Champions, At Risk, etc.), customer count, spend, and revenue shares. |
| 7 | **Regional Analysis** | Geographic | 35 States / Zones | Regional and state-level volume, gross revenue, net profit margin %, and AOV. |
| 8 | **Returns Analysis** | Operations | 7 Reasons / 5 Methods | Return root-cause distribution and payment channel performance share. |
| 9 | **Pivot Tables** | Source Transactions | **52,500 Rows** | Full transactional orders dataset formatted with headers and freeze-panes for custom Excel Pivot Tables. |
| 10 | **Customers Data** | Master Database | **5,200 Rows** | Complete cleaned customer master directory with signup date, city, state, and segment. |
| 11 | **Products Data** | Master Database | **550 Rows** | Complete cleaned product catalog with unit costs, base prices, categories, and suppliers. |
| 12 | **Payments Data** | Master Database | **52,500 Rows** | Complete cleaned payment transaction records with payment methods, status, and dates. |
| 13 | **Returns Data** | Master Database | **3,676 Rows** | Complete cleaned returns records with return reasons, refund amounts, and status. |
| 14 | **Customer RFM Profiles** | Analytical Dataset | **5,156 Profiles** | Individual customer RFM scores (1-5), Recency, Frequency, Monetary value, and assigned segment. |
| 15 | **Product Performance** | Analytical Dataset | **550 Products** | Complete product metrics table with unit margins, return rate %, total revenue, and portfolio tier. |
| 16 | **Monthly Sales Summary** | Analytical Dataset | **36 Months** | Aggregated monthly financial summary table. |
| 17 | **Data Audit & Dictionary** | Data Quality | 13 Tables Audited | Raw vs. cleaned data reconciliation audit, data quality scores, and full database schema dictionary. |

---

## 2. Core Excel Formulas Implemented

### 1. Dynamic Gross Sales Revenue (Responsive to Year & Month Slicers)
```excel
=IF(C5="All Months", IF(C4="All Years", SUM('Sales Analysis'!E4:E39), SUMIF('Sales Analysis'!B4:B39, C4, 'Sales Analysis'!E4:E39)), IF(C4="All Years", SUMIF('Sales Analysis'!C4:C39, C5, 'Sales Analysis'!E4:E39), SUMIFS('Sales Analysis'!E4:E39, 'Sales Analysis'!B4:B39, C4, 'Sales Analysis'!C4:C39, C5)))
```

### 2. Net Profit Margin Percentage
```excel
=Dashboard!D8 / Dashboard!B8
```
Calculates blended enterprise profit margin (`Total Net Profit / Total Gross Sales`).

### 3. Month-over-Month (MoM) Growth
```excel
=(E5 - E4) / E4
```
Calculates revenue acceleration from the previous month, formatted as `0.00%`.

### 4. Enterprise Order Return Rate via Conditional Counting
```excel
=COUNTIF('Pivot Tables'!H4:H55000, "Returned") / Dashboard!F8
```
Audits return rate across all 52,500 transactions against target threshold (<8.0%).

### 5. Table of Contents Hyperlink Navigation
```excel
=HYPERLINK("#'Dashboard'!A1", "Go to Dashboard")
```
Enables 1-click jump navigation directly to any sheet within the workbook.

---

## 3. How to Build Native Excel Pivot Tables

Follow these steps in Microsoft Excel to create interactive pivots from the `Pivot Tables` sheet:

### Step 1: Regional Profitability Pivot Table
1. Navigate to sheet `Pivot Tables`.
2. Select range `A3:L52503` (or click cell `A3` and press `Ctrl + A`).
3. Click **Insert > PivotTable** > **New Worksheet**.
4. Drag fields into areas:
   - **Rows**: `region`, `state`
   - **Values**: `sales_amount` (Summarize by Sum, Format Currency `₹ #,##0`), `profit_amount` (Summarize by Sum)
   - **Calculated Field**: Name = `Profit Margin %`, Formula = `=profit_amount / sales_amount`.

### Step 2: Customer Status Matrix
1. Select transaction table in `Pivot Tables`.
2. Drag `order_status` to **Rows**.
3. Drag `order_id` to **Values** (Summarize by Count).
4. Drag `sales_amount` to **Values** (Show Values As > **% of Grand Total**).
