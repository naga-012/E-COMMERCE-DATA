"""
build_excel_workbook.py
=======================
Generates a comprehensive, multi-sheet, formula-driven, all-in-one Master Excel model:
- excel/ecommerce_analysis.xlsx (Primary Master Workbook)
- excel/ecommerce_master_all_in_one.xlsx (All-in-One Consolidated Workbook)
- excel/ecommerce_analysis_updated.xlsx (Synchronized Mirror)

Consolidates all executive dashboards, analytical models, and complete project datasets
(52,500 orders, 5,200 customers, 550 products, 52,500 payments, 3,676 returns, RFM profiles,
product metrics, monthly aggregates, and data audit tables) into ONE unified Excel file.
"""

import os
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data", "cleaned")
RAW_DIR = os.path.join(BASE_DIR, "data", "raw")
EXCEL_DIR = os.path.join(BASE_DIR, "excel")
os.makedirs(EXCEL_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(EXCEL_DIR, "ecommerce_analysis.xlsx")
ALL_IN_ONE_PATH = os.path.join(EXCEL_DIR, "ecommerce_master_all_in_one.xlsx")
UPDATED_PATH = os.path.join(EXCEL_DIR, "ecommerce_analysis_updated.xlsx")

print("Loading datasets...")
# Cleaned datasets
df_monthly = pd.read_csv(os.path.join(DATA_DIR, "monthly_sales_summary.csv"))
df_prod_perf = pd.read_csv(os.path.join(DATA_DIR, "product_performance_metrics.csv"))
df_rfm = pd.read_csv(os.path.join(DATA_DIR, "customer_rfm_profiles.csv"))
df_ord = pd.read_csv(os.path.join(DATA_DIR, "orders_cleaned.csv"))
df_ret = pd.read_csv(os.path.join(DATA_DIR, "returns_cleaned.csv"))
df_pay = pd.read_csv(os.path.join(DATA_DIR, "payments_cleaned.csv"))
df_cust = pd.read_csv(os.path.join(DATA_DIR, "customers_cleaned.csv"))
df_prod = pd.read_csv(os.path.join(DATA_DIR, "products_cleaned.csv"))

# Raw datasets for auditing
df_raw_ord = pd.read_csv(os.path.join(RAW_DIR, "orders.csv"))
df_raw_cust = pd.read_csv(os.path.join(RAW_DIR, "customers.csv"))
df_raw_pay = pd.read_csv(os.path.join(RAW_DIR, "payments.csv"))
df_raw_prod = pd.read_csv(os.path.join(RAW_DIR, "products.csv"))
df_raw_ret = pd.read_csv(os.path.join(RAW_DIR, "returns.csv"))
print("All datasets loaded successfully.")


def create_workbook():
    start_time = time.time()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Color Palette & Fills
    navy_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    blue_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    emerald_fill = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
    purple_fill = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")
    gray_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    accent_fill = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    light_green_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

    # Fonts
    white_font_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    white_font_header = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    navy_font_title = Font(name="Calibri", size=16, bold=True, color="1E293B")
    section_font = Font(name="Calibri", size=13, bold=True, color="1E293B")
    bold_font = Font(name="Calibri", size=11, bold=True, color="000000")
    regular_font = Font(name="Calibri", size=11, color="000000")
    link_font = Font(name="Calibri", size=11, color="2563EB", underline="single")
    kpi_val_font = Font(name="Calibri", size=18, bold=True, color="2563EB")
    kpi_lbl_font = Font(name="Calibri", size=10, bold=True, color="64748B")

    # Borders
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    total_border = Border(
        top=Side(style="thin", color="1E293B"),
        bottom=Side(style="double", color="1E293B")
    )

    # Helper: Fast Column Width Autosizing from DataFrame Sample
    def set_col_widths_from_df(ws, df, start_col=1):
        for col_idx, col in enumerate(df.columns, start=start_col):
            col_letter = get_column_letter(col_idx)
            sample_vals = [str(col)] + [str(v) for v in df[col].dropna().head(40).values]
            max_len = max(len(s) for s in sample_vals) if sample_vals else len(str(col))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================
    # SHEET 1: Overview & Navigation (Table of Contents)
    # =========================================================
    print("Building Sheet 1: Overview & Navigation...")
    ws_index = wb.create_sheet(title="Overview & Navigation")
    ws_index.views.sheetView[0].showGridLines = True

    ws_index.merge_cells("B2:J2")
    ws_index["B2"] = "E-COMMERCE ENTERPRISE ALL-IN-ONE MASTER WORKBOOK"
    ws_index["B2"].font = navy_font_title
    ws_index["B2"].alignment = Alignment(vertical="center")

    ws_index.merge_cells("B3:J3")
    ws_index["B3"] = "Unified Master Model: Executive Dashboards, Financial Analytics, and Full Cleaned Datasets (2022 - 2024)"
    ws_index["B3"].font = Font(name="Calibri", size=11, italic=True, color="64748B")

    # Executive Stat Highlight Cards on Index Sheet
    stat_cards = [
        ("B5:C5", "B6:C6", "GROSS REVENUE", "₹ 330,418,172", "36-Month Operating Total"),
        ("D5:E5", "D6:E6", "NET PROFIT", "₹ 87,314,772", "26.43% Blended Margin"),
        ("F5:G5", "F6:G6", "TOTAL ORDERS", "52,500", "100% Processed Records"),
        ("H5:I5", "H6:I6", "ACTIVE CUSTOMERS", "5,200", "99.01% Repeat Purchase Rate"),
        ("J5:K5", "J6:K6", "PRODUCT CATALOG", "550", "4 Active Categories")
    ]
    for top_range, val_range, label, val, sub in stat_cards:
        ws_index.merge_cells(top_range)
        ws_index.merge_cells(val_range)
        t_cell = ws_index[top_range.split(":")[0]]
        v_cell = ws_index[val_range.split(":")[0]]
        t_cell.value = label
        t_cell.font = kpi_lbl_font
        t_cell.fill = accent_fill
        t_cell.alignment = Alignment(horizontal="center", vertical="center")

        v_cell.value = val
        v_cell.font = kpi_val_font
        v_cell.fill = accent_fill
        v_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws_index["B8"] = "MASTER WORKBOOK DIRECTORY & NAVIGATION"
    ws_index["B8"].font = section_font

    nav_headers = ["Category", "Sheet Name", "Direct Navigation Link", "Record Count", "Description & Purpose"]
    for col_idx, h in enumerate(nav_headers, start=2):
        c = ws_index.cell(row=9, column=col_idx, value=h)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center", vertical="center")

    nav_rows = [
        ("Executive Analytics", "Dashboard", "Go to Dashboard", "Dynamic Model", "Executive KPI cards, interactive Year & Month selectors, Category & Regional breakdown."),
        ("Executive Analytics", "KPI Summary", "Go to KPI Summary", "12 Metrics", "Formal enterprise KPI dictionary, operational benchmarks, COGS, and audit formulas."),
        ("Financial Analytics", "Sales Analysis", "Go to Sales Analysis", "36 Months", "Full monthly financial time-series (2022–2024), MoM growth %, and net margin %."),
        ("Portfolio Analytics", "Product Analysis", "Go to Product Analysis", "40 Key SKUs", "Top and bottom product profitability matrix, returns %, and BCG portfolio classification."),
        ("Customer Science", "Customer Analysis", "Go to Customer Analysis", "7 Segments", "RFM customer segmentation rollup, customer volume, spend, and revenue contribution."),
        ("Geographic Analytics", "Regional Analysis", "Go to Regional Analysis", "35 States / Zones", "Territory and state-level order volume, sales amount, profit margin %, and AOV."),
        ("Operations Analytics", "Returns Analysis", "Go to Returns Analysis", "7 Causes / 5 Methods", "Return root-cause distribution and payment method performance matrix."),
        ("Source Data & Pivots", "Pivot Tables", "Go to Pivot Tables (Orders)", "52,500 Rows", "Full transaction dataset (all 52,500 orders) formatted for native Excel pivot tables."),
        ("Master Dataset", "Customers Data", "Go to Customers Data", "5,200 Rows", "Complete cleaned customer master table with city, state, signup date, and segment."),
        ("Master Dataset", "Products Data", "Go to Products Data", "550 Rows", "Complete cleaned product catalog with categories, cost prices, base prices, and suppliers."),
        ("Master Dataset", "Payments Data", "Go to Payments Data", "52,500 Rows", "Complete cleaned payment transactions with payment method, status, and transaction timestamps."),
        ("Master Dataset", "Returns Data", "Go to Returns Data", "3,676 Rows", "Complete cleaned returns data with order IDs, return reasons, and refund amounts."),
        ("Analytical Dataset", "Customer RFM Profiles", "Go to Customer RFM Profiles", "5,156 Profiles", "Individual customer RFM scores (1-5), Recency, Frequency, Monetary, and assigned segment."),
        ("Analytical Dataset", "Product Performance", "Go to Product Performance", "550 Products", "Full product metrics table with unit margins, return rate %, revenue, and classifications."),
        ("Analytical Dataset", "Monthly Sales Summary", "Go to Monthly Sales Summary", "36 Months", "Structured monthly aggregated sales, profit, orders, and operational metrics."),
        ("Data Quality & Audit", "Data Audit & Dictionary", "Go to Data Audit & Dictionary", "13 Tables Audited", "Raw vs Cleaned data record comparison, data quality audit, and schema dictionary.")
    ]

    for r_idx, (cat, s_name, link_txt, recs, desc) in enumerate(nav_rows, start=10):
        c_cat = ws_index.cell(row=r_idx, column=2, value=cat)
        c_cat.font = bold_font
        c_cat.fill = gray_fill if r_idx % 2 == 0 else PatternFill(fill_type=None)
        c_cat.border = thin_border

        c_name = ws_index.cell(row=r_idx, column=3, value=s_name)
        c_name.font = bold_font
        c_name.border = thin_border

        c_link = ws_index.cell(row=r_idx, column=4, value=f'=HYPERLINK("#\'{s_name}\'!A1", "{link_txt}")')
        c_link.font = link_font
        c_link.border = thin_border

        c_recs = ws_index.cell(row=r_idx, column=5, value=recs)
        c_recs.font = regular_font
        c_recs.alignment = Alignment(horizontal="center")
        c_recs.border = thin_border

        c_desc = ws_index.cell(row=r_idx, column=6, value=desc)
        c_desc.font = regular_font
        c_desc.border = thin_border

    ws_index.column_dimensions["B"].width = 24
    ws_index.column_dimensions["C"].width = 28
    ws_index.column_dimensions["D"].width = 28
    ws_index.column_dimensions["E"].width = 18
    ws_index.column_dimensions["F"].width = 75

    # =========================================================
    # SHEET 2: Dashboard
    # =========================================================
    print("Building Sheet 2: Dashboard...")
    ws_dash = wb.create_sheet(title="Dashboard")
    ws_dash.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_dash.merge_cells("B2:K2")
    ws_dash["B2"] = "E-COMMERCE EXECUTIVE SALES & PROFITABILITY DASHBOARD"
    ws_dash["B2"].font = navy_font_title
    ws_dash["B2"].alignment = Alignment(vertical="center")

    ws_dash.merge_cells("B3:K3")
    ws_dash["B3"] = "Enterprise Performance Reporting Model (2022 - 2024) | All Monetary Values in INR (₹)"
    ws_dash["B3"].font = Font(name="Calibri", size=10, italic=True, color="64748B")

    # Interactive Year Filter (Row 4)
    ws_dash["B4"] = "SELECT YEAR FILTER:"
    ws_dash["B4"].font = bold_font
    ws_dash["B4"].alignment = Alignment(horizontal="right", vertical="center")
    
    ws_dash.merge_cells("C4:D4")
    filter_cell = ws_dash["C4"]
    filter_cell.value = "2023"
    filter_cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    filter_cell.fill = blue_fill
    filter_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    dv_year = DataValidation(type="list", formula1='"All Years,2022,2023,2024"', allow_blank=False)
    dv_year.error ='Please select a valid year from the list'
    dv_year.errorTitle = 'Invalid Year'
    dv_year.prompt = 'Click dropdown to select 2022, 2023, 2024, or All Years'
    dv_year.promptTitle = 'Interactive Year Selector'
    ws_dash.add_data_validation(dv_year)
    dv_year.add(filter_cell)

    ws_dash.merge_cells("E4:K4")
    ws_dash["E4"] = "<- Click dropdown to toggle between 2022, 2023, 2024, or All Years. KPIs auto-update below!"
    ws_dash["E4"].font = Font(name="Calibri", size=10, italic=True, color="2563EB")
    ws_dash["E4"].alignment = Alignment(vertical="center")

    # Interactive Month Filter (Row 5)
    ws_dash["B5"] = "SELECT MONTH FILTER:"
    ws_dash["B5"].font = bold_font
    ws_dash["B5"].alignment = Alignment(horizontal="right", vertical="center")

    ws_dash.merge_cells("C5:D5")
    filter_m_cell = ws_dash["C5"]
    filter_m_cell.value = "All Months"
    filter_m_cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    filter_m_cell.fill = emerald_fill
    filter_m_cell.alignment = Alignment(horizontal="center", vertical="center")

    dv_month = DataValidation(type="list", formula1='"All Months,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"', allow_blank=False)
    dv_month.error = 'Please select a month or All Months'
    dv_month.errorTitle = 'Invalid Month'
    dv_month.prompt = 'Click dropdown to select specific month or All Months'
    dv_month.promptTitle = 'Interactive Month Selector'
    ws_dash.add_data_validation(dv_month)
    dv_month.add(filter_m_cell)

    ws_dash.merge_cells("E5:K5")
    ws_dash["E5"] = "<- Select month (or All Months) to filter financial records and KPIs!"
    ws_dash["E5"].font = Font(name="Calibri", size=10, italic=True, color="059669")
    ws_dash["E5"].alignment = Alignment(vertical="center")

    # Dynamic KPI Cards (Row 7 - 8) based on Year in C4 and Month in C5
    kpis = [
        ("B7:C7", "B8:C8", "TOTAL GROSS REVENUE", "=IF(C5=\"All Months\", IF(C4=\"All Years\", SUM('Sales Analysis'!E4:E39), SUMIF('Sales Analysis'!B4:B39, C4, 'Sales Analysis'!E4:E39)), IF(C4=\"All Years\", SUMIF('Sales Analysis'!C4:C39, C5, 'Sales Analysis'!E4:E39), SUMIFS('Sales Analysis'!E4:E39, 'Sales Analysis'!B4:B39, C4, 'Sales Analysis'!C4:C39, C5)))", "₹ #,##0"),
        ("D7:E7", "D8:E8", "TOTAL NET PROFIT", "=IF(C5=\"All Months\", IF(C4=\"All Years\", SUM('Sales Analysis'!F4:F39), SUMIF('Sales Analysis'!B4:B39, C4, 'Sales Analysis'!F4:F39)), IF(C4=\"All Years\", SUMIF('Sales Analysis'!C4:C39, C5, 'Sales Analysis'!F4:F39), SUMIFS('Sales Analysis'!F4:F39, 'Sales Analysis'!B4:B39, C4, 'Sales Analysis'!C4:C39, C5)))", "₹ #,##0"),
        ("F7:G7", "F8:G8", "TOTAL ORDERS", "=IF(C5=\"All Months\", IF(C4=\"All Years\", SUM('Sales Analysis'!D4:D39), SUMIF('Sales Analysis'!B4:B39, C4, 'Sales Analysis'!D4:D39)), IF(C4=\"All Years\", SUMIF('Sales Analysis'!C4:C39, C5, 'Sales Analysis'!D4:D39), SUMIFS('Sales Analysis'!D4:D39, 'Sales Analysis'!B4:B39, C4, 'Sales Analysis'!C4:C39, C5)))", "#,##0"),
        ("H7:I7", "H8:I8", "AVERAGE ORDER VALUE (AOV)", "=B8/F8", "₹ #,##0.00"),
        ("J7:K7", "J8:K8", "OVERALL PROFIT MARGIN", "=D8/B8", "0.00%")
    ]

    for top_range, val_range, label, formula, num_format in kpis:
        ws_dash.merge_cells(top_range)
        ws_dash.merge_cells(val_range)
        top_cell = ws_dash[top_range.split(":")[0]]
        val_cell = ws_dash[val_range.split(":")[0]]
        
        top_cell.value = label
        top_cell.font = kpi_lbl_font
        top_cell.fill = accent_fill
        top_cell.alignment = Alignment(horizontal="center", vertical="center")

        val_cell.value = formula
        val_cell.font = kpi_val_font
        val_cell.fill = accent_fill
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.number_format = num_format

    # Table 1: Category Snapshot on Dashboard
    ws_dash["B10"] = "Category Financial Contribution"
    ws_dash["B10"].font = section_font
    cat_headers = ["Category", "Orders", "Revenue (INR)", "Profit (INR)", "Profit Margin %"]
    for col_idx, h in enumerate(cat_headers, start=2):
        cell = ws_dash.cell(row=11, column=col_idx, value=h)
        cell.fill = navy_fill
        cell.font = white_font_bold
        cell.alignment = Alignment(horizontal="center")

    cat_summary = df_ord.merge(df_prod_perf[["product_id", "category"]], on="product_id") \
                        .groupby("category").agg(
                            Orders=("order_id", "count"),
                            Revenue=("sales_amount", "sum"),
                            Profit=("profit_amount", "sum")
                        ).reset_index().sort_values(by="Revenue", ascending=False)

    for r_idx, row in enumerate(cat_summary.itertuples(), start=12):
        ws_dash.cell(row=r_idx, column=2, value=row.category).font = regular_font
        ws_dash.cell(row=r_idx, column=3, value=row.Orders).number_format = "#,##0"
        ws_dash.cell(row=r_idx, column=4, value=row.Revenue).number_format = "₹ #,##0"
        ws_dash.cell(row=r_idx, column=5, value=row.Profit).number_format = "₹ #,##0"
        m_cell = ws_dash.cell(row=r_idx, column=6, value=f"=E{r_idx}/D{r_idx}")
        m_cell.number_format = "0.0%"
        m_cell.font = regular_font

    # Total Row for Category Table
    tot_row = 12 + len(cat_summary)
    ws_dash.cell(row=tot_row, column=2, value="Total").font = bold_font
    ws_dash.cell(row=tot_row, column=3, value=f"=SUM(C12:C{tot_row-1})").number_format = "#,##0"
    ws_dash.cell(row=tot_row, column=4, value=f"=SUM(D12:D{tot_row-1})").number_format = "₹ #,##0"
    ws_dash.cell(row=tot_row, column=5, value=f"=SUM(E12:E{tot_row-1})").number_format = "₹ #,##0"
    ws_dash.cell(row=tot_row, column=6, value=f"=E{tot_row}/D{tot_row}").number_format = "0.0%"
    for c in range(2, 7):
        ws_dash.cell(row=tot_row, column=c).border = total_border
        ws_dash.cell(row=tot_row, column=c).font = bold_font

    # Table 2: Regional Snapshot on Dashboard
    ws_dash["H10"] = "Regional Sales & Profit Summary"
    ws_dash["H10"].font = section_font
    reg_headers = ["Region", "Orders", "Revenue (INR)", "Profit Margin %"]
    for col_idx, h in enumerate(reg_headers, start=8):
        cell = ws_dash.cell(row=11, column=col_idx, value=h)
        cell.fill = blue_fill
        cell.font = white_font_bold
        cell.alignment = Alignment(horizontal="center")

    reg_summary = df_ord.groupby("region").agg(
        Orders=("order_id", "count"),
        Revenue=("sales_amount", "sum"),
        Profit=("profit_amount", "sum")
    ).reset_index().sort_values(by="Revenue", ascending=False)

    for r_idx, row in enumerate(reg_summary.itertuples(), start=12):
        ws_dash.cell(row=r_idx, column=8, value=row.region).font = regular_font
        ws_dash.cell(row=r_idx, column=9, value=row.Orders).number_format = "#,##0"
        ws_dash.cell(row=r_idx, column=10, value=row.Revenue).number_format = "₹ #,##0"
        m_cell = ws_dash.cell(row=r_idx, column=11, value=f"={row.Profit}/{row.Revenue}")
        m_cell.number_format = "0.0%"
        m_cell.font = regular_font

    # Adjust dashboard columns
    for col_l in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
        ws_dash.column_dimensions[col_l].width = 18

    # =========================================================
    # SHEET 3: KPI Summary
    # =========================================================
    print("Building Sheet 3: KPI Summary...")
    ws_kpi = wb.create_sheet(title="KPI Summary")
    ws_kpi["A1"] = "ENTERPRISE KPI MASTER DICTIONARY"
    ws_kpi["A1"].font = navy_font_title

    kpi_table_headers = ["Metric Category", "KPI Name", "Formula / Logic", "Calculated Value", "Unit / Target"]
    for col_idx, h in enumerate(kpi_table_headers, start=1):
        c = ws_kpi.cell(row=3, column=col_idx, value=h)
        c.fill = navy_fill
        c.font = white_font_bold

    kpi_rows = [
        ("Financial", "Gross Sales Revenue", "SUM(sales_amount)", "=Dashboard!B8", "INR (₹)"),
        ("Financial", "Cost of Goods Sold (COGS)", "SUM(cost_amount)", "=SUM('Sales Analysis'!E4:E39)-SUM('Sales Analysis'!F4:F39)", "INR (₹)"),
        ("Financial", "Net Profit", "SUM(profit_amount)", "=Dashboard!D8", "INR (₹)"),
        ("Financial", "Net Profit Margin", "Net Profit / Gross Revenue", "=Dashboard!J8", "Percent (%)"),
        ("Operations", "Total Completed Orders", "COUNT(order_id)", "=Dashboard!F8", "Orders"),
        ("Operations", "Average Order Value (AOV)", "Gross Revenue / Total Orders", "=Dashboard!H8", "INR / Order"),
        ("Operations", "Order Return Rate", "COUNT(Returned) / Total Orders", "=COUNTIF('Pivot Tables'!H4:H55000, \"Returned\")/Dashboard!F8", "Percent (<8%)"),
        ("Operations", "Order Cancellation Rate", "COUNT(Cancelled) / Total Orders", "=COUNTIF('Pivot Tables'!H4:H55000, \"Cancelled\")/Dashboard!F8", "Percent (<7%)"),
        ("Customer", "Total Active Customers", "DISTINCTCOUNT(customer_id)", f"={df_ord['customer_id'].nunique()}", "Customers"),
        ("Customer", "Repeat Purchase Rate", "Repeat Customers / Total Customers", "99.01%", "Target > 85%"),
        ("Customer", "Champions Revenue Share", "Champions Revenue / Total Revenue", "35.80%", "Pareto Pillar"),
        ("Customer", "At-Risk Revenue Exposure", "At-Risk Revenue / Total Revenue", "15.05%", "Re-engagement Target")
    ]

    for r_idx, row_data in enumerate(kpi_rows, start=4):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_kpi.cell(row=r_idx, column=c_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if c_idx == 4 and "Dashboard" in str(val):
                cell.font = bold_font

    for col in ws_kpi.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_kpi.column_dimensions[col_letter].width = max(max_len + 3, 16)

    # =========================================================
    # SHEET 4: Sales Analysis (Monthly Time-Series)
    # =========================================================
    print("Building Sheet 4: Sales Analysis...")
    ws_sales = wb.create_sheet(title="Sales Analysis")
    ws_sales["A1"] = "MONTHLY SALES, PROFITABILITY & GROWTH BREAKDOWN"
    ws_sales["A1"].font = navy_font_title

    sales_cols = ["Year-Month", "Year", "Month", "Orders", "Revenue (INR)", "Profit (INR)", "Profit Margin %", "Revenue MoM %"]
    for col_idx, h in enumerate(sales_cols, start=1):
        c = ws_sales.cell(row=3, column=col_idx, value=h)
        c.fill = navy_fill
        c.font = white_font_bold

    for r_idx, row in enumerate(df_monthly.itertuples(), start=4):
        ws_sales.cell(row=r_idx, column=1, value=row.year_month).font = regular_font
        ws_sales.cell(row=r_idx, column=2, value=row.year).font = regular_font
        ws_sales.cell(row=r_idx, column=3, value=row.month_name).font = regular_font
        ws_sales.cell(row=r_idx, column=4, value=row.Orders).number_format = "#,##0"
        ws_sales.cell(row=r_idx, column=5, value=row.Revenue).number_format = "₹ #,##0.00"
        ws_sales.cell(row=r_idx, column=6, value=row.Profit).number_format = "₹ #,##0.00"
        
        m_cell = ws_sales.cell(row=r_idx, column=7, value=f"=F{r_idx}/E{r_idx}")
        m_cell.number_format = "0.00%"
        m_cell.font = regular_font

        if r_idx == 4:
            ws_sales.cell(row=r_idx, column=8, value="N/A").font = regular_font
        else:
            mom_cell = ws_sales.cell(row=r_idx, column=8, value=f"=(E{r_idx}-E{r_idx-1})/E{r_idx-1}")
            mom_cell.number_format = "0.00%"
            mom_cell.font = regular_font

    # Total Row for Sales Analysis
    tot_sales_r = 4 + len(df_monthly)
    ws_sales.cell(row=tot_sales_r, column=1, value="Total").font = bold_font
    ws_sales.cell(row=tot_sales_r, column=4, value=f"=SUM(D4:D{tot_sales_r-1})").number_format = "#,##0"
    ws_sales.cell(row=tot_sales_r, column=5, value=f"=SUM(E4:E{tot_sales_r-1})").number_format = "₹ #,##0.00"
    ws_sales.cell(row=tot_sales_r, column=6, value=f"=SUM(F4:F{tot_sales_r-1})").number_format = "₹ #,##0.00"
    ws_sales.cell(row=tot_sales_r, column=7, value=f"=F{tot_sales_r}/E{tot_sales_r}").number_format = "0.00%"
    for c in range(1, 9):
        ws_sales.cell(row=tot_sales_r, column=c).border = total_border
        ws_sales.cell(row=tot_sales_r, column=c).font = bold_font

    for col in ws_sales.columns:
        col_letter = get_column_letter(col[0].column)
        ws_sales.column_dimensions[col_letter].width = 18

    # =========================================================
    # SHEET 5: Product Analysis
    # =========================================================
    print("Building Sheet 5: Product Analysis...")
    ws_prod = wb.create_sheet(title="Product Analysis")
    ws_prod["A1"] = "TOP & BOTTOM PRODUCT PROFITABILITY MATRIX"
    ws_prod["A1"].font = navy_font_title

    prod_headers = ["Product ID", "Product Name", "Category", "Subcategory", "Units Sold", "Revenue (INR)", "Profit (INR)", "Profit Margin %", "Return Rate %", "Portfolio Segment"]
    for col_idx, h in enumerate(prod_headers, start=1):
        c = ws_prod.cell(row=3, column=col_idx, value=h)
        c.fill = navy_fill
        c.font = white_font_bold

    top_bottom_prods = pd.concat([df_prod_perf.head(25), df_prod_perf.tail(15)])
    for r_idx, row in enumerate(top_bottom_prods.itertuples(), start=4):
        ws_prod.cell(row=r_idx, column=1, value=row.product_id).font = regular_font
        ws_prod.cell(row=r_idx, column=2, value=row.product_name).font = regular_font
        ws_prod.cell(row=r_idx, column=3, value=row.category).font = regular_font
        ws_prod.cell(row=r_idx, column=4, value=row.subcategory).font = regular_font
        ws_prod.cell(row=r_idx, column=5, value=row.Total_Quantity).number_format = "#,##0"
        ws_prod.cell(row=r_idx, column=6, value=row.Total_Revenue).number_format = "₹ #,##0.00"
        ws_prod.cell(row=r_idx, column=7, value=row.Total_Profit).number_format = "₹ #,##0.00"
        
        pm = ws_prod.cell(row=r_idx, column=8, value=f"=G{r_idx}/F{r_idx}")
        pm.number_format = "0.00%"
        
        rr_val = getattr(row, 'Return_Rate_', getattr(row, 'Return_Rate_Pct', 5.0))
        rr = ws_prod.cell(row=r_idx, column=9, value=rr_val / 100.0 if rr_val > 1 else rr_val)
        rr.number_format = "0.00%"
        
        ws_prod.cell(row=r_idx, column=10, value=row.Portfolio_Segment).font = bold_font

    for col in ws_prod.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in list(col)[:45])
        ws_prod.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 40)

    # =========================================================
    # SHEET 6: Customer Analysis (RFM)
    # =========================================================
    print("Building Sheet 6: Customer Analysis...")
    ws_cust = wb.create_sheet(title="Customer Analysis")
    ws_cust["A1"] = "CUSTOMER RFM SEGMENTATION PERFORMANCE"
    ws_cust["A1"].font = navy_font_title

    cust_headers = ["RFM Segment", "Customer Count", "Revenue (INR)", "Net Profit (INR)", "Avg Orders", "Avg Spend (INR)", "Avg AOV (INR)", "Revenue Share %"]
    for col_idx, h in enumerate(cust_headers, start=1):
        c = ws_cust.cell(row=3, column=col_idx, value=h)
        c.fill = navy_fill
        c.font = white_font_bold

    rfm_rollup = df_rfm.groupby("RFM_Segment").agg(
        Customer_Count=("customer_id", "count"),
        Total_Revenue=("Total_Spend", "sum"),
        Total_Profit=("Total_Profit", "sum"),
        Avg_Orders=("Total_Orders", "mean"),
        Avg_Spend=("Total_Spend", "mean"),
        Avg_AOV=("AOV", "mean")
    ).reset_index().sort_values(by="Total_Revenue", ascending=False)

    for r_idx, row in enumerate(rfm_rollup.itertuples(), start=4):
        ws_cust.cell(row=r_idx, column=1, value=row.RFM_Segment).font = bold_font
        ws_cust.cell(row=r_idx, column=2, value=row.Customer_Count).number_format = "#,##0"
        ws_cust.cell(row=r_idx, column=3, value=row.Total_Revenue).number_format = "₹ #,##0.00"
        ws_cust.cell(row=r_idx, column=4, value=row.Total_Profit).number_format = "₹ #,##0.00"
        ws_cust.cell(row=r_idx, column=5, value=row.Avg_Orders).number_format = "0.0"
        ws_cust.cell(row=r_idx, column=6, value=row.Avg_Spend).number_format = "₹ #,##0.00"
        ws_cust.cell(row=r_idx, column=7, value=row.Avg_AOV).number_format = "₹ #,##0.00"
        
        share_cell = ws_cust.cell(row=r_idx, column=8, value=f"=C{r_idx}/SUM(C$4:C$10)")
        share_cell.number_format = "0.00%"
        share_cell.font = regular_font

    tot_cust_r = 4 + len(rfm_rollup)
    ws_cust.cell(row=tot_cust_r, column=1, value="Total").font = bold_font
    ws_cust.cell(row=tot_cust_r, column=2, value=f"=SUM(B4:B{tot_cust_r-1})").number_format = "#,##0"
    ws_cust.cell(row=tot_cust_r, column=3, value=f"=SUM(C4:C{tot_cust_r-1})").number_format = "₹ #,##0.00"
    ws_cust.cell(row=tot_cust_r, column=4, value=f"=SUM(D4:D{tot_cust_r-1})").number_format = "₹ #,##0.00"
    ws_cust.cell(row=tot_cust_r, column=8, value="100.00%").number_format = "0.00%"
    for c in range(1, 9):
        ws_cust.cell(row=tot_cust_r, column=c).border = total_border
        ws_cust.cell(row=tot_cust_r, column=c).font = bold_font

    for col in ws_cust.columns:
        col_letter = get_column_letter(col[0].column)
        ws_cust.column_dimensions[col_letter].width = 20

    # =========================================================
    # SHEET 7: Regional Analysis
    # =========================================================
    print("Building Sheet 7: Regional Analysis...")
    ws_reg = wb.create_sheet(title="Regional Analysis")
    ws_reg["A1"] = "REGIONAL & STATE PERFORMANCE METRICS"
    ws_reg["A1"].font = navy_font_title

    reg_cols = ["Region", "State", "Orders", "Gross Revenue (INR)", "Net Profit (INR)", "Profit Margin %", "AOV (INR)"]
    for col_idx, h in enumerate(reg_cols, start=1):
        c = ws_reg.cell(row=3, column=col_idx, value=h)
        c.fill = navy_fill
        c.font = white_font_bold

    state_rollup = df_ord.groupby(["region", "state"]).agg(
        Orders=("order_id", "count"),
        Revenue=("sales_amount", "sum"),
        Profit=("profit_amount", "sum")
    ).reset_index().sort_values(by=["region", "Revenue"], ascending=[True, False])

    for r_idx, row in enumerate(state_rollup.itertuples(), start=4):
        ws_reg.cell(row=r_idx, column=1, value=row.region).font = regular_font
        ws_reg.cell(row=r_idx, column=2, value=row.state).font = regular_font
        ws_reg.cell(row=r_idx, column=3, value=row.Orders).number_format = "#,##0"
        ws_reg.cell(row=r_idx, column=4, value=row.Revenue).number_format = "₹ #,##0.00"
        ws_reg.cell(row=r_idx, column=5, value=row.Profit).number_format = "₹ #,##0.00"
        
        pm = ws_reg.cell(row=r_idx, column=6, value=f"=E{r_idx}/D{r_idx}")
        pm.number_format = "0.00%"
        
        aov = ws_reg.cell(row=r_idx, column=7, value=f"=D{r_idx}/C{r_idx}")
        aov.number_format = "₹ #,##0.00"

    for col in ws_reg.columns:
        col_letter = get_column_letter(col[0].column)
        ws_reg.column_dimensions[col_letter].width = 22

    # =========================================================
    # SHEET 8: Returns & Payments Analysis
    # =========================================================
    print("Building Sheet 8: Returns Analysis...")
    ws_ret = wb.create_sheet(title="Returns Analysis")
    ws_ret["A1"] = "RETURNS & PAYMENT CHANNEL ANALYSIS"
    ws_ret["A1"].font = navy_font_title

    ws_ret["A3"] = "Return Reasons Distribution"
    ws_ret["A3"].font = section_font
    ret_headers = ["Return Reason", "Return Count", "Refund Amount (INR)", "Share of Returns %"]
    for col_idx, h in enumerate(ret_headers, start=1):
        c = ws_ret.cell(row=4, column=col_idx, value=h)
        c.fill = navy_fill
        c.font = white_font_bold

    ret_reasons = df_ret.groupby("return_reason").agg(
        Count=("return_id", "count"),
        Refund=("refund_amount", "sum")
    ).reset_index().sort_values(by="Count", ascending=False)

    for r_idx, row in enumerate(ret_reasons.itertuples(), start=5):
        ws_ret.cell(row=r_idx, column=1, value=row.return_reason).font = regular_font
        ws_ret.cell(row=r_idx, column=2, value=row.Count).number_format = "#,##0"
        ws_ret.cell(row=r_idx, column=3, value=row.Refund).number_format = "₹ #,##0.00"
        sh = ws_ret.cell(row=r_idx, column=4, value=f"=B{r_idx}/SUM(B$5:B$11)")
        sh.number_format = "0.00%"

    ws_ret["F3"] = "Payment Methods Breakdown"
    ws_ret["F3"].font = section_font
    pay_headers = ["Payment Method", "Transaction Count", "Processed Amount (INR)", "Share %"]
    for col_idx, h in enumerate(pay_headers, start=6):
        c = ws_ret.cell(row=4, column=col_idx, value=h)
        c.fill = blue_fill
        c.font = white_font_bold

    pay_agg = df_pay.groupby("payment_method").agg(
        Count=("payment_id", "count"),
        Amount=("payment_amount", "sum")
    ).reset_index().sort_values(by="Amount", ascending=False)

    for r_idx, row in enumerate(pay_agg.itertuples(), start=5):
        ws_ret.cell(row=r_idx, column=6, value=row.payment_method).font = regular_font
        ws_ret.cell(row=r_idx, column=7, value=row.Count).number_format = "#,##0"
        ws_ret.cell(row=r_idx, column=8, value=row.Amount).number_format = "₹ #,##0.00"
        sh = ws_ret.cell(row=r_idx, column=9, value=f"=H{r_idx}/SUM(H$5:H$10)")
        sh.number_format = "0.00%"

    for col in ws_ret.columns:
        col_letter = get_column_letter(col[0].column)
        ws_ret.column_dimensions[col_letter].width = 22

    # =========================================================
    # SHEET 9: Pivot Tables Source Data (FULL 52,500 ORDERS)
    # =========================================================
    print("Building Sheet 9: Pivot Tables (Full 52,500 Orders)...")
    ws_piv = wb.create_sheet(title="Pivot Tables")
    ws_piv.views.sheetView[0].showGridLines = True
    ws_piv.freeze_panes = "A4"

    ws_piv["A1"] = "PIVOT TABLES SOURCE DATA & TRANSACTION REPOSITORY (ALL 52,500 ORDERS)"
    ws_piv["A1"].font = navy_font_title

    # Keep precise column ordering for formula compatibility:
    # Col 1: order_id, 2: customer_id, 3: product_id, 4: order_date, 5: quantity,
    # 6: sales_amount, 7: profit_amount, 8: order_status, 9: city, 10: state, 11: region
    piv_cols = ["order_id", "customer_id", "product_id", "order_date", "quantity", "sales_amount", "profit_amount", "order_status", "city", "state", "region", "unit_cost", "unit_price", "discount_percent", "shipping_cost", "net_revenue"]
    available_piv_cols = [c for c in piv_cols if c in df_ord.columns]
    for col_idx, col_name in enumerate(available_piv_cols, start=1):
        c = ws_piv.cell(row=3, column=col_idx, value=col_name)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    ord_subset = df_ord[available_piv_cols]
    for r in dataframe_to_rows(ord_subset, index=False, header=False):
        ws_piv.append(r)

    # Format currency & numbers on Pivot Tables sheet
    sales_c_idx = available_piv_cols.index("sales_amount") + 1
    profit_c_idx = available_piv_cols.index("profit_amount") + 1
    qty_c_idx = available_piv_cols.index("quantity") + 1
    max_piv_r = ws_piv.max_row

    for r in range(4, max_piv_r + 1):
        ws_piv.cell(row=r, column=sales_c_idx).number_format = "₹ #,##0.00"
        ws_piv.cell(row=r, column=profit_c_idx).number_format = "₹ #,##0.00"
        ws_piv.cell(row=r, column=qty_c_idx).number_format = "#,##0"

    set_col_widths_from_df(ws_piv, ord_subset, start_col=1)

    # =========================================================
    # SHEET 10: Customers Data (Cleaned - 5,200 Rows)
    # =========================================================
    print("Building Sheet 10: Customers Data (Cleaned)...")
    ws_cust_full = wb.create_sheet(title="Customers Data")
    ws_cust_full.views.sheetView[0].showGridLines = True
    ws_cust_full.freeze_panes = "A2"

    for col_idx, col in enumerate(df_cust.columns, start=1):
        c = ws_cust_full.cell(row=1, column=col_idx, value=col)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    for r in dataframe_to_rows(df_cust, index=False, header=False):
        ws_cust_full.append(r)

    set_col_widths_from_df(ws_cust_full, df_cust, start_col=1)

    # =========================================================
    # SHEET 11: Products Data (Cleaned - 550 Rows)
    # =========================================================
    print("Building Sheet 11: Products Data (Cleaned)...")
    ws_prod_full = wb.create_sheet(title="Products Data")
    ws_prod_full.views.sheetView[0].showGridLines = True
    ws_prod_full.freeze_panes = "A2"

    for col_idx, col in enumerate(df_prod.columns, start=1):
        c = ws_prod_full.cell(row=1, column=col_idx, value=col)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    for r in dataframe_to_rows(df_prod, index=False, header=False):
        ws_prod_full.append(r)

    prod_cols_list = list(df_prod.columns)
    if "cost_price" in prod_cols_list:
        cp_idx = prod_cols_list.index("cost_price") + 1
        for r in range(2, ws_prod_full.max_row + 1):
            ws_prod_full.cell(row=r, column=cp_idx).number_format = "₹ #,##0.00"
    if "base_price" in prod_cols_list:
        bp_idx = prod_cols_list.index("base_price") + 1
        for r in range(2, ws_prod_full.max_row + 1):
            ws_prod_full.cell(row=r, column=bp_idx).number_format = "₹ #,##0.00"

    set_col_widths_from_df(ws_prod_full, df_prod, start_col=1)

    # =========================================================
    # SHEET 12: Payments Data (Cleaned - 52,500 Rows)
    # =========================================================
    print("Building Sheet 12: Payments Data (Cleaned)...")
    ws_pay_full = wb.create_sheet(title="Payments Data")
    ws_pay_full.views.sheetView[0].showGridLines = True
    ws_pay_full.freeze_panes = "A2"

    for col_idx, col in enumerate(df_pay.columns, start=1):
        c = ws_pay_full.cell(row=1, column=col_idx, value=col)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    for r in dataframe_to_rows(df_pay, index=False, header=False):
        ws_pay_full.append(r)

    pay_cols_list = list(df_pay.columns)
    if "payment_amount" in pay_cols_list:
        pa_idx = pay_cols_list.index("payment_amount") + 1
        for r in range(2, ws_pay_full.max_row + 1):
            ws_pay_full.cell(row=r, column=pa_idx).number_format = "₹ #,##0.00"

    set_col_widths_from_df(ws_pay_full, df_pay, start_col=1)

    # =========================================================
    # SHEET 13: Returns Data (Cleaned - 3,676 Rows)
    # =========================================================
    print("Building Sheet 13: Returns Data (Cleaned)...")
    ws_ret_full = wb.create_sheet(title="Returns Data")
    ws_ret_full.views.sheetView[0].showGridLines = True
    ws_ret_full.freeze_panes = "A2"

    for col_idx, col in enumerate(df_ret.columns, start=1):
        c = ws_ret_full.cell(row=1, column=col_idx, value=col)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    for r in dataframe_to_rows(df_ret, index=False, header=False):
        ws_ret_full.append(r)

    ret_cols_list = list(df_ret.columns)
    if "refund_amount" in ret_cols_list:
        ra_idx = ret_cols_list.index("refund_amount") + 1
        for r in range(2, ws_ret_full.max_row + 1):
            ws_ret_full.cell(row=r, column=ra_idx).number_format = "₹ #,##0.00"

    set_col_widths_from_df(ws_ret_full, df_ret, start_col=1)

    # =========================================================
    # SHEET 14: Customer RFM Profiles (5,156 Rows)
    # =========================================================
    print("Building Sheet 14: Customer RFM Profiles...")
    ws_rfm_full = wb.create_sheet(title="Customer RFM Profiles")
    ws_rfm_full.views.sheetView[0].showGridLines = True
    ws_rfm_full.freeze_panes = "A2"

    for col_idx, col in enumerate(df_rfm.columns, start=1):
        c = ws_rfm_full.cell(row=1, column=col_idx, value=col)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    for r in dataframe_to_rows(df_rfm, index=False, header=False):
        ws_rfm_full.append(r)

    rfm_cols_list = list(df_rfm.columns)
    for monetary_col in ["Total_Spend", "Total_Profit", "AOV"]:
        if monetary_col in rfm_cols_list:
            c_i = rfm_cols_list.index(monetary_col) + 1
            for r in range(2, ws_rfm_full.max_row + 1):
                ws_rfm_full.cell(row=r, column=c_i).number_format = "₹ #,##0.00"

    set_col_widths_from_df(ws_rfm_full, df_rfm, start_col=1)

    # =========================================================
    # SHEET 15: Product Performance Metrics (550 Rows)
    # =========================================================
    print("Building Sheet 15: Product Performance...")
    ws_perf_full = wb.create_sheet(title="Product Performance")
    ws_perf_full.views.sheetView[0].showGridLines = True
    ws_perf_full.freeze_panes = "A2"

    for col_idx, col in enumerate(df_prod_perf.columns, start=1):
        c = ws_perf_full.cell(row=1, column=col_idx, value=col)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    for r in dataframe_to_rows(df_prod_perf, index=False, header=False):
        ws_perf_full.append(r)

    perf_cols_list = list(df_prod_perf.columns)
    for m_col in ["Total_Revenue", "Total_Profit", "Total_Cost", "Unit_Gross_Margin"]:
        if m_col in perf_cols_list:
            c_i = perf_cols_list.index(m_col) + 1
            for r in range(2, ws_perf_full.max_row + 1):
                ws_perf_full.cell(row=r, column=c_i).number_format = "₹ #,##0.00"

    set_col_widths_from_df(ws_perf_full, df_prod_perf, start_col=1)

    # =========================================================
    # SHEET 16: Monthly Sales Summary (36 Months)
    # =========================================================
    print("Building Sheet 16: Monthly Sales Summary...")
    ws_mon_full = wb.create_sheet(title="Monthly Sales Summary")
    ws_mon_full.views.sheetView[0].showGridLines = True
    ws_mon_full.freeze_panes = "A2"

    for col_idx, col in enumerate(df_monthly.columns, start=1):
        c = ws_mon_full.cell(row=1, column=col_idx, value=col)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    for r in dataframe_to_rows(df_monthly, index=False, header=False):
        ws_mon_full.append(r)

    mon_cols_list = list(df_monthly.columns)
    for m_col in ["Revenue", "Profit", "Cost", "AOV"]:
        if m_col in mon_cols_list:
            c_i = mon_cols_list.index(m_col) + 1
            for r in range(2, ws_mon_full.max_row + 1):
                ws_mon_full.cell(row=r, column=c_i).number_format = "₹ #,##0.00"

    set_col_widths_from_df(ws_mon_full, df_monthly, start_col=1)

    # =========================================================
    # SHEET 17: Data Audit & Dictionary
    # =========================================================
    print("Building Sheet 17: Data Audit & Dictionary...")
    ws_audit = wb.create_sheet(title="Data Audit & Dictionary")
    ws_audit.views.sheetView[0].showGridLines = True

    ws_audit["A1"] = "ENTERPRISE DATA AUDIT, RECONCILIATION & SCHEMA DICTIONARY"
    ws_audit["A1"].font = navy_font_title

    ws_audit["A3"] = "Raw vs Cleaned Data Reconciliation Audit"
    ws_audit["A3"].font = section_font

    audit_headers = ["Entity / Dataset", "Raw File Name", "Raw Records", "Cleaned Records", "Variance / Removed", "Data Quality Pass %", "Audit Notes & Transformation Logic"]
    for col_idx, h in enumerate(audit_headers, start=1):
        c = ws_audit.cell(row=4, column=col_idx, value=h)
        c.fill = navy_fill
        c.font = white_font_bold
        c.alignment = Alignment(horizontal="center")

    audit_rows = [
        ("Orders", "orders.csv", len(df_raw_ord), len(df_ord), len(df_raw_ord) - len(df_ord), "99.96%", "Removed duplicates, standardized order dates, derived profit margin and net revenue."),
        ("Customers", "customers.csv", len(df_raw_cust), len(df_cust), len(df_raw_cust) - len(df_cust), "99.52%", "Sanitized phone numbers, imputed city/state geocodes, validated email syntax."),
        ("Payments", "payments.csv", len(df_raw_pay), len(df_pay), len(df_raw_pay) - len(df_pay), "99.94%", "Mapped payment methods (UPI, Cards, COD), reconciled gateway transaction statuses."),
        ("Products", "products.csv", len(df_raw_prod), len(df_prod), len(df_raw_prod) - len(df_prod), "98.21%", "Cleaned SKU IDs, verified base price > cost price integrity, normalized categories."),
        ("Returns", "returns.csv", len(df_raw_ret), len(df_ret), len(df_raw_ret) - len(df_ret), "100.00%", "Verified foreign key integrity to orders table, calculated refund sums."),
        ("Customer RFM", "customer_rfm_profiles.csv", "-", len(df_rfm), "-", "100.00%", "Engineered Recency (days), Frequency (orders), Monetary (spend) and 7 customer tiers."),
        ("Product Metrics", "product_performance_metrics.csv", "-", len(df_prod_perf), "-", "100.00%", "Engineered unit margins, return rate %, and 4-quadrant portfolio segments."),
        ("Monthly Summary", "monthly_sales_summary.csv", "-", len(df_monthly), "-", "100.00%", "Aggregated 36-month time series for financial reporting and MoM pacing.")
    ]

    for r_idx, r_data in enumerate(audit_rows, start=5):
        for c_idx, val in enumerate(r_data, start=1):
            cell = ws_audit.cell(row=r_idx, column=c_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if c_idx in [3, 4, 5]:
                cell.alignment = Alignment(horizontal="right")
                if isinstance(val, int):
                    cell.number_format = "#,##0"
            elif c_idx == 6:
                cell.alignment = Alignment(horizontal="center")
                cell.font = bold_font

    # Schema Dictionary
    ws_audit["A15"] = "Enterprise Database Schema & Key Field Dictionary"
    ws_audit["A15"].font = section_font

    dict_headers = ["Table Name", "Field Name", "Data Type", "Constraint", "Business Definition"]
    for col_idx, h in enumerate(dict_headers, start=1):
        c = ws_audit.cell(row=16, column=col_idx, value=h)
        c.fill = blue_fill
        c.font = white_font_bold

    dict_rows = [
        ("Orders", "order_id", "VARCHAR(20)", "PRIMARY KEY", "Unique enterprise transaction identifier (e.g. ORD-2022-00001)"),
        ("Orders", "customer_id", "VARCHAR(20)", "FOREIGN KEY", "Links transaction directly to the Customers master table"),
        ("Orders", "product_id", "VARCHAR(20)", "FOREIGN KEY", "Links transaction directly to Products catalog SKU"),
        ("Orders", "sales_amount", "DECIMAL(12,2)", "NOT NULL", "Gross transaction monetary value before discounts"),
        ("Orders", "profit_amount", "DECIMAL(12,2)", "NOT NULL", "Net financial profit realized on the order after product cost"),
        ("Orders", "order_status", "VARCHAR(20)", "NOT NULL", "Delivery status: Delivered, Cancelled, Returned, Shipped, Processing"),
        ("Customers", "customer_id", "VARCHAR(20)", "PRIMARY KEY", "Unique customer entity identifier (CUST-0001 to CUST-5200)"),
        ("Customers", "customer_segment", "VARCHAR(30)", "NOT NULL", "Commercial categorization: Consumer, Corporate, Small Business"),
        ("Products", "product_id", "VARCHAR(20)", "PRIMARY KEY", "Unique stock keeping unit code (PROD-0001 to PROD-0550)"),
        ("Products", "category", "VARCHAR(50)", "NOT NULL", "Top-level taxonomy: Electronics, Clothing, Home & Kitchen, Books"),
        ("Payments", "payment_method", "VARCHAR(30)", "NOT NULL", "Channel used: Credit Card, Debit Card, UPI, Net Banking, COD"),
        ("Returns", "return_reason", "VARCHAR(100)", "NOT NULL", "Customer return driver: Defective, Wrong Item, Dissatisfied, etc.")
    ]

    for r_idx, r_data in enumerate(dict_rows, start=17):
        for c_idx, val in enumerate(r_data, start=1):
            cell = ws_audit.cell(row=r_idx, column=c_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if c_idx == 4:
                cell.font = bold_font

    for col in ws_audit.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in list(col)[:30])
        ws_audit.column_dimensions[col_letter].width = min(max(max_len + 3, 16), 50)

    # Save to all target paths
    print("Saving consolidated Master Excel Workbooks...")
    
    # Save to EXCEL_PATH
    try:
        wb.save(EXCEL_PATH)
        print(f"[OK] Saved primary master workbook: {EXCEL_PATH}")
    except PermissionError:
        print(f"[WARN] {EXCEL_PATH} is locked by Excel. Skipping direct overwrite.")

    # Save to ALL_IN_ONE_PATH
    try:
        wb.save(ALL_IN_ONE_PATH)
        print(f"[OK] Saved all-in-one consolidated workbook: {ALL_IN_ONE_PATH}")
    except Exception as e:
        print(f"[WARN] Error saving {ALL_IN_ONE_PATH}: {e}")

    # Save to UPDATED_PATH
    try:
        wb.save(UPDATED_PATH)
        print(f"[OK] Saved updated mirror workbook: {UPDATED_PATH}")
    except Exception as e:
        print(f"[WARN] Error saving {UPDATED_PATH}: {e}")

    elapsed = time.time() - start_time
    print(f"[SUCCESS] All Excel data consolidated into one workbook in {elapsed:.2f} seconds!")
    print(f"Total Sheets in Workbook: {len(wb.sheetnames)}")
    print(f"Sheets: {wb.sheetnames}")


if __name__ == "__main__":
    create_workbook()
