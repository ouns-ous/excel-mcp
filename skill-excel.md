---
name: excel-master
description: "Expert guidance, decision trees, and best practices for automating Microsoft Excel (.xlsx) using the Excel MCP Server (223 Tools). Triggers: creating spreadsheets, financial models, executive KPI dashboards, charts, pivot tables, slicers, what-if sensitivity analysis, data cleaning, formula auditing, PDF export, and fast headless OOXML operations."
---

# Microsoft Excel Automation & Data Engineering (Excel MCP)

This skill equips AI assistants (Antigravity, Claude Desktop, Cursor, etc.) with a clear blueprint on **WHEN**, **WHERE**, and **HOW** to orchestrate the **223 tools** of the `excel-mcp` server.

---

## 🎯 Architecture: Two Powerful Engines

| Feature | Live Desktop COM Engine | Headless OOXML Engine (`openpyxl`) |
| :--- | :--- | :--- |
| **Tool Prefix** | `excel_*` (interactive COM calls) | `excel_*_headless`, `excel_create_from_markdown`, `excel_compare_two_workbooks`, etc. |
| **Execution** | Real-time desktop GUI (`win32com.client`) | In-memory OOXML stream processing |
| **When to Use** | When the user wants real-time visual feedback, interactive dashboards, Pivot Tables, Slicers, Goal Seek, native PDF export, or live formula evaluation. | When speed is critical, background batch processing without Excel GUI, comparing two files, or transforming Markdown/JSON/CSV into spreadsheets. |
| **Prerequisites** | Microsoft Excel installed on Windows | None (works standalone in Python) |

---

## 🧭 Intent Matrix: Which Tool to Use When?

### 1. Starting a New Spreadsheet
- **Live Interactive**: `excel_new_workbook(sheet_name="Dashboard")` -> creates visible workbook and focuses window.
- **From Text / Markdown Table**: `excel_create_from_markdown(markdown_table, output_path)` -> instant styled OOXML spreadsheet.
- **From JSON Data**: `excel_json_to_sheet(json_data, output_path)` -> converts structured API response to styled table.
- **From CSV**: `excel_csv_to_excel(csv_path, output_path)` -> styled workbook with auto-fitted columns.

### 2. Building Executive Dashboards
- **KPI Summary Cards**: `excel_create_kpi_card(start_cell="B2", title="TOTAL REVENUE", value="$4.8M", subtitle="+18.5% YoY", accent_color="#2563EB")`.
- **Compact Metric Tiles**: `excel_add_metric_tile(cell="E2", label="Active Subscriptions", value="14,250", delta="+8.2%", status="up")`.
- **Callout & Alert Banners**: `excel_add_callout_box(cell="B6", text="Q3 revenue exceeded forecast by 14%.", box_type="info")`.
- **Project Timelines**: `excel_create_gantt_block(start_cell="B10", tasks=[...], weeks_count=8)`.
- **Interactive Checkboxes**: `excel_add_checkbox(cell="B18", caption="Reviewed by Finance", linked_cell="C18")`.

### 3. Data Entry, Formulas & Formatting
- **Write Bulk Data**: `excel_write_range(start_cell="B8", data=[["Product", "Q1", "Q2"], ["App", 100, 150]])`.
- **Insert Formulas**: `excel_set_formula(cell="D9", formula="=SUM(B9:C9)")`.
- **Apply Number Formats**: `excel_set_number_format(range_address="C9:D50", format_code="$#,##0.00")`.
- **Wrap & Align Text**: `excel_wrap_text(range_address="B8:E8", wrap=True)`.
- **Zebra Row Shading**: `excel_apply_zebra_stripes(range_address="B9:E30", color_even="#F8FAFC", color_odd="#FFFFFF")`.
- **AutoFit Dimensions**: `excel_autofit_all()` or `excel_autofit_columns(columns="B:E")`.

### 4. Visualizations & Charting
- **Insert Native Chart**: `excel_add_chart(chart_type="column_clustered", source_range="B8:D12", title="Quarterly Growth", left=400, top=120)`.
- **Customize Chart Title & Style**: `excel_set_chart_title(chart_name="Chart 1", title="2026 Revenue Trajectory", font_size=14, bold=True)`.
- **Add Legend**: `excel_set_chart_legend(chart_name="Chart 1", position="bottom", show_legend=True)`.
- **Add Secondary Y-Axis**: `excel_set_chart_secondary_axis(chart_name="Chart 1", series_index=2)`.
- **Add Trendline**: `excel_add_chart_trendline(chart_name="Chart 1", series_index=1, trendline_type="linear")`.
- **Mini In-Cell Trends**: `excel_add_sparkline(data_range="C9:G9", location_cell="H9", sparkline_type="line")`.
- **Export Chart Image**: `excel_export_chart_as_image(chart_name="Chart 1", output_image_path="chart.png")`.

### 5. Pivot Tables & Interactive Filtering
- **Create Pivot Table**: `excel_create_pivot_table(source_range="A1:E100", target_cell="G2", row_fields=["Region"], col_fields=["Quarter"], data_fields=["Sales"])`.
- **Add Visual Slicers**: `excel_add_table_slicer(table_name="SalesTable", column_name="Region", left=550, top=80)`.
- **Change Calculation**: `excel_change_pivot_field_calculation(pivot_name="PivotTable1", data_field_name="Sum of Sales", calculation="average")`.
- **Switch Pivot Layout**: `excel_set_pivot_layout(pivot_name="PivotTable1", layout="tabular")`.

### 6. What-If Financial & Sensitivity Analysis
- **Goal Seek**: `excel_goal_seek(target_cell="C15", target_value=100000.0, changing_cell="C5")` -> finds the exact unit price needed to hit \$100k profit.
- **2-Variable Sensitivity Matrix**: `excel_create_two_variable_data_table(table_range="B20:G25", row_input_cell="C5", col_input_cell="C6")`.
- **Scenario Management**:
  - `excel_add_scenario(name="Best Case", changing_cells="C5:C7", values=[150, 2000, 0.05])`.
  - `excel_show_scenario(name="Best Case")`.

### 7. Sheet Polishing & Layout Aesthetics
- **Clean Dashboard Look**: `excel_toggle_gridlines(show=False)` (gives a clean, modern SaaS software feel).
- **Freeze Panes**: `excel_freeze_top_row()` or `excel_freeze_custom_panes(row=9, column=2)`.
- **Color Tabs**: `excel_set_tab_color(hex_color="#10B981")`.
- **Process Flows**: `excel_add_flowchart_process_block(...)` & `excel_add_connector_arrow(...)`.

### 8. Auditing & Troubleshooting
- **Find Formula Errors**: `excel_get_formula_errors()` -> returns all `#DIV/0!`, `#N/A`, `#VALUE!`.
- **Inspect Precedents/Dependents**: `excel_get_precedents(cell_address="D15")` and `excel_get_dependents(cell_address="C5")`.
- **Trace Arrows**: `excel_trace_precedents(range_address="D15")`.
- **List Merged Areas**: `excel_get_merged_ranges()`.

### 9. Printing & PDF Export
- **Print Titles (Repeat Header Rows)**: `excel_set_repeat_rows_cols(repeat_rows="1:8")`.
- **Margins & Page Fit**: `excel_set_sheet_margins(top_inches=0.5, bottom_inches=0.5)` and `excel_set_sheet_fit_to_pages(fit_to_wide=1, fit_to_tall=1)`.
- **Export to PDF**: `excel_export_pdf(output_path="C:\\Reports\\Report.pdf", sheet_only=True)`.

---

## 🎨 Enterprise Design System for Spreadsheets

Follow these rules for stunning, professional outputs:

1. **Header Row Styling**:
   - Dark Navy (`#1E293B`) or Slate Blue (`#1F4E79`).
   - Pure White text (`#FFFFFF`), Bold (`Font.Bold = True`).
   - Centered alignment with vertical centering.
2. **Data Rows**:
   - Alternating Zebra rows: `#FFFFFF` and `#F8FAFC`.
   - Text left-aligned; numbers and currency right-aligned; status badges centered.
3. **Accent Colors**:
   - Primary: `#2563EB` (Royal Blue)
   - Success / Positive: `#10B981` (Emerald)
   - Warning: `#F59E0B` (Amber)
   - Danger: `#EF4444` (Rose / Red)
4. **Number Formatting Standards**:
   - Currency: `"$#,##0.00"` or `"$#,##0"` (integers).
   - Percentages: `"0.0%"` or `"0%"`.
   - Dates: `"YYYY-MM-DD"`.
   - Accounting: `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`.

---

## 🚀 End-to-End Workflow Examples

### Workflow 1: Executive KPI Dashboard with Chart & PDF
```python
# 1. Create a clean workbook
excel_new_workbook(sheet_name="Executive Dashboard")
excel_toggle_gridlines(show=False)

# 2. Add Title & KPI Cards
excel_write_cell(cell="B2", value="EXECUTIVE REVENUE OVERVIEW")
excel_format_range(range_address="B2", font_size=18, bold=True, font_color="#1E293B")

excel_create_kpi_card(start_cell="B4", title="ARR REVENUE", value="$12.4M", subtitle="+24% YoY", accent_color="#2563EB")
excel_create_kpi_card(start_cell="E4", title="NET RETENTION", value="118.2%", subtitle="+3.4% YoY", accent_color="#10B981")
excel_create_kpi_card(start_cell="H4", title="CUSTOMER CAC", value="$420", subtitle="-12% improved", accent_color="#6366F1")

# 3. Add Data Table
data = [
    ["Quarter", "New ARR ($)", "Expansion ($)", "Churn ($)", "Total ARR ($)"],
    ["Q1 2026", 850000, 320000, -80000, "=SUM(C9:E9)"],
    ["Q2 2026", 920000, 340000, -75000, "=SUM(C10:E10)"],
    ["Q3 2026", 1100000, 410000, -90000, "=SUM(C11:E11)"],
    ["Q4 2026 (Est)", 1250000, 480000, -85000, "=SUM(C12:E12)"]
]
excel_write_range(start_cell="B8", data=data)
excel_format_table_headers(range_address="B8:F8", bg_color="#1E293B", text_color="#FFFFFF")
excel_set_number_format(range_address="C9:F12", format_code="$#,##0")
excel_apply_zebra_stripes(range_address="B9:F12", color_even="#F8FAFC", color_odd="#FFFFFF")

# 4. Add Interactive Clustered Column Chart
excel_add_chart(
    chart_type="column_clustered",
    source_range="B8:D12",
    title="Quarterly Expansion vs New ARR",
    left=70,
    top=280,
    width=520,
    height=260
)
excel_set_chart_legend(chart_name="Chart 1", position="bottom", show_legend=True)

# 5. Fit & Export to High-Resolution PDF
excel_set_sheet_fit_to_pages(fit_to_wide=1, fit_to_tall=1)
excel_export_pdf(output_path="C:\\Reports\\Executive_Dashboard.pdf")
```

### Workflow 2: Financial Sensitivity Analysis (Goal Seek & Data Table)
```python
# 1. Setup Base Financial Model
excel_write_range(start_cell="B2", data=[
    ["Parameter", "Value"],
    ["Unit Price", 50.0],
    ["Sales Volume", 10000],
    ["Fixed Costs", 150000],
    ["Variable Cost / Unit", 20.0],
    ["Total Revenue", "=C3*C4"],
    ["Total Cost", "=C5+(C6*C4)"],
    ["Operating Profit", "=C7-C8"]
])
excel_format_table_headers(range_address="B2:C2")

# 2. Goal Seek: What price is needed for $250k profit?
res = excel_goal_seek(target_cell="C9", target_value=250000.0, changing_cell="C3")

# 3. Create 2-Variable Sensitivity Matrix (Price vs Volume)
# Setup top row volumes: 8000, 9000, 10000, 11000, 12000
# Setup left column prices: 40, 45, 50, 55, 60
excel_create_two_variable_data_table(
    table_range="E2:J7",
    row_input_cell="C4", # Volume varies across columns
    col_input_cell="C3"  # Price varies down rows
)
```

### Workflow 3: Fast Headless Batch Operations
```python
# 1. Convert Markdown report into Excel without GUI
excel_create_from_markdown(markdown_table=md_str, output_path="output/sales.xlsx")

# 2. Audit workbook structure headlessly
audit = excel_audit_workbook(file_path="output/sales.xlsx")

# 3. Compare two versions of a financial model
diffs = excel_compare_two_workbooks("v1.xlsx", "v2.xlsx")
# diffs['cell_differences'] highlights all altered assumptions!
```
