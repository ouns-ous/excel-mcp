---
name: excel-master
description: "Expert guidance, decision trees, and best practices for automating Microsoft Excel (.xlsx) using the Excel MCP Server (243 Tools). Triggers: creating spreadsheets, financial models, executive KPI dashboards, charts, pivot tables, slicers, what-if sensitivity analysis, data cleaning, formula auditing, PDF export, data governance, and fast headless OOXML operations."
---

# Microsoft Excel Automation & Data Engineering (Excel MCP)

This skill equips AI assistants (Antigravity, Claude Desktop, Cursor, etc.) with a clear blueprint on **WHEN**, **WHERE**, and **HOW** to orchestrate the **243 tools** of the `excel-mcp` server.

---

## 🎯 Architecture: Two Powerful Engines

| Feature | Live Desktop COM Engine | Headless OOXML Engine (`openpyxl`) |
| :--- | :--- | :--- |
| **Tool Prefix** | `excel_*` (interactive COM calls) | `excel_*_headless`, `excel_create_from_markdown`, `excel_compare_two_workbooks`, etc. |
| **Execution** | Real-time desktop GUI (`win32com.client`) | In-memory OOXML stream processing |
| **When to Use** | When the user wants real-time visual feedback, interactive dashboards, Pivot Tables, Slicers, Goal Seek, native PDF export, or live formula evaluation. | When speed is critical, background batch processing without Excel GUI, comparing two files, transforming Markdown/JSON/CSV, data governance, or server environments. |
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

### 6. Data Cleaning & Text Intelligence (Headless)
- **Trim & Clean Control Chars**: `excel_clean_whitespace_headless(file_path="data.xlsx", range_address="A2:E500")` -> cleans spaces & invisible control chars.
- **Batch Case Conversion**: `excel_case_transform_headless(file_path="data.xlsx", range_address="A2:A500", mode="upper"|"lower"|"title"|"capitalize")`.
- **Regex Find & Replace**: `excel_regex_replace_headless(file_path="data.xlsx", range_address="C2:C500", pattern=r"INV-(\\d+)", replacement=r"REF-\\1")`.
- **Extract Text Patterns**: `excel_extract_by_pattern_headless(file_path="data.xlsx", source_col="D", target_col="E", pattern=r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\\.[a-zA-Z0-9-.]+)")`.
- **Fuzzy Deduplication**: `excel_fuzzy_deduplicate_headless(file_path="data.xlsx", key_column="A", threshold=0.85, action="flag"|"remove")`.
- **Data Privacy & Redaction**: `excel_mask_sensitive_data_headless(file_path="data.xlsx", range_address="B2:B500", mask_type="email"|"credit_card"|"phone"|"custom")`.
- **Detect Data Types & Schema**: `excel_detect_column_types_headless(file_path="data.xlsx")` -> infers column types, null counts, and samples.

### 7. Financial Modeling & Statistical Analysis (Headless)
- **Loan Amortization**: `excel_add_amortization_schedule_headless(file_path="loan.xlsx", principal=250000, annual_interest_rate=0.065, term_years=15)`.
- **CAGR Growth Analysis**: `excel_add_cagr_analysis_headless(file_path="model.xlsx", start_value=1200000, end_value=4800000, periods=4, output_cell="G2")`.
- **Multi-Year Depreciation**: `excel_add_depreciation_schedule_headless(file_path="capex.xlsx", asset_cost=80000, salvage_value=8000, useful_life_years=5, method="straight_line"|"declining_balance")`.
- **Correlation Matrix**: `excel_generate_correlation_matrix_headless(file_path="analytics.xlsx", columns=["B", "C", "D"], output_cell="F2")`.
- **Descriptive Statistics**: `excel_add_summary_statistics_headless(file_path="data.xlsx", column="C", output_cell="E2")`.
- **Break-Even Analysis**: `excel_calculate_break_even_headless(file_path="pricing.xlsx", fixed_costs=75000, variable_cost_per_unit=28.5, price_per_unit=65.0)`.
- **Waterfall Breakdown**: `excel_add_waterfall_breakdown_headless(file_path="arr.xlsx", categories=[...], amounts=[...], output_cell="B2")`.

### 8. Governance, Schemas & Multi-Workbook Management
- **Audit Trail Logging**: `excel_append_audit_log_headless(file_path="financials.xlsx", event_message="Model parameters revised", user_name="Analyst")`.
- **Batch Export to CSV**: `excel_batch_export_sheets_to_csv_headless(file_path="report.xlsx", output_dir="csv_dump/")`.
- **Merge Workbooks**: `excel_merge_multiple_workbooks_headless(file_paths=["q1.xlsx", "q2.xlsx"], output_path="master.xlsx", combine_mode="sheets"|"rows")`.
- **Partition & Split**: `excel_split_workbook_by_column_headless(file_path="global_sales.xlsx", split_column="A", output_dir="by_region/")`.
- **Schema Validation**: `excel_validate_schema_headless(file_path="import.xlsx", expected_columns=["ID", "Email", "Amount"], required_types={"Amount": "number"})`.
- **Markdown Data Dictionary**: `excel_generate_markdown_data_dictionary_headless(file_path="data.xlsx")`.

### 9. What-If Analysis, Auditing & Export
- **Goal Seek**: `excel_goal_seek(target_cell="C9", target_value=1000000, changing_cell="C3")`.
- **2-Variable Sensitivity Table**: `excel_create_two_variable_data_table(table_range="E2:J7", row_input_cell="C4", col_input_cell="C3")`.
- **Trace Precedents / Dependents**: `excel_trace_precedents(cell="C9")` or `excel_trace_dependents(cell="C3")`.
- **Export PDF (Auto-Fit)**: `excel_set_sheet_fit_to_pages(fit_to_wide=1, fit_to_tall=1)` then `excel_export_pdf(output_path="report.pdf")`.

---

## 🎨 Enterprise Design System: Excel Standards

When generating spreadsheets, apply the following enterprise standards:

### 1. Palette & Hierarchy
- **Header Fill**: Dark Navy `#1F4E79` or Slate `#1E293B` with bold white text `#FFFFFF`.
- **KPI Accent**: Royal Blue `#2563EB`, Emerald `#10B981`, Purple `#7C3AED`, or Amber `#D97706`.
- **Zebra Striping**: Even rows `#F8FAFC` (very light slate), Odd rows `#FFFFFF`.
- **Total Rows**: Top thin border `#94A3B8`, bottom double border `#1E293B`, bold font.

### 2. Number Formatting
- **Currency**: `"$#,##0"` (integers) or `"$#,##0.00"` (exact decimals).
- **Percentages**: `"0.0%"` or `"0.00%"`.
- **Dates**: `"YYYY-MM-DD"`.
- **Integers**: `#,##0`.

---

## ⚡ End-to-End Orchestration Workflows

### Workflow 1: Executive KPI Dashboard (Live Desktop COM)
```python
# 1. Initialize Workbook
excel_new_workbook(sheet_name="Executive Dashboard")
excel_set_view(show_gridlines=True)

# 2. Write Title & KPI Summary Cards
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

### Workflow 2: Financial Sensitivity & Break-Even Modeling
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
excel_create_two_variable_data_table(
    table_range="E2:J7",
    row_input_cell="C4", # Volume varies across columns
    col_input_cell="C3"  # Price varies down rows
)
```

### Workflow 3: Fast Headless Data Quality & Governance Pipeline
```python
# 1. Clean dirty data and remove invisible control characters
excel_clean_whitespace_headless(file_path="input_raw.xlsx", range_address="A2:H1000")

# 2. Deduplicate approximately spelled customer names
excel_fuzzy_deduplicate_headless(file_path="input_raw.xlsx", key_column="B", threshold=0.85, action="flag")

# 3. Mask PII/sensitive emails for compliance
excel_mask_sensitive_data_headless(file_path="input_raw.xlsx", range_address="D2:D1000", mask_type="email")

# 4. Validate schema compliance
schema_report = excel_validate_schema_headless(
    file_path="input_raw.xlsx",
    expected_columns=["ID", "Customer", "Email", "Revenue"],
    required_types={"Revenue": "number"}
)

# 5. Append audit log entry
excel_append_audit_log_headless(
    file_path="input_raw.xlsx",
    event_message="Data cleaning, deduplication, and PII masking pipeline executed.",
    user_name="DataOps_Agent"
)
```
