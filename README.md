# Microsoft Excel Live COM Desktop Automation & OOXML MCP Server 📊⚡

A comprehensive **Model Context Protocol (MCP)** server for Windows enabling AI coding assistants (**Antigravity**, **Claude Desktop**, **Cursor**, etc.) to automate **Microsoft Excel** in real-time desktop sessions with live visual feedback, paired with a fast headless OOXML (`openpyxl`) engine.

---

## 🌟 Highlights

- **Live Windows Desktop Automation**: Real-time workbook creation, sheet switching, cell reading/writing, formula evaluation, and formatting via `win32com.client`.
- **Fast Headless OOXML Engine**: Create, transform, and extract `.xlsx` files headlessly without requiring desktop Excel.
- **Executive Dashboards & KPI Cards**: Pre-styled executive metric blocks with trend badges and clean modern palettes.
- **Formulas & Number Formatting**: Full support for Excel formulas (`=SUM`, `=AVERAGE`, `=VLOOKUP`, etc.) and number formatting (currency, percent, dates).
- **Native Charts**: Interactive Column, Bar, Line, Pie, Doughnut, and Scatter charts inserted directly from cell data.
- **1-Click PDF Export**: Fast high-resolution PDF generation using Excel's native engine.
- **Markdown & JSON Interchange**: Instant bi-directional conversion between Markdown tables / JSON and Excel spreadsheets.

---

## 🛠️ Complete List of Tools (81 Tools)

### 1. Workbook & Window Management
- `excel_status()`: Get current running state, open workbooks count, active sheet, active cell, and used range.
- `excel_new_workbook(sheet_name)`: Create a new visible workbook.
- `excel_open(file_path)`: Open an existing workbook (`.xlsx`, `.xls`, `.xlsm`, `.csv`).
- `excel_save(file_path)`: Save the active workbook.
- `excel_export_pdf(output_path, sheet_only)`: Export workbook or active sheet directly to PDF.
- `excel_close(save_changes)`: Close the active workbook.
- `excel_focus_window()`: Bring Microsoft Excel to foreground and maximize.

### 2. Sheet Operations & Tab Styling
- `excel_list_sheets()`: List all worksheets with indices and visibility status.
- `excel_add_sheet(name, after)`: Add a new worksheet.
- `excel_activate_sheet(name)`: Switch active focus to a worksheet.
- `excel_rename_sheet(old_name, new_name)`: Rename a worksheet.
- `excel_delete_sheet(name)`: Delete a worksheet.
- `excel_set_tab_color(hex_color)`: Colorize worksheet tabs (e.g. Red, Blue, Green).
- `excel_toggle_gridlines(show)`: Show/hide gridlines (ideal for sleek dashboard presentations).
- `excel_toggle_headings(show)`: Show/hide row numbers and column letters.

### 3. Data & Formulas
- `excel_read_range(range_address, sheet_name)`: Read cell or 2D range values.
- `excel_write_range(start_cell, data, sheet_name)`: Write 2D array of values starting at `start_cell`.
- `excel_write_cell(cell, value, sheet_name)`: Write a single value or formula.
- `excel_set_formula(cell, formula, sheet_name)`: Set an Excel formula (`=SUM(...)`, etc.).
- `excel_clear_range(range_address, clear_type, sheet_name)`: Clear contents, formatting, or all.
- `excel_find_replace(find_text, replace_text, sheet_name)`: Search and replace text.
- `excel_find_all(search_term, match_case)`: Find all cell occurrences across sheet.

### 4. Formatting & Layout
- `excel_format_range(...)`: Apply bold, italic, font name, font size, colors (hex), alignment, borders, and number formats.
- `excel_autofit(columns, rows, sheet_name)`: Auto-fit column widths and row heights.
- `excel_freeze_panes(row, col, sheet_name)`: Freeze header rows and/or leading columns.
- `excel_unfreeze_panes()`: Unfreeze panes.
- `excel_add_hyperlink(cell, target, text, tooltip)`: Insert clickable web link or internal sheet link.
- `excel_remove_hyperlink(cell)`: Remove hyperlink.

### 5. Sorting & Filtering
- `excel_sort_range(range_address, key_column, order, has_headers)`: Sort tabular range by any column.
- `excel_apply_filter(range_address, column_index, criteria)`: Apply AutoFilter criteria.
- `excel_clear_filters()`: Clear all active AutoFilters across sheet.

### 6. Conditional Formatting & Validation
- `excel_add_conditional_formatting(...)`: Color scale heatmaps, data bars, or highlight rules.
- `excel_clear_conditional_formatting(range_address)`: Clear conditional formatting rules.
- `excel_add_data_validation(range_address, validation_type, formula1)`: Add dropdown lists or input constraints.
- `excel_clear_data_validation(range_address)`: Remove validation dropdowns.

### 7. Structural Rows, Columns & Merging
- `excel_insert_rows(row_index, count)`: Insert blank rows.
- `excel_delete_rows(row_index, count)`: Delete rows.
- `excel_insert_columns(column, count)`: Insert blank columns.
- `excel_delete_columns(column, count)`: Delete columns.
- `excel_hide_unhide_columns(columns, hide)`: Hide or unhide columns.
- `excel_hide_unhide_rows(rows, hide)`: Hide or unhide rows.
- `excel_set_column_width(columns, width)`: Set column width.
- `excel_set_row_height(rows, height)`: Set row height.
- `excel_merge_cells(range_address)`: Merge cell ranges.
- `excel_unmerge_cells(range_address)`: Unmerge cells.

### 8. Data Analysis, What-If & Cleaning
- `excel_goal_seek(target_cell, target_value, changing_cell)`: Native financial Goal Seek analysis.
- `excel_text_to_columns(range_address, destination_cell, delimiter)`: Split delimited single column into columns.
- `excel_remove_duplicates(range_address, columns, has_headers)`: Deduplicate table rows.
- `excel_get_column_summary(range_address)`: Instant statistical metrics (Sum, Count, Average, Min, Max).

### 9. Page Setup & Printing
- `excel_set_page_setup(orientation, paper_size, fit_to_pages)`: Configure page orientation and 1-page fit.
- `excel_set_print_area(range_address)`: Set explicit print region.
- `excel_clear_print_area()`: Clear print region.

### 10. Images & Range Screenshots
- `excel_add_image(image_path, cell, width, height)`: Insert image or logo onto worksheet.
- `excel_export_range_image(range_address, output_path)`: High-resolution PNG snapshot of a range.

### 11. Comments & Notes
- `excel_add_comment(cell, text, author)`: Attach comment/note to cell.
- `excel_list_comments()`: List all review comments in sheet.
- `excel_delete_comment(cell)`: Remove comment from cell.

### 12. Security & Protection
- `excel_protect_sheet(password)`: Lock sheet against edits.
- `excel_unprotect_sheet(password)`: Unlock worksheet.

### 13. Named Ranges & Formulas Audit
- `excel_create_named_range(name, range_address)`: Define named range for formulas.
- `excel_list_named_ranges()`: List all defined workbook names.
- `excel_delete_named_range(name)`: Delete named range.
- `excel_calculate_now()`: Force workbook recalculation.
- `excel_list_formulas(range_address)`: Audit all formulas, strings, and evaluated values.

### 14. Tables, Charts & Visualizations
- `excel_create_table(range_address, table_name, style_name, has_headers)`: Convert range to Excel Table (ListObject).
- `excel_add_table_total_row(table_name)`: Show table total row.
- `excel_add_chart(chart_type, source_range, title, left, top, width, height)`: Insert chart.
- `excel_export_chart_image(chart_index, output_path)`: Export chart as PNG.
- `excel_add_sparkline(data_range, location_cell, sparkline_type)`: Insert mini in-cell trend lines.
- `excel_create_pivot_table(...)`: Dynamic native Pivot Table.

### 15. Executive Components & Roadmaps
- `excel_create_kpi_card(start_cell, title, value, subtitle, accent_color)`: Render 3-row executive KPI block.
- `excel_create_gantt_block(start_cell, tasks, weeks_count)`: Render project Gantt timeline grid.

### 16. Headless OOXML Engine (openpyxl)
- `excel_create_from_markdown(markdown_table, output_path, sheet_name)`: Convert Markdown table to styled `.xlsx`.
- `excel_read_to_markdown(file_path, sheet_name)`: Read `.xlsx` table into clean Markdown.
- `excel_json_to_sheet(json_data, output_path, sheet_name)`: Export JSON records to styled `.xlsx`.
- `excel_csv_to_excel(csv_path, output_path, delimiter, sheet_name)`: Convert CSV to styled `.xlsx`.
- `excel_to_csv(file_path, output_csv_path, sheet_name, delimiter)`: Export sheet to CSV.
- `excel_merge_workbooks(file_paths, output_path)`: Consolidate multiple workbooks into one.
- `excel_audit_workbook(file_path)`: Deep structural inspection of an `.xlsx` file.
- `excel_batch_replace(file_path, replacements, output_path)`: Batch replace template variables.
- `excel_extract_column(file_path, column, sheet_name)`: Extract raw column values list.

---

## ⚙️ Configuration Setup

Add to your MCP configuration (`~/.gemini/config/mcp_config.json` or Claude Desktop):

```json
{
  "mcpServers": {
    "excel": {
      "command": "C:\\Users\\HP\\Desktop\\tkhrbi9\\mcps\\excel-mcp\\.venv\\Scripts\\python.exe",
      "args": [
        "C:\\Users\\HP\\Desktop\\tkhrbi9\\mcps\\excel-mcp\\server.py"
      ]
    }
  }
}
```

---

## 🧪 Testing

```bash
cd excel-mcp
uv venv
uv pip install -e .

# Basic Demo: KPI Cards, Formulas, Tables, Charts, PDF Export
uv run python test_demo.py

# Advanced Demo: Validation, Sorting, Conditional Formatting, Comments
uv run python test_advanced_utils.py

# Super Tools: Goal Seek, Gantt Roadmaps, Text-to-Columns, Sparklines, Audit
uv run python test_super_tools.py
```
