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

## 🛠️ Complete List of Tools (119 Tools)

### 1. Workbook & Window Management
- `excel_status()`: Get current running state, open workbooks count, active sheet, active cell, and used range.
- `excel_new_workbook(sheet_name)`: Create a new visible workbook.
- `excel_open(file_path)`: Open an existing workbook (`.xlsx`, `.xls`, `.xlsm`, `.csv`).
- `excel_save(file_path)`: Save the active workbook.
- `excel_export_pdf(output_path, sheet_only)`: Export workbook or active sheet directly to PDF.
- `excel_export_html(output_path, sheet_only)`: Export workbook or sheet as standalone interactive HTML web file.
- `excel_close(save_changes)`: Close the active workbook.
- `excel_focus_window()`: Bring Microsoft Excel to foreground and maximize.

### 2. Sheet Operations, Organization & Tab Styling
- `excel_list_sheets()`: List all worksheets with indices and visibility status.
- `excel_add_sheet(name, after)`: Add a new worksheet.
- `excel_duplicate_sheet(sheet_name, new_name)`: Clone an existing worksheet with all formulas and styles.
- `excel_move_sheet(sheet_name, before, after)`: Reorder worksheet tabs.
- `excel_activate_sheet(name)`: Switch active focus to a worksheet.
- `excel_rename_sheet(old_name, new_name)`: Rename a worksheet.
- `excel_delete_sheet(name)`: Delete a worksheet.
- `excel_hide_unhide_sheet(sheet_name, hide, very_hidden)`: Hide or unhide sheets (supports xlSheetVeryHidden).
- `excel_set_tab_color(hex_color)`: Colorize worksheet tabs (e.g. Red, Blue, Green).
- `excel_toggle_gridlines(show)`: Show/hide gridlines (ideal for sleek dashboard presentations).
- `excel_toggle_headings(show)`: Show/hide row numbers and column letters.
- `excel_set_zoom(zoom_percentage)`: Set window zoom level (10% to 400%).
- `excel_split_panes(row, col)`: Split or unsplit window view panes.

### 3. Data & Formulas
- `excel_read_range(range_address, sheet_name)`: Read cell or 2D range values.
- `excel_write_range(start_cell, data, sheet_name)`: Write 2D array of values starting at `start_cell`.
- `excel_write_cell(cell, value, sheet_name)`: Write a single value or formula.
- `excel_set_formula(cell, formula, sheet_name)`: Set an Excel formula (`=SUM(...)`, etc.).
- `excel_convert_formulas_to_values(range_address)`: Freeze formula results into static values.
- `excel_evaluate_expression(expression)`: Evaluate arbitrary Excel expression using calculation engine.
- `excel_clear_range(range_address, clear_type, sheet_name)`: Clear contents, formatting, or all.
- `excel_clear_all_formatting()`: Reset all formatting across used range to raw defaults.
- `excel_find_replace(find_text, replace_text, sheet_name)`: Search and replace text.
- `excel_find_all(search_term, match_case)`: Find all cell occurrences across sheet.

### 4. Formatting, Styles & Number Presets
- `excel_format_range(...)`: Apply bold, italic, font name, font size, colors (hex), alignment, borders, and number formats.
- `excel_apply_cell_style(range_address, style_name)`: Apply built-in Excel style ('Good', 'Bad', 'Neutral', 'Title', etc.).
- `excel_list_cell_styles()`: List all workbook cell styles.
- `excel_set_number_format_preset(range_address, preset)`: Preset formats (currency USD/EUR/MAD, percent, accounting, dates).
- `excel_autofit(columns, rows, sheet_name)`: Auto-fit column widths and row heights.
- `excel_freeze_panes(row, col, sheet_name)`: Freeze header rows and/or leading columns.
- `excel_unfreeze_panes()`: Unfreeze panes.
- `excel_add_hyperlink(cell, target, text, tooltip)`: Insert clickable web link or internal sheet link.
- `excel_remove_hyperlink(cell)`: Remove hyperlink.

### 5. Outlining, Grouping & Hierarchies
- `excel_group_rows(start_row, end_row)`: Group detail rows into collapsible hierarchies (+/-).
- `excel_ungroup_rows(start_row, end_row)`: Remove row grouping.
- `excel_group_columns(start_col, end_col)`: Group detail columns into collapsible views.
- `excel_ungroup_columns(start_col, end_col)`: Remove column grouping.
- `excel_show_outline_levels(row_levels, col_levels)`: Set outline depth (1=summarized, 2+=detailed).

### 6. Sorting & Filtering
- `excel_sort_range(range_address, key_column, order, has_headers)`: Sort tabular range by any column.
- `excel_apply_filter(range_address, column_index, criteria)`: Apply AutoFilter criteria.
- `excel_clear_filters()`: Clear all active AutoFilters across sheet.

### 7. Conditional Formatting & Validation
- `excel_add_conditional_formatting(...)`: Color scale heatmaps, data bars, or highlight rules.
- `excel_add_top_bottom_rule(range_address, rule_type, rank)`: Highlight Top N or Bottom N items/percentages.
- `excel_add_above_below_average_rule(range_address, above)`: Highlight cells above or below range average.
- `excel_add_icon_set_rule(range_address, icon_set_type)`: In-cell KPI icons (traffic lights, arrows, flags).
- `excel_highlight_duplicate_values(range_address)`: Highlight duplicate values dynamically.
- `excel_clear_conditional_formatting(range_address)`: Clear conditional formatting rules.
- `excel_add_data_validation(range_address, validation_type, formula1)`: Add dropdown lists or input constraints.
- `excel_clear_data_validation(range_address)`: Remove validation dropdowns.

### 8. Structural Rows, Columns & Merging
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

### 9. Formula Auditing & Tracer Arrows
- `excel_trace_precedents(cell)`: Draw arrows pointing to predecessor data cells.
- `excel_trace_dependents(cell)`: Draw arrows pointing to dependent formulas.
- `excel_remove_tracer_arrows()`: Clear all visual audit tracer arrows.
- `excel_list_formulas(range_address)`: Inspect formulas and current calculated values.
- `excel_calculate_now()`: Force workbook recalculation.

### 10. Data Analysis, What-If & Cleaning
- `excel_goal_seek(target_cell, target_value, changing_cell)`: Native financial Goal Seek analysis.
- `excel_text_to_columns(range_address, destination_cell, delimiter)`: Split delimited single column into columns.
- `excel_remove_duplicates(range_address, columns, has_headers)`: Deduplicate table rows.
- `excel_get_column_summary(range_address)`: Instant statistical metrics (Sum, Count, Average, Min, Max).
- `excel_get_special_cells_summary(cell_type)`: Inspect and count special cells ('blanks', 'constants', 'formulas', 'comments').

### 11. Page Setup, Printing & Headers/Footers
- `excel_set_page_setup(orientation, paper_size, fit_to_pages)`: Configure page orientation and 1-page fit.
- `excel_set_header_footer(...)`: Configure custom printed headers and footers (&P page, &D date, etc.).
- `excel_set_print_area(range_address)`: Set explicit print region.
- `excel_clear_print_area()`: Clear print region.

### 12. Images, Shapes, Badges & Watermarks
- `excel_add_image(image_path, cell, width, height)`: Insert image or logo onto worksheet.
- `excel_add_shape(shape_type, cell, width, height, text)`: Insert modern native shapes (rounded rectangle, arrow, oval) with text.
- `excel_list_shapes()`: List all shapes and graphic objects on the sheet.
- `excel_delete_shape(shape_name)`: Delete a graphic shape or object.
- `excel_add_watermark(text, font_size, color)`: Add semi-transparent diagonal watermark ('CONFIDENTIAL', 'DRAFT').
- `excel_export_range_image(range_address, output_path)`: High-resolution PNG snapshot of a range.

### 13. Comments & Notes
- `excel_add_comment(cell, text, author)`: Attach comment/note to cell.
- `excel_list_comments()`: List all review comments in sheet.
- `excel_delete_comment(cell)`: Remove comment from cell.

### 14. Security, Cell Permissions & Protection
- `excel_protect_sheet(password)`: Lock sheet against edits.
- `excel_unprotect_sheet(password)`: Unlock worksheet.
- `excel_protect_workbook(password, lock_structure)`: Prevent adding, deleting, or renaming sheets.
- `excel_unprotect_workbook(password)`: Unlock workbook structure.
- `excel_lock_unlock_cells(range_address, locked)`: Define cell-level edit permissions on protected sheets.

### 15. Named Ranges
- `excel_create_named_range(name, range_address)`: Define named range for formulas.
- `excel_list_named_ranges()`: List all defined workbook names.
- `excel_delete_named_range(name)`: Delete named range.

### 16. Tables, Charts & Visualizations
- `excel_create_table(range_address, table_name, style_name, has_headers)`: Convert range to Excel Table (ListObject).
- `excel_add_table_total_row(table_name)`: Show table total row.
- `excel_add_chart(chart_type, source_range, title, left, top, width, height)`: Insert chart.
- `excel_export_chart_image(chart_index, output_path)`: Export chart as PNG.
- `excel_add_sparkline(data_range, location_cell, sparkline_type)`: Insert mini in-cell trend lines.
- `excel_create_pivot_table(...)`: Dynamic native Pivot Table.

### 17. Executive Components, Controls & Roadmaps
- `excel_create_kpi_card(start_cell, title, value, subtitle, accent_color)`: Render 3-row executive KPI block.
- `excel_add_metric_tile(cell, label, value, delta, status)`: Compact 2-row metric tile.
- `excel_add_callout_box(cell, text, title, box_type)`: Executive alert banner shape.
- `excel_add_checkbox(cell, caption, linked_cell)`: Form Control interactive checkbox.
- `excel_create_gantt_block(start_cell, tasks, weeks_count)`: Render project Gantt timeline grid.

### 18. Headless OOXML Engine (openpyxl)
- `excel_create_from_markdown(markdown_table, output_path, sheet_name)`: Convert Markdown table to styled `.xlsx`.
- `excel_read_to_markdown(file_path, sheet_name)`: Read `.xlsx` table into clean Markdown.
- `excel_json_to_sheet(json_data, output_path, sheet_name)`: Export JSON records to styled `.xlsx`.
- `excel_csv_to_excel(csv_path, output_path, delimiter, sheet_name)`: Convert CSV to styled `.xlsx`.
- `excel_to_csv(file_path, output_csv_path, sheet_name, delimiter)`: Export sheet to CSV.
- `excel_merge_workbooks(file_paths, output_path)`: Consolidate multiple workbooks into one.
- `excel_audit_workbook(file_path)`: Deep structural inspection of an `.xlsx` file.
- `excel_batch_replace(file_path, replacements, output_path)`: Batch replace template variables.
- `excel_extract_column(file_path, column, sheet_name)`: Extract raw column values list.
- `excel_extract_metadata(file_path)`: Extract OOXML document properties.
- `excel_set_metadata(file_path, title, author, subject)`: Set/scrub document metadata.

---

## ⚙️ Configuration Setup

Add to your MCP configuration (`~/.gemini/config/mcp_config.json` or Claude Desktop):

```json
{
  "mcpServers": {
    "excel": {
      "command": "C:\\Users\\HP\\Desktop\\excel-mcp\\.venv\\Scripts\\python.exe",
      "args": [
        "C:\\Users\\HP\\Desktop\\excel-mcp\\server.py"
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

# Power Tools: Grouping, Outlines, Top/Bottom Rules, Icon Sets, Tracer Arrows, Shapes
uv run python test_power_tools.py
```
