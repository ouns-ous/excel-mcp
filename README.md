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

## 🛠️ Complete List of Tools (201 Tools)

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

### 18. Print Titles, Margins & Page Break Controls
- `excel_set_repeat_rows_cols(repeat_rows, repeat_cols)`: Set rows/columns to repeat on every printed page (Print Titles).
- `excel_insert_page_break(cell)`: Insert a horizontal and vertical page break at cell.
- `excel_reset_page_breaks()`: Reset all custom page breaks back to default.
- `excel_set_sheet_margins(top, bottom, left, right, header, footer)`: Set precise printing and PDF margins in inches.
- `excel_set_sheet_orientation_and_paper(orientation, paper_size)`: Set portrait/landscape and A4/Letter/Legal paper size.
- `excel_set_sheet_fit_to_pages(fit_to_wide, fit_to_tall)`: Fit sheet to an exact number of pages wide and tall.

### 19. Calculation Modes & Data Transformations
- `excel_set_calculation_mode(mode)`: Switch calculation engine between 'automatic', 'manual', and 'semiautomatic'.
- `excel_get_calculation_mode()`: Read current calculation engine state.
- `excel_transpose_range(source_range, target_start_cell)`: Transpose matrix rows into columns.
- `excel_fill_series(start_range, fill_range, series_type, step, stop)`: Auto-fill linear, growth, or date series.
- `excel_set_view_mode(mode)`: Switch view between 'normal', 'page_break_preview', and 'page_layout'.

### 20. Advanced What-If Sensitivity & Scenarios
- `excel_create_two_variable_data_table(table_range, row_input_cell, col_input_cell)`: Generate a 2-variable sensitivity Data Table matrix.
- `excel_add_scenario(name, changing_cells, values, comment)`: Create a What-If Scenario with varying cell values.
- `excel_show_scenario(name)`: Apply and display a saved What-If Scenario on the worksheet.
- `excel_list_scenarios()`: List all What-If Scenarios saved on the worksheet.

### 21. Interactive Table Slicers & Conversions
- `excel_add_table_slicer(table_name, column_name, slicer_name, left, top, width, height)`: Insert an interactive graphic visual Slicer button box.
- `excel_convert_table_to_range(table_name)`: Convert an official Excel Table back to a standard cell range.

### 22. External Data Connections & Power Query
- `excel_refresh_all_connections()`: Refresh all external data connections, Power Query queries, and PivotTables.
- `excel_get_data_connections()`: List all external data connections (ODBC, OLEDB, Power Query) in the workbook.
- `excel_refresh_connection(connection_name)`: Refresh a specific external data connection by name.

### 23. Chart Customization & Visual Export
- `excel_set_chart_secondary_axis(chart_name, series_index)`: Move a chart series to the secondary vertical (Y) axis.
- `excel_add_chart_trendline(chart_name, series_index, trendline_type)`: Add a trendline (linear, exponential, moving average).
- `excel_export_chart_as_image(chart_name, output_image_path)`: Export a specific chart as high-resolution PNG or JPG.

### 24. Cell Inspection, Auditing & Content Clearing
- `excel_get_formula_errors()`: Scan worksheet and return all cells with errors (#DIV/0!, #N/A, #VALUE!).
- `excel_find_cells_with_comments()`: Find and return all cells containing notes or comments.
- `excel_find_cells_with_data_validation()`: Find all cells with Data Validation rules applied.
- `excel_get_merged_ranges()`: List all merged cell areas on the worksheet.
- `excel_clear_cell_contents(range_address)`: Clear values and formulas while keeping formatting intact.
- `excel_clear_cell_formats(range_address)`: Clear styling while keeping values and formulas intact.
- `excel_export_sheet_to_csv(output_path)`: Direct COM export of active sheet to CSV.
- `excel_get_precedents(cell_address)`: Return addresses of cells referenced by this formula.
- `excel_get_dependents(cell_address)`: Return addresses of cells dependent on this cell.

### 25. Filter Controls, Alignment & Form Security
- `excel_clear_sheet_filters()`: Clear active filter selections without removing filter dropdown arrows.
- `excel_apply_autofilter_column(range_address, field_index, criteria)`: Apply filter criteria on a specific column.
- `excel_remove_autofilter()`: Remove AutoFilter mode and dropdowns completely.
- `excel_wrap_text(range_address, wrap)`: Toggle text wrapping on a range.
- `excel_shrink_to_fit(range_address, shrink)`: Toggle Shrink-to-Fit on cells.
- `excel_set_cell_orientation(range_address, degrees)`: Set text rotation orientation angle (-90 to +90).
- `excel_flash_fill_column(range_address)`: Trigger Excel Flash Fill down a column.
- `excel_set_chart_title(chart_name, title, font_size, bold)`: Directly format chart title.
- `excel_set_chart_legend(chart_name, position, show_legend)`: Place legend (top, bottom, left, right, corner).
- `excel_set_chart_axes_titles(chart_name, x_title, y_title)`: Set horizontal and vertical axis title labels.
- `excel_set_chart_style(chart_name, style_id)`: Apply built-in modern visual chart themes.
- `excel_lock_cells(range_address, locked)`: Lock or unlock cells for form templates.
- `excel_hide_formulas(range_address, hidden)`: Hide formulas from formula bar in protected sheets.
- `excel_is_sheet_protected()`: Query worksheet protection status.
- `excel_set_row_visibility(row_index, hidden)`: Hide or unhide specific rows.
- `excel_set_column_visibility(column, hidden)`: Hide or unhide specific columns.

### 26. Freeze Panes, AutoFit & Range Copy
- `excel_freeze_custom_panes(row, column)`: Freeze panes at any custom row and column intersection.
- `excel_freeze_top_row()`: Freeze top header row during vertical scrolling.
- `excel_freeze_first_column()`: Freeze first column during horizontal scrolling.
- `excel_copy_range(source_range, destination_cell)`: Copy a range with formulas and styles.
- `excel_autofit_rows(range_address)`: Auto-fit row heights to cell content.
- `excel_autofit_all()`: Auto-fit all rows and columns across the entire used range.

### 27. Typography, Window Toggles & Threaded Comments
- `excel_set_strikethrough(range_address, strikethrough)`: Apply or remove strikethrough text formatting.
- `excel_set_subscript(range_address, subscript)`: Apply or remove subscript formatting.
- `excel_set_superscript(range_address, superscript)`: Apply or remove superscript formatting.
- `excel_set_underline_style(range_address, underline_type)`: Single, double, or accounting underlines.
- `excel_toggle_formula_bar(show)`: Show or hide Excel formula bar.
- `excel_toggle_status_bar(show)`: Show or hide Excel bottom status bar.
- `excel_toggle_scroll_bars(horizontal, vertical)`: Show or hide horizontal and vertical scrollbars.
- `excel_toggle_sheet_tabs(show)`: Show or hide worksheet tabs bar.
- `excel_add_threaded_comment(cell, text)`: Modern Excel 365 threaded comments.
- `excel_delete_threaded_comment(cell)`: Delete threaded comments.
- `excel_list_threaded_comments()`: List threaded comments with author and replies count.

### 28. Workbook Protection & Print Quality
- `excel_protect_workbook_structure(password, protect_structure, protect_windows)`: Lock sheet add/delete/rename structure.
- `excel_unprotect_workbook_structure(password)`: Unlock workbook structure.
- `excel_is_workbook_protected()`: Query workbook structure protection status.
- `excel_set_print_quality(dpi)`: Set print resolution quality in DPI (300, 600).
- `excel_set_black_and_white_print(bw)`: Enable clean monochrome PDF export or printing.

### 29. Headless OOXML Engine (openpyxl)
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
- `excel_compare_two_workbooks(file_path_1, file_path_2)`: Compare two `.xlsx` files and return structural and cell differences.
- `excel_extract_hyperlinks(file_path)`: Extract all embedded hyperlinks and URLs headlessly.
- `excel_unhide_all_sheets_headless(file_path, output_path)`: Unhide all hidden and very hidden worksheets headlessly.
- `excel_copy_sheet_headless(file_path, source_sheet, new_sheet_name)`: Clone worksheet headlessly.
- `excel_protect_sheet_headless(file_path, password, sheet_name)`: Password protect sheet headlessly.
- `excel_delete_sheet_headless(file_path, sheet_name)`: Remove worksheet headlessly.
- `excel_set_sheet_color_headless(file_path, sheet_name, hex_color)`: Colorize worksheet tab headlessly.
- `excel_get_defined_names_headless(file_path)`: Extract all defined names/ranges headlessly.
- `excel_insert_rows_headless(file_path, row_idx, amount)`: Insert blank rows headlessly.
- `excel_delete_rows_headless(file_path, row_idx, amount)`: Delete rows headlessly.
- `excel_insert_cols_headless(file_path, col_idx, amount)`: Insert blank columns headlessly.
- `excel_delete_cols_headless(file_path, col_idx, amount)`: Delete columns headlessly.

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

# Ultra Tools: Workbook Comparison, Hyperlinks, Sheet Unhiding
uv run python test_ultra_tools.py

# Mega Tools: Sheet Cloning, Tab Color, Sheet Protection, Deletion
uv run python test_mega_tools.py

# Giga Tools: Headless Defined Names, Row/Column Insertions & Deletions
uv run python test_giga_tools.py
```



