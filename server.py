"""Microsoft Excel Live Automation & OOXML MCP Server.
Empowers AI coding assistants (Antigravity, Claude Desktop, Cursor) with full,
real-time control over Microsoft Excel on Windows, plus fast headless .xlsx processing.
"""

from typing import Any, Dict, List, Optional, Union
from mcp.server.fastmcp import FastMCP
from excel_controller import controller
from excel_engine import ExcelEngine

# Initialize FastMCP server
mcp = FastMCP("Excel-Controller")


# ── LIVE WINDOWS EXCEL DESKTOP AUTOMATION ────────────────────────────────────────

@mcp.tool()
def excel_status() -> Dict[str, Any]:
    """Get the current running status of Microsoft Excel and the active workbook.
    Returns whether Excel is running, open workbooks count, active workbook name,
    sheets list, active cell, and used range.
    """
    return controller.get_info()


@mcp.tool()
def excel_new_workbook(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Create a new blank workbook in Microsoft Excel desktop and bring it to front.
    Args:
        sheet_name: Optional name for the initial active worksheet.
    """
    return controller.new_workbook(sheet_name=sheet_name)


@mcp.tool()
def excel_open(file_path: str) -> Dict[str, Any]:
    """Open an existing Excel workbook (.xlsx, .xls, .xlsm, .csv) in desktop and bring to front.
    Args:
        file_path: Absolute or relative path to the workbook file.
    """
    return controller.open_workbook(file_path)


@mcp.tool()
def excel_save(file_path: Optional[str] = None) -> Dict[str, Any]:
    """Save the active workbook.
    Args:
        file_path: Optional new destination path (.xlsx). If omitted, saves in place.
    """
    return controller.save_workbook(file_path)


@mcp.tool()
def excel_export_pdf(output_path: Optional[str] = None, sheet_only: bool = False) -> Dict[str, Any]:
    """Export the active workbook or current sheet directly to a high-quality PDF using Excel's native engine.
    Args:
        output_path: Optional destination path (.pdf). If omitted, saves alongside active file.
        sheet_only: If True, exports only the active sheet; if False, exports entire workbook.
    """
    return controller.export_pdf(output_path=output_path, sheet_only=sheet_only)


@mcp.tool()
def excel_close(save_changes: bool = False) -> Dict[str, Any]:
    """Close the active workbook.
    Args:
        save_changes: Whether to save changes before closing (default False).
    """
    return controller.close_workbook(save_changes=save_changes)


@mcp.tool()
def excel_focus_window() -> Dict[str, Any]:
    """Bring Microsoft Excel window to foreground and maximize."""
    controller.focus_window()
    return {"status": "success", "message": "Excel focused."}


# ── WORKSHEET MANAGEMENT ────────────────────────────────────────────────────────

@mcp.tool()
def excel_list_sheets() -> List[Dict[str, Any]]:
    """List all worksheets in the active workbook with their names, indices, and visibility."""
    return controller.list_sheets()


@mcp.tool()
def excel_add_sheet(name: str, after: Optional[str] = None) -> Dict[str, Any]:
    """Add a new worksheet to the active workbook.
    Args:
        name: Name for the new worksheet (max 31 chars).
        after: Optional name of the sheet after which to insert.
    """
    return controller.add_sheet(name=name, after=after)


@mcp.tool()
def excel_activate_sheet(name: str) -> Dict[str, Any]:
    """Switch active focus to a specific worksheet by name.
    Args:
        name: Name of the sheet to activate.
    """
    return controller.activate_sheet(name=name)


@mcp.tool()
def excel_rename_sheet(old_name: str, new_name: str) -> Dict[str, Any]:
    """Rename an existing worksheet in the active workbook.
    Args:
        old_name: Current sheet name.
        new_name: New sheet name (max 31 chars).
    """
    return controller.rename_sheet(old_name=old_name, new_name=new_name)


@mcp.tool()
def excel_delete_sheet(name: str) -> Dict[str, Any]:
    """Delete a worksheet by name.
    Args:
        name: Name of the sheet to delete.
    """
    return controller.delete_sheet(name=name)


# ── CELL DATA & FORMULAS ────────────────────────────────────────────────────────

@mcp.tool()
def excel_read_range(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Read values from a specific cell or range (e.g. 'A1:D10' or 'B2').
    Args:
        range_address: Standard Excel range notation (e.g. 'A1:C5').
        sheet_name: Optional specific worksheet name. If omitted, uses active sheet.
    """
    return controller.read_range(range_address=range_address, sheet_name=sheet_name)


@mcp.tool()
def excel_write_range(start_cell: str, data: List[List[Any]], sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Write a 2D array of values starting from start_cell into the worksheet.
    Args:
        start_cell: Top-left cell coordinate (e.g. 'A1', 'B3').
        data: 2D list of values to write: [[row1_col1, row1_col2], [row2_col1, row2_col2]].
        sheet_name: Optional specific worksheet name.
    """
    return controller.write_range(start_cell=start_cell, data=data, sheet_name=sheet_name)


@mcp.tool()
def excel_write_cell(cell: str, value: Any, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Write a single value or formula into a specific cell.
    Args:
        cell: Cell coordinate (e.g. 'A1', 'C4').
        value: Value to write (string, number, or formula starting with '=').
        sheet_name: Optional specific worksheet name.
    """
    return controller.write_cell(cell=cell, value=value, sheet_name=sheet_name)


@mcp.tool()
def excel_set_formula(cell: str, formula: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set an Excel formula in a cell (e.g., '=SUM(B2:B10)', '=AVERAGE(C2:C20)', '=VLOOKUP(...)').
    Args:
        cell: Cell coordinate (e.g. 'B11').
        formula: Excel formula string (with or without leading '=').
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_formula(cell=cell, formula=formula, sheet_name=sheet_name)


@mcp.tool()
def excel_clear_range(range_address: str, clear_type: str = "all", sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Clear a range of cells.
    Args:
        range_address: Range to clear (e.g. 'A1:D20').
        clear_type: 'all' (contents + formats), 'contents' (data only), or 'formats' (styling only).
        sheet_name: Optional specific worksheet name.
    """
    return controller.clear_range(range_address=range_address, clear_type=clear_type, sheet_name=sheet_name)


@mcp.tool()
def excel_find_replace(find_text: str, replace_text: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Find and replace all occurrences of a text string across the worksheet.
    Args:
        find_text: Text to search for.
        replace_text: Replacement text.
        sheet_name: Optional specific worksheet name.
    """
    return controller.find_replace(find_text=find_text, replace_text=replace_text, sheet_name=sheet_name)


# ── STYLING & FORMATTING ────────────────────────────────────────────────────────

@mcp.tool()
def excel_format_range(
    range_address: str,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    horizontal_align: Optional[str] = None,
    vertical_align: Optional[str] = None,
    wrap_text: Optional[bool] = None,
    number_format: Optional[str] = None,
    border_style: Optional[str] = None,
    border_color: Optional[str] = None,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Apply rich styling and typography to a range of cells.
    Args:
        range_address: Range to format (e.g. 'A1:F1', 'B2:B10').
        bold: Set bold text.
        italic: Set italic text.
        font_name: Font family name (e.g. 'Segoe UI', 'Calibri', 'Arial').
        font_size: Font size in points (e.g. 11, 14, 18).
        font_color: Hex color string (e.g. '#FFFFFF', '#1F4E79').
        bg_color: Fill background hex color string (e.g. '#1F4E79', '#F1F5F9').
        horizontal_align: 'left', 'center', 'right', or 'justify'.
        vertical_align: 'top', 'center', or 'bottom'.
        wrap_text: Whether to enable text wrapping.
        number_format: Excel number format string (e.g. '$#,##0.00', '0.0%', 'yyyy-mm-dd', '#,##0').
        border_style: 'thin', 'medium', 'thick', 'double', or 'none'.
        border_color: Border hex color string (e.g. '#CBD5E1').
        sheet_name: Optional specific worksheet name.
    """
    return controller.format_range(
        range_address=range_address,
        bold=bold,
        italic=italic,
        font_name=font_name,
        font_size=font_size,
        font_color=font_color,
        bg_color=bg_color,
        horizontal_align=horizontal_align,
        vertical_align=vertical_align,
        wrap_text=wrap_text,
        number_format=number_format,
        border_style=border_style,
        border_color=border_color,
        sheet_name=sheet_name,
    )


@mcp.tool()
def excel_autofit(columns: Optional[str] = None, rows: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Automatically fit column widths and/or row heights to fit content perfectly.
    Args:
        columns: Specific column range (e.g. 'A:E') or None to autofit all columns.
        rows: Specific row range (e.g. '1:20') or None to autofit all rows.
        sheet_name: Optional specific worksheet name.
    """
    return controller.autofit(columns=columns, rows=rows, sheet_name=sheet_name)


@mcp.tool()
def excel_freeze_panes(row: int = 2, col: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Freeze panes so top headers or leading columns remain visible when scrolling.
    Args:
        row: Row index below which to freeze (e.g. 2 freezes row 1).
        col: Column index right of which to freeze (e.g. 1 freezes no columns, 2 freezes column A).
        sheet_name: Optional specific worksheet name.
    """
    return controller.freeze_panes(row=row, col=col, sheet_name=sheet_name)


@mcp.tool()
def excel_unfreeze_panes() -> Dict[str, Any]:
    """Unfreeze all frozen panes in the active worksheet."""
    return controller.unfreeze_panes()


# ── TABLES & CHARTS ─────────────────────────────────────────────────────────────

@mcp.tool()
def excel_create_table(
    range_address: str,
    table_name: Optional[str] = None,
    style_name: str = "TableStyleMedium9",
    has_headers: bool = True,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Convert a range of cells into an official Excel Table (ListObject) with banded rows and filtering.
    Args:
        range_address: Range of data (e.g. 'A1:E10').
        table_name: Unique name for the table (e.g. 'SalesData').
        style_name: Table style name (e.g. 'TableStyleMedium9', 'TableStyleLight1', 'TableStyleDark1').
        has_headers: Whether the first row contains column headers (default True).
        sheet_name: Optional specific worksheet name.
    """
    return controller.create_table(
        range_address=range_address,
        table_name=table_name,
        style_name=style_name,
        has_headers=has_headers,
        sheet_name=sheet_name,
    )


@mcp.tool()
def excel_add_table_total_row(table_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Enable and display the Total summary row at the bottom of an Excel Table.
    Args:
        table_name: Name of the table.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_table_total_row(table_name=table_name, sheet_name=sheet_name)


@mcp.tool()
def excel_add_chart(
    chart_type: str,
    source_range: str,
    title: Optional[str] = None,
    left: float = 350,
    top: float = 40,
    width: float = 480,
    height: float = 300,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Insert a native Excel chart connected directly to sheet data.
    Args:
        chart_type: One of 'column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter'.
        source_range: Range containing labels and series data (e.g. 'A1:C6').
        title: Optional chart title displayed at the top.
        left: Horizontal offset position in points.
        top: Vertical offset position in points.
        width: Chart width in points.
        height: Chart height in points.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_chart(
        chart_type=chart_type,
        source_range=source_range,
        title=title,
        left=left,
        top=top,
        width=width,
        height=height,
        sheet_name=sheet_name,
    )


@mcp.tool()
def excel_export_chart_image(chart_index: int, output_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Export an embedded chart to a high-resolution PNG image file.
    Args:
        chart_index: 1-based index of the chart on the sheet.
        output_path: Destination path (.png).
        sheet_name: Optional specific worksheet name.
    """
    return controller.export_chart_image(chart_index=chart_index, output_path=output_path, sheet_name=sheet_name)


# ── EXECUTIVE DASHBOARD BLOCKS ──────────────────────────────────────────────────

@mcp.tool()
def excel_create_kpi_card(
    start_cell: str,
    title: str,
    value: str,
    subtitle: Optional[str] = None,
    accent_color: str = "#2563EB",
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a modern executive KPI metric block with custom styling and indicators.
    Args:
        start_cell: Top-left cell coordinate (e.g. 'B2').
        title: Metric name (e.g. 'Total Revenue', 'Active Users').
        value: Primary metric number or text (e.g. '$1.24M', '98.5%').
        subtitle: Optional secondary trend or badge (e.g. '+14.2% YoY', 'Target Met').
        accent_color: Hex color for primary metric value (default '#2563EB').
        sheet_name: Optional specific worksheet name.
    """
    return controller.create_kpi_card(
        start_cell=start_cell,
        title=title,
        value=value,
        subtitle=subtitle,
        accent_color=accent_color,
        sheet_name=sheet_name,
    )


# ── HEADLESS OOXML PROCESSING (OPENPYXL) ────────────────────────────────────────

@mcp.tool()
def excel_create_from_markdown(markdown_table: str, output_path: str, sheet_name: str = "Data") -> Dict[str, Any]:
    """Convert a Markdown formatted table into a styled Excel (.xlsx) file without needing desktop Excel.
    Args:
        markdown_table: Markdown pipe-delimited table string.
        output_path: Destination file path (.xlsx).
        sheet_name: Name of the worksheet.
    """
    path = ExcelEngine.markdown_to_excel(markdown_table=markdown_table, output_path=output_path, sheet_name=sheet_name)
    return {"status": "success", "file": path, "message": f"Successfully created .xlsx at {path}"}


@mcp.tool()
def excel_read_to_markdown(file_path: str, sheet_name: Optional[str] = None) -> str:
    """Read a table from an Excel (.xlsx) file and return it as a clean Markdown table.
    Args:
        file_path: Absolute or relative path to the .xlsx file.
        sheet_name: Optional sheet name.
    """
    return ExcelEngine.excel_to_markdown(file_path=file_path, sheet_name=sheet_name)


@mcp.tool()
def excel_json_to_sheet(json_data: List[Dict[str, Any]], output_path: str, sheet_name: str = "Data") -> Dict[str, Any]:
    """Convert a list of JSON dictionaries into a formatted Excel (.xlsx) file headlessly.
    Args:
        json_data: List of key-value dictionaries.
        output_path: Destination file path (.xlsx).
        sheet_name: Name of the worksheet.
    """
    path = ExcelEngine.json_to_excel(json_data=json_data, output_path=output_path, sheet_name=sheet_name)
    return {"status": "success", "file": path, "message": f"Successfully exported JSON to {path}"}


# ── SORTING & FILTERING ─────────────────────────────────────────────────────────

@mcp.tool()
def excel_sort_range(
    range_address: str,
    key_column: int = 1,
    order: str = "ascending",
    has_headers: bool = True,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Sort data in a range by a specified key column index (1-based).
    Args:
        range_address: Range of data to sort (e.g. 'A2:F50' or 'A1:F50').
        key_column: 1-based column index to sort by (default 1).
        order: 'ascending' or 'descending'.
        has_headers: Whether the first row of the range is a header row.
        sheet_name: Optional specific worksheet name.
    """
    return controller.sort_range(
        range_address=range_address,
        key_column=key_column,
        order=order,
        has_headers=has_headers,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_apply_filter(
    range_address: str,
    column_index: int = 1,
    criteria: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Apply an AutoFilter to a range on a specific column index.
    Args:
        range_address: Range containing headers and data (e.g. 'A1:E50').
        column_index: 1-based column index within the range to filter.
        criteria: Filter criteria string (e.g. 'Pending', '>1000'). If None, activates filter dropdown without filtering.
        sheet_name: Optional specific worksheet name.
    """
    return controller.apply_filter(
        range_address=range_address,
        column_index=column_index,
        criteria=criteria,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_clear_filters(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Clear all active AutoFilters across the active worksheet."""
    return controller.clear_filters(sheet_name=sheet_name)


# ── CONDITIONAL FORMATTING ──────────────────────────────────────────────────────

@mcp.tool()
def excel_add_conditional_formatting(
    range_address: str,
    rule_type: str = "color_scale",
    operator: str = "greater",
    value: Optional[str] = None,
    bg_color: Optional[str] = None,
    font_color: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Add conditional formatting rules to highlight data dynamically.
    Args:
        range_address: Range of cells to format (e.g. 'D2:D50').
        rule_type: 'color_scale' (3-color heatmap), 'data_bar' (in-cell bar), or 'cell_value' (highlight by value).
        operator: For 'cell_value': 'greater', 'less', 'equal', 'between', 'not_equal'.
        value: Threshold value for the rule (e.g. '1000', '0.05').
        bg_color: Hex background highlight color (e.g. '#DC2626' red, '#16A34A' green).
        font_color: Hex font color.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_conditional_formatting(
        range_address=range_address,
        rule_type=rule_type,
        operator=operator,
        value=value,
        bg_color=bg_color,
        font_color=font_color,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_clear_conditional_formatting(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Remove all conditional formatting rules from a range of cells.
    Args:
        range_address: Range to clear (e.g. 'A1:Z100').
        sheet_name: Optional specific worksheet name.
    """
    return controller.clear_conditional_formatting(range_address=range_address, sheet_name=sheet_name)


# ── DATA VALIDATION ─────────────────────────────────────────────────────────────

@mcp.tool()
def excel_add_data_validation(
    range_address: str,
    validation_type: str = "list",
    formula1: str = "Yes,No,Pending",
    error_title: Optional[str] = None,
    error_message: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Add an interactive in-cell dropdown list or data validation rule to cells.
    Args:
        range_address: Range of cells (e.g. 'E2:E50').
        validation_type: 'list' (dropdown choices).
        formula1: Comma-separated list of items (e.g. 'Active,Inactive,On Leave') or formula range.
        error_title: Optional title for alert modal if invalid entry is entered.
        error_message: Optional custom validation message.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_data_validation(
        range_address=range_address,
        validation_type=validation_type,
        formula1=formula1,
        error_title=error_title,
        error_message=error_message,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_clear_data_validation(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Remove data validation dropdowns and restrictions from a range of cells.
    Args:
        range_address: Range to clear (e.g. 'E2:E50').
        sheet_name: Optional specific worksheet name.
    """
    return controller.clear_data_validation(range_address=range_address, sheet_name=sheet_name)


# ── STRUCTURAL ROWS & COLUMNS ───────────────────────────────────────────────────

@mcp.tool()
def excel_insert_rows(row_index: int, count: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Insert one or more blank rows before row_index.
    Args:
        row_index: 1-based row index where rows will be inserted.
        count: Number of rows to insert (default 1).
        sheet_name: Optional specific worksheet name.
    """
    return controller.insert_rows(row_index=row_index, count=count, sheet_name=sheet_name)


@mcp.tool()
def excel_delete_rows(row_index: int, count: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Delete one or more rows starting at row_index.
    Args:
        row_index: 1-based row index to delete.
        count: Number of rows to delete.
        sheet_name: Optional specific worksheet name.
    """
    return controller.delete_rows(row_index=row_index, count=count, sheet_name=sheet_name)


@mcp.tool()
def excel_insert_columns(column: Union[str, int], count: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Insert one or more blank columns before the specified column letter or index.
    Args:
        column: Column letter (e.g. 'C') or 1-based index (e.g. 3).
        count: Number of columns to insert.
        sheet_name: Optional specific worksheet name.
    """
    return controller.insert_columns(column=column, count=count, sheet_name=sheet_name)


@mcp.tool()
def excel_delete_columns(column: Union[str, int], count: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Delete one or more columns starting at the specified column letter or index.
    Args:
        column: Column letter (e.g. 'C') or 1-based index.
        count: Number of columns to delete.
        sheet_name: Optional specific worksheet name.
    """
    return controller.delete_columns(column=column, count=count, sheet_name=sheet_name)


@mcp.tool()
def excel_hide_unhide_columns(columns: str, hide: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Hide or unhide specific columns.
    Args:
        columns: Column range (e.g. 'B:D' or 'F:F').
        hide: True to hide, False to unhide.
        sheet_name: Optional specific worksheet name.
    """
    return controller.hide_unhide_columns(columns=columns, hide=hide, sheet_name=sheet_name)


@mcp.tool()
def excel_hide_unhide_rows(rows: str, hide: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Hide or unhide specific rows.
    Args:
        rows: Row range (e.g. '5:10').
        hide: True to hide, False to unhide.
        sheet_name: Optional specific worksheet name.
    """
    return controller.hide_unhide_rows(rows=rows, hide=hide, sheet_name=sheet_name)


@mcp.tool()
def excel_set_column_width(columns: str, width: float, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set explicit width for columns.
    Args:
        columns: Column range (e.g. 'B:D' or 'A:A').
        width: Column width in characters (e.g. 15, 25.5).
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_column_width(columns=columns, width=width, sheet_name=sheet_name)


@mcp.tool()
def excel_set_row_height(rows: str, height: float, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set explicit height for rows.
    Args:
        rows: Row range (e.g. '1:1' or '2:10').
        height: Row height in points (e.g. 24, 30).
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_row_height(rows=rows, height=height, sheet_name=sheet_name)


@mcp.tool()
def excel_merge_cells(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Merge a contiguous range of cells into a single cell.
    Args:
        range_address: Range to merge (e.g. 'A1:E1').
        sheet_name: Optional specific worksheet name.
    """
    return controller.merge_cells(range_address=range_address, sheet_name=sheet_name)


@mcp.tool()
def excel_unmerge_cells(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Unmerge previously merged cells in a range.
    Args:
        range_address: Range to unmerge (e.g. 'A1:E1').
        sheet_name: Optional specific worksheet name.
    """
    return controller.unmerge_cells(range_address=range_address, sheet_name=sheet_name)


# ── IMAGES & VISUAL EXPORT ──────────────────────────────────────────────────────

@mcp.tool()
def excel_add_image(
    image_path: str,
    cell: str = "B2",
    width: Optional[float] = None,
    height: Optional[float] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert an image or logo onto a worksheet anchored at a specific cell.
    Args:
        image_path: Absolute or relative path to image file (.png, .jpg, .bmp).
        cell: Anchor cell coordinate (default 'B2').
        width: Optional custom width in points.
        height: Optional custom height in points.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_image(
        image_path=image_path,
        cell=cell,
        width=width,
        height=height,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_export_range_image(range_address: str, output_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Take a high-resolution screenshot/image of an exact cell range and save as PNG.
    Args:
        range_address: Range to screenshot (e.g. 'B2:H20').
        output_path: Destination image path (.png).
        sheet_name: Optional specific worksheet name.
    """
    return controller.export_range_image(range_address=range_address, output_path=output_path, sheet_name=sheet_name)


# ── COMMENTS & NOTES ────────────────────────────────────────────────────────────

@mcp.tool()
def excel_add_comment(cell: str, text: str, author: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Add a review comment or note to a specific cell.
    Args:
        cell: Cell coordinate (e.g. 'C5').
        text: Comment message text.
        author: Optional author name.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_comment(cell=cell, text=text, author=author, sheet_name=sheet_name)


@mcp.tool()
def excel_list_comments(sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """List all comments and notes present across the active worksheet."""
    return controller.list_comments(sheet_name=sheet_name)


@mcp.tool()
def excel_delete_comment(cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Delete a comment from a specific cell.
    Args:
        cell: Cell coordinate (e.g. 'C5').
        sheet_name: Optional specific worksheet name.
    """
    return controller.delete_comment(cell=cell, sheet_name=sheet_name)


# ── SECURITY & PROTECTION ───────────────────────────────────────────────────────

@mcp.tool()
def excel_protect_sheet(password: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Protect the worksheet against unauthorized edits.
    Args:
        password: Optional protection password.
        sheet_name: Optional specific worksheet name.
    """
    return controller.protect_sheet(password=password, sheet_name=sheet_name)


@mcp.tool()
def excel_unprotect_sheet(password: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Unprotect a worksheet.
    Args:
        password: Optional protection password.
        sheet_name: Optional specific worksheet name.
    """
    return controller.unprotect_sheet(password=password, sheet_name=sheet_name)


# ── NAMED RANGES ────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_create_named_range(name: str, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Define a workbook-level named range for easier formula references.
    Args:
        name: Unique name identifier (e.g. 'SalesQ1', 'TaxRate').
        range_address: Cell range (e.g. 'B2:B50' or 'C1').
        sheet_name: Optional specific worksheet name.
    """
    return controller.create_named_range(name=name, range_address=range_address, sheet_name=sheet_name)


@mcp.tool()
def excel_list_named_ranges() -> List[Dict[str, Any]]:
    """List all defined named ranges in the workbook."""
    return controller.list_named_ranges()


@mcp.tool()
def excel_delete_named_range(name: str) -> Dict[str, Any]:
    """Delete a defined named range.
    Args:
        name: Name of the range to delete.
    """
    return controller.delete_named_range(name=name)


# ── CALCULATION & AUDIT ─────────────────────────────────────────────────────────

@mcp.tool()
def excel_calculate_now() -> Dict[str, Any]:
    """Force Microsoft Excel to recalculate all open formulas immediately."""
    return controller.calculate_now()


@mcp.tool()
def excel_list_formulas(range_address: Optional[str] = None, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """Audit and list all formula cells, their exact formula strings, and current calculated values.
    Args:
        range_address: Optional range to inspect (if None, inspects entire used range).
        sheet_name: Optional specific worksheet name.
    """
    return controller.list_formulas(range_address=range_address, sheet_name=sheet_name)


# ── PIVOT TABLES ────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_create_pivot_table(
    source_range: str,
    dest_cell: str,
    table_name: str = "SummaryPivot",
    row_fields: Optional[List[str]] = None,
    value_fields: Optional[List[str]] = None,
    col_fields: Optional[List[str]] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Create a dynamic native Excel Pivot Table from tabular data.
    Args:
        source_range: Source data range including headers (e.g. 'A1:E100').
        dest_cell: Destination cell where pivot table will be created (e.g. 'G2').
        table_name: Unique name for the pivot table.
        row_fields: List of column header names to group by in rows (e.g. ['Region', 'Product']).
        value_fields: List of numeric column header names to aggregate (e.g. ['Sales', 'Profit']).
        col_fields: List of column header names to place in columns (e.g. ['Quarter']).
        sheet_name: Optional specific worksheet name.
    """
    return controller.create_pivot_table(
        source_range=source_range,
        dest_cell=dest_cell,
        table_name=table_name,
        row_fields=row_fields,
        value_fields=value_fields,
        col_fields=col_fields,
        sheet_name=sheet_name
    )


# ── HEADLESS CSV & MERGE TOOLS ──────────────────────────────────────────────────

@mcp.tool()
def excel_csv_to_excel(csv_path: str, output_path: str, delimiter: str = ",", sheet_name: str = "Data") -> Dict[str, Any]:
    """Convert a CSV file into a styled Excel (.xlsx) file headlessly.
    Args:
        csv_path: Path to the source CSV file.
        output_path: Path for output .xlsx file.
        delimiter: CSV separator character (default ',').
        sheet_name: Sheet title name.
    """
    path = ExcelEngine.csv_to_excel(csv_path=csv_path, output_path=output_path, delimiter=delimiter, sheet_name=sheet_name)
    return {"status": "success", "file": path, "message": f"Converted CSV to .xlsx at {path}"}


@mcp.tool()
def excel_to_csv(file_path: str, output_csv_path: str, sheet_name: Optional[str] = None, delimiter: str = ",") -> Dict[str, Any]:
    """Extract a sheet from an Excel (.xlsx) file into a CSV file headlessly.
    Args:
        file_path: Source .xlsx file path.
        output_csv_path: Destination .csv file path.
        sheet_name: Optional sheet name (default active sheet).
        delimiter: CSV separator character (default ',').
    """
    path = ExcelEngine.excel_to_csv(file_path=file_path, output_csv_path=output_csv_path, sheet_name=sheet_name, delimiter=delimiter)
    return {"status": "success", "file": path, "message": f"Exported sheet to CSV at {path}"}


@mcp.tool()
def excel_merge_workbooks(file_paths: List[str], output_path: str) -> Dict[str, Any]:
    """Merge sheets from multiple .xlsx workbooks into a single consolidated workbook.
    Args:
        file_paths: List of file paths to source .xlsx workbooks.
        output_path: Destination path for consolidated workbook (.xlsx).
    """
    path = ExcelEngine.merge_workbooks(file_paths=file_paths, output_path=output_path)
    return {"status": "success", "file": path, "message": f"Merged {len(file_paths)} workbooks into {path}"}


# ── HYPERLINKS ──────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_add_hyperlink(
    cell: str,
    target: str,
    text_to_display: Optional[str] = None,
    tooltip: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Add a web hyperlink or an internal sheet navigation link to a cell.
    Args:
        cell: Anchor cell coordinate (e.g. 'A1', 'C10').
        target: Web URL ('https://...') or internal sheet address ('#Summary!A1').
        text_to_display: Text shown in the cell.
        tooltip: Optional screen tip shown on hover.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_hyperlink(
        cell=cell,
        target=target,
        text_to_display=text_to_display,
        tooltip=tooltip,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_remove_hyperlink(cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Remove a hyperlink from a cell while keeping its text content.
    Args:
        cell: Cell coordinate (e.g. 'A1').
        sheet_name: Optional specific worksheet name.
    """
    return controller.remove_hyperlink(cell=cell, sheet_name=sheet_name)


# ── TAB COLORS & VIEW SETTINGS ──────────────────────────────────────────────────

@mcp.tool()
def excel_set_tab_color(hex_color: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set the tab color of a worksheet for visual organization.
    Args:
        hex_color: Hex color string (e.g. '#2563EB', '#DC2626', '#16A34A').
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_tab_color(hex_color=hex_color, sheet_name=sheet_name)


@mcp.tool()
def excel_toggle_gridlines(show: bool = True) -> Dict[str, Any]:
    """Show or hide sheet gridlines (turn off for sleek executive dashboard presentations).
    Args:
        show: True to show gridlines, False to hide.
    """
    return controller.toggle_gridlines(show=show)


@mcp.tool()
def excel_toggle_headings(show: bool = True) -> Dict[str, Any]:
    """Show or hide row numbers and column letters.
    Args:
        show: True to show headings, False to hide.
    """
    return controller.toggle_headings(show=show)


# ── PAGE SETUP & PRINT ──────────────────────────────────────────────────────────

@mcp.tool()
def excel_set_page_setup(
    orientation: str = "landscape",
    paper_size: str = "A4",
    fit_to_pages: bool = True,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Configure page orientation, paper size, and print/PDF fit settings.
    Args:
        orientation: 'landscape' or 'portrait'.
        paper_size: 'A4' or 'Letter'.
        fit_to_pages: Whether to fit print output onto 1 page wide/tall.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_page_setup(
        orientation=orientation,
        paper_size=paper_size,
        fit_to_pages=fit_to_pages,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_print_area(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Define an explicit print/export area for the worksheet.
    Args:
        range_address: Range notation (e.g. 'A1:H45').
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_print_area(range_address=range_address, sheet_name=sheet_name)


@mcp.tool()
def excel_clear_print_area(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Clear custom print area restrictions."""
    return controller.clear_print_area(sheet_name=sheet_name)


# ── DATA ANALYSIS & WHAT-IF ─────────────────────────────────────────────────────

@mcp.tool()
def excel_find_all(search_term: str, match_case: bool = False, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """Search and find all occurrences of a term across the sheet with coordinates and values.
    Args:
        search_term: Text or numeric string to locate.
        match_case: Whether to match case.
        sheet_name: Optional specific worksheet name.
    """
    return controller.find_all(search_term=search_term, match_case=match_case, sheet_name=sheet_name)


@mcp.tool()
def excel_get_column_summary(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Compute rapid statistical metrics (Count, Sum, Average, Min, Max) on numeric cells.
    Args:
        range_address: Range of numbers (e.g. 'F4:F50').
        sheet_name: Optional specific worksheet name.
    """
    return controller.get_column_summary(range_address=range_address, sheet_name=sheet_name)


@mcp.tool()
def excel_remove_duplicates(
    range_address: str,
    columns: Optional[List[int]] = None,
    has_headers: bool = True,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Remove duplicate rows from a table or range based on specific column keys.
    Args:
        range_address: Data range (e.g. 'A1:E100').
        columns: Optional list of 1-based column indices to check for duplicates.
        has_headers: Whether the first row contains headers.
        sheet_name: Optional specific worksheet name.
    """
    return controller.remove_duplicates(
        range_address=range_address,
        columns=columns,
        has_headers=has_headers,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_goal_seek(
    target_cell: str,
    target_value: float,
    changing_cell: str,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Perform Goal Seek what-if analysis to find the necessary input to reach a target formula value.
    Args:
        target_cell: Cell containing the formula (e.g. 'B10').
        target_value: Desired outcome number (e.g. 50000.0).
        changing_cell: Cell that Excel will adjust (e.g. 'B4').
        sheet_name: Optional specific worksheet name.
    """
    return controller.goal_seek(
        target_cell=target_cell,
        target_value=target_value,
        changing_cell=changing_cell,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_text_to_columns(
    range_address: str,
    destination_cell: str,
    delimiter: str = ",",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Split delimited text in a single column into multiple columns.
    Args:
        range_address: Source single-column range (e.g. 'A1:A50').
        destination_cell: Target top-left cell to output columns (e.g. 'B1').
        delimiter: Separator character (e.g. ',', ';', '\t', ' ').
        sheet_name: Optional specific worksheet name.
    """
    return controller.text_to_columns(
        range_address=range_address,
        destination_cell=destination_cell,
        delimiter=delimiter,
        sheet_name=sheet_name
    )


# ── SPARKLINES & GANTT ROADMAPS ─────────────────────────────────────────────────

@mcp.tool()
def excel_add_sparkline(
    data_range: str,
    location_cell: str,
    sparkline_type: str = "line",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert a modern mini in-cell trend sparkline ('line', 'column', or 'win_loss').
    Args:
        data_range: Row or column of values (e.g. 'B2:G2').
        location_cell: Target cell where sparkline will appear (e.g. 'H2').
        sparkline_type: 'line', 'column', or 'win_loss'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_sparkline(
        data_range=data_range,
        location_cell=location_cell,
        sparkline_type=sparkline_type,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_create_gantt_block(
    start_cell: str,
    tasks: List[Dict[str, Any]],
    weeks_count: int = 6,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Create a project Gantt timeline grid with styled milestone duration bars.
    Args:
        start_cell: Top-left anchor cell (e.g. 'B3').
        tasks: List of task dicts: [{'name': 'Design', 'owner': 'Alice', 'start_week': 1, 'duration': 2, 'color': '#3B82F6'}].
        weeks_count: Total timeline span in weeks (default 6).
        sheet_name: Optional specific worksheet name.
    """
    return controller.create_gantt_block(
        start_cell=start_cell,
        tasks=tasks,
        weeks_count=weeks_count,
        sheet_name=sheet_name
    )


# ── HEADLESS WORKBOOK AUDIT & BATCH TOOLS ───────────────────────────────────────

@mcp.tool()
def excel_audit_workbook(file_path: str) -> Dict[str, Any]:
    """Perform a comprehensive structural audit of an .xlsx workbook (sheets, dimensions, formula counts, size).
    Args:
        file_path: Path to the .xlsx file.
    """
    return ExcelEngine.audit_workbook(file_path=file_path)


@mcp.tool()
def excel_batch_replace(file_path: str, replacements: Dict[str, str], output_path: str) -> Dict[str, Any]:
    """Batch replace placeholder template variables (e.g. {'{{YEAR}}': '2026'}) across all sheets.
    Args:
        file_path: Source template .xlsx file path.
        replacements: Dictionary of {token: replacement_value}.
        output_path: Destination .xlsx path.
    """
    path = ExcelEngine.batch_replace(file_path=file_path, replacements=replacements, output_path=output_path)
    return {"status": "success", "file": path, "message": f"Batch replaced placeholders into {path}"}


@mcp.tool()
def excel_extract_column(file_path: str, column: Union[str, int], sheet_name: Optional[str] = None) -> List[Any]:
    """Extract an entire column from an .xlsx file as a raw Python list.
    Args:
        file_path: Source .xlsx file path.
        column: Column letter (e.g. 'A', 'C') or 1-based index (e.g. 1, 3).
        sheet_name: Optional specific sheet name.
    """
    return ExcelEngine.extract_column(file_path=file_path, column=column, sheet_name=sheet_name)


# ── CELL STYLES & NUMBER PRESETS ────────────────────────────────────────────────

@mcp.tool()
def excel_apply_cell_style(range_address: str, style_name: str = "Good", sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Apply an official Excel built-in cell style ('Good', 'Bad', 'Neutral', 'Title', 'Heading 1', 'Total', 'Calculation', 'Input').
    Args:
        range_address: Target cell range (e.g. 'B2:F2', 'D5').
        style_name: Built-in style name.
        sheet_name: Optional specific worksheet name.
    """
    return controller.apply_cell_style(range_address=range_address, style_name=style_name, sheet_name=sheet_name)


@mcp.tool()
def excel_list_cell_styles() -> List[str]:
    """List all available workbook cell styles."""
    return controller.list_cell_styles()


@mcp.tool()
def excel_set_number_format_preset(range_address: str, preset: str = "currency_usd", sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Apply standard financial and business number format presets.
    Args:
        range_address: Cell range (e.g. 'C4:C50').
        preset: 'currency_usd', 'currency_eur', 'currency_mad', 'currency_int', 'percent', 'percent_int', 'accounting', 'date_short', 'date_long', 'integer', 'decimal_two'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_number_format_preset(range_address=range_address, preset=preset, sheet_name=sheet_name)


# ── ADVANCED SHEET WORKFLOWS ────────────────────────────────────────────────────

@mcp.tool()
def excel_duplicate_sheet(sheet_name: str, new_name: Optional[str] = None) -> Dict[str, Any]:
    """Clone an existing worksheet with all formulas, charts, styles, and data intact.
    Args:
        sheet_name: Name of the sheet to clone.
        new_name: Optional custom name for the new cloned sheet.
    """
    return controller.duplicate_sheet(sheet_name=sheet_name, new_name=new_name)


@mcp.tool()
def excel_move_sheet(sheet_name: str, before: Optional[str] = None, after: Optional[str] = None) -> Dict[str, Any]:
    """Reorder sheet tabs in the workbook.
    Args:
        sheet_name: Name of the sheet to move.
        before: Optional sheet before which to place.
        after: Optional sheet after which to place.
    """
    return controller.move_sheet(sheet_name=sheet_name, before=before, after=after)


@mcp.tool()
def excel_hide_unhide_sheet(sheet_name: str, hide: bool = True, very_hidden: bool = False) -> Dict[str, Any]:
    """Hide or unhide a worksheet (very_hidden prevents users from unhiding via standard Excel UI).
    Args:
        sheet_name: Name of the worksheet.
        hide: True to hide, False to unhide.
        very_hidden: If True, sets sheet to xlSheetVeryHidden.
    """
    return controller.hide_unhide_sheet(sheet_name=sheet_name, hide=hide, very_hidden=very_hidden)


# ── ADVANCED FORMULAS & CALCULATION ─────────────────────────────────────────────

@mcp.tool()
def excel_convert_formulas_to_values(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Paste/freeze formula results as static values in a range (locks data).
    Args:
        range_address: Range to convert (e.g. 'B4:F20').
        sheet_name: Optional specific worksheet name.
    """
    return controller.convert_formulas_to_values(range_address=range_address, sheet_name=sheet_name)


@mcp.tool()
def excel_evaluate_expression(expression: str) -> Dict[str, Any]:
    """Evaluate an arbitrary formula or expression directly in Excel's calculation engine (e.g. 'SUM(10, 20, 30)', 'NOW()').
    Args:
        expression: Excel formula string to compute.
    """
    return controller.evaluate_expression(expression=expression)


# ── VIEW, ZOOM & WINDOW SETTINGS ────────────────────────────────────────────────

@mcp.tool()
def excel_set_zoom(zoom_percentage: int = 100, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set the window zoom level (e.g. 75, 100, 120, 150).
    Args:
        zoom_percentage: Zoom percentage (10 to 400).
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_zoom(zoom_percentage=zoom_percentage, sheet_name=sheet_name)


@mcp.tool()
def excel_split_panes(row: Optional[int] = None, col: Optional[int] = None) -> Dict[str, Any]:
    """Split or unsplit the window view panes horizontally and/or vertically.
    Args:
        row: Row index for horizontal split (or None to unsplit).
        col: Column index for vertical split.
    """
    return controller.split_panes(row=row, col=col)


# ── EXECUTIVE VISUAL COMPONENTS ─────────────────────────────────────────────────

@mcp.tool()
def excel_add_callout_box(
    cell: str,
    text: str,
    title: Optional[str] = None,
    box_type: str = "info",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert a modern styled callout/alert banner shape with accent border and colors.
    Args:
        cell: Anchor cell (e.g. 'B2').
        text: Main message text.
        title: Optional banner title (e.g. 'NOTE', 'WARNING', 'CONFIDENTIAL').
        box_type: 'info', 'warning', 'success', or 'danger'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_callout_box(cell=cell, text=text, title=title, box_type=box_type, sheet_name=sheet_name)


@mcp.tool()
def excel_add_metric_tile(
    cell: str,
    label: str,
    value: str,
    delta: Optional[str] = None,
    status: str = "positive",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert a compact 2-row metric card with label, value, and colored delta indicator.
    Args:
        cell: Top-left cell anchor (e.g. 'B2').
        label: Metric label (e.g. 'Conversion Rate').
        value: Primary number (e.g. '4.8%').
        delta: Optional trend change string (e.g. '+0.6% MoM').
        status: 'positive', 'negative', or 'neutral'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_metric_tile(cell=cell, label=label, value=value, delta=delta, status=status, sheet_name=sheet_name)


# ── PRINT HEADERS, FOOTERS, FORMS & HTML EXPORT ─────────────────────────────────

@mcp.tool()
def excel_set_header_footer(
    header_left: Optional[str] = None,
    header_center: Optional[str] = None,
    header_right: Optional[str] = None,
    footer_left: Optional[str] = None,
    footer_center: Optional[str] = None,
    footer_right: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Configure printed page headers and footers (supports &P page number, &N total pages, &D date).
    Args:
        header_left: Left header text.
        header_center: Center header text.
        header_right: Right header text.
        footer_left: Left footer text.
        footer_center: Center footer text.
        footer_right: Right footer text.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_header_footer(
        header_left=header_left,
        header_center=header_center,
        header_right=header_right,
        footer_left=footer_left,
        footer_center=footer_center,
        footer_right=footer_right,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_add_checkbox(
    cell: str,
    caption: str = "Done",
    linked_cell: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert an interactive Form Control Checkbox connected to a boolean cell value.
    Args:
        cell: Anchor cell coordinate (e.g. 'B5').
        caption: Text label next to checkbox.
        linked_cell: Optional cell address (e.g. 'C5') that toggles TRUE/FALSE with checkbox clicks.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_checkbox(cell=cell, caption=caption, linked_cell=linked_cell, sheet_name=sheet_name)


@mcp.tool()
def excel_export_html(output_path: str, sheet_only: bool = False) -> Dict[str, Any]:
    """Export the active workbook or current sheet as a standalone interactive HTML web file.
    Args:
        output_path: Destination path (.html).
        sheet_only: If True, exports only current sheet; if False, exports workbook.
    """
    return controller.export_html(output_path=output_path, sheet_only=sheet_only)


# ── HEADLESS OOXML METADATA TOOLS ───────────────────────────────────────────────

@mcp.tool()
def excel_extract_metadata(file_path: str) -> Dict[str, Any]:
    """Extract core document metadata (Title, Author, Subject, Created, Modified) headlessly.
    Args:
        file_path: Path to the .xlsx file.
    """
    return ExcelEngine.extract_metadata(file_path=file_path)


@mcp.tool()
def excel_set_metadata(
    file_path: str,
    title: Optional[str] = None,
    author: Optional[str] = None,
    subject: Optional[str] = None,
    output_path: Optional[str] = None
) -> Dict[str, Any]:
    """Set or scrub document metadata (Title, Author, Subject) headlessly.
    Args:
        file_path: Source .xlsx file path.
        title: Optional document title.
        author: Optional author name.
        subject: Optional document subject.
        output_path: Optional destination path (if omitted, updates in place).
    """
    res = ExcelEngine.set_metadata(file_path=file_path, title=title, author=author, subject=subject, output_path=output_path)
    return {"status": "success", "file": res, "message": "Updated metadata."}


@mcp.tool()
def excel_add_watermark(
    text: str = "CONFIDENTIAL",
    font_size: int = 54,
    color: str = "#E2E8F0",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Add a semi-transparent diagonal watermark across the sheet (e.g. 'DRAFT', 'CONFIDENTIAL').
    Args:
        text: Watermark text.
        font_size: Point size (default 54).
        color: Hex color string.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_watermark(text=text, font_size=font_size, color=color, sheet_name=sheet_name)


@mcp.tool()
def excel_clear_all_formatting(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Reset all cell styling, colors, and borders across the active sheet back to raw defaults.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.clear_all_formatting(sheet_name=sheet_name)


def main():
    """Run the FastMCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
