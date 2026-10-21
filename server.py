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


# ── ROW & COLUMN OUTLINING / GROUPING ───────────────────────────────────────────

@mcp.tool()
def excel_group_rows(start_row: int, end_row: int, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Group rows together for expandable/collapsible outline hierarchy.
    Args:
        start_row: Starting row index (e.g. 5).
        end_row: Ending row index (e.g. 15).
        sheet_name: Optional specific worksheet name.
    """
    return controller.group_rows(start_row=start_row, end_row=end_row, sheet_name=sheet_name)


@mcp.tool()
def excel_ungroup_rows(start_row: int, end_row: int, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Ungroup rows to remove collapsible outline level.
    Args:
        start_row: Starting row index.
        end_row: Ending row index.
        sheet_name: Optional specific worksheet name.
    """
    return controller.ungroup_rows(start_row=start_row, end_row=end_row, sheet_name=sheet_name)


@mcp.tool()
def excel_group_columns(start_col: str, end_col: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Group columns together for collapsible outline views.
    Args:
        start_col: Starting column letter (e.g. 'C').
        end_col: Ending column letter (e.g. 'F').
        sheet_name: Optional specific worksheet name.
    """
    return controller.group_columns(start_col=start_col, end_col=end_col, sheet_name=sheet_name)


@mcp.tool()
def excel_ungroup_columns(start_col: str, end_col: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Ungroup columns to remove collapsible outline level.
    Args:
        start_col: Starting column letter.
        end_col: Ending column letter.
        sheet_name: Optional specific worksheet name.
    """
    return controller.ungroup_columns(start_col=start_col, end_col=end_col, sheet_name=sheet_name)


@mcp.tool()
def excel_show_outline_levels(row_levels: Optional[int] = None, col_levels: Optional[int] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set the outline expansion level (1 = all collapsed, 2+ = expanded details).
    Args:
        row_levels: Outline level for rows (1, 2, 3...).
        col_levels: Outline level for columns (1, 2, 3...).
        sheet_name: Optional specific worksheet name.
    """
    return controller.show_outline_levels(row_levels=row_levels, col_levels=col_levels, sheet_name=sheet_name)


# ── ADVANCED CONDITIONAL FORMATTING PRESETS ─────────────────────────────────────

@mcp.tool()
def excel_add_top_bottom_rule(
    range_address: str,
    rule_type: str = "top_percent",
    rank: int = 10,
    bg_color: str = "#DCFCE7",
    font_color: str = "#166534",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Highlight Top N or Bottom N items or percentages in a range.
    Args:
        range_address: Cell range (e.g. 'D2:D100').
        rule_type: 'top_items', 'top_percent', 'bottom_items', 'bottom_percent'.
        rank: Top or bottom rank count/percentage (default 10).
        bg_color: Background hex color.
        font_color: Text hex color.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_top_bottom_rule(range_address=range_address, rule_type=rule_type, rank=rank, bg_color=bg_color, font_color=font_color, sheet_name=sheet_name)


@mcp.tool()
def excel_add_above_below_average_rule(
    range_address: str,
    above: bool = True,
    bg_color: str = "#FEF3C7",
    font_color: str = "#92400E",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Highlight cells that are above or below the average value of a range.
    Args:
        range_address: Cell range (e.g. 'E2:E50').
        above: True to highlight above average, False for below average.
        bg_color: Background hex color.
        font_color: Text hex color.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_above_below_average_rule(range_address=range_address, above=above, bg_color=bg_color, font_color=font_color, sheet_name=sheet_name)


@mcp.tool()
def excel_add_icon_set_rule(
    range_address: str,
    icon_set_type: str = "3_traffic_lights",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Add visual icon set ratings directly inside cells.
    Args:
        range_address: Cell range (e.g. 'F2:F20').
        icon_set_type: '3_traffic_lights', '3_arrows', '3_flags', '3_symbols', '4_arrows', '5_quarters'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_icon_set_rule(range_address=range_address, icon_set_type=icon_set_type, sheet_name=sheet_name)


@mcp.tool()
def excel_highlight_duplicate_values(
    range_address: str,
    bg_color: str = "#FEE2E2",
    font_color: str = "#991B1B",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Automatically highlight all duplicate values across a cell range.
    Args:
        range_address: Cell range (e.g. 'A2:A100').
        bg_color: Fill hex color (default soft red).
        font_color: Text hex color.
        sheet_name: Optional specific worksheet name.
    """
    return controller.highlight_duplicate_values(range_address=range_address, bg_color=bg_color, font_color=font_color, sheet_name=sheet_name)


# ── FORMULA AUDITING & TRACER ARROWS ───────────────────────────────────────────

@mcp.tool()
def excel_trace_precedents(cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Draw graphic audit tracer arrows pointing to cells that feed data into this formula.
    Args:
        cell: Coordinate of the formula cell (e.g. 'C10').
        sheet_name: Optional specific worksheet name.
    """
    return controller.trace_precedents(cell=cell, sheet_name=sheet_name)


@mcp.tool()
def excel_trace_dependents(cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Draw graphic audit tracer arrows pointing to formulas that depend on this cell value.
    Args:
        cell: Coordinate of the cell (e.g. 'B4').
        sheet_name: Optional specific worksheet name.
    """
    return controller.trace_dependents(cell=cell, sheet_name=sheet_name)


@mcp.tool()
def excel_remove_tracer_arrows(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Clear all precedent and dependent audit tracer arrows from the worksheet.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.remove_tracer_arrows(sheet_name=sheet_name)


# ── WORKBOOK STRUCTURE SECURITY ─────────────────────────────────────────────────

@mcp.tool()
def excel_protect_workbook(password: Optional[str] = None, lock_structure: bool = True, lock_windows: bool = False) -> Dict[str, Any]:
    """Lock workbook structure to prevent adding, deleting, hiding, or renaming sheets.
    Args:
        password: Optional protection password.
        lock_structure: Whether to lock sheet structure (default True).
        lock_windows: Whether to lock window sizing/positioning.
    """
    return controller.protect_workbook(password=password, lock_structure=lock_structure, lock_windows=lock_windows)


@mcp.tool()
def excel_unprotect_workbook(password: Optional[str] = None) -> Dict[str, Any]:
    """Unlock workbook structure.
    Args:
        password: Optional protection password.
    """
    return controller.unprotect_workbook(password=password)


# ── SHAPES, ARROWS & BADGES ─────────────────────────────────────────────────────

@mcp.tool()
def excel_add_shape(
    shape_type: str = "rounded_rectangle",
    cell: str = "B2",
    width: float = 120,
    height: float = 40,
    fill_color: str = "#2563EB",
    line_color: str = "#1D4ED8",
    text: Optional[str] = None,
    font_color: str = "#FFFFFF",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert a modern native shape or flow badge with custom text and styling.
    Args:
        shape_type: 'rectangle', 'rounded_rectangle', 'oval', or 'arrow_right'.
        cell: Anchor cell coordinate (e.g. 'B2').
        width: Width in points.
        height: Height in points.
        fill_color: Hex fill color.
        line_color: Hex border color.
        text: Optional label text inside the shape.
        font_color: Hex text color.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_shape(
        shape_type=shape_type,
        cell=cell,
        width=width,
        height=height,
        fill_color=fill_color,
        line_color=line_color,
        text=text,
        font_color=font_color,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_delete_shape(shape_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Delete a shape, badge, or image by its exact name.
    Args:
        shape_name: Name identifier of the shape.
        sheet_name: Optional specific worksheet name.
    """
    return controller.delete_shape(shape_name=shape_name, sheet_name=sheet_name)


@mcp.tool()
def excel_list_shapes(sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """List all shapes, badges, and graphic objects on the sheet with position/size.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.list_shapes(sheet_name=sheet_name)


# ── SPECIAL CELLS INSPECTION & EDIT PERMISSIONS ─────────────────────────────────

@mcp.tool()
def excel_get_special_cells_summary(cell_type: str = "blanks", sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Inspect and count special cells across the worksheet: 'blanks', 'constants', 'formulas', or 'comments'.
    Args:
        cell_type: 'blanks', 'constants', 'formulas', or 'comments'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.get_special_cells_summary(cell_type=cell_type, sheet_name=sheet_name)


@mcp.tool()
def excel_lock_unlock_cells(range_address: str, locked: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set edit permissions for specific cells when sheet protection is active.
    Args:
        range_address: Range to set (e.g. 'B4:D20').
        locked: True to protect/lock, False to allow editing even on protected sheets.
        sheet_name: Optional specific worksheet name.
    """
    return controller.lock_unlock_cells(range_address=range_address, locked=locked, sheet_name=sheet_name)


# ── ADVANCED PRINTING & PAGE LAYOUT TITLES ──────────────────────────────────────

@mcp.tool()
def excel_set_repeat_rows_cols(
    repeat_rows: Optional[str] = None,
    repeat_cols: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Set repeating header rows and/or columns for multi-page printouts (Print Titles, e.g. '$1:$1').
    Args:
        repeat_rows: Row range string (e.g. '$1:$2').
        repeat_cols: Column range string (e.g. '$A:$A').
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_repeat_rows_cols(repeat_rows=repeat_rows, repeat_cols=repeat_cols, sheet_name=sheet_name)


@mcp.tool()
def excel_insert_page_break(before_cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Insert manual horizontal and vertical page breaks before a specific cell.
    Args:
        before_cell: Anchor cell coordinate (e.g. 'A25').
        sheet_name: Optional specific worksheet name.
    """
    return controller.insert_page_break(before_cell=before_cell, sheet_name=sheet_name)


@mcp.tool()
def excel_reset_page_breaks(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Reset all manual page breaks back to default pagination.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.reset_page_breaks(sheet_name=sheet_name)


# ── CALCULATION ENGINE MODES ────────────────────────────────────────────────────

@mcp.tool()
def excel_set_calculation_mode(mode: str = "automatic") -> Dict[str, Any]:
    """Configure Excel calculation mode ('automatic', 'manual', or 'semiautomatic').
    Args:
        mode: Calculation engine mode ('automatic', 'manual', 'semiautomatic').
    """
    return controller.set_calculation_mode(mode=mode)


@mcp.tool()
def excel_get_calculation_mode() -> Dict[str, Any]:
    """Get current Excel calculation mode."""
    return controller.get_calculation_mode()


# ── DATA TRANSFORMATIONS & UTILITIES ────────────────────────────────────────────

@mcp.tool()
def excel_transpose_range(source_range: str, dest_cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Transpose (swap rows and columns) of a cell range to a destination cell.
    Args:
        source_range: Source range (e.g. 'B2:F3').
        dest_cell: Target top-left cell (e.g. 'H2').
        sheet_name: Optional specific worksheet name.
    """
    return controller.transpose_range(source_range=source_range, dest_cell=dest_cell, sheet_name=sheet_name)


@mcp.tool()
def excel_fill_series(range_address: str, fill_type: str = "linear", step: float = 1.0, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Auto-fill a sequence or series (linear numbers, growth exponential, or dates).
    Args:
        range_address: Target series range (e.g. 'A1:A20').
        fill_type: 'linear', 'growth', or 'date'.
        step: Increment step value (default 1.0).
        sheet_name: Optional specific worksheet name.
    """
    return controller.fill_series(range_address=range_address, fill_type=fill_type, step=step, sheet_name=sheet_name)


# ── VIEW MODES & AUDITING ───────────────────────────────────────────────────────

@mcp.tool()
def excel_set_view_mode(view_mode: str = "normal", sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set window view mode: 'normal', 'page_break_preview', or 'page_layout'.
    Args:
        view_mode: 'normal', 'page_break_preview', or 'page_layout'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_view_mode(view_mode=view_mode, sheet_name=sheet_name)


@mcp.tool()
def excel_get_formula_errors(sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """Scan and list all formula error cells (#DIV/0!, #N/A, #VALUE!, #REF!) across the worksheet.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.get_formula_errors(sheet_name=sheet_name)


# ── WHAT-IF DATA TABLES & SLICERS ───────────────────────────────────────────────

@mcp.tool()
def excel_create_two_variable_data_table(
    table_range: str,
    row_input_cell: str,
    col_input_cell: str,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Generate a 2-variable sensitivity Data Table matrix (e.g. Price vs Volume sensitivity).
    Args:
        table_range: Data table range matrix (e.g. 'B4:F10').
        row_input_cell: Target input cell for row variables (e.g. 'B1').
        col_input_cell: Target input cell for column variables (e.g. 'B2').
        sheet_name: Optional specific worksheet name.
    """
    return controller.create_two_variable_data_table(
        table_range=table_range,
        row_input_cell=row_input_cell,
        col_input_cell=col_input_cell,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_add_table_slicer(
    table_name: str,
    column_name: str,
    slicer_name: Optional[str] = None,
    left: float = 400,
    top: float = 60,
    width: float = 160,
    height: float = 220,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert an interactive graphic visual Slicer button box for filtering an Excel Table.
    Args:
        table_name: Name of the table.
        column_name: Column to create slicer for.
        slicer_name: Optional custom slicer identifier.
        left: Horizontal position offset in points.
        top: Vertical position offset in points.
        width: Slicer width.
        height: Slicer height.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_table_slicer(
        table_name=table_name,
        column_name=column_name,
        slicer_name=slicer_name,
        left=left,
        top=top,
        width=width,
        height=height,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_convert_table_to_range(table_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Convert an official Excel Table back to a standard cell range (keeps data & formatting).
    Args:
        table_name: Name of the table to convert.
        sheet_name: Optional specific worksheet name.
    """
    return controller.convert_table_to_range(table_name=table_name, sheet_name=sheet_name)


# ── SHEET EXPORT & CELL CONTENT CLEARING ────────────────────────────────────────

@mcp.tool()
def excel_export_sheet_to_csv(output_path: str, sheet_name: Optional[str] = None) -> str:
    """Export the active worksheet directly to a CSV file.
    Args:
        output_path: Destination path for the CSV file.
        sheet_name: Optional specific worksheet name.
    """
    return controller.export_sheet_to_csv(output_path=output_path, sheet_name=sheet_name)


@mcp.tool()
def excel_get_merged_ranges(sheet_name: Optional[str] = None) -> List[str]:
    """Find and list all merged cell range addresses on the worksheet.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.get_merged_ranges(sheet_name=sheet_name)


@mcp.tool()
def excel_clear_cell_contents(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Clear values and formulas in a range while keeping cell formatting intact.
    Args:
        range_address: Target cell range (e.g. 'A2:C20').
        sheet_name: Optional specific worksheet name.
    """
    return controller.clear_cell_contents(range_address=range_address, sheet_name=sheet_name)


@mcp.tool()
def excel_clear_cell_formats(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Clear cell formatting (colors, borders, styles) while keeping values and formulas intact.
    Args:
        range_address: Target cell range (e.g. 'A2:C20').
        sheet_name: Optional specific worksheet name.
    """
    return controller.clear_cell_formats(range_address=range_address, sheet_name=sheet_name)



# ── DATA CONNECTIONS & POWER QUERY ──────────────────────────────────────────────

@mcp.tool()
def excel_refresh_all_connections() -> Dict[str, Any]:
    """Refresh all external data connections, Power Query queries, and PivotTables in the workbook."""
    return controller.refresh_all_connections()


@mcp.tool()
def excel_get_data_connections() -> List[Dict[str, Any]]:
    """List all external data connections (ODBC, OLEDB, Power Query) in the workbook."""
    return controller.get_data_connections()


@mcp.tool()
def excel_refresh_connection(connection_name: str) -> Dict[str, Any]:
    """Refresh a specific external data connection by name.
    Args:
        connection_name: Name of the connection to refresh.
    """
    return controller.refresh_connection(connection_name=connection_name)


# ── ADVANCED CHART CUSTOMIZATION & EXPORT ────────────────────────────────────────

@mcp.tool()
def excel_set_chart_secondary_axis(
    chart_name: str,
    series_index: int = 2,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Move a chart data series to the secondary vertical (Y) axis.
    Args:
        chart_name: Name of the chart object.
        series_index: 1-based index of the series to move.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_chart_secondary_axis(chart_name=chart_name, series_index=series_index, sheet_name=sheet_name)


@mcp.tool()
def excel_add_chart_trendline(
    chart_name: str,
    series_index: int = 1,
    trendline_type: str = "linear",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Add a trendline (linear, exponential, logarithmic, moving_average, polynomial) to a chart series.
    Args:
        chart_name: Name of the chart object.
        series_index: 1-based index of the series to add trendline to.
        trendline_type: Trendline type ('linear', 'exponential', 'logarithmic', 'moving_average', 'polynomial').
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_chart_trendline(
        chart_name=chart_name,
        series_index=series_index,
        trendline_type=trendline_type,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_export_chart_as_image(
    chart_name: str,
    output_image_path: str,
    sheet_name: Optional[str] = None
) -> str:
    """Export a specific chart as a high-resolution PNG or JPG image file.
    Args:
        chart_name: Name of the chart object.
        output_image_path: Destination path for the image (e.g. 'chart.png').
        sheet_name: Optional specific worksheet name.
    """
    return controller.export_chart_as_image(chart_name=chart_name, output_image_path=output_image_path, sheet_name=sheet_name)


# ── WHAT-IF SCENARIO MANAGER ────────────────────────────────────────────────────

@mcp.tool()
def excel_add_scenario(
    name: str,
    changing_cells: str,
    values: List[Any],
    comment: str = "",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Create a What-If Scenario with defined changing cell values (e.g. Best Case, Worst Case).
    Args:
        name: Name of the scenario.
        changing_cells: Range address of cells that vary (e.g. 'B2:B4').
        values: List of values to apply to changing cells in this scenario.
        comment: Optional scenario description.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_scenario(
        name=name,
        changing_cells=changing_cells,
        values=values,
        comment=comment,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_show_scenario(name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Apply and display a saved What-If Scenario on the worksheet.
    Args:
        name: Name of the scenario to activate.
        sheet_name: Optional specific worksheet name.
    """
    return controller.show_scenario(name=name, sheet_name=sheet_name)


@mcp.tool()
def excel_list_scenarios(sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """List all What-If Scenarios saved on the worksheet.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.list_scenarios(sheet_name=sheet_name)


# ── CELL INSPECTION & AUDITING TOOLS ────────────────────────────────────────────

@mcp.tool()
def excel_find_cells_with_comments(sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """Find and return all cells containing comments or notes on the sheet.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.find_cells_with_comments(sheet_name=sheet_name)


@mcp.tool()
def excel_find_cells_with_data_validation(sheet_name: Optional[str] = None) -> List[str]:
    """Find and return all cell addresses that have Data Validation rules applied.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.find_cells_with_data_validation(sheet_name=sheet_name)


# ── ADVANCED PAGE SETUP & PRINT METRICS ──────────────────────────────────────────

@mcp.tool()
def excel_set_sheet_margins(
    top_inches: float = 0.75,
    bottom_inches: float = 0.75,
    left_inches: float = 0.7,
    right_inches: float = 0.7,
    header_inches: float = 0.3,
    footer_inches: float = 0.3,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Configure page margins in inches for printing or PDF export.
    Args:
        top_inches: Top margin.
        bottom_inches: Bottom margin.
        left_inches: Left margin.
        right_inches: Right margin.
        header_inches: Header margin.
        footer_inches: Footer margin.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_sheet_margins(
        top_inches=top_inches,
        bottom_inches=bottom_inches,
        left_inches=left_inches,
        right_inches=right_inches,
        header_inches=header_inches,
        footer_inches=footer_inches,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_sheet_orientation_and_paper(
    orientation: str = "landscape",
    paper_size: str = "A4",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Set page orientation ('portrait' or 'landscape') and paper size ('A4', 'Letter', 'Legal', 'A3').
    Args:
        orientation: 'portrait' or 'landscape'.
        paper_size: 'A4', 'Letter', 'Legal', or 'A3'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_sheet_orientation_and_paper(
        orientation=orientation,
        paper_size=paper_size,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_sheet_fit_to_pages(
    fit_to_wide: int = 1,
    fit_to_tall: int = 1,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Fit worksheet to a specific number of pages wide and tall for printing / PDF export.
    Args:
        fit_to_wide: Maximum pages wide (e.g. 1).
        fit_to_tall: Maximum pages tall (e.g. 1).
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_sheet_fit_to_pages(
        fit_to_wide=fit_to_wide,
        fit_to_tall=fit_to_tall,
        sheet_name=sheet_name
    )


# ── HEADLESS OOXML ENHANCEMENTS ──────────────────────────────────────────────────

@mcp.tool()
def excel_compare_two_workbooks(file_path_1: str, file_path_2: str) -> Dict[str, Any]:
    """Compare two .xlsx workbooks headlessly and return structural and cell value differences.
    Args:
        file_path_1: First workbook file path.
        file_path_2: Second workbook file path.
    """
    return ExcelEngine.compare_two_workbooks(file_path_1=file_path_1, file_path_2=file_path_2)


@mcp.tool()
def excel_extract_hyperlinks(file_path: str, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """Extract all embedded URLs and hyperlinks from an .xlsx workbook headlessly.
    Args:
        file_path: Path to the .xlsx file.
        sheet_name: Optional specific worksheet name.
    """
    return ExcelEngine.extract_hyperlinks(file_path=file_path, sheet_name=sheet_name)


@mcp.tool()
def excel_unhide_all_sheets_headless(file_path: str, output_path: Optional[str] = None) -> str:
    """Unhide all hidden and very hidden worksheets in an .xlsx file headlessly.
    Args:
        file_path: Input .xlsx file path.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.unhide_all_sheets_headless(file_path=file_path, output_path=output_path)


# ── ADVANCED AUTOFILTER MANAGEMENT ──────────────────────────────────────────────

@mcp.tool()
def excel_clear_sheet_filters(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Clear active filter selections on the worksheet without removing filter dropdown arrows.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.clear_sheet_filters(sheet_name=sheet_name)


@mcp.tool()
def excel_apply_autofilter_column(
    range_address: str,
    field_index: int = 1,
    criteria: str = "",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Apply a filter criteria to a specific column within an AutoFilter range.
    Args:
        range_address: Target table or header range (e.g. 'A1:E100').
        field_index: 1-based column offset in the range.
        criteria: Filter condition (e.g. '>100', 'Pending', '=Completed').
        sheet_name: Optional specific worksheet name.
    """
    return controller.apply_autofilter_column(
        range_address=range_address,
        field_index=field_index,
        criteria=criteria,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_remove_autofilter(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Remove AutoFilter mode and dropdown arrows completely from the worksheet.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.remove_autofilter(sheet_name=sheet_name)


# ── ADVANCED ALIGNMENT & TEXT FITTING ────────────────────────────────────────────

@mcp.tool()
def excel_wrap_text(range_address: str, wrap: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Enable or disable text wrapping on a range of cells.
    Args:
        range_address: Target range (e.g. 'B2:B20').
        wrap: Whether text should wrap to multiple lines.
        sheet_name: Optional specific worksheet name.
    """
    return controller.wrap_text(range_address=range_address, wrap=wrap, sheet_name=sheet_name)


@mcp.tool()
def excel_shrink_to_fit(range_address: str, shrink: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Enable or disable Shrink-to-Fit on cells to automatically fit text within column width.
    Args:
        range_address: Target range (e.g. 'A1:C10').
        shrink: Whether to shrink font size visually to fit cell boundaries.
        sheet_name: Optional specific worksheet name.
    """
    return controller.shrink_to_fit(range_address=range_address, shrink=shrink, sheet_name=sheet_name)


@mcp.tool()
def excel_set_cell_orientation(range_address: str, degrees: int = 0, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set text rotation orientation angle in degrees (-90 to +90).
    Args:
        range_address: Target range (e.g. 'A1:G1').
        degrees: Angle in degrees (-90 to 90, e.g. 45 for diagonal headers).
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_cell_orientation(range_address=range_address, degrees=degrees, sheet_name=sheet_name)


@mcp.tool()
def excel_flash_fill_column(range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Trigger Excel's intelligent Flash Fill down a column based on pattern examples.
    Args:
        range_address: Range column with examples and blanks to fill (e.g. 'C2:C50').
        sheet_name: Optional specific worksheet name.
    """
    return controller.flash_fill_column(range_address=range_address, sheet_name=sheet_name)


# ── ADVANCED CHART ELEMENTS & STYLING ───────────────────────────────────────────

@mcp.tool()
def excel_set_chart_title(
    chart_name: str,
    title: str,
    font_size: float = 14.0,
    bold: bool = True,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Configure chart title text, font size, and bold formatting directly.
    Args:
        chart_name: Name of the chart object.
        title: Title string.
        font_size: Font size in points.
        bold: Whether title font is bold.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_chart_title(
        chart_name=chart_name,
        title=title,
        font_size=font_size,
        bold=bold,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_chart_legend(
    chart_name: str,
    position: str = "right",
    show_legend: bool = True,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Configure chart legend visibility and placement ('top', 'bottom', 'left', 'right', 'corner').
    Args:
        chart_name: Name of the chart object.
        position: Placement ('top', 'bottom', 'left', 'right', 'corner').
        show_legend: Whether legend is displayed.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_chart_legend(
        chart_name=chart_name,
        position=position,
        show_legend=show_legend,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_chart_axes_titles(
    chart_name: str,
    x_title: Optional[str] = None,
    y_title: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Set horizontal category (X) and vertical value (Y) axis title labels on a chart.
    Args:
        chart_name: Name of the chart object.
        x_title: Optional horizontal X-axis title label.
        y_title: Optional vertical Y-axis title label.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_chart_axes_titles(
        chart_name=chart_name,
        x_title=x_title,
        y_title=y_title,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_chart_style(chart_name: str, style_id: int = 201, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Apply a built-in modern visual chart style preset by ID (e.g. 201, 202, 215).
    Args:
        chart_name: Name of the chart object.
        style_id: Numeric Excel chart style identifier.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_chart_style(chart_name=chart_name, style_id=style_id, sheet_name=sheet_name)


# ── SECURITY & FORM CELL LOCKING ────────────────────────────────────────────────

@mcp.tool()
def excel_lock_cells(range_address: str, locked: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Lock or unlock cells. Unlocked cells can still be edited when worksheet protection is active.
    Args:
        range_address: Target cell range (e.g. 'C4:C20').
        locked: True to lock, False to unlock for editing in protected sheets.
        sheet_name: Optional specific worksheet name.
    """
    return controller.lock_cells(range_address=range_address, locked=locked, sheet_name=sheet_name)


@mcp.tool()
def excel_hide_formulas(range_address: str, hidden: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Hide formulas in the formula bar for cells when worksheet protection is active.
    Args:
        range_address: Target cell range (e.g. 'D2:D50').
        hidden: True to hide formulas from users.
        sheet_name: Optional specific worksheet name.
    """
    return controller.hide_formulas(range_address=range_address, hidden=hidden, sheet_name=sheet_name)


@mcp.tool()
def excel_is_sheet_protected(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Query worksheet protection status (contents, objects, scenarios).
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.is_sheet_protected(sheet_name=sheet_name)


# ── FORMULA AUDITING & PRECEDENTS/DEPENDENTS ────────────────────────────────────

@mcp.tool()
def excel_get_precedents(cell_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Return the range addresses of all cells directly referenced by the formula in this cell.
    Args:
        cell_address: Cell containing a formula (e.g. 'E10').
        sheet_name: Optional specific worksheet name.
    """
    return controller.get_precedents(cell_address=cell_address, sheet_name=sheet_name)


@mcp.tool()
def excel_get_dependents(cell_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Return the range addresses of all cells that depend on the value in this cell.
    Args:
        cell_address: Target cell address (e.g. 'B4').
        sheet_name: Optional specific worksheet name.
    """
    return controller.get_dependents(cell_address=cell_address, sheet_name=sheet_name)


# ── ROW & COLUMN VISIBILITY ─────────────────────────────────────────────────────

@mcp.tool()
def excel_set_row_visibility(
    row_index: Union[int, str],
    hidden: bool = True,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Hide or unhide specific rows.
    Args:
        row_index: Row number or range (e.g. 5 or '5:10').
        hidden: True to hide, False to unhide.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_row_visibility(row_index=row_index, hidden=hidden, sheet_name=sheet_name)


@mcp.tool()
def excel_set_column_visibility(
    column: Union[int, str],
    hidden: bool = True,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Hide or unhide specific columns.
    Args:
        column: Column letter or range (e.g. 'C' or 'C:E').
        hidden: True to hide, False to unhide.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_column_visibility(column=column, hidden=hidden, sheet_name=sheet_name)


# ── HEADLESS OOXML EXPANSIONS ───────────────────────────────────────────────────

@mcp.tool()
def excel_copy_sheet_headless(
    file_path: str,
    source_sheet: str,
    new_sheet_name: str,
    output_path: Optional[str] = None
) -> str:
    """Duplicate a worksheet with all its data, styling, and formulas headlessly.
    Args:
        file_path: Path to the .xlsx file.
        source_sheet: Name of the sheet to clone.
        new_sheet_name: Name for the cloned sheet.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.copy_sheet_headless(
        file_path=file_path,
        source_sheet=source_sheet,
        new_sheet_name=new_sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_protect_sheet_headless(
    file_path: str,
    password: str = "",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Enable password protection on a worksheet headlessly.
    Args:
        file_path: Path to the .xlsx file.
        password: Optional protection password.
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.protect_sheet_headless(
        file_path=file_path,
        password=password,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_delete_sheet_headless(
    file_path: str,
    sheet_name: str,
    output_path: Optional[str] = None
) -> str:
    """Delete a worksheet from a workbook headlessly.
    Args:
        file_path: Path to the .xlsx file.
        sheet_name: Sheet to delete.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.delete_sheet_headless(
        file_path=file_path,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_set_sheet_color_headless(
    file_path: str,
    sheet_name: str,
    hex_color: str,
    output_path: Optional[str] = None
) -> str:
    """Set worksheet tab color headlessly.
    Args:
        file_path: Path to the .xlsx file.
        sheet_name: Sheet to colorize.
        hex_color: Hex color string (e.g. '#2563EB' or '2563EB').
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.set_sheet_color_headless(
        file_path=file_path,
        sheet_name=sheet_name,
        hex_color=hex_color,
        output_path=output_path
    )


# ── FREEZE PANES & RANGE COPY CONTROLS ──────────────────────────────────────────

@mcp.tool()
def excel_freeze_custom_panes(row: int, column: int, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Freeze panes at an arbitrary row and column intersection.
    Args:
        row: Row index below which to freeze (e.g. 3).
        column: Column index to the right of which to freeze (e.g. 2 for col B).
        sheet_name: Optional specific worksheet name.
    """
    return controller.freeze_custom_panes(row=row, column=column, sheet_name=sheet_name)


@mcp.tool()
def excel_copy_range(source_range: str, destination_cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Copy a range with values, formulas, and formatting to a destination cell.
    Args:
        source_range: Source range address (e.g. 'A1:D10').
        destination_cell: Top-left cell where to paste (e.g. 'F1').
        sheet_name: Optional specific worksheet name.
    """
    return controller.copy_range(source_range=source_range, destination_cell=destination_cell, sheet_name=sheet_name)



@mcp.tool()
def excel_freeze_top_row(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Freeze the top header row so it stays visible during vertical scrolling.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.freeze_top_row(sheet_name=sheet_name)


@mcp.tool()
def excel_freeze_first_column(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Freeze the first column (Column A) so it stays visible during horizontal scrolling.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.freeze_first_column(sheet_name=sheet_name)


# ── ADVANCED AUTOFIT CONTROLS ───────────────────────────────────────────────────

@mcp.tool()
def excel_autofit_rows(range_address: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Auto-fit row heights to match cell content.
    Args:
        range_address: Optional range to fit (defaults to used range).
        sheet_name: Optional specific worksheet name.
    """
    return controller.autofit_rows(range_address=range_address, sheet_name=sheet_name)


@mcp.tool()
def excel_autofit_all(sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Auto-fit all rows and columns across the entire used range in one command.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.autofit_all(sheet_name=sheet_name)


# ── TYPOGRAPHY & FONT STYLES ────────────────────────────────────────────────────

@mcp.tool()
def excel_set_strikethrough(range_address: str, strikethrough: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Apply or remove strikethrough text formatting on a range.
    Args:
        range_address: Target cell range (e.g. 'A2:A10').
        strikethrough: True to apply, False to remove.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_strikethrough(range_address=range_address, strikethrough=strikethrough, sheet_name=sheet_name)


@mcp.tool()
def excel_set_subscript(range_address: str, subscript: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Apply or remove subscript formatting on a range (e.g. chemical formulas).
    Args:
        range_address: Target cell range (e.g. 'B3').
        subscript: True to apply, False to remove.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_subscript(range_address=range_address, subscript=subscript, sheet_name=sheet_name)


@mcp.tool()
def excel_set_superscript(range_address: str, superscript: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Apply or remove superscript formatting on a range (e.g. mathematical powers, notes).
    Args:
        range_address: Target cell range (e.g. 'C5').
        superscript: True to apply, False to remove.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_superscript(range_address=range_address, superscript=superscript, sheet_name=sheet_name)


@mcp.tool()
def excel_set_underline_style(
    range_address: str,
    underline_type: str = "single",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Set underline style ('none', 'single', 'double', 'single_accounting', 'double_accounting').
    Args:
        range_address: Target cell range (e.g. 'A1:E1').
        underline_type: 'none', 'single', 'double', 'single_accounting', 'double_accounting'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_underline_style(range_address=range_address, underline_type=underline_type, sheet_name=sheet_name)


# ── WINDOW & DISPLAY TOGGLES ────────────────────────────────────────────────────

@mcp.tool()
def excel_toggle_formula_bar(show: bool = True) -> Dict[str, Any]:
    """Show or hide Excel's formula bar.
    Args:
        show: True to show, False to hide.
    """
    return controller.toggle_formula_bar(show=show)


@mcp.tool()
def excel_toggle_status_bar(show: bool = True) -> Dict[str, Any]:
    """Show or hide Excel's bottom status bar.
    Args:
        show: True to show, False to hide.
    """
    return controller.toggle_status_bar(show=show)


@mcp.tool()
def excel_toggle_scroll_bars(horizontal: bool = True, vertical: bool = True) -> Dict[str, Any]:
    """Show or hide the horizontal and vertical scrollbars in the active window.
    Args:
        horizontal: True to show horizontal scrollbar.
        vertical: True to show vertical scrollbar.
    """
    return controller.toggle_scroll_bars(horizontal=horizontal, vertical=vertical)


@mcp.tool()
def excel_toggle_sheet_tabs(show: bool = True) -> Dict[str, Any]:
    """Show or hide worksheet tab bar at the bottom of the window.
    Args:
        show: True to show, False to hide.
    """
    return controller.toggle_sheet_tabs(show=show)


# ── THREADED COMMENTS (EXCEL 365 / 2021) ────────────────────────────────────────

@mcp.tool()
def excel_add_threaded_comment(cell: str, text: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Add a modern threaded comment to a cell.
    Args:
        cell: Target cell address (e.g. 'B4').
        text: Comment message text.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_threaded_comment(cell=cell, text=text, sheet_name=sheet_name)


@mcp.tool()
def excel_delete_threaded_comment(cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Delete a threaded or standard comment on a cell.
    Args:
        cell: Target cell address.
        sheet_name: Optional specific worksheet name.
    """
    return controller.delete_threaded_comment(cell=cell, sheet_name=sheet_name)


@mcp.tool()
def excel_list_threaded_comments(sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """List all threaded comments on the worksheet with author, text, date, and replies count.
    Args:
        sheet_name: Optional specific worksheet name.
    """
    return controller.list_threaded_comments(sheet_name=sheet_name)


# ── WORKBOOK STRUCTURE & WINDOWS PROTECTION ─────────────────────────────────────

@mcp.tool()
def excel_protect_workbook_structure(
    password: str = "",
    protect_structure: bool = True,
    protect_windows: bool = False
) -> Dict[str, Any]:
    """Protect workbook structure (preventing adding, deleting, hiding, or renaming sheets).
    Args:
        password: Optional protection password.
        protect_structure: Whether to lock sheet structure.
        protect_windows: Whether to lock window layout.
    """
    return controller.protect_workbook_structure(
        password=password,
        protect_structure=protect_structure,
        protect_windows=protect_windows
    )


@mcp.tool()
def excel_unprotect_workbook_structure(password: str = "") -> Dict[str, Any]:
    """Unprotect workbook structure.
    Args:
        password: Password used when workbook was protected.
    """
    return controller.unprotect_workbook_structure(password=password)


@mcp.tool()
def excel_is_workbook_protected() -> Dict[str, Any]:
    """Check if workbook structure or windows are protected."""
    return controller.is_workbook_protected()


# ── PRINT QUALITY & BLACK AND WHITE ─────────────────────────────────────────────

@mcp.tool()
def excel_set_print_quality(dpi: int = 300, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set print resolution quality in DPI (e.g. 300 or 600).
    Args:
        dpi: Print quality resolution (e.g. 300, 600).
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_print_quality(dpi=dpi, sheet_name=sheet_name)


@mcp.tool()
def excel_set_black_and_white_print(bw: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Enable or disable black and white print mode for clean monochrome PDF export or printing.
    Args:
        bw: True for monochrome print, False for color.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_black_and_white_print(bw=bw, sheet_name=sheet_name)


# ── HEADLESS OOXML DATA STRUCTURES & MANIPULATION ────────────────────────────────

@mcp.tool()
def excel_get_defined_names_headless(file_path: str) -> List[Dict[str, Any]]:
    """Read all workbook defined names / named ranges headlessly.
    Args:
        file_path: Path to the .xlsx file.
    """
    return ExcelEngine.get_defined_names_headless(file_path=file_path)


@mcp.tool()
def excel_insert_rows_headless(
    file_path: str,
    row_idx: int,
    amount: int = 1,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Insert blank rows headlessly at a specific row position.
    Args:
        file_path: Path to the .xlsx file.
        row_idx: 1-based row index to insert at.
        amount: Number of rows to insert.
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.insert_rows_headless(
        file_path=file_path,
        row_idx=row_idx,
        amount=amount,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_delete_rows_headless(
    file_path: str,
    row_idx: int,
    amount: int = 1,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Delete rows headlessly at a specific row position.
    Args:
        file_path: Path to the .xlsx file.
        row_idx: 1-based row index to delete.
        amount: Number of rows to delete.
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.delete_rows_headless(
        file_path=file_path,
        row_idx=row_idx,
        amount=amount,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_insert_cols_headless(
    file_path: str,
    col_idx: int,
    amount: int = 1,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Insert blank columns headlessly at a specific column position.
    Args:
        file_path: Path to the .xlsx file.
        col_idx: 1-based column index to insert at.
        amount: Number of columns to insert.
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.insert_cols_headless(
        file_path=file_path,
        col_idx=col_idx,
        amount=amount,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_delete_cols_headless(
    file_path: str,
    col_idx: int,
    amount: int = 1,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Delete columns headlessly at a specific column position.
    Args:
        file_path: Path to the .xlsx file.
        col_idx: 1-based column index to delete.
        amount: Number of columns to delete.
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.delete_cols_headless(
        file_path=file_path,
        col_idx=col_idx,
        amount=amount,
        sheet_name=sheet_name,
        output_path=output_path
    )


# ── SHAPE FORMATTING & MANIPULATION ─────────────────────────────────────────────

@mcp.tool()
def excel_set_shape_fill_color(
    shape_name: str,
    hex_color: str,
    transparency: float = 0.0,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Set solid fill color and transparency for an existing shape.
    Args:
        shape_name: Name of the shape.
        hex_color: Color string (e.g. '#2563EB').
        transparency: Opacity from 0.0 (opaque) to 1.0 (fully transparent).
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_shape_fill_color(
        shape_name=shape_name,
        hex_color=hex_color,
        transparency=transparency,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_shape_line_color(
    shape_name: str,
    hex_color: str,
    weight: float = 1.5,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Set outline border color and thickness for an existing shape.
    Args:
        shape_name: Name of the shape.
        hex_color: Color string (e.g. '#1E293B').
        weight: Line thickness in points.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_shape_line_color(
        shape_name=shape_name,
        hex_color=hex_color,
        weight=weight,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_shape_text(
    shape_name: str,
    text: str,
    font_size: float = 12.0,
    bold: bool = False,
    font_color: str = "#000000",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Set or update text inside an existing shape or callout box.
    Args:
        shape_name: Name of the shape.
        text: Text to display.
        font_size: Font size in points.
        bold: Whether font is bold.
        font_color: Font color hex string.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_shape_text(
        shape_name=shape_name,
        text=text,
        font_size=font_size,
        bold=bold,
        font_color=font_color,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_delete_shape_by_name(shape_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Delete an existing shape by its name identifier.
    Args:
        shape_name: Name of the shape to delete.
        sheet_name: Optional specific worksheet name.
    """
    return controller.delete_shape_by_name(shape_name=shape_name, sheet_name=sheet_name)


@mcp.tool()
def excel_bring_shape_to_front(shape_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Bring a shape to the front of the z-order stacking layer.
    Args:
        shape_name: Name of the shape.
        sheet_name: Optional specific worksheet name.
    """
    return controller.bring_shape_to_front(shape_name=shape_name, sheet_name=sheet_name)


@mcp.tool()
def excel_send_shape_to_back(shape_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Send a shape to the back of the z-order stacking layer.
    Args:
        shape_name: Name of the shape.
        sheet_name: Optional specific worksheet name.
    """
    return controller.send_shape_to_back(shape_name=shape_name, sheet_name=sheet_name)



# ── WORKBOOK WINDOW STATES & TILING ─────────────────────────────────────────────

@mcp.tool()
def excel_minimize_window() -> Dict[str, Any]:
    """Minimize the active Excel desktop window."""
    return controller.minimize_window()


@mcp.tool()
def excel_restore_window() -> Dict[str, Any]:
    """Restore the active Excel window to normal floating dimensions."""
    return controller.restore_window()


@mcp.tool()
def excel_maximize_window() -> Dict[str, Any]:
    """Maximize the active Excel desktop window."""
    return controller.maximize_window()


@mcp.tool()
def excel_tile_windows(tile_style: str = "horizontal") -> Dict[str, Any]:
    """Tile open workbook windows ('horizontal', 'vertical', 'tiled', 'cascade').
    Args:
        tile_style: 'horizontal', 'vertical', 'tiled', or 'cascade'.
    """
    return controller.tile_windows(tile_style=tile_style)


# ── FLOWCHART BLOCKS & CONNECTORS ───────────────────────────────────────────────

@mcp.tool()
def excel_add_connector_arrow(
    start_cell: str,
    end_cell: str,
    color: str = "#2563EB",
    weight: float = 2.0,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Draw an arrow line connecting the center of two cells.
    Args:
        start_cell: Cell where arrow begins (e.g. 'B2').
        end_cell: Cell where arrow points (e.g. 'D2').
        color: Line color hex string.
        weight: Line thickness in points.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_connector_arrow(
        start_cell=start_cell,
        end_cell=end_cell,
        color=color,
        weight=weight,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_add_flowchart_process_block(
    cell: str,
    text: str,
    width: float = 120.0,
    height: float = 50.0,
    bg_color: str = "#EFF6FF",
    border_color: str = "#3B82F6",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert a styled flowchart process rectangle block with centered text.
    Args:
        cell: Anchor cell coordinate (e.g. 'C3').
        text: Block text label.
        width: Box width in points.
        height: Box height in points.
        bg_color: Background hex color.
        border_color: Border line hex color.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_flowchart_process_block(
        cell=cell,
        text=text,
        width=width,
        height=height,
        bg_color=bg_color,
        border_color=border_color,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_add_flowchart_decision_diamond(
    cell: str,
    text: str,
    width: float = 100.0,
    height: float = 75.0,
    bg_color: str = "#FEF3C7",
    border_color: str = "#F59E0B",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Insert a styled flowchart decision diamond with centered text.
    Args:
        cell: Anchor cell coordinate (e.g. 'E3').
        text: Decision question label.
        width: Diamond width in points.
        height: Diamond height in points.
        bg_color: Background hex color.
        border_color: Border line hex color.
        sheet_name: Optional specific worksheet name.
    """
    return controller.add_flowchart_decision_diamond(
        cell=cell,
        text=text,
        width=width,
        height=height,
        bg_color=bg_color,
        border_color=border_color,
        sheet_name=sheet_name
    )


# ── ADVANCED PIVOT TABLE OPERATIONS ─────────────────────────────────────────────

@mcp.tool()
def excel_change_pivot_field_calculation(
    pivot_name: str,
    data_field_name: str,
    calculation: str = "sum",
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Change aggregation calculation on a PivotTable data field ('sum', 'count', 'average', 'max', 'min', 'product').
    Args:
        pivot_name: Name of the PivotTable.
        data_field_name: Name of the data value field.
        calculation: 'sum', 'count', 'average', 'max', 'min', or 'product'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.change_pivot_field_calculation(
        pivot_name=pivot_name,
        data_field_name=data_field_name,
        calculation=calculation,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_set_pivot_layout(pivot_name: str, layout: str = "tabular", sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Set PivotTable report layout ('tabular', 'compact', 'outline').
    Args:
        pivot_name: Name of the PivotTable.
        layout: 'tabular', 'compact', or 'outline'.
        sheet_name: Optional specific worksheet name.
    """
    return controller.set_pivot_layout(pivot_name=pivot_name, layout=layout, sheet_name=sheet_name)


@mcp.tool()
def excel_toggle_pivot_grand_totals(
    pivot_name: str,
    rows: bool = True,
    columns: bool = True,
    sheet_name: Optional[str] = None
) -> Dict[str, Any]:
    """Toggle row and column grand totals for a PivotTable.
    Args:
        pivot_name: Name of the PivotTable.
        rows: Whether to display row grand totals.
        columns: Whether to display column grand totals.
        sheet_name: Optional specific worksheet name.
    """
    return controller.toggle_pivot_grand_totals(
        pivot_name=pivot_name,
        rows=rows,
        columns=columns,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_refresh_pivot_table(pivot_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """Refresh data cache for a specific PivotTable.
    Args:
        pivot_name: Name of the PivotTable to update.
        sheet_name: Optional specific worksheet name.
    """
    return controller.refresh_pivot_table(pivot_name=pivot_name, sheet_name=sheet_name)


# ── HEADLESS OOXML SPREADSHEET ENHANCEMENTS ──────────────────────────────────────

@mcp.tool()
def excel_set_gridlines_headless(
    file_path: str,
    show: bool = True,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Toggle gridlines visibility in an .xlsx file headlessly.
    Args:
        file_path: Path to the .xlsx file.
        show: True to show gridlines, False to hide.
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.set_gridlines_headless(
        file_path=file_path,
        show=show,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_set_zoom_headless(
    file_path: str,
    zoom_scale: int = 100,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Set worksheet zoom scale percentage in an .xlsx file headlessly.
    Args:
        file_path: Path to the .xlsx file.
        zoom_scale: Zoom percentage (e.g. 85, 100, 120).
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.set_zoom_headless(
        file_path=file_path,
        zoom_scale=zoom_scale,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_add_image_headless(
    file_path: str,
    image_path: str,
    cell: str = "A1",
    width: Optional[float] = None,
    height: Optional[float] = None,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Embed an image into an Excel spreadsheet headlessly.
    Args:
        file_path: Path to the .xlsx file.
        image_path: Path to the image file (.png, .jpg).
        cell: Anchor cell coordinate (e.g. 'B2').
        width: Optional display width.
        height: Optional display height.
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.add_image_headless(
        file_path=file_path,
        image_path=image_path,
        cell=cell,
        width=width,
        height=height,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_auto_fit_columns_headless(
    file_path: str,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Compute and auto-fit column widths to text content in an .xlsx file headlessly.
    Args:
        file_path: Path to the .xlsx file.
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.auto_fit_columns_headless(
        file_path=file_path,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_freeze_panes_headless(
    file_path: str,
    cell: str = "A2",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Freeze panes at a specific cell coordinate in an .xlsx file headlessly.
    Args:
        file_path: Path to the .xlsx file.
        cell: Cell coordinate below/right of which to freeze (e.g. 'A2' for top row).
        sheet_name: Optional specific worksheet name.
        output_path: Optional destination path (overwrites input if not provided).
    """
    return ExcelEngine.freeze_panes_headless(
        file_path=file_path,
        cell=cell,
        sheet_name=sheet_name,
        output_path=output_path
    )




@mcp.tool()
def excel_clean_whitespace_headless(
    file_path: str,
    range_address: Optional[str] = None,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Strip leading/trailing whitespace, collapse internal multiple spaces, and remove control characters from cells headlessly."""
    return ExcelEngine.clean_whitespace_headless(
        file_path=file_path,
        range_address=range_address,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_case_transform_headless(
    file_path: str,
    range_address: str,
    mode: str = "upper",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Transform text case (upper, lower, title, capitalize) across a range headlessly."""
    return ExcelEngine.case_transform_headless(
        file_path=file_path,
        range_address=range_address,
        mode=mode,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_regex_replace_headless(
    file_path: str,
    range_address: str,
    pattern: str,
    replacement: str,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Find and replace text using regular expressions across a cell range headlessly."""
    return ExcelEngine.regex_replace_headless(
        file_path=file_path,
        range_address=range_address,
        pattern=pattern,
        replacement=replacement,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_extract_by_pattern_headless(
    file_path: str,
    source_col: str,
    target_col: str,
    pattern: str,
    start_row: int = 2,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Extract matching regex patterns (or group 1) from source column into target column headlessly."""
    return ExcelEngine.extract_by_pattern_headless(
        file_path=file_path,
        source_col=source_col,
        target_col=target_col,
        pattern=pattern,
        start_row=start_row,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_fuzzy_deduplicate_headless(
    file_path: str,
    key_column: str = "A",
    threshold: float = 0.85,
    action: str = "flag",
    start_row: int = 2,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> dict:
    """Detect approximate duplicate rows based on string similarity in a key column headlessly."""
    return ExcelEngine.fuzzy_deduplicate_headless(
        file_path=file_path,
        key_column=key_column,
        threshold=threshold,
        action=action,
        start_row=start_row,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_mask_sensitive_data_headless(
    file_path: str,
    range_address: str,
    mask_type: str = "email",
    mask_char: str = "*",
    visible_chars: int = 4,
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Mask sensitive data like emails, credit cards, or phone numbers in a cell range headlessly."""
    return ExcelEngine.mask_sensitive_data_headless(
        file_path=file_path,
        range_address=range_address,
        mask_type=mask_type,
        mask_char=mask_char,
        visible_chars=visible_chars,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_detect_column_types_headless(
    file_path: str,
    sheet_name: Optional[str] = None,
    sample_rows: int = 100
) -> dict:
    """Infer semantic data types (integer, float, date, boolean, email, text) across columns headlessly."""
    return ExcelEngine.detect_column_types_headless(
        file_path=file_path,
        sheet_name=sheet_name,
        sample_rows=sample_rows
    )




@mcp.tool()
def excel_add_amortization_schedule_headless(
    file_path: str,
    principal: float,
    annual_interest_rate: float,
    term_years: int,
    start_month: str = "2026-01",
    sheet_name: str = "Amortization",
    output_path: Optional[str] = None
) -> str:
    """Generate a complete loan amortization schedule table headlessly."""
    return ExcelEngine.add_amortization_schedule_headless(
        file_path=file_path,
        principal=principal,
        annual_interest_rate=annual_interest_rate,
        term_years=term_years,
        start_month=start_month,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_add_cagr_analysis_headless(
    file_path: str,
    start_value: float,
    end_value: float,
    periods: int,
    output_cell: str = "B2",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> dict:
    """Calculate Compound Annual Growth Rate (CAGR) and build formatted summary card headlessly."""
    return ExcelEngine.add_cagr_analysis_headless(
        file_path=file_path,
        start_value=start_value,
        end_value=end_value,
        periods=periods,
        output_cell=output_cell,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_add_depreciation_schedule_headless(
    file_path: str,
    asset_cost: float,
    salvage_value: float,
    useful_life_years: int,
    method: str = "straight_line",
    sheet_name: str = "Depreciation",
    output_path: Optional[str] = None
) -> str:
    """Generate a complete multi-year asset depreciation schedule table headlessly."""
    return ExcelEngine.add_depreciation_schedule_headless(
        file_path=file_path,
        asset_cost=asset_cost,
        salvage_value=salvage_value,
        useful_life_years=useful_life_years,
        method=method,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_generate_correlation_matrix_headless(
    file_path: str,
    columns: List[str],
    start_row: int = 2,
    end_row: Optional[int] = None,
    output_cell: str = "G2",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> dict:
    """Compute Pearson correlation matrix across numerical columns headlessly."""
    return ExcelEngine.generate_correlation_matrix_headless(
        file_path=file_path,
        columns=columns,
        start_row=start_row,
        end_row=end_row,
        output_cell=output_cell,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_add_summary_statistics_headless(
    file_path: str,
    column: str = "B",
    start_row: int = 2,
    end_row: Optional[int] = None,
    output_cell: str = "D2",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> dict:
    """Compute comprehensive statistical metrics (Count, Mean, Median, Min, Max, StdDev) headlessly."""
    return ExcelEngine.add_summary_statistics_headless(
        file_path=file_path,
        column=column,
        start_row=start_row,
        end_row=end_row,
        output_cell=output_cell,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_calculate_break_even_headless(
    file_path: str,
    fixed_costs: float,
    variable_cost_per_unit: float,
    price_per_unit: float,
    target_profit: float = 0.0,
    output_cell: str = "B2",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> dict:
    """Compute Break-Even units and revenue, contribution margin, and formatted card headlessly."""
    return ExcelEngine.calculate_break_even_headless(
        file_path=file_path,
        fixed_costs=fixed_costs,
        variable_cost_per_unit=variable_cost_per_unit,
        price_per_unit=price_per_unit,
        target_profit=target_profit,
        output_cell=output_cell,
        sheet_name=sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_add_waterfall_breakdown_headless(
    file_path: str,
    categories: List[str],
    amounts: List[float],
    output_cell: str = "B2",
    sheet_name: Optional[str] = None,
    output_path: Optional[str] = None
) -> str:
    """Create structured waterfall breakdown table (Base, Increase, Decrease, Cumulative) headlessly."""
    return ExcelEngine.add_waterfall_breakdown_headless(
        file_path=file_path,
        categories=categories,
        amounts=amounts,
        output_cell=output_cell,
        sheet_name=sheet_name,
        output_path=output_path
    )




@mcp.tool()
def excel_append_audit_log_headless(
    file_path: str,
    event_message: str,
    user_name: str = "AI Agent",
    log_sheet_name: str = "_AuditLog",
    output_path: Optional[str] = None
) -> str:
    """Append an audit trail event log entry to a workbook headlessly."""
    return ExcelEngine.append_audit_log_headless(
        file_path=file_path,
        event_message=event_message,
        user_name=user_name,
        log_sheet_name=log_sheet_name,
        output_path=output_path
    )


@mcp.tool()
def excel_batch_export_sheets_to_csv_headless(
    file_path: str,
    output_dir: Optional[str] = None,
    delimiter: str = ","
) -> List[str]:
    """Export each sheet in an Excel workbook to individual CSV files headlessly."""
    return ExcelEngine.batch_export_sheets_to_csv_headless(
        file_path=file_path,
        output_dir=output_dir,
        delimiter=delimiter
    )


@mcp.tool()
def excel_merge_multiple_workbooks_headless(
    file_paths: List[str],
    output_path: str,
    combine_mode: str = "sheets"
) -> str:
    """Merge multiple Excel workbooks into one (by adding sheets or stacking rows) headlessly."""
    return ExcelEngine.merge_multiple_workbooks_headless(
        file_paths=file_paths,
        output_path=output_path,
        combine_mode=combine_mode
    )


@mcp.tool()
def excel_split_workbook_by_column_headless(
    file_path: str,
    split_column: str = "A",
    output_dir: Optional[str] = None,
    sheet_name: Optional[str] = None
) -> List[str]:
    """Split a worksheet into multiple Excel files partitioned by unique values in a column headlessly."""
    return ExcelEngine.split_workbook_by_column_headless(
        file_path=file_path,
        split_column=split_column,
        output_dir=output_dir,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_validate_schema_headless(
    file_path: str,
    expected_columns: List[str],
    required_types: Optional[Dict[str, str]] = None,
    sheet_name: Optional[str] = None
) -> dict:
    """Validate worksheet headers and cell types against an expected schema headlessly."""
    return ExcelEngine.validate_schema_headless(
        file_path=file_path,
        expected_columns=expected_columns,
        required_types=required_types,
        sheet_name=sheet_name
    )


@mcp.tool()
def excel_generate_markdown_data_dictionary_headless(
    file_path: str,
    sheet_name: Optional[str] = None
) -> str:
    """Generate a complete Markdown data dictionary catalog from an Excel worksheet headlessly."""
    return ExcelEngine.generate_markdown_data_dictionary_headless(
        file_path=file_path,
        sheet_name=sheet_name
    )


def main():
    """Run the FastMCP server."""
    mcp.run()


if __name__ == "__main__":
    main()




