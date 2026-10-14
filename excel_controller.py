"""Microsoft Excel COM Automation Controller for Windows.
Provides real-time, live visual control of Microsoft Excel desktop via Windows COM.
"""

import os
import sys
import time
from typing import Any, Dict, List, Optional, Tuple, Union

if sys.platform == "win32":
    import pythoncom
    import win32com.client
    import win32gui
else:
    win32com = None
    win32gui = None

# Excel Chart Type Constants
XL_CHART_TYPES = {
    "column": 51,        # xlColumnClustered
    "column_stacked": 52,# xlColumnStacked
    "bar": 57,           # xlBarClustered
    "bar_stacked": 58,   # xlBarStacked
    "line": 4,           # xlLine
    "line_markers": 65,  # xlLineMarkers
    "pie": 5,            # xlPie
    "doughnut": -4120,   # xlDoughnut
    "area": 1,           # xlArea
    "scatter": -4169,    # xlXYScatter
}

# Alignment Constants
XL_H_ALIGN = {
    "left": -4131,       # xlLeft
    "center": -4108,     # xlCenter
    "right": -4152,      # xlRight
    "justify": -4130,    # xlJustify
}

XL_V_ALIGN = {
    "top": -4160,        # xlTop
    "center": -4108,     # xlCenter
    "bottom": -4107,     # xlBottom
}

# Border Constants
XL_BORDER_STYLES = {
    "thin": (1, 2),      # xlContinuous, xlThin
    "medium": (1, -4138),# xlContinuous, xlMedium
    "thick": (1, 4),     # xlContinuous, xlThick
    "double": (-4119, 2),# xlDouble, xlThin
    "none": (-4142, 2),  # xlLineStyleNone
}


def hex_to_bgr_int(hex_str: str) -> int:
    """Converts a hex color string (#RRGGBB or #RGB) to COM BGR integer."""
    hex_str = hex_str.strip().lstrip("#")
    if len(hex_str) == 3:
        hex_str = "".join([c * 2 for c in hex_str])
    if len(hex_str) != 6:
        return 0
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    # Excel COM color is BGR integer: R + (G * 256) + (B * 65536)
    return r + (g * 256) + (b * 65536)


class ExcelController:
    """Manages live visual interaction with Microsoft Excel desktop on Windows."""

    def __init__(self):
        self._app = None
        self._wb = None

    def get_app(self, bring_to_front: bool = True):
        """Connect to running Excel or launch a new visible instance."""
        if sys.platform != "win32":
            raise RuntimeError("Excel COM controller is only supported on Windows.")

        pythoncom.CoInitialize()

        try:
            self._app = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            self._app = win32com.client.Dispatch("Excel.Application")

        self._app.Visible = True
        self._app.DisplayAlerts = False

        try:
            self._app.WindowState = -4137  # xlMaximized
        except Exception:
            pass

        if bring_to_front:
            self.focus_window()

        return self._app

    def focus_window(self):
        """Bring Microsoft Excel window to foreground."""
        try:
            if win32gui:
                hwnd = win32gui.FindWindow("XLMAIN", None)
                if hwnd:
                    win32gui.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                    win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def get_workbook(self, create_if_none: bool = True):
        """Get the current active workbook, or create one if none is open."""
        app = self.get_app()
        try:
            if app.Workbooks.Count > 0:
                self._wb = app.ActiveWorkbook
                return self._wb
        except Exception:
            pass

        if create_if_none:
            self._wb = app.Workbooks.Add()
            return self._wb
        raise RuntimeError("No workbook is currently open in Microsoft Excel.")

    def get_sheet(self, sheet_name: Optional[str] = None):
        """Get a specific worksheet by name or the active sheet."""
        wb = self.get_workbook()
        if sheet_name:
            try:
                return wb.Worksheets(sheet_name)
            except Exception:
                raise ValueError(f"Sheet '{sheet_name}' not found in active workbook.")
        return wb.ActiveSheet

    # ── SYSTEM / WORKBOOK MANAGEMENT ──────────────────────────────────────────

    def get_info(self) -> Dict[str, Any]:
        """Return running status and active workbook information."""
        try:
            app = self.get_app(bring_to_front=False)
            wb_count = app.Workbooks.Count
            if wb_count == 0:
                return {
                    "running": True,
                    "workbooks_open": 0,
                    "active_workbook": None,
                    "sheets": [],
                    "status": "Excel is running with no open workbooks."
                }

            wb = app.ActiveWorkbook
            sheet = wb.ActiveSheet
            sheets = [s.Name for s in wb.Worksheets]

            used_range = ""
            try:
                ur = sheet.UsedRange
                used_range = ur.Address
            except Exception:
                pass

            return {
                "running": True,
                "workbooks_open": wb_count,
                "active_workbook": wb.Name,
                "active_sheet": sheet.Name,
                "sheets": sheets,
                "active_cell": app.ActiveCell.Address if app.ActiveCell else None,
                "used_range": used_range,
                "status": "Connected to active Excel session."
            }
        except Exception as e:
            return {
                "running": False,
                "error": str(e),
                "status": "Microsoft Excel is not running or unreachable."
            }

    def new_workbook(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Create a new visible workbook."""
        app = self.get_app(bring_to_front=True)
        wb = app.Workbooks.Add()
        self._wb = wb
        if sheet_name:
            wb.ActiveSheet.Name = sheet_name[:31]
        return {
            "status": "success",
            "message": f"Created new workbook '{wb.Name}'",
            "workbook": wb.Name,
            "active_sheet": wb.ActiveSheet.Name
        }

    def open_workbook(self, file_path: str) -> Dict[str, Any]:
        """Open an existing workbook (.xlsx, .xls, .xlsm, .csv)."""
        app = self.get_app(bring_to_front=True)
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")
        wb = app.Workbooks.Open(abs_path)
        self._wb = wb
        return {
            "status": "success",
            "message": f"Opened workbook '{wb.Name}'",
            "path": abs_path,
            "sheets": [s.Name for s in wb.Worksheets]
        }

    def save_workbook(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Save the active workbook."""
        wb = self.get_workbook()
        app = self.get_app(bring_to_front=False)
        app.DisplayAlerts = False
        try:
            if file_path:
                abs_path = os.path.abspath(file_path)
                os.makedirs(os.path.dirname(abs_path), exist_ok=True)
                if os.path.exists(abs_path):
                    try:
                        os.remove(abs_path)
                    except Exception:
                        pass
                ext = os.path.splitext(abs_path)[1].lower()
                fmt = 51 if ext == ".xlsx" else (52 if ext == ".xlsm" else (6 if ext == ".csv" else None))
                if fmt is not None:
                    wb.SaveAs(abs_path, fmt)
                else:
                    wb.SaveAs(abs_path)
                saved_to = abs_path
            else:
                wb.Save()
                saved_to = wb.FullName
        finally:
            app.DisplayAlerts = True

        return {
            "status": "success",
            "message": f"Workbook saved successfully to '{saved_to}'",
            "path": saved_to
        }

    def export_pdf(self, output_path: Optional[str] = None, sheet_only: bool = False) -> Dict[str, Any]:
        """Export the active workbook or current sheet directly to PDF."""
        wb = self.get_workbook()
        if not output_path:
            base = os.path.splitext(wb.FullName)[0] if wb.Path else os.path.join(os.getcwd(), wb.Name)
            output_path = base + ".pdf"
        abs_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)

        target = wb.ActiveSheet if sheet_only else wb
        # 0 = xlTypePDF
        target.ExportAsFixedFormat(0, abs_path)
        return {
            "status": "success",
            "message": f"Exported PDF successfully to '{abs_path}'",
            "output_path": abs_path
        }

    def close_workbook(self, save_changes: bool = False) -> Dict[str, Any]:
        """Close the active workbook."""
        wb = self.get_workbook(create_if_none=False)
        name = wb.Name
        wb.Close(SaveChanges=save_changes)
        return {"status": "success", "message": f"Closed workbook '{name}'."}

    # ── SHEET OPERATIONS ──────────────────────────────────────────────────────

    def list_sheets(self) -> List[Dict[str, Any]]:
        """List all sheets with their index and visibility."""
        wb = self.get_workbook()
        result = []
        for i, s in enumerate(wb.Worksheets, start=1):
            result.append({
                "index": i,
                "name": s.Name,
                "visible": s.Visible == -1,
                "is_active": s.Name == wb.ActiveSheet.Name
            })
        return result

    def add_sheet(self, name: str, after: Optional[str] = None) -> Dict[str, Any]:
        """Add a new worksheet with the specified name."""
        wb = self.get_workbook()
        name_clean = name[:31]
        if after:
            try:
                target = wb.Worksheets(after)
                sheet = wb.Worksheets.Add(After=target)
            except Exception:
                sheet = wb.Worksheets.Add()
        else:
            sheet = wb.Worksheets.Add()
        sheet.Name = name_clean
        sheet.Activate()
        return {
            "status": "success",
            "message": f"Added sheet '{name_clean}'",
            "name": name_clean
        }

    def activate_sheet(self, name: str) -> Dict[str, Any]:
        """Activate (switch to) a worksheet by name."""
        sheet = self.get_sheet(name)
        sheet.Activate()
        return {"status": "success", "message": f"Activated sheet '{sheet.Name}'."}

    def rename_sheet(self, old_name: str, new_name: str) -> Dict[str, Any]:
        """Rename an existing worksheet."""
        sheet = self.get_sheet(old_name)
        sheet.Name = new_name[:31]
        return {"status": "success", "message": f"Renamed sheet '{old_name}' to '{new_name[:31]}'"}

    def delete_sheet(self, name: str) -> Dict[str, Any]:
        """Delete a worksheet by name."""
        wb = self.get_workbook()
        app = self.get_app()
        app.DisplayAlerts = False
        try:
            sheet = wb.Worksheets(name)
            sheet.Delete()
            return {"status": "success", "message": f"Deleted sheet '{name}'"}
        finally:
            app.DisplayAlerts = True

    # ── CELL & RANGE OPERATIONS ───────────────────────────────────────────────

    def read_range(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Read data from a cell or range of cells."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        raw_val = rng.Value

        # Normalize raw_val into a 2D list
        if not isinstance(raw_val, (list, tuple)):
            values = [[raw_val]]
        elif isinstance(raw_val, tuple) and not isinstance(raw_val[0], tuple):
            values = [[v for v in raw_val]]
        else:
            values = [[c for c in row] for row in raw_val]

        return {
            "address": rng.Address,
            "rows": len(values),
            "columns": len(values[0]) if values else 0,
            "values": values,
            "formula": str(rng.Formula) if rng.Count == 1 else None
        }

    def write_range(self, start_cell: str, data: List[List[Any]], sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Write a 2D array of data starting from start_cell."""
        if not data or not data[0]:
            return {"status": "noop", "message": "No data provided"}

        sheet = self.get_sheet(sheet_name)
        start_rng = sheet.Range(start_cell)
        num_rows = len(data)
        num_cols = max(len(row) for row in data)

        # Pad rows to uniform columns length
        clean_data = []
        for row in data:
            row_list = list(row)
            if len(row_list) < num_cols:
                row_list += [None] * (num_cols - len(row_list))
            clean_data.append(tuple(row_list))

        end_cell = sheet.Cells(start_rng.Row + num_rows - 1, start_rng.Column + num_cols - 1)
        target_rng = sheet.Range(start_rng, end_cell)
        target_rng.Value = tuple(clean_data)

        return {
            "status": "success",
            "range": target_rng.Address,
            "rows_written": num_rows,
            "cols_written": num_cols
        }

    def write_cell(self, cell: str, value: Any, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Write a single value or formula to a specific cell."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(cell)
        str_val = str(value)
        if str_val.startswith("="):
            rng.Formula = str_val
        else:
            rng.Value = value
        return {
            "status": "success",
            "cell": cell,
            "value": value
        }

    def set_formula(self, cell: str, formula: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Set an Excel formula in a cell or range (e.g., '=SUM(B2:B10)')."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(cell)
        if not formula.startswith("="):
            formula = "=" + formula
        rng.Formula = formula
        return {
            "status": "success",
            "cell": cell,
            "formula": formula,
            "calculated_value": rng.Value
        }

    def clear_range(self, range_address: str, clear_type: str = "all", sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Clear a range: 'all', 'contents', or 'formats'."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        if clear_type == "contents":
            rng.ClearContents()
        elif clear_type == "formats":
            rng.ClearFormats()
        else:
            rng.Clear()
        return {"status": "success", "cleared": range_address, "type": clear_type}

    def find_replace(self, find_text: str, replace_text: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Find and replace text across the active sheet."""
        sheet = self.get_sheet(sheet_name)
        # xlPart = 2, xlByRows = 1
        sheet.Cells.Replace(What=find_text, Replacement=replace_text, LookAt=2, SearchOrder=1, MatchCase=False)
        return {
            "status": "success",
            "find": find_text,
            "replace": replace_text,
            "message": f"Replaced all occurrences of '{find_text}' with '{replace_text}'"
        }

    # ── FORMATTING & STYLING ──────────────────────────────────────────────────

    def format_range(
        self,
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
        """Format cells: font, background color, alignments, borders, and number formats."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)

        if bold is not None:
            rng.Font.Bold = bold
        if italic is not None:
            rng.Font.Italic = italic
        if font_name is not None:
            rng.Font.Name = font_name
        if font_size is not None:
            rng.Font.Size = font_size
        if font_color is not None:
            rng.Font.Color = hex_to_bgr_int(font_color)
        if bg_color is not None:
            rng.Interior.Color = hex_to_bgr_int(bg_color)
        if horizontal_align in XL_H_ALIGN:
            rng.HorizontalAlignment = XL_H_ALIGN[horizontal_align]
        if vertical_align in XL_V_ALIGN:
            rng.VerticalAlignment = XL_V_ALIGN[vertical_align]
        if wrap_text is not None:
            rng.WrapText = wrap_text
        if number_format is not None:
            rng.NumberFormat = number_format

        if border_style and border_style in XL_BORDER_STYLES:
            line_style, weight = XL_BORDER_STYLES[border_style]
            # Apply to borders
            for border_id in [7, 8, 9, 10, 11, 12]:  # left, top, bottom, right, inside_v, inside_h
                try:
                    b = rng.Borders(border_id)
                    b.LineStyle = line_style
                    b.Weight = weight
                    if border_color:
                        b.Color = hex_to_bgr_int(border_color)
                except Exception:
                    pass

        return {
            "status": "success",
            "range": range_address,
            "message": "Formatting applied successfully."
        }

    def autofit(self, columns: Optional[str] = None, rows: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """AutoFit column widths and row heights."""
        sheet = self.get_sheet(sheet_name)
        if columns:
            sheet.Columns(columns).AutoFit()
        else:
            sheet.Columns.AutoFit()
        if rows:
            sheet.Rows(rows).AutoFit()
        return {"status": "success", "message": "Columns/rows auto-fitted."}

    def freeze_panes(self, row: int = 2, col: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Freeze panes at the specified row and column (e.g. freeze header row = row 2, col 1)."""
        app = self.get_app()
        sheet = self.get_sheet(sheet_name)
        sheet.Activate()
        app.ActiveWindow.FreezePanes = False
        target = sheet.Cells(row, col)
        target.Select()
        app.ActiveWindow.FreezePanes = True
        return {"status": "success", "message": f"Panes frozen at row {row}, column {col}."}

    def unfreeze_panes(self) -> Dict[str, Any]:
        """Unfreeze panes."""
        app = self.get_app()
        app.ActiveWindow.FreezePanes = False
        return {"status": "success", "message": "Panes unfrozen."}

    # ── TABLES & STRUCTURED DATA ──────────────────────────────────────────────

    def create_table(
        self,
        range_address: str,
        table_name: Optional[str] = None,
        style_name: str = "TableStyleMedium9",
        has_headers: bool = True,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Convert a range of cells into an official Excel Table (ListObject)."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        # xlSrcRange = 1, xlYes = 1, xlNo = 2
        header_const = 1 if has_headers else 2
        table = sheet.ListObjects.Add(1, rng, None, header_const)
        if table_name:
            table.Name = table_name
        table.TableStyle = style_name
        return {
            "status": "success",
            "table_name": table.Name,
            "range": rng.Address,
            "style": style_name
        }

    def add_table_total_row(self, table_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Display the total summary row for an Excel Table."""
        sheet = self.get_sheet(sheet_name)
        tbl = sheet.ListObjects(table_name)
        tbl.ShowTotals = True
        return {"status": "success", "table": table_name, "totals_shown": True}

    # ── CHARTS & VISUALIZATIONS ───────────────────────────────────────────────

    def add_chart(
        self,
        chart_type: str,
        source_range: str,
        title: Optional[str] = None,
        left: float = 350,
        top: float = 40,
        width: float = 480,
        height: float = 300,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Insert a native Excel chart (column, bar, line, pie, doughnut, area, scatter)."""
        sheet = self.get_sheet(sheet_name)
        src_rng = sheet.Range(source_range)

        chart_code = XL_CHART_TYPES.get(chart_type.lower(), 51)  # default column clustered
        chart_obj = sheet.ChartObjects().Add(left, top, width, height)
        chart = chart_obj.Chart
        chart.SetSourceData(src_rng)
        chart.ChartType = chart_code

        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = title

        return {
            "status": "success",
            "chart_name": chart_obj.Name,
            "chart_type": chart_type,
            "source_range": source_range,
            "title": title
        }

    def export_chart_image(self, chart_index: int, output_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Export a chart as a high-resolution PNG image."""
        sheet = self.get_sheet(sheet_name)
        abs_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        chart_obj = sheet.ChartObjects(chart_index)
        chart_obj.Chart.Export(abs_path, "PNG")
        return {
            "status": "success",
            "message": f"Exported chart image to '{abs_path}'",
            "output_path": abs_path
        }

    # ── HIGH-LEVEL EXECUTIVE DASHBOARD BLOCKS ──────────────────────────────────

    def create_kpi_card(
        self,
        start_cell: str,
        title: str,
        value: str,
        subtitle: Optional[str] = None,
        accent_color: str = "#2563EB",
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create a modern executive KPI metric block in a 3-row, 2-column cell box."""
        sheet = self.get_sheet(sheet_name)
        top_left = sheet.Range(start_cell)
        r = top_left.Row
        c = top_left.Column

        # Define 3x2 card area
        card_rng = sheet.Range(sheet.Cells(r, c), sheet.Cells(r + 2, c + 2))
        card_rng.Interior.Color = hex_to_bgr_int("#F8FAFC")  # Light gray-slate background

        # Border
        for b_id in [7, 8, 9, 10]:
            try:
                b = card_rng.Borders(b_id)
                b.LineStyle = 1
                b.Weight = 2
                b.Color = hex_to_bgr_int("#E2E8F0")
            except Exception:
                pass

        # Title row
        title_cell = sheet.Cells(r, c)
        title_cell.Value = title.upper()
        title_cell.Font.Bold = True
        title_cell.Font.Size = 9
        title_cell.Font.Color = hex_to_bgr_int("#64748B")

        # Value row
        val_cell = sheet.Cells(r + 1, c)
        val_cell.Value = value
        val_cell.Font.Bold = True
        val_cell.Font.Size = 18
        val_cell.Font.Color = hex_to_bgr_int(accent_color)

        # Subtitle row
        if subtitle:
            sub_cell = sheet.Cells(r + 2, c)
            sub_cell.Value = subtitle
            sub_cell.Font.Italic = True
            sub_cell.Font.Size = 8
            sub_cell.Font.Color = hex_to_bgr_int("#059669" if "+" in subtitle else "#475569")

        return {
            "status": "success",
            "card_range": card_rng.Address,
            "title": title,
            "value": value
        }

    # ── SORTING & FILTERING ───────────────────────────────────────────────────

    def sort_range(
        self,
        range_address: str,
        key_column: int = 1,
        order: str = "ascending",
        has_headers: bool = True,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Sort data in range by a key column."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        # xlAscending = 1, xlDescending = 2; xlYes = 1, xlNo = 2
        order_const = 1 if order.lower() == "ascending" else 2
        header_const = 1 if has_headers else 2
        key_rng = rng.Columns(key_column)
        rng.Sort(Key1=key_rng, Order1=order_const, Header=header_const)
        return {
            "status": "success",
            "range": range_address,
            "key_column": key_column,
            "order": order
        }

    def apply_filter(
        self,
        range_address: str,
        column_index: int = 1,
        criteria: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Apply an AutoFilter to range on a specific column index."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        if criteria:
            rng.AutoFilter(Field=column_index, Criteria1=str(criteria))
        else:
            rng.AutoFilter(Field=column_index)
        return {"status": "success", "range": range_address, "column": column_index, "criteria": criteria}

    def clear_filters(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Clear all active AutoFilters on the sheet."""
        sheet = self.get_sheet(sheet_name)
        try:
            if sheet.AutoFilterMode:
                sheet.AutoFilterMode = False
        except Exception:
            pass
        return {"status": "success", "message": "All filters cleared."}

    # ── CONDITIONAL FORMATTING ────────────────────────────────────────────────

    def add_conditional_formatting(
        self,
        range_address: str,
        rule_type: str = "color_scale",
        operator: str = "greater",
        value: Optional[str] = None,
        bg_color: Optional[str] = None,
        font_color: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add conditional formatting: 'color_scale', 'data_bar', or 'cell_value' rule."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)

        if rule_type == "color_scale":
            cs = rng.FormatConditions.AddColorScale(3)
            # Default 3-color scale (Green - Yellow - Red)
        elif rule_type == "data_bar":
            db = rng.FormatConditions.AddDatabar()
            if bg_color:
                db.BarColor.Color = hex_to_bgr_int(bg_color)
        else:
            # Cell value rule: xlCellValue = 1
            op_map = {
                "greater": 3,     # xlGreater
                "less": 4,        # xlLess
                "equal": 7,       # xlEqual
                "between": 5,     # xlBetween
                "not_equal": 8,   # xlNotEqual
            }
            op_const = op_map.get(operator.lower(), 3)
            fc = rng.FormatConditions.Add(1, op_const, str(value or "0"))
            if bg_color:
                fc.Interior.Color = hex_to_bgr_int(bg_color)
            if font_color:
                fc.Font.Color = hex_to_bgr_int(font_color)

        return {
            "status": "success",
            "range": range_address,
            "rule_type": rule_type,
            "message": f"Applied {rule_type} conditional formatting."
        }

    def clear_conditional_formatting(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Remove all conditional formatting rules from range."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        rng.FormatConditions.Delete()
        return {"status": "success", "range": range_address, "message": "Conditional formatting cleared."}

    # ── DATA VALIDATION ───────────────────────────────────────────────────────

    def add_data_validation(
        self,
        range_address: str,
        validation_type: str = "list",
        formula1: str = "Yes,No,Pending",
        error_title: Optional[str] = None,
        error_message: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add interactive dropdown list or value validation to a range of cells."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        rng.Validation.Delete()

        # xlValidateList = 3, xlValidAlertStop = 1
        rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula1)
        rng.Validation.IgnoreBlank = True
        rng.Validation.InCellDropdown = True
        if error_title:
            rng.Validation.ErrorTitle = error_title
        if error_message:
            rng.Validation.ErrorMessage = error_message

        return {
            "status": "success",
            "range": range_address,
            "validation_type": validation_type,
            "formula": formula1
        }

    def clear_data_validation(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Remove data validation dropdowns/rules from range."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        rng.Validation.Delete()
        return {"status": "success", "range": range_address, "message": "Data validation removed."}

    # ── STRUCTURAL ROWS & COLUMNS ─────────────────────────────────────────────

    def insert_rows(self, row_index: int, count: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Insert blank rows before row_index."""
        sheet = self.get_sheet(sheet_name)
        for _ in range(count):
            sheet.Rows(row_index).Insert()
        return {"status": "success", "inserted_at": row_index, "count": count}

    def delete_rows(self, row_index: int, count: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Delete rows starting from row_index."""
        sheet = self.get_sheet(sheet_name)
        for _ in range(count):
            sheet.Rows(row_index).Delete()
        return {"status": "success", "deleted_at": row_index, "count": count}

    def insert_columns(self, column: Union[str, int], count: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Insert blank columns before the specified column letter or index."""
        sheet = self.get_sheet(sheet_name)
        for _ in range(count):
            sheet.Columns(column).Insert()
        return {"status": "success", "inserted_at": str(column), "count": count}

    def delete_columns(self, column: Union[str, int], count: int = 1, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Delete columns starting from column letter or index."""
        sheet = self.get_sheet(sheet_name)
        for _ in range(count):
            sheet.Columns(column).Delete()
        return {"status": "success", "deleted_at": str(column), "count": count}

    def hide_unhide_columns(self, columns: str, hide: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Hide or unhide specific columns (e.g. 'C:E')."""
        sheet = self.get_sheet(sheet_name)
        sheet.Columns(columns).Hidden = hide
        return {"status": "success", "columns": columns, "hidden": hide}

    def hide_unhide_rows(self, rows: str, hide: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Hide or unhide specific rows (e.g. '10:15')."""
        sheet = self.get_sheet(sheet_name)
        sheet.Rows(rows).Hidden = hide
        return {"status": "success", "rows": rows, "hidden": hide}

    def set_column_width(self, columns: str, width: float, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Set explicit column width (e.g. columns='B:D', width=20)."""
        sheet = self.get_sheet(sheet_name)
        sheet.Columns(columns).ColumnWidth = width
        return {"status": "success", "columns": columns, "width": width}

    def set_row_height(self, rows: str, height: float, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Set explicit row height in points (e.g. rows='1:3', height=25)."""
        sheet = self.get_sheet(sheet_name)
        sheet.Rows(rows).RowHeight = height
        return {"status": "success", "rows": rows, "height": height}

    def merge_cells(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Merge cells in range."""
        sheet = self.get_sheet(sheet_name)
        app = self.get_app(bring_to_front=False)
        app.DisplayAlerts = False
        try:
            sheet.Range(range_address).Merge()
        finally:
            app.DisplayAlerts = True
        return {"status": "success", "merged": range_address}

    def unmerge_cells(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Unmerge cells in range."""
        sheet = self.get_sheet(sheet_name)
        sheet.Range(range_address).UnMerge()
        return {"status": "success", "unmerged": range_address}

    # ── IMAGES & VISUAL EXPORT ────────────────────────────────────────────────

    def add_image(
        self,
        image_path: str,
        cell: str = "B2",
        width: Optional[float] = None,
        height: Optional[float] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Insert an image file onto a worksheet anchored at a cell."""
        sheet = self.get_sheet(sheet_name)
        abs_path = os.path.abspath(image_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Image not found: {abs_path}")

        target_cell = sheet.Range(cell)
        pic = sheet.Pictures().Insert(abs_path)
        pic.Left = target_cell.Left
        pic.Top = target_cell.Top
        if width:
            pic.Width = width
        if height:
            pic.Height = height

        return {
            "status": "success",
            "image": abs_path,
            "cell": cell,
            "width": pic.Width,
            "height": pic.Height
        }

    def export_range_image(
        self,
        range_address: str,
        output_path: str,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Take a high-resolution snapshot of a range and save as PNG image."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        abs_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)

        # 1 = xlScreen, 2 = xlBitmap
        rng.CopyPicture(1, 2)
        chart_obj = sheet.ChartObjects().Add(0, 0, rng.Width, rng.Height)
        chart = chart_obj.Chart
        chart.Paste()
        chart.Export(abs_path, "PNG")
        chart_obj.Delete()

        return {
            "status": "success",
            "range": range_address,
            "output_path": abs_path
        }

    # ── COMMENTS & NOTES ──────────────────────────────────────────────────────

    def add_comment(self, cell: str, text: str, author: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Add a note or comment to a specific cell."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(cell)
        try:
            rng.ClearComments()
        except Exception:
            pass
        prefix = f"{author}: " if author else ""
        rng.AddComment(prefix + text)
        return {"status": "success", "cell": cell, "comment": text}

    def list_comments(self, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """List all cell comments across the active sheet."""
        sheet = self.get_sheet(sheet_name)
        comments = []
        try:
            for c in sheet.Comments:
                comments.append({
                    "cell": c.Parent.Address,
                    "author": c.Author,
                    "text": c.Text()
                })
        except Exception:
            pass
        return comments

    def delete_comment(self, cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Delete comment on a specific cell."""
        sheet = self.get_sheet(sheet_name)
        sheet.Range(cell).ClearComments()
        return {"status": "success", "cell": cell, "message": "Comment removed."}

    # ── SECURITY & PROTECTION ─────────────────────────────────────────────────

    def protect_sheet(self, password: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Protect a worksheet to prevent accidental edits."""
        sheet = self.get_sheet(sheet_name)
        if password:
            sheet.Protect(Password=password)
        else:
            sheet.Protect()
        return {"status": "success", "sheet": sheet.Name, "protected": True}

    def unprotect_sheet(self, password: Optional[str] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Unprotect a worksheet."""
        sheet = self.get_sheet(sheet_name)
        if password:
            sheet.Unprotect(Password=password)
        else:
            sheet.Unprotect()
        return {"status": "success", "sheet": sheet.Name, "protected": False}

    # ── NAMED RANGES ──────────────────────────────────────────────────────────

    def create_named_range(self, name: str, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Create a workbook-level named range."""
        sheet = self.get_sheet(sheet_name)
        wb = self.get_workbook()
        refers_to = f"='{sheet.Name}'!{sheet.Range(range_address).Address}"
        wb.Names.Add(Name=name, RefersTo=refers_to)
        return {"status": "success", "name": name, "refers_to": refers_to}

    def list_named_ranges(self) -> List[Dict[str, Any]]:
        """List all defined named ranges in the workbook."""
        wb = self.get_workbook()
        result = []
        try:
            for n in wb.Names:
                result.append({"name": n.Name, "refers_to": n.RefersTo})
        except Exception:
            pass
        return result

    def delete_named_range(self, name: str) -> Dict[str, Any]:
        """Delete a named range."""
        wb = self.get_workbook()
        wb.Names(name).Delete()
        return {"status": "success", "deleted": name}

    # ── CALCULATION & AUDIT ───────────────────────────────────────────────────

    def calculate_now(self) -> Dict[str, Any]:
        """Force Excel to recalculate all formulas immediately."""
        app = self.get_app(bring_to_front=False)
        app.Calculate()
        return {"status": "success", "message": "Formulas recalculated."}

    def list_formulas(self, range_address: Optional[str] = None, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """List all cells containing formulas in a range or sheet."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address) if range_address else sheet.UsedRange
        results = []
        try:
            # xlCellTypeFormulas = -4123
            formula_cells = rng.SpecialCells(-4123)
            for c in formula_cells:
                results.append({
                    "cell": c.Address,
                    "formula": str(c.Formula),
                    "value": c.Value
                })
        except Exception:
            pass
        return results

    # ── PIVOT TABLES ──────────────────────────────────────────────────────────

    def create_pivot_table(
        self,
        source_range: str,
        dest_cell: str,
        table_name: str = "SummaryPivot",
        row_fields: Optional[List[str]] = None,
        value_fields: Optional[List[str]] = None,
        col_fields: Optional[List[str]] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a native Excel Pivot Table from a data range."""
        sheet = self.get_sheet(sheet_name)
        wb = self.get_workbook()
        src_rng = sheet.Range(source_range)
        dest_rng = sheet.Range(dest_cell)

        # xlDatabase = 1
        pc = wb.PivotCaches().Create(SourceType=1, SourceData=src_rng)
        pt = pc.CreatePivotTable(TableDestination=dest_rng, TableName=table_name)

        if row_fields:
            for f in row_fields:
                pt.PivotFields(f).Orientation = 1  # xlRowField

        if col_fields:
            for f in col_fields:
                pt.PivotFields(f).Orientation = 2  # xlColumnField

        if value_fields:
            for f in value_fields:
                # 4 = xlDataField, -4157 = xlSum
                field = pt.PivotFields(f)
                field.Orientation = 4
                field.Function = -4157

        return {
            "status": "success",
            "table_name": table_name,
            "destination": dest_cell,
            "source": source_range
        }

    # ── HYPERLINKS ────────────────────────────────────────────────────────────

    def add_hyperlink(
        self,
        cell: str,
        target: str,
        text_to_display: Optional[str] = None,
        tooltip: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add an external web URL or internal sheet link to a cell."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(cell)
        display_text = text_to_display or target
        is_internal = target.startswith("#") or "!" in target

        if is_internal:
            sub_addr = target.lstrip("#")
            sheet.Hyperlinks.Add(
                Anchor=rng,
                Address="",
                SubAddress=sub_addr,
                ScreenTip=tooltip or "",
                TextToDisplay=display_text
            )
        else:
            sheet.Hyperlinks.Add(
                Anchor=rng,
                Address=target,
                ScreenTip=tooltip or "",
                TextToDisplay=display_text
            )

        return {"status": "success", "cell": cell, "target": target, "text": display_text}

    def remove_hyperlink(self, cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Remove hyperlink from cell while preserving text content."""
        sheet = self.get_sheet(sheet_name)
        sheet.Range(cell).Hyperlinks.Delete()
        return {"status": "success", "cell": cell, "message": "Hyperlink removed."}

    # ── TAB COLOR & VIEW SETTINGS ─────────────────────────────────────────────

    def set_tab_color(self, hex_color: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Set worksheet tab color."""
        sheet = self.get_sheet(sheet_name)
        sheet.Tab.Color = hex_to_bgr_int(hex_color)
        return {"status": "success", "sheet": sheet.Name, "color": hex_color}

    def toggle_gridlines(self, show: bool = True) -> Dict[str, Any]:
        """Toggle worksheet gridlines visibility (hide for presentation-ready dashboards)."""
        app = self.get_app(bring_to_front=False)
        app.ActiveWindow.DisplayGridlines = show
        return {"status": "success", "gridlines_visible": show}

    def toggle_headings(self, show: bool = True) -> Dict[str, Any]:
        """Toggle worksheet row numbers and column letters visibility."""
        app = self.get_app(bring_to_front=False)
        app.ActiveWindow.DisplayHeadings = show
        return {"status": "success", "headings_visible": show}

    # ── PAGE SETUP & PRINT ────────────────────────────────────────────────────

    def set_page_setup(
        self,
        orientation: str = "landscape",
        paper_size: str = "A4",
        fit_to_pages: bool = True,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Configure page orientation, paper size, and print scaling."""
        sheet = self.get_sheet(sheet_name)
        ps = sheet.PageSetup

        # xlLandscape = 2, xlPortrait = 1
        ps.Orientation = 2 if orientation.lower() == "landscape" else 1

        # xlPaperA4 = 9, xlPaperLetter = 1
        ps.PaperSize = 9 if paper_size.upper() == "A4" else 1

        if fit_to_pages:
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1

        return {
            "status": "success",
            "sheet": sheet.Name,
            "orientation": orientation,
            "paper_size": paper_size,
            "fit_to_pages": fit_to_pages
        }

    def set_print_area(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Set specific print area for the worksheet."""
        sheet = self.get_sheet(sheet_name)
        sheet.PageSetup.PrintArea = range_address
        return {"status": "success", "print_area": range_address}

    def clear_print_area(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Clear the print area for the worksheet."""
        sheet = self.get_sheet(sheet_name)
        sheet.PageSetup.PrintArea = ""
        return {"status": "success", "message": "Print area cleared."}

    # ── ADVANCED DATA ANALYSIS & SEARCH ───────────────────────────────────────

    def find_all(self, search_term: str, match_case: bool = False, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """Find all cell occurrences of a text term across the active sheet."""
        sheet = self.get_sheet(sheet_name)
        # xlPart = 2
        first_found = sheet.Cells.Find(What=search_term, LookAt=2, MatchCase=match_case)
        if not first_found:
            return []

        results = []
        first_addr = first_found.Address
        current = first_found

        while current:
            results.append({
                "cell": current.Address,
                "value": current.Value,
                "formula": str(current.Formula) if current.HasFormula else None
            })
            current = sheet.Cells.FindNext(current)
            if not current or current.Address == first_addr:
                break

        return results

    def get_column_summary(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Calculate quick descriptive statistics for numeric data in a range."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        app = self.get_app(bring_to_front=False)
        wf = app.WorksheetFunction

        try:
            total_sum = wf.Sum(rng)
            avg = wf.Average(rng)
            min_val = wf.Min(rng)
            max_val = wf.Max(rng)
            count = wf.Count(rng)
            return {
                "range": range_address,
                "count": count,
                "sum": total_sum,
                "average": round(avg, 4),
                "min": min_val,
                "max": max_val
            }
        except Exception as e:
            return {"range": range_address, "error": f"Could not calculate summary: {e}"}

    def remove_duplicates(
        self,
        range_address: str,
        columns: Optional[List[int]] = None,
        has_headers: bool = True,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Remove duplicate rows from a data table."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        # xlYes = 1, xlNo = 2
        header_const = 1 if has_headers else 2
        col_indices = tuple(columns) if columns else tuple(range(1, rng.Columns.Count + 1))

        rng.RemoveDuplicates(Columns=col_indices, Header=header_const)
        return {"status": "success", "range": range_address, "columns_checked": list(col_indices)}

    def goal_seek(
        self,
        target_cell: str,
        target_value: float,
        changing_cell: str,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Perform Goal Seek what-if analysis to find the input needed for a desired formula result."""
        sheet = self.get_sheet(sheet_name)
        formula_rng = sheet.Range(target_cell)
        changing_rng = sheet.Range(changing_cell)

        success = formula_rng.GoalSeek(Goal=float(target_value), ChangingCell=changing_rng)
        return {
            "status": "success" if success else "failed",
            "target_cell": target_cell,
            "target_value": target_value,
            "changing_cell": changing_cell,
            "result_value": changing_rng.Value
        }

    def text_to_columns(
        self,
        range_address: str,
        destination_cell: str,
        delimiter: str = ",",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Split delimited text in a single column into multiple columns."""
        sheet = self.get_sheet(sheet_name)
        src_rng = sheet.Range(range_address)
        dest_rng = sheet.Range(destination_cell)

        # 1 = xlDelimited
        is_comma = delimiter == ","
        is_tab = delimiter == "\t"
        is_semicolon = delimiter == ";"
        is_space = delimiter == " "
        other_char = delimiter if (not is_comma and not is_tab and not is_semicolon and not is_space) else False

        src_rng.TextToColumns(
            Destination=dest_rng,
            DataType=1,
            TextQualifier=1,
            ConsecutiveDelimiter=False,
            Tab=is_tab,
            Semicolon=is_semicolon,
            Comma=is_comma,
            Space=is_space,
            Other=bool(other_char),
            OtherChar=other_char if other_char else ""
        )

        return {"status": "success", "source": range_address, "destination": destination_cell, "delimiter": delimiter}

    # ── SPARKLINES & GANTT BLOCKS ─────────────────────────────────────────────

    def add_sparkline(
        self,
        data_range: str,
        location_cell: str,
        sparkline_type: str = "line",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Insert a mini in-cell trend sparkline (line, column, win_loss)."""
        sheet = self.get_sheet(sheet_name)
        target_cell = sheet.Range(location_cell)
        data_rng = sheet.Range(data_range)

        # xlSparkLine = 1, xlSparkColumn = 2, xlSparkWinLoss = 3
        type_map = {"line": 1, "column": 2, "win_loss": 3}
        sp_type = type_map.get(sparkline_type.lower(), 1)

        target_cell.SparklineGroups.Add(Type=sp_type, SourceData=data_rng.Address)
        return {
            "status": "success",
            "location": location_cell,
            "data_source": data_range,
            "sparkline_type": sparkline_type
        }

    def create_gantt_block(
        self,
        start_cell: str,
        tasks: List[Dict[str, Any]],
        weeks_count: int = 6,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a modern project Gantt timeline grid."""
        sheet = self.get_sheet(sheet_name)
        tl = sheet.Range(start_cell)
        r = tl.Row
        c = tl.Column

        # Headers
        headers = ["Task Name", "Owner", "Start Wk", "Duration (Wks)"] + [f"W{i}" for i in range(1, weeks_count + 1)]
        for col_idx, h in enumerate(headers):
            cell = sheet.Cells(r, c + col_idx)
            cell.Value = h
            cell.Font.Bold = True
            cell.Interior.Color = hex_to_bgr_int("#1E293B")
            cell.Font.Color = hex_to_bgr_int("#FFFFFF")
            cell.HorizontalAlignment = -4108  # center

        # Task rows
        for t_idx, t in enumerate(tasks, start=1):
            curr_r = r + t_idx
            sheet.Cells(curr_r, c).Value = t.get("name", f"Task {t_idx}")
            sheet.Cells(curr_r, c + 1).Value = t.get("owner", "Team")
            st_wk = int(t.get("start_week", 1))
            dur = int(t.get("duration", 1))
            sheet.Cells(curr_r, c + 2).Value = st_wk
            sheet.Cells(curr_r, c + 3).Value = dur

            # Fill bar cells
            for w in range(1, weeks_count + 1):
                bar_cell = sheet.Cells(curr_r, c + 3 + w)
                if st_wk <= w < (st_wk + dur):
                    bar_cell.Interior.Color = hex_to_bgr_int(t.get("color", "#3B82F6"))
                else:
                    bar_cell.Interior.Color = hex_to_bgr_int("#F8FAFC")

        # Borders
        grid_rng = sheet.Range(sheet.Cells(r, c), sheet.Cells(r + len(tasks), c + len(headers) - 1))
        for b_id in [7, 8, 9, 10, 11, 12]:
            try:
                b = grid_rng.Borders(b_id)
                b.LineStyle = 1
                b.Weight = 2
                b.Color = hex_to_bgr_int("#CBD5E1")
            except Exception:
                pass

        return {
            "status": "success",
            "gantt_range": grid_rng.Address,
            "tasks_count": len(tasks),
            "weeks": weeks_count
        }

    # ── CELL STYLES & PRESETS ─────────────────────────────────────────────────

    def apply_cell_style(self, range_address: str, style_name: str = "Good", sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Apply an official Excel built-in style ('Good', 'Bad', 'Neutral', 'Title', 'Heading 1', 'Total', 'Input', etc.)."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        rng.Style = style_name
        return {"status": "success", "range": range_address, "style": style_name}

    def list_cell_styles(self) -> List[str]:
        """List all available workbook cell styles."""
        wb = self.get_workbook()
        styles = []
        try:
            for s in wb.Styles:
                styles.append(s.Name)
        except Exception:
            pass
        return styles

    def set_number_format_preset(self, range_address: str, preset: str = "currency_usd", sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Apply standard financial and business number format presets."""
        preset_map = {
            "currency_usd": "$#,##0.00",
            "currency_eur": "#,##0.00 €",
            "currency_mad": "#,##0.00 \"DH\"",
            "currency_int": "$#,##0",
            "percent": "0.0%",
            "percent_int": "0%",
            "accounting": "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
            "date_short": "yyyy-mm-dd",
            "date_long": "dd mmmm yyyy",
            "integer": "#,##0",
            "decimal_two": "0.00"
        }
        fmt = preset_map.get(preset.lower(), "$#,##0.00")
        sheet = self.get_sheet(sheet_name)
        sheet.Range(range_address).NumberFormat = fmt
        return {"status": "success", "range": range_address, "preset": preset, "format": fmt}

    # ── ADVANCED SHEET WORKFLOWS ──────────────────────────────────────────────

    def duplicate_sheet(self, sheet_name: str, new_name: Optional[str] = None) -> Dict[str, Any]:
        """Clone an existing worksheet with all formulas, charts, and styling intact."""
        src_sheet = self.get_sheet(sheet_name)
        src_sheet.Copy(After=src_sheet)
        new_sheet = self.get_sheet()
        if new_name:
            new_sheet.Name = new_name[:31]
        return {
            "status": "success",
            "source_sheet": sheet_name,
            "new_sheet": new_sheet.Name
        }

    def move_sheet(self, sheet_name: str, before: Optional[str] = None, after: Optional[str] = None) -> Dict[str, Any]:
        """Reorder worksheet tabs."""
        wb = self.get_workbook()
        target_sheet = self.get_sheet(sheet_name)
        if before:
            target_sheet.Move(Before=wb.Worksheets(before))
        elif after:
            target_sheet.Move(After=wb.Worksheets(after))
        return {"status": "success", "moved_sheet": sheet_name}

    def hide_unhide_sheet(self, sheet_name: str, hide: bool = True, very_hidden: bool = False) -> Dict[str, Any]:
        """Hide or unhide a worksheet (very_hidden prevents unhiding from Excel UI)."""
        sheet = self.get_sheet(sheet_name)
        if not hide:
            # -1 = xlSheetVisible
            sheet.Visible = -1
        else:
            # 2 = xlSheetVeryHidden, 0 = xlSheetHidden
            sheet.Visible = 2 if very_hidden else 0
        return {"status": "success", "sheet": sheet_name, "visible_state": sheet.Visible}

    # ── ADVANCED FORMULAS & CALCULATION ───────────────────────────────────────

    def convert_formulas_to_values(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Convert formula cells into static values (freeze calculated data)."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        rng.Value = rng.Value
        return {"status": "success", "range": range_address, "message": "Formulas converted to static values."}

    def evaluate_expression(self, expression: str) -> Dict[str, Any]:
        """Evaluate an arbitrary Excel formula expression directly using Excel's calculation engine."""
        app = self.get_app(bring_to_front=False)
        result = app.Evaluate(expression)
        return {
            "status": "success",
            "expression": expression,
            "result": str(result)
        }

    # ── VIEW, ZOOM & WINDOW SETTINGS ──────────────────────────────────────────

    def set_zoom(self, zoom_percentage: int = 100, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Set zoom scale (e.g. 75, 100, 120) for the active window."""
        app = self.get_app(bring_to_front=False)
        sheet = self.get_sheet(sheet_name)
        sheet.Activate()
        app.ActiveWindow.Zoom = max(10, min(400, zoom_percentage))
        return {"status": "success", "zoom": app.ActiveWindow.Zoom}

    def split_panes(self, row: Optional[int] = None, col: Optional[int] = None) -> Dict[str, Any]:
        """Split or unsplit the window pane."""
        app = self.get_app(bring_to_front=False)
        win = app.ActiveWindow
        if row is None and col is None:
            win.Split = False
            return {"status": "success", "split": False}
        win.Split = True
        if row:
            win.SplitRow = row
        if col:
            win.SplitColumn = col
        return {"status": "success", "split": True, "split_row": row, "split_col": col}

    # ── EXECUTIVE VISUAL COMPONENTS ───────────────────────────────────────────

    def add_callout_box(
        self,
        cell: str,
        text: str,
        title: Optional[str] = None,
        box_type: str = "info",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Insert a styled callout rectangle shape with colored accent."""
        sheet = self.get_sheet(sheet_name)
        target = sheet.Range(cell)

        # 1 = msoShapeRectangle
        shape = sheet.Shapes.AddShape(1, target.Left, target.Top, 350, 60)

        # Palette
        palette = {
            "info": ("#EFF6FF", "#3B82F6", "#1E40AF"),      # bg, border, text
            "warning": ("#FEFCE8", "#EAB308", "#854D0E"),
            "success": ("#F0FDF4", "#22C55E", "#166534"),
            "danger": ("#FEF2F2", "#EF4444", "#991B1B"),
        }
        bg, border, text_col = palette.get(box_type.lower(), palette["info"])

        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = hex_to_bgr_int(bg)
        shape.Line.ForeColor.RGB = hex_to_bgr_int(border)
        shape.Line.Weight = 1.5

        full_text = f"{title.upper()}\n{text}" if title else text
        tf = shape.TextFrame2
        tf.TextRange.Text = full_text
        tf.TextRange.Font.Fill.ForeColor.RGB = hex_to_bgr_int(text_col)
        tf.TextRange.Font.Size = 10
        tf.TextRange.Font.Name = "Segoe UI"

        return {
            "status": "success",
            "cell": cell,
            "box_type": box_type,
            "shape_name": shape.Name
        }

    def add_metric_tile(
        self,
        cell: str,
        label: str,
        value: str,
        delta: Optional[str] = None,
        status: str = "positive",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a compact 2-row metric tile."""
        sheet = self.get_sheet(sheet_name)
        top = sheet.Range(cell)
        r, c = top.Row, top.Column

        # Tile block (2x2)
        tile_rng = sheet.Range(sheet.Cells(r, c), sheet.Cells(r + 1, c + 1))
        tile_rng.Interior.Color = hex_to_bgr_int("#FFFFFF")

        for b_id in [7, 8, 9, 10]:
            try:
                b = tile_rng.Borders(b_id)
                b.LineStyle = 1
                b.Weight = 2
                b.Color = hex_to_bgr_int("#E2E8F0")
            except Exception:
                pass

        # Label
        lbl_cell = sheet.Cells(r, c)
        lbl_cell.Value = label.upper()
        lbl_cell.Font.Size = 8
        lbl_cell.Font.Bold = True
        lbl_cell.Font.Color = hex_to_bgr_int("#64748B")

        # Value
        val_cell = sheet.Cells(r + 1, c)
        val_cell.Value = value
        val_cell.Font.Size = 14
        val_cell.Font.Bold = True
        val_cell.Font.Color = hex_to_bgr_int("#0F172A")

        # Delta
        if delta:
            d_cell = sheet.Cells(r + 1, c + 1)
            d_cell.Value = delta
            d_cell.Font.Size = 9
            d_cell.Font.Bold = True
            delta_color = "#059669" if status == "positive" else ("#DC2626" if status == "negative" else "#64748B")
            d_cell.Font.Color = hex_to_bgr_int(delta_color)

        return {"status": "success", "tile_range": tile_rng.Address, "label": label, "value": value}

    # ── PRINT HEADERS, FOOTERS & HTML EXPORT ──────────────────────────────────

    def set_header_footer(
        self,
        header_left: Optional[str] = None,
        header_center: Optional[str] = None,
        header_right: Optional[str] = None,
        footer_left: Optional[str] = None,
        footer_center: Optional[str] = None,
        footer_right: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Configure page header and footer strings (supports &P page, &N pages, &D date)."""
        sheet = self.get_sheet(sheet_name)
        ps = sheet.PageSetup

        if header_left is not None:
            ps.LeftHeader = header_left
        if header_center is not None:
            ps.CenterHeader = header_center
        if header_right is not None:
            ps.RightHeader = header_right
        if footer_left is not None:
            ps.LeftFooter = footer_left
        if footer_center is not None:
            ps.CenterFooter = footer_center
        if footer_right is not None:
            ps.RightFooter = footer_right

        return {"status": "success", "sheet": sheet.Name, "message": "Headers and footers updated."}

    def add_checkbox(
        self,
        cell: str,
        caption: str = "Completed",
        linked_cell: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Insert an interactive Form Control Checkbox anchored at a cell."""
        sheet = self.get_sheet(sheet_name)
        target = sheet.Range(cell)

        # xlCheckBox = 1
        cb = sheet.CheckBoxes().Add(target.Left + 2, target.Top + 2, 100, target.Height - 4)
        cb.Caption = caption
        if linked_cell:
            cb.LinkedCell = linked_cell

        return {"status": "success", "cell": cell, "caption": caption, "linked_cell": linked_cell}

    def export_html(self, output_path: str, sheet_only: bool = False) -> Dict[str, Any]:
        """Export workbook or current sheet as a standalone HTML document."""
        wb = self.get_workbook()
        abs_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)

        app = self.get_app(bring_to_front=False)
        app.DisplayAlerts = False
        try:
            # 44 = xlHtml
            if sheet_only:
                # 44 = xlHtml
                wb.PublishObjects.Add(
                    SourceType=4, # xlSourceSheet
                    Filename=abs_path,
                    Sheet=wb.ActiveSheet.Name,
                    HtmlType=0 # xlHtmlStatic
                ).Publish(True)
            else:
                wb.SaveAs(abs_path, 44)
        finally:
            app.DisplayAlerts = True

        return {"status": "success", "output_path": abs_path}

    def add_watermark(
        self,
        text: str = "CONFIDENTIAL",
        font_size: int = 54,
        color: str = "#E2E8F0",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a semi-transparent diagonal watermark across the sheet."""
        sheet = self.get_sheet(sheet_name)
        # 0 = msoTextEffect1
        art = sheet.Shapes.AddTextEffect(0, text, "Segoe UI", font_size, False, False, 150, 150)
        art.Rotation = -35
        art.Fill.Solid()
        art.Fill.ForeColor.RGB = hex_to_bgr_int(color)
        art.Line.Visible = False
        return {"status": "success", "watermark": text, "shape": art.Name}

    def clear_all_formatting(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Reset all formatting across the used range to standard defaults."""
        sheet = self.get_sheet(sheet_name)
        sheet.UsedRange.ClearFormats()
        return {"status": "success", "message": "All formatting cleared across used range."}

    # ── ROW & COLUMN OUTLINING / GROUPING ─────────────────────────────────────

    def group_rows(self, start_row: int, end_row: int, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Group rows for collapsibility/outline."""
        sheet = self.get_sheet(sheet_name)
        sheet.Rows(f"{start_row}:{end_row}").Group()
        return {"status": "success", "start_row": start_row, "end_row": end_row, "message": "Rows grouped."}

    def ungroup_rows(self, start_row: int, end_row: int, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Ungroup rows to remove outline level."""
        sheet = self.get_sheet(sheet_name)
        sheet.Rows(f"{start_row}:{end_row}").Ungroup()
        return {"status": "success", "start_row": start_row, "end_row": end_row, "message": "Rows ungrouped."}

    def group_columns(self, start_col: str, end_col: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Group columns for collapsibility/outline."""
        sheet = self.get_sheet(sheet_name)
        sheet.Columns(f"{start_col}:{end_col}").Group()
        return {"status": "success", "start_col": start_col, "end_col": end_col, "message": "Columns grouped."}

    def ungroup_columns(self, start_col: str, end_col: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Ungroup columns to remove outline level."""
        sheet = self.get_sheet(sheet_name)
        sheet.Columns(f"{start_col}:{end_col}").Ungroup()
        return {"status": "success", "start_col": start_col, "end_col": end_col, "message": "Columns ungrouped."}

    def show_outline_levels(self, row_levels: Optional[int] = None, col_levels: Optional[int] = None, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Expand or collapse grouped rows/columns to specified outline depth (1=collapsed, 2+=expanded)."""
        sheet = self.get_sheet(sheet_name)
        if row_levels is not None:
            sheet.Outline.ShowLevels(RowLevels=row_levels)
        if col_levels is not None:
            sheet.Outline.ShowLevels(ColumnLevels=col_levels)
        return {"status": "success", "row_levels": row_levels, "col_levels": col_levels}

    # ── ADVANCED CONDITIONAL FORMATTING PRESETS ────────────────────────────────

    def add_top_bottom_rule(
        self,
        range_address: str,
        rule_type: str = "top_percent",
        rank: int = 10,
        bg_color: str = "#DCFCE7",
        font_color: str = "#166534",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Highlight Top N or Bottom N items/percents in a numeric range."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        # xlTop10 = 5
        cond = rng.FormatConditions.AddTop10()
        is_bottom = "bottom" in rule_type.lower()
        is_percent = "percent" in rule_type.lower()
        cond.TopBottom = 2 if is_bottom else 1  # 1 = xlTop10Top, 2 = xlTop10Bottom
        cond.Percent = is_percent
        cond.Rank = rank
        cond.Interior.Color = hex_to_bgr_int(bg_color)
        cond.Font.Color = hex_to_bgr_int(font_color)
        return {"status": "success", "range": range_address, "rule_type": rule_type, "rank": rank}

    def add_above_below_average_rule(
        self,
        range_address: str,
        above: bool = True,
        bg_color: str = "#FEF3C7",
        font_color: str = "#92400E",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Highlight cells that are above or below the range average."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        # xlAboveAverageCondition = 12
        cond = rng.FormatConditions.AddAboveAverage()
        cond.AboveBelow = 0 if above else 1  # 0 = xlAboveAverage, 1 = xlBelowAverage
        cond.Interior.Color = hex_to_bgr_int(bg_color)
        cond.Font.Color = hex_to_bgr_int(font_color)
        return {"status": "success", "range": range_address, "above": above}

    def add_icon_set_rule(
        self,
        range_address: str,
        icon_set_type: str = "3_traffic_lights",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add visual icon set (traffic lights, arrows, flags) to cells."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        # xlIconSet = 6
        cond = rng.FormatConditions.AddIconSetCondition()
        icon_map = {
            "3_traffic_lights": 1, # xl3TrafficLights1
            "3_arrows": 2,         # xl3Arrows
            "3_flags": 3,          # xl3Flags
            "3_symbols": 6,        # xl3Symbols
            "4_arrows": 7,         # xl4Arrows
            "5_quarters": 17       # xl5Quarters
        }
        cond.IconSet = icon_map.get(icon_set_type.lower(), 1)
        return {"status": "success", "range": range_address, "icon_set": icon_set_type}

    def highlight_duplicate_values(
        self,
        range_address: str,
        bg_color: str = "#FEE2E2",
        font_color: str = "#991B1B",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Highlight duplicate values across a range in real-time."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        # xlUniqueValues = 8
        cond = rng.FormatConditions.AddUniqueValues()
        cond.DupeUnique = 1  # 1 = xlDuplicate, 0 = xlUnique
        cond.Interior.Color = hex_to_bgr_int(bg_color)
        cond.Font.Color = hex_to_bgr_int(font_color)
        return {"status": "success", "range": range_address, "message": "Duplicates highlighted."}

    # ── FORMULA AUDITING & TRACING ────────────────────────────────────────────

    def trace_precedents(self, cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Draw visual audit arrows pointing to the cells that provide data to this formula."""
        sheet = self.get_sheet(sheet_name)
        sheet.Range(cell).ShowPrecedents()
        return {"status": "success", "cell": cell, "message": "Precedent tracer arrows shown."}

    def trace_dependents(self, cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Draw visual audit arrows pointing to cells that depend on this cell value."""
        sheet = self.get_sheet(sheet_name)
        sheet.Range(cell).ShowDependents()
        return {"status": "success", "cell": cell, "message": "Dependent tracer arrows shown."}

    def remove_tracer_arrows(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Clear all precedent and dependent audit tracer arrows."""
        sheet = self.get_sheet(sheet_name)
        sheet.ClearArrows()
        return {"status": "success", "message": "All audit tracer arrows cleared."}

    # ── WORKBOOK STRUCTURE & WINDOW SECURITY ──────────────────────────────────

    def protect_workbook(self, password: Optional[str] = None, lock_structure: bool = True, lock_windows: bool = False) -> Dict[str, Any]:
        """Protect workbook structure to prevent adding, deleting, or renaming sheets."""
        wb = self.get_workbook()
        if password:
            wb.Protect(Password=password, Structure=lock_structure, Windows=lock_windows)
        else:
            wb.Protect(Structure=lock_structure, Windows=lock_windows)
        return {"status": "success", "workbook": wb.Name, "structure_locked": lock_structure}

    def unprotect_workbook(self, password: Optional[str] = None) -> Dict[str, Any]:
        """Remove workbook structure protection."""
        wb = self.get_workbook()
        if password:
            wb.Unprotect(Password=password)
        else:
            wb.Unprotect()
        return {"status": "success", "workbook": wb.Name, "protected": False}

    # ── SHAPES, ARROWS & BADGES ───────────────────────────────────────────────

    def add_shape(
        self,
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
        """Insert a native shape (rectangle, rounded_rectangle, oval, arrow_right) with optional text."""
        sheet = self.get_sheet(sheet_name)
        target = sheet.Range(cell)
        shape_map = {
            "rectangle": 1,           # msoShapeRectangle
            "rounded_rectangle": 5,   # msoShapeRoundedRectangle
            "oval": 9,                # msoShapeOval
            "arrow_right": 13         # msoShapeRightArrow
        }
        mso_type = shape_map.get(shape_type.lower(), 5)
        shape = sheet.Shapes.AddShape(mso_type, target.Left, target.Top, width, height)
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = hex_to_bgr_int(fill_color)
        shape.Line.ForeColor.RGB = hex_to_bgr_int(line_color)
        if text:
            tf = shape.TextFrame2
            tf.TextRange.Text = text
            tf.TextRange.Font.Fill.ForeColor.RGB = hex_to_bgr_int(font_color)
            tf.TextRange.Font.Name = "Segoe UI"
            tf.TextRange.Font.Size = 10
            tf.TextRange.Font.Bold = True
        return {"status": "success", "shape_name": shape.Name, "cell": cell, "type": shape_type}

    def delete_shape(self, shape_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Delete a shape, image, or drawing object by name."""
        sheet = self.get_sheet(sheet_name)
        sheet.Shapes(shape_name).Delete()
        return {"status": "success", "deleted_shape": shape_name}

    def list_shapes(self, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """List all shapes, images, and drawing objects on the sheet."""
        sheet = self.get_sheet(sheet_name)
        items = []
        for s in sheet.Shapes:
            items.append({
                "name": s.Name,
                "type": s.Type,
                "left": round(s.Left, 1),
                "top": round(s.Top, 1),
                "width": round(s.Width, 1),
                "height": round(s.Height, 1)
            })
        return items

    # ── SPECIAL CELLS SELECTION & FORMULA AUDIT ───────────────────────────────

    def get_special_cells_summary(self, cell_type: str = "blanks", sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Inspect and count special cells across the sheet: 'blanks', 'constants', 'formulas', 'comments'."""
        sheet = self.get_sheet(sheet_name)
        type_map = {
            "blanks": 4,       # xlCellTypeBlanks
            "constants": 2,    # xlCellTypeConstants
            "formulas": -4123, # xlCellTypeFormulas
            "comments": -4144  # xlCellTypeComments
        }
        xl_type = type_map.get(cell_type.lower(), 4)
        try:
            cells = sheet.UsedRange.SpecialCells(xl_type)
            return {"status": "success", "type": cell_type, "count": cells.Count, "addresses": cells.Address}
        except Exception:
            return {"status": "success", "type": cell_type, "count": 0, "addresses": None}

    def lock_unlock_cells(self, range_address: str, locked: bool = True, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Lock or unlock specific cells (unlocked cells can still be edited when sheet is protected)."""
        sheet = self.get_sheet(sheet_name)
        sheet.Range(range_address).Locked = locked
        return {"status": "success", "range": range_address, "locked": locked}

    # ── ADVANCED PRINTING & PAGE LAYOUT TITLES ────────────────────────────────

    def set_repeat_rows_cols(
        self,
        repeat_rows: Optional[str] = None,
        repeat_cols: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Set repeating header rows and/or columns for multi-page printouts (Print Titles)."""
        sheet = self.get_sheet(sheet_name)
        if repeat_rows is not None:
            sheet.PageSetup.PrintTitleRows = repeat_rows
        if repeat_cols is not None:
            sheet.PageSetup.PrintTitleColumns = repeat_cols
        return {"status": "success", "repeat_rows": repeat_rows, "repeat_cols": repeat_cols}

    def insert_page_break(self, before_cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Insert manual horizontal and vertical page breaks before a specific cell."""
        sheet = self.get_sheet(sheet_name)
        target = sheet.Range(before_cell)
        sheet.HPageBreaks.Add(Before=target)
        sheet.VPageBreaks.Add(Before=target)
        return {"status": "success", "before_cell": before_cell, "message": "Page break inserted."}

    def reset_page_breaks(self, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Reset all manual page breaks back to default pagination."""
        sheet = self.get_sheet(sheet_name)
        sheet.ResetAllPageBreaks()
        return {"status": "success", "message": "All manual page breaks reset."}

    # ── CALCULATION ENGINE MODES ──────────────────────────────────────────────

    def set_calculation_mode(self, mode: str = "automatic") -> Dict[str, Any]:
        """Configure Excel calculation mode ('automatic', 'manual', or 'semiautomatic')."""
        app = self.get_app(bring_to_front=False)
        # xlCalculationAutomatic = -4105, xlCalculationManual = -4135, xlCalculationSemiautomatic = 2
        mode_map = {
            "automatic": -4105,
            "manual": -4135,
            "semiautomatic": 2
        }
        val = mode_map.get(mode.lower(), -4105)
        app.Calculation = val
        return {"status": "success", "mode": mode, "calculation_val": val}

    def get_calculation_mode(self) -> Dict[str, Any]:
        """Get current Excel calculation mode."""
        app = self.get_app(bring_to_front=False)
        inv_map = {-4105: "automatic", -4135: "manual", 2: "semiautomatic"}
        return {"status": "success", "mode": inv_map.get(app.Calculation, "unknown")}

    # ── DATA TRANSFORMATIONS & ADVANCED UTILITIES ─────────────────────────────

    def transpose_range(self, source_range: str, dest_cell: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Transpose (swap rows and columns) of a cell range to a destination cell."""
        sheet = self.get_sheet(sheet_name)
        src = sheet.Range(source_range)
        src.Copy()
        dest = sheet.Range(dest_cell)
        # xlPasteAll = -4104, xlPasteSpecialOperationNone = -4142
        dest.PasteSpecial(Paste=-4104, Operation=-4142, SkipBlanks=False, Transpose=True)
        self.get_app(bring_to_front=False).CutCopyMode = False
        return {"status": "success", "source": source_range, "destination": dest_cell}

    def fill_series(self, range_address: str, fill_type: str = "linear", step: float = 1.0, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Auto-fill a sequence or series (linear, growth, date)."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(range_address)
        # xlDataSeriesLinear = -4132, xlGrowth = 2, xlChronological = 3
        type_map = {"linear": -4132, "growth": 2, "date": 3}
        xl_type = type_map.get(fill_type.lower(), -4132)
        # xlRows = 1, xlColumns = 2
        series_by = 1 if rng.Rows.Count == 1 else 2
        rng.DataSeries(Rowcol=series_by, Type=xl_type, Step=step)
        return {"status": "success", "range": range_address, "fill_type": fill_type, "step": step}

    # ── ADVANCED WINDOW & VIEW PERSPECTIVES ────────────────────────────────────

    def set_view_mode(self, view_mode: str = "normal", sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Set workbook view mode: 'normal', 'page_break_preview', or 'page_layout'."""
        app = self.get_app(bring_to_front=False)
        sheet = self.get_sheet(sheet_name)
        sheet.Activate()
        # xlNormalView = 1, xlPageBreakPreview = 2, xlPageLayoutView = 3
        view_map = {"normal": 1, "page_break_preview": 2, "page_layout": 3}
        val = view_map.get(view_mode.lower(), 1)
        app.ActiveWindow.View = val
        return {"status": "success", "view_mode": view_mode}

    def get_formula_errors(self, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """Scan sheet for formula errors like #DIV/0!, #N/A, #VALUE!, #REF!, #NAME?."""
        sheet = self.get_sheet(sheet_name)
        errors = []
        try:
            # 16 = xlErrors
            error_cells = sheet.UsedRange.SpecialCells(-4123, 16)
            for c in error_cells:
                errors.append({
                    "cell": c.Address,
                    "formula": str(c.Formula),
                    "error_text": str(c.Text)
                })
        except Exception:
            pass
        return errors

    # ── ADVANCED SENSITIVITY WHAT-IF: DATA TABLES ─────────────────────────────

    def create_two_variable_data_table(
        self,
        table_range: str,
        row_input_cell: str,
        col_input_cell: str,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Generate a 2-variable sensitivity Data Table matrix (e.g. Price vs Volume sensitivity)."""
        sheet = self.get_sheet(sheet_name)
        tbl_rng = sheet.Range(table_range)
        row_inp = sheet.Range(row_input_cell)
        col_inp = sheet.Range(col_input_cell)
        tbl_rng.Table(RowInput=row_inp, ColumnInput=col_inp)
        return {
            "status": "success",
            "table_range": table_range,
            "row_input_cell": row_input_cell,
            "col_input_cell": col_input_cell
        }

    # ── EXECUTIVE NATIVE SLICERS & FILTER ENHANCEMENTS ────────────────────────

    def add_table_slicer(
        self,
        table_name: str,
        column_name: str,
        slicer_name: Optional[str] = None,
        left: float = 400,
        top: float = 60,
        width: float = 160,
        height: float = 220,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Insert an interactive graphic visual Slicer button box for an Excel Table."""
        sheet = self.get_sheet(sheet_name)
        wb = self.get_workbook()
        tbl = sheet.ListObjects(table_name)
        s_name = slicer_name or f"Slicer_{column_name}"
        sc = wb.SlicerCaches.Add2(tbl, column_name)
        slicer = sc.Slicers.Add(sheet, Top=top, Left=left, Width=width, Height=height, Name=s_name)
        return {"status": "success", "table": table_name, "column": column_name, "slicer_name": slicer.Name}

    def convert_table_to_range(self, table_name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Convert an official Excel Table back to a standard cell range (keeps all data & formatting)."""
        sheet = self.get_sheet(sheet_name)
        tbl = sheet.ListObjects(table_name)
        addr = tbl.Range.Address
        tbl.Unlist()
        return {"status": "success", "table_converted": table_name, "range": addr}

    # ── SHEET EXPORT & CELL CONTENT CLEARING ──────────────────────────────────

    def export_sheet_to_csv(self, output_path: str, sheet_name: Optional[str] = None) -> str:
        """Export the active worksheet directly to a CSV file."""
        sheet = self.get_sheet(sheet_name)
        abs_out = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_out), exist_ok=True)
        app = self.get_app(bring_to_front=False)
        app.DisplayAlerts = False
        try:
            # Copy to temp new workbook, then save as CSV (xlCSV = 6)
            sheet.Copy()
            new_wb = app.ActiveWorkbook
            new_wb.SaveAs(abs_out, 6)
            new_wb.Close(SaveChanges=False)
        finally:
            app.DisplayAlerts = True
        return abs_out

    def get_merged_ranges(self, sheet_name: Optional[str] = None) -> List[str]:
        """Find and list all merged cell range addresses on the worksheet."""
        sheet = self.get_sheet(sheet_name)
        merged = []
        seen = set()
        for cell in sheet.UsedRange:
            try:
                if cell.MergeCells:
                    ma = cell.MergeArea.Address
                    if ma not in seen:
                        seen.add(ma)
                        merged.append(ma)
            except Exception:
                pass
        return merged

    def clear_cell_contents(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Clear values and formulas in a range while keeping cell formatting intact."""
        sheet = self.get_sheet(sheet_name)
        sheet.Range(range_address).ClearContents()
        return {"status": "success", "range": range_address, "cleared": "contents"}

    def clear_cell_formats(self, range_address: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Clear cell formatting (colors, borders, styles) while keeping values and formulas intact."""
        sheet = self.get_sheet(sheet_name)
        sheet.Range(range_address).ClearFormats()
        return {"status": "success", "range": range_address, "cleared": "formats"}

    # ── DATA CONNECTIONS & POWER QUERY ────────────────────────────────────────

    def refresh_all_connections(self) -> Dict[str, Any]:
        """Refresh all external data connections, queries, and PivotTables in the workbook."""
        wb = self.get_workbook()
        wb.RefreshAll()
        return {"status": "success", "refreshed": "all_connections"}

    def get_data_connections(self) -> List[Dict[str, Any]]:
        """List all external data connections (ODBC, OLEDB, Power Query) in the workbook."""
        wb = self.get_workbook()
        conns = []
        for c in wb.Connections:
            try:
                conns.append({
                    "name": c.Name,
                    "type": getattr(c, "Type", None),
                    "description": getattr(c, "Description", "")
                })
            except Exception:
                pass
        return conns

    def refresh_connection(self, connection_name: str) -> Dict[str, Any]:
        """Refresh a specific data connection by name."""
        wb = self.get_workbook()
        wb.Connections(connection_name).Refresh()
        return {"status": "success", "connection_refreshed": connection_name}

    # ── ADVANCED CHART CUSTOMIZATION & EXPORT ──────────────────────────────────

    def set_chart_secondary_axis(
        self,
        chart_name: str,
        series_index: int = 2,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Move a chart data series to the secondary vertical (Y) axis."""
        sheet = self.get_sheet(sheet_name)
        chart = sheet.ChartObjects(chart_name).Chart
        series = chart.SeriesCollection(series_index)
        # xlSecondary = 2
        series.AxisGroup = 2
        return {"status": "success", "chart": chart_name, "series_index": series_index, "axis_group": "secondary"}

    def add_chart_trendline(
        self,
        chart_name: str,
        series_index: int = 1,
        trendline_type: str = "linear",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Add a trendline (linear, exponential, logarithmic, moving_average, polynomial) to a chart series."""
        sheet = self.get_sheet(sheet_name)
        chart = sheet.ChartObjects(chart_name).Chart
        series = chart.SeriesCollection(series_index)
        type_map = {
            "linear": -4132,          # xlLinear
            "exponential": 1,         # xlExponential
            "logarithmic": -4133,     # xlLogarithmic
            "moving_average": 4,      # xlMovingAvg
            "polynomial": 6           # xlPolynomial
        }
        xl_t = type_map.get(trendline_type.lower(), -4132)
        trend = series.Trendlines().Add(Type=xl_t)
        return {"status": "success", "chart": chart_name, "series": series_index, "trendline_type": trendline_type}

    def export_chart_as_image(
        self,
        chart_name: str,
        output_image_path: str,
        sheet_name: Optional[str] = None
    ) -> str:
        """Export a specific chart as a high-resolution PNG or JPG image file."""
        sheet = self.get_sheet(sheet_name)
        chart = sheet.ChartObjects(chart_name).Chart
        abs_out = os.path.abspath(output_image_path)
        os.makedirs(os.path.dirname(abs_out), exist_ok=True)
        chart.Export(abs_out)
        return abs_out

    # ── WHAT-IF SCENARIO MANAGER ──────────────────────────────────────────────

    def add_scenario(
        self,
        name: str,
        changing_cells: str,
        values: List[Any],
        comment: str = "",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a What-If Scenario with defined changing cell values (e.g. Best Case, Worst Case)."""
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(changing_cells)
        sheet.Scenarios.Create(name, rng, values, comment)
        return {"status": "success", "scenario": name, "changing_cells": changing_cells, "values": values}

    def show_scenario(self, name: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Apply and display a saved What-If Scenario on the worksheet."""
        sheet = self.get_sheet(sheet_name)
        sheet.Scenarios(name).Show()
        return {"status": "success", "scenario_applied": name}

    def list_scenarios(self, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """List all What-If Scenarios saved on the worksheet."""
        sheet = self.get_sheet(sheet_name)
        scs = []
        for s in sheet.Scenarios:
            try:
                scs.append({
                    "name": s.Name,
                    "comment": getattr(s, "Comment", ""),
                    "changing_cells": s.ChangingCells.Address
                })
            except Exception:
                pass
        return scs

    # ── CELL INSPECTION & AUDITING TOOLS ──────────────────────────────────────

    def find_cells_with_comments(self, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """Find and return all cells containing comments or notes on the sheet."""
        sheet = self.get_sheet(sheet_name)
        res = []
        try:
            # -4144 = xlCellTypeComments
            cells = sheet.UsedRange.SpecialCells(-4144)
            for c in cells:
                c_text = ""
                try:
                    if c.Comment:
                        c_text = c.Comment.Text()
                except Exception:
                    pass
                res.append({
                    "cell": c.Address,
                    "value": str(c.Value),
                    "comment": c_text
                })
        except Exception:
            pass
        return res

    def find_cells_with_data_validation(self, sheet_name: Optional[str] = None) -> List[str]:
        """Find and return all cell addresses that have Data Validation rules applied."""
        sheet = self.get_sheet(sheet_name)
        addresses = []
        try:
            # -4174 = xlCellTypeAllValidation
            cells = sheet.UsedRange.SpecialCells(-4174)
            for c in cells:
                addresses.append(c.Address)
        except Exception:
            pass
        return addresses

    # ── ADVANCED PAGE SETUP & PRINT METRICS ────────────────────────────────────

    def set_sheet_margins(
        self,
        top_inches: float = 0.75,
        bottom_inches: float = 0.75,
        left_inches: float = 0.7,
        right_inches: float = 0.7,
        header_inches: float = 0.3,
        footer_inches: float = 0.3,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Configure page margins in inches for printing or PDF export."""
        sheet = self.get_sheet(sheet_name)
        ps = sheet.PageSetup
        ps.TopMargin = self.app.InchesToPoints(top_inches)
        ps.BottomMargin = self.app.InchesToPoints(bottom_inches)
        ps.LeftMargin = self.app.InchesToPoints(left_inches)
        ps.RightMargin = self.app.InchesToPoints(right_inches)
        ps.HeaderMargin = self.app.InchesToPoints(header_inches)
        ps.FooterMargin = self.app.InchesToPoints(footer_inches)
        return {
            "status": "success",
            "margins": {
                "top": top_inches,
                "bottom": bottom_inches,
                "left": left_inches,
                "right": right_inches,
                "header": header_inches,
                "footer": footer_inches
            }
        }

    def set_sheet_orientation_and_paper(
        self,
        orientation: str = "landscape",
        paper_size: str = "A4",
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Set page orientation ('portrait' or 'landscape') and paper size ('A4', 'Letter', 'Legal')."""
        sheet = self.get_sheet(sheet_name)
        ps = sheet.PageSetup
        # xlPortrait = 1, xlLandscape = 2
        ps.Orientation = 2 if orientation.lower() == "landscape" else 1

        # xlPaperLetter = 1, xlPaperLegal = 5, xlPaperA4 = 9, xlPaperA3 = 8
        paper_map = {
            "letter": 1,
            "legal": 5,
            "a4": 9,
            "a3": 8
        }
        ps.PaperSize = paper_map.get(paper_size.lower(), 9)
        return {"status": "success", "orientation": orientation, "paper_size": paper_size}

    def set_sheet_fit_to_pages(
        self,
        fit_to_wide: int = 1,
        fit_to_tall: int = 1,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Fit worksheet to a specific number of pages wide and tall for printing / PDF export."""
        sheet = self.get_sheet(sheet_name)
        ps = sheet.PageSetup
        ps.Zoom = False
        ps.FitToPagesWide = fit_to_wide
        ps.FitToPagesTall = fit_to_tall
        return {"status": "success", "fit_to_wide": fit_to_wide, "fit_to_tall": fit_to_tall}


# Singleton instance
controller = ExcelController()
