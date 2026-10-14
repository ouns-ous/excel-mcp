"""Headless Excel OOXML Engine using openpyxl.
Allows fast headless creation, reading, and transformation of Excel (.xlsx) files
without requiring Microsoft Excel desktop to be installed or running.
"""

import os
import re
from typing import Any, Dict, List, Optional, Union
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class ExcelEngine:
    """Headless processor for .xlsx files using openpyxl."""

    @staticmethod
    def create_workbook(
        output_path: str,
        data: Optional[List[List[Any]]] = None,
        sheet_name: str = "Sheet1",
        header_bg: str = "1F4E79",
        header_fg: str = "FFFFFF"
    ) -> str:
        """Create a new formatted .xlsx workbook headlessly."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

        if data:
            for r_idx, row in enumerate(data, start=1):
                for c_idx, val in enumerate(row, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    if r_idx == 1:
                        cell.font = Font(bold=True, color=header_fg.lstrip("#"), name="Segoe UI")
                        cell.fill = PatternFill(start_color=header_bg.lstrip("#"), end_color=header_bg.lstrip("#"), fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.font = Font(name="Segoe UI")

            # Auto-fit columns
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

        abs_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        wb.save(abs_path)
        return abs_path

    @staticmethod
    def read_sheet_data(
        file_path: str,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """Read all rows and columns from an .xlsx file headlessly."""
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"File not found: {abs_path}")

        wb = openpyxl.load_workbook(abs_path, data_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

        rows = []
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                rows.append(list(row))

        return {
            "file": abs_path,
            "sheet": ws.title,
            "total_rows": len(rows),
            "total_cols": len(rows[0]) if rows else 0,
            "rows": rows
        }

    @staticmethod
    def markdown_to_excel(markdown_table: str, output_path: str, sheet_name: str = "Data") -> str:
        """Convert a Markdown formatted table into a styled Excel (.xlsx) spreadsheet."""
        lines = [line.strip() for line in markdown_table.strip().splitlines() if line.strip()]
        data = []

        for line in lines:
            if not line.startswith("|") or not line.endswith("|"):
                continue
            cells = [c.strip() for c in line.split("|")[1:-1]]
            # Skip separator line (e.g. |---|---|)
            if all(re.match(r"^:?-+:?$", c) for c in cells):
                continue
            data.append(cells)

        if not data:
            raise ValueError("No valid Markdown table detected.")

        return ExcelEngine.create_workbook(output_path, data, sheet_name=sheet_name)

    @staticmethod
    def excel_to_markdown(file_path: str, sheet_name: Optional[str] = None) -> str:
        """Extract table from .xlsx and format as a standard Markdown table."""
        res = ExcelEngine.read_sheet_data(file_path, sheet_name)
        rows = res["rows"]
        if not rows:
            return "*(Empty table)*"

        md_lines = []
        header = [str(c or "") for c in rows[0]]
        md_lines.append("| " + " | ".join(header) + " |")
        md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")

        for row in rows[1:]:
            cells = [str(c or "") for c in row]
            # Match header width
            while len(cells) < len(header):
                cells.append("")
            md_lines.append("| " + " | ".join(cells[:len(header)]) + " |")

        return "\n".join(md_lines)

    @staticmethod
    def json_to_excel(json_data: List[Dict[str, Any]], output_path: str, sheet_name: str = "Data") -> str:
        """Convert a list of dictionaries (JSON records) to a formatted Excel file."""
        if not json_data:
            raise ValueError("json_data is empty.")

        headers = list(json_data[0].keys())
        rows = [headers]
        for item in json_data:
            rows.append([item.get(h, "") for h in headers])

        return ExcelEngine.create_workbook(output_path, rows, sheet_name=sheet_name)

    @staticmethod
    def csv_to_excel(csv_path: str, output_path: str, delimiter: str = ",", sheet_name: str = "Data") -> str:
        """Convert a CSV file into a styled Excel (.xlsx) spreadsheet."""
        import csv
        abs_csv = os.path.abspath(csv_path)
        if not os.path.exists(abs_csv):
            raise FileNotFoundError(f"CSV file not found: {abs_csv}")

        rows = []
        with open(abs_csv, "r", encoding="utf-8", errors="replace") as f:
            reader = csv.reader(f, delimiter=delimiter)
            for r in reader:
                rows.append(r)

        return ExcelEngine.create_workbook(output_path, rows, sheet_name=sheet_name)

    @staticmethod
    def excel_to_csv(file_path: str, output_csv_path: str, sheet_name: Optional[str] = None, delimiter: str = ",") -> str:
        """Extract a sheet from an Excel file into a CSV file."""
        import csv
        res = ExcelEngine.read_sheet_data(file_path, sheet_name)
        rows = res["rows"]
        abs_out = os.path.abspath(output_csv_path)
        os.makedirs(os.path.dirname(abs_out), exist_ok=True)

        with open(abs_out, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f, delimiter=delimiter)
            for r in rows:
                writer.writerow(r)

        return abs_out

    @staticmethod
    def merge_workbooks(file_paths: List[str], output_path: str) -> str:
        """Merge sheets from multiple .xlsx workbooks into a single workbook."""
        merged_wb = openpyxl.Workbook()
        # Remove default sheet
        merged_wb.remove(merged_wb.active)

        for path in file_paths:
            abs_p = os.path.abspath(path)
            if not os.path.exists(abs_p):
                continue
            src_wb = openpyxl.load_workbook(abs_p, data_only=True)
            base_name = os.path.splitext(os.path.basename(abs_p))[0][:15]

            for s_name in src_wb.sheetnames:
                src_sheet = src_wb[s_name]
                new_title = f"{base_name}_{s_name}"[:31]
                target_sheet = merged_wb.create_sheet(title=new_title)

                for row in src_sheet.iter_rows(values_only=True):
                    target_sheet.append(list(row))

        abs_out = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_out), exist_ok=True)
        merged_wb.save(abs_out)
        return abs_out

    @staticmethod
    def audit_workbook(file_path: str) -> Dict[str, Any]:
        """Deep inspect a workbook headlessly: sheets, dimensions, formulas, file size."""
        abs_p = os.path.abspath(file_path)
        if not os.path.exists(abs_p):
            raise FileNotFoundError(f"File not found: {abs_p}")

        size_kb = round(os.path.getsize(abs_p) / 1024, 2)
        wb = openpyxl.load_workbook(abs_p, data_only=False)

        sheets_meta = []
        total_cells = 0
        total_formulas = 0

        for name in wb.sheetnames:
            ws = wb[name]
            max_r = ws.max_row or 0
            max_c = ws.max_column or 0
            cell_count = max_r * max_c
            total_cells += cell_count

            sheet_formulas = 0
            for row in ws.iter_rows(values_only=False):
                for cell in row:
                    if cell.value and str(cell.value).startswith("="):
                        sheet_formulas += 1
            total_formulas += sheet_formulas

            sheets_meta.append({
                "sheet": name,
                "max_row": max_r,
                "max_column": max_c,
                "total_cells": cell_count,
                "formulas_count": sheet_formulas
            })

        return {
            "file": abs_p,
            "size_kb": size_kb,
            "sheets_count": len(wb.sheetnames),
            "sheets": sheets_meta,
            "total_cells_estimate": total_cells,
            "total_formulas": total_formulas
        }

    @staticmethod
    def batch_replace(file_path: str, replacements: Dict[str, str], output_path: str) -> str:
        """Batch replace placeholders (e.g. {{COMPANY}}) across all sheets in a workbook."""
        abs_in = os.path.abspath(file_path)
        wb = openpyxl.load_workbook(abs_in)

        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        val = cell.value
                        for k, v in replacements.items():
                            if k in val:
                                val = val.replace(k, str(v))
                        cell.value = val

        abs_out = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(abs_out), exist_ok=True)
        wb.save(abs_out)
        return abs_out

    @staticmethod
    def extract_column(file_path: str, column: Union[str, int], sheet_name: Optional[str] = None) -> List[Any]:
        """Extract all values from a specific column as a list."""
        abs_in = os.path.abspath(file_path)
        wb = openpyxl.load_workbook(abs_in, data_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

        col_idx = column if isinstance(column, int) else openpyxl.utils.column_index_from_string(column)
        values = []
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
            values.append(row[0])
        return values

    @staticmethod
    def extract_metadata(file_path: str) -> Dict[str, Any]:
        """Extract core OOXML document metadata (Author, Created, Title, Modified)."""
        abs_in = os.path.abspath(file_path)
        wb = openpyxl.load_workbook(abs_in)
        p = wb.properties
        return {
            "title": p.title,
            "creator": p.creator,
            "subject": p.subject,
            "description": p.description,
            "created": str(p.created) if p.created else None,
            "modified": str(p.modified) if p.modified else None,
            "lastModifiedBy": p.lastModifiedBy
        }

    @staticmethod
    def set_metadata(file_path: str, title: Optional[str] = None, author: Optional[str] = None, subject: Optional[str] = None, output_path: Optional[str] = None) -> str:
        """Set or update document properties/metadata."""
        abs_in = os.path.abspath(file_path)
        wb = openpyxl.load_workbook(abs_in)
        if title is not None:
            wb.properties.title = title
        if author is not None:
            wb.properties.creator = author
            wb.properties.lastModifiedBy = author
        if subject is not None:
            wb.properties.subject = subject

        dest = os.path.abspath(output_path) if output_path else abs_in
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        wb.save(dest)
        return dest

    @staticmethod
    def compare_two_workbooks(file_path_1: str, file_path_2: str) -> Dict[str, Any]:
        """Compare two .xlsx workbooks headlessly and return structural and cell value differences."""
        f1 = os.path.abspath(file_path_1)
        f2 = os.path.abspath(file_path_2)
        wb1 = openpyxl.load_workbook(f1, data_only=True)
        wb2 = openpyxl.load_workbook(f2, data_only=True)

        sheets1 = set(wb1.sheetnames)
        sheets2 = set(wb2.sheetnames)

        diffs = {
            "file_1": f1,
            "file_2": f2,
            "only_in_file_1": list(sheets1 - sheets2),
            "only_in_file_2": list(sheets2 - sheets1),
            "cell_differences": []
        }

        common_sheets = sheets1.intersection(sheets2)
        for s in common_sheets:
            ws1 = wb1[s]
            ws2 = wb2[s]
            max_r = max(ws1.max_row or 0, ws2.max_row or 0)
            max_c = max(ws1.max_column or 0, ws2.max_column or 0)

            for r in range(1, max_r + 1):
                for c in range(1, max_c + 1):
                    v1 = ws1.cell(row=r, column=c).value
                    v2 = ws2.cell(row=r, column=c).value
                    if v1 != v2:
                        diffs["cell_differences"].append({
                            "sheet": s,
                            "cell": f"{get_column_letter(c)}{r}",
                            "file_1_val": str(v1) if v1 is not None else None,
                            "file_2_val": str(v2) if v2 is not None else None
                        })
                        if len(diffs["cell_differences"]) >= 200:
                            diffs["truncated"] = True
                            return diffs

        return diffs

    @staticmethod
    def extract_hyperlinks(file_path: str, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """Extract all embedded URLs and hyperlinks from an .xlsx workbook headlessly."""
        f = os.path.abspath(file_path)
        wb = openpyxl.load_workbook(f)
        sheets = [wb[sheet_name]] if sheet_name and sheet_name in wb.sheetnames else wb.worksheets

        links = []
        for ws in sheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.hyperlink and cell.hyperlink.target:
                        links.append({
                            "sheet": ws.title,
                            "cell": cell.coordinate,
                            "text": str(cell.value) if cell.value else "",
                            "target": str(cell.hyperlink.target)
                        })
        return links

    @staticmethod
    def unhide_all_sheets_headless(file_path: str, output_path: Optional[str] = None) -> str:
        """Unhide all hidden and very hidden worksheets headlessly."""
        f = os.path.abspath(file_path)
        wb = openpyxl.load_workbook(f)
        for ws in wb.worksheets:
            ws.sheet_state = "visible"

        dest = os.path.abspath(output_path) if output_path else f
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        wb.save(dest)
        return dest

