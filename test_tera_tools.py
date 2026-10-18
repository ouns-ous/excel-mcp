"""Test suite for the tera batch of headless tools in excel-mcp.
Tests headless gridlines, zoom scale, auto-fitting column dimensions, and freeze panes.
"""

import os
import openpyxl
from excel_engine import ExcelEngine


def test_tera_headless_tools():
    test_dir = os.path.abspath("test_out")
    os.makedirs(test_dir, exist_ok=True)
    base_file = os.path.join(test_dir, "tera_test.xlsx")

    # 1. Create base file
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "Long Header Company Name"
    ws["B1"] = "Quarterly Revenue Statistics"
    ws["A2"] = "Acme Corp"
    ws["B2"] = 1250000
    wb.save(base_file)

    # 2. Test Gridlines Headless
    ExcelEngine.set_gridlines_headless(base_file, show=False)
    wb_gl = openpyxl.load_workbook(base_file)
    assert wb_gl["Report"].sheet_view.showGridLines is False
    print("[OK] Set gridlines headless passed.")

    # 3. Test Zoom Scale Headless
    ExcelEngine.set_zoom_headless(base_file, zoom_scale=85)
    wb_zm = openpyxl.load_workbook(base_file)
    assert wb_zm["Report"].sheet_view.zoomScale == 85
    print("[OK] Set zoom headless passed.")

    # 4. Test Auto-fit Columns Headless
    ExcelEngine.auto_fit_columns_headless(base_file)
    wb_af = openpyxl.load_workbook(base_file)
    # Column A width should be adjusted larger than default 10
    assert wb_af["Report"].column_dimensions["A"].width > 15
    print("[OK] Auto-fit columns headless passed.")

    # 5. Test Freeze Panes Headless
    ExcelEngine.freeze_panes_headless(base_file, cell="A2")
    wb_fp = openpyxl.load_workbook(base_file)
    assert wb_fp["Report"].freeze_panes == "A2"
    print("[OK] Freeze panes headless passed.")


if __name__ == "__main__":
    test_tera_headless_tools()
    print("[ALL TERA HEADLESS TESTS PASSED!]")
