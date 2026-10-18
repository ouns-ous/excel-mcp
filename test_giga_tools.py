"""Test suite for the giga batch of headless tools in excel-mcp.
Tests defined names reading, row insertions/deletions, column insertions/deletions.
"""

import os
import openpyxl
from excel_engine import ExcelEngine


def test_giga_headless_tools():
    test_dir = os.path.abspath("test_out")
    os.makedirs(test_dir, exist_ok=True)
    base_file = os.path.join(test_dir, "giga_test.xlsx")

    # 1. Create base file with 3 rows and 3 columns
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DataSheet"
    ws["A1"] = "ColA"
    ws["B1"] = "ColB"
    ws["A2"] = "Row2Val"
    ws["B2"] = 100
    # Add a defined name
    from openpyxl.workbook.defined_name import DefinedName
    new_dn = DefinedName("TaxRate", attr_text="DataSheet!$B$2")
    wb.defined_names.add(new_dn)
    wb.save(base_file)


    # 2. Test Get Defined Names Headless
    names = ExcelEngine.get_defined_names_headless(base_file)
    assert any(n["name"] == "TaxRate" for n in names)
    print("[OK] Get defined names headless passed.")

    # 3. Test Insert Rows Headless
    ExcelEngine.insert_rows_headless(base_file, row_idx=2, amount=2)
    wb_ins_r = openpyxl.load_workbook(base_file)
    ws_r = wb_ins_r["DataSheet"]
    # Previous row 2 should now be pushed down to row 4
    assert ws_r["B4"].value == 100
    print("[OK] Insert rows headless passed.")

    # 4. Test Delete Rows Headless
    ExcelEngine.delete_rows_headless(base_file, row_idx=2, amount=2)
    wb_del_r = openpyxl.load_workbook(base_file)
    ws_del_r = wb_del_r["DataSheet"]
    assert ws_del_r["B2"].value == 100
    print("[OK] Delete rows headless passed.")

    # 5. Test Insert Columns Headless
    ExcelEngine.insert_cols_headless(base_file, col_idx=2, amount=1)
    wb_ins_c = openpyxl.load_workbook(base_file)
    ws_c = wb_ins_c["DataSheet"]
    # Previous col B should now be pushed to col C
    assert ws_c["C2"].value == 100
    print("[OK] Insert columns headless passed.")

    # 6. Test Delete Columns Headless
    ExcelEngine.delete_cols_headless(base_file, col_idx=2, amount=1)
    wb_del_c = openpyxl.load_workbook(base_file)
    ws_del_c = wb_del_c["DataSheet"]
    assert ws_del_c["B2"].value == 100
    print("[OK] Delete columns headless passed.")


if __name__ == "__main__":
    test_giga_headless_tools()
    print("[ALL GIGA HEADLESS TESTS PASSED!]")
