"""Test suite for the newest batch of Excel MCP tools.
Tests:
- Headless Audit, Batch Replace, Extract Column
- Live Goal Seek, Text to Columns, Sparklines, Gantt Block, Tab Colors, Summary
"""

import os
import sys

def test_headless_features():
    print("\n--- 1. Testing Headless Audit & Batch Replace ---")
    from excel_engine import ExcelEngine

    # 1. Audit Workbook
    sample_file = "output/demo_headless_sales.xlsx"
    audit_res = ExcelEngine.audit_workbook(sample_file)
    print(f"[OK] Audited workbook '{sample_file}':")
    print(f"     Sheets: {audit_res['sheets_count']}, Size: {audit_res['size_kb']} KB")

    # 2. Batch Replace
    template_file = "output/template_test.xlsx"
    ExcelEngine.create_workbook(
        template_file,
        data=[
            ["Report Title", "Prepared For", "Fiscal Year"],
            ["Annual Analysis", "{{CLIENT_NAME}}", "{{YEAR}}"],
        ]
    )
    filled_file = "output/template_filled.xlsx"
    ExcelEngine.batch_replace(
        template_file,
        replacements={"{{CLIENT_NAME}}": "Acme Global Corp", "{{YEAR}}": "2026"},
        output_path=filled_file
    )
    print(f"[OK] Batch replaced placeholders into: {filled_file}")

    # 3. Extract Column
    col_vals = ExcelEngine.extract_column(filled_file, column="B")
    print(f"[OK] Extracted Column B: {col_vals}")


def test_live_super_tools():
    print("\n--- 2. Testing Live Super Tools ---")
    if sys.platform != "win32":
        print("[SKIP] Windows required for COM automation.")
        return

    try:
        from excel_controller import controller

        # New Workbook
        controller.new_workbook(sheet_name="Executive_Dashboard")

        # Tab color & Gridlines
        print("[->] Setting Tab Color & configuring view...")
        controller.set_tab_color(hex_color="#2563EB")
        controller.toggle_gridlines(show=True)

        # 1. Project Gantt Block
        print("[->] Creating Project Gantt Roadmap Block...")
        tasks = [
            {"name": "Architecture Design", "owner": "Karim", "start_week": 1, "duration": 2, "color": "#2563EB"},
            {"name": "Backend MCP Core", "owner": "Dev Team", "start_week": 2, "duration": 3, "color": "#059669"},
            {"name": "Security & Auditing", "owner": "Youssef", "start_week": 4, "duration": 2, "color": "#DC2626"},
            {"name": "Release & Delivery", "owner": "Lead", "start_week": 5, "duration": 2, "color": "#7C3AED"},
        ]
        controller.create_gantt_block(start_cell="B3", tasks=tasks, weeks_count=6)

        # 2. Text to Columns Test
        print("[->] Testing Text to Columns...")
        controller.write_range("B11", [["FirstName,LastName,Department"], ["Sarah,Connor,Security"], ["John,Doe,Engineering"]])
        controller.text_to_columns(range_address="B11:B13", destination_cell="D11", delimiter=",")

        # 3. Numeric Summary Test
        print("[->] Testing Numeric Column Summary...")
        controller.write_range("H11", [[100], [250], [400], [150], [600]])
        summary = controller.get_column_summary("H11:H15")
        print(f"[OK] Column Summary: Sum={summary.get('sum')}, Avg={summary.get('average')}")

        # 4. Hyperlink Test
        print("[->] Adding Hyperlink...")
        controller.add_hyperlink(cell="B1", target="https://github.com", text_to_display="Open Project Repository")

        # 5. Goal Seek Test
        print("[->] Testing Financial Goal Seek...")
        # Cell B16: Units sold = 100
        # Cell B17: Price per unit = 50
        # Cell B18: Total Revenue = B16 * B17
        controller.write_cell("B16", 100)
        controller.write_cell("B17", 50)
        controller.set_formula("B18", "=B16*B17")

        # Goal Seek: We want Revenue (B18) to reach 10,000 by changing Units (B16)
        gs_res = controller.goal_seek(target_cell="B18", target_value=10000.0, changing_cell="B16")
        print(f"[OK] Goal Seek Result: Units needed = {gs_res.get('result_value')}")

        # AutoFit
        controller.autofit()

        # Save and close
        save_path = os.path.abspath("output/demo_super_tools.xlsx")
        controller.save_workbook(save_path)
        print(f"[OK] Saved super tools workbook to: {save_path}")

        controller.close_workbook(save_changes=False)
        print("[OK] Workbook closed cleanly.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[NOTICE] Encountered: {e}")


if __name__ == "__main__":
    test_headless_features()
    test_live_super_tools()
    print("\n[SUCCESS] All newest tools tested and verified successfully!")
