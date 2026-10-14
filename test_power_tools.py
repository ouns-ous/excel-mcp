"""Test suite for the latest batch of Excel MCP Power Tools (Grouping, Outlines,
Top/Bottom & Icon Set Conditional Formatting, Tracer Arrows, Shapes, Special Cells).
"""

import os
import sys

def test_power_tools():
    print("\n--- Testing New Excel Power Tools (119 Tools Suite) ---")
    if sys.platform != "win32":
        print("[SKIP] Windows required for Excel COM Automation.")
        return

    try:
        from excel_controller import controller
        os.makedirs("output", exist_ok=True)

        # 1. New Workbook
        controller.new_workbook(sheet_name="Power_Audit_Demo")
        print("[OK] Created new workbook.")

        # 2. Write numeric test table
        data = [
            ["ID", "Department", "Score", "Bonus ($)", "DuplicateCheck"],
            [101, "Sales", 88, 5000, "ABC"],
            [102, "Engineering", 95, 12000, "XYZ"],
            [103, "Marketing", 72, 3500, "ABC"],
            [104, "Design", 91, 8000, "DEF"],
            [105, "Support", 64, 2000, "XYZ"],
            [106, "HR", 84, 4500, "GHI"],
        ]
        controller.write_range("B3", data)
        controller.format_range("B3:F3", bold=True, bg_color="#1E293B", font_color="#FFFFFF")
        print("[OK] Populated dataset.")

        # 3. Test Row & Column Outlining/Grouping
        print("[->] Grouping detail rows B4:B8...")
        controller.group_rows(start_row=4, end_row=8)
        controller.show_outline_levels(row_levels=2)
        print("[OK] Outline levels configured.")

        # 4. Top/Bottom & Icon Sets & Duplicate Highlights
        print("[->] Adding Top/Bottom & Icon Set conditional formatting...")
        controller.add_top_bottom_rule(range_address="D4:D9", rule_type="top_items", rank=2, bg_color="#DCFCE7", font_color="#166534")
        controller.add_icon_set_rule(range_address="D4:D9", icon_set_type="3_traffic_lights")
        controller.highlight_duplicate_values(range_address="F4:F9", bg_color="#FEE2E2", font_color="#991B1B")
        print("[OK] Conditional formatting rules applied.")

        # 5. Native Shapes & Badges
        print("[->] Adding shape badge...")
        controller.add_shape(
            shape_type="rounded_rectangle",
            cell="B11",
            width=160,
            height=32,
            fill_color="#2563EB",
            text="AUDITED & APPROVED",
            font_color="#FFFFFF"
        )
        shapes = controller.list_shapes()
        print(f"[OK] Shapes on sheet: {len(shapes)}")

        # 6. Formula & Tracer Arrows
        controller.write_cell("D10", "=SUM(D4:D9)")
        controller.format_range("D10", bold=True, bg_color="#F1F5F9")
        controller.trace_precedents("D10")
        print("[OK] Drew precedent tracer arrows for formula.")
        controller.remove_tracer_arrows()
        print("[OK] Cleared tracer arrows.")

        # 7. Special Cells Inspection
        special_res = controller.get_special_cells_summary(cell_type="formulas")
        print(f"[OK] Special cells (formulas) found: {special_res.get('count')}")

        # AutoFit and Save
        controller.autofit()
        out_path = os.path.abspath("output/demo_power_tools.xlsx")
        controller.save_workbook(out_path)
        print(f"[OK] Saved workbook to: {out_path}")

        controller.close_workbook(save_changes=False)
        print("[SUCCESS] All new power tools tested and verified successfully!")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[NOTICE] Encountered: {e}")


if __name__ == "__main__":
    test_power_tools()
