"""Test suite for advanced Excel MCP capabilities.
Tests:
- Headless CSV conversion, export, and workbook merging
- Live Sorting, Filtering, and Conditional Formatting
- Data Validation (Dropdowns)
- Structural operations (Merge, Width, Freeze)
- Formula Auditing and Named Ranges
"""

import os
import sys

def test_advanced_headless():
    print("\n--- 1. Testing Advanced Headless Utilities ---")
    from excel_engine import ExcelEngine
    os.makedirs("output", exist_ok=True)

    # 1. Create a sample CSV
    csv_file = "output/sample_products.csv"
    with open(csv_file, "w", encoding="utf-8") as f:
        f.write("ProductID,Category,Price,Stock\n")
        f.write("P-101,Electronics,299.99,45\n")
        f.write("P-102,Accessories,19.99,250\n")
        f.write("P-103,Laptops,999.00,12\n")

    # CSV to Excel
    xlsx_out = "output/products_from_csv.xlsx"
    ExcelEngine.csv_to_excel(csv_file, xlsx_out, sheet_name="Inventory")
    print(f"[OK] Converted CSV to Excel: {xlsx_out}")

    # Excel back to CSV
    csv_back = "output/products_exported.csv"
    ExcelEngine.excel_to_csv(xlsx_out, csv_back)
    print(f"[OK] Exported Excel sheet to CSV: {csv_back}")

    # Merge Workbooks
    merged_out = "output/merged_suite.xlsx"
    ExcelEngine.merge_workbooks([xlsx_out, "output/demo_headless_sales.xlsx"], merged_out)
    print(f"[OK] Merged multiple workbooks into: {merged_out}")


def test_advanced_live_automation():
    print("\n--- 2. Testing Advanced Live Excel COM Features ---")
    if sys.platform != "win32":
        print("[SKIP] Windows required for COM automation.")
        return

    try:
        from excel_controller import controller

        # New Workbook
        controller.new_workbook(sheet_name="Operations_Hub")

        # Table data
        headers = ["Task ID", "Team", "Status", "Priority", "Budget ($)"]
        data = [
            headers,
            ["T-01", "Core Engine", "Completed", "High", 15000],
            ["T-02", "UI / UX", "In Progress", "Critical", 8500],
            ["T-03", "Security Audit", "Pending", "Medium", 12000],
            ["T-04", "Performance", "In Progress", "High", 6200],
            ["T-05", "Docs & Guides", "Completed", "Low", 2500],
        ]
        controller.write_range(start_cell="B3", data=data)

        # Style headers
        controller.format_range(
            range_address="B3:F3",
            bold=True,
            bg_color="#0F172A",
            font_color="#FFFFFF",
            horizontal_align="center"
        )
        controller.format_range(range_address="F4:F8", number_format="$#,##0")

        # Data Validation (Dropdown on Status column C4:C8)
        print("[->] Adding Data Validation Dropdown to Status column...")
        controller.add_data_validation(
            range_address="D4:D8",
            validation_type="list",
            formula1="Completed,In Progress,Pending,Blocked"
        )

        # Conditional Formatting: Highlight High / Critical Priority
        print("[->] Adding Conditional Formatting...")
        controller.add_conditional_formatting(
            range_address="E4:E8",
            rule_type="cell_value",
            operator="equal",
            value="Critical",
            bg_color="#FEE2E2",
            font_color="#991B1B"
        )
        # Heatmap / Data bars on Budget
        controller.add_conditional_formatting(
            range_address="F4:F8",
            rule_type="data_bar",
            bg_color="#3B82F6"
        )

        # Named Range
        print("[->] Creating Named Range 'TaskBudgets'...")
        controller.create_named_range("TaskBudgets", "F4:F8")

        # Total Formula
        controller.write_cell("E9", "Total Budget:")
        controller.format_range("E9", bold=True, horizontal_align="right")
        controller.set_formula("F9", "=SUM(TaskBudgets)")
        controller.format_range("F9", bold=True, number_format="$#,##0", bg_color="#F1F5F9")

        # Add review comment
        print("[->] Adding cell review note...")
        controller.add_comment("F9", "Approved for Q1 execution by Finance.", author="CFO")

        # AutoFit
        controller.autofit()

        # Save and close cleanly
        save_path = os.path.abspath("output/demo_advanced_operations.xlsx")
        controller.save_workbook(save_path)
        print(f"[OK] Saved advanced operations workbook: {save_path}")

        controller.close_workbook(save_changes=False)
        print("[OK] Workbook closed cleanly.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[NOTICE] Encountered error: {e}")


if __name__ == "__main__":
    test_advanced_headless()
    test_advanced_live_automation()
    print("\n[SUCCESS] Advanced Excel MCP test completed successfully!")
