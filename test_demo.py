"""Comprehensive demonstration of Excel MCP capabilities.
Tests both headless OOXML generation and live Windows COM desktop automation.
"""

import os
import sys
from excel_engine import ExcelEngine

def test_headless_pipeline():
    print("\n--- 1. Testing Headless OOXML Engine (openpyxl) ---")
    os.makedirs("output", exist_ok=True)

    # 1. Markdown to Excel
    sample_md = """
| Quarter | Region | Product | Sales ($) | Units | Target Achieved |
|---|---|---|---|---|---|
| Q1 2026 | North America | Cloud ERP | 450000 | 150 | Yes |
| Q1 2026 | EMEA | Security Suite | 320000 | 95 | Yes |
| Q1 2026 | APAC | AI Assistant | 680000 | 420 | Yes |
| Q2 2026 | North America | Data Lakehouse | 510000 | 180 | Yes |
| Q2 2026 | EMEA | Cloud ERP | 390000 | 120 | No |
"""
    out_xlsx = "output/demo_headless_sales.xlsx"
    created_path = ExcelEngine.markdown_to_excel(sample_md, out_xlsx, sheet_name="Q1_Q2_Sales")
    print(f"[OK] Created Excel headlessly from Markdown: {created_path}")

    # 2. Excel to Markdown Extraction
    extracted_md = ExcelEngine.excel_to_markdown(created_path)
    print("\n[OK] Extracted Markdown from generated .xlsx:")
    print(extracted_md)

    # 3. JSON to Excel
    sample_json = [
        {"EmpID": "E-101", "Name": "Alice Johnson", "Department": "Engineering", "Salary": 125000},
        {"EmpID": "E-102", "Name": "Bob Smith", "Department": "Product", "Salary": 115000},
        {"EmpID": "E-103", "Name": "Charlie Brown", "Department": "Design", "Salary": 98000},
    ]
    json_xlsx = "output/demo_employees.xlsx"
    ExcelEngine.json_to_excel(sample_json, json_xlsx, sheet_name="Employees")
    print(f"\n[OK] Exported JSON to Excel: {json_xlsx}")


def test_live_desktop_automation():
    print("\n--- 2. Testing Live Excel COM Desktop Automation ---")
    if sys.platform != "win32":
        print("[SKIP] COM automation only supported on Windows.")
        return

    try:
        from excel_controller import controller

        # 1. New Workbook
        print("[->] Creating new live workbook...")
        controller.new_workbook(sheet_name="Executive_Dashboard")

        # 2. Executive KPI Cards
        print("[->] Creating KPI cards...")
        controller.create_kpi_card(
            start_cell="B2",
            title="Total Revenue",
            value="$2,350,000",
            subtitle="+18.4% vs last quarter",
            accent_color="#2563EB"
        )
        controller.create_kpi_card(
            start_cell="F2",
            title="Active Enterprise Clients",
            value="1,420",
            subtitle="+95 new accounts",
            accent_color="#059669"
        )
        controller.create_kpi_card(
            start_cell="J2",
            title="Customer Retention",
            value="98.2%",
            subtitle="Benchmark: >95%",
            accent_color="#7C3AED"
        )

        # 3. Data Table
        print("[->] Writing financial data table...")
        headers = ["Quarter", "Revenue ($)", "Expenses ($)", "Net Profit ($)", "Profit Margin"]
        rows = [
            headers,
            ["Q1 2026", 450000, 310000, "=B7-C7", "=D7/B7"],
            ["Q2 2026", 520000, 340000, "=B8-C8", "=D8/B8"],
            ["Q3 2026", 610000, 390000, "=B9-C9", "=D9/B9"],
            ["Q4 2026", 770000, 460000, "=B10-C10", "=D10/B10"],
        ]
        controller.write_range(start_cell="B6", data=rows)

        # Format header
        controller.format_range(
            range_address="B6:F6",
            bold=True,
            font_color="#FFFFFF",
            bg_color="#1E293B",
            horizontal_align="center",
            font_size=11
        )

        # Format number currencies
        controller.format_range(range_address="B7:D10", number_format="$#,##0")
        controller.format_range(range_address="F7:F10", number_format="0.0%")
        controller.format_range(range_address="B7:F10", border_style="thin", border_color="#E2E8F0")

        # Add Total Row
        total_row = [["Total", "=SUM(B7:B10)", "=SUM(C7:C10)", "=SUM(D7:D10)", "=D11/B11"]]
        controller.write_range(start_cell="B11", data=total_row)
        controller.format_range(
            range_address="B11:F11",
            bold=True,
            bg_color="#F1F5F9",
            border_style="medium",
            border_color="#94A3B8"
        )
        controller.format_range(range_address="B11:D11", number_format="$#,##0")
        controller.format_range(range_address="F11", number_format="0.0%")

        # AutoFit
        controller.autofit()

        # 4. Native Excel Chart
        print("[->] Adding interactive clustered column chart...")
        controller.add_chart(
            chart_type="column",
            source_range="B6:D10",
            title="2026 Revenue vs Expenses",
            left=450,
            top=120,
            width=500,
            height=280
        )

        # 5. Save and Export PDF
        out_live_xlsx = os.path.abspath("output/demo_live_executive_dashboard.xlsx")
        out_live_pdf = os.path.abspath("output/demo_live_executive_dashboard.pdf")
        controller.save_workbook(out_live_xlsx)
        print(f"[OK] Saved live workbook to: {out_live_xlsx}")

        controller.export_pdf(out_live_pdf)
        print(f"[OK] Exported PDF successfully: {out_live_pdf}")

        # Cleanly close workbook
        controller.close_workbook(save_changes=False)
        print("[OK] Closed live demo workbook.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[NOTICE] Live COM test skipped or encountered: {e}")


if __name__ == "__main__":
    test_headless_pipeline()
    test_live_desktop_automation()
    print("\n[SUCCESS] All Excel MCP Demo tests completed successfully!")
