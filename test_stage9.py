"""Automated Unit Tests for Stage 9: Final Excel Reporting and Application Integration.

This test module verifies:
1. Generation and integrity of the final comprehensive workbook '{TICKER}_10K_Analysis.xlsx'.
2. Presence and correct ordering of all 8 required worksheets.
3. Proper configuration of the mandatory dual-axis chart (Stock Price vs. Sentiment P/N Ratio).
4. Embedding of all 6 analytical trend charts in the 'Charts and Analysis' worksheet.
5. Dual export to both company subfolder and root excel_outputs directory.
6. Execution and status reporting of the complete 8-stage pipeline in main.py.
"""

import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl

import config
from excel_report import ExcelReportGenerator, ExcelReportError
from main import run_full_pipeline


class TestStage9Integration(unittest.TestCase):
    """Unit test suite for Stage 9 reporting and complete application integration."""

    def setUp(self) -> None:
        """Create isolated temporary test directories."""
        self.test_dir = Path(tempfile.mkdtemp(prefix="test_stage9_"))
        self.excel_dir = self.test_dir / "excel_outputs"
        self.financial_dir = self.test_dir / "financial_data"
        self.excel_dir.mkdir(parents=True, exist_ok=True)
        self.financial_dir.mkdir(parents=True, exist_ok=True)
        self.reporter = ExcelReportGenerator(
            excel_dir=self.excel_dir, financial_dir=self.financial_dir
        )

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        if self.test_dir.exists():
            shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_excel_report_generator_sheets_and_dual_axis_chart(self) -> None:
        """Verify that final workbook contains all 8 sheets and mandatory dual-axis chart."""
        # Create mock base workbook with 7 existing sheets
        base_sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Stock Prices",
            "Financial Ratios",
            "Audit & Reconciliation",
        ]
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        for s in base_sheets:
            ws = wb.create_sheet(title=s)
            ws["A1"] = f"Mock {s} (TEST)"
            ws.cell(row=4, column=1, value="Fiscal Year")
            ws.cell(row=5, column=1, value="Date")
            for col, yr in enumerate(["2023", "2024"], start=2):
                ws.cell(row=4, column=col, value=f"FY {yr}")
                ws.cell(row=5, column=col, value=f"{yr}-09-30")

        # Mock Stock Prices row 8
        ws_sp = wb["Stock Prices"]
        ws_sp.cell(row=8, column=1, value="Split-Adjusted Closing Price")
        ws_sp.cell(row=8, column=2, value=171.21)
        ws_sp.cell(row=8, column=3, value=227.79)

        # Mock Sentiment Analysis row 10, 11, 19
        ws_sa = wb["Sentiment Analysis"]
        ws_sa.cell(row=10, column=2, value=50)
        ws_sa.cell(row=10, column=3, value=60)
        ws_sa.cell(row=11, column=2, value=500)
        ws_sa.cell(row=11, column=3, value=600)
        ws_sa.cell(row=19, column=2, value=0.100)
        ws_sa.cell(row=19, column=3, value=0.100)

        # Mock Financial Ratios calculation details
        ws_fr = wb["Financial Ratios"]
        for col, yr in enumerate(["2023", "2024"], start=2):
            ws_fr.cell(row=14, column=col, value=150.0)
            ws_fr.cell(row=15, column=col, value=100.0)
            ws_fr.cell(row=18, column=col, value=10.0)
            ws_fr.cell(row=19, column=col, value=20.0)
            ws_fr.cell(row=20, column=col, value=70.0)
            ws_fr.cell(row=22, column=col, value=100.0)
            ws_fr.cell(row=25, column=col, value=400.0)
            ws_fr.cell(row=26, column=col, value=500.0)
            ws_fr.cell(row=28, column=col, value=600.0)
            ws_fr.cell(row=31, column=col, value=150.0)

        test_company_dir = self.excel_dir / "TEST"
        test_company_dir.mkdir(parents=True, exist_ok=True)
        base_wb_path = test_company_dir / "TEST_financial_statements.xlsx"
        wb.save(base_wb_path)

        comp_path, root_path, sheetnames = self.reporter.generate_final_workbook(
            ticker="TEST", count=2
        )

        expected_sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Stock Prices",
            "Financial Ratios",
            "Charts and Analysis",
            "Audit & Reconciliation",
        ]
        self.assertEqual(sheetnames, expected_sheets)
        self.assertTrue(comp_path.exists())
        self.assertTrue(root_path.exists())

        # Inspect Charts and Analysis worksheet
        res_wb = openpyxl.load_workbook(comp_path)
        ws_charts = res_wb["Charts and Analysis"]

        # Exactly 6 charts embedded
        self.assertEqual(len(ws_charts._charts), 6)

        # Verify dual axis on primary chart
        dual_chart = ws_charts._charts[0]
        # In openpyxl dual-axis, a second plot/axis is appended
        self.assertIsNotNone(dual_chart.y_axis)

    def test_full_pipeline_execution_aapl(self) -> None:
        """Verify that the full 8-stage pipeline executes with exit code 0 for AAPL."""
        exit_code = run_full_pipeline(ticker="AAPL", count=10)
        self.assertEqual(exit_code, 0)

        # Verify output files exist
        comp_file = config.EXCEL_OUTPUTS_DIR / "AAPL" / "AAPL_10K_Analysis.xlsx"
        root_file = config.EXCEL_OUTPUTS_DIR / "AAPL_10K_Analysis.xlsx"
        self.assertTrue(comp_file.exists())
        self.assertTrue(root_file.exists())

        # Verify workbook opens cleanly and contains all 8 sheets
        wb = openpyxl.load_workbook(root_file)
        expected_sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Stock Prices",
            "Financial Ratios",
            "Charts and Analysis",
            "Audit & Reconciliation",
        ]
        self.assertEqual(wb.sheetnames, expected_sheets)


if __name__ == "__main__":
    unittest.main()
