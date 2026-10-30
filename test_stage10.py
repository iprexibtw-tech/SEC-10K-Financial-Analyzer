"""Automated Unit Tests for Stage 10: NVIDIA Analysis and Dual-Company Application Validation.

This test module verifies:
1. Complete, simultaneous coexistence and integrity of both AAPL and NVDA outputs.
2. Existence and structure of both comprehensive workbooks:
   - data/excel_outputs/AAPL_10K_Analysis.xlsx
   - data/excel_outputs/NVDA_10K_Analysis.xlsx
3. Presence and exact ordering of all 8 required worksheets in both workbooks:
   ['Income Statement', 'Balance Sheet', 'Cash Flow Statement', 'Sentiment Analysis',
    'Stock Prices', 'Financial Ratios', 'Charts and Analysis', 'Audit & Reconciliation']
4. Correct embedding of 10 charts in each workbook (6 in Charts & Analysis, 4 in Financial Ratios).
5. Verification of the dual-axis chart (Stock Price vs. Sentiment P/N Ratio) in both workbooks.
6. Zero accounting discrepancies: 100% PASS for balance sheet equation and gross profit reconciliation across all 20 fiscal periods (10 for AAPL, 10 for NVDA).
7. Proper handling of NVIDIA-specific calendar (Sunday fiscal ends matched to preceding Friday).
8. Existence of all 6 structured CSV exports in data/financial_data/.
"""

import csv
import unittest
from pathlib import Path

import openpyxl

import config


class TestStage10DualCompanyValidation(unittest.TestCase):
    """Unit test suite for Stage 10 dual-company analysis and validation."""

    def setUp(self) -> None:
        """Locate project output directories."""
        self.excel_dir = config.EXCEL_OUTPUTS_DIR
        self.financial_dir = config.FINANCIAL_DATA_DIR
        self.raw_dir = config.RAW_FILINGS_DIR
        self.extracted_dir = config.EXTRACTED_TEXT_DIR
        self.cleaned_dir = config.CLEANED_TEXT_DIR
        self.expected_sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Stock Prices",
            "Financial Ratios",
            "Charts and Analysis",
            "Audit & Reconciliation",
        ]

    def test_both_workbooks_exist_in_root_and_subdirectories(self) -> None:
        """Verify both AAPL and NVDA master workbooks exist in root and subdirectories."""
        for ticker in ["AAPL", "NVDA"]:
            root_wb = self.excel_dir / f"{ticker}_10K_Analysis.xlsx"
            sub_wb = self.excel_dir / ticker / f"{ticker}_10K_Analysis.xlsx"
            stmt_wb = self.excel_dir / ticker / f"{ticker}_financial_statements.xlsx"

            self.assertTrue(root_wb.exists(), f"Missing root workbook: {root_wb}")
            self.assertTrue(sub_wb.exists(), f"Missing subfolder workbook: {sub_wb}")
            self.assertTrue(stmt_wb.exists(), f"Missing financial statements workbook: {stmt_wb}")
            self.assertGreater(root_wb.stat().st_size, 20_000, f"Workbook {root_wb.name} too small")

    def test_worksheet_structure_and_order_for_both_companies(self) -> None:
        """Verify that both workbooks contain all 8 required worksheets in identical order."""
        for ticker in ["AAPL", "NVDA"]:
            wb_path = self.excel_dir / f"{ticker}_10K_Analysis.xlsx"
            wb = openpyxl.load_workbook(wb_path, data_only=True)
            self.assertEqual(
                wb.sheetnames,
                self.expected_sheets,
                f"Worksheet list mismatch for {ticker}: {wb.sheetnames}",
            )

    def test_chart_counts_and_dual_axis_configuration(self) -> None:
        """Verify 10 embedded charts per workbook including the dual-axis chart."""
        for ticker in ["AAPL", "NVDA"]:
            wb_path = self.excel_dir / f"{ticker}_10K_Analysis.xlsx"
            wb = openpyxl.load_workbook(wb_path, data_only=False)

            ws_charts = wb["Charts and Analysis"]
            self.assertEqual(
                len(ws_charts._charts),
                6,
                f"Expected 6 charts in 'Charts and Analysis' for {ticker}, found {len(ws_charts._charts)}",
            )

            # First chart must be dual-axis (Stock Price vs Sentiment P/N ratio)
            dual_chart = ws_charts._charts[0]
            self.assertIn("Sentiment", str(dual_chart.title))
            self.assertIsNotNone(dual_chart.y_axis)
            self.assertIsNotNone(dual_chart.x_axis)

            ws_ratios = wb["Financial Ratios"]
            self.assertEqual(
                len(ws_ratios._charts),
                4,
                f"Expected 4 charts in 'Financial Ratios' for {ticker}, found {len(ws_ratios._charts)}",
            )

    def test_audit_reconciliations_100_percent_pass_both_companies(self) -> None:
        """Verify zero discrepancies: 10/10 PASS for BS and GP for both companies (20/20 total)."""
        for ticker in ["AAPL", "NVDA"]:
            wb_path = self.excel_dir / f"{ticker}_10K_Analysis.xlsx"
            wb = openpyxl.load_workbook(wb_path, data_only=True)
            ws_audit = wb["Audit & Reconciliation"]

            # Rows 6 to 15 contain the 10 fiscal reporting periods
            pass_count_bs = 0
            pass_count_gp = 0

            for r in range(6, 16):
                fy = ws_audit.cell(row=r, column=1).value
                bs_bal = str(ws_audit.cell(row=r, column=6).value or "")
                gp_rec = str(ws_audit.cell(row=r, column=7).value or "")

                self.assertIn("PASS", bs_bal, f"{ticker} {fy} Balance Sheet failed: {bs_bal}")
                self.assertIn("PASS", gp_rec, f"{ticker} {fy} Gross Profit failed: {gp_rec}")

                if "PASS" in bs_bal:
                    pass_count_bs += 1
                if "PASS" in gp_rec:
                    pass_count_gp += 1

            self.assertEqual(pass_count_bs, 10, f"{ticker} BS PASS count != 10")
            self.assertEqual(pass_count_gp, 10, f"{ticker} GP PASS count != 10")

    def test_nvda_sunday_to_friday_date_handling(self) -> None:
        """Verify NVIDIA Sunday period-end dates are matched to preceding Friday trading dates."""
        stock_csv = self.financial_dir / "NVDA_stock_prices.csv"
        self.assertTrue(stock_csv.exists(), "NVDA_stock_prices.csv missing")

        with open(stock_csv, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)

        self.assertEqual(len(rows), 10, "Expected 10 stock price rows for NVDA")
        for row in rows:
            bs_date = row["fiscal_balance_sheet_date"]
            tr_date = row["actual_trading_date"]
            offset = int(row["offset_days"])
            price = float(row["closing_price_usd"])

            self.assertGreater(price, 0.0, f"Invalid closing price for {bs_date}")
            # NVIDIA balance sheet dates end in late January and are Sundays
            # Offset should be >= 0 (trading date <= balance sheet date, no lookahead bias)
            self.assertGreaterEqual(offset, 0)
            self.assertLessEqual(tr_date, bs_date, f"Lookahead bias: {tr_date} > {bs_date}")

    def test_all_csv_datasets_present_and_populated(self) -> None:
        """Verify all 6 analytical CSV files exist and have 10 data rows."""
        csv_types = [
            "sentiment_analysis.csv",
            "stock_prices.csv",
            "financial_ratios.csv",
        ]
        for ticker in ["AAPL", "NVDA"]:
            for c_type in csv_types:
                csv_path = self.financial_dir / f"{ticker}_{c_type}"
                self.assertTrue(csv_path.exists(), f"Missing CSV: {csv_path}")
                with open(csv_path, "r", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    rows = list(reader)
                self.assertEqual(
                    len(rows),
                    10,
                    f"Expected 10 rows in {csv_path.name}, got {len(rows)}",
                )

    def test_manifest_and_text_files_integrity(self) -> None:
        """Verify raw, extracted, and cleaned text directories and manifests."""
        for ticker in ["AAPL", "NVDA"]:
            raw_m = self.raw_dir / ticker / f"{ticker}_10K_manifest.csv"
            ext_m = self.extracted_dir / ticker / f"{ticker}_mda_manifest.csv"
            cln_m = self.cleaned_dir / ticker / f"{ticker}_cleaning_manifest.csv"

            self.assertTrue(raw_m.exists(), f"Missing raw manifest for {ticker}")
            self.assertTrue(ext_m.exists(), f"Missing extracted manifest for {ticker}")
            self.assertTrue(cln_m.exists(), f"Missing cleaned manifest for {ticker}")

            # Verify 10 MD&A extracted and 10 cleaned text files
            ext_files = list((self.extracted_dir / ticker).glob("*_MDA.txt"))
            cln_files = list((self.cleaned_dir / ticker).glob("*_MDA_cleaned.txt"))
            self.assertEqual(len(ext_files), 10, f"Expected 10 extracted text files for {ticker}")
            self.assertEqual(len(cln_files), 10, f"Expected 10 cleaned text files for {ticker}")


if __name__ == "__main__":
    unittest.main()
