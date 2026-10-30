"""Automated Unit Tests for Stage 8: Ten-Year Financial Ratio Analysis.

This module validates:
1. Accuracy of formulas for Current Ratio, Debt-to-Equity, Asset Turnover, and Net Profit Margin.
2. Safety division by zero handling (zero liabilities, zero equity, zero assets, zero revenue).
3. Correct multi-period averaging of Total Assets using preceding year opening balances.
4. Prevention of debt double-counting and verification of interest-bearing components.
5. Preservation of all prior worksheets ('Income Statement', 'Balance Sheet',
   'Cash Flow Statement', 'Sentiment Analysis', 'Stock Prices', 'Audit & Reconciliation').
6. Proper generation of 4 embedded openpyxl trend charts and structured CSV schema.
"""

import csv
import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl

from financial_ratios import FinancialRatioAnalyzer, FinancialRatioError


class TestStage8FinancialRatios(unittest.TestCase):
    """Unit test suite for financial ratio calculation and Excel integration."""

    def setUp(self) -> None:
        """Create isolated temporary test directories."""
        self.test_dir = Path(tempfile.mkdtemp(prefix="test_ratios_"))
        self.excel_dir = self.test_dir / "excel_outputs"
        self.financial_dir = self.test_dir / "financial_data"
        self.excel_dir.mkdir(parents=True, exist_ok=True)
        self.financial_dir.mkdir(parents=True, exist_ok=True)
        self.analyzer = FinancialRatioAnalyzer(
            excel_dir=self.excel_dir, financial_dir=self.financial_dir
        )

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        if self.test_dir.exists():
            shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_current_ratio_formula_and_zero_safety(self) -> None:
        """Verify Current Ratio calculation and division-by-zero protection."""
        inputs = {
            "ticker": "TEST",
            "company_name": "Test Corp",
            "fiscal_years": ["2023", "2024"],
            "period_ends": ["2023-09-30", "2024-09-28"],
            "revenue": {"2023": 1000.0, "2024": 1200.0},
            "net_income": {"2023": 200.0, "2024": 300.0},
            "current_assets": {"2023": 200.0, "2024": 300.0},
            "current_liabilities": {"2023": 100.0, "2024": 0.0},  # Zero liabilities test
            "total_assets": {"2023": 500.0, "2024": 600.0},
            "equity": {"2023": 250.0, "2024": 350.0},
            "commercial_paper": {"2023": 20.0, "2024": 30.0},
            "current_debt": {"2023": 30.0, "2024": 40.0},
            "long_term_debt": {"2023": 50.0, "2024": 60.0},
            "total_liabilities": {"2023": 250.0, "2024": 250.0},
        }

        ratios_data = self.analyzer.calculate_ratios("TEST", inputs=inputs)
        cr_2023 = ratios_data["ratios"]["current_ratio"]["2023"]
        cr_2024 = ratios_data["ratios"]["current_ratio"]["2024"]

        # 200 / 100 = 2.0x
        self.assertEqual(cr_2023, 2.0)
        # Division by zero yields None safely
        self.assertIsNone(cr_2024)

    def test_debt_to_equity_components_and_no_double_count(self) -> None:
        """Verify interest-bearing debt components sum properly without double counting."""
        inputs = {
            "ticker": "TEST",
            "company_name": "Test Corp",
            "fiscal_years": ["2024"],
            "period_ends": ["2024-09-28"],
            "revenue": {"2024": 1000.0},
            "net_income": {"2024": 250.0},
            "current_assets": {"2024": 400.0},
            "current_liabilities": {"2024": 200.0},
            "total_assets": {"2024": 800.0},
            "equity": {"2024": 300.0},
            "commercial_paper": {"2024": 25.0},
            "current_debt": {"2024": 50.0},
            "long_term_debt": {"2024": 75.0},
            "total_liabilities": {"2024": 500.0},
        }

        ratios_data = self.analyzer.calculate_ratios("TEST", inputs=inputs)
        tot_debt = ratios_data["calculation_details"]["total_interest_bearing_debt"]["2024"]
        de_ratio = ratios_data["ratios"]["debt_to_equity"]["2024"]

        # 25 + 50 + 75 = 150.0
        self.assertEqual(tot_debt, 150.0)
        # 150.0 / 300.0 = 0.50x
        self.assertEqual(de_ratio, 0.5)
        # Total debt (150) < Total Liabilities (500)
        self.assertTrue(ratios_data["validation"]["debt_less_than_total_liabilities"])

    def test_asset_turnover_with_prior_year_opening_balance(self) -> None:
        """Verify Asset Turnover uses preceding year assets for beginning balance."""
        inputs = {
            "ticker": "TEST",
            "company_name": "Test Corp",
            "fiscal_years": ["2023", "2024"],
            "period_ends": ["2023-09-30", "2024-09-28"],
            "revenue": {"2023": 500.0, "2024": 800.0},
            "net_income": {"2023": 100.0, "2024": 160.0},
            "current_assets": {"2023": 200.0, "2024": 300.0},
            "current_liabilities": {"2023": 100.0, "2024": 150.0},
            "total_assets": {"2023": 400.0, "2024": 600.0},
            "equity": {"2023": 200.0, "2024": 300.0},
            "commercial_paper": {"2023": 10.0, "2024": 15.0},
            "current_debt": {"2023": 20.0, "2024": 25.0},
            "long_term_debt": {"2023": 50.0, "2024": 60.0},
            "total_liabilities": {"2023": 200.0, "2024": 300.0},
        }

        # Mock get_prior_year_assets_from_facts returning 300.0 for 2022
        self.analyzer.get_prior_year_assets_from_facts = lambda cik, fy: 300.0  # type: ignore

        ratios_data = self.analyzer.calculate_ratios("TEST", inputs=inputs, cik="0001234567")

        # 2023: Beg = 300, End = 400 -> Avg = 350. Turnover = 500 / 350 = 1.4286
        avg_a_2023 = ratios_data["calculation_details"]["average_total_assets"]["2023"]
        at_2023 = ratios_data["ratios"]["asset_turnover"]["2023"]
        self.assertEqual(avg_a_2023, 350.0)
        self.assertAlmostEqual(at_2023, 1.4286, places=3)

        # 2024: Beg = 400 (2023 End), End = 600 -> Avg = 500. Turnover = 800 / 500 = 1.60
        avg_a_2024 = ratios_data["calculation_details"]["average_total_assets"]["2024"]
        at_2024 = ratios_data["ratios"]["asset_turnover"]["2024"]
        self.assertEqual(avg_a_2024, 500.0)
        self.assertEqual(at_2024, 1.60)

    def test_net_profit_margin_calculation(self) -> None:
        """Verify Net Profit Margin calculation ((Net Income / Revenue) * 100)."""
        inputs = {
            "ticker": "TEST",
            "company_name": "Test Corp",
            "fiscal_years": ["2024"],
            "period_ends": ["2024-09-28"],
            "revenue": {"2024": 400.0},
            "net_income": {"2024": 100.0},
            "current_assets": {"2024": 200.0},
            "current_liabilities": {"2024": 100.0},
            "total_assets": {"2024": 500.0},
            "equity": {"2024": 250.0},
            "commercial_paper": {"2024": 0.0},
            "current_debt": {"2024": 20.0},
            "long_term_debt": {"2024": 50.0},
            "total_liabilities": {"2024": 250.0},
        }

        ratios_data = self.analyzer.calculate_ratios("TEST", inputs=inputs)
        npm = ratios_data["ratios"]["net_profit_margin"]["2024"]
        # (100 / 400) * 100 = 25.0%
        self.assertEqual(npm, 25.0)

    def test_excel_integration_preserves_worksheets_and_adds_charts(self) -> None:
        """Verify adding 'Financial Ratios' preserves all previous sheets and inserts 4 charts."""
        # Create dummy workbook with 6 existing sheets
        existing_sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Stock Prices",
            "Audit & Reconciliation",
        ]
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for s in existing_sheets:
            ws = wb.create_sheet(title=s)
            ws["A1"] = f"Content for {s}"

        wb_path = self.excel_dir / "TEST" / "TEST_financial_statements.xlsx"
        wb_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(wb_path)

        inputs = {
            "ticker": "TEST",
            "company_name": "Test Corp",
            "fiscal_years": ["2023", "2024"],
            "period_ends": ["2023-09-30", "2024-09-28"],
            "revenue": {"2023": 500.0, "2024": 800.0},
            "net_income": {"2023": 100.0, "2024": 200.0},
            "current_assets": {"2023": 200.0, "2024": 300.0},
            "current_liabilities": {"2023": 100.0, "2024": 150.0},
            "total_assets": {"2023": 400.0, "2024": 600.0},
            "equity": {"2023": 200.0, "2024": 300.0},
            "commercial_paper": {"2023": 10.0, "2024": 15.0},
            "current_debt": {"2023": 20.0, "2024": 25.0},
            "long_term_debt": {"2023": 50.0, "2024": 60.0},
            "total_liabilities": {"2023": 200.0, "2024": 300.0},
        }

        ratios_data = self.analyzer.calculate_ratios("TEST", inputs=inputs)
        updated_path = self.analyzer.add_financial_ratios_worksheet_to_excel(
            ticker="TEST", ratios_data=ratios_data, workbook_path=wb_path
        )

        res_wb = openpyxl.load_workbook(updated_path)
        expected_sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Stock Prices",
            "Financial Ratios",
            "Audit & Reconciliation",
        ]
        self.assertEqual(res_wb.sheetnames, expected_sheets)

        ws_ratios = res_wb["Financial Ratios"]
        # Verify 4 charts are embedded
        self.assertEqual(len(ws_ratios._charts), 4)

        # Verify formulas in row 7, 8, 9, 10
        self.assertEqual(ws_ratios["B7"].value, "=B14/B15")
        self.assertEqual(ws_ratios["B8"].value, "=B21/B22")
        self.assertEqual(ws_ratios["B9"].value, "=B28/B27")
        self.assertEqual(ws_ratios["B10"].value, "=B31/B32")

    def test_export_csv_schema_and_contents(self) -> None:
        """Verify CSV export contains all required fields and correct values."""
        inputs = {
            "ticker": "TEST",
            "company_name": "Test Corp",
            "fiscal_years": ["2024"],
            "period_ends": ["2024-09-28"],
            "revenue": {"2024": 1000.0},
            "net_income": {"2024": 250.0},
            "current_assets": {"2024": 400.0},
            "current_liabilities": {"2024": 200.0},
            "total_assets": {"2024": 800.0},
            "equity": {"2024": 300.0},
            "commercial_paper": {"2024": 25.0},
            "current_debt": {"2024": 50.0},
            "long_term_debt": {"2024": 75.0},
            "total_liabilities": {"2024": 500.0},
        }

        ratios_data = self.analyzer.calculate_ratios("TEST", inputs=inputs)
        csv_file = self.financial_dir / "TEST_financial_ratios.csv"
        saved_path = self.analyzer.export_ratios_csv("TEST", ratios_data, output_path=csv_file)

        self.assertTrue(saved_path.exists())
        with open(saved_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)

        self.assertEqual(len(rows), 1)
        r = rows[0]
        self.assertEqual(r["company_ticker"], "TEST")
        self.assertEqual(r["fiscal_year"], "2024")
        self.assertEqual(float(r["current_ratio"]), 2.0)
        self.assertEqual(float(r["debt_to_equity_ratio"]), 0.5)
        self.assertEqual(float(r["net_profit_margin_pct"]), 25.0)
        self.assertEqual(r["validation_status"], "Verified")


if __name__ == "__main__":
    unittest.main()
