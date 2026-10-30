"""Unit tests for Stage 4 Financial Statement Extraction and Excel Integration.

Tests XBRL fact extraction, duration/instant filtering, accounting reconciliations,
and multi-tab Excel workbook generation.
"""

import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl

from financial_extractor import FinancialExtractor


class TestStage4FinancialLogic(unittest.TestCase):
    """Test suite verifying financial statement extraction, reconciliations, and Excel output."""

    def setUp(self) -> None:
        """Set up temporary directories and test extractor."""
        self.test_dir = Path(tempfile.mkdtemp(prefix="test_fin_"))
        self.financial_dir = self.test_dir / "financial_data"
        self.excel_dir = self.test_dir / "excel_outputs"
        self.financial_dir.mkdir(parents=True, exist_ok=True)
        self.excel_dir.mkdir(parents=True, exist_ok=True)

        self.extractor = FinancialExtractor(
            user_agent="Test User test@university.edu",
            financial_dir=self.financial_dir,
            excel_dir=self.excel_dir,
        )

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_extract_fact_value_duration_and_instant(self) -> None:
        """Verify distinction between annual duration facts and instant facts."""
        mock_gaap = {
            "Revenues": {
                "units": {
                    "USD": [
                        # Full year (duration ~365 days)
                        {"start": "2023-10-01", "end": "2024-09-28", "val": 391035000000, "form": "10-K", "fp": "FY", "filed": "2024-11-01"},
                        # Quarter (duration ~90 days) - should be ignored
                        {"start": "2024-06-30", "end": "2024-09-28", "val": 94930000000, "form": "10-Q", "fp": "Q4", "filed": "2024-11-01"},
                    ]
                }
            },
            "Assets": {
                "units": {
                    "USD": [
                        # Instant fact (no start date)
                        {"end": "2024-09-28", "val": 364980000000, "form": "10-K", "fp": "FY", "filed": "2024-11-01"}
                    ]
                }
            },
        }

        # 1. Test annual duration fact
        val_dur = self.extractor._extract_fact_value(
            gaap_facts=mock_gaap,
            tag_list=["Revenues"],
            period_end="2024-09-28",
            is_duration=True,
        )
        self.assertEqual(val_dur, 391035000000.0)

        # 2. Test instant fact
        val_inst = self.extractor._extract_fact_value(
            gaap_facts=mock_gaap,
            tag_list=["Assets"],
            period_end="2024-09-28",
            is_duration=False,
        )
        self.assertEqual(val_inst, 364980000000.0)

    def test_balance_sheet_reconciliation_logic(self) -> None:
        """Verify accounting reconciliation detects balanced and unbalanced statements."""
        fys = ["2023", "2024"]

        # 2023 is balanced (Assets 100 = Liab 60 + Eq 40)
        # 2024 is out of balance (Assets 120 != Liab 60 + Eq 40 = 100)
        balance_rows = [
            {"title": "Total Assets", "values": {"2023": 100.0, "2024": 120.0}},
            {"title": "Total Liabilities", "values": {"2023": 60.0, "2024": 60.0}},
            {"title": "Total Shareholders' Equity", "values": {"2023": 40.0, "2024": 40.0}},
        ]
        income_rows = [
            {"title": "Revenue (Total Net Sales)", "is_subtotal": False, "values": {"2023": 200.0, "2024": 250.0}},
            {"title": "Cost of Revenue (Cost of Goods Sold)", "is_subtotal": False, "values": {"2023": 120.0, "2024": 150.0}},
            {"title": "Gross Profit", "is_subtotal": True, "values": {"2023": 80.0, "2024": 100.0}},
        ]
        cashflow_rows = [
            {"title": "Operating Activities", "values": {"2023": 50.0, "2024": 60.0}},
            {"title": "Investing Activities", "values": {"2023": -20.0, "2024": -25.0}},
            {"title": "Financing Activities", "values": {"2023": -25.0, "2024": -30.0}},
            {"title": "Net Increase (Decrease)", "values": {"2023": 5.0, "2024": 5.0}},
        ]

        rec = self.extractor._run_reconciliations(
            fys=fys,
            income_rows=income_rows,
            balance_rows=balance_rows,
            cashflow_rows=cashflow_rows,
        )

        self.assertFalse(rec["balance_sheet_balanced"])
        self.assertTrue(rec["balance_sheet_checks"]["2023"]["balanced"])
        self.assertFalse(rec["balance_sheet_checks"]["2024"]["balanced"])
        self.assertTrue(rec["gross_profit_checks"]["2023"]["reconciled"])
        self.assertTrue(rec["gross_profit_checks"]["2024"]["reconciled"])

    def test_excel_workbook_structure_and_formatting(self) -> None:
        """Verify that openpyxl workbook creates 4 worksheets with freeze panes and headers."""
        periods = [
            {"fiscal_year": "2023", "period_end": "2023-09-30"},
            {"fiscal_year": "2024", "period_end": "2024-09-28"},
        ]
        dummy_data = {
            "ticker": "AAPL",
            "company_name": "Apple Inc.",
            "cik": "0000320193",
            "periods": periods,
            "income_statement": [
                {"title": "Revenue (Total Net Sales)", "is_subtotal": False, "is_total": False, "values": {"2023": 383285.0, "2024": 391035.0}},
                {"title": "Gross Profit", "is_subtotal": True, "is_total": False, "values": {"2023": 169148.0, "2024": 180683.0}},
                {"title": "Net Income", "is_subtotal": False, "is_total": True, "values": {"2023": 96995.0, "2024": 93736.0}},
            ],
            "balance_sheet": [
                {"title": "Total Assets", "is_subtotal": False, "is_total": True, "values": {"2023": 352583.0, "2024": 364980.0}},
                {"title": "Total Liabilities", "is_subtotal": True, "is_total": False, "values": {"2023": 290437.0, "2024": 308030.0}},
                {"title": "Total Shareholders' Equity", "is_subtotal": True, "is_total": False, "values": {"2023": 62146.0, "2024": 56950.0}},
            ],
            "cash_flows": [
                {"title": "Net Cash Provided by Operating Activities", "is_subtotal": True, "is_total": False, "values": {"2023": 110543.0, "2024": 118254.0}},
                {"title": "Net Increase (Decrease) in Cash", "is_subtotal": False, "is_total": True, "values": {"2023": -5762.0, "2024": 743.0}},
            ],
            "reconciliations": {
                "balance_sheet_balanced": True,
                "balance_sheet_checks": {
                    "2023": {"assets": 352583.0, "liab_plus_eq": 352583.0, "diff": 0.0, "balanced": True},
                    "2024": {"assets": 364980.0, "liab_plus_eq": 364980.0, "diff": 0.0, "balanced": True},
                },
                "gross_profit_checks": {
                    "2023": {"revenue": 383285.0, "cogs": 214137.0, "reported_gp": 169148.0, "diff": 0.0, "reconciled": True},
                    "2024": {"revenue": 391035.0, "cogs": 210352.0, "reported_gp": 180683.0, "diff": 0.0, "reconciled": True},
                },
                "cash_flow_checks": {},
                "warnings": [],
            },
        }

        out_path = self.test_dir / "test_out.xlsx"
        result_path = self.extractor.create_excel_workbook(dummy_data, output_path=out_path)

        self.assertTrue(result_path.exists())

        wb = openpyxl.load_workbook(result_path)
        expected_sheets = ["Income Statement", "Balance Sheet", "Cash Flow Statement", "Audit & Reconciliation"]
        self.assertEqual(wb.sheetnames, expected_sheets)

        # Verify freeze panes and headers on Income Statement
        ws_is = wb["Income Statement"]
        self.assertEqual(ws_is.freeze_panes, "B6")
        self.assertIn("Apple Inc.", ws_is["A1"].value)
        self.assertEqual(ws_is["B4"].value, "FY 2023")
        self.assertEqual(ws_is["C4"].value, "FY 2024")


if __name__ == "__main__":
    unittest.main()
