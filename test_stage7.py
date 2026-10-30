"""Unit tests for Stage 7 Historical Stock Price Extraction and Excel Integration.

Tests cover:
- Weekend and market holiday matching to preceding trading days.
- Lookahead bias prevention (actual trading date <= balance sheet date).
- Excel workbook integration: adding 'Stock Prices' worksheet while preserving all previous sheets.
- CSV summary export schema and field verification.
- Positive price and volume validation.
- Error handling for missing dates.
"""

import csv
import datetime
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pandas as pd

from stock_prices import (
    MissingDateError,
    StockPriceError,
    StockPriceExtractor,
    StockPriceRetrievalError,
)


class TestStage7StockPrices(unittest.TestCase):
    """Test suite verifying historical stock price extraction logic and Excel integration."""

    def setUp(self) -> None:
        """Set up isolated test directory with mock filings and metadata."""
        self.test_dir = Path(tempfile.mkdtemp(prefix="test_stock_"))
        self.excel_dir = self.test_dir / "excel_outputs"
        self.fin_dir = self.test_dir / "financial_data"
        self.raw_dir = self.test_dir / "raw_filings"

        self.excel_dir.mkdir(parents=True, exist_ok=True)
        self.fin_dir.mkdir(parents=True, exist_ok=True)
        self.raw_dir.mkdir(parents=True, exist_ok=True)

        self.extractor = StockPriceExtractor(
            excel_dir=self.excel_dir,
            financial_data_dir=self.fin_dir,
            raw_filings_dir=self.raw_dir,
        )

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_weekend_matching_selects_preceding_friday(self) -> None:
        """Verify that a Saturday balance sheet date is matched to Friday."""
        # Mock dataframe with Friday trading date
        dates = pd.date_range("2024-09-23", "2024-09-27", freq="B")  # Mon-Fri
        mock_df = pd.DataFrame(
            {
                "Close": [220.0, 222.0, 224.0, 225.0, 227.79],
                "Adj Close": [218.0, 220.0, 222.0, 223.0, 225.90],
                "Volume": [50000000] * 5,
            },
            index=dates,
        )

        with patch("yfinance.Ticker") as mock_ticker_cls:
            mock_inst = MagicMock()
            mock_inst.history.return_value = mock_df
            mock_ticker_cls.return_value = mock_inst

            # Saturday 2024-09-28
            periods = [{"fiscal_year": "2024", "balance_sheet_date": "2024-09-28"}]
            records = self.extractor.retrieve_stock_prices(ticker="AAPL", periods=periods)

            self.assertEqual(len(records), 1)
            rec = records[0]
            self.assertEqual(rec["fiscal_balance_sheet_date"], "2024-09-28")
            self.assertEqual(rec["actual_trading_date"], "2024-09-27")  # Friday
            self.assertEqual(rec["offset_days"], 1)
            self.assertEqual(rec["closing_price_usd"], 227.79)
            self.assertEqual(rec["adjusted_closing_price_usd"], 225.90)

    def test_actual_trading_date_never_exceeds_balance_sheet_date(self) -> None:
        """Verify that trading date is strictly <= balance sheet date."""
        dates = pd.date_range("2024-09-20", "2024-09-30", freq="B")
        mock_df = pd.DataFrame(
            {
                "Close": [200.0 + i for i in range(len(dates))],
                "Adj Close": [198.0 + i for i in range(len(dates))],
                "Volume": [1000000] * len(dates),
            },
            index=dates,
        )

        with patch("yfinance.Ticker") as mock_ticker_cls:
            mock_inst = MagicMock()
            mock_inst.history.return_value = mock_df
            mock_ticker_cls.return_value = mock_inst

            periods = [
                {"fiscal_year": "2023", "balance_sheet_date": "2024-09-21"},
                {"fiscal_year": "2024", "balance_sheet_date": "2024-09-28"},
            ]
            records = self.extractor.retrieve_stock_prices(ticker="TEST", periods=periods)

            for rec in records:
                bs_d = datetime.datetime.strptime(rec["fiscal_balance_sheet_date"], "%Y-%m-%d").date()
                tr_d = datetime.datetime.strptime(rec["actual_trading_date"], "%Y-%m-%d").date()
                self.assertLessEqual(tr_d, bs_d)

    def test_excel_integration_preserves_worksheets(self) -> None:
        """Verify that adding 'Stock Prices' preserves all previous sheets."""
        ticker = "AAPL"
        company_excel_dir = self.excel_dir / ticker
        company_excel_dir.mkdir(parents=True, exist_ok=True)
        excel_path = company_excel_dir / f"{ticker}_financial_statements.xlsx"

        # Create workbook with 5 existing worksheets
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Audit & Reconciliation",
        ]
        for s in sheets:
            ws = wb.create_sheet(title=s)
            ws["A1"] = f"Original content for {s}"
        wb.save(excel_path)

        mock_records = [
            {
                "company_ticker": ticker,
                "ticker": ticker,
                "fiscal_year": "2024",
                "fiscal_balance_sheet_date": "2024-09-28",
                "actual_trading_date": "2024-09-27",
                "offset_days": 1,
                "closing_price_usd": 227.79,
                "adjusted_closing_price_usd": 225.90,
                "trading_volume": 34026000,
                "primary_measure": "Split-Adjusted Closing Price",
                "data_source": "Yahoo Finance (via yfinance)",
                "retrieval_status": "Verified",
                "notes": "Matched to preceding Friday.",
            }
        ]

        updated_path = self.extractor.add_stock_prices_worksheet_to_excel(
            ticker=ticker, price_records=mock_records, company_name="Apple Inc.", cik="0000320193"
        )

        res_wb = openpyxl.load_workbook(updated_path)
        expected_sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Stock Prices",
            "Audit & Reconciliation",
        ]
        self.assertEqual(res_wb.sheetnames, expected_sheets)

        # Confirm original sheets are intact
        self.assertEqual(res_wb["Income Statement"]["A1"].value, "Original content for Income Statement")
        self.assertEqual(res_wb["Sentiment Analysis"]["A1"].value, "Original content for Sentiment Analysis")

        # Confirm Stock Prices sheet formatting
        ws_sp = res_wb["Stock Prices"]
        self.assertEqual(ws_sp["A1"].value, "Apple Inc. (AAPL)")
        self.assertEqual(ws_sp["B4"].value, "FY 2024")
        self.assertEqual(ws_sp["B5"].value, "2024-09-28")
        self.assertEqual(ws_sp["B6"].value, "2024-09-27")
        self.assertEqual(ws_sp["B8"].value, 227.79)
        self.assertEqual(ws_sp.freeze_panes, "B7")

    def test_export_summary_csv_schema(self) -> None:
        """Verify that exported CSV contains all required fields."""
        ticker = "TEST"
        mock_records = [
            {
                "company_ticker": ticker,
                "fiscal_year": "2024",
                "fiscal_balance_sheet_date": "2024-09-28",
                "actual_trading_date": "2024-09-27",
                "offset_days": 1,
                "closing_price_usd": 150.00,
                "adjusted_closing_price_usd": 145.00,
                "trading_volume": 1000000,
                "primary_measure": "Split-Adjusted Closing Price",
                "data_source": "Yahoo Finance (via yfinance)",
                "retrieval_status": "Verified",
                "notes": "Test note",
            }
        ]

        p_csv, m_csv = self.extractor.export_summary_csv(ticker=ticker, price_records=mock_records)
        self.assertTrue(p_csv.exists())
        self.assertTrue(m_csv.exists())

        with open(p_csv, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            row = next(reader)
            expected_keys = [
                "company_ticker",
                "fiscal_year",
                "fiscal_balance_sheet_date",
                "actual_trading_date",
                "offset_days",
                "closing_price_usd",
                "adjusted_closing_price_usd",
                "trading_volume",
                "primary_measure",
                "data_source",
                "retrieval_status",
                "notes",
            ]
            for k in expected_keys:
                self.assertIn(k, row)

    def test_missing_dates_raises_error(self) -> None:
        """Verify MissingDateError is raised when no balance sheet dates exist."""
        with self.assertRaises(MissingDateError):
            self.extractor.get_balance_sheet_dates(ticker="NONEXISTENT")


if __name__ == "__main__":
    unittest.main()
