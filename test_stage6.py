"""Unit tests for Stage 6 Loughran-McDonald Sentiment Analysis and Excel Integration.

Tests cover:
- Auto-discovery of the Loughran-McDonald Master Dictionary (CSV and Excel).
- Parsing active years (>0 active, <0 deleted, 0 absent).
- Calculation of word counts, positive/negative frequencies, and P/N ratios.
- Undefined ratio handling when negative words = 0.
- Words contributing to multiple sentiment categories where applicable.
- Excel workbook integration: adding 'Sentiment Analysis' tab while preserving all previous sheets.
- Verification of CSV manifest fields and formatting.
"""

import csv
import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl

from sentiment_analysis import (
    DictionaryFormatError,
    DictionaryNotFoundError,
    LoughranMcDonaldDictionary,
    MDASentimentAnalyzer,
)


class TestStage6SentimentAnalysis(unittest.TestCase):
    """Test suite verifying Loughran-McDonald sentiment scoring and Excel integration."""

    def setUp(self) -> None:
        """Set up isolated temporary directory for testing."""
        self.test_dir = Path(tempfile.mkdtemp(prefix="test_sentiment_"))
        self.data_dir = self.test_dir / "data"
        self.cleaned_dir = self.test_dir / "cleaned_text"
        self.excel_dir = self.test_dir / "excel_outputs"
        self.raw_dir = self.test_dir / "raw_filings"
        self.fin_dir = self.test_dir / "financial_data"

        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.cleaned_dir.mkdir(parents=True, exist_ok=True)
        self.excel_dir.mkdir(parents=True, exist_ok=True)
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.fin_dir.mkdir(parents=True, exist_ok=True)

        # Create a mock Loughran-McDonald CSV
        self.mock_dict_csv = self.data_dir / "Loughran-McDonald_MasterDictionary_Test.csv"
        headers = [
            "Word", "Seq_num", "Negative", "Positive", "Uncertainty",
            "Litigious", "Strong_Modal", "Weak_Modal", "Constraining", "Complexity"
        ]
        rows = [
            # Word, Seq, Neg, Pos, Uncert, Litig, StrongM, WeakM, Constrain, Complex
            ["LOSS", "1", "2009", "0", "0", "0", "0", "0", "0", "0"],
            ["ADVERSE", "2", "2009", "0", "0", "0", "0", "0", "0", "0"],
            ["DECLINE", "3", "2014", "0", "0", "0", "0", "0", "0", "0"],
            ["IMPROVE", "4", "0", "2009", "0", "0", "0", "0", "0", "0"],
            ["BENEFIT", "5", "0", "-2020", "0", "0", "0", "0", "0", "0"],  # Deleted in 2020 (<0)
            ["EXCELLENT", "6", "0", "2012", "0", "0", "0", "0", "0", "0"],
            ["UNCERTAIN", "7", "0", "0", "2009", "0", "0", "0", "0", "0"],
            ["RISK", "8", "0", "0", "2009", "0", "0", "0", "0", "0"],
            ["LITIGATION", "9", "0", "0", "0", "2009", "0", "0", "0", "0"],
            ["MUST", "10", "0", "0", "0", "0", "2009", "0", "0", "0"],
            ["COULD", "11", "0", "0", "2009", "0", "0", "2009", "0", "0"],  # Uncertainty AND Weak_Modal
            ["RESTRICT", "12", "0", "0", "0", "0", "0", "0", "2009", "0"],
            ["NORMAL", "13", "0", "0", "0", "0", "0", "0", "0", "0"],  # Neutral
        ]
        with open(self.mock_dict_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)

        self.mock_dict = LoughranMcDonaldDictionary(dict_path=self.mock_dict_csv)

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_locate_dictionary_auto_discovery(self) -> None:
        """Verify automatic location of LM dictionary CSV in candidate directories."""
        found = LoughranMcDonaldDictionary.locate_dictionary(search_dirs=[self.data_dir])
        self.assertEqual(found, self.mock_dict_csv)

    def test_excel_dictionary_format_support(self) -> None:
        """Verify that dictionary can be loaded from an Excel (.xlsx) file."""
        excel_dict_path = self.data_dir / "Loughran-McDonald_Test.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Word", "Negative", "Positive", "Uncertainty", "Litigious", "Strong_Modal", "Weak_Modal", "Constraining"])
        ws.append(["LOSS", 2009, 0, 0, 0, 0, 0, 0])
        ws.append(["PROFIT", 0, 2015, 0, 0, 0, 0, 0])
        ws.append(["UNKNOWN", 0, 0, 2009, 0, 0, 0, 0])
        wb.save(excel_dict_path)

        loaded_dict = LoughranMcDonaldDictionary(dict_path=excel_dict_path)
        self.assertIn("LOSS", loaded_dict.categories["Negative"])
        self.assertIn("PROFIT", loaded_dict.categories["Positive"])
        self.assertIn("UNKNOWN", loaded_dict.categories["Uncertainty"])

    def test_multiple_category_contribution(self) -> None:
        """Verify words can belong to multiple categories if indicated by dictionary (e.g. COULD in Uncertainty & Weak_Modal)."""
        analyzer = MDASentimentAnalyzer(dictionary=self.mock_dict)
        tokens = ["COULD"]
        res = analyzer.score_tokens(tokens)
        self.assertEqual(res["counts"]["Uncertainty"], 1)
        self.assertEqual(res["counts"]["Weak_Modal"], 1)
        self.assertEqual(res["counts"]["Modal"], 1)

    def test_score_tokens_metrics(self) -> None:
        """Verify accurate calculation of token counts, P/N ratios, and word frequencies."""
        analyzer = MDASentimentAnalyzer(dictionary=self.mock_dict)
        # 10 tokens: 2 positive, 4 negative, 1 uncertainty, 1 strong modal, 1 weak modal, 1 neutral
        tokens = [
            "IMPROVE", "EXCELLENT",  # 2 Pos
            "LOSS", "ADVERSE", "DECLINE", "LOSS",  # 4 Neg
            "UNCERTAIN",  # 1 Uncert
            "MUST",  # 1 Strong Modal
            "COULD",  # 1 Weak Modal (and Uncertainty)
            "NORMAL",  # 1 Neutral
        ]
        res = analyzer.score_tokens(tokens)
        self.assertEqual(res["total_word_count"], 10)
        self.assertEqual(res["counts"]["Positive"], 2)
        self.assertEqual(res["counts"]["Negative"], 4)
        self.assertEqual(res["counts"]["Uncertainty"], 2)  # UNCERTAIN + COULD
        self.assertEqual(res["counts"]["Strong_Modal"], 1)
        self.assertEqual(res["counts"]["Weak_Modal"], 1)
        self.assertEqual(res["counts"]["Modal"], 2)

        # P/N Ratio = 2 / 4 = 0.500
        self.assertAlmostEqual(res["pn_ratio"], 0.500, places=3)
        # Positive Word Frequency = 2 / 10 = 20.0%
        self.assertAlmostEqual(res["positive_word_frequency"], 20.0, places=1)
        # Negative Word Frequency = 4 / 10 = 40.0%
        self.assertAlmostEqual(res["negative_word_frequency"], 40.0, places=1)

    def test_zero_negative_division_safety_returns_undefined(self) -> None:
        """Verify that zero negative words yields P/N ratio of 'Undefined' without division-by-zero error."""
        analyzer = MDASentimentAnalyzer(dictionary=self.mock_dict)
        tokens = ["IMPROVE", "EXCELLENT", "NORMAL"]
        res = analyzer.score_tokens(tokens)
        self.assertEqual(res["counts"]["Negative"], 0)
        self.assertEqual(res["pn_ratio"], "Undefined")
        self.assertEqual(res["positive_word_frequency"], 2 / 3 * 100.0)

    def test_excel_integration_preserves_worksheets_and_names_tab(self) -> None:
        """Verify worksheet is named 'Sentiment Analysis' and added to Excel while preserving all existing tabs."""
        ticker = "TEST"
        company_excel_dir = self.excel_dir / ticker
        company_excel_dir.mkdir(parents=True, exist_ok=True)
        excel_path = company_excel_dir / f"{ticker}_financial_statements.xlsx"

        # Create workbook with existing statement tabs
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sheet in ["Income Statement", "Balance Sheet", "Cash Flow Statement", "Audit & Reconciliation"]:
            ws = wb.create_sheet(title=sheet)
            ws["A1"] = f"Original content for {sheet}"
        wb.save(excel_path)

        analyzer = MDASentimentAnalyzer(
            dictionary=self.mock_dict,
            cleaned_dir=self.cleaned_dir,
            excel_dir=self.excel_dir,
            raw_filings_dir=self.raw_dir,
            financial_data_dir=self.fin_dir,
        )

        mock_records = [
            {
                "company_ticker": ticker,
                "ticker": ticker,
                "fiscal_reporting_year": "2024",
                "fiscal_year": "2024",
                "fiscal_period_end_date": "2024-09-28",
                "source_mda_filename": f"{ticker}_2024_MDA.txt",
                "cleaned_filename": f"{ticker}_2024_MDA_CLEANED.txt",
                "cleaned_text_words": 2500,
                "total_word_count": 2000,
                "positive_word_count": 10,
                "negative_word_count": 20,
                "uncertainty_word_count": 15,
                "litigious_word_count": 8,
                "strong_modal_word_count": 5,
                "weak_modal_word_count": 7,
                "constraining_word_count": 12,
                "modal_word_count": 12,
                "pn_ratio": 0.500,
                "positive_word_frequency": 0.5,
                "negative_word_frequency": 1.0,
                "net_polarity": -0.3333,
                "proportions": {
                    "Positive": 0.5,
                    "Negative": 1.0,
                    "Uncertainty": 0.75,
                    "Litigious": 0.4,
                    "Constraining": 0.6,
                    "Modal": 0.6,
                },
                "dictionary_version": self.mock_dict.version,
                "processing_status": "Verified",
                "top_negative": ["LOSS", "ADVERSE"],
                "top_positive": ["IMPROVE"],
                "top_uncertainty": ["UNCERTAIN"],
            }
        ]

        updated_path = analyzer.add_sentiment_worksheet_to_excel(
            ticker=ticker, records=mock_records, company_name="Test Corp", cik="0001234567"
        )

        # Inspect updated workbook
        res_wb = openpyxl.load_workbook(updated_path)
        expected_sheets = [
            "Income Statement",
            "Balance Sheet",
            "Cash Flow Statement",
            "Sentiment Analysis",
            "Audit & Reconciliation",
        ]
        self.assertEqual(res_wb.sheetnames, expected_sheets)
        self.assertIn("Sentiment Analysis", res_wb.sheetnames)

        # Confirm original sheets still have their content
        self.assertEqual(res_wb["Income Statement"]["A1"].value, "Original content for Income Statement")
        self.assertEqual(res_wb["Audit & Reconciliation"]["A1"].value, "Original content for Audit & Reconciliation")

        # Confirm sentiment worksheet has correct values and formatting
        ws_sent = res_wb["Sentiment Analysis"]
        self.assertEqual(ws_sent["A1"].value, "Test Corp (TEST)")
        self.assertEqual(ws_sent["B4"].value, "FY 2024")
        self.assertEqual(ws_sent["B5"].value, "2024-09-28")
        self.assertEqual(ws_sent.freeze_panes, "B6")

    def test_export_summary_csv_fields(self) -> None:
        """Verify that CSV output contains all required fields specified in Stage 6 prompt."""
        ticker = "TEST"
        analyzer = MDASentimentAnalyzer(
            dictionary=self.mock_dict,
            cleaned_dir=self.cleaned_dir,
            excel_dir=self.excel_dir,
            raw_filings_dir=self.raw_dir,
            financial_data_dir=self.fin_dir,
        )

        mock_records = [
            {
                "company_ticker": ticker,
                "ticker": ticker,
                "fiscal_reporting_year": "2024",
                "fiscal_year": "2024",
                "fiscal_period_end_date": "2024-09-28",
                "source_mda_filename": f"{ticker}_2024_MDA.txt",
                "total_word_count": 2000,
                "positive_word_count": 10,
                "negative_word_count": 20,
                "uncertainty_word_count": 15,
                "litigious_word_count": 8,
                "strong_modal_word_count": 5,
                "weak_modal_word_count": 7,
                "constraining_word_count": 12,
                "pn_ratio": 0.500,
                "positive_word_frequency": 0.5,
                "negative_word_frequency": 1.0,
                "dictionary_version": self.mock_dict.version,
                "processing_status": "Verified",
            }
        ]

        p_csv, m_csv = analyzer.export_summary_csv(ticker=ticker, records=mock_records)
        self.assertTrue(p_csv.exists())
        self.assertTrue(m_csv.exists())

        with open(p_csv, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            row = next(reader)
            expected_keys = [
                "company_ticker",
                "fiscal_reporting_year",
                "fiscal_period_end_date",
                "source_mda_filename",
                "total_word_count",
                "positive_word_count",
                "negative_word_count",
                "uncertainty_word_count",
                "litigious_word_count",
                "strong_modal_word_count",
                "weak_modal_word_count",
                "constraining_word_count",
                "pn_ratio",
                "positive_word_frequency",
                "negative_word_frequency",
                "dictionary_version",
                "processing_status",
            ]
            for k in expected_keys:
                self.assertIn(k, row)


if __name__ == "__main__":
    unittest.main()
