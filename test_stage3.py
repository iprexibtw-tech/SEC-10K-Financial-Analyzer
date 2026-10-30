"""Unit tests for Stage 3 SEC Form 10-K Item 7 (MD&A) Extraction.

Tests boundary identification, TOC discrimination, quality control validation,
and manifest recording.
"""

import shutil
import tempfile
import unittest
from pathlib import Path

from mda_extractor import MDAExtractor


class TestStage3MDALogic(unittest.TestCase):
    """Test suite verifying MD&A boundary parsing, TOC avoidance, and quality checks."""

    def setUp(self) -> None:
        """Create a temporary directory for isolated test operations."""
        self.test_dir = Path(tempfile.mkdtemp(prefix="test_mda_"))
        self.raw_dir = self.test_dir / "raw"
        self.out_dir = self.test_dir / "out"
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.out_dir.mkdir(parents=True, exist_ok=True)
        self.extractor = MDAExtractor(
            raw_filings_dir=self.raw_dir, extracted_text_dir=self.out_dir
        )

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_toc_discrimination_and_boundary_finding(self) -> None:
        """Verify that TOC Item 7 is ignored and substantive Item 7 is located."""
        sample_html = """
        <html>
        <head><title>Form 10-K</title></head>
        <body>
            <div>TABLE OF CONTENTS</div>
            <div>Item 1. Business 1</div>
            <div>Item 1A. Risk Factors 5</div>
            <div>Item 7. Management's Discussion and Analysis of Financial Condition and Results of Operations 21</div>
            <div>Item 7A. Quantitative and Qualitative Disclosures About Market Risk 35</div>
            <div>Item 8. Financial Statements 38</div>
            <div>Part I</div>
            <div>Some narrative business content...</div>
            <div>Item 1A. Risk Factors</div>
            <div>Risks discussion...</div>
            <div>Part II</div>
            <div>Item 7. Management's Discussion and Analysis of Financial Condition and Results of Operations</div>
            <div>Overview and financial performance for the fiscal year. Net revenue increased 10%...</div>
            <div>Segment performance: Americas, Europe, Greater China...</div>
            <div>Capital resources and liquidity remain strong with $50 billion cash.</div>
            <div>Item 7A. Quantitative and Qualitative Disclosures About Market Risk</div>
            <div>Interest rate risk and foreign currency hedging activities...</div>
            <div>Item 8. Financial Statements and Supplementary Data</div>
        </body>
        </html>
        """
        raw_file = self.raw_dir / "AAPL_10K_FY2025_2025-10-31.htm"
        raw_file.write_text(sample_html, encoding="utf-8")

        lines = self.extractor._clean_html_to_lines(sample_html)
        start_idx, end_idx, start_label, end_label = self.extractor.find_mda_boundaries(lines)

        self.assertIsNotNone(start_idx)
        self.assertIsNotNone(end_idx)
        self.assertEqual(start_label, "Item 7")
        self.assertEqual(end_label, "Item 7A")

        # Substantive start line must not be the TOC line
        mda_lines = lines[start_idx:end_idx]
        mda_text = "\n".join(mda_lines)

        self.assertIn("Overview and financial performance", mda_text)
        self.assertIn("Capital resources and liquidity", mda_text)
        # Must not contain Item 7A or Item 8 content
        self.assertNotIn("Interest rate risk and foreign currency", mda_text)
        self.assertNotIn("TABLE OF CONTENTS", mda_text)

    def test_validation_flags_short_extractions(self) -> None:
        """Verify quality control flags suspiciously short content (< 1,000 words)."""
        short_text = "Item 7. Management's Discussion and Analysis. Very brief summary of results."
        status, warnings = self.extractor.validate_extraction(
            extracted_text=short_text,
            start_line_text="Item 7. Management's Discussion and Analysis",
            end_label="Item 7A",
            total_source_lines=100,
            extracted_line_count=2,
        )
        self.assertEqual(status, "Verified (with warnings)")
        self.assertTrue(any("short" in w.lower() for w in warnings))

    def test_validation_flags_empty_text(self) -> None:
        """Verify quality control marks empty text as Failed."""
        status, warnings = self.extractor.validate_extraction(
            extracted_text="",
            start_line_text="",
            end_label="",
            total_source_lines=10,
            extracted_line_count=0,
        )
        self.assertEqual(status, "Failed")

    def test_fallback_to_item_8_when_item_7a_omitted(self) -> None:
        """Verify boundary finder falls back to Item 8 if Item 7A is omitted."""
        sample_lines = [
            "Part II",
            "Item 7. Management's Discussion and Analysis of Financial Condition and Results of Operations",
            "Detailed operational overview...",
            "Liquidity discussion...",
            "Item 8. Financial Statements and Supplementary Data",
            "Consolidated balance sheets...",
        ]

        start_idx, end_idx, start_label, end_label = self.extractor.find_mda_boundaries(sample_lines)
        self.assertEqual(start_idx, 1)
        self.assertEqual(end_idx, 4)
        self.assertEqual(end_label, "Item 8")


if __name__ == "__main__":
    unittest.main()
