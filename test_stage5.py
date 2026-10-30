"""Unit tests for Stage 5 MD&A Text Cleaning and Sentiment Preparation.

Tests text cleaning normalization, running header removal, quote standardization,
preservation of financial and negation vocabulary, uppercase tokenization,
and quality control validation.
"""

import shutil
import tempfile
import unittest
from pathlib import Path

from text_cleaner import MDATextCleaner


class TestStage5TextCleaner(unittest.TestCase):
    """Test suite verifying MD&A text cleaning rules, normalization, and tokenization."""

    def setUp(self) -> None:
        """Create a temporary directory for isolated test operations."""
        self.test_dir = Path(tempfile.mkdtemp(prefix="test_clean_"))
        self.extracted_dir = self.test_dir / "extracted_text"
        self.cleaned_dir = self.test_dir / "cleaned_text"
        self.extracted_dir.mkdir(parents=True, exist_ok=True)
        self.cleaned_dir.mkdir(parents=True, exist_ok=True)
        self.cleaner = MDATextCleaner(
            extracted_dir=self.extracted_dir, cleaned_dir=self.cleaned_dir
        )

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_running_headers_removal(self) -> None:
        """Verify SEC running page headers/footers are stripped cleanly."""
        sample = (
            "Results of operations were favorable.\n\n"
            "Apple Inc. | 2024 Form 10-K | 26\n\n"
            "Operating expenses increased by 10%."
        )
        cleaned = self.cleaner.clean_text(sample)
        self.assertNotIn("Apple Inc. | 2024 Form 10-K | 26", cleaned)
        self.assertIn("Results of operations were favorable.", cleaned)
        self.assertIn("Operating expenses increased by 10%.", cleaned)

    def test_quotes_and_entities_normalization(self) -> None:
        """Verify curly quotes, non-breaking spaces, and HTML entities are normalized."""
        sample = (
            "The Company\u2019s \u201cServices\u201d revenue &amp; net sales grew.\n\n"
            "Product\xa0sales\xa0reached $50,000 million."
        )
        cleaned = self.cleaner.clean_text(sample)
        # Curly apostrophe replaced by standard single quote
        self.assertIn("Company's", cleaned)
        # Curly quotes replaced by standard double quotes
        self.assertIn('"Services"', cleaned)
        # &amp; decoded to &
        self.assertIn("&", cleaned)
        # Non-breaking spaces normalized
        self.assertNotIn("\xa0", cleaned)
        self.assertIn("Product sales reached $50,000 million.", cleaned)

    def test_preservation_of_critical_financial_and_negation_words(self) -> None:
        """Verify that negative, uncertainty, risk, and numerical words remain intact."""
        sample = (
            "The company experienced a loss due to litigation and economic uncertainty. "
            "Net income did not decline, and there was no material impairment. "
            "Total long-term debt was $105,000 million in 2024."
        )
        cleaned = self.cleaner.clean_text(sample)

        # Confirm all words preserved in cleaned text
        for word in ["loss", "litigation", "uncertainty", "not", "no", "decline", "impairment", "long-term", "$105,000"]:
            self.assertIn(word, cleaned)

        # Confirm tokenization preserves these tokens in uppercase
        tokens = self.cleaner.tokenize_for_sentiment(cleaned)
        token_set = set(tokens)
        self.assertIn("LOSS", token_set)
        self.assertIn("LITIGATION", token_set)
        self.assertIn("UNCERTAINTY", token_set)
        self.assertIn("NOT", token_set)
        self.assertIn("NO", token_set)
        self.assertIn("DECLINE", token_set)
        self.assertIn("IMPAIRMENT", token_set)
        self.assertIn("LONG-TERM", token_set)

    def test_reproducibility(self) -> None:
        """Verify that running the cleaning process repeatedly produces identical output."""
        sample = "Revenue increased 5% to $391,035 million in FY 2024.\n\nOperating margin was 31.5%."
        run1 = self.cleaner.clean_text(sample)
        run2 = self.cleaner.clean_text(sample)
        self.assertEqual(run1, run2)

    def test_validation_flags_empty_text(self) -> None:
        """Verify quality control flags empty cleaned text as Failed."""
        status, warnings = self.cleaner.validate_cleaned_text(
            original_text="",
            cleaned_text="",
            tokens=[],
            fiscal_year="2024",
        )
        self.assertEqual(status, "Failed")


if __name__ == "__main__":
    unittest.main()
