"""MD&A Text Cleaning and Preparation Module for Financial Sentiment Analysis.

This module processes raw extracted Item 7 MD&A text files into cleaned,
standardized text suitable for Loughran-McDonald financial sentiment analysis.

Key Responsibilities:
- Remove residual HTML tags, HTML entities, and formatting artifacts.
- Remove running headers, footers, and decorative list markers.
- Normalize whitespace and paragraph boundaries.
- Standardize punctuation, curly quotes, and hyphenation.
- Preserve all meaningful financial terminology, numbers, negations, and risk words.
- Generate uppercase token stream compatible with Loughran-McDonald dictionary.
- Output UTF-8 files: data/cleaned_text/{TICKER}/{TICKER}_{YEAR}_MDA_CLEANED.txt.
- Produce quality control audit manifest in CSV.
"""

import csv
import html
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import config

logger = logging.getLogger("Text-Cleaner")


class TextCleanerError(Exception):
    """Base exception for text cleaning operations."""


class MDATextCleaner:
    """Cleans, normalizes, and tokenizes MD&A text for sentiment analysis."""

    def __init__(
        self,
        extracted_dir: Optional[Path] = None,
        cleaned_dir: Optional[Path] = None,
    ) -> None:
        """Initialize the text cleaner with source and destination directories.

        Args:
            extracted_dir (Optional[Path]): Directory containing extracted MD&A files.
            cleaned_dir (Optional[Path]): Directory for saving cleaned text files.
        """
        self.extracted_dir = extracted_dir or config.EXTRACTED_TEXT_DIR
        self.cleaned_dir = cleaned_dir or config.CLEANED_TEXT_DIR

    @staticmethod
    def clean_text(text: str) -> str:
        """Perform comprehensive, reproducible cleaning on extracted MD&A text.

        Preserves financial vocabulary, numbers, paragraph structure, and negations
        while stripping residual HTML, running page headers/footers, and non-substantive
        decorative artifacts.

        Args:
            text (str): Raw extracted MD&A text string.

        Returns:
            str: Cleaned and standardized text.
        """
        if not text:
            return ""

        # 1. Decode standard and numeric HTML entities (&amp;, &#8217;, &nbsp;, etc.)
        cleaned = html.unescape(text)

        # 2. Normalize non-breaking spaces and special unicode spaces to standard ASCII space
        cleaned = (
            cleaned.replace("\xa0", " ")
            .replace("\u200b", "")  # zero-width space
            .replace("\ufeff", "")  # byte order mark
            .replace("\u202f", " ")  # narrow no-break space
        )

        # 3. Strip any residual HTML/XML tags
        cleaned = re.sub(r"<[^>]+>", " ", cleaned)

        # 4. Standardize unicode quotation marks and apostrophes
        cleaned = (
            cleaned.replace("“", '"')
            .replace("”", '"')
            .replace("‘", "'")
            .replace("’", "'")
            .replace("`", "'")
        )

        # 5. Remove SEC running headers and page break markers
        # Examples: "Apple Inc. | 2024 Form 10-K | 26" or "NVIDIA Corporation | 2024 Form 10-K | 45"
        cleaned = re.sub(
            r"(?:Apple\s*Inc\.|NVIDIA\s*Corp(?:\.|oration)?)\s*\|\s*\d{4}\s*Form\s*10-K\s*\|\s*\d{1,3}",
            "",
            cleaned,
            flags=re.IGNORECASE,
        )

        # 6. Normalize em-dashes and en-dashes to standard hyphen or spaced hyphen
        cleaned = cleaned.replace("—", " - ").replace("–", " - ")

        # 7. Remove non-substantive decorative bullet markers and table lines
        # Bullets: •, ●, ▪, ◦, ‣, ∙
        cleaned = re.sub(r"[\u2022\u25cf\u25aa\u25e6\u2023\u2219]", " ", cleaned)
        # Decorative divider lines: e.g. "----", "____", "===="
        cleaned = re.sub(r"^[\-\_\=\*\~]{3,}\s*$", "", cleaned, flags=re.MULTILINE)

        # 8. Clean line-by-line whitespace
        lines = []
        for line in cleaned.split("\n"):
            stripped = line.strip()
            # Collapse multiple inline spaces and tabs into a single space
            stripped = re.sub(r"[ \t]+", " ", stripped)
            lines.append(stripped)

        # 9. Normalize multiple blank lines into clean paragraph breaks (max 2 newlines)
        rejoined = "\n".join(lines)
        rejoined = re.sub(r"\n{3,}", "\n\n", rejoined)

        return rejoined.strip()

    @staticmethod
    def tokenize_for_sentiment(cleaned_text: str) -> List[str]:
        """Tokenize text into uppercase words for Loughran-McDonald dictionary matching.

        Extracts alphabetic tokens and hyphenated compound words, standardizing
        to uppercase. Preserves contractions and negations without stemming or
        stopword removal.

        Args:
            cleaned_text (str): Cleaned MD&A text.

        Returns:
            List[str]: List of uppercase word tokens.
        """
        if not cleaned_text:
            return []

        # Match alphabetic tokens and compound words (e.g. "LONG-TERM", "COST-EFFECTIVE")
        # Words of at least 2 characters, or standalone single-letter words like "A"
        raw_tokens = re.findall(r"\b[A-Za-z]+(?:-[A-Za-z]+)*\b", cleaned_text)

        tokens = [t.upper() for t in raw_tokens if t]
        return tokens

    def validate_cleaned_text(
        self,
        original_text: str,
        cleaned_text: str,
        tokens: List[str],
        fiscal_year: str,
    ) -> Tuple[str, List[str]]:
        """Validate quality of cleaned text and token stream.

        Args:
            original_text (str): Raw extracted text.
            cleaned_text (str): Cleaned text.
            tokens (List[str]): Tokenized words.
            fiscal_year (str): Fiscal reporting year.

        Returns:
            Tuple[str, List[str]]: (status, list_of_warnings)
        """
        warnings_list: List[str] = []

        if not cleaned_text or not cleaned_text.strip():
            return "Failed", ["Cleaned text is empty."]

        orig_words = len(original_text.split())
        clean_words = len(cleaned_text.split())

        # Check percentage reduction
        pct_change = ((clean_words - orig_words) / orig_words) * 100.0 if orig_words else 0.0

        if pct_change < -30.0:
            warnings_list.append(
                f"Unusually large word count reduction ({pct_change:.1f}%). Possible truncation."
            )
        elif pct_change > 10.0:
            warnings_list.append(
                f"Unusually large word count increase ({pct_change:+.1f}%). Possible duplication."
            )

        # Verify key financial sentiment words remain intact
        token_set = set(tokens)

        # Check for top-line terminology (accounting conventions use 'Revenue' or 'Net Sales')
        if not any(w in token_set for w in ("REVENUE", "REVENUES", "SALES")):
            warnings_list.append("Missing expected top-line terminology (REVENUE or SALES).")

        # Essential corporate financial disclosure vocabulary
        essential_vocab = ["OPERATING", "FINANCIAL", "RESULTS", "RISK", "MANAGEMENT"]
        missing_vocab = [w for w in essential_vocab if w not in token_set]
        if missing_vocab:
            warnings_list.append(
                f"Missing expected standard financial vocabulary: {', '.join(missing_vocab)}."
            )

        # Verify negation words are not removed
        negation_terms = ["NOT", "NO"]
        if not any(neg in token_set for neg in negation_terms):
            warnings_list.append("No standard negation words found in cleaned tokens.")

        if not warnings_list:
            status = "Verified"
        elif any("Failed" in w or "truncation" in w for w in warnings_list):
            status = "Failed"
        else:
            status = "Verified (with warnings)"

        return status, warnings_list

    def clean_mda_file(
        self,
        file_path: Path,
        ticker: str,
        fiscal_year: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Clean an individual MD&A text file and write the cleaned version to disk.

        Args:
            file_path (Path): Path to the source extracted MD&A text file.
            ticker (str): Normalized company ticker symbol.
            fiscal_year (Optional[str]): Fiscal reporting year. If None, parses from filename.

        Returns:
            Dict[str, Any]: Quality audit record with word counts and status.
        """
        clean_ticker = ticker.strip().upper()

        if not fiscal_year:
            # Parse year from filename e.g. AAPL_2025_MDA.txt
            match = re.search(r"_(\d{4})_MDA", file_path.name)
            fiscal_year = match.group(1) if match else "UNKNOWN"

        logger.info(f"Cleaning {clean_ticker} FY{fiscal_year} from {file_path.name}...")

        try:
            raw_text = file_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            raw_text = file_path.read_text(encoding="utf-8", errors="replace")
            logger.warning(f"UTF-8 decoding issues encountered in {file_path.name}; replaced invalid bytes.")

        cleaned_text = self.clean_text(raw_text)
        tokens = self.tokenize_for_sentiment(cleaned_text)

        orig_word_count = len(raw_text.split())
        clean_word_count = len(cleaned_text.split())
        token_count = len(tokens)

        pct_change = (
            ((clean_word_count - orig_word_count) / orig_word_count) * 100.0
            if orig_word_count else 0.0
        )

        status, warnings_list = self.validate_cleaned_text(
            original_text=raw_text,
            cleaned_text=cleaned_text,
            tokens=tokens,
            fiscal_year=fiscal_year,
        )

        # Output directory
        company_cleaned_dir = self.cleaned_dir / clean_ticker
        company_cleaned_dir.mkdir(parents=True, exist_ok=True)

        out_filename = f"{clean_ticker}_{fiscal_year}_MDA_CLEANED.txt"
        out_path = company_cleaned_dir / out_filename
        out_path.write_text(cleaned_text, encoding="utf-8")

        # Excerpts for audit and demonstration
        orig_excerpt = raw_text[:200].replace("\n", " ").strip()
        clean_excerpt = cleaned_text[:200].replace("\n", " ").strip()

        warnings_str = "; ".join(warnings_list) if warnings_list else "None"

        record: Dict[str, Any] = {
            "ticker": clean_ticker,
            "fiscal_year": fiscal_year,
            "source_mda_filename": file_path.name,
            "cleaned_filename": out_filename,
            "cleaned_file_path": str(out_path),
            "original_word_count": orig_word_count,
            "cleaned_word_count": clean_word_count,
            "sentiment_tokens_count": token_count,
            "pct_change_word_count": round(pct_change, 2),
            "cleaning_status": status,
            "warnings": warnings_str,
            "original_excerpt": orig_excerpt,
            "cleaned_excerpt": clean_excerpt,
        }

        logger.info(
            f"Cleaned {clean_ticker} FY{fiscal_year}: {orig_word_count:,} -> {clean_word_count:,} words "
            f"({pct_change:+.2f}%) | Tokens: {token_count:,} [{status}]"
        )
        return record

    def clean_company_mdas(self, ticker: str) -> Tuple[List[Dict[str, Any]], Path]:
        """Clean all extracted MD&A text files for a company and record CSV manifest.

        Args:
            ticker (str): Company ticker symbol.

        Returns:
            Tuple[List[Dict[str, Any]], Path]: Audit records and CSV manifest path.

        Raises:
            TextCleanerError: If no extracted MD&A files are found for the ticker.
        """
        clean_ticker = ticker.strip().upper()
        company_extracted_dir = self.extracted_dir / clean_ticker

        if not company_extracted_dir.exists():
            raise TextCleanerError(
                f"No extracted MD&A directory found for {clean_ticker} at {company_extracted_dir}. "
                "Please run Stage 3 extraction first."
            )

        source_files = sorted(
            [f for f in company_extracted_dir.glob("*_MDA.txt") if not f.name.endswith(".tmp")],
            reverse=True,
        )

        if not source_files:
            raise TextCleanerError(
                f"No extracted MD&A text files (*_MDA.txt) found in {company_extracted_dir}."
            )

        company_cleaned_dir = self.cleaned_dir / clean_ticker
        company_cleaned_dir.mkdir(parents=True, exist_ok=True)

        records: List[Dict[str, Any]] = []
        for file_path in source_files:
            rec = self.clean_mda_file(file_path=file_path, ticker=clean_ticker)
            records.append(rec)

        # Sort descending by fiscal year
        records.sort(key=lambda r: r.get("fiscal_year", ""), reverse=True)

        # Save manifest
        manifest_path = company_cleaned_dir / f"{clean_ticker}_cleaning_manifest.csv"
        self._save_manifest(records, manifest_path)

        return records, manifest_path

    @staticmethod
    def _save_manifest(records: List[Dict[str, Any]], manifest_path: Path) -> None:
        """Write text cleaning summary records to a CSV manifest file.

        Args:
            records (List[Dict[str, Any]]): List of audit record dicts.
            manifest_path (Path): Destination CSV path.
        """
        fieldnames = [
            "ticker",
            "fiscal_year",
            "source_mda_filename",
            "cleaned_filename",
            "cleaned_file_path",
            "original_word_count",
            "cleaned_word_count",
            "sentiment_tokens_count",
            "pct_change_word_count",
            "cleaning_status",
            "warnings",
        ]

        with open(manifest_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            for r in records:
                writer.writerow(r)

        logger.info(f"Saved text cleaning manifest to {manifest_path}")
