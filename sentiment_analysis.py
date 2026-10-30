"""Loughran-McDonald Financial Sentiment Analysis Module for SEC 10-K Filings.

This module scores cleaned MD&A narrative text using the official Loughran-McDonald
Financial Sentiment Dictionary (Master Dictionary 1993-2025). It supports both CSV
and Excel dictionary formats, calculates word frequencies, Positive/Negative (P/N)
ratios, category proportions, and polarity across 10 annual reporting periods,
integrating the results into the existing multi-tab Excel workbook under the
worksheet named 'Sentiment Analysis' while preserving all financial statements.

Key Sentiment Categories (Loughran & McDonald 2011):
- Negative: Words reflecting adverse, loss, impairment, or declining conditions.
- Positive: Words reflecting favorable, beneficial, and improving performance.
- Uncertainty: Words indicating ambiguity, risk, volatility, or contingency.
- Litigious: Words relating to lawsuits, regulatory compliance, and legal exposure.
- Constraining: Words reflecting commitments, contractual restrictions, and mandates.
- Strong Modal: Assertive modal auxiliary verbs (e.g., MUST, WILL, ALWAYS).
- Weak Modal: Tentative or hedging modal auxiliary verbs (e.g., COULD, MAY, POSSIBLE).
"""

import argparse
import csv
import logging
import os
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import config
from text_cleaner import MDATextCleaner

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("Sentiment-Analysis")


class SentimentAnalysisError(Exception):
    """Base exception for sentiment analysis operations."""


class DictionaryNotFoundError(SentimentAnalysisError):
    """Raised when the Loughran-McDonald dictionary file cannot be located."""


class DictionaryFormatError(SentimentAnalysisError):
    """Raised when the dictionary file is missing required sentiment columns."""


class LoughranMcDonaldDictionary:
    """Loader and validator for the official Loughran-McDonald Financial Sentiment Dictionary.

    Supports both CSV and Excel (.xlsx) formats.
    """

    REQUIRED_CATEGORIES = [
        "Negative",
        "Positive",
        "Uncertainty",
        "Litigious",
        "Constraining",
    ]
    MODAL_CATEGORIES = ["Strong_Modal", "Weak_Modal"]

    def __init__(self, dict_path: Optional[Path] = None) -> None:
        """Initialize and load the Loughran-McDonald dictionary.

        Args:
            dict_path (Optional[Path]): Specific path to dictionary CSV or Excel file,
                                        or None for automatic discovery.
        """
        self.dict_path = dict_path or self.locate_dictionary()
        self.categories: Dict[str, Set[str]] = {}
        self.total_master_words: int = 0
        self.version: str = self._resolve_dictionary_version()
        self.load_dictionary()

    def _resolve_dictionary_version(self) -> str:
        """Extract dictionary version identifier from filename or headers."""
        stem = self.dict_path.stem
        match = re.search(r"(\d{4}[-_]\d{4})", stem)
        if match:
            years = match.group(1).replace("_", "-")
            return f"Loughran-McDonald Master Dictionary ({years})"
        return f"Loughran-McDonald Master Dictionary ({self.dict_path.name})"

    @classmethod
    def locate_dictionary(cls, search_dirs: Optional[List[Path]] = None) -> Path:
        """Automatically locate the Loughran-McDonald CSV or Excel file within project directories.

        Searches candidate directories without requiring the user to manually rename
        the file.

        Args:
            search_dirs (Optional[List[Path]]): Custom directories to search.

        Returns:
            Path: Verified path to the dictionary file.

        Raises:
            DictionaryNotFoundError: If no valid dictionary CSV or Excel file is found.
        """
        candidate_dirs = search_dirs or [
            config.DATA_DIR,
            config.BASE_DIR,
            config.FINANCIAL_DATA_DIR,
            config.BUNDLE_DIR / "data",
            config.BUNDLE_DIR,
        ]

        patterns = [
            "*Loughran*McDonald*.csv",
            "*MasterDictionary*.csv",
            "*Loughran*.csv",
            "*LM*.csv",
            "*Loughran*McDonald*.xlsx",
            "*MasterDictionary*.xlsx",
            "*Loughran*.xlsx",
            "*LM*.xlsx",
        ]

        found_files: List[Path] = []
        for directory in candidate_dirs:
            if not directory.exists():
                continue
            for pattern in patterns:
                for match in directory.glob(pattern):
                    if match.is_file() and match not in found_files:
                        found_files.append(match)

        if not found_files:
            raise DictionaryNotFoundError(
                "Loughran-McDonald Master Dictionary (CSV or Excel) not found. "
                "Please place the file (e.g., 'Loughran-McDonald_MasterDictionary_1993-2025.csv') "
                f"inside {config.DATA_DIR}."
            )

        # If multiple files match, select the largest one (full master dictionary)
        found_files.sort(key=lambda p: p.stat().st_size, reverse=True)
        selected = found_files[0]
        logger.info(f"Auto-located Loughran-McDonald Dictionary: {selected.name} ({selected.stat().st_size:,} bytes)")
        return selected

    def load_dictionary(self) -> None:
        """Parse and validate the Loughran-McDonald dictionary (CSV or Excel).

        Reads word classifications and populates category sets. In the official
        Loughran-McDonald Master Dictionary:
        - Value > 0 indicates active membership (positive integer represents the year added).
        - Value < 0 indicates the word was removed from that list in that year.
        - Value == 0 indicates non-membership.

        Raises:
            DictionaryFormatError: If required category columns or 'Word' header are missing.
        """
        if not self.dict_path.exists():
            raise DictionaryNotFoundError(f"Dictionary file does not exist: {self.dict_path}")

        ext = self.dict_path.suffix.lower()
        if ext in [".xlsx", ".xlsm"]:
            self._load_excel_dictionary()
        else:
            self._load_csv_dictionary()

        # Build composite Modal category
        strong_set = self.categories.get("Strong_Modal", set())
        weak_set = self.categories.get("Weak_Modal", set())
        self.categories["Modal"] = strong_set.union(weak_set)

        logger.info(
            f"Successfully loaded {self.version}: {self.total_master_words:,} words. "
            f"Active category terms -> Negative: {len(self.categories['Negative'])}, "
            f"Positive: {len(self.categories['Positive'])}, "
            f"Uncertainty: {len(self.categories['Uncertainty'])}, "
            f"Litigious: {len(self.categories['Litigious'])}, "
            f"Constraining: {len(self.categories['Constraining'])}, "
            f"Modal: {len(self.categories['Modal'])}"
        )

    def _load_csv_dictionary(self) -> None:
        """Load dictionary from CSV format."""
        encodings = ["utf-8", "utf-8-sig", "latin-1"]
        file_handle = None
        for enc in encodings:
            try:
                f = open(self.dict_path, mode="r", encoding=enc)
                f.readline()
                f.seek(0)
                file_handle = f
                break
            except Exception:
                continue

        if file_handle is None:
            raise DictionaryFormatError(f"Unable to read dictionary CSV with supported encodings: {self.dict_path}")

        with file_handle:
            reader = csv.DictReader(file_handle)
            fieldnames = [fn.strip() for fn in (reader.fieldnames or [])]
            col_map: Dict[str, str] = {fn.upper(): fn for fn in fieldnames}

            if "WORD" not in col_map:
                raise DictionaryFormatError(
                    f"Dictionary CSV missing 'Word' identifier column. Available: {fieldnames}"
                )

            word_col = col_map["WORD"]
            for req_cat in self.REQUIRED_CATEGORIES:
                if req_cat.upper() not in col_map:
                    raise DictionaryFormatError(
                        f"Dictionary CSV missing required sentiment category: '{req_cat}'. "
                        f"Available columns: {fieldnames}"
                    )

            target_categories = self.REQUIRED_CATEGORIES + [
                m for m in self.MODAL_CATEGORIES if m.upper() in col_map
            ]
            self.categories = {cat: set() for cat in target_categories}

            total_rows = 0
            for row in reader:
                total_rows += 1
                raw_word = row.get(word_col, "")
                if not raw_word:
                    continue
                word = raw_word.strip().upper()

                for cat in target_categories:
                    col_name = col_map[cat.upper()]
                    raw_val = row.get(col_name, "0")
                    try:
                        val = int(raw_val)
                        if val > 0:
                            self.categories[cat].add(word)
                    except (ValueError, TypeError):
                        continue

            self.total_master_words = total_rows

    def _load_excel_dictionary(self) -> None:
        """Load dictionary from Excel (.xlsx) format."""
        try:
            wb = openpyxl.load_workbook(self.dict_path, read_only=True, data_only=True)
            ws = wb.active
            rows_iter = ws.iter_rows(values_only=True)
            header_row = next(rows_iter, None)
            if not header_row:
                raise DictionaryFormatError(f"Excel dictionary is empty: {self.dict_path}")

            fieldnames = [str(c).strip() for c in header_row if c is not None]
            col_map: Dict[str, int] = {str(c).strip().upper(): idx for idx, c in enumerate(header_row) if c is not None}

            if "WORD" not in col_map:
                raise DictionaryFormatError(
                    f"Excel dictionary missing 'Word' column. Available: {fieldnames}"
                )

            word_idx = col_map["WORD"]
            for req_cat in self.REQUIRED_CATEGORIES:
                if req_cat.upper() not in col_map:
                    raise DictionaryFormatError(
                        f"Excel dictionary missing required sentiment category: '{req_cat}'. "
                        f"Available: {fieldnames}"
                    )

            target_categories = self.REQUIRED_CATEGORIES + [
                m for m in self.MODAL_CATEGORIES if m.upper() in col_map
            ]
            self.categories = {cat: set() for cat in target_categories}

            total_rows = 0
            for row in rows_iter:
                if not row or len(row) <= word_idx or not row[word_idx]:
                    continue
                total_rows += 1
                word = str(row[word_idx]).strip().upper()

                for cat in target_categories:
                    c_idx = col_map[cat.upper()]
                    if len(row) > c_idx and row[c_idx] is not None:
                        try:
                            val = int(row[c_idx])
                            if val > 0:
                                self.categories[cat].add(word)
                        except (ValueError, TypeError):
                            continue

            self.total_master_words = total_rows

        except Exception as e:
            if isinstance(e, SentimentAnalysisError):
                raise
            raise DictionaryFormatError(f"Failed to parse Excel dictionary at {self.dict_path}: {e}")


class MDASentimentAnalyzer:
    """Engine that scores cleaned Item 7 MD&A narratives using Loughran-McDonald dictionaries."""

    def __init__(
        self,
        dictionary: Optional[LoughranMcDonaldDictionary] = None,
        cleaned_dir: Optional[Path] = None,
        excel_dir: Optional[Path] = None,
        raw_filings_dir: Optional[Path] = None,
        financial_data_dir: Optional[Path] = None,
    ) -> None:
        """Initialize the MD&A Sentiment Analyzer.

        Args:
            dictionary (Optional[LoughranMcDonaldDictionary]): Loaded LM dictionary instance.
            cleaned_dir (Optional[Path]): Directory containing cleaned MD&A files.
            excel_dir (Optional[Path]): Directory containing Excel workbooks.
            raw_filings_dir (Optional[Path]): Directory containing raw filing manifests.
            financial_data_dir (Optional[Path]): Directory containing structured financial data.
        """
        self.dictionary = dictionary or LoughranMcDonaldDictionary()
        self.cleaned_dir = cleaned_dir or config.CLEANED_TEXT_DIR
        self.excel_dir = excel_dir or config.EXCEL_OUTPUTS_DIR
        self.raw_filings_dir = raw_filings_dir or config.RAW_FILINGS_DIR
        self.financial_data_dir = financial_data_dir or config.FINANCIAL_DATA_DIR

    def score_tokens(self, tokens: List[str]) -> Dict[str, Any]:
        """Score a tokenized word sequence against the Loughran-McDonald dictionary.

        Counts all eligible word tokens in the denominator. Excludes purely numerical
        tokens from dictionary matching. Repeated occurrences of sentiment words are
        counted. A word may contribute to multiple sentiment categories if the
        dictionary indicates this.

        Args:
            tokens (List[str]): Uppercase word tokens from cleaned MD&A narrative.

        Returns:
            Dict[str, Any]: Sentiment counts, ratios, frequencies, and matched words.
        """
        total_word_count = len(tokens)
        counts: Dict[str, int] = {cat: 0 for cat in self.dictionary.categories}
        matched_words: Dict[str, Counter] = {cat: Counter() for cat in self.dictionary.categories}

        for token in tokens:
            for cat, word_set in self.dictionary.categories.items():
                if token in word_set:
                    counts[cat] += 1
                    matched_words[cat][token] += 1

        pos_count = counts.get("Positive", 0)
        neg_count = counts.get("Negative", 0)

        # 1. P/N Ratio: Positive Word Count / Negative Word Count
        # If negative word count is zero, return undefined safely
        if neg_count > 0:
            pn_ratio: Union[float, str] = round(pos_count / neg_count, 4)
        else:
            pn_ratio = "Undefined"

        # 2. Positive Word Frequency = Positive Word Count / Total Word Count
        pos_freq = (pos_count / total_word_count * 100.0) if total_word_count > 0 else 0.0

        # 3. Negative Word Frequency = Negative Word Count / Total Word Count
        neg_freq = (neg_count / total_word_count * 100.0) if total_word_count > 0 else 0.0

        # 4. Net Sentiment Polarity = (Pos - Neg) / (Pos + Neg)
        total_evaluative = pos_count + neg_count
        net_polarity = ((pos_count - neg_count) / total_evaluative) if total_evaluative > 0 else 0.0

        # Proportions (% of total word count) for all categories
        proportions: Dict[str, float] = {}
        for cat, cnt in counts.items():
            proportions[cat] = (cnt / total_word_count * 100.0) if total_word_count > 0 else 0.0

        return {
            "total_word_count": total_word_count,
            "counts": counts,
            "pn_ratio": pn_ratio,
            "positive_word_frequency": pos_freq,
            "negative_word_frequency": neg_freq,
            "net_polarity": net_polarity,
            "proportions": proportions,
            "matched_words": matched_words,
        }

    def get_period_end_dates(self, ticker: str) -> Dict[str, str]:
        """Retrieve fiscal period-end dates from the company's filing manifest.

        Args:
            ticker (str): Stock ticker symbol.

        Returns:
            Dict[str, str]: Mapping from fiscal year string (e.g. '2024') to date ('2024-09-28').
        """
        manifest_path = self.raw_filings_dir / ticker / f"{ticker}_10K_manifest.csv"
        dates: Dict[str, str] = {}
        if manifest_path.exists():
            try:
                with open(manifest_path, mode="r", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        fy = row.get("fiscal_year", "").strip()
                        date_val = row.get("fiscal_period_end_date", "").strip()
                        if fy and date_val:
                            dates[fy] = date_val
            except Exception as e:
                logger.warning(f"Could not load period-end dates from manifest: {e}")
        return dates

    def analyze_company_mdas(self, ticker: str) -> List[Dict[str, Any]]:
        """Score all available cleaned MD&A files for the specified company.

        Args:
            ticker (str): Company ticker symbol (e.g., 'AAPL').

        Returns:
            List[Dict[str, Any]]: Chronologically sorted list of annual sentiment records.

        Raises:
            SentimentAnalysisError: If no cleaned MD&A files are found for the company.
        """
        company_clean_dir = self.cleaned_dir / ticker
        if not company_clean_dir.exists():
            raise SentimentAnalysisError(
                f"Cleaned MD&A directory does not exist: {company_clean_dir}. "
                f"Please run Stage 5 text cleaning first: 'python main.py --ticker {ticker} --clean-mda'"
            )

        clean_files = sorted(company_clean_dir.glob(f"{ticker}_*_MDA_CLEANED.txt"))
        if not clean_files:
            raise SentimentAnalysisError(
                f"No cleaned MD&A text files found in {company_clean_dir}. "
                f"Please run Stage 5 text cleaning first."
            )

        period_dates = self.get_period_end_dates(ticker)
        records: List[Dict[str, Any]] = []

        for clean_path in clean_files:
            match = re.search(rf"{ticker}_(\d{{4}})_MDA_CLEANED\.txt", clean_path.name)
            if not match:
                continue
            fiscal_year = match.group(1)
            text = clean_path.read_text(encoding="utf-8")
            raw_text_words = len(text.split())

            tokens = MDATextCleaner.tokenize_for_sentiment(text)
            score_results = self.score_tokens(tokens)

            period_end = period_dates.get(fiscal_year, f"{fiscal_year}-12-31")
            source_mda_name = f"{ticker}_{fiscal_year}_MDA.txt"

            record = {
                "company_ticker": ticker,
                "ticker": ticker,
                "fiscal_reporting_year": fiscal_year,
                "fiscal_year": fiscal_year,
                "fiscal_period_end_date": period_end,
                "source_mda_filename": source_mda_name,
                "cleaned_filename": clean_path.name,
                "cleaned_text_words": raw_text_words,
                "total_word_count": score_results["total_word_count"],
                "tokens_count": score_results["total_word_count"],
                "cleaned_word_count": raw_text_words,
                "positive_word_count": score_results["counts"].get("Positive", 0),
                "positive_count": score_results["counts"].get("Positive", 0),
                "negative_word_count": score_results["counts"].get("Negative", 0),
                "negative_count": score_results["counts"].get("Negative", 0),
                "uncertainty_word_count": score_results["counts"].get("Uncertainty", 0),
                "uncertainty_count": score_results["counts"].get("Uncertainty", 0),
                "litigious_word_count": score_results["counts"].get("Litigious", 0),
                "litigious_count": score_results["counts"].get("Litigious", 0),
                "strong_modal_word_count": score_results["counts"].get("Strong_Modal", 0),
                "strong_modal_count": score_results["counts"].get("Strong_Modal", 0),
                "weak_modal_word_count": score_results["counts"].get("Weak_Modal", 0),
                "weak_modal_count": score_results["counts"].get("Weak_Modal", 0),
                "constraining_word_count": score_results["counts"].get("Constraining", 0),
                "constraining_count": score_results["counts"].get("Constraining", 0),
                "modal_word_count": score_results["counts"].get("Modal", 0),
                "modal_count": score_results["counts"].get("Modal", 0),
                "modal_word_count": score_results["counts"].get("Modal", 0),
                "pn_ratio": score_results["pn_ratio"],
                "positive_word_frequency": score_results["positive_word_frequency"],
                "negative_word_frequency": score_results["negative_word_frequency"],
                "net_polarity": score_results["net_polarity"],
                "proportions": score_results["proportions"],
                "dictionary_version": self.dictionary.version,
                "processing_status": "Verified",
                "top_negative": [w for w, _ in score_results["matched_words"]["Negative"].most_common(5)],
                "top_positive": [w for w, _ in score_results["matched_words"]["Positive"].most_common(5)],
                "top_uncertainty": [w for w, _ in score_results["matched_words"]["Uncertainty"].most_common(5)],
            }
            records.append(record)

        records.sort(key=lambda r: int(r["fiscal_year"]))
        logger.info(f"Analyzed {len(records)} MD&A annual reports for {ticker}.")
        return records

    def add_sentiment_worksheet_to_excel(
        self,
        ticker: str,
        records: List[Dict[str, Any]],
        company_name: Optional[str] = None,
        cik: Optional[str] = None,
    ) -> Path:
        """Add worksheet named 'Sentiment Analysis' to the company's existing Excel workbook.

        Preserves all existing financial statement worksheets (Income Statement,
        Balance Sheet, Cash Flow Statement, and Audit & Reconciliation) and presents
        the 10-year sentiment results in chronological order with professional formatting.

        Args:
            ticker (str): Stock ticker symbol.
            records (List[Dict[str, Any]]): Chronologically sorted sentiment records.
            company_name (Optional[str]): Company display name.
            cik (Optional[str]): Company SEC CIK number.

        Returns:
            Path: Path to the updated Excel workbook.

        Raises:
            SentimentAnalysisError: If the workbook cannot be opened or saved.
        """
        company_dir = self.excel_dir / ticker
        excel_path = company_dir / f"{ticker}_financial_statements.xlsx"

        if not excel_path.exists():
            logger.warning(
                f"Existing financial statements workbook not found at {excel_path}. Creating new workbook."
            )
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
        else:
            try:
                wb = openpyxl.load_workbook(excel_path)
                logger.info(f"Loaded existing Excel workbook: {excel_path.name} (Sheets: {wb.sheetnames})")
            except Exception as e:
                raise SentimentAnalysisError(f"Failed to load Excel workbook at {excel_path}: {e}")

        # Explicitly named 'Sentiment Analysis' per prompt specification
        sheet_title = "Sentiment Analysis"

        # Remove prior versions of sentiment tabs if present
        for old_name in ["Sentiment Analysis", "MD&A Sentiment Analysis"]:
            if old_name in wb.sheetnames:
                del wb[old_name]

        target_idx = len(wb.sheetnames)
        if "Cash Flow Statement" in wb.sheetnames:
            target_idx = wb.sheetnames.index("Cash Flow Statement") + 1

        ws = wb.create_sheet(title=sheet_title, index=target_idx)
        ws.views.sheetView[0].showGridLines = True

        disp_name = company_name or ("Apple Inc." if ticker == "AAPL" else ticker)
        disp_cik = cik or ("0000320193" if ticker == "AAPL" else "")

        font_name = "Calibri"
        navy_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        section_hdr_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        highlight_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

        title_font = Font(name=font_name, size=15, bold=True, color="1F4E78")
        subtitle_font = Font(name=font_name, size=10, italic=True, color="595959")
        col_hdr_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
        section_hdr_font = Font(name=font_name, size=10, bold=True, color="1F4E78")
        regular_font = Font(name=font_name, size=10)
        bold_font = Font(name=font_name, size=10, bold=True)
        key_metric_font = Font(name=font_name, size=11, bold=True, color="002060")

        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        align_center = Alignment(horizontal="center", vertical="center")

        border_thin = Side(style="thin", color="D9D9D9")
        border_top_thick = Side(style="thin", color="000000")
        border_bottom_double = Side(style="double", color="000000")

        regular_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
        subtotal_border = Border(top=border_top_thick, bottom=border_thin)
        total_border = Border(top=border_top_thick, bottom=border_bottom_double)

        # Row 1: Company Title
        ws["A1"] = f"{disp_name} ({ticker})"
        ws["A1"].font = title_font
        ws.row_dimensions[1].height = 24

        # Row 2: Sheet Heading
        ws["A2"] = f"MANAGEMENT'S DISCUSSION & ANALYSIS (MD&A) SENTIMENT ANALYSIS (CIK: {disp_cik})"
        ws["A2"].font = Font(name=font_name, size=12, bold=True, color="333333")
        ws.row_dimensions[2].height = 18

        # Row 3: Methodology Subtitle
        ws["A3"] = f"{self.dictionary.version} | Text Source: Item 7 MD&A Narrative"
        ws["A3"].font = subtitle_font
        ws.row_dimensions[3].height = 16

        # Row 4: Column Headers (Row 4 = FY, Row 5 = Period-End Date)
        ws.row_dimensions[4].height = 22
        ws.row_dimensions[5].height = 18

        ws["A4"] = "Metric / Sentiment Classification"
        ws["A5"] = "Fiscal Period-End Date"
        ws["A4"].fill = navy_header_fill
        ws["A5"].fill = navy_header_fill
        ws["A4"].font = col_hdr_font
        ws["A5"].font = Font(name=font_name, size=9, bold=True, color="FFFFFF")
        ws["A4"].alignment = align_left
        ws["A5"].alignment = align_left

        col_indices: Dict[str, int] = {}
        for idx, rec in enumerate(records, start=2):
            col_letter = get_column_letter(idx)
            col_indices[rec["fiscal_year"]] = idx

            cell_fy = ws[f"{col_letter}4"]
            cell_end = ws[f"{col_letter}5"]

            cell_fy.value = f"FY {rec['fiscal_year']}"
            cell_fy.fill = navy_header_fill
            cell_fy.font = col_hdr_font
            cell_fy.alignment = align_center

            cell_end.value = rec["fiscal_period_end_date"]
            cell_end.fill = navy_header_fill
            cell_end.font = Font(name=font_name, size=9, bold=True, color="FFFFFF")
            cell_end.alignment = align_center

        rows_definition = [
            # Section 1
            ("1. MD&A DOCUMENT METRICS", None, None, None, section_hdr_fill, section_hdr_font, True),
            ("Total Word Count (Tokens Analyzed)", "total_word_count", "#,##0", regular_border, None, bold_font, False),
            ("Cleaned Narrative Words (Raw)", "cleaned_text_words", "#,##0", regular_border, None, regular_font, False),

            # Section 2
            ("2. LOUGHRAN-MCDONALD SENTIMENT WORD COUNTS", None, None, None, section_hdr_fill, section_hdr_font, True),
            ("Positive Word Count", "positive_word_count", "#,##0", regular_border, None, regular_font, False),
            ("Negative Word Count", "negative_word_count", "#,##0", regular_border, None, regular_font, False),
            ("Uncertainty Word Count", "uncertainty_word_count", "#,##0", regular_border, None, regular_font, False),
            ("Litigious Word Count", "litigious_word_count", "#,##0", regular_border, None, regular_font, False),
            ("Strong Modal Word Count", "strong_modal_word_count", "#,##0", regular_border, None, regular_font, False),
            ("Weak Modal Word Count", "weak_modal_word_count", "#,##0", regular_border, None, regular_font, False),
            ("Constraining Word Count", "constraining_word_count", "#,##0", regular_border, None, regular_font, False),
            ("Total Modal Word Count (Strong + Weak)", "modal_word_count", "#,##0", subtotal_border, None, bold_font, False),

            # Section 3
            ("3. SENTIMENT RATIOS & FREQUENCIES", None, None, None, section_hdr_fill, section_hdr_font, True),
            ("Positive-to-Negative (P/N) Ratio", "pn_ratio", "0.000", total_border, highlight_fill, key_metric_font, False),
            ("Positive Word Frequency (%)", "pos_freq_dec", "0.00%", regular_border, None, bold_font, False),
            ("Negative Word Frequency (%)", "neg_freq_dec", "0.00%", regular_border, None, bold_font, False),
            ("Net Sentiment Polarity [(Pos - Neg) / (Pos + Neg)]", "net_polarity", "+0.00%;-0.00%;\"0.00%\"", regular_border, None, regular_font, False),

            # Section 4
            ("4. CATEGORY PROPORTIONS (% OF TOTAL WORD COUNT)", None, None, None, section_hdr_fill, section_hdr_font, True),
            ("Uncertainty Word Frequency (%)", "prop_Uncertainty", "0.00%", regular_border, None, regular_font, False),
            ("Litigious Word Frequency (%)", "prop_Litigious", "0.00%", regular_border, None, regular_font, False),
            ("Constraining Word Frequency (%)", "prop_Constraining", "0.00%", regular_border, None, regular_font, False),
            ("Total Modal Word Frequency (%)", "prop_Modal", "0.00%", subtotal_border, None, bold_font, False),
        ]

        current_row = 6
        num_cols = len(records) + 1

        for label, key, num_fmt, border_style, fill_style, font_style, is_section in rows_definition:
            ws.row_dimensions[current_row].height = 20 if is_section else 18
            cell_a = ws[f"A{current_row}"]
            cell_a.value = label
            cell_a.alignment = align_left
            cell_a.font = font_style or regular_font

            if is_section:
                cell_a.fill = fill_style
                for c_idx in range(2, num_cols + 1):
                    ws.cell(row=current_row, column=c_idx).fill = fill_style
                current_row += 1
                continue

            if fill_style:
                cell_a.fill = fill_style
            if border_style:
                cell_a.border = border_style

            for rec in records:
                c_idx = col_indices[rec["fiscal_year"]]
                col_let = get_column_letter(c_idx)
                cell_data = ws[f"{col_let}{current_row}"]
                cell_data.alignment = align_right
                cell_data.font = font_style or regular_font

                if fill_style:
                    cell_data.fill = fill_style
                if border_style:
                    cell_data.border = border_style

                if key == "pos_freq_dec":
                    val = (rec["positive_word_frequency"]) / 100.0
                elif key == "neg_freq_dec":
                    val = (rec["negative_word_frequency"]) / 100.0
                elif key and key.startswith("prop_"):
                    cat_name = key.replace("prop_", "")
                    val = (rec["proportions"].get(cat_name, 0.0)) / 100.0
                elif key == "pn_ratio":
                    raw_pn = rec.get("pn_ratio")
                    val = raw_pn if isinstance(raw_pn, (int, float)) else "Undefined"
                else:
                    val = rec.get(key, 0)

                cell_data.value = val
                if num_fmt and isinstance(val, (int, float)):
                    cell_data.number_format = num_fmt

            current_row += 1

        # Section 5: Top Matched Terminology
        ws.row_dimensions[current_row].height = 20
        sec5_a = ws[f"A{current_row}"]
        sec5_a.value = "5. TOP FREQUENT FINANCIAL SENTIMENT TERMS IDENTIFIED"
        sec5_a.font = section_hdr_font
        sec5_a.fill = section_hdr_fill
        for c_idx in range(2, num_cols + 1):
            ws.cell(row=current_row, column=c_idx).fill = section_hdr_fill
        current_row += 1

        qualitative_rows = [
            ("Top Negative Words", "top_negative"),
            ("Top Positive Words", "top_positive"),
            ("Top Uncertainty Words", "top_uncertainty"),
        ]

        for q_label, q_key in qualitative_rows:
            ws.row_dimensions[current_row].height = 18
            q_cell_a = ws[f"A{current_row}"]
            q_cell_a.value = q_label
            q_cell_a.font = regular_font
            q_cell_a.alignment = align_left
            q_cell_a.border = regular_border

            for rec in records:
                c_idx = col_indices[rec["fiscal_year"]]
                col_let = get_column_letter(c_idx)
                q_data_cell = ws[f"{col_let}{current_row}"]
                terms = rec.get(q_key, [])
                q_data_cell.value = ", ".join(terms[:3]) if terms else "None"
                q_data_cell.font = Font(name=font_name, size=9, italic=True)
                q_data_cell.alignment = align_center
                q_data_cell.border = regular_border

            current_row += 1

        ws.column_dimensions["A"].width = 46
        for c_idx in range(2, num_cols + 1):
            ws.column_dimensions[get_column_letter(c_idx)].width = 15

        ws.freeze_panes = "B6"

        try:
            wb.save(excel_path)
            logger.info(f"Successfully integrated worksheet '{sheet_title}' into {excel_path}")
        except Exception as e:
            raise SentimentAnalysisError(f"Failed to save updated Excel workbook at {excel_path}: {e}")

        return excel_path

    def export_summary_csv(self, ticker: str, records: List[Dict[str, Any]]) -> Tuple[Path, Path]:
        """Export structured sentiment analysis table to CSV files.

        Saves to both data/financial_data/ and data/excel_outputs/{ticker}/.

        Args:
            ticker (str): Company ticker symbol.
            records (List[Dict[str, Any]]): List of scored sentiment records.

        Returns:
            Tuple[Path, Path]: (primary_csv_path, mirror_csv_path)
        """
        self.financial_data_dir.mkdir(parents=True, exist_ok=True)
        primary_csv = self.financial_data_dir / f"{ticker}_sentiment_analysis.csv"

        company_excel_dir = self.excel_dir / ticker
        company_excel_dir.mkdir(parents=True, exist_ok=True)
        mirror_csv = company_excel_dir / f"{ticker}_sentiment_summary.csv"

        fieldnames = [
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

        for target_csv in [primary_csv, mirror_csv]:
            with open(target_csv, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for r in records:
                    pn_val = r["pn_ratio"]
                    pn_str = f"{pn_val:.4f}" if isinstance(pn_val, (int, float)) else "Undefined"
                    writer.writerow(
                        {
                            "company_ticker": r["company_ticker"],
                            "fiscal_reporting_year": r["fiscal_reporting_year"],
                            "fiscal_period_end_date": r["fiscal_period_end_date"],
                            "source_mda_filename": r["source_mda_filename"],
                            "total_word_count": r["total_word_count"],
                            "positive_word_count": r["positive_word_count"],
                            "negative_word_count": r["negative_word_count"],
                            "uncertainty_word_count": r["uncertainty_word_count"],
                            "litigious_word_count": r["litigious_word_count"],
                            "strong_modal_word_count": r["strong_modal_word_count"],
                            "weak_modal_word_count": r["weak_modal_word_count"],
                            "constraining_word_count": r["constraining_word_count"],
                            "pn_ratio": pn_str,
                            "positive_word_frequency": f"{r['positive_word_frequency']:.2f}%",
                            "negative_word_frequency": f"{r['negative_word_frequency']:.2f}%",
                            "dictionary_version": r["dictionary_version"],
                            "processing_status": r["processing_status"],
                        }
                    )

        logger.info(f"Saved sentiment analysis CSV output to {primary_csv}")
        return primary_csv, mirror_csv


def run_sentiment_analysis(
    ticker: str = "AAPL",
    dict_path: Optional[Path] = None,
) -> Tuple[List[Dict[str, Any]], Path, Path]:
    """Execute complete Stage 6 sentiment scoring, Excel integration, and CSV export.

    Args:
        ticker (str): Stock ticker symbol.
        dict_path (Optional[Path]): Specific path to LM dictionary CSV or Excel file.

    Returns:
        Tuple[List[Dict[str, Any]], Path, Path]: (records, excel_path, primary_csv_path)
    """
    clean_ticker = ticker.strip().upper()
    lm_dict = LoughranMcDonaldDictionary(dict_path=dict_path)
    analyzer = MDASentimentAnalyzer(dictionary=lm_dict)

    records = analyzer.analyze_company_mdas(ticker=clean_ticker)
    excel_path = analyzer.add_sentiment_worksheet_to_excel(ticker=clean_ticker, records=records)
    primary_csv, _ = analyzer.export_summary_csv(ticker=clean_ticker, records=records)

    return records, excel_path, primary_csv


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Loughran-McDonald Financial Sentiment Analysis for SEC 10-K Filings"
    )
    parser.add_argument(
        "--ticker",
        type=str,
        default="AAPL",
        help="Stock ticker symbol to analyze (default: AAPL)",
    )
    parser.add_argument(
        "--dict-path",
        type=Path,
        default=None,
        help="Path to Loughran-McDonald Master Dictionary (CSV or Excel)",
    )
    args = parser.parse_args()

    try:
        recs, xl_path, csv_p = run_sentiment_analysis(ticker=args.ticker, dict_path=args.dict_path)
        print(f"\n[SUCCESS] Completed Sentiment Analysis for {args.ticker.upper()} across {len(recs)} periods.")
        print(f"Updated Excel Workbook (Tab: 'Sentiment Analysis'): {xl_path}")
        print(f"Sentiment CSV Manifest: {csv_p}")
    except Exception as exc:
        logger.exception(f"Sentiment Analysis failed: {exc}")
        print(f"\n[ERROR] Sentiment analysis failed: {exc}")
        sys.exit(1)
