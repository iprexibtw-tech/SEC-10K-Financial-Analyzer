"""SEC Form 10-K Item 7 (MD&A) Text Extraction Module.

This module extracts the substantive Item 7 Management's Discussion and Analysis
of Financial Condition and Results of Operations (MD&A) section from downloaded
SEC 10-K filings, ending before Item 7A (or Item 8 where Item 7A is omitted).

Features:
- Robust HTML parsing preserving financial figures, paragraphs, and structure.
- Discrimination between true MD&A content, Table of Contents, and cross-references.
- Output saved as individual UTF-8 text files named {TICKER}_{YEAR}_MDA.txt.
- Rigorous validation (boundary verification, word count checks, TOC contamination detection).
- Generation of an extraction audit manifest (CSV).
"""

import csv
import logging
import re
import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning

import config

# Suppress XMLParsedAsHTMLWarning when parsing XML/inline XBRL SEC filings with lxml
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

logger = logging.getLogger("MDA-Extractor")


class MDAExtractionError(Exception):
    """Base exception for MD&A extraction errors."""


class MDAExtractor:
    """Extracts, validates, and stores Item 7 MD&A sections from 10-K filings."""

    def __init__(
        self,
        raw_filings_dir: Optional[Path] = None,
        extracted_text_dir: Optional[Path] = None,
    ) -> None:
        """Initialize the extractor with source and output directories.

        Args:
            raw_filings_dir (Optional[Path]): Directory containing raw filings.
            extracted_text_dir (Optional[Path]): Directory for saving extracted text.
        """
        self.raw_dir = raw_filings_dir or config.RAW_FILINGS_DIR
        self.extracted_dir = extracted_text_dir or config.EXTRACTED_TEXT_DIR

    @staticmethod
    def _clean_html_to_lines(html_content: str) -> List[str]:
        """Convert SEC filing HTML into clean, structure-preserving lines.

        Decomposes scripts, styles, and inline XBRL metadata tags, converts
        block-level containers into breaks, decodes HTML entities, and strips
        empty lines while preserving paragraph boundaries.

        Args:
            html_content (str): Raw HTML string of the 10-K filing.

        Returns:
            List[str]: List of non-empty cleaned text lines.
        """
        soup = BeautifulSoup(html_content, "lxml")

        # Decompose non-content and metadata elements
        for tag in soup(["script", "style", "ix:header", "ix:hidden", "noscript"]):
            tag.decompose()

        # Extract text using newline separator to preserve vertical flow
        raw_text = soup.get_text("\n")

        # Normalize non-breaking spaces and unicode apostrophes
        normalized = (
            raw_text.replace("\xa0", " ")
            .replace("&#160;", " ")
            .replace("&nbsp;", " ")
        )

        lines: List[str] = []
        for line in normalized.split("\n"):
            cleaned = line.strip()
            if cleaned:
                lines.append(cleaned)

        return lines

    @staticmethod
    def _is_table_of_contents_entry(lines: List[str], index: int) -> bool:
        """Determine if a heading candidate is a Table of Contents entry.

        Table of contents entries appear early in the document (within the first
        25% of lines), are clustered closely within 1-4 lines of adjacent Items,
        and often include page number references.

        Args:
            lines (List[str]): All document lines.
            index (int): Line index of the candidate heading.

        Returns:
            bool: True if candidate is in the TOC, False otherwise.
        """
        total = len(lines)
        # TOC never appears in the second half or beyond the first 35% of a 10-K
        if total > 50 and index > total * 0.35:
            return False

        # Look ahead up to 5 lines for consecutive Item headings (e.g. Item 7 followed by 7A or 8)
        lookahead_window = lines[index + 1 : min(total, index + 6)]
        for ahead in lookahead_window:
            norm_ahead = re.sub(r"\s+", " ", ahead)
            if re.match(r"^ITEM\s*(?:7A|8|9|10)\b", norm_ahead, re.IGNORECASE):
                return True

        # Check if line ends with a TOC page number (e.g. 'Item 7 ... 21')
        combined = lines[index]
        if index + 1 < total:
            combined += " " + lines[index + 1]
        if re.search(r"\b(?:Item\s*7A?|Operations)\s+\d{1,3}$", combined.strip(), re.IGNORECASE):
            return True

        return False

    def find_mda_boundaries(
        self, lines: List[str]
    ) -> Tuple[Optional[int], Optional[int], str, str]:
        """Locate start line of Item 7 and end line (start of Item 7A or Item 8).

        Args:
            lines (List[str]): Document lines.

        Returns:
            Tuple[Optional[int], Optional[int], str, str]:
                (start_idx, end_idx, start_label, end_label)
        """
        item7_candidates: List[int] = []
        item7a_candidates: List[int] = []
        item8_candidates: List[int] = []

        total_lines = len(lines)

        for i in range(total_lines):
            line = lines[i]

            # Look ahead 1-2 lines to handle headings split across tags/lines
            combined = line
            if i + 1 < total_lines:
                combined += " " + lines[i + 1]
            if i + 2 < total_lines:
                combined += " " + lines[i + 2]

            norm = re.sub(r"\s+", " ", combined).replace("\u2019", "'")

            # Match Item 7 heading: starts with Item 7 followed by Management's Discussion
            if re.match(
                r"^ITEM\s*7[\.\:\-\–\—\s]+Management'?s\s*Discussion",
                norm,
                re.IGNORECASE,
            ):
                item7_candidates.append(i)

            # Match Item 7A heading: starts with Item 7A followed by Quantitative
            elif re.match(
                r"^ITEM\s*7A[\.\:\-\–\—\s]+Quantitative",
                norm,
                re.IGNORECASE,
            ):
                item7a_candidates.append(i)

            # Match Item 8 heading: starts with Item 8 followed by Financial Statements
            elif re.match(
                r"^ITEM\s*8[\.\:\-\–\—\s]+Financial\s*Statements",
                norm,
                re.IGNORECASE,
            ):
                item8_candidates.append(i)

        if not item7_candidates:
            return None, None, "", ""

        # Select substantive Item 7 (filter out TOC entries)
        substantive_item7: Optional[int] = None
        for cand in item7_candidates:
            if not self._is_table_of_contents_entry(lines, cand):
                substantive_item7 = cand
                break

        # Fallback: if all matched look like TOC, pick the latest occurrence
        if substantive_item7 is None and item7_candidates:
            substantive_item7 = item7_candidates[-1]

        if substantive_item7 is None:
            return None, None, "", ""

        # Locate end boundary: First Item 7A appearing strictly after substantive Item 7
        end_idx: Optional[int] = None
        end_label = ""
        for cand in item7a_candidates:
            if cand > substantive_item7:
                end_idx = cand
                end_label = "Item 7A"
                break

        # Fallback to Item 8 if Item 7A is not found
        if end_idx is None:
            for cand in item8_candidates:
                if cand > substantive_item7:
                    end_idx = cand
                    end_label = "Item 8"
                    break

        start_label = "Item 7"
        return substantive_item7, end_idx, start_label, end_label

    def validate_extraction(
        self,
        extracted_text: str,
        start_line_text: str,
        end_label: str,
        total_source_lines: int,
        extracted_line_count: int,
    ) -> Tuple[str, List[str]]:
        """Perform comprehensive quality control and validation on extracted MD&A.

        Args:
            extracted_text (str): The extracted MD&A text.
            start_line_text (str): Content of the starting line.
            end_label (str): Label of the ending boundary ('Item 7A' or 'Item 8').
            total_source_lines (int): Total line count of the source filing.
            extracted_line_count (int): Lines extracted for MD&A.

        Returns:
            Tuple[str, List[str]]: (validation_status, list_of_warnings)
        """
        warnings_list: List[str] = []

        if not extracted_text or not extracted_text.strip():
            return "Failed", ["Extracted text is empty."]

        words = len(extracted_text.split())

        # Check word count boundaries
        if words < 1000:
            warnings_list.append(
                f"Suspiciously short MD&A extraction ({words:,} words). Possible TOC or premature cutoff."
            )
        elif words > 45000:
            warnings_list.append(
                f"Unusually long MD&A extraction ({words:,} words). Possible boundary overshoot."
            )

        # Check beginning boundary integrity (first 250 characters of extracted text)
        norm_start = re.sub(r"\s+", " ", extracted_text[:250]).replace("\u2019", "'")
        if not re.search(
            r"^Item\s*7[\.\:\-\–\—\s]+Management'?s\s*Discussion",
            norm_start,
            re.IGNORECASE,
        ):
            warnings_list.append(
                f"Starting boundary does not strictly match Item 7 heading: '{norm_start[:80]}'."
            )

        # Check for Table of Contents contamination
        toc_pattern = r"Item\s*1\.\s+Business.*Item\s*2\.\s+Properties.*Item\s*7\."
        if re.search(toc_pattern, extracted_text[:3000], re.DOTALL | re.IGNORECASE):
            warnings_list.append("Detected potential Table of Contents contamination in extraction.")

        # Check for boundary inclusion error: Item 7A heading inside extracted text
        # (Excluding the very start, which is Item 7)
        if re.search(r"\nITEM\s*7A[\.\:\s]+Quantitative", extracted_text[500:], re.IGNORECASE):
            warnings_list.append("Item 7A heading found inside extracted text body.")

        # Check that an end boundary was identified
        if not end_label:
            warnings_list.append("Ending boundary (Item 7A or Item 8) was not identified.")

        if not warnings_list:
            status = "Verified"
        elif any("Failed" in w or "empty" in w for w in warnings_list):
            status = "Failed"
        else:
            status = "Verified (with warnings)"

        return status, warnings_list

    def extract_mda_from_file(
        self,
        file_path: Path,
        ticker: str,
        fiscal_year: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Extract MD&A section from an individual 10-K filing file and save to disk.

        Args:
            file_path (Path): Path to the source 10-K filing document.
            ticker (str): Normalized company ticker.
            fiscal_year (Optional[str]): Fiscal reporting year. If None, parses from filename.

        Returns:
            Dict[str, Any]: Extraction audit record with metrics and status.
        """
        clean_ticker = ticker.strip().upper()

        # Parse fiscal year from filename if not provided
        # Pattern: TICKER_10K_FY{YYYY}_{DATE}.htm
        if not fiscal_year:
            match = re.search(r"FY(\d{4})", file_path.name)
            if match:
                fiscal_year = match.group(1)
            else:
                # Fallback: look for 4-digit year in filename
                match_year = re.search(r"(\d{4})", file_path.name)
                fiscal_year = match_year.group(1) if match_year else "UNKNOWN"

        logger.info(f"Processing {clean_ticker} FY{fiscal_year} from {file_path.name}...")

        html_content = file_path.read_text(encoding="utf-8", errors="ignore")
        lines = self._clean_html_to_lines(html_content)

        start_idx, end_idx, start_label, end_label = self.find_mda_boundaries(lines)

        if start_idx is None or end_idx is None or start_idx >= end_idx:
            logger.error(
                f"Failed to identify valid MD&A boundaries in {file_path.name} "
                f"(Start: {start_idx}, End: {end_idx})"
            )
            return {
                "ticker": clean_ticker,
                "fiscal_year": fiscal_year,
                "source_filename": file_path.name,
                "extracted_filename": "",
                "extracted_file_path": "",
                "word_count": 0,
                "start_line": start_idx if start_idx is not None else -1,
                "end_line": end_idx if end_idx is not None else -1,
                "end_boundary_label": end_label,
                "status": "Failed (Boundaries Not Found)",
                "warnings": "Could not identify Item 7 and Item 7A/8 section boundaries.",
                "beginning_excerpt": "",
                "ending_excerpt": "",
            }

        # Extract lines
        mda_lines = lines[start_idx:end_idx]
        extracted_text = "\n\n".join(mda_lines)
        word_count = len(extracted_text.split())

        start_line_text = lines[start_idx]
        status, warnings_list = self.validate_extraction(
            extracted_text=extracted_text,
            start_line_text=start_line_text,
            end_label=end_label,
            total_source_lines=len(lines),
            extracted_line_count=len(mda_lines),
        )

        # Excerpts for verification report
        beg_excerpt = extracted_text[:220].replace("\n", " ").strip()
        end_excerpt = extracted_text[-220:].replace("\n", " ").strip()

        # Save to company-specific directory in extracted_text/
        company_out_dir = self.extracted_dir / clean_ticker
        company_out_dir.mkdir(parents=True, exist_ok=True)

        out_filename = f"{clean_ticker}_{fiscal_year}_MDA.txt"
        out_path = company_out_dir / out_filename
        out_path.write_text(extracted_text, encoding="utf-8")

        warnings_str = "; ".join(warnings_list) if warnings_list else "None"

        record: Dict[str, Any] = {
            "ticker": clean_ticker,
            "fiscal_year": fiscal_year,
            "source_filename": file_path.name,
            "extracted_filename": out_filename,
            "extracted_file_path": str(out_path),
            "word_count": word_count,
            "start_line": start_idx,
            "end_line": end_idx,
            "end_boundary_label": end_label,
            "status": status,
            "warnings": warnings_str,
            "beginning_excerpt": beg_excerpt,
            "ending_excerpt": end_excerpt,
        }

        logger.info(
            f"Extracted {clean_ticker} FY{fiscal_year} ({word_count:,} words) -> {out_filename} [{status}]"
        )
        return record

    def extract_company_mdas(self, ticker: str) -> Tuple[List[Dict[str, Any]], Path]:
        """Extract MD&A sections for all downloaded 10-K filings of a company.

        Args:
            ticker (str): Company ticker symbol.

        Returns:
            Tuple[List[Dict[str, Any]], Path]: List of extraction records and manifest path.

        Raises:
            MDAExtractionError: If no downloaded raw filings are found for the ticker.
        """
        clean_ticker = ticker.strip().upper()
        company_raw_dir = self.raw_dir / clean_ticker

        if not company_raw_dir.exists():
            raise MDAExtractionError(
                f"No downloaded filings directory found for {clean_ticker} at {company_raw_dir}. "
                "Please run Stage 2 retrieval first."
            )

        filing_files = sorted(
            [f for f in company_raw_dir.glob("*.htm*") if not f.name.endswith(".tmp")],
            reverse=True,
        )

        if not filing_files:
            raise MDAExtractionError(
                f"No 10-K HTML filing files found in {company_raw_dir}."
            )

        company_out_dir = self.extracted_dir / clean_ticker
        company_out_dir.mkdir(parents=True, exist_ok=True)

        records: List[Dict[str, Any]] = []
        for file_path in filing_files:
            record = self.extract_mda_from_file(file_path=file_path, ticker=clean_ticker)
            records.append(record)

        # Sort records descending by fiscal year
        records.sort(key=lambda r: r.get("fiscal_year", ""), reverse=True)

        # Write extraction manifest
        manifest_path = company_out_dir / f"{clean_ticker}_mda_manifest.csv"
        self._save_manifest(records, manifest_path)

        return records, manifest_path

    @staticmethod
    def _save_manifest(records: List[Dict[str, Any]], manifest_path: Path) -> None:
        """Write extraction summary records to CSV manifest.

        Args:
            records (List[Dict[str, Any]]): List of extraction audit dicts.
            manifest_path (Path): Target CSV path.
        """
        fieldnames = [
            "ticker",
            "fiscal_year",
            "source_filename",
            "extracted_filename",
            "extracted_file_path",
            "word_count",
            "start_line",
            "end_line",
            "end_boundary_label",
            "status",
            "warnings",
        ]

        with open(manifest_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            for r in records:
                writer.writerow(r)

        logger.info(f"Saved MD&A manifest to {manifest_path}")
