"""SEC EDGAR 10-K Filing Retrieval Module.

This module retrieves company CIK numbers, identifies the most recent annual
Form 10-K filings from the official SEC EDGAR submissions API, downloads the
primary filing documents, organizes them into company-specific directories,
and records a verifiable CSV manifest.

Complies with SEC Fair Access policies by enforcing rate limits, retries with
exponential backoff, and compliant User-Agent headers.
"""

import csv
import logging
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests

import config

logger = logging.getLogger("SEC-Downloader")


class SECDownloaderError(Exception):
    """Base exception for SEC downloader operations."""


class SECAuthenticationError(SECDownloaderError):
    """Raised when SEC rejects request due to invalid or missing User-Agent."""


class SECCompanyNotFoundError(SECDownloaderError):
    """Raised when a company ticker symbol cannot be resolved."""


class SECFilingsNotFoundError(SECDownloaderError):
    """Raised when no matching Form 10-K filings can be found."""


class SECRateLimitError(SECDownloaderError):
    """Raised when SEC rate limit is exceeded repeatedly."""


class SECDownloader:
    """Manages SEC EDGAR queries, document downloads, and rate limiting."""

    def __init__(self, user_agent: Optional[str] = None) -> None:
        """Initialize the SEC Downloader with a validated User-Agent.

        Args:
            user_agent (Optional[str]): SEC-compliant User-Agent header string.
                If None, loads from environment or local configuration.

        Raises:
            SECAuthenticationError: If no valid User-Agent is provided or found.
        """
        resolved_agent = user_agent or config.get_sec_user_agent()
        if not resolved_agent:
            raise SECAuthenticationError(
                "A descriptive SEC-compliant User-Agent with contact information "
                "is required by SEC EDGAR fair access policies before making requests. "
                "Format: 'YourName/Organization your_email@domain.com'"
            )

        self.user_agent = resolved_agent.strip()
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": self.user_agent,
                "Accept-Encoding": "gzip, deflate",
                "Accept": "application/json, text/html, */*",
            }
        )
        self._last_request_time: float = 0.0
        self._min_request_interval: float = config.SEC_REQUEST_DELAY

    def _rate_limit_wait(self) -> None:
        """Enforce delay between requests to strictly respect SEC rate limits."""
        elapsed = time.time() - self._last_request_time
        if elapsed < self._min_request_interval:
            sleep_time = self._min_request_interval - elapsed
            time.sleep(sleep_time)
        self._last_request_time = time.time()

    def _get(self, url: str, stream: bool = False) -> requests.Response:
        """Execute an HTTP GET request with rate limiting and exponential backoff.

        Args:
            url (str): Target URL to query.
            stream (bool): Whether to stream the response content.

        Returns:
            requests.Response: Successful HTTP response.

        Raises:
            SECAuthenticationError: On HTTP 403 Forbidden.
            SECRateLimitError: On persistent HTTP 429 Too Many Requests.
            SECDownloaderError: On network failures or unrecoverable HTTP errors.
        """
        for attempt in range(1, config.SEC_MAX_RETRIES + 1):
            self._rate_limit_wait()
            try:
                response = self.session.get(
                    url,
                    timeout=config.SEC_REQUEST_TIMEOUT,
                    stream=stream,
                )

                if response.status_code == 200:
                    return response

                if response.status_code == 403:
                    raise SECAuthenticationError(
                        f"SEC EDGAR returned HTTP 403 Forbidden for {url}. "
                        f"Ensure your User-Agent header '{self.user_agent}' contains "
                        "a valid, recognizable name and email address."
                    )

                if response.status_code == 404:
                    raise SECDownloaderError(f"SEC resource not found (404): {url}")

                if response.status_code == 429:
                    backoff = (config.SEC_BACKOFF_FACTOR ** attempt) * 2.0
                    logger.warning(
                        f"SEC Rate limit hit (429). Backing off for {backoff:.2f}s "
                        f"(Attempt {attempt}/{config.SEC_MAX_RETRIES})..."
                    )
                    time.sleep(backoff)
                    continue

                if response.status_code in (500, 502, 503, 504):
                    backoff = (config.SEC_BACKOFF_FACTOR ** attempt) * 1.5
                    logger.warning(
                        f"SEC server error ({response.status_code}). Retrying in {backoff:.2f}s "
                        f"(Attempt {attempt}/{config.SEC_MAX_RETRIES})..."
                    )
                    time.sleep(backoff)
                    continue

                response.raise_for_status()

            except (requests.ConnectionError, requests.Timeout) as conn_err:
                if attempt == config.SEC_MAX_RETRIES:
                    raise SECDownloaderError(
                        f"Failed to connect to SEC EDGAR at {url} after "
                        f"{config.SEC_MAX_RETRIES} attempts: {conn_err}"
                    ) from conn_err
                backoff = (config.SEC_BACKOFF_FACTOR ** attempt) * 1.5
                logger.warning(f"Connection error to {url}: {conn_err}. Retrying in {backoff:.2f}s...")
                time.sleep(backoff)

        raise SECRateLimitError(
            f"SEC EDGAR request to {url} failed after {config.SEC_MAX_RETRIES} retries."
        )

    def get_tickers_registry(self) -> Dict[str, Any]:
        """Load or fetch the official SEC company tickers registry.

        Returns:
            Dict[str, Any]: Dictionary of SEC registered companies containing
                'cik_str', 'ticker', and 'title'.
        """
        import json

        tickers_cache_path = config.FINANCIAL_DATA_DIR / "company_tickers.json"
        data: Optional[Dict[str, Any]] = None

        # Check local cache first (valid if cached within 7 days)
        if tickers_cache_path.exists():
            try:
                file_age = time.time() - tickers_cache_path.stat().st_mtime
                if file_age < 7 * 86400:
                    with open(tickers_cache_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
            except Exception:
                data = None

        # Check bundled copy as fallback before fetching from network
        if not data and config.IS_FROZEN:
            bundled_tickers = config.BUNDLE_DIR / "data" / "financial_data" / "company_tickers.json"
            if bundled_tickers.exists():
                try:
                    with open(bundled_tickers, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    data = None

        # Fetch fresh copy from SEC if not cached or cache is stale
        if not data:
            logger.info("Fetching SEC company tickers registry...")
            resp = self._get(config.SEC_COMPANY_TICKERS_URL)
            data = resp.json()
            # Save local cache safely
            try:
                with open(tickers_cache_path, "w", encoding="utf-8") as f:
                    json.dump(data, f)
            except OSError as err:
                logger.warning(f"Could not cache company tickers locally: {err}")

        return data or {}

    def get_company_cik(self, ticker: str) -> Tuple[str, str]:
        """Lookup company CIK and official company title from SEC ticker data.

        Args:
            ticker (str): Stock ticker symbol (e.g. 'AAPL').

        Returns:
            Tuple[str, str]: 10-digit CIK string and official company title.

        Raises:
            SECCompanyNotFoundError: If ticker symbol is not in SEC's registry.
        """
        clean_ticker = ticker.strip().upper()
        data = self.get_tickers_registry()

        # Search for ticker
        for entry in data.values():
            if entry.get("ticker", "").upper() == clean_ticker:
                cik_int = entry.get("cik_str")
                company_title = entry.get("title", "")
                cik_10_digits = str(cik_int).zfill(10)
                logger.info(
                    f"Resolved ticker '{clean_ticker}' -> CIK: {cik_10_digits} "
                    f"({company_title})"
                )
                return cik_10_digits, company_title

        raise SECCompanyNotFoundError(
            f"Ticker symbol '{clean_ticker}' was not found in SEC EDGAR registry. "
            "Please check the ticker spelling and try again."
        )

    def search_companies(self, query: str, limit: int = 10) -> List[Dict[str, str]]:
        """Search for publicly traded companies by name or ticker using official SEC registry.

        Supports direct ticker matching, company name prefix matching, word boundary
        matching, and substring matching.

        Args:
            query (str): Company name or ticker query (e.g. 'Apple', 'NVIDIA', 'Microsoft').
            limit (int): Maximum number of ranked results to return (default: 10).

        Returns:
            List[Dict[str, str]]: Ranked list of matching companies with keys:
                'ticker', 'title', 'cik'.
        """
        q = query.strip()
        if not q:
            return []

        q_upper = q.upper()
        q_words = [w for w in re.split(r"\s+", q_upper) if w]

        # Common parent company / brand aliases for user convenience
        aliases = {
            "GOOGLE": "GOOGL",
            "FACEBOOK": "META",
        }
        alias_ticker = aliases.get(q_upper)

        data = self.get_tickers_registry()
        results: List[Dict[str, Any]] = []
        seen_tickers = set()

        for entry in data.values():
            ticker = entry.get("ticker", "").upper()
            title = entry.get("title", "")
            title_upper = title.upper()
            cik_str = str(entry.get("cik_str", "")).zfill(10)

            if not ticker or ticker in seen_tickers:
                continue

            score = 0
            if ticker == q_upper:
                score += 1000
            elif alias_ticker and ticker == alias_ticker:
                score += 900
            elif title_upper == q_upper:
                score += 800
            elif title_upper.startswith(q_upper):
                score += 500
            elif re.search(r"\b" + re.escape(q_upper) + r"\b", title_upper):
                score += 300
            elif len(q_words) > 1 and all(
                re.search(r"\b" + re.escape(w) + r"\b", title_upper) for w in q_words
            ):
                score += 200
            elif q_upper in title_upper:
                score += 100
            elif q_upper in ticker:
                score += 50
            else:
                continue

            seen_tickers.add(ticker)
            results.append({
                "ticker": ticker,
                "title": title,
                "cik": cik_str,
                "score": score,
            })

        # Sort descending by relevance score, then ascending by title length (prefer cleaner titles), then alphabetically
        results.sort(key=lambda x: (-x["score"], len(x["title"]), x["title"]))
        return [
            {"ticker": r["ticker"], "title": r["title"], "cik": r["cik"]}
            for r in results[:limit]
        ]

    def _extract_10ks_from_submission_data(
        self,
        recent_data: Dict[str, List[Any]],
        cik: str,
        ticker: str,
        company_name: str,
    ) -> List[Dict[str, Any]]:
        """Extract Form 10-K entries from an SEC submissions data block.

        Args:
            recent_data (Dict[str, List[Any]]): Submission array columns.
            cik (str): 10-digit CIK string.
            ticker (str): Normalized ticker symbol.
            company_name (str): Official company name.

        Returns:
            List[Dict[str, Any]]: List of parsed 10-K filing metadata dicts.
        """
        forms = recent_data.get("form", [])
        filing_dates = recent_data.get("filingDate", [])
        report_dates = recent_data.get("reportDate", [])
        accessions = recent_data.get("accessionNumber", [])
        primary_docs = recent_data.get("primaryDocument", [])

        cik_int_str = str(int(cik))
        results: List[Dict[str, Any]] = []

        total_entries = len(forms)
        for i in range(total_entries):
            # Strict 10-K filter: exclude 10-K/A, 10-KT, etc.
            if forms[i] != "10-K":
                continue

            accession = accessions[i]
            filing_date = filing_dates[i]
            report_date = report_dates[i] if i < len(report_dates) else filing_date
            primary_doc = primary_docs[i] if i < len(primary_docs) else ""

            if not primary_doc:
                continue

            # Format accession number without hyphens for URL construction
            accession_no_hyphens = accession.replace("-", "")

            # Construct official SEC Archive URL
            source_url = (
                f"{config.SEC_ARCHIVES_BASE_URL}/{cik_int_str}/"
                f"{accession_no_hyphens}/{primary_doc}"
            )

            # Determine fiscal year (from report date YYYY-MM-DD or filing date)
            fiscal_year = report_date[:4] if report_date else filing_date[:4]

            # Extension preservation
            _, ext = os.path.splitext(primary_doc)
            if not ext:
                ext = ".htm"

            # Construct standardized descriptive filename
            saved_filename = f"{ticker}_10K_FY{fiscal_year}_{filing_date}{ext}"

            results.append(
                {
                    "ticker": ticker,
                    "cik": cik,
                    "company_name": company_name,
                    "form": "10-K",
                    "fiscal_year": fiscal_year,
                    "fiscal_period_end_date": report_date,
                    "filing_date": filing_date,
                    "accession_number": accession,
                    "primary_document": primary_doc,
                    "sec_source_url": source_url,
                    "saved_filename": saved_filename,
                }
            )

        return results

    def get_10k_filings_metadata(
        self, cik: str, ticker: str, company_name: str, count: int = 10
    ) -> List[Dict[str, Any]]:
        """Query SEC EDGAR submissions to identify the N most recent distinct 10-Ks.

        Args:
            cik (str): 10-digit CIK string.
            ticker (str): Normalized ticker symbol.
            company_name (str): Official company name.
            count (int): Number of annual filings requested (default: 10).

        Returns:
            List[Dict[str, Any]]: Metadata for the N most recent distinct 10-K filings.

        Raises:
            SECFilingsNotFoundError: If no Form 10-K filings are available.
        """
        submissions_url = config.SEC_SUBMISSIONS_URL_TEMPLATE.format(cik=cik)
        logger.info(f"Querying SEC submissions for CIK {cik}...")
        resp = self._get(submissions_url)
        sub_json = resp.json()

        filings_dict = sub_json.get("filings", {})
        recent = filings_dict.get("recent", {})

        raw_candidates = self._extract_10ks_from_submission_data(
            recent, cik, ticker, company_name
        )

        # Check older submission files if more filings are needed
        older_files = filings_dict.get("files", [])
        if len(raw_candidates) < count and older_files:
            logger.info("Checking older SEC submission files to reach requested count...")
            for old_file_info in older_files:
                old_file_name = old_file_info.get("name")
                if not old_file_name:
                    continue
                old_url = f"https://data.sec.gov/submissions/{old_file_name}"
                try:
                    old_resp = self._get(old_url)
                    old_data = old_resp.json()
                    older_10ks = self._extract_10ks_from_submission_data(
                        old_data, cik, ticker, company_name
                    )
                    raw_candidates.extend(older_10ks)
                    if len(raw_candidates) >= count:
                        break
                except Exception as old_err:
                    logger.warning(f"Could not retrieve older submission file {old_file_name}: {old_err}")

        if not raw_candidates:
            raise SECFilingsNotFoundError(
                f"No annual Form 10-K filings were found for {ticker} (CIK: {cik})."
            )

        # Deduplicate by distinct annual reporting period (report_date / fiscal year)
        # Sort candidates descending by filing date first
        raw_candidates.sort(key=lambda x: x["filing_date"], reverse=True)

        distinct_filings: List[Dict[str, Any]] = []
        seen_periods = set()

        for cand in raw_candidates:
            period_key = cand["fiscal_period_end_date"] or cand["fiscal_year"]
            if period_key not in seen_periods:
                seen_periods.add(period_key)
                distinct_filings.append(cand)

            if len(distinct_filings) == count:
                break

        actual_count = len(distinct_filings)
        if actual_count < count:
            logger.warning(
                f"Requested {count} filings, but only {actual_count} distinct annual "
                f"10-K filings are available for {ticker}."
            )
        else:
            logger.info(
                f"Identified {actual_count} distinct annual 10-K filings for {ticker}."
            )

        return distinct_filings

    def verify_filing_content(self, file_path: Path) -> Tuple[bool, str]:
        """Verify that a local filing document is valid, non-empty, and substantive.

        Args:
            file_path (Path): Path to the saved document.

        Returns:
            Tuple[bool, str]: (is_valid, verification_detail_message)
        """
        if not file_path.exists():
            return False, "File not found on disk"

        size = file_path.stat().st_size
        if size < 2048:
            return False, f"File size too small ({size} bytes)"

        try:
            # Inspect first 4KB and last 4KB for content verification
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                head = f.read(4096).lower()
                f.seek(max(0, size - 4096))
                tail = f.read(4096).lower()

            # Check for SEC error pages or block notifications
            error_signatures = [
                "your request has been blocked",
                "access denied",
                "403 forbidden",
                "sec edgar - error",
                "file not found",
            ]
            for sig in error_signatures:
                if sig in head:
                    return False, f"Content indicates SEC block or error ('{sig}')"

            # Check for substantive 10-K or HTML/text indicators
            valid_signatures = ["10-k", "annual report", "form 10-k", "<html", "<document", "table"]
            has_valid_sig = any(sig in head or sig in tail for sig in valid_signatures)
            if not has_valid_sig:
                return False, "Document missing standard 10-K or HTML structural markers"

            return True, f"Valid ({size:,} bytes)"

        except Exception as err:
            return False, f"Read error: {err}"

    def download_filing(
        self, filing_meta: Dict[str, Any], company_dir: Path
    ) -> Dict[str, Any]:
        """Download or verify an individual 10-K primary document.

        Avoids duplicate downloads if the file is already valid on disk.

        Args:
            filing_meta (Dict[str, Any]): Filing metadata dictionary.
            company_dir (Path): Company-specific directory for saving files.

        Returns:
            Dict[str, Any]: Updated metadata record with local path and status.
        """
        target_path = company_dir / filing_meta["saved_filename"]
        source_url = filing_meta["sec_source_url"]

        record = dict(filing_meta)
        record["saved_file_path"] = str(target_path)

        # 1. Check if file already exists and is valid (prevent duplicate downloads)
        if target_path.exists():
            is_valid, msg = self.verify_filing_content(target_path)
            if is_valid:
                record["verification_status"] = "Verified (Cached)"
                record["file_size_bytes"] = target_path.stat().st_size
                logger.info(
                    f"Preserved existing file [{record['saved_filename']}] - {msg}"
                )
                return record

            logger.warning(
                f"Existing file {target_path.name} failed verification ({msg}). Re-downloading..."
            )

        # 2. Download from SEC Archive
        logger.info(f"Downloading: {target_path.name} from {source_url}...")
        resp = self._get(source_url)
        content_bytes = resp.content

        # Write to disk
        target_path.write_bytes(content_bytes)

        # 3. Verify downloaded content
        is_valid, msg = self.verify_filing_content(target_path)
        if not is_valid:
            record["verification_status"] = f"Failed ({msg})"
            record["file_size_bytes"] = len(content_bytes)
            logger.error(f"Download verification failed for {target_path.name}: {msg}")
        else:
            record["verification_status"] = "Verified (Downloaded)"
            record["file_size_bytes"] = len(content_bytes)
            logger.info(f"Successfully downloaded and verified [{target_path.name}] ({msg})")

        return record

    def save_manifest(
        self, records: List[Dict[str, Any]], manifest_path: Path
    ) -> None:
        """Write records to a CSV manifest file.

        Args:
            records (List[Dict[str, Any]]): List of filing record dictionaries.
            manifest_path (Path): Destination CSV path.
        """
        fieldnames = [
            "ticker",
            "cik",
            "fiscal_year",
            "fiscal_period_end_date",
            "filing_date",
            "accession_number",
            "primary_document",
            "sec_source_url",
            "saved_filename",
            "saved_file_path",
            "file_size_bytes",
            "verification_status",
        ]

        with open(manifest_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            for rec in records:
                writer.writerow(rec)

        logger.info(f"Saved manifest to {manifest_path}")

    def download_company_10ks(
        self, ticker: str, count: int = 10
    ) -> Tuple[List[Dict[str, Any]], Path]:
        """Orchestrate downloading and organizing 10-K filings for a company.

        Args:
            ticker (str): Ticker symbol.
            count (int): Number of filings to retrieve (default: 10).

        Returns:
            Tuple[List[Dict[str, Any]], Path]: Processed records and manifest path.
        """
        clean_ticker = ticker.strip().upper()
        company_dir = config.RAW_FILINGS_DIR / clean_ticker
        company_dir.mkdir(parents=True, exist_ok=True)

        # 1. Resolve CIK
        cik, company_name = self.get_company_cik(clean_ticker)

        # 2. Query 10-K metadata
        filings_meta = self.get_10k_filings_metadata(
            cik=cik, ticker=clean_ticker, company_name=company_name, count=count
        )

        # 3. Download/Verify documents
        processed_records: List[Dict[str, Any]] = []
        for meta in filings_meta:
            rec = self.download_filing(meta, company_dir)
            processed_records.append(rec)

        # 4. Save CSV Manifest
        manifest_path = company_dir / f"{clean_ticker}_10K_manifest.csv"
        self.save_manifest(processed_records, manifest_path)

        return processed_records, manifest_path
