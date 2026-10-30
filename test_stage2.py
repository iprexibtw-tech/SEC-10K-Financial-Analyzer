"""Unit and integration verification tests for Stage 2 SEC Downloader.

Tests input validation, User-Agent checks, 10-K filtering (excluding 10-K/A),
distinct fiscal period selection, filename generation, manifest generation,
content verification, and duplicate caching.
"""

import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import config
from sec_downloader import (
    SECAuthenticationError,
    SECCompanyNotFoundError,
    SECDownloader,
    SECDownloaderError,
    SECFilingsNotFoundError,
)


class TestStage2SECLogic(unittest.TestCase):
    """Test suite verifying SEC Downloader logic and business rules."""

    def setUp(self) -> None:
        """Create a temporary directory for isolated file operations."""
        self.test_dir = Path(tempfile.mkdtemp(prefix="test_sec_"))
        self.valid_ua = "Kevin Student kevin.student@university.edu"

    def tearDown(self) -> None:
        """Clean up temporary test directory."""
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_user_agent_validation(self) -> None:
        """Verify that user-agent validation enforces name and email."""
        # Valid user agents
        self.assertTrue(config.is_valid_user_agent("Kevin Smith kevin@uni.edu"))
        self.assertTrue(config.is_valid_user_agent("AppleLab Research j.doe@mit.edu"))

        # Invalid user agents
        self.assertFalse(config.is_valid_user_agent(""))
        self.assertFalse(config.is_valid_user_agent("student@university.edu"))
        self.assertFalse(config.is_valid_user_agent("JustANameWithoutEmail"))
        self.assertFalse(config.is_valid_user_agent("sample@domain.com"))

    def test_missing_user_agent_raises_error(self) -> None:
        """Verify SECDownloader raises SECAuthenticationError when no UA is set."""
        with patch.object(config, "get_sec_user_agent", return_value=None):
            with self.assertRaises(SECAuthenticationError):
                SECDownloader(user_agent=None)

    def test_cik_lookup_parsing(self) -> None:
        """Verify that ticker lookup resolves 10-digit CIK correctly."""
        downloader = SECDownloader(user_agent=self.valid_ua)
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."},
            "1": {"cik_str": 789019, "ticker": "MSFT", "title": "MICROSOFT CORP"},
        }

        with patch.object(config, "FINANCIAL_DATA_DIR", self.test_dir):
            with patch.object(downloader, "_get", return_value=mock_response):
                cik, title = downloader.get_company_cik("aapl")
                self.assertEqual(cik, "0000320193")
                self.assertEqual(title, "Apple Inc.")

                with self.assertRaises(SECCompanyNotFoundError):
                    downloader.get_company_cik("NONEXISTENT")

    def test_extract_10ks_excludes_amendments_and_other_forms(self) -> None:
        """Verify strict exclusion of 10-K/A, 10-Q, 8-K filings."""
        downloader = SECDownloader(user_agent=self.valid_ua)

        sample_submissions = {
            "form": ["10-K", "10-K/A", "10-Q", "8-K", "10-K"],
            "filingDate": [
                "2023-11-03",
                "2023-12-01",
                "2023-08-04",
                "2023-05-10",
                "2022-10-28",
            ],
            "reportDate": [
                "2023-09-30",
                "2023-09-30",
                "2023-07-01",
                "2023-05-01",
                "2022-09-24",
            ],
            "accessionNumber": [
                "0000320193-23-000106",
                "0000320193-23-000115",
                "0000320193-23-000077",
                "0000320193-23-000055",
                "0000320193-22-000108",
            ],
            "primaryDocument": [
                "aapl-20230930.htm",
                "aapl-20230930a.htm",
                "aapl-20230701.htm",
                "aapl-8k.htm",
                "aapl-20220924.htm",
            ],
        }

        records = downloader._extract_10ks_from_submission_data(
            sample_submissions,
            cik="0000320193",
            ticker="AAPL",
            company_name="Apple Inc.",
        )

        # Only the two pure 10-Ks must be returned
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["fiscal_year"], "2023")
        self.assertEqual(records[0]["filing_date"], "2023-11-03")
        self.assertEqual(records[0]["form"], "10-K")
        self.assertEqual(records[0]["saved_filename"], "AAPL_10K_FY2023_2023-11-03.htm")

        self.assertEqual(records[1]["fiscal_year"], "2022")
        self.assertEqual(records[1]["filing_date"], "2022-10-28")
        self.assertEqual(records[1]["saved_filename"], "AAPL_10K_FY2022_2022-10-28.htm")

    def test_distinct_reporting_periods(self) -> None:
        """Verify that multiple 10-Ks for the same period are deduplicated."""
        downloader = SECDownloader(user_agent=self.valid_ua)

        # Mock submissions with two 10-Ks sharing the same report date
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "filings": {
                "recent": {
                    "form": ["10-K", "10-K"],
                    "filingDate": ["2023-11-05", "2023-11-03"],
                    "reportDate": ["2023-09-30", "2023-09-30"],
                    "accessionNumber": [
                        "0000320193-23-000108",
                        "0000320193-23-000106",
                    ],
                    "primaryDocument": [
                        "aapl-20230930-v2.htm",
                        "aapl-20230930.htm",
                    ],
                }
            }
        }

        with patch.object(downloader, "_get", return_value=mock_response):
            results = downloader.get_10k_filings_metadata(
                cik="0000320193",
                ticker="AAPL",
                company_name="Apple Inc.",
                count=10,
            )
            # Only 1 unique reporting period (2023-09-30) should be included
            self.assertEqual(len(results), 1)
            self.assertEqual(results[0]["filing_date"], "2023-11-05")

    def test_content_verification_detects_blocks_and_valid_files(self) -> None:
        """Verify content verification detects SEC blocks and validates HTML."""
        downloader = SECDownloader(user_agent=self.valid_ua)

        # 1. Nonexistent file
        is_val, msg = downloader.verify_filing_content(self.test_dir / "nonexistent.htm")
        self.assertFalse(is_val)

        # 2. Too small file
        tiny = self.test_dir / "tiny.htm"
        tiny.write_text("short", encoding="utf-8")
        is_val, msg = downloader.verify_filing_content(tiny)
        self.assertFalse(is_val)
        self.assertIn("too small", msg.lower())

        # 3. SEC Block message file
        blocked = self.test_dir / "blocked.htm"
        blocked.write_text(
            "<html><body>Your Request Has Been Blocked by SEC EDGAR</body></html>"
            + (" " * 3000),
            encoding="utf-8",
        )
        is_val, msg = downloader.verify_filing_content(blocked)
        self.assertFalse(is_val)
        self.assertIn("block", msg.lower())

        # 4. Valid substantive 10-K file
        valid_filing = self.test_dir / "valid_10k.htm"
        valid_filing.write_text(
            "<html><head><title>Form 10-K Annual Report</title></head>"
            "<body>UNITED STATES SECURITIES AND EXCHANGE COMMISSION"
            "<h1>Apple Inc. Form 10-K</h1>" + ("<p>Content</p>" * 200) + "</body></html>",
            encoding="utf-8",
        )
        is_val, msg = downloader.verify_filing_content(valid_filing)
        self.assertTrue(is_val)
        self.assertIn("valid", msg.lower())

    def test_caching_avoids_duplicate_download(self) -> None:
        """Verify that an existing valid file on disk is not re-downloaded."""
        downloader = SECDownloader(user_agent=self.valid_ua)

        target_file = self.test_dir / "AAPL_10K_FY2023_2023-11-03.htm"
        target_file.write_text(
            "<html><body><h1>Form 10-K Annual Report</h1>"
            + ("<p>Item 1</p>" * 300)
            + "</body></html>",
            encoding="utf-8",
        )

        meta = {
            "ticker": "AAPL",
            "cik": "0000320193",
            "saved_filename": target_file.name,
            "sec_source_url": "https://www.sec.gov/Archives/fake.htm",
        }

        # _get should NOT be called since the valid file already exists
        with patch.object(downloader, "_get") as mock_get:
            record = downloader.download_filing(meta, self.test_dir)
            mock_get.assert_not_called()
            self.assertEqual(record["verification_status"], "Verified (Cached)")

    def test_manifest_creation(self) -> None:
        """Verify that CSV manifest records all required fields."""
        downloader = SECDownloader(user_agent=self.valid_ua)
        manifest_path = self.test_dir / "manifest.csv"

        records = [
            {
                "ticker": "AAPL",
                "cik": "0000320193",
                "fiscal_year": "2023",
                "fiscal_period_end_date": "2023-09-30",
                "filing_date": "2023-11-03",
                "accession_number": "0000320193-23-000106",
                "primary_document": "aapl-20230930.htm",
                "sec_source_url": "https://www.sec.gov/Archives/edgar/data/320193/000032019323000106/aapl-20230930.htm",
                "saved_filename": "AAPL_10K_FY2023_2023-11-03.htm",
                "saved_file_path": str(self.test_dir / "AAPL_10K_FY2023_2023-11-03.htm"),
                "file_size_bytes": 1054320,
                "verification_status": "Verified (Downloaded)",
            }
        ]

        downloader.save_manifest(records, manifest_path)
        self.assertTrue(manifest_path.exists())

        content = manifest_path.read_text(encoding="utf-8")
        self.assertIn("AAPL", content)
        self.assertIn("0000320193", content)
        self.assertIn("2023-09-30", content)
        self.assertIn("0000320193-23-000106", content)
        self.assertIn("Verified (Downloaded)", content)

    def test_search_companies_exact_ticker(self) -> None:
        """Verify searching by exact ticker symbol returns the target company first."""
        downloader = SECDownloader(user_agent=self.valid_ua)
        results = downloader.search_companies("AAPL", limit=5)
        self.assertTrue(len(results) > 0)
        self.assertEqual(results[0]["ticker"], "AAPL")
        self.assertEqual(results[0]["cik"], "0000320193")
        self.assertIn("Apple", results[0]["title"])

    def test_search_companies_by_company_name(self) -> None:
        """Verify searching by company name 'Apple' prioritizes Apple Inc. (AAPL)."""
        downloader = SECDownloader(user_agent=self.valid_ua)
        results = downloader.search_companies("Apple", limit=5)
        self.assertTrue(len(results) > 0)
        self.assertEqual(results[0]["ticker"], "AAPL")
        self.assertEqual(results[0]["cik"], "0000320193")
        self.assertEqual(results[0]["title"], "Apple Inc.")

    def test_search_companies_nvidia(self) -> None:
        """Verify searching by company name 'NVIDIA' returns NVIDIA CORP (NVDA)."""
        downloader = SECDownloader(user_agent=self.valid_ua)
        results = downloader.search_companies("NVIDIA", limit=5)
        self.assertTrue(len(results) > 0)
        self.assertEqual(results[0]["ticker"], "NVDA")
        self.assertEqual(results[0]["cik"], "0001045810")
        self.assertIn("NVIDIA", results[0]["title"].upper())

    def test_search_companies_brand_alias(self) -> None:
        """Verify brand alias mapping (e.g. Google -> GOOGL)."""
        downloader = SECDownloader(user_agent=self.valid_ua)
        results = downloader.search_companies("Google", limit=5)
        self.assertTrue(len(results) > 0)
        self.assertEqual(results[0]["ticker"], "GOOGL")

    def test_search_companies_empty_and_unknown(self) -> None:
        """Verify empty and nonexistent searches handle gracefully."""
        downloader = SECDownloader(user_agent=self.valid_ua)
        self.assertEqual(downloader.search_companies(""), [])
        self.assertEqual(downloader.search_companies("   "), [])
        self.assertEqual(downloader.search_companies("XYZ_NONEXISTENT_COMPANY_99999"), [])

    def test_search_ranking_scoring_tiers(self) -> None:
        """Verify relevance ranking scoring tiers using controlled mock data."""
        downloader = SECDownloader(user_agent=self.valid_ua)
        mock_registry = {
            "0": {"cik_str": 100, "ticker": "SUB", "title": "A Company With Test Substring"},
            "1": {"cik_str": 200, "ticker": "TEST", "title": "Other Corporation"},
            "2": {"cik_str": 300, "ticker": "PRF", "title": "Test Prefix Incorporated"},
            "3": {"cik_str": 400, "ticker": "EXC", "title": "Test"},
        }
        with patch.object(downloader, "get_tickers_registry", return_value=mock_registry):
            results = downloader.search_companies("Test", limit=10)
            # Order should be:
            # 1. Exact ticker TEST (score 1000)
            # 2. Exact title Test (score 800)
            # 3. Title prefix Test Prefix Incorporated (score 500)
            # 4. Substring A Company With Test Substring (score 300/100)
            self.assertEqual(results[0]["ticker"], "TEST")
            self.assertEqual(results[1]["ticker"], "EXC")
            self.assertEqual(results[2]["ticker"], "PRF")
            self.assertEqual(results[3]["ticker"], "SUB")


if __name__ == "__main__":
    unittest.main()
