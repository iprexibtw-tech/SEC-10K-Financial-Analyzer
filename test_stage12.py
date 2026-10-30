"""Unit and integration verification tests for Desktop GUI and Standalone Executable.

Tests CustomTkinter desktop interface initialization, interactive company search,
metrics display, background pipeline integration, PyInstaller distribution integrity,
bundled data assets, and Windows desktop shortcut generation.
"""

import os
import subprocess
import sys
import unittest
from pathlib import Path

import config
import gui_app


class TestDesktopApplication(unittest.TestCase):
    """Test suite verifying CustomTkinter GUI and packaged Windows application."""

    def setUp(self) -> None:
        """Create and update GUI instance in memory without event-loop block."""
        self.app = gui_app.SEC10KAnalyzerApp()
        self.app.update()

    def tearDown(self) -> None:
        """Destroy GUI window cleanly."""
        try:
            self.app.destroy()
        except Exception:
            pass

    def test_gui_window_properties(self) -> None:
        """Verify CustomTkinter window title, dimensions, and default selection."""
        self.assertEqual(self.app.title(), "SEC 10-K Financial Analyzer")
        self.assertEqual(self.app.selected_ticker, "AAPL")
        self.assertIn("Apple Inc.", self.app.selected_company_title)
        self.assertEqual(self.app.selected_cik, "0000320193")

    def test_gui_ticker_entry_and_switching(self) -> None:
        """Verify changing ticker updates company details and loads existing workbook."""
        self.app.ticker_entry.delete(0, "end")
        self.app.ticker_entry.insert(0, "NVDA")
        self.app._on_ticker_entered(notify_log=False)
        self.app.update()

        self.assertEqual(self.app.selected_ticker, "NVDA")
        self.assertIn("NVIDIA", self.app.selected_company_title.upper())
        self.assertEqual(self.app.selected_cik, "0001045810")
        self.assertIsNotNone(self.app.last_completed_workbook)

    def test_gui_company_name_search(self) -> None:
        """Verify searching company name returns ranked results in GUI frame."""
        self.app.search_entry.delete(0, "end")
        self.app.search_entry.insert(0, "Microsoft")
        self.app._on_search_clicked()
        self.app.update()

        # Children in search results frame should be > 0 (header + results)
        children = self.app.search_results_frame.winfo_children()
        self.assertTrue(len(children) >= 2)

    def test_packaged_executable_structure(self) -> None:
        """Verify that PyInstaller build outputs the standalone executable and internal assets."""
        exe_path = Path("dist/SEC10KAnalyzer/SEC10KAnalyzer.exe")
        self.assertTrue(exe_path.exists(), f"Packaged executable not found at {exe_path}")
        self.assertGreater(exe_path.stat().st_size, 10_000_000, "Executable file size unexpectedly small")

        internal_dir = Path("dist/SEC10KAnalyzer/_internal")
        self.assertTrue(internal_dir.exists(), "Internal libraries folder missing")

        # Verify bundled dictionary
        dict_candidates = list((internal_dir / "data").glob("*MasterDictionary*.csv"))
        self.assertTrue(len(dict_candidates) > 0, "Bundled Loughran-McDonald dictionary missing")

        # Verify company tickers registry
        tickers_path = internal_dir / "data" / "financial_data" / "company_tickers.json"
        self.assertTrue(tickers_path.exists(), "Bundled company_tickers.json missing")

    def test_desktop_shortcut_creation(self) -> None:
        """Verify Windows desktop shortcut creation script."""
        desktop_dir = Path(os.environ.get("USERPROFILE", "")) / "OneDrive" / "Desktop"
        if not desktop_dir.exists():
            desktop_dir = Path(os.environ.get("USERPROFILE", "")) / "Desktop"

        shortcut_file = desktop_dir / "SEC 10-K Financial Analyzer.lnk"
        self.assertTrue(shortcut_file.exists(), f"Desktop shortcut not found at {shortcut_file}")

    def test_packaged_executable_execution(self) -> None:
        """Verify packaged executable runs outside project and executes Stage 1 checks."""
        exe_path = str(Path("dist/SEC10KAnalyzer/SEC10KAnalyzer.exe").resolve())
        proc = subprocess.run(
            [exe_path, "--verify-stage1"],
            cwd=os.environ.get("USERPROFILE", os.path.expanduser("~")),
            capture_output=True,
            text=True,
        )
        self.assertEqual(proc.returncode, 0, f"Packaged executable failed: {proc.stderr}")
        self.assertIn("[PASS] Directories and environment verified.", proc.stdout)
        self.assertIn("SEC10KAnalyzer", proc.stdout)

    def test_portable_onefile_executable(self) -> None:
        """Verify portable one-file executable exists and passes self-test outside project directory."""
        portable_path = Path("dist/portable/SEC10KAnalyzer.exe")
        if not portable_path.exists():
            portable_path = Path("dist/SEC10KAnalyzer_portable.exe")
        self.assertTrue(portable_path.exists(), f"Portable executable missing: {portable_path}")
        self.assertGreater(portable_path.stat().st_size, 40_000_000, "Portable binary unexpectedly small")

        # Run verification check from user home directory
        proc = subprocess.run(
            [str(portable_path.resolve()), "--verify-stage1"],
            cwd=os.environ.get("USERPROFILE", os.path.expanduser("~")),
            capture_output=True,
            text=True,
        )
        self.assertEqual(proc.returncode, 0, f"Portable execution failed: {proc.stderr}")
        self.assertIn("[PASS] Directories and environment verified.", proc.stdout)

    def test_distribution_folder_structure(self) -> None:
        """Verify final distribution folder and organized company folders."""
        dist_dir = Path("SEC 10-K Financial Analyzer")
        self.assertTrue(dist_dir.exists(), "Distribution folder 'SEC 10-K Financial Analyzer' missing")
        self.assertTrue((dist_dir / "SEC10KAnalyzer.exe").exists(), "SEC10KAnalyzer.exe missing in dist folder")
        self.assertTrue((dist_dir / "README.txt").exists(), "README.txt missing in dist folder")

        comp_dir = dist_dir / "Company Analyses"
        self.assertTrue(comp_dir.exists(), "Company Analyses folder missing in dist folder")

        for ticker, name in [("AAPL", "Apple Inc"), ("NVDA", "NVIDIA Corporation")]:
            c_folder = comp_dir / f"{ticker} - {name}"
            self.assertTrue(c_folder.exists(), f"Company folder missing: {c_folder}")
            self.assertTrue((c_folder / f"{ticker}_10K_Analysis.xlsx").exists(), f"Excel analysis missing for {ticker}")
            self.assertTrue((c_folder / f"{ticker}_financial_statements.xlsx").exists(), f"Financial statements missing for {ticker}")
            self.assertTrue((c_folder / "SEC Filings").exists(), f"SEC Filings folder missing for {ticker}")
            self.assertTrue((c_folder / "MD&A Text").exists(), f"MD&A Text folder missing for {ticker}")
            self.assertTrue((c_folder / f"{ticker}_financial_ratios.csv").exists(), f"Ratios CSV missing for {ticker}")
            self.assertTrue((c_folder / f"{ticker}_sentiment_analysis.csv").exists(), f"Sentiment CSV missing for {ticker}")
            self.assertTrue((c_folder / f"{ticker}_stock_prices.csv").exists(), f"Prices CSV missing for {ticker}")

    def test_distribution_zip_archive(self) -> None:
        """Verify compressed ZIP archive integrity and required files."""
        zip_path = Path("SEC_10K_Financial_Analyzer.zip")
        self.assertTrue(zip_path.exists(), "SEC_10K_Financial_Analyzer.zip missing")
        self.assertGreater(zip_path.stat().st_size, 40_000_000, "ZIP archive unexpectedly small")

        import zipfile
        with zipfile.ZipFile(zip_path, "r") as zf:
            namelist = zf.namelist()
            self.assertTrue(any("SEC10KAnalyzer.exe" in n for n in namelist), "Executable missing in ZIP")
            self.assertTrue(any("README.txt" in n for n in namelist), "README.txt missing in ZIP")
            self.assertTrue(any("AAPL_10K_Analysis.xlsx" in n for n in namelist), "AAPL workbook missing in ZIP")
            self.assertTrue(any("NVDA_10K_Analysis.xlsx" in n for n in namelist), "NVDA workbook missing in ZIP")

    def test_gui_history_loading(self) -> None:
        """Verify GUI loads analyzed companies directly from history without recalculation."""
        for ticker in ["AAPL", "NVDA"]:
            proc = subprocess.run(
                [sys.executable, "gui_app.py", "--test-gui-load", ticker],
                capture_output=True,
                text=True,
            )
            self.assertEqual(proc.returncode, 0, f"GUI history load failed for {ticker}: {proc.stderr}")
            self.assertIn(f"Loaded {ticker}. Workbook exists: True", proc.stdout)


if __name__ == "__main__":
    unittest.main()
