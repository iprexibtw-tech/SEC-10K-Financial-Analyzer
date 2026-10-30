"""Project configuration and path management module.

This module centralizes path definitions, project settings, SEC EDGAR
parameters, and directory initialization for the SEC 10-K financial analysis
project.
"""

import os
import platform
import re
import shutil
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Base Directories and Frozen State Detection (PyInstaller Support)
# ---------------------------------------------------------------------------
def _is_dir_writable(path: Path) -> bool:
    """Check if a directory is writable by attempting a temporary file write."""
    try:
        path.mkdir(parents=True, exist_ok=True)
        test_file = path / f".sec_write_test_{os.getpid()}.tmp"
        test_file.touch(exist_ok=True)
        test_file.unlink(missing_ok=True)
        return True
    except Exception:
        return False


def sanitize_filename(name: str) -> str:
    """Sanitize strings for safe Windows folder and file names."""
    # Remove illegal Windows characters: \ / : * ? " < > |
    clean = re.sub(r'[\\/:*?"<>|]', '', name)
    clean = re.sub(r'\s+', ' ', clean).strip('. ')
    return clean or "Company"


IS_FROZEN: bool = getattr(sys, "frozen", False)

if IS_FROZEN:
    # PyInstaller bundled extraction directory (read-only assets)
    BUNDLE_DIR: Path = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    EXE_DIR: Path = Path(sys.executable).resolve().parent
    LOCAL_APP_DATA: Path = Path(os.getenv("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "SEC10KAnalyzer"

    # Portable file storage: Prefer executable folder if writable; fallback to %LOCALAPPDATA%
    if _is_dir_writable(EXE_DIR):
        BASE_DIR: Path = EXE_DIR
        IS_PORTABLE_STORAGE: bool = True
        STORAGE_STATUS_DESC: str = f"Portable Directory ({EXE_DIR.name})"
    else:
        BASE_DIR: Path = LOCAL_APP_DATA
        IS_PORTABLE_STORAGE: bool = False
        STORAGE_STATUS_DESC: str = f"User AppData ({LOCAL_APP_DATA})"
else:
    # Running in development environment
    BUNDLE_DIR: Path = Path(__file__).resolve().parent
    EXE_DIR: Path = BUNDLE_DIR
    BASE_DIR: Path = BUNDLE_DIR
    IS_PORTABLE_STORAGE: bool = True
    STORAGE_STATUS_DESC: str = "Project Workspace (Local)"

DATA_DIR: Path = BASE_DIR / "data"
COMPANY_ANALYSES_DIR: Path = BASE_DIR / "Company Analyses"

# Organized project data subdirectories
RAW_FILINGS_DIR: Path = DATA_DIR / "raw_filings"
EXTRACTED_TEXT_DIR: Path = DATA_DIR / "extracted_text"
CLEANED_TEXT_DIR: Path = DATA_DIR / "cleaned_text"
FINANCIAL_DATA_DIR: Path = DATA_DIR / "financial_data"
EXCEL_OUTPUTS_DIR: Path = DATA_DIR / "excel_outputs"

# Local settings file to persist user configuration (like User-Agent) safely
SETTINGS_FILE: Path = BASE_DIR / ".sec_settings"


def get_company_folder(ticker: str, company_name: Optional[str] = None) -> Path:
    """Resolve or create the organized company analysis folder.

    Searches for existing matching company folders in COMPANY_ANALYSES_DIR
    (e.g., 'AAPL - Apple Inc.' or 'AAPL - Apple Inc'), or creates a new one.

    Args:
        ticker (str): Company ticker symbol.
        company_name (Optional[str]): Legal company name.

    Returns:
        Path: Path to the company-specific analysis folder.
    """
    clean_ticker = ticker.strip().upper()
    COMPANY_ANALYSES_DIR.mkdir(parents=True, exist_ok=True)

    # Check if an existing folder matching '<TICKER> - ' or '<TICKER>' already exists
    for folder in COMPANY_ANALYSES_DIR.iterdir():
        if folder.is_dir():
            name_upper = folder.name.upper()
            if name_upper == clean_ticker or name_upper.startswith(f"{clean_ticker} -"):
                return folder

    # Build canonical name: '<TICKER> - <Clean Company Name>'
    if company_name:
        clean_name = sanitize_filename(company_name)
        folder_name = f"{clean_ticker} - {clean_name}"
    else:
        folder_name = f"{clean_ticker}"

    folder_path = COMPANY_ANALYSES_DIR / folder_name
    folder_path.mkdir(parents=True, exist_ok=True)
    return folder_path


def get_analyzed_companies() -> List[Tuple[str, str, Path]]:
    """Discover all analyzed company folders within Company Analyses/.

    Returns:
        List[Tuple[str, str, Path]]: List of (ticker, display_name, folder_path).
    """
    results: List[Tuple[str, str, Path]] = []
    if not COMPANY_ANALYSES_DIR.exists():
        return results

    for folder in sorted(COMPANY_ANALYSES_DIR.iterdir(), key=lambda p: p.name.upper()):
        if folder.is_dir():
            # Match pattern: '<TICKER> - <NAME>' or '<TICKER>'
            match = re.match(r"^([A-Z0-9]+)\s*-\s*(.+)$", folder.name, re.IGNORECASE)
            if match:
                ticker = match.group(1).upper()
                name = match.group(2).strip()
                results.append((ticker, f"{ticker} - {name}", folder))
            else:
                ticker = folder.name.upper()
                results.append((ticker, ticker, folder))

    return results

# ---------------------------------------------------------------------------
# SEC EDGAR Endpoints and Settings
# ---------------------------------------------------------------------------
SEC_COMPANY_TICKERS_URL: str = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL_TEMPLATE: str = "https://data.sec.gov/submissions/CIK{cik}.json"
SEC_ARCHIVES_BASE_URL: str = "https://www.sec.gov/Archives/edgar/data"

# Rate limiting: SEC permits up to 10 requests/second; we use 0.2s delay (max 5 req/s)
SEC_REQUEST_DELAY: float = 0.2
SEC_MAX_RETRIES: int = 4
SEC_BACKOFF_FACTOR: float = 1.5
SEC_REQUEST_TIMEOUT: int = 15

# Default assignment parameters
DEFAULT_TICKER: str = "AAPL"
DEFAULT_FILINGS_COUNT: int = 10

# Local settings file to persist user configuration (like User-Agent) safely
SETTINGS_FILE: Path = BASE_DIR / ".sec_settings"


def is_valid_user_agent(user_agent: str) -> bool:
    """Validate that a User-Agent string meets SEC EDGAR requirements.

    The SEC requires a descriptive user agent with an email address:
    Format: 'Sample Company Name AdminContact@<sample company domain>.com'

    Args:
        user_agent (str): The User-Agent string to validate.

    Returns:
        bool: True if valid according to SEC guidelines, False otherwise.
    """
    if not user_agent or not isinstance(user_agent, str):
        return False
    user_agent = user_agent.strip()
    # Check for basic format: contains text and an email-like pattern
    email_pattern = r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+"
    has_email = re.search(email_pattern, user_agent) is not None
    has_name = len(user_agent.split()) >= 2
    is_not_default = "student@university.edu" not in user_agent and "sample" not in user_agent.lower()
    return bool(has_email and has_name and is_not_default)


def get_sec_user_agent() -> Optional[str]:
    """Retrieve the configured SEC User-Agent header.

    Checks:
    1. Environment variable `SEC_USER_AGENT`.
    2. Streamlit secrets `st.secrets["SEC_USER_AGENT"]` (for Streamlit Community Cloud).
    3. Local `.sec_settings` file if available.
    4. Bundled settings file if frozen.

    Returns:
        Optional[str]: Configured User-Agent string, or None if not set.
    """
    env_agent = os.getenv("SEC_USER_AGENT")
    if env_agent and is_valid_user_agent(env_agent):
        return env_agent.strip()

    # Check Streamlit secrets if running in Streamlit
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "SEC_USER_AGENT" in st.secrets:
            secret_ua = str(st.secrets["SEC_USER_AGENT"]).strip()
            if is_valid_user_agent(secret_ua):
                return secret_ua
    except Exception:
        pass

    if SETTINGS_FILE.exists():
        try:
            stored = SETTINGS_FILE.read_text(encoding="utf-8").strip()
            if is_valid_user_agent(stored):
                return stored
        except OSError:
            pass

    # If frozen, check bundled settings file as fallback
    if IS_FROZEN:
        bundled_settings = BUNDLE_DIR / ".sec_settings"
        if bundled_settings.exists():
            try:
                stored = bundled_settings.read_text(encoding="utf-8").strip()
                if is_valid_user_agent(stored):
                    try:
                        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
                        SETTINGS_FILE.write_text(stored, encoding="utf-8")
                    except OSError:
                        pass
                    return stored
            except OSError:
                pass

    return None


def save_sec_user_agent(user_agent: str) -> None:
    """Save the validated SEC User-Agent to the local settings file.

    Args:
        user_agent (str): Validated User-Agent string.

    Raises:
        ValueError: If the user agent does not meet SEC standards.
    """
    if not is_valid_user_agent(user_agent):
        raise ValueError(
            "User-Agent must include your name/organization and a valid contact email "
            "(e.g., 'Kevin Smith kevin@university.edu')."
        )
    SETTINGS_FILE.write_text(user_agent.strip(), encoding="utf-8")


def init_directories() -> Dict[str, Path]:
    """Ensure all required project data directories exist.

    Returns:
        Dict[str, Path]: Mapping of directory names to Path objects.

    Raises:
        OSError: If any directory cannot be created due to permission issues.
    """
    directories: Dict[str, Path] = {
        "data": DATA_DIR,
        "raw_filings": RAW_FILINGS_DIR,
        "extracted_text": EXTRACTED_TEXT_DIR,
        "cleaned_text": CLEANED_TEXT_DIR,
        "financial_data": FINANCIAL_DATA_DIR,
        "excel_outputs": EXCEL_OUTPUTS_DIR,
        "company_analyses": COMPANY_ANALYSES_DIR,
    }

    for name, path in directories.items():
        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError as err:
            raise OSError(
                f"Failed to create directory '{name}' at {path}: {err}"
            ) from err

    # When running as packaged application, seed user-data directory from bundled resources
    if IS_FROZEN and BUNDLE_DIR.exists() and BUNDLE_DIR != BASE_DIR:
        # Seed Company Analyses folder if bundled
        bundled_analyses = BUNDLE_DIR / "Company Analyses"
        if bundled_analyses.exists():
            for c_folder in bundled_analyses.iterdir():
                if c_folder.is_dir():
                    dest_c_folder = COMPANY_ANALYSES_DIR / c_folder.name
                    if not dest_c_folder.exists():
                        try:
                            shutil.copytree(c_folder, dest_c_folder)
                        except OSError:
                            pass

        bundled_data = BUNDLE_DIR / "data"
        if bundled_data.exists():
            # 1. Seed Loughran-McDonald dictionary
            for dict_file in bundled_data.glob("*MasterDictionary*.csv"):
                dest = DATA_DIR / dict_file.name
                if not dest.exists():
                    try:
                        shutil.copy2(dict_file, dest)
                    except OSError:
                        pass
            # 2. Seed financial reference data (e.g. company_tickers.json, cached facts)
            bundled_fin = bundled_data / "financial_data"
            if bundled_fin.exists():
                for fin_file in bundled_fin.glob("*.*"):
                    dest = FINANCIAL_DATA_DIR / fin_file.name
                    if not dest.exists():
                        try:
                            shutil.copy2(fin_file, dest)
                        except OSError:
                            pass
            # 3. Seed pre-generated Excel workbooks (AAPL & NVDA)
            bundled_excel = bundled_data / "excel_outputs"
            if bundled_excel.exists():
                for xl_file in bundled_excel.glob("*.xlsx"):
                    dest = EXCEL_OUTPUTS_DIR / xl_file.name
                    if not dest.exists():
                        try:
                            shutil.copy2(xl_file, dest)
                        except OSError:
                            pass
                for company_dir in bundled_excel.iterdir():
                    if company_dir.is_dir():
                        dest_company = EXCEL_OUTPUTS_DIR / company_dir.name
                        dest_company.mkdir(parents=True, exist_ok=True)
                        for item in company_dir.iterdir():
                            dest_item = dest_company / item.name
                            if not dest_item.exists() and item.is_file():
                                try:
                                    shutil.copy2(item, dest_item)
                                except OSError:
                                    pass

    return directories


def get_environment_info() -> Dict[str, str]:
    """Inspect and return current Python environment and system details.

    Returns:
        Dict[str, str]: Dictionary containing OS, Python version, executable,
            and virtual environment status.
    """
    in_venv = sys.prefix != sys.base_prefix
    return {
        "operating_system": f"{platform.system()} {platform.release()}",
        "python_version": platform.python_version(),
        "python_executable": sys.executable,
        "virtual_environment_active": "Yes" if in_venv else "No",
        "project_root": str(BASE_DIR),
        "sec_user_agent_configured": "Yes" if get_sec_user_agent() else "No",
    }
