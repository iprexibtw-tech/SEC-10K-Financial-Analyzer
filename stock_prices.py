"""Historical Stock Price Extraction and Excel Integration Module.

This module retrieves verified historical stock prices corresponding to the fiscal
year-end balance sheet dates of annual SEC 10-K filings for publicly traded companies
(such as Apple Inc. and NVIDIA Corporation).

Key Features:
- Automatically reads fiscal balance sheet period-end dates from verified filing metadata.
- Retrieves market data via yfinance without hardcoding prices or dates.
- Matches fiscal dates to actual trading days; if a balance sheet date falls on a weekend
  or market holiday, it selects the closing price from the most recent preceding trading day.
- Records both unadjusted/split-adjusted closing prices and dividend-adjusted closing prices.
- Clearly distinguishes between split-adjusted Close and total-return Adj Close.
- Generates a professionally formatted 'Stock Prices' worksheet in the company's Excel workbook
  while strictly preserving all financial statement and sentiment analysis sheets.
- Exports audit-ready CSV records to the project's financial_data directory.
"""

import argparse
import csv
import datetime
import logging
import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Suppress pandas/yfinance deprecation warnings for cleaner logging
warnings.filterwarnings("ignore", category=Warning)
import yfinance as yf

import config

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("Stock-Prices")


class StockPriceError(Exception):
    """Base exception for stock price extraction operations."""


class StockPriceRetrievalError(StockPriceError):
    """Raised when market data cannot be retrieved from the provider."""


class MissingDateError(StockPriceError):
    """Raised when balance sheet dates cannot be located or resolved."""


class StockPriceExtractor:
    """Retrieves and manages historical stock price data for SEC 10-K balance sheet dates."""

    DATA_SOURCE_NAME = "Yahoo Finance (via yfinance)"
    PRIMARY_MEASURE_NAME = "Split-Adjusted Closing Price"

    def __init__(
        self,
        excel_dir: Optional[Path] = None,
        financial_data_dir: Optional[Path] = None,
        raw_filings_dir: Optional[Path] = None,
    ) -> None:
        """Initialize the stock price extractor with project directories.

        Args:
            excel_dir (Optional[Path]): Directory containing Excel workbooks.
            financial_data_dir (Optional[Path]): Directory containing structured financial data.
            raw_filings_dir (Optional[Path]): Directory containing filing manifests.
        """
        self.excel_dir = excel_dir or config.EXCEL_OUTPUTS_DIR
        self.financial_data_dir = financial_data_dir or config.FINANCIAL_DATA_DIR
        self.raw_filings_dir = raw_filings_dir or config.RAW_FILINGS_DIR

    def get_balance_sheet_dates(self, ticker: str) -> List[Dict[str, str]]:
        """Resolve the 10 fiscal reporting period-end dates for the specified ticker.

        Reads from available metadata files in priority order:
        1. data/financial_data/{ticker}_sentiment_analysis.csv
        2. data/raw_filings/{ticker}/{ticker}_10K_manifest.csv
        3. data/excel_outputs/{ticker}/{ticker}_financial_statements.xlsx

        Args:
            ticker (str): Company stock ticker.

        Returns:
            List[Dict[str, str]]: Chronologically sorted list of dicts with 'fiscal_year'
                                  and 'balance_sheet_date'.

        Raises:
            MissingDateError: If balance sheet dates cannot be discovered.
        """
        periods: List[Dict[str, str]] = []

        # 1. Try sentiment analysis CSV
        sent_csv = self.financial_data_dir / f"{ticker}_sentiment_analysis.csv"
        if sent_csv.exists():
            try:
                with open(sent_csv, mode="r", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        fy = row.get("fiscal_reporting_year") or row.get("fiscal_year")
                        date_val = row.get("fiscal_period_end_date")
                        if fy and date_val:
                            periods.append({"fiscal_year": fy.strip(), "balance_sheet_date": date_val.strip()})
                if periods:
                    periods.sort(key=lambda p: int(p["fiscal_year"]))
                    return periods
            except Exception as e:
                logger.warning(f"Failed to read dates from {sent_csv}: {e}")

        # 2. Try raw filings manifest
        manifest_path = self.raw_filings_dir / ticker / f"{ticker}_10K_manifest.csv"
        if manifest_path.exists():
            try:
                with open(manifest_path, mode="r", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        fy = row.get("fiscal_year")
                        date_val = row.get("fiscal_period_end_date")
                        if fy and date_val:
                            periods.append({"fiscal_year": fy.strip(), "balance_sheet_date": date_val.strip()})
                if periods:
                    # Remove potential duplicates and sort
                    unique = {p["fiscal_year"]: p["balance_sheet_date"] for p in periods}
                    sorted_periods = [
                        {"fiscal_year": fy, "balance_sheet_date": dt}
                        for fy, dt in sorted(unique.items(), key=lambda x: int(x[0]))
                    ]
                    return sorted_periods
            except Exception as e:
                logger.warning(f"Failed to read dates from {manifest_path}: {e}")

        # 3. Try reading directly from Excel workbook header row
        excel_path = self.excel_dir / ticker / f"{ticker}_financial_statements.xlsx"
        if excel_path.exists():
            try:
                wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
                target_sheet = "Balance Sheet" if "Balance Sheet" in wb.sheetnames else wb.sheetnames[0]
                ws = wb[target_sheet]
                fy_row = [ws.cell(4, c).value for c in range(2, 14)]
                date_row = [ws.cell(5, c).value for c in range(2, 14)]
                for fy_cell, date_cell in zip(fy_row, date_row):
                    if fy_cell and date_cell:
                        fy_str = str(fy_cell).replace("FY", "").strip()
                        date_str = str(date_cell).strip()
                        periods.append({"fiscal_year": fy_str, "balance_sheet_date": date_str})
                if periods:
                    periods.sort(key=lambda p: int(p["fiscal_year"]))
                    return periods
            except Exception as e:
                logger.warning(f"Failed to read dates from Excel workbook: {e}")

        raise MissingDateError(
            f"Could not discover fiscal balance sheet dates for {ticker}. "
            f"Ensure Stage 2 or Stage 4 manifests exist in {self.raw_filings_dir / ticker}."
        )

    def retrieve_stock_prices(
        self,
        ticker: str,
        periods: Optional[List[Dict[str, str]]] = None,
    ) -> List[Dict[str, Any]]:
        """Retrieve verified historical stock prices for each fiscal year-end date.

        If a fiscal year-end date falls on a weekend or exchange holiday, retrieves
        the closing price from the most recent preceding trading day.

        Args:
            ticker (str): Public company ticker symbol.
            periods (Optional[List[Dict[str, str]]]): Explicit list of fiscal periods and dates.

        Returns:
            List[Dict[str, Any]]: Chronologically sorted list of stock price records.

        Raises:
            StockPriceRetrievalError: If historical price data cannot be retrieved.
        """
        clean_ticker = ticker.strip().upper()
        target_periods = periods or self.get_balance_sheet_dates(clean_ticker)

        if not target_periods:
            raise StockPriceRetrievalError(f"No fiscal periods available to retrieve stock prices for {clean_ticker}.")

        # Determine date query boundary with safety margins
        parsed_dates = [datetime.datetime.strptime(p["balance_sheet_date"], "%Y-%m-%d").date() for p in target_periods]
        earliest_date = min(parsed_dates) - datetime.timedelta(days=15)
        latest_date = max(parsed_dates) + datetime.timedelta(days=7)

        logger.info(
            f"Querying {self.DATA_SOURCE_NAME} for {clean_ticker} from "
            f"{earliest_date.strftime('%Y-%m-%d')} to {latest_date.strftime('%Y-%m-%d')}..."
        )

        try:
            yf_ticker = yf.Ticker(clean_ticker)
            # auto_adjust=False provides both split-adjusted Close and dividend-adjusted Adj Close
            hist_df = yf_ticker.history(
                start=earliest_date.strftime("%Y-%m-%d"),
                end=latest_date.strftime("%Y-%m-%d"),
                auto_adjust=False,
            )
        except Exception as err:
            raise StockPriceRetrievalError(
                f"Failed to retrieve market data for {clean_ticker} from {self.DATA_SOURCE_NAME}: {err}"
            )

        if hist_df is None or hist_df.empty:
            raise StockPriceRetrievalError(
                f"No historical price data returned for {clean_ticker} between {earliest_date} and {latest_date}."
            )

        results: List[Dict[str, Any]] = []

        for p in target_periods:
            fy = p["fiscal_year"]
            bs_date_str = p["balance_sheet_date"]
            bs_date = datetime.datetime.strptime(bs_date_str, "%Y-%m-%d").date()

            # Filter data up to and including the balance sheet date
            # This automatically selects the most recent preceding trading day if bs_date is a weekend/holiday
            sub_df = hist_df.loc[:bs_date_str]
            if sub_df.empty:
                raise StockPriceRetrievalError(
                    f"No trading data available on or before {bs_date_str} for FY{fy} ({clean_ticker})."
                )

            last_row = sub_df.iloc[-1]
            trade_timestamp = sub_df.index[-1]
            trade_date = trade_timestamp.date() if hasattr(trade_timestamp, "date") else trade_timestamp
            trade_date_str = trade_date.strftime("%Y-%m-%d")

            # Verification check: actual trading date must NOT be later than balance sheet date
            if trade_date > bs_date:
                raise StockPriceError(
                    f"Integrity error: Trading date {trade_date_str} is later than balance sheet date {bs_date_str}."
                )

            offset_days = (bs_date - trade_date).days
            close_price = round(float(last_row["Close"]), 2)
            adj_close_price = round(float(last_row.get("Adj Close", last_row["Close"])), 2)
            volume = int(last_row.get("Volume", 0))

            # Validate prices are positive numbers
            if close_price <= 0 or adj_close_price <= 0:
                raise StockPriceError(
                    f"Invalid non-positive price detected for FY{fy} ({clean_ticker}): "
                    f"Close=${close_price}, Adj Close=${adj_close_price}."
                )

            # Document non-trading day explanation
            if offset_days == 0:
                trading_day_note = "Exact balance sheet date was an active market trading day."
            else:
                trading_day_note = (
                    f"Balance sheet date was a non-trading day (weekend/holiday). "
                    f"Price matched to preceding trading day ({offset_days} day(s) earlier)."
                )

            record = {
                "company_ticker": clean_ticker,
                "ticker": clean_ticker,
                "fiscal_year": fy,
                "fiscal_balance_sheet_date": bs_date_str,
                "actual_trading_date": trade_date_str,
                "offset_days": offset_days,
                "closing_price_usd": close_price,
                "adjusted_closing_price_usd": adj_close_price,
                "trading_volume": volume,
                "primary_measure": self.PRIMARY_MEASURE_NAME,
                "data_source": self.DATA_SOURCE_NAME,
                "retrieval_status": "Verified",
                "notes": trading_day_note,
            }
            results.append(record)

        # Sort chronologically (oldest to newest: FY2016 -> FY2025)
        results.sort(key=lambda r: int(r["fiscal_year"]))
        logger.info(f"Successfully matched stock prices for {len(results)} fiscal periods for {clean_ticker}.")
        return results

    def add_stock_prices_worksheet_to_excel(
        self,
        ticker: str,
        price_records: List[Dict[str, Any]],
        company_name: Optional[str] = None,
        cik: Optional[str] = None,
    ) -> Path:
        """Add worksheet named 'Stock Prices' to the existing Excel workbook.

        Preserves all existing worksheets (Income Statement, Balance Sheet, Cash Flow
        Statement, Sentiment Analysis, and Audit & Reconciliation) and presents the 10
        fiscal periods in chronological order with professional formatting.

        Args:
            ticker (str): Company ticker symbol.
            price_records (List[Dict[str, Any]]): Chronologically sorted price records.
            company_name (Optional[str]): Company display name.
            cik (Optional[str]): Company SEC CIK number.

        Returns:
            Path: Path to the updated Excel workbook.

        Raises:
            StockPriceError: If the workbook cannot be loaded or saved.
        """
        company_dir = self.excel_dir / ticker
        excel_path = company_dir / f"{ticker}_financial_statements.xlsx"

        if not excel_path.exists():
            logger.warning(f"Workbook not found at {excel_path}. Creating new workbook.")
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
        else:
            try:
                wb = openpyxl.load_workbook(excel_path)
                logger.info(f"Loaded existing Excel workbook: {excel_path.name} (Sheets: {wb.sheetnames})")
            except Exception as e:
                raise StockPriceError(f"Failed to load Excel workbook at {excel_path}: {e}")

        sheet_title = "Stock Prices"

        # If worksheet already exists, remove it to refresh cleanly
        if sheet_title in wb.sheetnames:
            del wb[sheet_title]

        # Determine target insertion index: place right after Sentiment Analysis
        target_idx = len(wb.sheetnames)
        if "Sentiment Analysis" in wb.sheetnames:
            target_idx = wb.sheetnames.index("Sentiment Analysis") + 1
        elif "Audit & Reconciliation" in wb.sheetnames:
            target_idx = wb.sheetnames.index("Audit & Reconciliation")

        ws = wb.create_sheet(title=sheet_title, index=target_idx)
        ws.views.sheetView[0].showGridLines = True

        disp_name = company_name or ("Apple Inc." if ticker == "AAPL" else ticker)
        disp_cik = cik or ("0000320193" if ticker == "AAPL" else "")

        # Styling definitions matching Stage 4 and Stage 6 design system
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
        ws["A2"] = f"HISTORICAL STOCK PRICES FOR 10-K BALANCE SHEET DATES (CIK: {disp_cik})"
        ws["A2"].font = Font(name=font_name, size=12, bold=True, color="333333")
        ws.row_dimensions[2].height = 18

        # Row 3: Methodology Subtitle
        ws["A3"] = f"Data Source: {self.DATA_SOURCE_NAME} | Split-Adjusted Market Closing Prices"
        ws["A3"].font = subtitle_font
        ws.row_dimensions[3].height = 16

        # Row 4: Column Headers (Row 4 = FY, Row 5 = Balance Sheet Date, Row 6 = Actual Trading Date)
        ws.row_dimensions[4].height = 22
        ws.row_dimensions[5].height = 18
        ws.row_dimensions[6].height = 18

        ws["A4"] = "Metric / Stock Price Classification"
        ws["A5"] = "Fiscal Balance Sheet Date"
        ws["A6"] = "Actual Trading Date Used"

        for row_num, label in [(4, "Metric / Stock Price Classification"), (5, "Fiscal Balance Sheet Date"), (6, "Actual Trading Date Used")]:
            cell = ws[f"A{row_num}"]
            cell.fill = navy_header_fill
            cell.font = col_hdr_font if row_num == 4 else Font(name=font_name, size=9, bold=True, color="FFFFFF")
            cell.alignment = align_left

        col_indices: Dict[str, int] = {}
        for idx, rec in enumerate(price_records, start=2):
            col_letter = get_column_letter(idx)
            col_indices[rec["fiscal_year"]] = idx

            cell_fy = ws[f"{col_letter}4"]
            cell_bs = ws[f"{col_letter}5"]
            cell_tr = ws[f"{col_letter}6"]

            cell_fy.value = f"FY {rec['fiscal_year']}"
            cell_fy.fill = navy_header_fill
            cell_fy.font = col_hdr_font
            cell_fy.alignment = align_center

            cell_bs.value = rec["fiscal_balance_sheet_date"]
            cell_bs.fill = navy_header_fill
            cell_bs.font = Font(name=font_name, size=9, bold=True, color="FFFFFF")
            cell_bs.alignment = align_center

            cell_tr.value = rec["actual_trading_date"]
            cell_tr.fill = navy_header_fill
            cell_tr.font = Font(name=font_name, size=9, bold=True, color="FFFFFF")
            cell_tr.alignment = align_center

        # Row definitions
        rows_definition = [
            # Section 1
            ("1. HISTORICAL STOCK PRICE MEASURES ($ USD)", None, None, None, section_hdr_fill, section_hdr_font, True),
            ("Split-Adjusted Closing Price (Primary Analysis Measure)", "closing_price_usd", "$#,##0.00", total_border, highlight_fill, key_metric_font, False),
            ("Adjusted Closing Price (Splits & Cumulative Dividends)", "adjusted_closing_price_usd", "$#,##0.00", regular_border, None, bold_font, False),

            # Section 2
            ("2. MARKET TRADING ACTIVITY & VERIFICATION METRICS", None, None, None, section_hdr_fill, section_hdr_font, True),
            ("Trading Volume (Shares Traded on Date)", "trading_volume", "#,##0", regular_border, None, regular_font, False),
            ("Preceding Trading Day Offset (Days Prior to BS Date)", "offset_days", "0", regular_border, None, regular_font, False),
            ("Historical Market Data Source", "data_source", None, regular_border, None, regular_font, False),
            ("Price Observation Verification Status", "retrieval_status", None, subtotal_border, None, bold_font, False),

            # Section 3
            ("3. METHODOLOGY & STOCK SPLIT AUDIT NOTES", None, None, None, section_hdr_fill, section_hdr_font, True),
            ("Trading Day Matching Rule", "notes", None, regular_border, None, Font(name=font_name, size=9, italic=True), False),
        ]

        current_row = 7
        num_cols = len(price_records) + 1

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

            for rec in price_records:
                c_idx = col_indices[rec["fiscal_year"]]
                col_let = get_column_letter(c_idx)
                cell_data = ws[f"{col_let}{current_row}"]
                cell_data.font = font_style or regular_font

                val = rec.get(key)
                cell_data.value = val

                if isinstance(val, (int, float)):
                    cell_data.alignment = align_right
                    if num_fmt:
                        cell_data.number_format = num_fmt
                else:
                    cell_data.alignment = align_center

                if fill_style:
                    cell_data.fill = fill_style
                if border_style:
                    cell_data.border = border_style

            current_row += 1

        # Column widths
        ws.column_dimensions["A"].width = 52
        for c_idx in range(2, num_cols + 1):
            ws.column_dimensions[get_column_letter(c_idx)].width = 16

        # Freeze panes at cell B7 (keeps classifications in Col A and header dates in Rows 4-6 visible)
        ws.freeze_panes = "B7"

        try:
            wb.save(excel_path)
            logger.info(f"Successfully integrated worksheet '{sheet_title}' into {excel_path}")
        except Exception as e:
            raise StockPriceError(f"Failed to save updated Excel workbook at {excel_path}: {e}")

        return excel_path

    def export_summary_csv(self, ticker: str, price_records: List[Dict[str, Any]]) -> Tuple[Path, Path]:
        """Export structured stock price records to CSV files.

        Saves to both data/financial_data/ and data/excel_outputs/{ticker}/.

        Args:
            ticker (str): Public company ticker symbol.
            price_records (List[Dict[str, Any]]): List of stock price records.

        Returns:
            Tuple[Path, Path]: (primary_csv_path, mirror_csv_path)
        """
        self.financial_data_dir.mkdir(parents=True, exist_ok=True)
        primary_csv = self.financial_data_dir / f"{ticker}_stock_prices.csv"

        company_excel_dir = self.excel_dir / ticker
        company_excel_dir.mkdir(parents=True, exist_ok=True)
        mirror_csv = company_excel_dir / f"{ticker}_stock_prices_summary.csv"

        fieldnames = [
            "company_ticker",
            "fiscal_year",
            "fiscal_balance_sheet_date",
            "actual_trading_date",
            "offset_days",
            "closing_price_usd",
            "adjusted_closing_price_usd",
            "trading_volume",
            "primary_measure",
            "data_source",
            "retrieval_status",
            "notes",
        ]

        for target_csv in [primary_csv, mirror_csv]:
            with open(target_csv, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for r in price_records:
                    writer.writerow(
                        {
                            "company_ticker": r["company_ticker"],
                            "fiscal_year": r["fiscal_year"],
                            "fiscal_balance_sheet_date": r["fiscal_balance_sheet_date"],
                            "actual_trading_date": r["actual_trading_date"],
                            "offset_days": r["offset_days"],
                            "closing_price_usd": f"{r['closing_price_usd']:.2f}",
                            "adjusted_closing_price_usd": f"{r['adjusted_closing_price_usd']:.2f}",
                            "trading_volume": r["trading_volume"],
                            "primary_measure": r["primary_measure"],
                            "data_source": r["data_source"],
                            "retrieval_status": r["retrieval_status"],
                            "notes": r["notes"],
                        }
                    )

        logger.info(f"Saved stock prices CSV output to {primary_csv}")
        return primary_csv, mirror_csv


def run_stock_price_extraction(ticker: str = "AAPL") -> Tuple[List[Dict[str, Any]], Path, Path]:
    """Execute complete Stage 7 stock price retrieval, Excel integration, and CSV export.

    Args:
        ticker (str): Public company ticker symbol.

    Returns:
        Tuple[List[Dict[str, Any]], Path, Path]: (records, excel_path, primary_csv_path)
    """
    clean_ticker = ticker.strip().upper()
    extractor = StockPriceExtractor()
    records = extractor.retrieve_stock_prices(ticker=clean_ticker)
    excel_path = extractor.add_stock_prices_worksheet_to_excel(ticker=clean_ticker, price_records=records)
    primary_csv, _ = extractor.export_summary_csv(ticker=clean_ticker, price_records=records)
    return records, excel_path, primary_csv


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Historical Stock Price Extraction for SEC 10-K Balance Sheet Dates"
    )
    parser.add_argument(
        "--ticker",
        type=str,
        default="AAPL",
        help="Stock ticker symbol to retrieve (default: AAPL)",
    )
    args = parser.parse_args()

    try:
        recs, xl_path, csv_p = run_stock_price_extraction(ticker=args.ticker)
        print(f"\n[SUCCESS] Retrieved stock prices for {args.ticker.upper()} across {len(recs)} periods.")
        print(f"Updated Excel Workbook (Tab: 'Stock Prices'): {xl_path}")
        print(f"Stock Prices CSV Output: {csv_p}")
    except Exception as exc:
        logger.exception(f"Stock Price Extraction failed: {exc}")
        print(f"\n[ERROR] Stock price extraction failed: {exc}")
        sys.exit(1)
