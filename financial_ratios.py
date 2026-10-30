"""Ten-Year Financial Ratio Analysis and Excel Integration Module.

This module calculates, verifies, and analyzes four key financial ratios across
10 annual SEC 10-K fiscal reporting periods for Apple (AAPL) and NVIDIA (NVDA):
1. Liquidity — Current Ratio (Current Assets / Current Liabilities)
2. Solvency — Debt-to-Equity Ratio (Total Interest-Bearing Debt / Shareholders' Equity)
3. Efficiency — Total Asset Turnover (Net Revenue / Average Total Assets)
4. Profitability — Net Profit Margin ((Net Income / Net Revenue) * 100)

Generates a dedicated 'Financial Ratios' worksheet in the company Excel workbook
with live formulas, supporting calculation details, professional 10-year trend
charts, and exports structured CSV datasets.
"""

import csv
import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import config

logger = logging.getLogger("Financial-Ratios")


class FinancialRatioError(Exception):
    """Base exception for financial ratio operations."""


class FinancialRatioAnalyzer:
    """Calculates, validates, and charts 10-year financial ratios."""

    def __init__(
        self,
        excel_dir: Optional[Path] = None,
        financial_dir: Optional[Path] = None,
    ) -> None:
        """Initialize the FinancialRatioAnalyzer.

        Args:
            excel_dir (Optional[Path]): Directory where Excel workbooks are stored.
            financial_dir (Optional[Path]): Directory where financial datasets and facts reside.
        """
        self.excel_dir = excel_dir or config.EXCEL_OUTPUTS_DIR
        self.financial_dir = financial_dir or config.FINANCIAL_DATA_DIR

    def get_prior_year_assets_from_facts(
        self, cik: str, earliest_fy: int
    ) -> Optional[float]:
        """Retrieve total assets for the year preceding the earliest analyzed fiscal year.

        Queries the local SEC Company Facts cache to obtain opening asset balances.

        Args:
            cik (str): 10-digit CIK string.
            earliest_fy (int): Earliest fiscal year analyzed (e.g., 2016).

        Returns:
            Optional[float]: Total assets in USD Millions, or None if unavailable.
        """
        cik_10 = cik.zfill(10)
        facts_path = self.financial_dir / f"companyfacts_{cik_10}.json"
        if not facts_path.exists():
            return None

        try:
            with open(facts_path, "r", encoding="utf-8") as f:
                facts_data = json.load(f)

            assets_units = (
                facts_data.get("facts", {})
                .get("us-gaap", {})
                .get("Assets", {})
                .get("units", {})
                .get("USD", [])
            )
            if not assets_units:
                return None

            prior_fy = earliest_fy - 1
            # Look for 10-K entries matching prior fiscal year
            matches = [
                u
                for u in assets_units
                if u.get("form") in ("10-K", "10-K/A")
                and (u.get("fy") == prior_fy or str(prior_fy) in u.get("end", ""))
            ]
            # Prioritize latest filed date (audited comparative facts in later 10-Ks)
            matches.sort(
                key=lambda x: (x.get("form") == "10-K", x.get("filed", "")),
                reverse=True,
            )

            if matches:
                val = matches[0].get("val")
                if val is not None:
                    return round(float(val) / 1e6, 1)
        except Exception as err:
            logger.warning(f"Error querying prior year assets from facts for CIK {cik}: {err}")

        return None

    def extract_financial_inputs_from_workbook(
        self, ticker: str, workbook_path: Optional[Path] = None
    ) -> Dict[str, Any]:
        """Extract required numerator and denominator inputs from existing Excel workbook.

        Reads values from 'Income Statement' and 'Balance Sheet' worksheets.

        Args:
            ticker (str): Company ticker symbol.
            workbook_path (Optional[Path]): Path to company Excel workbook.

        Returns:
            Dict[str, Any]: Dictionary containing structured accounting inputs across 10 years.

        Raises:
            FinancialRatioError: If required worksheets or accounts are missing.
        """
        clean_ticker = ticker.strip().upper()
        target_wb_path = (
            workbook_path
            or (self.excel_dir / clean_ticker / f"{clean_ticker}_financial_statements.xlsx")
        )

        if not target_wb_path.exists():
            raise FinancialRatioError(
                f"Financial statement workbook not found: {target_wb_path}. "
                f"Please run Stage 4 financial extraction first."
            )

        wb = openpyxl.load_workbook(target_wb_path, data_only=True)
        if "Income Statement" not in wb.sheetnames or "Balance Sheet" not in wb.sheetnames:
            raise FinancialRatioError(
                f"Workbook {target_wb_path.name} is missing required statement sheets "
                f"('Income Statement' and 'Balance Sheet')."
            )

        ws_is = wb["Income Statement"]
        ws_bs = wb["Balance Sheet"]

        # Parse Fiscal Years and Period-End Dates from Headers (Rows 4 and 5)
        # Find column bounds
        fys: List[str] = []
        period_ends: List[str] = []
        col_indices: List[int] = []

        for col in range(2, ws_is.max_column + 1):
            fy_val = ws_is.cell(row=4, column=col).value
            pend_val = ws_is.cell(row=5, column=col).value
            if fy_val and str(fy_val).strip().startswith("FY"):
                fys.append(str(fy_val).replace("FY", "").strip())
                period_ends.append(str(pend_val).strip() if pend_val else "")
                col_indices.append(col)

        if not fys:
            raise FinancialRatioError("No fiscal year columns detected in Income Statement worksheet.")

        def read_row_by_title_prefix(ws: Any, prefixes: List[str]) -> Dict[str, Optional[float]]:
            for r in range(6, ws.max_row + 1):
                cell_title = str(ws.cell(row=r, column=1).value or "").strip().lower()
                for pref in prefixes:
                    if pref.lower() in cell_title:
                        vals = {}
                        for fy, col in zip(fys, col_indices):
                            c_val = ws.cell(row=r, column=col).value
                            try:
                                vals[fy] = float(c_val) if c_val is not None and str(c_val) != "N/A" else None
                            except (ValueError, TypeError):
                                vals[fy] = None
                        return vals
            return {fy: None for fy in fys}

        # Income Statement items
        revenue = read_row_by_title_prefix(ws_is, ["Revenue (Total Net Sales)", "Revenue", "Sales"])
        net_income = read_row_by_title_prefix(ws_is, ["Net Income"])

        # Balance Sheet items
        current_assets = read_row_by_title_prefix(ws_bs, ["Total Current Assets"])
        current_liabilities = read_row_by_title_prefix(ws_bs, ["Total Current Liabilities"])
        total_assets = read_row_by_title_prefix(ws_bs, ["Total Assets"])
        equity = read_row_by_title_prefix(ws_bs, ["Total Shareholders' Equity", "Stockholders' Equity"])
        commercial_paper = read_row_by_title_prefix(ws_bs, ["Commercial Paper"])
        current_debt = read_row_by_title_prefix(ws_bs, ["Current Term Debt", "Current Portion of Long-Term Debt"])
        long_term_debt = read_row_by_title_prefix(ws_bs, ["Long-Term Debt"])
        total_liabilities = read_row_by_title_prefix(ws_bs, ["Total Liabilities"])

        # Company Title from cell A1
        company_name = str(ws_is["A1"].value or f"{clean_ticker}").split("(")[0].strip()

        return {
            "ticker": clean_ticker,
            "company_name": company_name,
            "fiscal_years": fys,
            "period_ends": period_ends,
            "revenue": revenue,
            "net_income": net_income,
            "current_assets": current_assets,
            "current_liabilities": current_liabilities,
            "total_assets": total_assets,
            "equity": equity,
            "commercial_paper": commercial_paper,
            "current_debt": current_debt,
            "long_term_debt": long_term_debt,
            "total_liabilities": total_liabilities,
        }

    def calculate_ratios(
        self, ticker: str, inputs: Optional[Dict[str, Any]] = None, cik: Optional[str] = None
    ) -> Dict[str, Any]:
        """Calculate the four required financial ratios and supporting metrics across 10 years.

        Args:
            ticker (str): Company ticker symbol.
            inputs (Optional[Dict[str, Any]]): Extracted accounting inputs; loads from workbook if None.
            cik (Optional[str]): Company CIK (used for prior-year asset lookup).

        Returns:
            Dict[str, Any]: Complete ratios, calculation details, and validation checks.
        """
        clean_ticker = ticker.strip().upper()
        data = inputs or self.extract_financial_inputs_from_workbook(clean_ticker)

        fys = data["fiscal_years"]
        period_ends = data["period_ends"]

        # Resolve CIK if not provided
        resolved_cik = cik
        if not resolved_cik:
            from sec_downloader import SECDownloader
            try:
                downloader = SECDownloader()
                resolved_cik, _ = downloader.get_company_cik(clean_ticker)
            except Exception:
                resolved_cik = ""

        # -------------------------------------------------------------------
        # 1. Beginning Assets Selection (Opening Balance for Earliest Year)
        # -------------------------------------------------------------------
        earliest_fy = int(fys[0]) if fys and fys[0].isdigit() else 2016
        prior_year_assets = None
        if resolved_cik:
            prior_year_assets = self.get_prior_year_assets_from_facts(resolved_cik, earliest_fy)

        beg_assets_dict: Dict[str, Optional[float]] = {}
        avg_assets_dict: Dict[str, Optional[float]] = {}
        asset_turnover_note = "Average Total Assets = (Beginning Assets + Ending Assets) / 2"

        for idx, fy in enumerate(fys):
            end_a = data["total_assets"].get(fy)
            if idx == 0:
                if prior_year_assets is not None:
                    beg_a = prior_year_assets
                    asset_turnover_note += f". FY{fy} opening assets (${beg_a:,.1f}M) retrieved from preceding 10-K disclosures"
                else:
                    beg_a = end_a
                    asset_turnover_note += f". FY{fy} opening assets unavailable from SEC facts; ending assets used as proxy (methodological limitation)"
            else:
                prev_fy = fys[idx - 1]
                beg_a = data["total_assets"].get(prev_fy)

            beg_assets_dict[fy] = beg_a
            if beg_a is not None and end_a is not None:
                avg_assets_dict[fy] = round((beg_a + end_a) / 2.0, 1)
            else:
                avg_assets_dict[fy] = None

        # -------------------------------------------------------------------
        # 2. Total Interest-Bearing Debt Construction
        # -------------------------------------------------------------------
        total_debt_dict: Dict[str, Optional[float]] = {}
        for fy in fys:
            cp = data["commercial_paper"].get(fy) or 0.0
            cd = data["current_debt"].get(fy) or 0.0
            ltd = data["long_term_debt"].get(fy) or 0.0
            # If all components are None/absent, mark as None
            if (
                data["commercial_paper"].get(fy) is None
                and data["current_debt"].get(fy) is None
                and data["long_term_debt"].get(fy) is None
            ):
                total_debt_dict[fy] = None
            else:
                total_debt_dict[fy] = round(cp + cd + ltd, 1)

        # -------------------------------------------------------------------
        # 3. Ratio Calculations & Safety Division
        # -------------------------------------------------------------------
        current_ratios: Dict[str, Optional[float]] = {}
        debt_to_equity_ratios: Dict[str, Optional[float]] = {}
        asset_turnovers: Dict[str, Optional[float]] = {}
        net_profit_margins: Dict[str, Optional[float]] = {}

        for fy in fys:
            # A. Current Ratio = CA / CL
            ca = data["current_assets"].get(fy)
            cl = data["current_liabilities"].get(fy)
            if ca is not None and cl is not None and cl != 0:
                current_ratios[fy] = round(ca / cl, 4)
            else:
                current_ratios[fy] = None

            # B. Debt-to-Equity = Total Debt / Equity
            td = total_debt_dict.get(fy)
            eq = data["equity"].get(fy)
            if td is not None and eq is not None and eq != 0:
                debt_to_equity_ratios[fy] = round(td / eq, 4)
            else:
                debt_to_equity_ratios[fy] = None

            # C. Total Asset Turnover = Net Revenue / Average Total Assets
            rev = data["revenue"].get(fy)
            avg_a = avg_assets_dict.get(fy)
            if rev is not None and avg_a is not None and avg_a != 0:
                asset_turnovers[fy] = round(rev / avg_a, 4)
            else:
                asset_turnovers[fy] = None

            # D. Net Profit Margin = (Net Income / Net Revenue) * 100
            ni = data["net_income"].get(fy)
            if ni is not None and rev is not None and rev != 0:
                net_profit_margins[fy] = round((ni / rev) * 100.0, 4)
            else:
                net_profit_margins[fy] = None

        # -------------------------------------------------------------------
        # 4. Rigorous Validation and Accounting Integrity Checks
        # -------------------------------------------------------------------
        validation = self._validate_ratios(
            fys=fys,
            data=data,
            total_debt=total_debt_dict,
            avg_assets=avg_assets_dict,
            current_ratios=current_ratios,
            debt_to_equity=debt_to_equity_ratios,
            asset_turnovers=asset_turnovers,
            net_margins=net_profit_margins,
        )

        return {
            "ticker": clean_ticker,
            "company_name": data["company_name"],
            "fiscal_years": fys,
            "period_ends": period_ends,
            "ratios": {
                "current_ratio": current_ratios,
                "debt_to_equity": debt_to_equity_ratios,
                "asset_turnover": asset_turnovers,
                "net_profit_margin": net_profit_margins,
            },
            "calculation_details": {
                "current_assets": data["current_assets"],
                "current_liabilities": data["current_liabilities"],
                "commercial_paper": data["commercial_paper"],
                "current_debt": data["current_debt"],
                "long_term_debt": data["long_term_debt"],
                "total_interest_bearing_debt": total_debt_dict,
                "shareholders_equity": data["equity"],
                "beginning_total_assets": beg_assets_dict,
                "ending_total_assets": data["total_assets"],
                "average_total_assets": avg_assets_dict,
                "net_revenue": data["revenue"],
                "net_income": data["net_income"],
                "total_liabilities": data["total_liabilities"],
            },
            "methodology_notes": {
                "liquidity": "Current Ratio = Current Assets / Current Liabilities",
                "solvency": "Debt-to-Equity Ratio = Total Interest-Bearing Debt / Shareholders' Equity (Includes Commercial Paper, Current Term Debt, and Non-Current Long-Term Debt)",
                "efficiency": asset_turnover_note,
                "profitability": "Net Profit Margin = (Net Income / Net Revenue) * 100",
            },
            "validation": validation,
        }

    def _validate_ratios(
        self,
        fys: List[str],
        data: Dict[str, Any],
        total_debt: Dict[str, Optional[float]],
        avg_assets: Dict[str, Optional[float]],
        current_ratios: Dict[str, Optional[float]],
        debt_to_equity: Dict[str, Optional[float]],
        asset_turnovers: Dict[str, Optional[float]],
        net_margins: Dict[str, Optional[float]],
    ) -> Dict[str, Any]:
        """Perform automated validation checks on calculated ratios."""
        checks: Dict[str, Any] = {
            "all_ratios_calculated": True,
            "no_duplicate_years": len(fys) == len(set(fys)),
            "debt_less_than_total_liabilities": True,
            "warnings": [],
            "status_by_year": {},
        }

        for fy in fys:
            yr_ok = True
            # 1. Check all 4 ratios exist
            cr = current_ratios.get(fy)
            de = debt_to_equity.get(fy)
            at = asset_turnovers.get(fy)
            npm = net_margins.get(fy)

            if any(r is None for r in (cr, de, at, npm)):
                checks["all_ratios_calculated"] = False
                yr_ok = False
                checks["warnings"].append(f"FY{fy}: Incomplete ratio calculation (missing numerator or denominator).")

            # 2. Check debt <= total liabilities (prevents double-counting / misclassification)
            td = total_debt.get(fy)
            tl = data["total_liabilities"].get(fy)
            if td is not None and tl is not None:
                if td > (tl + 1.0):  # allow $1M rounding tolerance
                    checks["debt_less_than_total_liabilities"] = False
                    yr_ok = False
                    checks["warnings"].append(
                        f"FY{fy}: Interest-bearing debt (${td}M) exceeds total liabilities (${tl}M). Possible double counting."
                    )

            # 3. Check for unusual economic boundaries
            if cr is not None and (cr < 0.4 or cr > 5.0):
                checks["warnings"].append(f"FY{fy}: Unusual Current Ratio value: {cr:.2f}x.")

            if de is not None and de < 0:
                checks["warnings"].append(f"FY{fy}: Negative Debt-to-Equity ({de:.2f}x) due to negative shareholders' equity.")

            if npm is not None and (npm < -20.0 or npm > 70.0):
                checks["warnings"].append(f"FY{fy}: Unusual Net Profit Margin: {npm:.1f}%.")

            if at is not None and (at < 0.2 or at > 3.0):
                checks["warnings"].append(f"FY{fy}: Unusual Asset Turnover: {at:.2f}x.")

            checks["status_by_year"][fy] = "Verified" if yr_ok else "Warning"

        return checks

    def export_ratios_csv(
        self, ticker: str, ratios_data: Dict[str, Any], output_path: Optional[Path] = None
    ) -> Path:
        """Export comprehensive 10-year financial ratios and supporting data to CSV.

        Args:
            ticker (str): Company ticker symbol.
            ratios_data (Dict[str, Any]): Dictionary returned by calculate_ratios.
            output_path (Optional[Path]): Target CSV file path.

        Returns:
            Path: Path to the generated CSV file.
        """
        clean_ticker = ticker.strip().upper()
        target_path = (
            output_path
            or (self.financial_dir / f"{clean_ticker}_financial_ratios.csv")
        )
        target_path.parent.mkdir(parents=True, exist_ok=True)

        fieldnames = [
            "company_ticker",
            "fiscal_year",
            "period_end_date",
            "current_ratio",
            "debt_to_equity_ratio",
            "total_asset_turnover",
            "net_profit_margin_pct",
            "current_assets_usd_m",
            "current_liabilities_usd_m",
            "commercial_paper_usd_m",
            "current_debt_usd_m",
            "long_term_debt_usd_m",
            "total_interest_bearing_debt_usd_m",
            "shareholders_equity_usd_m",
            "beginning_total_assets_usd_m",
            "ending_total_assets_usd_m",
            "average_total_assets_usd_m",
            "net_revenue_usd_m",
            "net_income_usd_m",
            "validation_status",
        ]

        fys = ratios_data["fiscal_years"]
        period_ends = ratios_data["period_ends"]
        ratios = ratios_data["ratios"]
        details = ratios_data["calculation_details"]
        val_status = ratios_data["validation"]["status_by_year"]

        with open(target_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for idx, fy in enumerate(fys):
                writer.writerow({
                    "company_ticker": clean_ticker,
                    "fiscal_year": fy,
                    "period_end_date": period_ends[idx] if idx < len(period_ends) else "",
                    "current_ratio": ratios["current_ratio"].get(fy),
                    "debt_to_equity_ratio": ratios["debt_to_equity"].get(fy),
                    "total_asset_turnover": ratios["asset_turnover"].get(fy),
                    "net_profit_margin_pct": ratios["net_profit_margin"].get(fy),
                    "current_assets_usd_m": details["current_assets"].get(fy),
                    "current_liabilities_usd_m": details["current_liabilities"].get(fy),
                    "commercial_paper_usd_m": details["commercial_paper"].get(fy),
                    "current_debt_usd_m": details["current_debt"].get(fy),
                    "long_term_debt_usd_m": details["long_term_debt"].get(fy),
                    "total_interest_bearing_debt_usd_m": details["total_interest_bearing_debt"].get(fy),
                    "shareholders_equity_usd_m": details["shareholders_equity"].get(fy),
                    "beginning_total_assets_usd_m": details["beginning_total_assets"].get(fy),
                    "ending_total_assets_usd_m": details["ending_total_assets"].get(fy),
                    "average_total_assets_usd_m": details["average_total_assets"].get(fy),
                    "net_revenue_usd_m": details["net_revenue"].get(fy),
                    "net_income_usd_m": details["net_income"].get(fy),
                    "validation_status": val_status.get(fy, "Verified"),
                })

        logger.info(f"Saved financial ratios CSV output to {target_path}")
        return target_path

    def add_financial_ratios_worksheet_to_excel(
        self,
        ticker: str,
        ratios_data: Dict[str, Any],
        workbook_path: Optional[Path] = None,
    ) -> Path:
        """Create and integrate the 'Financial Ratios' worksheet and 4 trend charts into Excel.

        Preserves all existing worksheets ('Income Statement', 'Balance Sheet',
        'Cash Flow Statement', 'Sentiment Analysis', 'Stock Prices', 'Audit & Reconciliation').

        Args:
            ticker (str): Company ticker symbol.
            ratios_data (Dict[str, Any]): Dictionary returned by calculate_ratios.
            workbook_path (Optional[Path]): Path to the target Excel file.

        Returns:
            Path: Path to the updated Excel workbook.
        """
        clean_ticker = ticker.strip().upper()
        target_wb_path = (
            workbook_path
            or (self.excel_dir / clean_ticker / f"{clean_ticker}_financial_statements.xlsx")
        )

        if not target_wb_path.exists():
            raise FinancialRatioError(
                f"Excel workbook not found at {target_wb_path}. Run Stage 4 first."
            )

        wb = openpyxl.load_workbook(target_wb_path)
        logger.info(
            f"Loaded existing Excel workbook: {target_wb_path.name} (Sheets: {wb.sheetnames})"
        )

        sheet_title = "Financial Ratios"
        # If sheet already exists, remove it cleanly to refresh
        if sheet_title in wb.sheetnames:
            wb.remove(wb[sheet_title])

        # Position sheet before 'Audit & Reconciliation' if present, else append
        if "Audit & Reconciliation" in wb.sheetnames:
            insert_idx = wb.sheetnames.index("Audit & Reconciliation")
            ws = wb.create_sheet(title=sheet_title, index=insert_idx)
        else:
            ws = wb.create_sheet(title=sheet_title)

        ws.views.sheetView[0].showGridLines = True

        # Styling constants
        font_name = "Calibri"
        navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        subtotal_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")

        title_font = Font(name=font_name, size=15, bold=True, color="1F4E78")
        subtitle_font = Font(name=font_name, size=10, italic=True, color="595959")
        meta_font = Font(name=font_name, size=9, color="595959")
        header_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
        section_font = Font(name=font_name, size=10, bold=True, color="1F4E78")
        bold_font = Font(name=font_name, size=10, bold=True)
        regular_font = Font(name=font_name, size=10)

        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        align_center = Alignment(horizontal="center", vertical="center")

        border_thin = Side(style="thin", color="D9D9D9")
        border_dark_thin = Side(style="thin", color="000000")
        border_double = Side(style="double", color="000000")

        cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
        subtotal_border = Border(top=border_dark_thin, bottom=border_thin, left=border_thin, right=border_thin)
        summary_border = Border(top=border_dark_thin, bottom=border_double, left=border_thin, right=border_thin)

        # Number formats
        multiple_format = '0.00"x"'
        percent_format = "0.0%"
        currency_format = '$#,##0.0;($#,##0.0);"-"'

        # Rows 1-3: Headers
        ws["A1"] = f"{ratios_data['company_name']} ({clean_ticker})"
        ws["A1"].font = title_font
        ws.row_dimensions[1].height = 24

        ws["A2"] = "TEN-YEAR FINANCIAL RATIO ANALYSIS (LIQUIDITY, SOLVENCY, EFFICIENCY, PROFITABILITY)"
        ws["A2"].font = subtitle_font
        ws.row_dimensions[2].height = 18

        ws["A3"] = "Amounts in USD Millions ($M) | Formulas link directly to underlying calculation details below"
        ws["A3"].font = meta_font
        ws.row_dimensions[3].height = 16

        fys = ratios_data["fiscal_years"]
        period_ends = ratios_data["period_ends"]
        num_fys = len(fys)

        # Row 4: Fiscal Year Header
        ws.cell(row=4, column=1, value="Fiscal Year").font = header_font
        ws.cell(row=4, column=1).fill = navy_fill
        ws.cell(row=4, column=1).alignment = align_left
        ws.row_dimensions[4].height = 22

        # Row 5: Period End Date
        ws.cell(row=5, column=1, value="Fiscal Period-End Date").font = header_font
        ws.cell(row=5, column=1).fill = navy_fill
        ws.cell(row=5, column=1).alignment = align_left
        ws.row_dimensions[5].height = 18

        for idx, (fy, pend) in enumerate(zip(fys, period_ends), start=2):
            c4 = ws.cell(row=4, column=idx, value=f"FY {fy}")
            c4.font = header_font
            c4.fill = navy_fill
            c4.alignment = align_center

            c5 = ws.cell(row=5, column=idx, value=pend)
            c5.font = Font(name=font_name, size=9, bold=True, color="FFFFFF")
            c5.fill = navy_fill
            c5.alignment = align_center

        # Row 6: Section Header - Financial Ratios Summary
        ws.cell(row=6, column=1, value="1. FINANCIAL RATIOS SUMMARY").font = section_font
        ws.cell(row=6, column=1).fill = section_fill
        for c in range(2, num_fys + 2):
            ws.cell(row=6, column=c).fill = section_fill
        ws.row_dimensions[6].height = 20

        # Rows 7-10: Four Financial Ratios (Using Live Excel Formulas)
        ratio_rows_config = [
            (7, "Current Ratio (Liquidity)", "14", "15", multiple_format, "CA / CL"),
            (8, "Debt-to-Equity Ratio (Solvency)", "21", "22", multiple_format, "Total Debt / Equity"),
            (9, "Total Asset Turnover (Efficiency)", "28", "27", multiple_format, "Revenue / Average Assets"),
            (10, "Net Profit Margin (Profitability)", "31", "32", percent_format, "Net Income / Revenue"),
        ]

        for r_idx, title, num_row, den_row, n_fmt, desc in ratio_rows_config:
            ws.cell(row=r_idx, column=1, value=title).font = bold_font
            ws.cell(row=r_idx, column=1).alignment = align_left
            ws.cell(row=r_idx, column=1).border = cell_border
            ws.row_dimensions[r_idx].height = 20

            for col in range(2, num_fys + 2):
                col_letter = get_column_letter(col)
                formula_str = f"={col_letter}{num_row}/{col_letter}{den_row}"
                cell = ws.cell(row=r_idx, column=col, value=formula_str)
                cell.font = bold_font
                cell.alignment = align_right
                cell.number_format = n_fmt
                cell.border = cell_border

        # Row 11: Blank
        ws.row_dimensions[11].height = 10

        # Row 12: Section Header - Calculation Details
        ws.cell(row=12, column=1, value="2. SUPPORTING ACCOUNTING DATA & CALCULATION DETAILS ($M)").font = section_font
        ws.cell(row=12, column=1).fill = section_fill
        for c in range(2, num_fys + 2):
            ws.cell(row=12, column=c).fill = section_fill
        ws.row_dimensions[12].height = 20

        details = ratios_data["calculation_details"]

        # Row 13: Sub-section Liquidity
        ws.cell(row=13, column=1, value="Liquidity Inputs:").font = Font(name=font_name, size=10, bold=True, italic=True)
        # Row 14: Current Assets
        # Row 15: Current Liabilities
        calc_rows = [
            (14, "  Total Current Assets", details["current_assets"], regular_font, cell_border, None),
            (15, "  Total Current Liabilities", details["current_liabilities"], regular_font, cell_border, None),
            (16, None, None, None, None, None),  # blank
            (17, "Solvency Inputs:", None, Font(name=font_name, size=10, bold=True, italic=True), None, None),
            (18, "  Commercial Paper / Short-Term Borrowings", details["commercial_paper"], regular_font, cell_border, None),
            (19, "  Current Portion of Long-Term Debt", details["current_debt"], regular_font, cell_border, None),
            (20, "  Long-Term Debt (Non-Current)", details["long_term_debt"], regular_font, cell_border, None),
            (21, "  Total Interest-Bearing Debt", None, bold_font, subtotal_border, "SUM_DEBT"),
            (22, "  Total Shareholders' Equity", details["shareholders_equity"], regular_font, cell_border, None),
            (23, None, None, None, None, None),  # blank
            (24, "Efficiency Inputs:", None, Font(name=font_name, size=10, bold=True, italic=True), None, None),
            (25, "  Beginning Total Assets", details["beginning_total_assets"], regular_font, cell_border, None),
            (26, "  Ending Total Assets", details["ending_total_assets"], regular_font, cell_border, None),
            (27, "  Average Total Assets", None, bold_font, subtotal_border, "AVG_ASSETS"),
            (28, "  Net Revenue (Total Net Sales)", details["net_revenue"], regular_font, cell_border, None),
            (29, None, None, None, None, None),  # blank
            (30, "Profitability Inputs:", None, Font(name=font_name, size=10, bold=True, italic=True), None, None),
            (31, "  Net Income", details["net_income"], regular_font, cell_border, None),
            (32, "  Net Revenue (Total Net Sales)", details["net_revenue"], regular_font, cell_border, None),
        ]

        for r_idx, title, val_dict, f_style, b_style, special in calc_rows:
            if title is None:
                ws.row_dimensions[r_idx].height = 8
                continue

            ws.cell(row=r_idx, column=1, value=title).font = f_style or regular_font
            ws.cell(row=r_idx, column=1).alignment = align_left
            if b_style:
                ws.cell(row=r_idx, column=1).border = b_style
            ws.row_dimensions[r_idx].height = 18

            if special == "SUM_DEBT":
                for col in range(2, num_fys + 2):
                    col_letter = get_column_letter(col)
                    cell = ws.cell(row=r_idx, column=col, value=f"=SUM({col_letter}18:{col_letter}20)")
                    cell.font = bold_font
                    cell.alignment = align_right
                    cell.number_format = currency_format
                    cell.fill = subtotal_fill
                    cell.border = b_style
            elif special == "AVG_ASSETS":
                for col in range(2, num_fys + 2):
                    col_letter = get_column_letter(col)
                    cell = ws.cell(row=r_idx, column=col, value=f"=AVERAGE({col_letter}25:{col_letter}26)")
                    cell.font = bold_font
                    cell.alignment = align_right
                    cell.number_format = currency_format
                    cell.fill = subtotal_fill
                    cell.border = b_style
            elif val_dict:
                for col, fy in enumerate(fys, start=2):
                    v = val_dict.get(fy)
                    cell = ws.cell(row=r_idx, column=col, value=v if v is not None else 0.0)
                    cell.font = f_style or regular_font
                    cell.alignment = align_right
                    cell.number_format = currency_format
                    if b_style:
                        cell.border = b_style

        # Row 33: Blank
        ws.row_dimensions[33].height = 12

        # Row 34: Section Header - Visualizations
        ws.cell(row=34, column=1, value="3. SUPPORTING VISUALIZATIONS (10-YEAR FINANCIAL RATIO TRENDS)").font = section_font
        ws.cell(row=34, column=1).fill = section_fill
        for c in range(2, num_fys + 2):
            ws.cell(row=34, column=c).fill = section_fill
        ws.row_dimensions[34].height = 20

        # -------------------------------------------------------------------
        # 5. Add 4 Professional Excel Trend Charts (2x2 Grid)
        # -------------------------------------------------------------------
        chart_configs = [
            (7, "Current Ratio Trend (10-Year Liquidity)", "Current Ratio (x)", "B36", 13),
            (8, "Debt-to-Equity Ratio Trend (10-Year Solvency)", "Debt-to-Equity (x)", "H36", 14),
            (9, "Total Asset Turnover Trend (10-Year Efficiency)", "Asset Turnover (x)", "B52", 15),
            (10, "Net Profit Margin Trend (10-Year Profitability)", "Net Profit Margin (%)", "H52", 16),
        ]

        cats = Reference(ws, min_col=2, min_row=4, max_col=num_fys + 1, max_row=4)

        for row_idx, chart_title, y_title, cell_pos, chart_style in chart_configs:
            chart = LineChart()
            chart.title = chart_title
            chart.style = chart_style
            chart.y_axis.title = y_title
            chart.x_axis.title = "Fiscal Year"
            chart.width = 15
            chart.height = 10
            chart.legend = None  # Hide redundant single-series legend

            data_ref = Reference(ws, min_col=1, min_row=row_idx, max_col=num_fys + 1, max_row=row_idx)
            chart.add_data(data_ref, titles_from_data=True, from_rows=True)
            chart.set_categories(cats)
            ws.add_chart(chart, cell_pos)

        # Freeze Panes at Cell B7 so row titles stay anchored during horizontal scroll
        ws.freeze_panes = "B7"

        # Set Column Widths
        ws.column_dimensions["A"].width = 44
        for col in range(2, num_fys + 2):
            col_letter = get_column_letter(col)
            ws.column_dimensions[col_letter].width = 15

        # Save updated workbook
        wb.save(target_wb_path)
        logger.info(
            f"Successfully integrated worksheet '{sheet_title}' with 4 charts into {target_wb_path}"
        )
        return target_wb_path
