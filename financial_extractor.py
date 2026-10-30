"""Financial Statement Extraction and Excel Integration Module.

This module retrieves, structures, and reconciles 10 years of annual financial
statements (Income Statement, Balance Sheet, Statement of Cash Flows) from
official SEC EDGAR Company Facts XBRL data, cross-referencing downloaded 10-K
filings.

Outputs a professionally formatted, multi-tab Excel workbook with freeze panes,
custom accounting formatting, and rigorous accounting reconciliations.
"""

import csv
import json
import logging
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import config
from sec_downloader import SECDownloader, SECDownloaderError

logger = logging.getLogger("Financial-Extractor")


class FinancialExtractionError(Exception):
    """Base exception for financial extraction operations."""


class FinancialExtractor:
    """Extracts, reconciles, and exports 10-K financial statements to Excel."""

    def __init__(
        self,
        user_agent: Optional[str] = None,
        financial_dir: Optional[Path] = None,
        excel_dir: Optional[Path] = None,
    ) -> None:
        """Initialize the FinancialExtractor.

        Args:
            user_agent (Optional[str]): SEC User-Agent header string.
            financial_dir (Optional[Path]): Directory for caching facts.
            excel_dir (Optional[Path]): Directory for saving Excel workbooks.
        """
        self.user_agent = user_agent or config.get_sec_user_agent()
        self.financial_dir = financial_dir or config.FINANCIAL_DATA_DIR
        self.excel_dir = excel_dir or config.EXCEL_OUTPUTS_DIR
        self.downloader = SECDownloader(user_agent=self.user_agent)

    def get_company_facts(self, cik: str) -> Dict[str, Any]:
        """Fetch and cache official SEC Company Facts XBRL data for a company.

        Args:
            cik (str): 10-digit CIK string.

        Returns:
            Dict[str, Any]: Loaded JSON facts dictionary.

        Raises:
            FinancialExtractionError: If facts cannot be retrieved from SEC.
        """
        cik_10 = cik.zfill(10)
        cache_path = self.financial_dir / f"companyfacts_{cik_10}.json"

        # Check local cache first (valid for 7 days)
        if cache_path.exists():
            try:
                age_days = (time.time() - cache_path.stat().st_mtime) / 86400.0
                if age_days < 7.0 and cache_path.stat().st_size > 10000:
                    with open(cache_path, "r", encoding="utf-8") as f:
                        return json.load(f)
            except Exception as read_err:
                logger.warning(f"Error reading facts cache: {read_err}. Re-fetching...")

        url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik_10}.json"
        logger.info(f"Retrieving official SEC Company Facts from {url}...")

        try:
            resp = self.downloader._get(url)
            facts_data = resp.json()
            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(facts_data, f)
            logger.info(f"Cached Company Facts ({len(resp.content):,} bytes) to {cache_path.name}")
            return facts_data
        except Exception as err:
            raise FinancialExtractionError(
                f"Failed to retrieve SEC Company Facts for CIK {cik_10}: {err}"
            ) from err

    def get_reporting_periods(self, ticker: str, count: int = 10) -> List[Dict[str, str]]:
        """Identify the 10 distinct annual reporting periods (FY and Period End Date).

        Uses the Stage 2 manifest if present; otherwise queries SEC submissions.

        Args:
            ticker (str): Normalized ticker symbol.
            count (int): Number of fiscal years (default: 10).

        Returns:
            List[Dict[str, str]]: Sorted list of period dicts [{'fy': '2016', 'period_end': '2016-09-24'}, ...]
        """
        clean_ticker = ticker.strip().upper()
        manifest_path = config.RAW_FILINGS_DIR / clean_ticker / f"{clean_ticker}_10K_manifest.csv"

        periods: Dict[str, str] = {}
        if manifest_path.exists():
            try:
                with open(manifest_path, "r", encoding="utf-8") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        fy = row.get("fiscal_year", "")
                        pend = row.get("fiscal_period_end_date", "")
                        if fy and pend:
                            periods[fy] = pend
            except Exception as err:
                logger.warning(f"Could not read manifest {manifest_path}: {err}")

        # If manifest was incomplete or missing, query submissions
        if len(periods) < count:
            cik, company_title = self.downloader.get_company_cik(clean_ticker)
            meta_list = self.downloader.get_10k_filings_metadata(
                cik=cik, ticker=clean_ticker, company_name=company_title, count=count
            )
            for m in meta_list:
                periods[m["fiscal_year"]] = m["fiscal_period_end_date"]

        # Sort fiscal years ascending (oldest to newest for horizontal display)
        sorted_fys = sorted(periods.keys())[-count:]
        result = [{"fiscal_year": fy, "period_end": periods[fy]} for fy in sorted_fys]
        return result

    @staticmethod
    def _extract_fact_value(
        gaap_facts: Dict[str, Any],
        tag_list: List[str],
        period_end: str,
        is_duration: bool = True,
    ) -> Optional[float]:
        """Extract fact value for a target date matching duration or instant criteria.

        Args:
            gaap_facts (Dict[str, Any]): us-gaap facts dictionary from SEC.
            tag_list (List[str]): List of candidate XBRL tags in priority order.
            period_end (str): Target period-end date (YYYY-MM-DD).
            is_duration (bool): True for duration (Income/CF), False for instant (BS).

        Returns:
            Optional[float]: Value in USD, or None if unavailable.
        """
        for tag in tag_list:
            if tag not in gaap_facts:
                continue

            units = gaap_facts[tag].get("units", {}).get("USD", [])
            if not units:
                continue

            # Prioritize 10-K filings, then latest filed date (audited/restated)
            filtered = [
                it for it in units
                if it.get("end") == period_end and it.get("form") in ("10-K", "10-K/A")
            ]
            if not filtered:
                # Secondary check: allow any form if exact end date matches
                filtered = [it for it in units if it.get("end") == period_end]

            filtered.sort(key=lambda x: (x.get("form") == "10-K", x.get("filed", "")), reverse=True)

            for it in filtered:
                val = it.get("val")
                if val is None:
                    continue

                if is_duration:
                    start = it.get("start")
                    if start:
                        try:
                            d_start = datetime.strptime(start, "%Y-%m-%d")
                            d_end = datetime.strptime(period_end, "%Y-%m-%d")
                            days = (d_end - d_start).days
                            # Annual duration is approximately 330 to 390 days
                            if 320 <= days <= 395:
                                return float(val)
                        except ValueError:
                            pass
                    # If fp is FY without explicit start, accept if marked 10-K
                    elif it.get("fp") == "FY":
                        return float(val)
                else:
                    # Instant fact (Balance Sheet)
                    if "start" not in it:
                        return float(val)

        return None

    def extract_financial_data(
        self, ticker: str, count: int = 10
    ) -> Dict[str, Any]:
        """Extract all 3 financial statements across 10 annual reporting periods.

        Args:
            ticker (str): Company ticker symbol.
            count (int): Number of periods to extract (default: 10).

        Returns:
            Dict[str, Any]: Dictionary containing metadata, periods, statements, and reconciliations.
        """
        clean_ticker = ticker.strip().upper()
        cik, company_title = self.downloader.get_company_cik(clean_ticker)
        facts_data = self.get_company_facts(cik)
        gaap = facts_data.get("facts", {}).get("us-gaap", {})

        periods = self.get_reporting_periods(clean_ticker, count=count)
        fys = [p["fiscal_year"] for p in periods]
        period_ends = [p["period_end"] for p in periods]

        logger.info(f"Extracting financial statements for {company_title} ({clean_ticker}) across {len(fys)} years: {fys}")

        # -------------------------------------------------------------------
        # Statement Definitions (Display Name, Tags, is_duration, is_subtotal, is_total)
        # -------------------------------------------------------------------
        is_schema = [
            ("Revenue (Total Net Sales)", ["RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet", "Revenues"], True, False, False),
            ("Cost of Revenue (Cost of Goods Sold)", ["CostOfGoodsAndServicesSold", "CostOfRevenue", "CostOfGoodsSold"], True, False, False),
            ("Gross Profit", ["GrossProfit"], True, True, False),
            ("Research and Development (R&D)", ["ResearchAndDevelopmentExpense"], True, False, False),
            ("Selling, General and Administrative (SG&A)", ["SellingGeneralAndAdministrativeExpense", "SellingAndMarketingExpense"], True, False, False),
            ("Total Operating Expenses", ["OperatingExpenses"], True, True, False),
            ("Operating Income (Loss)", ["OperatingIncomeLoss"], True, True, False),
            ("Total Non-Operating / Other Income (Expense)", ["NonoperatingIncomeExpense", "OtherNonoperatingIncomeExpense"], True, False, False),
            ("Interest and Investment Income", ["InvestmentIncomeInterestAndDividend", "InvestmentIncomeInterest", "InterestAndOtherIncome", "InvestmentIncomeNet"], True, False, False),
            ("Interest Expense", ["InterestExpense", "InterestExpenseDebt", "InterestExpenseNonoperating"], True, False, False),
            ("Income Before Provision for Income Taxes", ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"], True, True, False),
            ("Provision for Income Taxes", ["IncomeTaxExpenseBenefit"], True, False, False),
            ("Net Income", ["NetIncomeLoss", "ProfitLoss"], True, False, True),
        ]

        bs_schema = [
            # Current Assets
            ("Cash and Cash Equivalents", ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"], False, False, False),
            ("Marketable Securities (Current)", ["MarketableSecuritiesCurrent", "AvailableForSaleSecuritiesCurrent"], False, False, False),
            ("Accounts Receivable, Net", ["AccountsReceivableNetCurrent"], False, False, False),
            ("Inventories", ["InventoryNet"], False, False, False),
            ("Vendor Non-Trade Receivables", ["NontradeReceivablesCurrent"], False, False, False),
            ("Other Current Assets", ["OtherAssetsCurrent", "PrepaidExpenseAndOtherAssetsCurrent"], False, False, False),
            ("Total Current Assets", ["AssetsCurrent"], False, True, False),
            # Non-Current Assets
            ("Marketable Securities (Non-Current)", ["MarketableSecuritiesNoncurrent", "AvailableForSaleSecuritiesNoncurrent"], False, False, False),
            ("Property, Plant and Equipment, Net", ["PropertyPlantAndEquipmentNet"], False, False, False),
            ("Other Non-Current Assets", ["OtherAssetsNoncurrent"], False, False, False),
            ("Total Non-Current Assets", ["AssetsNoncurrent"], False, True, False),
            ("Total Assets", ["Assets"], False, False, True),
            # Current Liabilities
            ("Accounts Payable", ["AccountsPayableCurrent"], False, False, False),
            ("Other Current Liabilities", ["OtherLiabilitiesCurrent", "AccruedLiabilitiesCurrent"], False, False, False),
            ("Deferred Revenue (Current)", ["ContractWithCustomerLiabilityCurrent", "DeferredRevenueCurrent"], False, False, False),
            ("Commercial Paper", ["CommercialPaper"], False, False, False),
            ("Current Term Debt", ["DebtCurrent", "LongTermDebtCurrent", "ConvertibleDebtCurrent"], False, False, False),
            ("Total Current Liabilities", ["LiabilitiesCurrent"], False, True, False),
            # Non-Current Liabilities
            ("Long-Term Debt", ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations", "LongTermDebt"], False, False, False),
            ("Other Non-Current Liabilities", ["OtherLiabilitiesNoncurrent"], False, False, False),
            ("Total Non-Current Liabilities", ["LiabilitiesNoncurrent"], False, True, False),
            ("Total Liabilities", ["Liabilities"], False, True, False),
            # Mezzanine / Temporary Equity
            ("Temporary / Mezzanine Equity", ["TemporaryEquityValueExcludingAdditionalPaidInCapital", "TemporaryEquityCarryingAmountAttributableToParent", "TemporaryEquityCarryingAmount", "RedeemableNoncontrollingInterestEquityCarryingAmount"], False, False, False),
            # Shareholders' Equity
            ("Common Stock and Additional Paid-In Capital", ["CommonStocksIncludingAdditionalPaidInCapital", "CommonStockValue"], False, False, False),
            ("Retained Earnings (Accumulated Deficit)", ["RetainedEarningsAccumulatedDeficit"], False, False, False),
            ("Accumulated Other Comprehensive Income (Loss)", ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"], False, False, False),
            ("Total Shareholders' Equity", ["StockholdersEquity", "CommonStockholdersEquity"], False, True, False),
            ("Total Liabilities and Shareholders' Equity", ["LiabilitiesAndStockholdersEquity"], False, False, True),
        ]

        cf_schema = [
            ("Net Income", ["NetIncomeLoss", "ProfitLoss"], True, False, False),
            ("Depreciation and Amortization", ["DepreciationDepletionAndAmortization", "DepreciationAmortizationAndAccretionNet"], True, False, False),
            ("Share-Based Compensation", ["ShareBasedCompensation"], True, False, False),
            ("Other Non-Cash Operating Adjustments", ["OtherNoncashIncomeExpense"], True, False, False),
            ("Change in Accounts Receivable", ["IncreaseDecreaseInAccountsReceivable"], True, False, False),
            ("Change in Inventories", ["IncreaseDecreaseInInventories"], True, False, False),
            ("Change in Accounts Payable", ["IncreaseDecreaseInAccountsPayable"], True, False, False),
            ("Change in Other Operating Assets and Liabilities", ["IncreaseDecreaseInOtherOperatingCapitalNet", "IncreaseDecreaseInOtherOperatingLiabilities"], True, False, False),
            ("Net Cash Provided by Operating Activities", ["NetCashProvidedByUsedInOperatingActivities"], True, True, False),
            ("Payments for Property, Plant and Equipment (CapEx)", ["PaymentsToAcquireProductiveAssets", "PaymentsToAcquirePropertyPlantAndEquipment"], True, False, False),
            ("Purchases of Marketable Securities", ["PaymentsToAcquireMarketableSecurities", "PaymentsToAcquireAvailableForSaleSecurities"], True, False, False),
            ("Proceeds from Sales/Maturities of Marketable Securities", ["ProceedsFromSaleAndMaturityOfMarketableSecurities", "ProceedsFromSaleOfAvailableForSaleSecurities", "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities"], True, False, False),
            ("Other Investing Activities", ["PaymentsToAcquireBusinessesNetOfCashAcquired", "PaymentsToAcquireOtherInvestments"], True, False, False),
            ("Net Cash Provided by (Used in) Investing Activities", ["NetCashProvidedByUsedInInvestingActivities"], True, True, False),
            ("Payments for Common Stock Repurchases", ["PaymentsForRepurchaseOfCommonStock"], True, False, False),
            ("Payments for Dividends and Dividend Equivalents", ["PaymentsOfDividendsCommonStock", "PaymentsOfDividends"], True, False, False),
            ("Proceeds from / (Repayments of) Debt, Net", ["ProceedsFromIssuanceOfLongTermDebt", "ProceedsFromIssuanceOfDebt", "RepaymentsOfLongTermDebt"], True, False, False),
            ("Other Financing Activities", ["ProceedsFromIssuanceOfCommonStock"], True, False, False),
            ("Net Cash Provided by (Used in) Financing Activities", ["NetCashProvidedByUsedInFinancingActivities"], True, True, False),
            ("Effect of Exchange Rate Changes on Cash", ["EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents", "EffectOfExchangeRateOnCashAndCashEquivalents"], True, False, False),
            ("Net Increase (Decrease) in Cash and Cash Equivalents", ["CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect", "CashAndCashEquivalentsPeriodIncreaseDecrease"], True, False, True),
        ]

        def build_statement_rows(schema: List[Tuple[str, List[str], bool, bool, bool]]) -> List[Dict[str, Any]]:
            rows = []
            for line_title, tags, is_dur, is_sub, is_tot in schema:
                vals_by_fy = {}
                for p in periods:
                    fy = p["fiscal_year"]
                    pend = p["period_end"]
                    v = self._extract_fact_value(gaap, tags, pend, is_duration=is_dur)
                    # Convert to USD Millions rounded to 1 decimal place ($M)
                    vals_by_fy[fy] = round(v / 1e6, 1) if v is not None else None

                rows.append({
                    "title": line_title,
                    "values": vals_by_fy,
                    "is_subtotal": is_sub,
                    "is_total": is_tot,
                })
            return rows

        income_rows = build_statement_rows(is_schema)
        balance_rows = build_statement_rows(bs_schema)
        cashflow_rows = build_statement_rows(cf_schema)

        # -------------------------------------------------------------------
        # Post-Processing & Logical Inferences
        # -------------------------------------------------------------------
        for fy in fys:
            # 1. Income Statement: Gross profit = Revenue - Cost of Revenue (if Gross Profit missing)
            rev_row = next((r for r in income_rows if "Revenue" in r["title"] and not r["is_subtotal"]), None)
            cogs_row = next((r for r in income_rows if "Cost of Revenue" in r["title"]), None)
            gp_row = next((r for r in income_rows if r["title"] == "Gross Profit"), None)
            if gp_row and gp_row["values"].get(fy) is None and rev_row and cogs_row:
                if rev_row["values"].get(fy) is not None and cogs_row["values"].get(fy) is not None:
                    gp_row["values"][fy] = round(rev_row["values"][fy] - cogs_row["values"][fy], 1)

            # 2. Balance Sheet: Total Non-Current Assets = Total Assets - Total Current Assets
            tot_assets_row = next((r for r in balance_rows if r["title"] == "Total Assets"), None)
            cur_assets_row = next((r for r in balance_rows if r["title"] == "Total Current Assets"), None)
            noncur_assets_row = next((r for r in balance_rows if r["title"] == "Total Non-Current Assets"), None)
            if noncur_assets_row and noncur_assets_row["values"].get(fy) is None:
                if tot_assets_row and cur_assets_row and tot_assets_row["values"].get(fy) is not None and cur_assets_row["values"].get(fy) is not None:
                    noncur_assets_row["values"][fy] = round(tot_assets_row["values"][fy] - cur_assets_row["values"][fy], 1)

            # 3. Balance Sheet: Total Non-Current Liabilities = Total Liabilities - Total Current Liabilities
            tot_liab_row = next((r for r in balance_rows if r["title"] == "Total Liabilities"), None)
            cur_liab_row = next((r for r in balance_rows if r["title"] == "Total Current Liabilities"), None)
            noncur_liab_row = next((r for r in balance_rows if r["title"] == "Total Non-Current Liabilities"), None)
            if noncur_liab_row and noncur_liab_row["values"].get(fy) is None:
                if tot_liab_row and cur_liab_row and tot_liab_row["values"].get(fy) is not None and cur_liab_row["values"].get(fy) is not None:
                    noncur_liab_row["values"][fy] = round(tot_liab_row["values"][fy] - cur_liab_row["values"][fy], 1)

            # 4. Balance Sheet: Total Liabilities and Equity (only compute if missing from XBRL)
            tot_mezzanine_row = next((r for r in balance_rows if "Mezzanine" in r["title"] or "Temporary" in r["title"]), None)
            tot_equity_row = next((r for r in balance_rows if r["title"] == "Total Shareholders' Equity"), None)
            tot_liab_eq_row = next((r for r in balance_rows if r["title"] == "Total Liabilities and Shareholders' Equity"), None)
            if tot_liab_eq_row and tot_liab_eq_row["values"].get(fy) is None:
                l = (tot_liab_row["values"].get(fy) or 0.0) if tot_liab_row else 0.0
                m = (tot_mezzanine_row["values"].get(fy) or 0.0) if tot_mezzanine_row else 0.0
                e = (tot_equity_row["values"].get(fy) or 0.0) if tot_equity_row else 0.0
                tot_liab_eq_row["values"][fy] = round(l + m + e, 1)

        # -------------------------------------------------------------------
        # Accounting Reconciliations & Quality Controls
        # -------------------------------------------------------------------
        reconciliations = self._run_reconciliations(
            fys=fys,
            income_rows=income_rows,
            balance_rows=balance_rows,
            cashflow_rows=cashflow_rows,
        )

        return {
            "ticker": clean_ticker,
            "company_name": company_title,
            "cik": cik,
            "periods": periods,
            "income_statement": income_rows,
            "balance_sheet": balance_rows,
            "cash_flows": cashflow_rows,
            "reconciliations": reconciliations,
        }

    def _run_reconciliations(
        self,
        fys: List[str],
        income_rows: List[Dict[str, Any]],
        balance_rows: List[Dict[str, Any]],
        cashflow_rows: List[Dict[str, Any]],
    ) -> Dict[str, Any]:
        """Perform automated accounting validation and reconciliation checks.

        Args:
            fys (List[str]): List of fiscal year strings.
            income_rows (List[Dict[str, Any]]): Income statement data.
            balance_rows (List[Dict[str, Any]]): Balance sheet data.
            cashflow_rows (List[Dict[str, Any]]): Cash flow statement data.

        Returns:
            Dict[str, Any]: Dictionary of validation results and discrepancies.
        """
        results: Dict[str, Any] = {
            "balance_sheet_balanced": True,
            "balance_sheet_checks": {},
            "gross_profit_checks": {},
            "cash_flow_checks": {},
            "warnings": [],
        }

        # Row lookups
        tot_assets = next((r for r in balance_rows if r["title"] == "Total Assets"), None)
        tot_liab = next((r for r in balance_rows if r["title"] == "Total Liabilities"), None)
        tot_mezzanine = next((r for r in balance_rows if "Mezzanine" in r["title"] or "Temporary" in r["title"]), None)
        tot_eq = next((r for r in balance_rows if r["title"] == "Total Shareholders' Equity"), None)
        tot_liab_eq = next((r for r in balance_rows if r["title"] == "Total Liabilities and Shareholders' Equity"), None)

        rev_r = next((r for r in income_rows if "Revenue" in r["title"] and not r["is_subtotal"]), None)
        cogs_r = next((r for r in income_rows if "Cost of Revenue" in r["title"]), None)
        gp_r = next((r for r in income_rows if r["title"] == "Gross Profit"), None)

        op_cf = next((r for r in cashflow_rows if r["title"] == "Net Cash Provided by Operating Activities"), None)
        inv_cf = next((r for r in cashflow_rows if r["title"] == "Net Cash Provided by (Used in) Investing Activities"), None)
        fin_cf = next((r for r in cashflow_rows if r["title"] == "Net Cash Provided by (Used in) Financing Activities"), None)
        net_cf = next((r for r in cashflow_rows if r["title"] == "Net Increase (Decrease) in Cash and Cash Equivalents"), None)

        for fy in fys:
            # 1. Balance Sheet: Assets = Liabilities + Equity (tolerance: $1.5M for rounding)
            if tot_assets:
                a = tot_assets["values"].get(fy)
                le = None
                if tot_liab_eq and tot_liab_eq["values"].get(fy) is not None:
                    le = tot_liab_eq["values"][fy]
                elif tot_liab and tot_eq:
                    l = tot_liab["values"].get(fy) or 0.0
                    m = (tot_mezzanine["values"].get(fy) or 0.0) if tot_mezzanine else 0.0
                    e = tot_eq["values"].get(fy) or 0.0
                    le = round(l + m + e, 1)

                if a is not None and le is not None:
                    diff = round(a - le, 1)
                    is_ok = abs(diff) <= 1.5
                    results["balance_sheet_checks"][fy] = {
                        "assets": a,
                        "liab_plus_eq": le,
                        "diff": diff,
                        "balanced": is_ok,
                    }
                    if not is_ok:
                        results["balance_sheet_balanced"] = False
                        results["warnings"].append(
                            f"FY{fy} Balance Sheet out of balance by ${diff}M (Assets: ${a}M, Liab+Eq: ${le}M)"
                        )

            # 2. Gross Profit Check: GP = Revenue - COGS
            if rev_r and cogs_r and gp_r:
                r_val = rev_r["values"].get(fy)
                c_val = cogs_r["values"].get(fy)
                g_val = gp_r["values"].get(fy)
                if r_val is not None and c_val is not None and g_val is not None:
                    gp_diff = round(g_val - (r_val - c_val), 1)
                    gp_ok = abs(gp_diff) <= 1.0
                    results["gross_profit_checks"][fy] = {
                        "revenue": r_val,
                        "cogs": c_val,
                        "reported_gp": g_val,
                        "calculated_gp": round(r_val - c_val, 1),
                        "diff": gp_diff,
                        "reconciled": gp_ok,
                    }
                    if not gp_ok:
                        results["warnings"].append(
                            f"FY{fy} Gross Profit discrepancy of ${gp_diff}M"
                        )

            # 3. Cash Flow Check: Sum of Operating + Investing + Financing
            if op_cf and inv_cf and fin_cf:
                o_v = op_cf["values"].get(fy)
                i_v = inv_cf["values"].get(fy)
                f_v = fin_cf["values"].get(fy)
                n_v = net_cf["values"].get(fy) if net_cf else None
                if o_v is not None and i_v is not None and f_v is not None and n_v is not None:
                    cf_sum = round(o_v + i_v + f_v, 1)
                    cf_diff = round(n_v - cf_sum, 1)
                    # Note: Exchange rate effect explains minor difference if any
                    results["cash_flow_checks"][fy] = {
                        "operating": o_v,
                        "investing": i_v,
                        "financing": f_v,
                        "sum_components": cf_sum,
                        "reported_net": n_v,
                        "diff_fx_unexplained": cf_diff,
                    }

        return results

    def create_excel_workbook(
        self, financial_data: Dict[str, Any], output_path: Optional[Path] = None
    ) -> Path:
        """Create a professionally styled multi-tab Excel workbook with financial data.

        Args:
            financial_data (Dict[str, Any]): Financial data dictionary from extract_financial_data.
            output_path (Optional[Path]): Destination path. If None, saves in data/excel_outputs.

        Returns:
            Path: Path to the generated Excel file.
        """
        ticker = financial_data["ticker"]
        company_name = financial_data["company_name"]
        cik = financial_data["cik"]
        periods = financial_data["periods"]
        fys = [p["fiscal_year"] for p in periods]

        company_excel_dir = self.excel_dir / ticker
        company_excel_dir.mkdir(parents=True, exist_ok=True)

        target_file = output_path or (company_excel_dir / f"{ticker}_financial_statements.xlsx")

        wb = openpyxl.Workbook()
        # Remove default sheet
        wb.remove(wb.active)

        # Styling definitions
        font_name = "Calibri"
        navy_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        subtotal_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
        title_font = Font(name=font_name, size=15, bold=True, color="1F4E78")
        subtitle_font = Font(name=font_name, size=10, italic=True, color="595959")
        header_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
        regular_font = Font(name=font_name, size=10)
        subtotal_font = Font(name=font_name, size=10, bold=True)
        total_font = Font(name=font_name, size=11, bold=True)

        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        border_thin = Side(style="thin", color="D9D9D9")
        border_top_thick = Side(style="thin", color="000000")
        border_bottom_double = Side(style="double", color="000000")

        regular_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
        subtotal_border = Border(top=border_top_thick, bottom=border_thin)
        total_border = Border(top=border_top_thick, bottom=border_bottom_double)

        # Standard accounting number format: negative in parentheses, dash for zero
        number_format = "$#,##0.0;($#,##0.0);\"-\""

        sheets_info = [
            ("Income Statement", "CONSOLIDATED STATEMENTS OF OPERATIONS", financial_data["income_statement"]),
            ("Balance Sheet", "CONSOLIDATED BALANCE SHEETS", financial_data["balance_sheet"]),
            ("Cash Flow Statement", "CONSOLIDATED STATEMENTS OF CASH FLOWS", financial_data["cash_flows"]),
        ]

        for sheet_name, stmt_heading, row_data in sheets_info:
            ws = wb.create_sheet(title=sheet_name)
            ws.views.sheetView[0].showGridLines = True

            # Row 1: Company Title
            ws["A1"] = f"{company_name} ({ticker})"
            ws["A1"].font = title_font
            ws.row_dimensions[1].height = 24

            # Row 2: Statement Heading
            ws["A2"] = f"{stmt_heading} (CIK: {cik})"
            ws["A2"].font = Font(name=font_name, size=12, bold=True, color="333333")
            ws.row_dimensions[2].height = 18

            # Row 3: Unit Notice
            ws["A3"] = "Amounts in USD Millions ($M) | Extracted from SEC Form 10-K Filings"
            ws["A3"].font = subtitle_font
            ws.row_dimensions[3].height = 16

            # Row 4: Column Headers (Row 4 = FY, Row 5 = Period End Date)
            ws.row_dimensions[4].height = 22
            ws.row_dimensions[5].height = 18

            ws["A4"] = "Line Item / Reporting Period"
            ws["A5"] = "Fiscal Period-End Date"
            ws["A4"].fill = navy_header_fill
            ws["A5"].fill = navy_header_fill
            ws["A4"].font = header_font
            ws["A5"].font = Font(name=font_name, size=9, bold=True, color="FFFFFF")
            ws["A4"].alignment = align_left
            ws["A5"].alignment = align_left

            col_idx = 2
            for p in periods:
                col_letter = get_column_letter(col_idx)
                cell_fy = ws[f"{col_letter}4"]
                cell_end = ws[f"{col_letter}5"]

                cell_fy.value = f"FY {p['fiscal_year']}"
                cell_end.value = p["period_end"]

                cell_fy.fill = navy_header_fill
                cell_end.fill = navy_header_fill
                cell_fy.font = header_font
                cell_end.font = Font(name=font_name, size=9, bold=True, color="FFFFFF")
                cell_fy.alignment = align_center
                cell_end.alignment = align_center

                col_idx += 1

            # Freeze Panes below headers (Row 6) and right of line items (Column B)
            ws.freeze_panes = "B6"

            # Populate Rows
            current_row = 6
            for item in row_data:
                ws.row_dimensions[current_row].height = 20
                title_cell = ws.cell(row=current_row, column=1, value=item["title"])
                title_cell.alignment = align_left

                is_sub = item.get("is_subtotal", False)
                is_tot = item.get("is_total", False)

                if is_tot:
                    title_cell.font = total_font
                    row_border = total_border
                    row_fill = subtotal_fill
                elif is_sub:
                    title_cell.font = subtotal_font
                    row_border = subtotal_border
                    row_fill = subtotal_fill
                else:
                    title_cell.font = regular_font
                    row_border = regular_border
                    row_fill = None

                title_cell.border = row_border
                if row_fill:
                    title_cell.fill = row_fill

                col_num = 2
                for p in periods:
                    fy = p["fiscal_year"]
                    val = item["values"].get(fy)
                    data_cell = ws.cell(row=current_row, column=col_num)

                    if val is not None:
                        data_cell.value = val
                        data_cell.number_format = number_format
                    else:
                        data_cell.value = "N/A"
                        data_cell.alignment = align_center

                    if is_tot:
                        data_cell.font = total_font
                    elif is_sub:
                        data_cell.font = subtotal_font
                    else:
                        data_cell.font = regular_font

                    data_cell.border = row_border
                    if row_fill:
                        data_cell.fill = row_fill
                    if val is not None:
                        data_cell.alignment = align_right

                    col_num += 1

                current_row += 1

            # Auto-fit column widths
            ws.column_dimensions["A"].width = 46
            for col_i in range(2, len(periods) + 2):
                col_letter = get_column_letter(col_i)
                ws.column_dimensions[col_letter].width = 16

        # -------------------------------------------------------------------
        # Worksheet 4: Reconciliations & Quality Audit Metadata
        # -------------------------------------------------------------------
        ws_audit = wb.create_sheet(title="Audit & Reconciliation")
        ws_audit.views.sheetView[0].showGridLines = True

        ws_audit["A1"] = f"{company_name} ({ticker}) - Data Quality & Reconciliation Audit"
        ws_audit["A1"].font = title_font
        ws_audit["A2"] = f"CIK: {cik} | Reporting Periods: {len(periods)} Years ({periods[0]['fiscal_year']} - {periods[-1]['fiscal_year']})"
        ws_audit["A2"].font = subtitle_font

        ws_audit["A4"] = "Accounting Reconciliation Tests"
        ws_audit["A4"].font = Font(name=font_name, size=12, bold=True, color="1F4E78")

        headers = ["Fiscal Year", "Period End Date", "BS Assets ($M)", "BS Liab + Eq ($M)", "BS Difference ($M)", "BS Balanced?", "Gross Profit Reconciled?", "Cash Flow Net ($M)"]
        ws_audit.row_dimensions[5].height = 22
        for i, h in enumerate(headers, start=1):
            cell = ws_audit.cell(row=5, column=i, value=h)
            cell.fill = navy_header_fill
            cell.font = header_font
            cell.alignment = align_center

        row_idx = 6
        rec = financial_data["reconciliations"]
        for p in periods:
            fy = p["fiscal_year"]
            pend = p["period_end"]
            bs_chk = rec["balance_sheet_checks"].get(fy, {})
            gp_chk = rec["gross_profit_checks"].get(fy, {})
            cf_chk = rec["cash_flow_checks"].get(fy, {})

            ws_audit.cell(row=row_idx, column=1, value=f"FY {fy}").alignment = align_center
            ws_audit.cell(row=row_idx, column=2, value=pend).alignment = align_center

            assets_val = bs_chk.get("assets", "N/A")
            c3 = ws_audit.cell(row=row_idx, column=3, value=assets_val)
            if isinstance(assets_val, (int, float)):
                c3.number_format = number_format
                c3.alignment = align_right

            liab_eq_val = bs_chk.get("liab_plus_eq", "N/A")
            c4 = ws_audit.cell(row=row_idx, column=4, value=liab_eq_val)
            if isinstance(liab_eq_val, (int, float)):
                c4.number_format = number_format
                c4.alignment = align_right

            diff_val = bs_chk.get("diff", "N/A")
            c5 = ws_audit.cell(row=row_idx, column=5, value=diff_val)
            if isinstance(diff_val, (int, float)):
                c5.number_format = "$#,##0.0;($#,##0.0);\"0.0\""
                c5.alignment = align_right

            c6 = ws_audit.cell(row=row_idx, column=6, value="YES [PASS]" if bs_chk.get("balanced") else "NO [CHECK]")
            c6.alignment = align_center
            c6.font = Font(name=font_name, bold=True, color="006100" if bs_chk.get("balanced") else "9C0006")

            gp_ok = gp_chk.get("reconciled", True)
            c7 = ws_audit.cell(row=row_idx, column=7, value="YES [PASS]" if gp_ok else f"Diff: ${gp_chk.get('diff', 0)}M")
            c7.alignment = align_center
            c7.font = Font(name=font_name, bold=True, color="006100" if gp_ok else "9C0006")

            cf_net = cf_chk.get("reported_net", "N/A")
            c8 = ws_audit.cell(row=row_idx, column=8, value=cf_net)
            if isinstance(cf_net, (int, float)):
                c8.number_format = number_format
                c8.alignment = align_right

            for c in range(1, 9):
                ws_audit.cell(row=row_idx, column=c).border = regular_border

            row_idx += 1

        # Audit column widths
        ws_audit.column_dimensions["A"].width = 14
        ws_audit.column_dimensions["B"].width = 16
        ws_audit.column_dimensions["C"].width = 18
        ws_audit.column_dimensions["D"].width = 20
        ws_audit.column_dimensions["E"].width = 20
        ws_audit.column_dimensions["F"].width = 16
        ws_audit.column_dimensions["G"].width = 24
        ws_audit.column_dimensions["H"].width = 20

        wb.save(target_file)
        logger.info(f"Excel workbook successfully written to {target_file}")
        return target_file
