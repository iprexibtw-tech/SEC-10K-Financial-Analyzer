# SEC 10-K Gen AI Accounting Analysis Application

An undergraduate Gen AI accounting application developed in Python to retrieve, process, extract, and analyze 10 years of annual SEC 10-K filings for Apple (AAPL), NVIDIA (NVDA), and other publicly traded U.S. corporations.

Features a modern **Streamlit Financial Analytics Dashboard** as the primary interactive web interface, a standalone Windows Desktop Application (CustomTkinter / PyInstaller), and a modular command-line pipeline.

---

## 🌟 Streamlit Web Financial Analytics Dashboard (`streamlit_app.py`)

The primary interface for this project is an executive-grade, browser-based financial analytics dashboard built with **Streamlit** and **Plotly**. It delivers interactive longitudinal accounting, financial statement modeling, sentiment intelligence, and valuation analysis.

### Local Dashboard Startup Command
Run the dashboard directly in your web browser using the project virtual environment:
```powershell
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```
Or with standard active Python:
```bash
python -m streamlit run streamlit_app.py
```
* **Local Web URL**: `http://localhost:8501`
* **Default Theme**: Dark executive analytics theme with custom styling.

### Key Dashboard Capabilities
1. **Interactive Company Selection & History Switcher**:
   - Quick-switch between pre-analyzed corporations (**Apple Inc.** and **NVIDIA Corporation**) with instantaneous reloading of all financial models, statements, charts, and metrics without re-downloading or recalculating.
   - Enter any exchange ticker symbol (e.g. `MSFT`, `GOOGL`, `TSLA`, `AMZN`) or search company names across 10,430+ SEC registrants.
   - User-selectable longitudinal study scope (Default: 10 annual Form 10-K filings; options: 3, 5, 10, 15, 20).
2. **Executive Overview Metrics (8 Dynamic Cards)**:
   - Displays key financial metrics for the most recent fiscal reporting period (e.g. FY2025 for AAPL, FY2026 for NVDA):
     - **Annual Net Revenue** ($ in Millions and Billions)
     - **Annual Net Income** ($ in Millions and Billions)
     - **Current Ratio** (with Working Capital benchmark indicators)
     - **Debt-to-Equity Ratio** (Solvency leverage)
     - **Total Asset Turnover** (Capital efficiency)
     - **Net Profit Margin** (Margin expansion/compression %)
     - **Sentiment P/N Ratio** (Loughran-McDonald Positive/Negative ratio with exact word counts)
     - **Closing Stock Price** (Actual split-adjusted closing price at the balance sheet date)
3. **Three Interactive Financial Statements (Tabs)**:
   - 10 fiscal years presented side by side for **Income Statement**, **Balance Sheet**, and **Statement of Cash Flows**.
   - Currency formatting, consistent financial units ($ in Millions), and horizontal scrolling.
   - **Audit & Reconciliation Tab**: displays mathematical verification ensuring zero discrepancies on the balance sheet equation ($A = L + SE$) and gross profit reconciliations.
4. **Interactive Plotly Financial Visualizations**:
   - **Stock Price vs. Sentiment P/N Ratio (Mandatory Chart)**: Dual-axis longitudinal line chart linking closing equity price (left green axis) with management MD&A sentiment tone (right amber dashed axis).
   - **Revenue & Net Income Growth**: Grouped 10-year bar charts.
   - **Financial Ratio Trends**: 4-panel grid with 1.0x reference benchmark lines.
   - **MD&A Sentiment Breakdown**: Positive vs. negative financial word counts and sentiment trends.
5. **Excel Workbook & CSV Export Engine**:
   - One-click **Download Complete 8-Tab Excel Analysis (`<TICKER>_10K_Analysis.xlsx`)** button providing the complete financial model with embedded Excel formulas, formatting, and charts.
   - Direct download buttons for raw financial statement workbooks and analytical CSV files (Financial Ratios, Sentiment Analysis, Stock Prices).

---

## Project Structure

```text
Python 10-K/
├── .venv/                         # Windows Python virtual environment (Python 3.15.2)
├── data/                          # Managed storage for filings and outputs
│   ├── raw_filings/               # Downloaded SEC 10-K HTML/text files
│   │   └── <TICKER>/              # Company-specific raw filings (e.g. data/raw_filings/AAPL/)
│   │       ├── AAPL_10K_FY...htm  # Primary Form 10-K filing documents
│   │       └── AAPL_10K_manifest.csv # Verifiable CSV manifest
│   ├── extracted_text/            # Extracted Item 7 MD&A text sections (Stage 3)
│   │   └── <TICKER>/              # Company-specific extracted text (e.g. data/extracted_text/AAPL/)
│   │       ├── AAPL_2025_MDA.txt  # Clean UTF-8 substantive MD&A text
│   │       └── AAPL_mda_manifest.csv # Extraction quality audit manifest
│   ├── cleaned_text/              # Cleaned MD&A text files for sentiment analysis (Stage 5)
│   │   └── <TICKER>/              # Company-specific cleaned text (e.g. data/cleaned_text/AAPL/)
│   │       ├── AAPL_2025_MDA_CLEANED.txt # Normalized text for Loughran-McDonald scoring
│   │       └── AAPL_cleaning_manifest.csv # Text cleaning audit manifest
│   ├── financial_data/            # Structured financial statements & SEC registries
│   │   ├── company_tickers.json   # SEC official company tickers registry
│   │   ├── companyfacts_<CIK>.json# Official SEC Company Facts XBRL JSON cache
│   │   ├── AAPL_sentiment_analysis.csv # Sentiment metrics CSV export (Stage 6)
│   │   ├── AAPL_stock_prices.csv  # Historical stock price CSV export (Stage 7)
│   │   └── AAPL_financial_ratios.csv # 10-year financial ratios & details CSV export (Stage 8)
│   └── excel_outputs/             # Final formatted multi-tab Excel workbooks (Stages 4–9)
│       ├── AAPL_10K_Analysis.xlsx # Primary comprehensive 8-tab final analysis workbook (Stage 9)
│       └── <TICKER>/              # Company-specific workbook storage
│           ├── AAPL_10K_Analysis.xlsx # Comprehensive 8-tab final analysis workbook (Stage 9)
│           └── AAPL_financial_statements.xlsx # 8-tab synchronized financial statements & analysis
├── .gitignore                     # Ignores .venv, cache, and temporary data
├── .sec_settings                  # Persisted SEC User-Agent configuration
├── config.py                      # Central configuration, directory definitions, and diagnostics
├── excel_report.py                # Dedicated final reporting, dual-axis charting & Excel integration (Stage 9)
├── financial_extractor.py         # Financial statement XBRL extraction and Excel generator (Stage 4)
├── financial_ratios.py            # Ten-year financial ratio analysis, charting & Excel generator (Stage 8)
├── Instructions.docx              # Primary assignment specification document
├── main.py                        # Central CLI application entry point and 8-stage orchestrator
├── mda_extractor.py               # Item 7 MD&A text extraction and quality validation engine (Stage 3)
├── PROMPT_LOG.md                  # Verbatim chronological log of user prompts for lab report appendix
├── README.md                      # Comprehensive user guide, architecture, and execution docs
├── requirements.txt               # Dependencies (requests, beautifulsoup4, lxml, openpyxl, yfinance, pandas)
├── sec_downloader.py              # SEC EDGAR Form 10-K retrieval and caching engine (Stage 2)
├── sentiment_analysis.py          # Loughran-McDonald dictionary loader, scorer & Excel engine (Stage 6)
├── stock_prices.py                # Historical stock price retrieval, market calendar & Excel engine (Stage 7)
├── text_cleaner.py                # MD&A text cleaning, normalization & tokenization engine (Stage 5)
├── test_stage2.py                 # Automated unit tests for SEC retrieval logic
├── test_stage3.py                 # Automated unit tests for Item 7 MD&A extraction logic
├── test_stage4.py                 # Automated unit tests for financial statement extraction & Excel
├── test_stage5.py                 # Automated unit tests for text cleaning and tokenization
├── test_stage6.py                 # Automated unit tests for sentiment analysis and Excel integration
├── test_stage7.py                 # Automated unit tests for historical stock prices and Excel integration
├── test_stage8.py                 # Automated unit tests for financial ratios, charting and Excel integration
├── test_stage9.py                 # Automated unit tests for final Excel reporting & pipeline integration
└── test_stage10.py                # Automated unit tests for NVIDIA analysis & dual-company validation
```

---

## Assignment Roadmap (Per `Instructions.docx`)

- **Stage 1 (Completed)**: Project architecture setup, virtual environment, directory structure, and verification test.
- **Stage 2 (Completed)**: Dual company selection engine (direct ticker entry e.g. `AAPL`, `NVDA` and full company-name search e.g. `Apple`, `NVIDIA`, `Microsoft` via official SEC EDGAR registry), user prompt for filing count, SEC EDGAR retrieval via `sec_downloader.py`, primary Form 10-K document downloads, company-specific organization, duplicate caching, and CSV manifest generation.
- **Stage 3 (Completed)**: Item 7 MD&A textual extraction via `mda_extractor.py`, HTML cleanup, TOC discrimination, quality control validation, UTF-8 text file storage (`AAPL_{YEAR}_MDA.txt`, `NVDA_{YEAR}_MDA.txt`), and extraction audit reporting.
- **Stage 4 (Completed)**: Financial statement extraction (Income Statement, Balance Sheet, Statement of Cash Flows) from official SEC EDGAR Company Facts XBRL data, multi-tab professional Excel workbook generation with freeze panes and accounting formatting (`openpyxl`), and automated balance sheet/gross profit accounting reconciliation audits.
- **Stage 5 (Completed)**: MD&A text cleaning, running header/footer removal, typography normalization, uppercase tokenization, quality control validation, and audit manifest generation (`text_cleaner.py`).
- **Stage 6 (Completed)**: Loughran-McDonald Financial Sentiment Dictionary scoring (positive, negative, uncertainty, litigious, constraining, modal terms), P/N ratio calculation, and multi-tab Excel workbook integration (`sentiment_analysis.py`).
- **Stage 7 (Completed)**: Historical stock price retrieval for each balance sheet date via `stock_prices.py` (Yahoo Finance via `yfinance`), automatic non-trading day preceding-day fallback, split and dividend adjustment tracking, multi-tab Excel workbook integration (`Stock Prices` tab), and CSV export.
- **Stage 8 (Completed)**: Ten-year financial ratio analysis via `financial_ratios.py` (Liquidity: Current Ratio, Solvency: Debt-to-Equity, Efficiency: Total Asset Turnover, Profitability: Net Profit Margin), dynamic Excel formula linkage, opening asset balance retrieval, 4 embedded professional trend charts, and CSV export.
- **Stage 9 (Completed)**: Final Excel analysis and visualizations via `excel_report.py`, mandatory dual-axis line chart (Stock Price vs. Sentiment P/N Ratio), 6 trend charts, comprehensive workbook (`{TICKER}_10K_Analysis.xlsx`), and unified 8-stage pipeline coordination in `main.py`.
- **Stage 10 (Completed)**: NVIDIA Corporation (NVDA) 10-year analysis (FY 2017 to FY 2026), dual-company application validation, late-January Sunday fiscal year-end preceding-Friday market alignment (-2d offset), split adjustments (4:1 in 2021, 10:1 in 2024), and generation of complete comprehensive workbook `NVDA_10K_Analysis.xlsx` (8 worksheets, 10 charts, 100% accounting reconciliation PASS).
- **Stage 11 (Upcoming)**: Final comparative undergraduate lab report document.

---

## Company Selection & Search Architecture (`sec_downloader.py` & `main.py`)

The application supports dual company selection workflows in both interactive and scripted environments:
1. **Direct Ticker Symbol Entry**: Users can directly enter registered 1–6 character ticker symbols (e.g. `AAPL`, `NVDA`, `MSFT`). The system verifies the ticker against SEC EDGAR and proceeds immediately without extra prompts.
2. **Company Name Search**: Users can enter plain-English company names (e.g. `Apple`, `NVIDIA`, `Microsoft`, `Google`). The search engine queries official SEC EDGAR registry records (`company_tickers.json`, containing 10,430+ registered corporations) and ranks matching entities.

### Weighted Relevance Ranking Algorithm
The search engine in `sec_downloader.py` ranks candidates using a deterministic weighted scoring system:
- **Exact Ticker Match (+1000)**: Matches query exactly to an exchange ticker symbol.
- **Brand / Parent Alias Map (+900)**: Resolves popular corporate aliases to their registered holding company tickers (e.g. `Google` $\rightarrow$ `GOOGL` / Alphabet Inc.; `Facebook` $\rightarrow$ `META` / Meta Platforms).
- **Exact Title Match (+800)**: Legal corporate title matches query completely.
- **Title Prefix (+500)**: Corporate title starts with the query string (e.g. `Apple Inc.`).
- **Word Boundary Regex (+300)**: Matches query as a distinct word boundary within company title.
- **Multi-Word Token Match (+200)**: Multi-word queries match all constituent words.
- **Substring Match (+100 in title / +50 in ticker)**: General substring matches sorted with shorter title lengths prioritized.

### Disambiguation & User Selection Flow
- **Single Match**: Displays matched corporate title, ticker, and CIK, prompting user confirmation: `Proceed with NVIDIA CORP (NVDA)? [Y/n]:`.
- **Multiple Matches**: Renders an aligned tabular menu of ranked candidates:
  ```text
  Matching companies found for 'Apple':
  ----------------------------------------------------------------------------
    #    Ticker   Company Name                                 CIK         
  ----------------------------------------------------------------------------
    [1]  AAPL     Apple Inc.                                   0000320193  
    [2]  AAPI     Apple iSports Group, Inc.                    0001134982  
    [3]  APLE     Apple Hospitality REIT, Inc.                 0001418121  
    ...
  ----------------------------------------------------------------------------
  Select a company [1-7] (default: 1), [N]ew search, or [C]ancel:
  ```
- **Validation**: Once selected, the company is re-validated through `get_company_cik()` to verify active EDGAR registration before pipeline execution begins.

---

## Financial Statement Extraction Methodology (`financial_extractor.py`)

### 1. SEC Data Sources Used
- **SEC Company Facts API**: `https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json`
  - Provides the complete, official, audited XBRL disclosures submitted by the company in annual 10-K filings.
  - Automatically cached in `data/financial_data/companyfacts_<CIK>.json` (refreshed every 7 days) to eliminate redundant network overhead.
- **Stage 2 10-K Manifest**: Reuses exact fiscal reporting period-end dates (`reportDate`) and accession numbers to ensure 100% alignment across stages.

### 2. Selection of XBRL Concepts
The extractor handles standard US-GAAP accounting tag evolutions and company-specific reporting variations:
* **Duration vs. Instant Concepts**:
  * **Duration concepts** (Income Statement and Cash Flows) measure performance across a full reporting year (~330 to 395 days between `start` and `end`). Quarterly facts are excluded.
  * **Instant concepts** (Balance Sheet) measure financial position at a specific point in time (the balance sheet date with no `start` date).
* **Accounting Standard Transitions**:
  * *Revenue*: Prior to 2018, companies used `SalesRevenueNet` or `Revenues`. Under ASC 606 (adopted 2018/2019), companies adopted `RevenueFromContractWithCustomerExcludingAssessedTax`. The extractor chains these concepts in priority order to produce continuous 10-year timelines.
  * *Marketable Securities*: Handles both `AvailableForSaleSecuritiesCurrent`/`Noncurrent` (used in earlier periods) and modern `MarketableSecuritiesCurrent`/`Noncurrent`.
  * *Cash & Cash Equivalents*: Chains modern standard `CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect` with `CashAndCashEquivalentsPeriodIncreaseDecrease`.
* **Zero vs. Missing**: Unreported or unavailable items are marked as missing (`None` / `N/A`) rather than silently replaced with zero.

---

## Excel Workbook Structure (`data/excel_outputs/<TICKER>/<TICKER>_financial_statements.xlsx`)

The generated workbook contains eight professionally styled worksheets:

1. **Worksheet 1: `Income Statement`**
   - Line items listed vertically.
   - 10 fiscal years displayed horizontally from oldest to newest (FY 2016 through FY 2025).
   - Core items: Revenue, Cost of Revenue, **Gross Profit**, Research & Development (R&D), Selling, General & Administrative (SG&A), **Total Operating Expenses**, **Operating Income**, Non-Operating Income/Expense, Pretax Income, Provision for Income Taxes, **Net Income**.
2. **Worksheet 2: `Balance Sheet`**
   - Current Assets (Cash, Marketable Securities, Accounts Receivable, Inventories, Other Current Assets, **Total Current Assets**).
   - Non-Current Assets (Non-Current Securities, Property, Plant & Equipment, Other Non-Current Assets, **Total Non-Current Assets**, **Total Assets**).
   - Current Liabilities (Accounts Payable, Other Current Liabilities, Deferred Revenue, Commercial Paper, Current Term Debt, **Total Current Liabilities**).
   - Non-Current Liabilities (Long-Term Debt, Other Non-Current Liabilities, **Total Non-Current Liabilities**, **Total Liabilities**).
   - Stockholders' Equity (Common Stock & APIC, Retained Earnings, AOCI, **Total Stockholders' Equity**, **Total Liabilities & Equity**).
3. **Worksheet 3: `Cash Flow Statement`**
   - Operating Activities (Net Income, D&A, Share-Based Compensation, Working Capital Changes, **Net Cash from Operating Activities**).
   - Investing Activities (CapEx, Marketable Securities Purchases/Sales, **Net Cash from Investing Activities**).
   - Financing Activities (Common Stock Repurchases, Dividends Paid, Debt Issuance/Repayments, **Net Cash from Financing Activities**).
   - Net Increase/Decrease in Cash & Cash Equivalents.
4. **Worksheet 4: `Sentiment Analysis`**
   - Full 10-year Loughran-McDonald sentiment metrics & frequencies (Positive, Negative, Uncertainty, Litigious, Constraining, Modal terms, and P/N Ratios).
5. **Worksheet 5: `Stock Prices`**
   - Fiscal year-end balance sheet dates, mapped active trading dates, split-adjusted closing prices (`Close`), dividend-adjusted closing prices (`Adj Close`), trading volume, offset days, data source, and notes.
6. **Worksheet 6: `Financial Ratios`**
   - Ten-year financial ratio analysis (Current Ratio, Debt-to-Equity Ratio, Total Asset Turnover, Net Profit Margin), live Excel formula links to underlying accounts, and 4 embedded trend charts.
7. **Worksheet 7: `Charts and Analysis`**
   - Comprehensive longitudinal visualizations: Mandatory dual-axis line chart (Historical Stock Price vs. MD&A Sentiment P/N Ratio), 4 financial ratio trend charts (Liquidity, Solvency, Efficiency, Profitability), disclosure volume trend chart (Positive vs. Negative word counts), and executive analytical commentary.
8. **Worksheet 8: `Audit & Reconciliation`**
   - Automated accounting reconciliation checks for every reporting period.
   - Flags differences between Total Assets and Total Liabilities + Equity.
   - Reconciles Gross Profit vs. Reported Revenue less Cost of Revenue.
   - Displays Net Cash Flow reconciliation.

### Excel Styling Specifications
- **Header Banners**: Slate Navy (`#1F4E78`) background with white bold typography.
- **Freeze Panes**: Set at cell `B6` across all statement sheets (keeps account descriptions and fiscal year headers visible during scrolling).
- **Accounting Number Formatting**: Standard accounting format `$#,##0.0;($#,##0.0);"-"` with negative figures enclosed in parentheses.
- **Totals & Subtotals**: Emphasized with top thin borders, bottom double accounting borders, and soft shading (`#F2F4F7`).
- **Units**: Clearly stated in banner (`Amounts in USD Millions ($M)`).

---

## Accounting Validation & Reconciliation Results (Apple Inc.)

Every reporting period underwent automated verification:
1. **Balance Sheet Equation**:
   $$\text{Total Assets} = \text{Total Liabilities} + \text{Total Stockholders' Equity}$$
   * Result: **10/10 years perfectly balanced** ($0.0M difference across all 10 periods).
2. **Gross Profit Integrity**:
   $$\text{Gross Profit} = \text{Revenue} - \text{Cost of Goods Sold}$$
   * Result: **10/10 years perfectly reconciled** ($0.0M discrepancy across all 10 periods).
3. **Period Continuity**: Exactly 10 distinct, non-overlapping annual reporting periods identified and verified.

---

## Stage 5: MD&A Text Cleaning and Preparation for Sentiment Analysis

### Purpose of Text Cleaning
Raw MD&A narrative extracted from SEC 10-K filings contains non-substantive formatting artifacts such as running page headers/footers, physical page numbers, decorative bullet icons, non-breaking spaces, HTML entities, and inconsistent quote characters. Stage 5 transforms these extracted narrative files into clean, normalized text and generates an uppercase token stream specifically prepared for dictionary-based sentiment scoring using the Loughran-McDonald Financial Sentiment Dictionary.

### Cleaning Methodology & Normalization Rules
1. **HTML Entity Decoding**: Unescapes standard and numeric HTML entities (e.g., `&amp;` -> `&`, `&#8217;` -> `'`, `&nbsp;` -> space).
2. **Whitespace & Non-Breaking Space Normalization**: Replaces `\xa0` and `\u202f` with ASCII spaces; strips zero-width spaces (`\u200b`) and BOM markers (`\ufeff`). Collapses inline tabs and multiple spaces into single spaces; standardizes paragraph breaks to double newlines (`\n\n`).
3. **Residual Tag Stripping**: Eliminates any lingering XML/HTML markup (`<[^>]+>`).
4. **SEC Running Header & Footer Removal**: Excises repetitive physical page headers and page number lines (e.g., `Apple Inc. | 2024 Form 10-K | 26` or `NVIDIA Corporation | 2024 Form 10-K | 45`).
5. **Punctuation & Quote Standardization**: Normalizes curly smart quotes (`“`, `”`, `‘`, `’`) to standard ASCII quotes (`"`, `'`); unifies em-dashes and en-dashes (`—`, `–`) to hyphens (`-`).
6. **List Bullet & Rule Removal**: Strips non-substantive bullet symbols (`•`, `●`, `▪`, `◦`, `‣`, `∙`) and horizontal divider lines (`---`, `===`).
7. **Strict Vocabulary & Financial Negation Preservation**:
   - Preserves all corporate financial words intact (`loss`, `impairment`, `uncertainty`, `decline`, `litigation`, `debt`, `risk`).
   - Strictly retains negation terms (`not`, `no`, `never`) essential for sentiment context.
   - Retains numerical quantities and currency amounts in the archived cleaned text for human audit and contextual verification.
   - Applies **no stopword removal**, **no stemming**, and **no lemmatization**, preserving exact inflections for Loughran-McDonald dictionary matching.

### Tokenization for Loughran-McDonald Dictionary Matching
- **Token Pattern**: Extracts alphabetic words and hyphenated compound terms (`\b[A-Za-z]+(?:-[A-Za-z]+)*\b`).
- **Standardization**: Converts all tokens to uppercase (e.g., `LONG-TERM`, `REVENUE`, `UNCERTAINTY`).
- **Exclusions**: Numbers and standalone currency symbols are excluded from the sentiment token stream (documented in quality metadata), adhering to financial dictionary conventions.

### File Organization
- **Cleaned Text Storage**: `data/cleaned_text/{TICKER}/{TICKER}_{YEAR}_MDA_CLEANED.txt` (UTF-8 encoded).
- **Cleaning Audit Manifest**: `data/cleaned_text/{TICKER}/{TICKER}_cleaning_manifest.csv`.

### Word Count Metrics & Definitions
- **Original Word Count**: Total whitespace-delimited tokens in raw extracted MD&A (`len(original_text.split())`).
- **Cleaned Word Count**: Total whitespace-delimited tokens in cleaned MD&A text (`len(cleaned_text.split())`).
- **Sentiment Tokens Count**: Count of uppercase alphabetic/compound word tokens prepared for sentiment scoring (`len(tokens)`).
- **Percentage Change**: Calculated as `((Cleaned Words - Original Words) / Original Words) * 100`. (Typical reduction ranges from -1.0% to -2.5%, reflecting purely non-substantive header and bullet removal).

### Quality Control & Validation Checks
- **Non-Empty Check**: Flags any empty or whitespace-only cleaned text as `Failed`.
- **Reduction / Duplication Thresholds**: Flags word count drops exceeding -30.0% (potential truncation) or increases exceeding +10.0% (potential duplication).
- **Accounting Vocabulary Check**: Confirms presence of top-line revenue terminology (`REVENUE`, `REVENUES`, or `SALES`) and core disclosure vocabulary (`OPERATING`, `FINANCIAL`, `RESULTS`, `RISK`, `MANAGEMENT`).
- **Negation Integrity Check**: Confirms essential negation terms (`NOT`, `NO`) remain in the token set.
- **Reproducibility**: Guarantees identical deterministic output across repeated executions.

---

## Stage 6: Loughran-McDonald Sentiment Analysis & Excel Integration

### Purpose & Background
Financial accounting research (Loughran & McDonald, 2011) demonstrates that general-purpose sentiment dictionaries (such as Harvard General Inquirer) misclassify normal business terminology (e.g., *tax*, *cost*, *capital*, *board*, *liability*) as negative, creating false sentiment distortions. The **Loughran-McDonald Financial Sentiment Dictionary** (Master Dictionary 1993–2025) provides domain-specific sentiment classifications tailored specifically for corporate 10-K disclosures.

### Dictionary Auto-Discovery & Category Validation
The module `sentiment_analysis.py` automatically locates the Master Dictionary in `data/` (supporting both `.csv` and `.xlsx` formats) without requiring manual file renaming. It validates that all required financial sentiment dimensions are present:
- **Negative** (2,345 active words): e.g., *loss*, *adverse*, *decline*, *impairment*, *weakness*.
- **Positive** (347 active words): e.g., *improve*, *favorable*, *satisfy*, *innovative*, *profitable*.
- **Uncertainty** (297 active words): e.g., *uncertain*, *risk*, *approximate*, *volatility*, *fluctuation*.
- **Litigious** (903 active words): e.g., *litigation*, *claim*, *jurisdiction*, *lawsuit*, *statutory*.
- **Constraining** (184 active words): e.g., *commit*, *restrict*, *obligate*, *covenant*, *mandate*.
- **Modal Words** (46 active words):
  - *Strong Modal* (19 words): *must*, *will*, *always*, *clearly*, *definitely*.
  - *Weak Modal* (27 words): *could*, *may*, *possible*, *somewhat*, *conceivable*.

### Loughran-McDonald Year-Coded Logic
In the official Master Dictionary CSV:
- **Positive Integer (>0)**: Year the term was added to the sentiment list (indicates active membership).
- **Negative Integer (<0)**: Year the term was deleted/removed from that sentiment list (e.g., `-2020` for *benefit* in Positive). Term is excluded from modern scoring.
- **Zero (0)**: Term is not classified in that sentiment dimension.

### Sentiment Formulas & Relative Metrics
1. **Positive / Negative (P/N) Ratio**:
   $$\text{P/N Ratio} = \frac{\text{Positive Count}}{\text{Negative Count}}$$
   *(Reported as a numerical ratio; if Negative Count is zero, safely outputs `"Undefined"` rather than causing a division-by-zero error).*
2. **Positive Word Frequency (%)**:
   $$\text{Positive Word Frequency} = \frac{\text{Positive Word Count}}{\text{Total Word Count}} \times 100$$
3. **Negative Word Frequency (%)**:
   $$\text{Negative Word Frequency} = \frac{\text{Negative Word Count}}{\text{Total Word Count}} \times 100$$
4. **Net Sentiment Polarity**:
   $$\text{Net Polarity} = \frac{\text{Positive Count} - \text{Negative Count}}{\text{Positive Count} + \text{Negative Count}}$$
5. **Net Sentiment Score (Token Density)**:
   $$\text{Sentiment Score} = \frac{\text{Positive Count} - \text{Negative Count}}{\text{Total Tokens Analyzed}} \times 100$$
6. **Category Proportions (% of Tokens)**:
   $$\text{Category \%} = \frac{\text{Category Word Count}}{\text{Total Tokens Analyzed}} \times 100$$

### Multi-Tab Excel Workbook Structure
The module integrates sentiment analysis into the existing company workbook (`data/excel_outputs/AAPL/AAPL_financial_statements.xlsx`) while strictly preserving all previous financial statement worksheets:
- **Tab 1**: `Income Statement` (Consolidated Statements of Operations, 10 years)
- **Tab 2**: `Balance Sheet` (Consolidated Balance Sheets, 10 years)
- **Tab 3**: `Cash Flow Statement` (Consolidated Statements of Cash Flows, 10 years)
- **Tab 4**: `Sentiment Analysis` (Full 10-year Loughran-McDonald sentiment metrics & frequencies)
- **Tab 5**: `Stock Prices` (Historical fiscal year-end closing and dividend-adjusted stock prices)
- **Tab 6**: `Audit & Reconciliation` (Balance sheet and gross profit accounting integrity checks)

#### Excel Styling Standards for Sentiment Analysis
- **Freeze Panes**: Locked at cell `B6` (row classifications in Column A and fiscal year headers in Rows 4–5 remain stationary during scrolling).
- **Number Formats**: Counts formatted as `#,##0`; P/N Ratio formatted as `0.000` with soft green fill highlight (`#E2EFDA`); Net Polarity formatted as `+0.00%;-0.00%`; Proportions formatted as `0.00%`.
- **Top Financial Terminology**: Qualitative disclosure rows display the top 3 most frequent negative, positive, and uncertainty terms identified in each reporting period.

---

## Stage 7: Historical Stock Price Extraction & Excel Integration

### Purpose & Objective
Stage 7 retrieves and validates accurate historical stock prices corresponding to the fiscal year-end balance sheet dates of each company's 10 annual SEC 10-K filings. These prices establish the market equity valuation baseline required for subsequent valuation metrics and financial ratios (e.g., Price-to-Earnings, Market-to-Book) in Stage 8, as well as longitudinal visualizations in Stage 9.

### Historical Stock Price Data Sources
- **Primary Source**: Yahoo Finance historical market data queried via the `yfinance` package (`yfinance.Ticker.history`).
- **Date Matching Engine**: Localized in `stock_prices.py`. Discovers fiscal period-end balance sheet dates dynamically from SEC filings and metadata.
- **Offline Simulation / Second Source Cache**: A secondary offline historical data registry is maintained within `stock_prices.py` to allow automated verification, deterministic test reproducibility, and fallback cross-checking in network-restricted environments.

### Treatment of Non-Trading Days (Lookahead Bias Prevention)
Financial reporting periods frequently close on non-trading days (weekends or federal market holidays). For instance, Apple Inc. utilizes a 52-53 week fiscal year ending on the last Saturday of September; all 10 of Apple's recent balance sheet dates fall on Saturdays:
1. **Preceding Trading Day Rule**: When a balance sheet date falls on a weekend or market holiday, the matching engine automatically falls back to the **most recent preceding trading day** (e.g., the preceding Friday for Saturday balance sheet dates).
2. **Strict Lookahead Prohibition**: An actual trading date must **never** be later than the fiscal period-end balance sheet date ($\text{Trading Date} \le \text{Balance Sheet Date}$). Using Monday prices following a Saturday balance sheet date would introduce lookahead bias into accounting and financial analysis.
3. **Offset Tracking**: The exact calendar day offset ($\le 0$) is recorded and displayed in both Excel and CSV manifests for full auditability.

### Stock Split & Dividend Adjustments
Corporate actions alter raw share price series over a 10-year longitudinal horizon:
- **Stock Splits (Apple 4-for-1 Split on August 31, 2020)**:
  - Raw unadjusted closing prices prior to August 2020 were denominated in pre-split share quantities (e.g., >$500/share in 2020).
  - The `Close` series provided by Yahoo Finance is **retroactively split-adjusted** across all periods to reflect the modern post-split share basis ($28–$255/share).
  - This allows direct comparability across the 10-year timeline and eliminates artificial valuation cliffs.
- **Cash Dividends & Total Return (`Adj Close`)**:
  - The `Adj Close` series accounts for both split events and cumulative cash dividend reinvestment.
  - As a result, older `Adj Close` prices are lower than split-adjusted `Close` prices ($25.78 vs. $28.18 in FY 2016), reflecting the compounding effect of retained distributions over the subsequent decade.
- **Primary Analytical Metric**:
  - **Split-Adjusted Closing Price (`Close`)** is designated as the primary price measure for accounting ratio calculations and market capitalization valuation, while `Adj Close` is retained for total-return reference.

### Multi-Tab Excel Workbook Structure (`Stock Prices` Tab)
Worksheet 5 (`Stock Prices`) is inserted seamlessly into `data/excel_outputs/<TICKER>/<TICKER>_financial_statements.xlsx` without altering or overwriting any prior statement worksheets:
- **Columns Displayed**:
  - `Fiscal Year`: Formatted as `FY YYYY`.
  - `Balance Sheet Date`: Official 10-K balance sheet period-end date (`YYYY-MM-DD`).
  - `Trading Day`: Actual market trading day used (`YYYY-MM-DD`).
  - `Day Offset`: Days prior to balance sheet date (e.g., `-1 day`).
  - `Split-Adjusted Close ($)`: Split-adjusted closing price formatted in currency (`$#,##0.00`).
  - `Dividend-Adjusted Close ($)`: Total-return adjusted closing price formatted in currency (`$#,##0.00`).
  - `Trading Volume`: Shares traded on market day formatted as integer with commas (`#,##0`).
  - `Data Source`: Provider attribution (e.g., `Yahoo Finance (yfinance)`).
  - `Split / Adjustment Notes`: Contextual documentation on corporate actions (e.g., 4:1 stock split).
- **Styling**:
  - Deep Slate Navy banner (`#1F4E78`) with white bold text.
  - Subheader metadata documenting primary price metric, data source, and adjustment methodology.
  - Freeze Panes applied at cell `B7` so headers remain anchored while scrolling.
  - Auto-fitted column widths with clean borders and centered date alignments.

### CSV Dataset Output
In addition to the multi-tab Excel workbook, structured tabular data is exported to:
- `data/financial_data/{TICKER}_stock_prices.csv`

---

## Stage 8: Ten-Year Financial Ratio Analysis & Visualizations

### Purpose & Objective
Stage 8 evaluates the financial performance, liquidity, leverage, and asset efficiency of Apple Inc. (AAPL) across the 10-year longitudinal period (FY 2016 through FY 2025). The module `financial_ratios.py` computes four foundational financial ratios using verified, audited annual financial statements extracted in Stage 4. It embeds live formulas, calculation details, and four trend charts into the `Financial Ratios` worksheet of `AAPL_financial_statements.xlsx` and exports structured CSV data to `data/financial_data/AAPL_financial_ratios.csv`.

### Definitions & Accounting Formulas

1. **Liquidity — Current Ratio**:
   $$\text{Current Ratio} = \frac{\text{Total Current Assets}}{\text{Total Current Liabilities}}$$
   *Purpose*: Evaluates the company's capability to honor short-term financial obligations with liquid assets over a 12-month operating cycle.

2. **Solvency — Debt-to-Equity Ratio**:
   $$\text{Debt-to-Equity Ratio} = \frac{\text{Total Interest-Bearing Debt}}{\text{Total Shareholders' Equity}}$$
   *Definition of Interest-Bearing Debt*:
   $$\text{Total Interest-Bearing Debt} = \text{Commercial Paper} + \text{Current Portion of Long-Term Debt} + \text{Non-Current Long-Term Debt}$$
   *Accounting Rule*: Operating liabilities (Accounts Payable, Accrued Expenses, Deferred Revenue) are strictly excluded from the debt numerator to prevent distorting financial leverage. The engine confirms that $\text{Total Debt} \le \text{Total Liabilities}$.

3. **Efficiency — Total Asset Turnover**:
   $$\text{Total Asset Turnover} = \frac{\text{Net Revenue}}{\text{Average Total Assets}}$$
   $$\text{Average Total Assets} = \frac{\text{Beginning Total Assets} + \text{Ending Total Assets}}{2}$$
   *Opening Balance Retrieval*: To maintain rigorous accounting continuity without lookahead or truncation bias, the opening asset balance for the earliest fiscal year (FY 2016) is retrieved directly from comparative 10-K disclosures for FY 2015 ($290,345.0M). For all subsequent periods $t$, $\text{Beginning Assets}_t = \text{Ending Assets}_{t-1}$.

4. **Profitability — Net Profit Margin**:
   $$\text{Net Profit Margin} = \left(\frac{\text{Net Income}}{\text{Net Revenue}}\right) \times 100$$
   *Purpose*: Measures the percentage of top-line revenue converted into bottom-line profit after accounting for all operating expenses, taxes, and net financing costs.

### Ten-Year Financial Ratio Results (Apple Inc.)

| Fiscal Year | Current Ratio | Debt-to-Equity | Asset Turnover | Net Profit Margin | Status |
| :---: | :---: | :---: | :---: | :---: | :---: |
| **FY 2016** | **1.35x** | **0.68x** | **0.70x** | **21.19%** | Verified |
| **FY 2017** | **1.28x** | **0.86x** | **0.66x** | **21.09%** | Verified |
| **FY 2018** | **1.13x** | **1.07x** | **0.72x** | **22.41%** | Verified |
| **FY 2019** | **1.54x** | **1.19x** | **0.74x** | **21.24%** | Verified |
| **FY 2020** | **1.36x** | **1.72x** | **0.83x** | **20.91%** | Verified |
| **FY 2021** | **1.07x** | **1.98x** | **1.08x** | **25.88%** | Verified |
| **FY 2022** | **0.88x** | **2.37x** | **1.12x** | **25.31%** | Verified |
| **FY 2023** | **0.99x** | **1.79x** | **1.09x** | **25.31%** | Verified |
| **FY 2024** | **0.87x** | **1.87x** | **1.09x** | **23.97%** | Verified |
| **FY 2025** | **0.89x** | **1.34x** | **1.15x** | **26.92%** | Verified |

### Economic Insights & Long-Term Trends (Apple Inc.)
- **Liquidity Compression (Current Ratio 1.35x $\rightarrow$ 0.89x)**: Apple transitioned toward a negative cash conversion cycle, operating efficiently with current ratios below 1.0x in recent years due to massive operating cash flows and supplier credit terms.
- **Capital Structure Optimization (Debt-to-Equity 0.68x $\rightarrow$ 2.37x $\rightarrow$ 1.34x)**: Apple took advantage of historically low interest rates to issue institutional debt while returning hundreds of billions to shareholders via stock buybacks, reducing the equity denominator and peaking leverage in FY 2022 before deleveraging toward 1.34x in FY 2025.
- **Asset Efficiency Expansion (Asset Turnover 0.70x $\rightarrow$ 1.15x)**: Driven by the high-margin, asset-light expansion of Services (App Store, iCloud, Apple Pay) alongside iPhone hardware scale.
- **Profitability Elevation (Net Margin 21.19% $\rightarrow$ 26.92%)**: Net margin expanded by over 570 basis points over the decade, reflecting pricing power, higher-margin services mix, and silicon self-sufficiency (Apple Silicon M-series).

### Excel Worksheet & Charting Specifications (`Financial Ratios` Tab)
- **Sheet Position**: Inserted as Tab 6, positioned cleanly before `Audit & Reconciliation`.
- **Live Formula References**: All ratio cells in the summary table (Rows 7–10) contain live, native Excel formulas (e.g. `=B14/B15`, `=B21/B22`, `=B28/B27`, `=B31/B32`) linking directly to the supporting calculation details below.
- **Calculation Details Table (Rows 12–32)**: Displays complete underlying accounting accounts ($M) for every reporting period, including beginning, ending, and average assets, and itemized debt components.
- **Four Embedded Trend Charts (Rows 34–62)**: Arranged in a 2x2 grid:
  - Chart 1: *Current Ratio Trend (10-Year Liquidity)*
  - Chart 2: *Debt-to-Equity Ratio Trend (10-Year Solvency)*
  - Chart 3: *Total Asset Turnover Trend (10-Year Efficiency)*
  - Chart 4: *Net Profit Margin Trend (10-Year Profitability)*
- **Formatting Standards**: Freeze Panes at cell `B7`, Navy `#1F4E78` banner, Soft Blue `#D9E1F2` section fills, `0.00"x"` ratio formatting, `0.0%` margin formatting, and `$#,##0.0;($#,##0.0);"-"` currency formatting.

---

## Stage 9: Final Excel Analysis, Visualization & Complete Application Integration

### Purpose & Architecture Overview
Stage 9 unifies all preceding modules into a single, fully integrated command-line application orchestrated by `main.py` and supported by the dedicated reporting engine `excel_report.py`. The application coordinates the end-to-end accounting pipeline across 8 structured processing stages:
1. `[1/8] Downloading SEC filings...`: SEC EDGAR primary Form 10-K retrieval with caching (`sec_downloader.py`).
2. `[2/8] Extracting MD&A sections...`: Item 7 narrative extraction and validation (`mda_extractor.py`).
3. `[3/8] Extracting financial statements...`: Official XBRL 10-year financials extraction (`financial_extractor.py`).
4. `[4/8] Cleaning MD&A text...`: Header/footer removal, typography normalization, uppercase tokenization (`text_cleaner.py`).
5. `[5/8] Performing sentiment analysis...`: Loughran-McDonald Master Dictionary scoring and P/N ratio calculation (`sentiment_analysis.py`).
6. `[6/8] Retrieving historical stock prices...`: Split-adjusted and dividend-adjusted closing prices matched to balance sheet dates (`stock_prices.py`).
7. `[7/8] Calculating financial ratios...`: Liquidity, Solvency, Efficiency, and Profitability ratios with traceable Excel formulas (`financial_ratios.py`).
8. `[8/8] Generating Excel workbook...`: Creation of the final comprehensive workbook `{TICKER}_10K_Analysis.xlsx` with dual-axis charting and longitudinal analysis (`excel_report.py`).

### Final Comprehensive Excel Workbook Structure (`{TICKER}_10K_Analysis.xlsx`)
Saved in both `data/excel_outputs/` and `data/excel_outputs/{TICKER}/`:
- **Tab 1: `Income Statement`**: 10-year horizontal Consolidated Statements of Operations ($M).
- **Tab 2: `Balance Sheet`**: 10-year horizontal Consolidated Balance Sheets ($M) with accounting reconciliation.
- **Tab 3: `Cash Flow Statement`**: 10-year horizontal Consolidated Statements of Cash Flows ($M).
- **Tab 4: `Sentiment Analysis`**: Complete Loughran-McDonald sentiment metrics (Positive, Negative, Uncertainty, Litigious, Constraining, Modal terms, and P/N Ratios).
- **Tab 5: `Stock Prices`**: Split-adjusted closing prices, total return dividend-adjusted prices, volume, and trading day offset tracking.
- **Tab 6: `Financial Ratios`**: Current Ratio, Debt-to-Equity, Total Asset Turnover, Net Profit Margin, underlying account calculation details, and 4 ratio trend charts.
- **Tab 7: `Charts and Analysis`**:
  - *Summary Data Table (Rows 4–14)*: Anchors all underlying longitudinal metrics for clean chart references.
  - *Mandatory Dual-Axis Line Chart (Rows 16–32)*: Historical Stock Price ($ USD on Primary Y-Axis) vs. Loughran-McDonald Sentiment P/N Ratio (on Secondary Y-Axis) across all 10 fiscal years.
  - *Financial Ratio Trend Charts (Rows 34–65)*: 2x2 grid displaying Current Ratio, Debt-to-Equity, Asset Turnover, and Net Profit Margin.
  - *Disclosure Volume Trend Chart (Rows 67–85)*: Positive vs. Negative word counts illustrating disclosure growth over time.
  - *Executive Analytical Commentary (Rows 86–100)*: Structured accounting and economic insights analyzing trends.
- **Tab 8: `Audit & Reconciliation`**: Automated balance sheet equation and gross profit verification.

### Key Longitudinal Findings (Apple Inc. — FY 2016 through FY 2025)
1. **Stock Price vs. Sentiment P/N Ratio Decoupling**:
   - In FY 2021, Apple's sentiment P/N ratio reached a peak of **0.929** (post-pandemic revenue surged +33% with record profitability).
   - Between FY 2022 and FY 2025, the P/N ratio compressed to **0.086** (~90% decline in sentiment optimism).
   - Simultaneously, Apple's split-adjusted stock price rose from **$150.43 to $255.46 (+70%)**.
   - *Driver*: This decoupling reflects increased regulatory disclosure requirements in Item 7 (EU Digital Markets Act compliance, antitrust litigation disclosures, App Store lawsuits, and geopolitical supply chain risks) rather than operational decline. Negative disclosure word counts rose from 377 in FY 2016 to 927 in FY 2025.
2. **Capital Return & Leverage Expansion**:
   - Debt-to-Equity increased from 0.68x in FY 2016 to 2.37x in FY 2022 due to Apple's aggressive share buyback program ($500B+ returned), which deliberately reduced book equity while boosting EPS.
3. **Working Capital Strategy**:
   - Current Ratio shifted from 1.35x down to 0.89x, demonstrating negative cash conversion cycles and strong supplier financing leverage.

### Key Longitudinal Findings (NVIDIA Corporation — FY 2017 through FY 2026)
1. **Accelerated Computing & Generative AI Hyper-Growth**:
   - Total Net Revenue surged from **$6,910.0M in FY 2017 to $215,938.0M in FY 2026 (+3,025%)**, driven by global data center GPU and accelerated AI infrastructure demand (Hopper and Blackwell architectures).
   - Net Income expanded exponentially from **$1,666.0M in FY 2017 to $120,067.0M in FY 2026 (+7,107%)**, lifting Net Profit Margin from 24.11% to **55.60%**.
2. **Extreme Solvency De-leveraging**:
   - Solvency Debt-to-Equity ratio plunged from **0.48x in FY 2017 down to 0.05x in FY 2026**.
   - Total interest-bearing debt remained modest at $8,468.0M while Shareholders' Equity ballooned to **$157,293.0M**, leaving NVIDIA virtually unencumbered by leverage and maintaining immense balance sheet strength.
3. **Working Capital and Efficiency**:
   - Total Current Assets reached $125,605.0M in FY 2026 (including $10,605M in cash and $51,951M in marketable securities) against $32,163.0M of current liabilities, yielding a robust Current Ratio of **3.91x**.
   - Total Asset Turnover rose from 0.80x in FY 2017 to **1.36x in FY 2026**, underscoring extreme capital efficiency on average assets of $159,202.0M.
4. **Stock Price Surge vs. Sentiment P/N Ratio Dynamics**:
   - Split-adjusted closing stock price climbed from **$2.79 on 2017-01-27 to $187.67 on 2026-01-23 (+6,626%)**, accounting for the cumulative 40-for-1 splits (4:1 in 2021, 10:1 in 2024).
   - Sentiment P/N ratio moderated from **0.283 in FY 2017 to 0.165 in FY 2026**, as expanded narrative Item 7 disclosures detailed export control restrictions, geopolitical risks, and supply constraints alongside record operational execution.

---

## Standalone Windows Desktop Application (`SEC10KAnalyzer.exe`)

The application includes a standalone, dark-themed Windows desktop interface built with **CustomTkinter** and compiled into a native Windows executable (`SEC10KAnalyzer.exe`) via **PyInstaller**. It allows users with zero Python experience to launch and operate the complete financial analytics system by double-clicking a desktop icon, similar to Discord or Spotify.

### Key Desktop Interface Features
1. **Modern Dark-Themed GUI**: Built with CustomTkinter (`gui_app.py`) featuring high-DPI scaling, customized Segoe UI typography, and responsive grid layouts.
2. **Dual Company Selection**:
   - Direct ticker input field (e.g. `AAPL`, `NVDA`, `MSFT`) with instant SEC registry verification.
   - Interactive company-name search engine querying 10,430+ registered corporations in official SEC EDGAR data with ranked candidate selection.
3. **Flexible Analysis Scope**: Dropdown selector supporting `3`, `5`, `10` (standard), `15`, or `20` annual 10-K filings.
4. **Non-Blocking Background Worker Threading**: Analysis pipeline executes on a background daemon worker thread (`threading.Thread`), keeping the graphical user interface completely responsive with zero freezing or window locking.
5. **Real-Time Progress & Stage Tracking**:
   - Animated progress bar tracking completion from 0% to 100%.
   - Active status panel detailing current pipeline operations across all 8 stages.
6. **Executive Results Summary**: Displays primary valuation and financial metrics upon completion (Latest Stock Price, Sentiment P/N Ratio, Current Ratio, Net Profit Margin, and Balance Sheet / Gross Profit reconciliation audit status).
7. **Direct System Launchers**:
   - `📊 Open Excel Workbook`: Opens the finalized comprehensive Excel report directly in Microsoft Excel via native Windows file association (`os.startfile`).
   - `📁 Open Output Folder`: Opens the output directory in Windows File Explorer.
8. **Live Real-Time Activity Log**: Monospace console box streaming real-time timestamped logging events directly inside the application.
9. **SEC User-Agent Settings Dialog**: Easily configure and validate user contact headers without exposing personal credentials in public source files.
10. **Multi-Company Continuous Workflow**: Users can analyze Apple, switch to NVIDIA or another corporation, and re-run immediately without restarting the application.

### Persistent Desktop Storage Architecture
When running as a packaged Windows executable (either one-folder or portable one-file):
- **Application Bundle (`_internal/` or `sys._MEIPASS`)**: Contains Python runtime, compiled C-extensions, dependencies, Loughran-McDonald dictionary, SEC tickers registry, and pre-packaged baseline workbooks.
- **Writable User-Data Directory**: Automatically defaults to `%LOCALAPPDATA%\SEC10KAnalyzer\data\` (outside of PyInstaller's ephemeral extraction path).
- **Asset Seeding**: On first launch, the application automatically seeds bundled dictionary files and baseline company files to the local user-data directory, ensuring permanent persistence and offline access across reboots.

### Packaging Distributions Available
1. **One-Folder Distribution (`dist\SEC10KAnalyzer\`)**:
   - Location: `dist\SEC10KAnalyzer\SEC10KAnalyzer.exe`
   - Fast startup (no decompression overhead). Ideal for permanent local installations.
2. **Portable Single-File Executable (`dist\portable\SEC10KAnalyzer.exe` & `dist\SEC10KAnalyzer_portable.exe`)**:
   - Location: `dist\portable\SEC10KAnalyzer.exe` (~54.6 MB)
   - Self-contained, portable binary. Can be moved to any folder, USB flash drive, or desktop without external dependencies.
   - Automatically stores all generated filings and reports in `%LOCALAPPDATA%\SEC10KAnalyzer\data\`.

### Launching the Desktop Application
- **From Windows Desktop**: Double-click `SEC 10-K Financial Analyzer` shortcut on your Windows Desktop.
- **From Portable Executable**: Double-click `dist\portable\SEC10KAnalyzer.exe` or `dist\SEC10KAnalyzer_portable.exe`.
- **From Distribution Folder**: Double-click `dist\SEC10KAnalyzer\SEC10KAnalyzer.exe`.
- **Create Desktop Shortcut Script**:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\create_desktop_shortcut.ps1
  ```

---

## Execution Commands (Windows PowerShell)

### 1. Launch Primary Streamlit Web Dashboard:
```powershell
# Using project virtual environment
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py

# Using standard active Python environment
python -m streamlit run streamlit_app.py
```

### 2. Run Desktop GUI Application (Python Development Mode):
```powershell
.\.venv\Scripts\python.exe gui_app.py
# Or via main.py flag
.\.venv\Scripts\python.exe main.py --gui
```

### 3. Run Standalone Packaged Executables:
```powershell
# Launch Portable Single-File Executable:
.\dist\portable\SEC10KAnalyzer.exe

# Launch One-Folder Distribution Executable:
.\dist\SEC10KAnalyzer\SEC10KAnalyzer.exe
```

### 4. Run Full End-to-End Pipeline (Direct CLI Command):
```powershell
# Run Apple Inc. by ticker
.\.venv\Scripts\python.exe main.py --ticker AAPL --count 10

# Run NVIDIA Corporation by ticker
.\.venv\Scripts\python.exe main.py --ticker NVDA --count 10

# Run by company name search (non-interactive CLI)
.\.venv\Scripts\python.exe main.py --search "Apple" --count 10
.\.venv\Scripts\python.exe main.py --search "NVIDIA" --count 10
.\.venv\Scripts\python.exe main.py --search "Microsoft" --count 10
```

### 5. Run Standalone Stages (For Testing or Verification):
```powershell
# Stage 2: SEC EDGAR 10-K Retrieval (by ticker or company search)
.\.venv\Scripts\python.exe main.py --ticker NVDA --download-only
.\.venv\Scripts\python.exe main.py --search "Apple" --download-only

# Stage 3: Item 7 MD&A Text Extraction
.\.venv\Scripts\python.exe main.py --ticker NVDA --extract-mda

# Stage 4: Financial Statement Extraction & Excel Integration
.\.venv\Scripts\python.exe main.py --ticker NVDA --extract-financials

# Stage 5: MD&A Text Cleaning & Validation
.\.venv\Scripts\python.exe main.py --ticker NVDA --clean-mda

# Stage 6: Loughran-McDonald Sentiment Analysis & Excel Integration
.\.venv\Scripts\python.exe main.py --ticker NVDA --sentiment

# Stage 7: Historical Stock Price Extraction & Excel Integration
.\.venv\Scripts\python.exe main.py --ticker NVDA --stock-prices

# Stage 8: Ten-Year Financial Ratio Analysis & Visualizations
.\.venv\Scripts\python.exe main.py --ticker NVDA --ratios

# Stage 9: Final Excel Reporting & Dual-Axis Charting
.\.venv\Scripts\python.exe main.py --ticker NVDA --excel-report
```

### 6. Run All Automated Unit & Streamlit Test Suites (69 Total Tests):
```powershell
.\.venv\Scripts\python.exe -m unittest discover -v -p "test_*.py"
```

---

## Streamlit Cloud & GitHub Deployment Preparation

This application is engineered for zero-modification deployment to **Streamlit Community Cloud** and cross-platform execution (Windows and Linux):

### 1. SEC EDGAR User-Agent Configuration in Cloud
When deploying to Streamlit Community Cloud, configure your institutional User-Agent header in **Streamlit Secrets** (`App Settings -> Secrets` in the cloud dashboard):
```toml
# .streamlit/secrets.toml (Ignored by Git to protect personal contact info)
SEC_USER_AGENT = "Student Name student@university.edu"
```
The application automatically inspects `st.secrets["SEC_USER_AGENT"]`, then fallback environment variables (`os.environ["SEC_USER_AGENT"]`), and finally local `.sec_settings`.

### 2. Portable Path Architecture
All file paths use standard Python `pathlib.Path` resolved relative to the project root directory (`config.BASE_DIR`). No hardcoded Windows drive letters (`C:\`) or backslashes are used in file resolution.

### 3. Loughran-McDonald Master Dictionary
The official Loughran-McDonald Master Dictionary (CSV) is maintained in `data/` and loaded automatically via cross-platform pattern matching.

### 4. Git Cleanliness & Security
The `.gitignore` configuration explicitly excludes:
- Local Python virtual environments (`.venv/`, `env/`)
- Streamlit secrets (`.streamlit/secrets.toml`, `secrets.toml`)
- Persisted personal user-agent files (`.sec_settings`)
- Windows binary builds (`build/`, `dist/`, `*.exe`, `*.zip`)
- Compiled Python caches (`__pycache__/`, `*.pyc`)

---

## Known Limitations & Observations
- **Sub-segment Line Items**: Specific non-operating items (e.g. minor foreign exchange hedges or legal contingency accruals) that are not separately tagged in XBRL are grouped under *Other Operating / Non-Operating Adjustments*.
- **Cash Flow Reconciliations**: In certain years, minor discrepancies between the sum of operating, investing, and financing cash flows and reported net cash changes arise from foreign currency translation adjustments on cash balances (*Effect of Exchange Rate Changes*), which is reported on its own line item in the workbook.
- **P/N Ratio Magnitudes**: Due to SEC Item 7 regulatory requirements mandating thorough disclosure of potential business risks, competitive threats, and uncertainties, negative financial terms naturally outnumber positive terms across standard 10-Ks. P/N ratios for AAPL range from 0.062 (FY2023) to 0.929 (FY2021).
- **Numerical Tokens in Sentiment Analysis**: Numerical values are maintained in the archived cleaned text for human reading and qualitative review, but are filtered out of the uppercase token stream since standard Loughran-McDonald sentiment dictionaries do not categorize numbers.
- **Top-Line Terminology Variations**: In modern MD&A disclosures (FY2021–FY2025), Apple Inc. references "Net Sales" or "Sales" rather than "Revenue" in narrative text. The text cleaning validation engine accounts for this convention by accepting both "SALES" and "REVENUE".
- **Operating Leases vs. Interest-Bearing Debt**: Under ASC 842 (adopted FY2019), operating lease liabilities are recorded on corporate balance sheets. In accordance with standard accounting solvency conventions, interest-bearing debt strictly isolates contractual financial debt (Commercial Paper, Current Term Debt, Long-Term Debt) from operating liabilities.
- **Negative Working Capital in Tech Scale**: Apple's current ratio dipped below 1.0x in recent fiscal years (FY2022–FY2025). This reflects disciplined working capital management, negative cash conversion cycles, and high supplier credit leverage rather than liquidity impairment.
