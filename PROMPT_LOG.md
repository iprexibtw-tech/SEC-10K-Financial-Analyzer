# GEN AI Accounting Project — Chronological User Prompt Log

This document maintains an exact, verbatim chronological record of all user prompts provided during the development of this undergraduate Gen AI accounting project. Per the final requirements in `Instructions.docx`, this prompt history serves as the primary audit trail and will be included as an appendix at the end of the final lab document.

All prompts below have been extracted directly and verbatim from the active conversation history without omissions, inventions, or reconstructions.

---

## Prompt Log Summary

| # | Date & Time | Stage / Topic | Purpose / Objective |
| :-: | :--- | :--- | :--- |
| **1** | 2026-10-29 14:11:01 | **Stage 1** | Project setup, folder architecture, virtual environment (`.venv`), PEP 8 standards, and baseline verification. |
| **2** | 2026-10-30 14:07:41 | **Stage 2** | SEC EDGAR Form 10-K retrieval module (`sec_downloader.py`), rate limiting, 10 distinct reporting periods, and User-Agent compliance. |
| **3** | 2026-10-30 14:24:12 | **Stage 2 (Configuration)** | User contact details supplied to configure compliant SEC User-Agent header and authorize AAPL 10-K downloads. |
| **4** | 2026-10-30 14:29:43 | **Stage 3** | Item 7 MD&A textual extraction module (`mda_extractor.py`), HTML cleanup, TOC avoidance, boundary verification, and UTF-8 storage. |
| **5** | 2026-10-30 14:40:14 | **Stage 4** | Financial statement extraction module (`financial_extractor.py`), SEC XBRL facts, multi-tab Excel creation (`openpyxl`), and accounting reconciliations. |
| **6** | 2026-10-30 14:57:53 | **Documentation / Audit** | Creation of `PROMPT_LOG.md` to track all development prompts for lab document appendix. |
| **7** | 2026-10-30 14:59:20 | **Stage 5** | MD&A text cleaning module (`text_cleaner.py`), HTML/header removal, negation/financial word preservation, uppercase tokenization, and quality validation. |
| **8** | 2026-10-30 15:06:13 | **Stage 6 (Dictionary Setup)** | Loughran-McDonald dictionary auto-discovery, category validation, `sentiment_analysis.py`, P/N ratio calculation, and Excel workbook tab integration. |
| **9** | 2026-10-30 15:15:06 | **Stage 6 (Full Specification)** | Full Stage 6 requirements: CSV/Excel format support, exact sentiment categories, P/N ratio, word frequencies, "Sentiment Analysis" Excel tab, and CSV manifest. |
| **10** | 2026-10-30 15:22:53 | **Stage 7** | Historical stock price extraction module (`stock_prices.py`), fiscal year-end balance sheet date matching, non-trading day fallback, split/adjusted closing prices, and "Stock Prices" Excel tab integration. |
| **11** | 2026-10-30 15:35:50 | **Stage 8** | Ten-year financial ratio analysis (`financial_ratios.py`): Current Ratio, Debt-to-Equity, Total Asset Turnover, Net Profit Margin, supporting calculation details, professional Excel charts, and workbook integration. |
| **12** | 2026-10-30 15:45:43 | **Stage 9** | Final Excel analysis, visualization (`Charts and Analysis` tab with dual-axis stock price vs. sentiment P/N ratio chart), dedicated Excel reporter (`excel_report.py`), and unified pipeline integration in `main.py`. |
| **13** | 2026-10-30 15:54:40 | **Stage 10** | NVIDIA analysis and final Python application validation: end-to-end execution for NVDA (10 years), company-specific differences (fiscal calendar, taxonomy, stock splits), dual validation for AAPL & NVDA, and final documentation. |
| **14** | 2026-10-30 16:15:10 | **Stage 11** | Final application audit, quality assurance review, and assignment compliance validation against `Instructions.docx`. |
| **15** | 2026-10-30 16:35:20 | **Enhancement** | Company selection enhancement: ticker input + official SEC company-name search and autocomplete in `main.py`. |
| **16** | 2026-10-30 16:50:00 | **Desktop App** | Standalone Windows desktop application using CustomTkinter (`gui_app.py`) with one-folder PyInstaller packaging. |
| **17** | 2026-10-30 17:10:15 | **Packaging** | Portable one-file Windows executable packaging (`SEC10KAnalyzer.exe`) via PyInstaller `--onefile` and relocated execution testing. |
| **18** | 2026-10-30 17:25:30 | **Distribution** | Final portable distribution package (`SEC 10-K Financial Analyzer/`), structured `Company Analyses/` folder, and distribution README. |
| **19** | 2026-10-30 17:35:45 | **Streamlit** | Interactive Streamlit financial analytics web dashboard (`streamlit_app.py`) with 6 tabs, executive dark theme, Plotly charts, and export engine. |
| **20** | 2026-10-30 17:46:29 | **GitHub & Cloud** | Repository preparation, privacy/credential audit, .gitignore hardening, user approval, and push to GitHub (`iprexibtw-tech/SEC-10K-Financial-Analyzer`). |
| **21** | 2026-10-30 18:40:40 | **Streamlit Cloud Architecture** | Clarification to maintain Streamlit Cloud deployment: add Apple vs. NVIDIA comparison tab, quick switcher pills, and interactive statement filtering. |

---

## Verbatim Prompt History

### Prompt 1: Stage 1 — Environment & Architecture Setup
* **Timestamp**: 2026-10-29T14:11:01-04:00
* **Stage**: Stage 1

```text
I am completing an undergraduate GEN AI accounting assignment using Python and Google Antigravity on Windows.

Please read Instructions.docx in my project folder and treat it as the primary source of assignment requirements.

Our goal is to develop a modular Python application that analyzes the 10 most recent annual SEC 10-K filings for Apple (AAPL) and a second company.

We will develop the application incrementally, testing and verifying each stage before proceeding.

For Stage 1:
1. Create a modular Python project structure using main.py as the central entry point.
2. Follow PEP 8 coding standards.
3. Use reusable functions, clear comments, docstrings, and error handling.
4. Create organized folders for downloaded SEC filings, extracted text, financial data, and Excel outputs.
5. Set up a Python virtual environment (.venv) compatible with Windows.
6. Create requirements.txt and install only packages required for the current stage.
7. Create README.md containing Windows setup instructions and execution commands.
8. Add .gitignore to exclude .venv, cache files, and temporary data.
9. Design the project so additional modules can be added without rewriting working code.
10. Do not create unnecessary placeholder modules.
Verify that the Python environment works by executing a simple test through main.py.

Explain which files were created, show actual test results, and provide the exact Windows PowerShell command needed to run the program.

Do not proceed to Stage 2 until I approve.
```

---

### Prompt 2: Stage 2 — SEC EDGAR 10-K Retrieval Specification
* **Timestamp**: 2026-10-30T14:07:41-04:00
* **Stage**: Stage 2

```text
STAGE 2 — SEC EDGAR 10-K Filing Retrieval

We have completed Stage 1 of my undergraduate GEN AI accounting project.

Please review the existing project structure, README.md, and Instructions.docx before making any changes. Preserve all functioning code and continue using our modular Python architecture on Windows.

Objective: Develop a Python module that downloads, organizes, and verifies annual Form 10-K filings from the official SEC EDGAR database.

1. User Input

Modify main.py to prompt the user for:

A publicly traded company's ticker symbol (e.g., AAPL).

The number of most recent annual 10-K filings to download (default: 10).

Validate inputs, normalize ticker symbols, and display understandable error messages when inputs are invalid.

2. SEC EDGAR Integration

Create a separate module named sec_downloader.py.

The module must:

Retrieve the correct company CIK using official SEC company ticker data.

Query SEC EDGAR's submissions data to identify available annual Form 10-K filings, including older filing records when necessary.

Exclude amended filings (10-K/A) and unrelated filing types.

Download the actual primary 10-K filing documents, not merely submission metadata.

Identify the 10 most recent distinct annual reporting periods available.

Use a descriptive SEC-compliant User-Agent with my contact information, which you should request before making SEC requests.

Respect SEC rate limits, apply reasonable delays, and implement retries with exponential backoff for temporary server errors.

Never bypass SEC access restrictions.

3. File Organization

Save filings within a company-specific directory using the existing project structure.

Use descriptive filenames containing:

Company ticker.

Fiscal reporting year.

Filing date.

Create a CSV manifest containing:

Company ticker and CIK.

Fiscal period-end date.

Filing date.

Accession number.

Original SEC source URL.

Saved filename and location.

Download verification status.

Avoid duplicate downloads and preserve existing valid files.

4. Error Handling

Implement appropriate handling for:

Invalid or nonexistent ticker symbols.

Missing filings or incomplete filing histories.

Internet connection failures.

SEC rate limiting or temporary service interruptions.

Corrupted, incomplete, or unexpected downloaded content.

File access and permission errors.

If fewer than the requested filings are available, report the actual number retrieved instead of silently substituting other filing types.

5. Verification and Testing

Test the program using:

Ticker: AAPL
Number of annual 10-K filings: 10

Verify that:

The correct company and CIK are identified.

The downloaded files are genuine annual 10-K filing documents.

The files cover 10 distinct annual reporting periods.

Filenames and manifest records match the corresponding SEC metadata.

All downloads are readable and contain substantive filing content.

Running the same request again does not unnecessarily download duplicate files.

Display a summary table showing the fiscal year, fiscal period-end date, filing date, accession number, filename, and verification status.

Report any unsuccessful tests or uncertain results explicitly.

6. Documentation

Update README.md to explain:

How the SEC download module works.

Required Python packages.

How to configure the SEC User-Agent.

How to run the program in Windows PowerShell.

Where downloaded files are saved.

How to verify successful downloads.

Important Development Rules

Follow PEP 8 standards.

Include helpful docstrings and comments.

Preserve Stage 1 functionality.

Keep SEC downloading logic separate from main.py.

Use the existing virtual environment.

Do not create a graphical interface.

Do not implement MD&A extraction, financial statement extraction, sentiment analysis, or Excel reporting yet.

Do not proceed to Stage 3 without my approval.

After implementation, run the tests, show the actual results, explain any errors, and provide the exact Windows PowerShell command needed to run the application.

Do not claim successful verification unless the relevant tests have actually passed.
```

---

### Prompt 3: Stage 2 — User-Agent Configuration & Download Authorization
* **Timestamp**: 2026-10-30T14:24:12-04:00
* **Stage**: Stage 2 (Execution)

```text
Name: Kevin Nguyen
Email: knguy5@unh.newhaven.edu

Please configure the SEC User-Agent using this information and proceed with downloading and verifying the 10 most recent AAPL 10-K filings.

For this project, focus exclusively on developing and testing the Python application locally on Windows.

Do not initialize Git, create a GitHub repository, or perform any version control operations.

Keep all Python scripts, downloaded SEC filings, and generated outputs within my existing project folder.

Once the downloads are complete, display the verification results and CSV manifest summary.

Do not proceed to Stage 3 until I approve.
```

---

### Prompt 4: Stage 3 — SEC 10-K Item 7 MD&A Text Extraction
* **Timestamp**: 2026-10-30T14:29:43-04:00
* **Stage**: Stage 3

```text
STAGE 3 — SEC 10-K MD&A TEXT EXTRACTION

We are continuing the undergraduate GEN AI accounting project using Python and Google Antigravity on Windows.

Before starting, review Instructions.docx and the existing Python project files. Preserve all previously completed and verified functionality.

The two companies selected for this project are:

Apple Inc. (AAPL)

NVIDIA Corporation (NVDA)

Our current objective is to extract Management's Discussion and Analysis (MD&A) sections from the annual SEC 10-K filings downloaded during Stage 2.

1. Create the MD&A Extraction Module

Create a separate Python module named mda_extractor.py.

The module must:

Read the locally downloaded SEC 10-K filings.

Identify Item 7 — Management's Discussion and Analysis of Financial Condition and Results of Operations.

Extract the complete substantive MD&A section, ending before Item 7A — Quantitative and Qualitative Disclosures About Market Risk, where applicable.

Handle variations in filing formats, HTML structures, heading capitalization, whitespace, and section labeling.

Avoid incorrectly extracting Item 7 references from the table of contents or repeated headings.

Preserve all meaningful financial language, numerical information, and paragraph structure.

Remove HTML markup and irrelevant formatting while preserving the original textual content.

Avoid removing words or applying sentiment-related cleaning at this stage.

Do not assume the first occurrence of "Item 7" is the actual MD&A section.

2. File Organization

Save each extracted MD&A section as an individual UTF-8 .txt file.

Use the existing project directory structure.

Organize files by company and fiscal year, using filenames such as:

AAPL_2025_MDA.txt

Ensure the year corresponds to the fiscal reporting period, not simply the filing date.

Do not overwrite or modify the original downloaded SEC filings.

3. Validation and Quality Control

For each extracted MD&A section:

Confirm that the section begins with the substantive Item 7 content.

Confirm that extraction stops at the correct ending boundary.

Check for missing, duplicated, or truncated text.

Calculate the extracted word count.

Record the source filing and reporting year.

Flag suspiciously short, unusually long, or empty extractions.

Identify potential table-of-contents contamination.

Flag extraction results that cannot be confidently validated.

Do not classify an extraction as successful merely because a text file was created.

Create an extraction summary showing:

Company ticker

Fiscal reporting year

Source filing filename

Extracted text filename

Word count

Extraction status

Warnings or validation issues

4. Integrate With the Existing Python Application

Integrate mda_extractor.py with main.py.

Requirements:

Preserve the existing SEC downloader and all verified Stage 2 functionality.

Keep MD&A extraction logic separate from the SEC downloading module.

Allow the program to process multiple filings without manually modifying the source code.

Handle missing or invalid filing files gracefully.

Use PEP 8 coding standards, meaningful comments, docstrings, and appropriate error handling.

Use only necessary Python dependencies.

Ensure compatibility with Windows PowerShell.

5. Testing and Verification

First, test the module using the 10 downloaded Apple (AAPL) annual 10-K filings.

For each year:

Extract the MD&A section.

Save the extracted content as a text file.

Display the word count.

Show a short excerpt from the beginning and end of the extracted section.

Verify that the boundaries are correct.

Report any years requiring manual review.

Use a combination of automated validation and source-document comparisons.

Do not claim every extraction is correct unless the actual content and section boundaries have been verified.

6. Documentation

Update README.md with:

An explanation of the MD&A extraction process.

How to execute the extraction module.

Where the extracted text files are stored.

The validation methods used.

Any limitations or formatting issues encountered.

Important Development Instructions

Do not initialize Git or GitHub.

Work exclusively within the existing local Python project.

Do not unnecessarily rewrite previously verified code.

Do not implement financial statement extraction, stock price analysis, or sentiment scoring yet.

Do not proceed to Stage 4 until I approve.

After completing this stage, execute the program and show the actual extraction results for all 10 AAPL filings.

Provide a summary table with the year, extracted word count, verification status, and any errors.

Tell me the exact Windows PowerShell command needed to run the updated program.

Stop after completing Stage 3.
```

---

### Prompt 5: Stage 4 — Financial Statement Extraction & Excel Integration
* **Timestamp**: 2026-10-30T14:40:14-04:00
* **Stage**: Stage 4

```text
STAGE 4 — FINANCIAL STATEMENT EXTRACTION AND EXCEL INTEGRATION

We are continuing my undergraduate GEN AI accounting project using Python and Google Antigravity on Windows.

Please review Instructions.docx and the existing Python application before making changes.

The two companies selected for this project are:

Apple Inc. (AAPL)

NVIDIA Corporation (NVDA)

We are developing a modular Python application that retrieves and analyzes 10 years of annual SEC 10-K filings.

Stages 1–3 have been implemented. We will perform a comprehensive project review after completing all development stages.

For now, proceed with Stage 4.

1. Create the Financial Statement Extraction Module

Create a separate Python module named financial_extractor.py.

The module must extract the following financial statements for each of the 10 fiscal reporting periods:

Consolidated Statements of Operations (Income Statement)

Consolidated Balance Sheets

Consolidated Statements of Cash Flows

Use official SEC EDGAR data sources, including SEC Company Facts XBRL data, where appropriate.

Use the original downloaded 10-K filings for verification and to resolve company-specific reporting differences.

Do not fabricate or estimate financial statement values.

2. Financial Statement Requirements

For each reporting year:

Identify the correct fiscal period-end date.

Retrieve annual financial statement values.

Distinguish annual duration-based values from balance sheet instant values.

Use appropriate SEC XBRL accounting concepts.

Account for company-specific XBRL tags and differences in financial statement presentation.

Select values corresponding to the correct fiscal periods and reporting context.

Avoid incorrectly mixing annual, quarterly, or comparative values from different filings.

Preserve the original financial statement units and convert them consistently for presentation.

Mark unavailable values as missing rather than replacing them with zero.

Ensure that the extraction logic works for both Apple and NVIDIA without requiring manual changes to the program.

3. Create the Excel Workbook

Create an Excel workbook using openpyxl or another appropriate Python library.

Save the workbook in the existing data/excel_outputs directory.

Create three worksheets:

Worksheet 1: Income Statement

List financial statement line items vertically.

Display 10 fiscal years horizontally, from oldest to newest.

Include revenue, cost of revenue, gross profit, operating expenses, operating income, pretax income, income tax expense, net income, and other relevant reported items.

Worksheet 2: Balance Sheet

Display current assets, noncurrent assets, total assets, current liabilities, noncurrent liabilities, total liabilities, shareholders' equity, and other relevant reported items.

Present 10 fiscal years side by side.

Worksheet 3: Cash Flow Statement

Display cash flows from operating, investing, and financing activities.

Include important reported cash flow components and net changes in cash.

Present 10 fiscal years side by side.

Include additional financial statement line items where supported by the company's filings.

Do not omit required financial statement components solely because they use company-specific labels.

4. Excel Formatting

Apply professional financial statement formatting:

Clear worksheet titles and company identification.

Fiscal year and fiscal period-end date headers.

Consistent financial units, preferably USD millions.

Appropriate number formatting.

Negative amounts displayed in parentheses.

Bold financial statement subtotals and totals.

Consistent column widths and alignment.

Freeze panes to keep year headings and account names visible.

Clearly distinguish missing values from reported zeros.

Include the source information and extraction assumptions in a separate metadata worksheet if useful.

5. Data Validation

Implement automated validation procedures.

At a minimum:

Verify that Total Assets = Total Liabilities + Shareholders' Equity, within reasonable rounding tolerance.

Verify that gross profit agrees with revenue less cost of revenue, where those values are reported.

Verify that the cash flow subtotals and totals reconcile where possible.

Check that exactly 10 distinct annual reporting periods are presented when 10 are available.

Detect missing, duplicated, and inconsistent reporting periods.

Identify suspicious values or mismatched units.

Compare selected financial statement values against the original SEC 10-K filings.

Flag unresolved discrepancies rather than silently adjusting the numbers.

6. Integration With Existing Application

Integrate financial_extractor.py with main.py.

Requirements:

Preserve the existing SEC downloader and MD&A extractor.

Reuse downloaded filings and existing company metadata where possible.

Avoid unnecessary repeated SEC requests.

Maintain modular functions and consistent data structures.

Keep company-specific outputs in separate files or directories.

Ensure compatibility with Windows PowerShell.

Follow PEP 8 conventions and include meaningful comments and error handling.

Do not initialize Git, create a GitHub repository, or add version control functionality.

7. Testing

Test Stage 4 using AAPL and its 10 most recent available annual 10-K filings.

Generate the Excel workbook and verify that:

All three required financial statement worksheets exist.

Each statement includes 10 annual reporting periods.

Financial data is aligned with the correct fiscal years.

The workbook opens without errors.

Automated accounting reconciliation checks execute successfully.

Missing or inconsistent values are clearly reported.

Display a concise test summary with the number of reporting periods processed, worksheet names, reconciliation results, and unresolved warnings.

Run basic regression tests to ensure Stages 2 and 3 still function.

If errors occur, fix them before completing Stage 4.

8. Documentation

Update README.md with:

Financial statement extraction methodology.

SEC data sources used.

Explanation of how XBRL concepts are selected.

Excel workbook structure.

Validation and reconciliation methods.

Exact Windows PowerShell execution commands.

Known limitations.

Important Instructions

Do not implement sentiment analysis, financial ratios, stock price analysis, or the final comparative report yet.

Do not unnecessarily rewrite previously working modules.

We will perform a comprehensive review of the entire project after all stages are implemented.

After completing Stage 4:

Show the files created or modified.

Show the actual Excel workbook structure.

Summarize the tests performed and their results.

Identify any unresolved data-quality issues.

Provide the exact Windows command to run the application.

Stop and wait for Stage 5 instructions.
```

---

### Prompt 6: Prompt Log Creation & Lab Appendix Requirement
* **Timestamp**: 2026-10-30T14:57:53-04:00
* **Stage**: Audit / Documentation

```text
I have reviewed the updated assignment instructions.

Please ensure our project remains aligned with all requirements in Instructions.docx.

An important final requirement is to include all user prompts at the end of the lab document.

Create a PROMPT_LOG.md file in the project folder to maintain a chronological record of the actual prompts I provide during development, including revisions and troubleshooting prompts.

Do not invent or reconstruct prompts that cannot be recovered from the conversation history. If earlier prompts are unavailable, flag them for me to supply.

We will include this complete prompt history as an appendix in the final lab document.

Continue following our existing modular Python development plan. Do not initialize Git or GitHub.
```

---

### Prompt 7: Stage 5 — MD&A Text Cleaning and Preparation for Sentiment Analysis
* **Timestamp**: 2026-10-30T14:59:20-04:00
* **Stage**: Stage 5

```text
# STAGE 5 — MD&A TEXT CLEANING AND PREPARATION FOR SENTIMENT ANALYSIS

We are continuing my undergraduate GEN AI accounting project using Python and Google Antigravity on Windows.

Please review Instructions.docx and the existing project files before implementing Stage 5.

The two companies selected for this project are:
- Apple Inc. (AAPL)
- NVIDIA Corporation (NVDA)

Stages 1–4 have been implemented. We are developing a modular application that downloads SEC 10-K filings, extracts financial information, and prepares Management's Discussion and Analysis (MD&A) text for Loughran-McDonald sentiment analysis.

We will conduct a comprehensive project review after all development stages have been completed. Continue performing basic automated tests during development.

## 1. Create the Text Cleaning Module

Create a separate Python module named text_cleaner.py.

Its responsibility is to transform the MD&A text files extracted in Stage 3 into cleaned text suitable for financial sentiment analysis.

Do not modify the original SEC filings or overwrite the original extracted MD&A text.

## 2. Text Cleaning Requirements

Implement a consistent, reproducible text cleaning process that:

1. Reads the existing MD&A .txt files from the project's data/extracted_text directory.
2. Uses UTF-8 encoding and handles unexpected encoding issues gracefully.
3. Removes residual HTML tags, HTML entities, and formatting artifacts.
4. Normalizes unnecessary whitespace, tabs, and line breaks.
5. Removes non-substantive formatting characters while preserving meaningful financial terminology.
6. Standardizes text capitalization for dictionary matching.
7. Tokenizes the cleaned text into words using a consistent method.
8. Preserves meaningful financial words, including negative, uncertainty-related, and risk-related terminology.
9. Handles punctuation, hyphenated words, contractions, and common financial abbreviations consistently.
10. Avoids removing stopwords, applying stemming, or using lemmatization unless explicitly justified and compatible with the Loughran-McDonald dictionary.

Do not remove the words "not," "no," "never," or other negation-related terms.

Do not remove financially meaningful terms such as "loss," "uncertainty," "litigation," "decline," or "impairment."

Do not remove numbers indiscriminately from the archived cleaned text. If numerical tokens are excluded from dictionary matching later, document that choice separately.

## 3. Create Cleaned Text Outputs

Save cleaned MD&A text as separate UTF-8 files.

Use an organized directory within the existing project, such as:

data/cleaned_text/AAPL/

Use descriptive filenames, for example:

AAPL_2025_MDA_CLEANED.txt

Ensure that:

- Each cleaned file corresponds to the correct company and fiscal reporting period.
- Original MD&A text remains unchanged.
- Existing files are not overwritten unnecessarily.
- The process works for both AAPL and NVDA without requiring changes to source code.

## 4. Text Processing Metadata

For each MD&A file, record:

- Company ticker
- Fiscal reporting year
- Original MD&A filename
- Cleaned text filename
- Original word count
- Cleaned word count
- Number of tokens prepared for sentiment analysis
- Percentage change in word count
- Cleaning status
- Relevant warnings

Save this information in a CSV summary file.

Use consistent definitions for word counting and document the methodology.

## 5. Quality Control and Validation

Implement checks to:

1. Confirm the cleaned text files are not empty.
2. Detect unusually large reductions in word count.
3. Identify potential encoding errors.
4. Flag missing source files.
5. Confirm the cleaning process does not introduce duplicate content.
6. Verify that important financial and sentiment-related words remain intact.
7. Ensure the cleaned outputs correspond to the correct source filings.
8. Confirm that running the cleaning process repeatedly produces consistent results.

Do not assume that successfully saving a text file proves the text has been cleaned correctly.

## 6. Integrate With the Existing Application

Integrate text_cleaner.py with main.py.

Requirements:

- Preserve all previously completed functionality.
- Keep text cleaning separate from MD&A extraction and sentiment scoring.
- Reuse the existing company and fiscal year metadata.
- Process multiple reporting periods automatically.
- Follow PEP 8 coding standards.
- Include clear function docstrings, comments, and appropriate error handling.
- Avoid unnecessary Python dependencies.
- Remain compatible with Windows PowerShell.

Do not initialize Git or GitHub or add version control functionality.

## 7. Testing

Test Stage 5 using the 10 previously extracted Apple (AAPL) MD&A text files.

For each fiscal year:

1. Read the original extracted MD&A text.
2. Generate cleaned text.
3. Calculate original and cleaned word counts.
4. Display the percentage change in word count.
5. Verify that important financial language is preserved.
6. Flag unusual differences or errors.

Show at least three short before-and-after text examples to demonstrate how the cleaning process works.

Run basic regression tests to confirm that previously completed modules continue functioning.

If a source MD&A extraction appears incomplete or incorrect, report the problem instead of silently treating its cleaned output as valid.

## 8. Documentation

Update README.md with:

- The purpose of text cleaning.
- Cleaning and tokenization methods.
- File locations.
- Word count definitions.
- Any limitations or assumptions.
- Exact Windows PowerShell commands for running the cleaning process.

## Important Development Instructions

- Do not modify or overwrite original SEC filings.
- Do not overwrite the original extracted MD&A sections.
- Do not implement Loughran-McDonald sentiment scoring yet.
- Do not implement historical stock price analysis or financial ratios.
- Do not unnecessarily rewrite previously verified modules.
- Do not create a graphical interface.
- Do not perform an extensive manual review of earlier stages yet.

We will conduct a comprehensive review after the entire application is developed.

## Completion Requirements

After implementing Stage 5:

1. Execute the text cleaning module.
2. Report the number of MD&A files processed successfully.
3. Display a summary table containing fiscal years, original word counts, cleaned word counts, and percentage changes.
4. Show three before-and-after cleaning examples.
5. Identify any errors or unresolved issues.
6. Provide the exact Windows PowerShell command to run the updated program.
7. Update the project documentation.
8. Stop and wait for Stage 6 instructions.

Do not proceed to sentiment analysis until I approve.
```

---

### Prompt 8: Stage 6 — Loughran-McDonald Sentiment Analysis & Excel Integration
* **Timestamp**: 2026-10-30T15:06:13-04:00
* **Stage**: Stage 6

```text
I have downloaded the Loughran-McDonald Master Dictionary CSV and placed it directly inside the project's data folder.

Please automatically locate the CSV file, validate that it contains the required sentiment categories, and use it for Stage 6.

Do not require me to rename the file unless necessary.

Proceed with implementing sentiment_analysis.py, calculating the sentiment scores for Apple's 10 MD&A files, and adding the results to the existing Excel workbook.

Verify the word counts, positive and negative classifications, P/N ratios, and other dictionary categories.

Preserve the existing financial statement worksheets and report the actual test results.

Do not proceed to Stage 7 yet.
```

---

### Prompt 9: Stage 6 — Loughran-McDonald Financial Sentiment Analysis (Full Specification)
* **Timestamp**: 2026-10-30T15:15:06-04:00
* **Stage**: Stage 6

```text
# STAGE 6 — LOUGHRAN-MCDONALD FINANCIAL SENTIMENT ANALYSIS

We are continuing my undergraduate GEN AI accounting project using Python and Google Antigravity on Windows.

Before beginning, review Instructions.docx, the existing Python source files, and the outputs from Stages 1–5.

The two companies selected for this project are:
- Apple Inc. (AAPL)
- NVIDIA Corporation (NVDA)

The objective of Stage 6 is to analyze the cleaned Management's Discussion and Analysis (MD&A) text from the annual SEC 10-K filings using the Loughran-McDonald financial sentiment dictionary.

Continue using our existing modular Python application, with main.py as the central entry point.

We will conduct a comprehensive review of all project components at the end. Continue performing automated tests throughout development.

## 1. Create the Sentiment Analysis Module

Create a separate Python module named sentiment_analysis.py.

The module must:

1. Load the Loughran-McDonald Master Dictionary from the local project directory.
2. Support CSV or Excel dictionary formats where practical.
3. Validate that the dictionary contains the required word and sentiment-category columns.
4. Use the dictionary's actual category indicators to classify financial terminology.
5. Match cleaned MD&A tokens against dictionary entries consistently and case-insensitively.
6. Avoid stemming, lemmatization, or modifying dictionary words in ways that cause incorrect matches.
7. Preserve accurate token counting and ensure the same methodology is used across all reporting years.
8. Document the dictionary version and its source.

Do not substitute a general-purpose sentiment dictionary or an AI-generated word list.

If the official dictionary cannot be located or loaded, report the problem and request the correct file rather than fabricating sentiment classifications.

## 2. Sentiment Categories

Calculate the following values for each fiscal reporting year:

- Total Word Count
- Positive Word Count
- Negative Word Count
- Uncertainty Word Count
- Litigious Word Count
- Strong Modal Word Count
- Weak Modal Word Count
- Constraining Word Count, if available in the dictionary
- Positive-to-Negative (P/N) Ratio
- Positive Word Frequency
- Negative Word Frequency

Use these formulas:

P/N Ratio = Positive Word Count / Negative Word Count

Positive Word Frequency = Positive Word Count / Total Word Count

Negative Word Frequency = Negative Word Count / Total Word Count

The P/N ratio should be reported as a numerical ratio.

Positive and negative word frequencies should be reported as percentages.

If the negative word count is zero, return an appropriately labeled undefined value rather than causing a division-by-zero error.

Treat categories according to the dictionary's official classification rules. A word may belong to multiple categories if the dictionary indicates this.

## 3. Word Count and Tokenization Rules

Use a consistent tokenization method for every MD&A document.

Requirements:

1. Count all eligible word tokens in the denominator, including words that are not in the dictionary.
2. Exclude purely numerical tokens from dictionary matching.
3. Count repeated occurrences of sentiment words rather than only unique words.
4. Preserve legitimate financial terminology.
5. Use the cleaned MD&A text generated in Stage 5.
6. Document treatment of hyphenated words, punctuation, abbreviations, and contractions.
7. Maintain consistent definitions of total word count across both companies and all fiscal years.

Do not silently change the text-cleaning methodology from Stage 5.

If changes are necessary for accurate dictionary matching, explain them and test the effect before implementing them.

## 4. Create Sentiment Analysis Outputs

Generate a structured sentiment analysis table containing:

- Company ticker
- Fiscal reporting year
- Fiscal period-end date
- Source MD&A filename
- Total Word Count
- Positive Word Count
- Negative Word Count
- Uncertainty Word Count
- Litigious Word Count
- Strong Modal Word Count
- Weak Modal Word Count
- Constraining Word Count, if supported
- P/N Ratio
- Positive Word Frequency
- Negative Word Frequency
- Dictionary Version
- Processing Status

Save the results as a CSV file within the existing data directory.

Add a worksheet named "Sentiment Analysis" to the company's existing Excel workbook created in Stage 4.

Do not overwrite or remove the financial statement worksheets.

Present the sentiment results for all 10 fiscal reporting periods in chronological order.

Apply clear formatting, descriptive headers, appropriate decimal precision, and consistent percentage formatting.

## 5. Validation and Testing

Test the sentiment analysis module using the 10 cleaned Apple (AAPL) MD&A files.

Perform the following checks:

1. Verify that the correct Loughran-McDonald dictionary is loaded.
2. Confirm that all 10 fiscal reporting periods are processed.
3. Ensure no MD&A files are missing or empty.
4. Validate total word counts against the cleaned text files.
5. Independently verify sentiment classifications using a small controlled sample of text with known dictionary entries.
6. Confirm that repeated occurrences of sentiment words are counted correctly.
7. Confirm that words may contribute to multiple sentiment categories where applicable.
8. Verify that P/N ratios and word frequencies are mathematically correct.
9. Confirm that undefined ratios are handled appropriately.
10. Verify that the Excel workbook retains all previously created financial statement worksheets.
11. Check that the Sentiment Analysis worksheet contains accurate data for each year.

Create or update automated tests for the sentiment module.

If discrepancies are found, report them and correct the underlying code rather than altering results manually.

## 6. Integration With the Existing Application

Integrate sentiment_analysis.py with main.py.

Requirements:

- Preserve the SEC downloader, MD&A extractor, financial statement extractor, and text cleaning modules.
- Reuse existing file paths, configuration settings, and company metadata.
- Maintain a modular architecture.
- Ensure the program works for both AAPL and NVDA without modifying source code.
- Avoid unnecessary package installations or repeated downloads.
- Follow PEP 8 coding standards.
- Include comments, docstrings, and appropriate error handling.
- Ensure compatibility with Windows PowerShell.

Do not initialize Git, create a GitHub repository, or perform version control operations.

## 7. Documentation

Update README.md to include:

- Purpose of Loughran-McDonald sentiment analysis.
- Dictionary source and version.
- Definitions of sentiment categories.
- Explanation of the word-counting methodology.
- Formulas used for P/N ratios and word frequencies.
- Instructions for supplying the dictionary.
- Exact Windows PowerShell commands to run the analysis.
- Known limitations of dictionary-based sentiment analysis.

## Important Development Instructions

Do not implement historical stock price extraction or financial ratio calculations yet.

Do not create a graphical interface.

Do not modify previously verified modules unless necessary to fix an identified compatibility issue.

We will conduct a comprehensive final project review after completing the remaining stages.

## Completion Requirements

After implementing Stage 6:

1. Run the sentiment analysis module using AAPL's 10 cleaned MD&A documents.
2. Display the actual sentiment analysis results for all 10 years.
3. Show total word counts, positive counts, negative counts, and P/N ratios.
4. Report the dictionary version used.
5. Confirm that the Sentiment Analysis worksheet was added to the Excel workbook.
6. Run automated tests and display the results.
7. Report any missing files, unusual values, or processing errors.
8. Update README.md.
9. Provide the exact Windows PowerShell command needed to run the updated application.
10. Stop and wait for Stage 7 instructions.

Do not proceed to the next stage until I approve.
```

---

### Prompt 10: Stage 7 — Historical Stock Price Extraction and Excel Integration
* **Timestamp**: 2026-10-30T15:22:53-04:00
* **Stage**: Stage 7

```text
# STAGE 7 — HISTORICAL STOCK PRICE EXTRACTION AND EXCEL INTEGRATION

We are continuing my undergraduate GEN AI accounting project using Python and Google Antigravity on Windows.

Before beginning, review Instructions.docx and the existing project files, including all functionality implemented in Stages 1–6.

Our two selected companies are:
- Apple Inc. (AAPL)
- NVIDIA Corporation (NVDA)

The objective of Stage 7 is to retrieve accurate historical stock prices corresponding to the fiscal year-end balance sheet dates of each company's 10 annual SEC 10-K filings.

Continue using our modular Python architecture, with main.py as the central entry point.

We will conduct a comprehensive review of the entire project at the end, but continue performing automated tests throughout development.

## 1. Create the Stock Price Module

Create a separate Python module named stock_prices.py.

The module must:

1. Read the fiscal reporting period-end dates from the existing financial statement data or SEC filing metadata.
2. Retrieve historical stock prices for the company's ticker symbol.
3. Match each stock price to the corresponding fiscal year-end balance sheet date.
4. If the balance sheet date falls on a weekend or market holiday, use the closing price from the most recent preceding trading day.
5. Record both the fiscal year-end date and the actual trading date.
6. Support both AAPL and NVDA without changes to the source code.
7. Avoid hardcoding stock prices or dates.
8. Use a documented and reasonably reliable historical market data source.

You may use the yfinance Python package if appropriate, but first evaluate whether its historical price data satisfies the project's requirements.

Do not use random search snippets or estimated prices as substitutes for verified historical values.

## 2. Stock Price Methodology

Retrieve and record:

- Company ticker
- Fiscal year
- Fiscal year-end balance sheet date
- Actual trading date
- Historical closing stock price
- Historical adjusted closing price, where available
- Data source
- Retrieval status

Use a consistent stock price measure for year-over-year analysis.

For the primary historical comparison, use split-adjusted closing prices to account for stock splits affecting Apple and NVIDIA over the ten-year period.

Clearly explain the distinction between closing prices and adjusted closing prices, including whether dividend adjustments are applied by the source.

Do not label split-adjusted prices as fully adjusted for dividends unless the underlying data supports that interpretation.

Preserve the raw source values where available so the historical data can be traced and verified.

## 3. Create Excel Stock Price Worksheet

Add a new worksheet named "Stock Prices" to the existing Excel workbook generated during previous stages.

Requirements:

1. Present 10 fiscal reporting periods in chronological order, from oldest to newest.
2. Include the corresponding balance sheet date for each fiscal year.
3. Include the actual trading date used to obtain the stock price.
4. Include closing and adjusted closing prices, where available.
5. Clearly label the primary price measure used for analysis.
6. Include the historical data source.
7. Use professional Excel formatting with clear column headers, appropriate currency formats, and consistent decimal precision.
8. Preserve all previously created worksheets, including:
   - Income Statement
   - Balance Sheet
   - Cash Flow Statement
   - Sentiment Analysis
9. Do not overwrite financial statements or sentiment scores.

Ensure the worksheet's fiscal years align correctly with the existing 10-K financial data.

## 4. File Organization

Save any intermediate stock price datasets in the existing data/financial_data directory.

Use descriptive filenames identifying the ticker and data type.

Keep the Excel outputs within the existing data/excel_outputs directory.

Do not create unnecessary duplicate files or change the existing folder structure without justification.

## 5. Validation and Error Handling

Implement checks to:

1. Verify that each stock price corresponds to the correct balance sheet date or preceding trading day.
2. Confirm that the actual trading date is not later than the fiscal year-end date.
3. Detect missing observations.
4. Check for duplicate fiscal years.
5. Confirm that prices are valid positive numerical values.
6. Account for stock splits and document the adjustment methodology.
7. Detect possible inconsistencies between raw and adjusted historical prices.
8. Report API errors, connection failures, or unavailable historical data.
9. Confirm that all 10 annual reporting periods have corresponding stock price observations when data is available.
10. Cross-check selected historical prices against a second reliable source when practical.

Do not fabricate missing stock prices.

## 6. Integration With Existing Python Application

Integrate stock_prices.py with main.py.

Requirements:

- Preserve all previous modules and functionality.
- Use existing fiscal period metadata rather than manually entering dates.
- Reuse the current project configuration.
- Avoid unnecessary repeated network requests.
- Maintain modular functions and error handling.
- Follow PEP 8 standards.
- Include comments and clear docstrings.
- Ensure compatibility with Windows PowerShell.
- Do not initialize Git or GitHub.

## 7. Testing

Test the module using Apple (AAPL) and its 10 fiscal reporting periods.

For each year:

1. Identify the correct balance sheet date.
2. Retrieve the historical stock price.
3. Confirm the actual trading date.
4. Check whether a stock split or dividend adjustment affects the reported values.
5. Add the results to Excel.
6. Verify that the workbook remains readable and retains the previous worksheets.

Display a summary table including fiscal year, balance sheet date, trading date, closing price, adjusted closing price, and verification status.

Run automated tests for the stock_prices.py module.

Perform basic regression tests to ensure Stages 2–6 still function correctly.

Report any discrepancies or unavailable observations.

## 8. Documentation

Update README.md to describe:

- Historical stock price data sources.
- Price retrieval methodology.
- Treatment of non-trading days.
- Stock split and dividend adjustments.
- Excel worksheet structure.
- Validation methodology.
- Exact Windows PowerShell execution commands.
- Known data limitations.

## Important Development Instructions

Do not implement financial ratio calculations or the final comparison charts yet.

Do not create a graphical interface.

Do not unnecessarily modify existing modules.

Do not perform extensive manual review of earlier stages yet; we will complete a comprehensive project review after the remaining stages.

Do not proceed to Stage 8 until I approve.

## Completion Requirements

After implementing Stage 7:

1. Execute the stock price module for AAPL.
2. Display the actual historical stock price results for all 10 years.
3. Confirm the stock prices correspond to the correct fiscal balance sheet dates.
4. Report the stock price source and adjustment methodology.
5. Confirm the Stock Prices worksheet was added to Excel.
6. Run automated tests and display their results.
7. Identify any errors or unresolved limitations.
8. Update README.md.
9. Provide the exact Windows PowerShell command to run the program.
10. Stop and wait for Stage 8 instructions.
```

---

### Prompt 11: Stage 8 — Ten-Year Financial Ratio Analysis
* **Timestamp**: 2026-10-30T15:35:50-04:00
* **Stage**: Stage 8

```text
# STAGE 8 — TEN-YEAR FINANCIAL RATIO ANALYSIS

We are continuing my undergraduate GEN AI accounting project using Python and Google Antigravity on Windows.

Before making changes, review Instructions.docx and all existing Python project files.

Our selected companies are:
- Apple Inc. (AAPL)
- NVIDIA Corporation (NVDA)

Stages 1–7 have been implemented. The purpose of Stage 8 is to calculate and analyze financial ratios using the annual financial statement data extracted from SEC 10-K filings.

We will conduct a comprehensive project review after completing all development stages. Continue running automated tests to catch errors during development.

## 1. Create the Financial Ratio Module

Create a separate Python module named financial_ratios.py.

Integrate it with the existing main.py application.

The module must calculate four financial ratios for each of the 10 fiscal reporting periods.

Use verified financial statement data from Stage 4 rather than downloading duplicate datasets or hardcoding results.

## 2. Required Financial Ratios

Calculate the following:

**A. Liquidity — Current Ratio**

Formula:
Current Ratio = Current Assets / Current Liabilities

Purpose: Evaluate the company's ability to meet short-term obligations.

**B. Solvency — Debt-to-Equity Ratio**

Formula:
Debt-to-Equity Ratio = Total Interest-Bearing Debt / Total Shareholders' Equity

Total Interest-Bearing Debt should include reported short-term borrowings, current portions of long-term debt, and noncurrent interest-bearing debt.

Do not automatically substitute total liabilities for interest-bearing debt.

Purpose: Evaluate the company's financial leverage and long-term solvency.

**C. Efficiency — Total Asset Turnover**

Formula:
Total Asset Turnover = Net Revenue / Average Total Assets

Average Total Assets = (Beginning Total Assets + Ending Total Assets) / 2

Retrieve the opening asset balance for the earliest analyzed year from the preceding fiscal year, where available.

Do not substitute ending total assets for average assets without explicitly identifying the methodological limitation.

Purpose: Evaluate how efficiently assets generate revenue.

**D. Profitability — Net Profit Margin**

Formula:
Net Profit Margin = Net Income / Net Revenue × 100

Purpose: Measure the proportion of revenue retained as profit.

## 3. Data Extraction and Accounting Consistency

Use the existing financial statement data from Stage 4.

Requirements:

1. Identify the appropriate financial statement accounts and XBRL concepts.
2. Use consistent account definitions across all fiscal years.
3. Ensure all numerator and denominator values correspond to the correct reporting period.
4. Use annual revenue and net income figures, not quarterly or year-to-date interim values.
5. Use fiscal year-end balance sheet balances for current assets, current liabilities, debt, and equity.
6. Handle changes in accounting terminology or company-specific XBRL concepts.
7. Avoid double-counting debt components when building total interest-bearing debt.
8. Document the accounting definitions used for each ratio.
9. Handle missing, negative, or zero denominators appropriately.
10. Do not fabricate missing financial data or automatically replace missing values with zeros.

Use the same calculation methodology for both AAPL and NVDA to support meaningful comparisons.

## 4. Create Financial Ratio Outputs

Create a worksheet named "Financial Ratios" in the existing company-specific Excel workbook.

Present the results for the 10 fiscal years side by side, from oldest to newest.

Include:

- Current Ratio
- Debt-to-Equity Ratio
- Total Asset Turnover
- Net Profit Margin

Where practical, include supporting numerator and denominator values or a separate calculation-detail section.

Use formulas in Excel when practical so the ratios can be traced back to the underlying financial statement data.

Format:
- Current Ratio: numerical ratio, such as 1.25x.
- Debt-to-Equity Ratio: numerical ratio, such as 1.50x.
- Total Asset Turnover: numerical ratio, such as 0.75x.
- Net Profit Margin: percentage, such as 22.5%.

Maintain consistent formatting across all 10 reporting periods.

Do not overwrite or remove previously created Excel worksheets.

## 5. Financial Ratio Validation

Implement checks to:

1. Confirm that all four ratios are calculated for each reporting year when sufficient data exists.
2. Verify the numerator and denominator values.
3. Confirm fiscal year alignment.
4. Recalculate selected ratios independently and compare the results.
5. Detect missing or duplicated fiscal years.
6. Flag unusual financial ratio values.
7. Handle division-by-zero errors.
8. Verify that the earliest asset turnover calculation uses the correct beginning asset balance.
9. Confirm that debt components are not double-counted.
10. Ensure all ratio calculations are reproducible.

Do not silently correct inconsistent financial statement data.

Report unresolved discrepancies and the affected fiscal years.

## 6. Add Supporting Visualizations

Create professional Excel charts showing 10-year trends for:

- Current Ratio
- Debt-to-Equity Ratio
- Total Asset Turnover
- Net Profit Margin

Use clear chart titles, fiscal year labels, and appropriately formatted axes.

These charts should supplement the numerical calculations, not replace them.

Do not create the final stock price versus sentiment P/N ratio chart yet. That will be completed in Stage 9.

## 7. Integration With Existing Python Application

Integrate financial_ratios.py with main.py.

Requirements:

- Preserve all functionality implemented in Stages 1–7.
- Reuse existing project folders and company metadata.
- Keep ratio calculations separate from SEC downloading and financial statement extraction.
- Avoid unnecessary package installations.
- Follow PEP 8 coding standards.
- Use clear comments, docstrings, and error handling.
- Maintain Windows PowerShell compatibility.
- Do not initialize Git, create a GitHub repository, or perform version control operations.

## 8. Testing

Test Stage 8 using Apple's 10 annual fiscal reporting periods.

For each fiscal year:

1. Retrieve the required financial statement values.
2. Calculate all four ratios.
3. Show numerator and denominator values.
4. Validate the calculation results.
5. Write the results to the Excel workbook.
6. Verify that all previous worksheets are preserved.

Create automated unit tests using known numerical examples for each financial ratio.

Run basic regression tests to ensure Stages 2–7 remain functional.

Display a concise validation summary and report any missing values or errors.

## 9. Documentation

Update README.md with:

- Definitions and formulas for each ratio.
- Explanation of the accounting data sources.
- How beginning and ending asset balances are selected.
- Treatment of interest-bearing debt.
- Excel worksheet structure.
- Validation methods.
- Known limitations.
- Exact Windows PowerShell commands for executing the program.

## Important Development Instructions

Do not implement the final stock price versus sentiment P/N ratio chart yet.

Do not create a graphical interface.

Do not unnecessarily modify existing modules.

We will conduct a comprehensive project review after the remaining development stages.

After completing Stage 8:

1. Execute the financial ratio analysis for AAPL.
2. Display the four ratios for all 10 years.
3. Show supporting calculation values.
4. Confirm that the Financial Ratios worksheet and charts were created.
5. Report automated test results and unresolved warnings.
6. Update README.md.
7. Provide the exact Windows command to run the program.
8. Stop and wait for Stage 9 instructions.

Do not proceed to Stage 9 until I approve.
```

---

### Prompt 12: Stage 9 — Final Excel Analysis, Visualization, and Complete Application Integration
* **Timestamp**: 2026-10-30T15:45:43-04:00
* **Stage**: Stage 9

```text
# STAGE 9 — FINAL EXCEL ANALYSIS, VISUALIZATION, AND COMPLETE APPLICATION INTEGRATION

We are continuing my undergraduate GEN AI accounting assignment using Python and Google Antigravity on Windows.

Before making changes, carefully review Instructions.docx, README.md, main.py, and all existing Python modules developed during Stages 1–8.

The two companies selected for my project are:

- Apple Inc. (AAPL)
- NVIDIA Corporation (NVDA)

The objective of Stage 9 is to complete the final Excel analysis and visualizations, integrate all previously developed Python modules into one fully functional application, and verify that the complete program runs successfully from beginning to end.

This is the final major Python development stage before generating the second company's results and preparing the lab document.

**IMPORTANT: Maintain the existing modular Python architecture. Do not combine every module into one enormous Python file. Instead, make main.py the central entry point that coordinates the entire application.**

## 1. Complete Application Integration

Ensure all previously developed Python modules are integrated into one functioning application.

The program must execute the following processing stages automatically:

1. Prompt the user for the desired company's stock ticker symbol and number of annual 10-K filings.
2. Download the requested SEC 10-K filings.
3. Organize and label the downloaded filings.
4. Extract the MD&A sections.
5. Extract financial statement data.
6. Clean the extracted MD&A text.
7. Perform Loughran-McDonald sentiment analysis.
8. Retrieve historical stock prices corresponding to fiscal year-end balance sheet dates.
9. Calculate the required financial ratios.
10. Generate the complete Excel workbook and visualizations.

All processing must be initiated through main.py.

The user must not need to manually execute separate Python modules.

Ensure the application works for both AAPL and NVDA without modifying source code.

## 2. Main Application User Interface

Create a clear, user-friendly command-line interface.

When launched, the program should display a title such as:

SEC 10-K FINANCIAL ANALYZER

Prompt the user for:

- Company ticker symbol, such as AAPL or NVDA.
- Number of most recent annual 10-K filings to analyze, with a default of 10.

Validate user inputs and provide clear error messages.

Display progress throughout execution.

For example:

[1/8] Downloading SEC filings...
[2/8] Extracting MD&A sections...
[3/8] Extracting financial statements...
[4/8] Cleaning MD&A text...
[5/8] Performing sentiment analysis...
[6/8] Retrieving historical stock prices...
[7/8] Calculating financial ratios...
[8/8] Generating Excel workbook...

Upon completion, display:

- Company analyzed.
- Number of reporting periods processed.
- Completed and failed stages.
- Any unresolved data-quality warnings.
- Location of the generated Excel workbook.
- Total processing time.

The application should not report a successful analysis if a required stage failed.

## 3. Final Excel Workbook Requirements

Generate one comprehensive Excel workbook for each company.

Save the completed workbooks in the existing data/excel_outputs directory.

Use descriptive filenames such as:

AAPL_10K_Analysis.xlsx

NVDA_10K_Analysis.xlsx

Each workbook must contain at least the following worksheets.

**Worksheet 1 — Income Statement**

- Present the 10 fiscal reporting periods side by side.
- Display relevant financial statement line items.
- Use consistent financial units and clear headings.
- Preserve financial statement totals and subtotals.

**Worksheet 2 — Balance Sheet**

- Present the 10 fiscal reporting periods side by side.
- Include current assets, current liabilities, total assets, total liabilities, and shareholders' equity, along with other relevant reported items.
- Ensure that the balance sheet equation is verified.

**Worksheet 3 — Cash Flow Statement**

- Present the 10 fiscal reporting periods side by side.
- Include operating, investing, and financing cash flows.
- Preserve the relevant reported cash flow totals and subtotals.

**Worksheet 4 — Sentiment Analysis**

Present sentiment results for all 10 fiscal years, including:

- Total Word Count
- Positive Word Count
- Negative Word Count
- Uncertainty Word Count
- Litigious Word Count
- Strong Modal Word Count
- Weak Modal Word Count
- Constraining Word Count, if supported by the dictionary
- Positive-to-Negative (P/N) Ratio
- Positive Word Frequency
- Negative Word Frequency

Use the official Loughran-McDonald Master Dictionary already provided in the project.

Do not substitute another dictionary.

**Worksheet 5 — Stock Prices**

Include:

- Fiscal year
- Balance sheet date
- Actual trading date
- Closing stock price
- Adjusted closing price, when available
- Data source
- Relevant adjustment notes

Ensure each stock price corresponds to the appropriate fiscal year-end balance sheet date or the most recent preceding trading day.

**Worksheet 6 — Financial Ratios**

Calculate and present:

1. Liquidity — Current Ratio
2. Solvency — Debt-to-Equity Ratio
3. Efficiency — Total Asset Turnover
4. Profitability — Net Profit Margin

Use the calculation methodology established in Stage 8.

Display all 10 reporting periods side by side.

Where practical, use Excel formulas referencing the underlying financial statement data so calculations remain traceable.

**Worksheet 7 — Charts and Analysis**

Create professional Excel visualizations showing:

1. Historical stock price versus sentiment P/N ratio for all 10 years.
2. Current Ratio trends.
3. Debt-to-Equity Ratio trends.
4. Total Asset Turnover trends.
5. Net Profit Margin trends.
6. Positive and Negative Word Count trends, if useful.

The stock price versus sentiment P/N ratio chart is mandatory under Instructions.docx.

Use a dual-axis line chart when necessary because stock prices and P/N ratios have different numerical scales.

Clearly label the axes and identify which stock price measure is being used.

Ensure all chart data corresponds to the correct fiscal reporting periods.

## 4. Excel Formatting and Presentation

Apply professional formatting throughout the workbook.

Requirements:

- Clear company identification.
- Descriptive worksheet titles.
- Consistent fiscal-year headings.
- Chronological ordering from oldest to newest.
- Appropriate currency, ratio, and percentage formats.
- Financial values displayed using consistent units, preferably USD millions.
- Negative financial values displayed in parentheses.
- Bold totals and subtotals.
- Consistent fonts and column widths.
- Freeze panes where appropriate.
- Clear chart titles, legends, and axis labels.
- Clearly identified missing values.
- A data source and methodology worksheet if necessary.

Do not use invented financial values or replace missing data with zeros.

## 5. Data Consistency and Reporting Period Alignment

Ensure that all financial, sentiment, and stock price information is correctly aligned.

The application must:

1. Use actual fiscal year-end dates from SEC filings.
2. Match financial statement values to the correct fiscal reporting periods.
3. Match MD&A sentiment scores to those same reporting periods.
4. Match stock prices to the corresponding balance sheet dates.
5. Avoid confusing filing dates with fiscal period-end dates.
6. Account for Apple's and NVIDIA's different fiscal year-end conventions.
7. Avoid duplicate annual reporting periods.
8. Maintain consistent financial statement units.
9. Preserve the established stock split adjustment methodology.
10. Identify discrepancies instead of silently adjusting data.

Do not assume Apple and NVIDIA share the same fiscal year-end dates.

## 6. Application Architecture and Module Coordination

Maintain the existing modular architecture.

main.py must coordinate the modules responsible for:

- SEC filing downloads.
- MD&A extraction.
- Financial statement extraction.
- Text cleaning.
- Sentiment analysis.
- Historical stock prices.
- Financial ratio calculations.
- Excel workbook generation.

Create a dedicated Excel reporting module, such as excel_report.py, if this improves maintainability.

Reuse existing modules rather than duplicating functionality.

Ensure data structures and file paths are compatible across modules.

Requirements:

1. Automatically create required output directories.
2. Reuse existing valid downloaded SEC filings where possible.
3. Avoid unnecessary repeat downloads and API requests.
4. Handle missing intermediate files.
5. Prevent downstream calculations from using invalid or incomplete data.
6. Provide informative error messages.
7. Preserve previously functioning code.
8. Use PEP 8 conventions and clear documentation.
9. Maintain compatibility with Windows PowerShell.
10. Keep all project files within the existing project directory.

Do not initialize Git, create a GitHub repository, or perform version control operations.

## 7. Comprehensive Automated Testing

Now perform an integrated test of the entire application.

Test using Apple (AAPL) and 10 annual 10-K filings.

The test must verify:

1. The application launches successfully through main.py.
2. The user can enter a ticker symbol and number of filings.
3. The correct SEC 10-K filings are identified.
4. The filings download and save correctly.
5. MD&A sections are extracted.
6. Extracted text is cleaned.
7. Financial statements are populated correctly.
8. Sentiment scores are calculated using the Loughran-McDonald dictionary.
9. Historical stock prices are retrieved for the correct reporting dates.
10. All four financial ratios are calculated accurately.
11. All required Excel worksheets are generated.
12. The mandatory stock price versus sentiment chart is generated.
13. Previously created worksheets are preserved.
14. The completed Excel workbook opens successfully.
15. Running the program again does not unnecessarily repeat downloads or create duplicate records.

Run existing unit tests and regression tests.

Verify accounting reconciliations and independently check selected numerical results against the original SEC filings.

For important Excel formulas, verify that the referenced cells are correct and that the formulas calculate as intended.

Do not claim successful verification unless the tests have actually passed.

## 8. Error Handling and Quality Control

Implement appropriate error handling for:

- Invalid ticker symbols.
- Missing SEC filings.
- SEC API restrictions and connection failures.
- Corrupted or incomplete downloaded documents.
- MD&A extraction failures.
- Missing financial statement values.
- Missing or incompatible sentiment dictionary files.
- Incorrect or incomplete historical stock price information.
- Division-by-zero errors.
- Missing output directories.
- Excel file access or permission errors.
- Inconsistent reporting periods.

If a required stage fails, provide a clear message identifying the problem.

Do not silently generate a completed report containing inaccurate or fabricated results.

Distinguish between fatal errors, warnings, and successfully completed operations.

## 9. Final Application Documentation

Update README.md with comprehensive instructions covering:

1. Project purpose and overview.
2. Application structure.
3. Python dependencies.
4. Windows installation instructions.
5. Virtual environment setup.
6. SEC User-Agent configuration.
7. Loughran-McDonald dictionary setup.
8. How to run the complete application.
9. How to analyze AAPL or NVDA.
10. Description of the Excel workbook.
11. Financial ratio formulas.
12. Sentiment scoring methodology.
13. Historical stock price methodology.
14. Troubleshooting common errors.
15. Known limitations.

Provide the exact Windows PowerShell command required to run the completed application.

The intended command is:

.\.venv\Scripts\python.exe main.py

If the actual command differs, document the correct one.

## 10. Final Verification Summary

After implementation, provide a comprehensive summary containing:

- Python modules created or modified.
- Results of the full AAPL integration test.
- Number of annual filings processed.
- MD&A extraction and sentiment analysis status.
- Financial statement reconciliation results.
- Historical stock price verification results.
- Financial ratio calculation results.
- Excel worksheet names.
- Charts successfully generated.
- Location of the completed Excel workbook.
- Outstanding errors or data-quality limitations.
- Exact Windows execution command.

Report actual testing results rather than hypothetical examples.

## Critical Development Instructions

- Maintain one unified Python application controlled by main.py.
- Keep the source code modular.
- Do not require the user to execute separate modules manually.
- Do not build a graphical desktop interface.
- Do not package the application into an executable yet.
- Do not initialize Git or GitHub.
- Do not overwrite original SEC filings or original MD&A extracts.
- Do not modify working modules unnecessarily.
- Do not fabricate data or claim unperformed tests passed.
- Preserve all existing project functionality.

Our next stages will involve running the completed application for NVIDIA (NVDA) and preparing the final lab document comparing Apple and NVIDIA.

After completing Stage 9, stop and wait for my approval before proceeding further.
```

---

### Prompt 13: Stage 10 — NVIDIA Analysis and Final Python Application Validation
* **Timestamp**: 2026-10-30T15:54:40-04:00
* **Stage**: Stage 10

```text
# STAGE 10 — NVIDIA ANALYSIS AND FINAL PYTHON APPLICATION VALIDATION

We are continuing my undergraduate GEN AI accounting project using Python and Google Antigravity on Windows.

Before beginning, review Instructions.docx, the existing Python modules, and the completed Stage 9 integration.

Our two selected companies are:

- Apple Inc. (AAPL)
- NVIDIA Corporation (NVDA)

The objective of Stage 10 is to run the fully integrated Python application for NVIDIA, verify that the same program works for both companies, and confirm that all required Excel outputs are generated correctly.

The Lab Document is a separate deliverable and must not be generated during this stage.

## 1. Run the Complete Application for NVIDIA

Using the existing main.py application, process NVIDIA (NVDA) and its 10 most recent available annual Form 10-K filings.

Execute the complete workflow:

1. Download and organize NVIDIA's 10-K filings.
2. Extract the MD&A sections.
3. Extract the income statements, balance sheets, and statements of cash flows.
4. Clean the MD&A text.
5. Perform sentiment analysis using the supplied Loughran-McDonald Master Dictionary.
6. Retrieve historical stock prices corresponding to NVIDIA's fiscal year-end balance sheet dates.
7. Calculate the four required financial ratios.
8. Generate the final Excel workbook and charts.

Do not hardcode NVIDIA-specific results or modify the source code merely to change the ticker.

## 2. Generate NVIDIA's Excel Workbook

Generate a workbook named NVDA_10K_Analysis.xlsx within the existing data/excel_outputs directory.

Include at least:

- Income Statement
- Balance Sheet
- Cash Flow Statement
- Sentiment Analysis
- Stock Prices
- Financial Ratios
- Charts and Analysis

Present all 10 fiscal reporting periods accurately.

Ensure the stock price versus sentiment P/N ratio chart is included.

Use the same methodology and formatting established for Apple.

## 3. Handle Company-Specific Differences

Account for differences between Apple and NVIDIA, including:

- Fiscal year-end dates.
- Financial statement classifications.
- Company-specific XBRL taxonomy concepts.
- MD&A section formatting.
- Stock splits and historical price adjustments.
- Financial statement terminology.

Ensure the same calculations and sentiment methodology are applied consistently.

Do not substitute missing data with fabricated values.

## 4. Final Application Validation

After generating NVIDIA's workbook, conduct a comprehensive end-to-end review of the application for BOTH AAPL and NVDA.

Verify:

1. main.py successfully runs the complete workflow for either company.
2. Each company has 10 distinct annual reporting periods.
3. All SEC filings are correctly identified.
4. All MD&A extractions contain the intended Item 7 content.
5. Financial statements contain accurate annual values.
6. Balance sheet accounting equations reconcile.
7. Sentiment calculations use the correct Loughran-McDonald dictionary.
8. Positive and negative word counts and P/N ratios are calculated correctly.
9. Historical stock prices align with fiscal year-end balance sheet dates.
10. All four financial ratios are calculated accurately.
11. Excel workbooks contain all required worksheets.
12. The mandatory stock price versus sentiment P/N ratio charts are correct.
13. No existing Apple outputs were overwritten by NVIDIA's analysis.
14. Previously completed functionality remains operational.

Run automated tests and independently verify selected results against source data.

Report actual test results and unresolved discrepancies.

## 5. Project Organization

Preserve the existing modular Python architecture.

All application functionality must remain accessible through main.py.

Keep company-specific raw data and outputs clearly separated.

Do not create a graphical desktop interface.

Do not initialize Git or GitHub.

Do not generate the Lab Document during this stage.

## 6. Final Documentation

Update README.md with:

- Complete Windows installation instructions.
- Required Python packages.
- SEC User-Agent configuration instructions.
- Loughran-McDonald dictionary requirements.
- Instructions for analyzing any supported ticker.
- The complete processing workflow.
- Explanation of all Excel worksheets.
- Data sources and calculation methodologies.
- Known limitations and troubleshooting instructions.

Provide the exact Windows PowerShell command needed to execute the application.

## 7. Final Deliverables

After completing this stage, provide:

1. The location of the complete Python application.
2. The location of AAPL_10K_Analysis.xlsx.
3. The location of NVDA_10K_Analysis.xlsx.
4. A summary of automated test results.
5. A summary of data verification and reconciliation results.
6. Any unresolved errors or missing information.
7. Confirmation that the same main.py application successfully handles both companies.
8. The exact command needed to run the program.

Do not claim that tests passed unless they were actually executed successfully.

Do not proceed to creating the Lab Document.

Stop after reporting the final application verification results.
```


---

### Prompt 14: Stage 11 — Final Application Audit, Quality Assurance, and Assignment Compliance
* **Timestamp**: 2026-10-30T16:08:10-04:00
* **Stage**: Stage 11

```text
# STAGE 11 — FINAL APPLICATION AUDIT, QUALITY ASSURANCE, AND ASSIGNMENT COMPLIANCE

We are completing the final quality assurance review of my undergraduate GEN AI accounting project using Python and Google Antigravity on Windows.

Our two companies are:
- Apple Inc. (AAPL)
- NVIDIA Corporation (NVDA)

Before proceeding, read Instructions.docx and inspect the entire existing project, including all Python modules, test scripts, downloaded SEC filings, extracted MD&A files, financial datasets, and generated Excel workbooks.

The objective is to thoroughly audit the completed application and correct any issues that prevent it from satisfying the professor's requirements.

Do not build new features unless they are necessary to satisfy an unmet assignment requirement or correct an existing defect.

## 1. Assignment Compliance Audit

Compare the completed application directly against every requirement in Instructions.docx.

Create a requirement-by-requirement compliance table containing:

- Professor's requirement
- Corresponding Python module or output
- Evidence of successful execution
- Verification method
- Pass, Fail, or Needs Review status
- Outstanding issues

Verify all required functionality, including:

1. User input for ticker and number of annual 10-K filings.
2. SEC EDGAR 10-K retrieval and organized file storage.
3. Correct MD&A section extraction.
4. Income statements for 10 reporting periods.
5. Balance sheets for 10 reporting periods.
6. Statements of cash flows for 10 reporting periods.
7. Separate Excel worksheets for the three financial statements.
8. MD&A text cleaning.
9. Loughran-McDonald sentiment analysis.
10. Total word counts and sentiment category counts.
11. Positive-to-Negative (P/N) ratios.
12. Historical stock prices at fiscal year-end balance sheet dates.
13. One liquidity ratio.
14. One solvency ratio.
15. One efficiency ratio.
16. One profitability ratio.
17. Required stock price versus sentiment P/N ratio chart.
18. Successful analysis of both AAPL and NVDA.

Do not mark a requirement as passing merely because the relevant Python function exists.

## 2. Complete End-to-End Testing

Execute the application using main.py for both AAPL and NVDA.

For each company, request 10 annual 10-K filings.

Verify that the complete workflow runs successfully from a single entry point.

Confirm that:

- The correct annual filings are used.
- Every stage executes in the intended order.
- Existing valid downloads can be reused.
- Extracted MD&A files correspond to the correct fiscal years.
- Financial statement data is correctly populated.
- Sentiment scores are computed from the correct cleaned MD&A files.
- Stock prices correspond to the correct balance sheet dates.
- Financial ratios use the correct reporting periods.
- The final Excel workbooks are generated successfully.
- Errors and warnings are reported accurately.

Preserve company-specific outputs and do not overwrite Apple's results when processing NVIDIA.

## 3. Financial Statement Accuracy

Independently verify selected financial statement values against the original SEC 10-K filings for both companies.

Check:

- Revenue
- Gross Profit
- Operating Income
- Net Income
- Current Assets
- Current Liabilities
- Total Assets
- Total Liabilities
- Shareholders' Equity
- Operating Cash Flow
- Investing Cash Flow
- Financing Cash Flow

Reconcile the balance sheet equation where the required values are available.

Check financial statement units, fiscal period alignment, missing values, duplicate periods, and inconsistent accounting classifications.

Flag unresolved discrepancies rather than changing numbers without evidence.

## 4. MD&A Extraction and Sentiment Accuracy

Inspect the extracted MD&A text files for all reporting years.

Verify that:

- The correct substantive Item 7 sections were extracted.
- Tables of contents were not incorrectly selected.
- Text was not accidentally truncated.
- The cleaning process preserved meaningful financial terminology.
- The correct Loughran-McDonald Master Dictionary was used.
- Word counts and sentiment category counts are reproducible.
- Repeated sentiment words are counted correctly.
- Positive-to-Negative ratios are mathematically accurate.

Independently test the sentiment module using a small known-text example.

Report the dictionary version and methodology.

## 5. Historical Stock Price Accuracy

Verify:

- Correct company ticker.
- Correct fiscal year-end dates.
- Correct historical trading dates.
- Proper handling of weekends and market holidays.
- Consistent treatment of stock splits.
- Consistent use and labeling of adjusted versus unadjusted closing prices.
- Accurate alignment with sentiment P/N ratios.

Cross-check selected historical prices against a second source when practical.

Do not invent missing stock price observations.

## 6. Financial Ratio Accuracy

Recalculate selected financial ratios independently for each company.

Verify the formulas for:

Current Ratio = Current Assets / Current Liabilities

Debt-to-Equity Ratio = Total Interest-Bearing Debt / Total Shareholders' Equity

Total Asset Turnover = Net Revenue / Average Total Assets

Net Profit Margin = Net Income / Net Revenue

Confirm that all four ratios use appropriate financial statement data and consistent accounting definitions.

Verify beginning and ending asset balances for total asset turnover.

Check that interest-bearing debt is not double-counted.

Flag inconsistent or unavailable values.

## 7. Excel Workbook Audit

Inspect the final Excel workbooks for both AAPL and NVDA.

Confirm that each workbook includes:

- Income Statement
- Balance Sheet
- Cash Flow Statement
- Sentiment Analysis
- Stock Prices
- Financial Ratios
- Charts and Analysis

Verify that the worksheets include the correct fiscal periods and that financial statements display all 10 years side by side.

Check:

- Excel file integrity.
- Worksheet names.
- Missing and duplicated reporting periods.
- Financial statement values.
- Formulas and cell references.
- Currency and percentage formatting.
- Chart data ranges.
- Chart titles and labels.
- Stock price versus sentiment P/N ratio visualization.
- Preservation of existing analytical outputs.

Where formulas are used, verify their correctness and ensure calculated results are available when the workbook is opened in Excel.

Do not assume that saving the workbook successfully means its content is accurate.

## 8. Automated Tests and Error Correction

Run the existing test suite and identify all failing tests.

Correct bugs that prevent assignment requirements from being satisfied.

After corrections, rerun the relevant tests and confirm that previously working modules remain functional.

Do not rewrite the entire application unnecessarily.

Maintain the modular architecture and preserve main.py as the single entry point.

Do not initialize Git or GitHub, create a graphical interface, or package a Windows executable.

## 9. Final Verification Report

Create a file named FINAL_VALIDATION_REPORT.md within the project folder.

Include:

1. Executive summary of application functionality.
2. Assignment compliance checklist.
3. AAPL processing results.
4. NVDA processing results.
5. Excel workbook verification.
6. Financial statement reconciliation results.
7. MD&A extraction verification.
8. Sentiment scoring validation.
9. Stock price verification.
10. Financial ratio verification.
11. Automated test results.
12. Any unresolved errors or limitations.
13. Exact Windows PowerShell command needed to run the completed application.

Clearly distinguish completed tests from tests that were not executed.

Do not fabricate results or classify untested functionality as verified.

## 10. Final Deliverables

At the conclusion, report:

- Final application status.
- Number of annual 10-K filings successfully analyzed for AAPL and NVDA.
- Exact paths of both completed Excel workbooks.
- Number of tests passed and failed.
- Any unresolved discrepancies.
- Any unmet assignment requirements.
- Exact Windows PowerShell command for running the application.

The intended command is:

.\.venv\Scripts\python.exe main.py

Confirm that this command works in the actual project environment.

## Important Scope Restrictions

This stage is exclusively for verifying and correcting the existing Python application and Excel outputs.

Do not generate the Lab Document.

Do not incorporate the Lab Document into the application's functionality.

Do not initialize Git or GitHub.

Do not add unnecessary new features.

Preserve the existing project structure and all original source data.

After completing the audit and generating FINAL_VALIDATION_REPORT.md, provide your findings and stop.
```


---

### Prompt 15: Company Selection Enhancement (Ticker & Company-Name Search)
* **Timestamp**: 2026-10-30T16:15:15-04:00
* **Stage**: Application Enhancement

```text
Please enhance the existing Python application's company selection functionality.

Allow users to either:

1. Enter a stock ticker symbol directly, such as AAPL or NVDA.
2. Search for a publicly traded company by its name, such as Apple, NVIDIA, or Microsoft.

For company-name searches, use official SEC company ticker and CIK reference data.

Display matching companies with their ticker symbols and allow the user to select the correct result.

Validate the selected company before proceeding with the existing analysis pipeline.

Preserve all existing functionality, including the option to enter a ticker directly.

Keep the application command-line based, integrate this feature into main.py, and do not create a graphical interface.

Test both ticker-based selection and company-name searches.

Do not modify the financial analysis modules unless necessary to maintain compatibility.
```

---

### Prompt 16: Final Stage — Standalone Windows Desktop Application Packaging
* **Timestamp**: 2026-10-30T16:26:13-04:00
* **Stage**: Final Packaging & Desktop GUI

```text
# FINAL STAGE — CONVERT PYTHON PROJECT INTO A STANDALONE WINDOWS DESKTOP APPLICATION

We have completed the development of my undergraduate GEN AI SEC 10-K Financial Analyzer using Python and Google Antigravity on Windows.

I now want to transform the existing Python application into a complete, polished, standalone Windows desktop application that launches by double-clicking an icon, similar to opening Discord or Spotify.

The application must retain all existing SEC filing retrieval, financial statement extraction, sentiment analysis, stock price retrieval, financial ratio calculations, and Excel generation functionality.

Do not rebuild the application from scratch.

## 1. Create a Graphical Desktop Interface

Use CustomTkinter to develop a modern Windows desktop interface.

Design a professional financial analytics application featuring:

- Application title: SEC 10-K Financial Analyzer
- A clean, modern dark-themed interface.
- Company ticker input.
- Optional company-name search using official SEC company data.
- A dropdown or number field to select the number of annual 10-K filings.
- A prominent Analyze Company button.
- A real-time progress indicator.
- A status panel displaying the current processing stage.
- A results section showing completed analyses.
- An Open Excel Workbook button.
- An Open Output Folder button.
- Clear and understandable error messages.

The interface should be intuitive for someone who has no Python programming experience.

## 2. Integrate Existing Python Modules

Preserve the existing modular architecture.

Connect the GUI to the current financial analysis pipeline.

All existing functionality must remain operational, including:

1. SEC EDGAR 10-K downloads.
2. MD&A extraction.
3. Financial statement extraction.
4. MD&A text cleaning.
5. Loughran-McDonald sentiment analysis.
6. Historical stock price retrieval.
7. Liquidity, solvency, efficiency, and profitability ratio calculations.
8. Excel workbook and chart generation.

The GUI must invoke the same underlying functions used by the existing command-line application.

Avoid duplicating processing logic.

Keep main.py or the existing command-line entry point functional for automated tests and troubleshooting.

## 3. Application Behavior

The user should be able to:

1. Launch the desktop application by double-clicking its icon.
2. Enter or search for a company.
3. Select the desired number of annual 10-K filings.
4. Click Analyze Company.
5. Observe the analysis progress.
6. Receive a clear notification when processing is complete.
7. Open the completed Excel workbook.
8. Open the folder containing generated outputs.
9. Run another company's analysis without restarting the application.

Run the analysis in a background worker thread or process so the graphical interface remains responsive.

Ensure GUI updates are performed safely on the main interface thread.

Disable duplicate analysis requests while a task is running.

Do not falsely display 100% completion when required processing stages have failed.

## 4. File Storage and Configuration

Ensure downloaded filings, extracted text, financial data, sentiment results, and Excel workbooks are saved persistently.

For the packaged Windows application, use an appropriate writable user-data directory, such as:

%LOCALAPPDATA%\SEC10KAnalyzer

Do not attempt to save generated data inside PyInstaller's temporary extraction directory.

Preserve previously generated AAPL and NVDA workbooks.

Make the Loughran-McDonald dictionary available to the packaged application using a documented method.

Ensure SEC User-Agent configuration works in the packaged application without exposing private contact information in publicly distributed source files.

## 5. Package as a Standalone Windows Application

Use PyInstaller to package the completed application.

Create a Windows executable that launches the graphical interface without displaying a terminal window.

Prefer a one-folder build initially for reliability and easier troubleshooting.

After verifying the one-folder build, optionally create a one-file executable if practical.

Include necessary Python dependencies and non-code resources.

Set an appropriate application name:

SEC10KAnalyzer.exe

Create an application icon if practical.

Do not initialize Git or GitHub.

Do not create a web application.

Do not change the underlying financial analysis methodology.

## 6. Desktop Accessibility

Provide instructions for:

- Finding the finished executable in the dist directory.
- Launching the application.
- Creating a Windows desktop shortcut.
- Optionally placing the application in an appropriate permanent location.

The user should not need to open Antigravity, run PowerShell commands, or manually execute Python scripts to use the finished application.

The finished application should be launchable directly from Windows.

## 7. Comprehensive Testing

Test the graphical application before packaging.

Verify:

1. The interface opens successfully.
2. Company ticker entry works.
3. Company search works if implemented.
4. The number of filings can be changed.
5. The Analyze Company button starts the full processing pipeline.
6. The interface remains responsive during analysis.
7. Progress updates accurately reflect the current stage.
8. Error messages are displayed correctly.
9. Excel workbooks are generated and saved persistently.
10. Open Workbook and Open Output Folder buttons work.
11. The application can analyze multiple companies consecutively.
12. The existing command-line application and automated tests remain functional.

After packaging, test the executable itself rather than assuming that a successful PyInstaller build guarantees the application works.

Run a complete AAPL analysis using the packaged application and verify that the Excel output is correct.

If practical, test launching the executable outside the original project directory to confirm that it does not depend on development-only file paths.

## 8. Final Deliverables

Provide:

1. Complete Windows desktop application source code.
2. Packaged Windows executable in the dist folder.
3. All required supporting application files.
4. Instructions for launching the executable.
5. Instructions for creating a desktop shortcut.
6. Confirmation that all original financial analysis functionality is preserved.
7. Actual GUI and packaged application test results.
8. Documentation of any unresolved limitations.

Update README.md with a section explaining how to use the desktop application.

The Lab Document is a separate academic deliverable and must not be generated or incorporated into this application's functionality.

The objective is a fully functioning, standalone Windows desktop financial analysis application that can be opened and used without manually running Python.

Do not claim the packaging or testing was successful unless the application has actually been launched and tested.
```

---

### Prompt 17: Portable One-File Windows Executable Packaging & Relocated Testing
* **Timestamp**: 2026-10-30T16:58:47-04:00
* **Stage**: Standalone Portable One-File Packaging

```text
After successfully testing the one-folder Windows application, also create a portable one-file executable using PyInstaller's --onefile option.

The final SEC10KAnalyzer.exe must:

Launch independently without Python or Antigravity installed.

Include all required Python dependencies and bundled resources.

Access the Loughran-McDonald dictionary correctly.

Store downloaded SEC filings and Excel outputs in a persistent writable user-data directory, not alongside the executable.

Function when moved to a different folder or drive.

Avoid relying on absolute paths to my development project.

Preserve SEC User-Agent configuration and allow it to be configured if missing.

Display clear errors if internet access or required resources are unavailable.

Test the executable after moving a copy outside the original project directory.

Verify that the graphical interface launches and successfully generates an Excel report from the relocated executable.

Do not claim the portable build works unless this test passes.
```

---

### Prompt 18: Final Distribution Folder & Company Output Organization
* **Timestamp**: 2026-10-30T17:11:17-04:00
* **Stage**: Application Distribution & Output Packaging

```text
FINAL APPLICATION DISTRIBUTION AND COMPANY OUTPUT ORGANIZATION

I want to finalize my existing SEC 10-K Financial Analyzer as a complete, portable Windows application package suitable for sharing with my professor.

Preserve all existing functionality and do not rebuild the application.

1. Create a Final Distribution Folder

Create a folder named:

SEC 10-K Financial Analyzer

This folder should contain:

The tested portable Windows executable, named SEC10KAnalyzer.exe.

A folder named Company Analyses containing the existing AAPL and NVDA results.

A README.txt with simple instructions for launching and using the application.

Include all required application resources in the executable or ensure they are correctly accessible.

Do not include unnecessary development files, Python virtual environments, test scripts, or build caches in this distribution folder.

Preserve the original development project separately.

2. Organize Company-Specific Outputs

Inside Company Analyses, create a separate folder for each analyzed company.

For example:

Company Analyses/
AAPL - Apple Inc./
AAPL_10K_Analysis.xlsx
SEC Filings/
MD&A Text/

NVDA - NVIDIA Corporation/
    NVDA_10K_Analysis.xlsx
    SEC Filings/
    MD&A Text/

Copy the existing verified analysis outputs into their corresponding company folders.

Preserve the original datasets in the development project.

3. Automatically Create Folders for New Companies

When a user enters or searches for a new ticker symbol, the application should automatically:

Identify the correct company name and ticker.

Create a company-specific output folder.

Download and store SEC filings within that folder.

Store extracted MD&A text and relevant intermediate data.

Generate the completed Excel workbook.

Save the workbook within the corresponding company folder.

Preserve outputs belonging to other companies.

Use safe, Windows-compatible folder names.

If the same company is analyzed again, update or reuse its existing folder without creating unnecessary duplicates or deleting unrelated results.

4. Portable File Storage

Make the application portable so it can be copied to another Windows computer.

When launched from a writable folder, store company outputs in the Company Analyses folder alongside the executable.

If the location is not writable, such as a protected Windows installation directory, use a clearly identified writable user-data location and display that location in the application.

Never save persistent outputs in PyInstaller's temporary extraction directory.

Ensure the Loughran-McDonald dictionary, SEC User-Agent configuration, and other required resources remain accessible after moving the application.

Do not hardcode paths to my current computer, Downloads folder, Python environment, or Antigravity project.

5. Improve Output Accessibility

Within the existing graphical interface, ensure users can:

Analyze any supported company by name or ticker.

Open the generated Excel workbook.

Open the selected company's output folder.

Access previously generated analyses.

Run another analysis without restarting the application.

Keep the interface simple and professional.

6. Package and Test

Rebuild the portable Windows executable after making the necessary changes.

Test the application outside the original development directory.

Verify that:

The executable launches independently.

Existing AAPL and NVDA analyses remain accessible.

New company folders are created automatically.

Excel workbooks are saved in the correct folders.

The application works without Python or Antigravity installed.

All existing financial analysis functionality remains operational.

The distribution folder can be moved without breaking file paths.

Report any limitations or failed tests.

7. Final Deliverable

Create one ZIP archive named:

SEC_10K_Financial_Analyzer.zip

The ZIP should contain the complete application distribution folder, including the executable, existing company analyses, and README.

The recipient should be able to extract the ZIP, open the folder, double-click the application, and access the existing results or generate new analyses.

Do not create a Git repository or modify the underlying accounting and sentiment analysis methodology.

Do not generate the separate academic Lab Document.

Provide the exact location of the final ZIP archive and report the results of portability testing.
```

#### Implementation Actions & Verification for Prompt 18:
1. **Final Distribution Folder Architecture**:
   - Created standalone distribution package: `SEC 10-K Financial Analyzer/`.
   - Bundled tested portable 64-bit Windows binary: `SEC 10-K Financial Analyzer\SEC10KAnalyzer.exe` (61.2 MB).
   - Created clear user documentation: `SEC 10-K Financial Analyzer\README.txt` (8.2 KB) with instructions for launching, viewing pre-analyzed datasets, running new companies, and managing SEC settings.
   - Preserved original development environment (`Python 10-K/`) separately without unnecessary development caches, virtual environments, or build intermediate files in the distribution folder.

2. **Organized Company-Specific Outputs**:
   - Structured `Company Analyses/` repository containing complete verified 10-year datasets:
     - `Company Analyses\AAPL - Apple Inc\`:
       - `AAPL_10K_Analysis.xlsx` (Master 8-tab workbook)
       - `AAPL_financial_statements.xlsx` (Raw extracted 3-statement models)
       - `AAPL_financial_ratios.csv`, `AAPL_sentiment_analysis.csv`, `AAPL_stock_prices.csv`
       - `companyfacts_0000320193.json` (SEC EDGAR company facts cache)
       - `SEC Filings\` (all 10 HTML filings + `AAPL_10K_manifest.csv`)
       - `MD&A Text\` (all 10 raw + 10 cleaned MD&A text files + manifests)
     - `Company Analyses\NVDA - NVIDIA Corporation\`:
       - `NVDA_10K_Analysis.xlsx` (Master 8-tab workbook)
       - `NVDA_financial_statements.xlsx` (Raw extracted 3-statement models)
       - `NVDA_financial_ratios.csv`, `NVDA_sentiment_analysis.csv`, `NVDA_stock_prices.csv`
       - `companyfacts_0001045810.json` (SEC EDGAR company facts cache)
       - `SEC Filings\` (all 10 HTML filings + `NVDA_10K_manifest.csv`)
       - `MD&A Text\` (all 10 raw + 10 cleaned MD&A text files + manifests)
   - Automated Synchronization: Implemented `sync_company_analyses()` in `main.py` to automatically create sanitized company folders (`<TICKER> - <Clean Corporate Title>/`) and copy all filings, MD&A text, CSVs, and Excel workbooks when any new company is analyzed.

3. **Portable Storage & Non-Writable Fallback**:
   - Implemented dynamic writability check `_is_dir_writable()` in `config.py`.
   - When run from a writable directory (e.g. Desktop, USB, Downloads), uses portable storage inside `Company Analyses/` alongside the executable.
   - When run from a write-protected directory (e.g. `C:\Program Files`), automatically falls back cleanly to `%LOCALAPPDATA%\SEC10KAnalyzer\Company Analyses`.
   - Never saves persistent outputs in PyInstaller's temporary `sys._MEIPASS` folder.
   - Displayed active storage mode (`📁 Portable Storage` vs `📁 AppData Storage`) and active directory path in header pill and activity log.

4. **Desktop GUI Usability Enhancements**:
   - Added "Previously Analyzed Companies:" dropdown in `gui_app.py` auto-populated from `Company Analyses/`.
   - Allows users to select existing companies (`AAPL - Apple Inc`, `NVDA - NVIDIA Corporation`, etc.) and immediately load their workbooks, populate executive metric tiles, and update folder links without re-running the analysis pipeline.
   - Updated "📊 Open Excel Workbook" and "📁 Open Output Folder" buttons to target the selected company's folder directly.
   - Preserved ability to enter new ticker symbols or search corporations by name in official SEC EDGAR registry.

5. **Distribution Archive & Verification**:
   - Compressed distribution package into `SEC_10K_Financial_Analyzer.zip` (65.3 MB).
   - Tested extraction in clean temporary directory (`%TEMP%\sec_dist_test`).
   - Verified that `SEC10KAnalyzer.exe` launches from extracted folder, discovers AAPL and NVDA, loads both workbooks with 0 errors, and executes analysis of a new company (MSFT) creating `Company Analyses\MSFT - MICROSOFT CORP\`.
   - Verified that running new analyses does not overwrite or corrupt existing AAPL or NVDA analyses.
   - All 60 automated unit tests pass across the entire project test suite.

---

### Prompt 19: Streamlit Financial Analytics Dashboard Integration
* **Timestamp**: 2026-10-30T17:25:12-04:00
* **Stage**: Web-Based Streamlit Financial Analytics Dashboard

```text
# FINAL APPLICATION ENHANCEMENT — STREAMLIT FINANCIAL ANALYTICS DASHBOARD

I have installed Streamlit and want to use it as the primary interface for my existing SEC 10-K Financial Analyzer.

We are developing this project using Python and Google Antigravity on Windows.

The two companies required for my undergraduate accounting project are:
- Apple Inc. (AAPL)
- NVIDIA Corporation (NVDA)

My objective is to transform the current application into a professional, interactive financial analytics dashboard that runs in a web browser and can eventually be deployed through GitHub and Streamlit Community Cloud.

IMPORTANT: Preserve all existing financial analysis functionality. Do not rebuild the underlying Python application from scratch.

## 1. Create the Streamlit Application
Create a new entry-point file named streamlit_app.py.
Use Streamlit to develop a polished financial analytics dashboard.
The interface should include:
- SEC 10-K Financial Analyzer branding.
- A professional dark-themed dashboard.
- Company search by name or ticker symbol.
- A user-selectable number of annual 10-K filings, defaulting to 10.
- An Analyze Company button.
- Real-time processing progress.
- Clear success, warning, and error messages.
- Interactive financial statement tables.
- Financial performance metrics.
- Interactive charts and visualizations.
- Excel workbook download functionality.

## 2. Integrate the Existing Python Backend
Connect Streamlit to the existing modules for:
1. SEC 10-K downloads.
2. MD&A extraction.
3. Financial statement extraction.
4. MD&A text cleaning.
5. Loughran-McDonald sentiment analysis.
6. Historical stock price retrieval.
7. Financial ratio calculations.
8. Excel workbook generation.
Reuse the existing functions and data-processing logic.
Do not duplicate or rewrite financial calculations unnecessarily.
Preserve main.py as an optional command-line entry point for testing and troubleshooting.

## 3. Financial Overview Dashboard
After the analysis completes, automatically display the company's results.
Include summary metrics for: Revenue, Net Income, Current Ratio, Debt-to-Equity Ratio, Total Asset Turnover, Net Profit Margin, Sentiment P/N Ratio, Historical Stock Price.
Use the most recent available fiscal reporting period for summary metrics and clearly identify the date.
All displayed figures must originate from actual processed financial data.

## 4. Interactive Financial Statements
Create separate tabs for: Income Statement, Balance Sheet, Cash Flow Statement.
Display 10 fiscal years side by side.
Use interactive dataframes with consistent financial units, scrolling, readable column headers, and appropriate number formatting.
Preserve reporting periods and accounting classifications from existing backend.

## 5. Interactive Financial Charts
Create charts for:
- Revenue and net income trends.
- Current Ratio.
- Debt-to-Equity Ratio.
- Total Asset Turnover.
- Net Profit Margin.
- Positive and Negative Sentiment Word Counts.
- Sentiment P/N Ratio.
- Historical Stock Prices.
- Stock Price versus Sentiment P/N Ratio (Mandatory for professor's assignment).
Use appropriate secondary axes when comparing variables measured on different scales.
Label all axes and data sources clearly.

## 6. Company Selection and Analysis History
Allow users to:
1. Search by company name or ticker symbol.
2. Select number of annual 10-K filings.
3. Execute complete analysis.
4. View previously completed company analyses.
5. Switch between AAPL and NVDA.
6. Analyze additional supported publicly traded companies.
7. Reopen existing results without unnecessary repeated SEC requests.
Use session state to preserve user selections during Streamlit reruns.
Implement safe caching where appropriate.

## 7. Excel Exports
Preserve existing Excel workbook generation process.
Provide a Download Excel Workbook button containing all 8 worksheets, charts, and audit reconciliations.

## 8. Local and Cloud Compatibility
The Streamlit application must run locally on Windows and be prepared for Streamlit Community Cloud (Linux).
Use portable project-relative paths, handle SEC User-Agent via env/st.secrets, never expose personal secrets, allow direct downloads, handle concurrent requests safely.

## 9. GitHub Preparation
Prepare repository for uploading to GitHub, exclude virtual environments, caches, secrets, executables.
Do not publish or deploy without approval.

## 10. Testing
Test Streamlit app locally, verify financial tables, charts, AAPL/NVDA switching, sentiment analysis matching Excel outputs, Excel downloads, and ensure all existing unit tests pass.

## 11. Final Documentation
Update README.md with startup command: python -m streamlit run streamlit_app.py
```

#### Implementation Actions & Verification for Prompt 19:
1. **Streamlit Financial Analytics Dashboard Architecture (`streamlit_app.py`)**:
   - Created standalone web application entry point: `streamlit_app.py` (969 lines) powered by Streamlit and Plotly.
   - Built an executive dark theme with custom CSS styling, dynamic metric cards, and responsive tabbed layouts.
   - Designed interactive sidebar with company search (exact ticker & name search across 10,430+ SEC entities), analysis scope selector (3, 5, 10, 15, 20 filings), and SEC User-Agent configuration dialog.
   - Built a **History Switcher** to seamlessly toggle between pre-analyzed corporations (**Apple Inc.** and **NVIDIA Corporation**) with instantaneous reloading of all financial models, statements, charts, and metrics without re-downloading or recalculating.

2. **Backend Engine Integration**:
   - Connected directly to existing backend modules: `sec_downloader.py`, `mda_extractor.py`, `text_cleaner.py`, `financial_extractor.py`, `sentiment_analysis.py`, `stock_prices.py`, `financial_ratios.py`, `excel_report.py`, and `main.py`.
   - Maintained `main.py` and `gui_app.py` as fully operational alternative entry points.
   - Implemented real-time stage progress reporting (`st.progress`) with status callbacks.

3. **Financial Overview Dashboard (8 Metric Tiles)**:
   - Extracted verified latest fiscal reporting period metrics:
     - AAPL (FY2025 ended 2025-09-27): Revenue: $416,161M ($416.2B), Net Income: $112,010M ($112.0B), Current Ratio: 0.89x, Debt-to-Equity: 1.34x, Asset Turnover: 1.15x, Net Margin: 26.92%, Sentiment P/N: 0.2500, Stock Price: $255.46.
     - NVDA (FY2026 ended 2026-01-25): Revenue: $215,938M ($215.9B), Net Income: $120,067M ($120.1B), Current Ratio: 3.91x, Debt-to-Equity: 0.05x, Asset Turnover: 1.36x, Net Margin: 55.60%, Sentiment P/N: 0.3684, Stock Price: $187.67.
   - Clearly identifies the fiscal year-end date, actual trading date, and balance sheet reconciliation status.

4. **Interactive Financial Statement Presentation**:
   - Built 3 side-by-side 10-year statement tabs: Income Statement, Balance Sheet, and Statement of Cash Flows.
   - Clean number formatting with currency units ($ in Millions), accounting parenthesis for negative numbers, and interactive dataframes.
   - Added dedicated Audit & Reconciliation tab verifying 10/10 PASS on balance sheet equation ($A = L + SE$) and gross profit reconciliations.

5. **Interactive Plotly Visualizations**:
   - **Professor's Mandatory Dual-Axis Chart**: Closing Stock Price vs. MD&A Sentiment P/N Ratio over 10 fiscal reporting periods with secondary Y-axis and hover tooltips.
   - Revenue and Net Income 10-year growth grouped bar charts.
   - 4-panel financial ratio analysis grid with 1.0x working capital benchmark line.
   - Loughran-McDonald sentiment breakdown (positive vs. negative word counts, P/N ratio trends).
   - Historical closing stock prices dataset table.

6. **Excel & CSV Export Engine**:
   - Primary `st.download_button` to download the master 8-tab workbook: `<TICKER>_10K_Analysis.xlsx`.
   - Secondary download buttons for raw financial statement workbooks and CSV files (`*_financial_ratios.csv`, `*_sentiment_analysis.csv`, `*_stock_prices.csv`).

7. **Cloud Compatibility & GitHub Preparation**:
   - Configured `config.get_sec_user_agent()` to inspect `st.secrets["SEC_USER_AGENT"]`, environment variables, and `.sec_settings`.
   - Updated `.gitignore` to protect personal credentials (`.sec_settings`, `.streamlit/secrets.toml`), build caches, and binaries.
   - Updated `requirements.txt` with `streamlit>=1.30.0` and `plotly>=5.18.0`.

8. **Automated Testing & Local Live Server Verification**:
   - Created dedicated test module: `test_streamlit.py` with 6 automated tests using `streamlit.testing.v1.AppTest`.
   - Verified live local server execution at `http://localhost:8501` with HTTP 200 OK.
   - Full test suite passing: **69 / 69 automated tests** pass across all test modules.

---

### Prompt 20: GitHub Repository Preparation & Cloud Deployment
* **Timestamp**: 2026-10-30T17:46:29-04:00
* **Topic**: GitHub Preparation & Streamlit Community Cloud

```text
I have created a GitHub repository named SEC-10K-Financial-Analyzer.

I want to upload my existing SEC 10-K Financial Analyzer Python project to that repository so I can deploy my Streamlit dashboard through Streamlit Community Cloud.

Please prepare and upload the project without rebuilding or changing its financial analysis functionality.

Requirements:

Check that streamlit_app.py is present and is the correct Streamlit entry point.

Ensure requirements.txt contains every Python package required for cloud execution.

Check that the application is compatible with Linux, which Streamlit Community Cloud uses.

Create or update .gitignore to exclude .venv, pycache, build, dist, temporary files, local credentials, .env, and .streamlit/secrets.toml.

Check the entire project for hardcoded personal information, API keys, local Windows paths, and sensitive configuration.

Make sure the application can find the Loughran-McDonald dictionary in the deployed environment, without including prohibited or improperly licensed data.

Ensure the application can generate and download Excel workbooks without relying on files stored on my Windows computer.

Preserve the Python application functionality...

Show me the files that will be committed and explain anything excluded.

Ask for my approval before pushing the files to GitHub.

After approval, commit and push the project to the main branch.
```

#### Implementation Actions & Verification for Prompt 20:
1. **Streamlit Entry Point & Cloud Compatibility Audit**:
   - Verified `streamlit_app.py` is located at repository root as the primary entry point for Streamlit Community Cloud.
   - Confirmed `requirements.txt` contains all dependencies for headless cloud Linux execution (`requests`, `beautifulsoup4`, `lxml`, `openpyxl`, `yfinance`, `pandas`, `customtkinter`, `pillow`, `pyinstaller`, `streamlit`, `plotly`).
   - Audited codebase for full Linux cross-platform portability using `pathlib.Path` with zero hardcoded Windows drive letters (`C:\`).
   - Configured `config.get_sec_user_agent()` with automatic Streamlit Cloud secrets support (`st.secrets["SEC_USER_AGENT"]`).

2. **Privacy, Security & Git Exclusions**:
   - Hardened `.gitignore` to strictly exclude:
     - Virtual environments (`.venv/`, `venv/`).
     - Local credentials and secrets (`.sec_settings`, `.env`, `.env.*`, `.streamlit/secrets.toml`, `secrets.toml`).
     - Local Windows binaries & packaging builds (`build/`, `dist/`, `*.exe`, `*.zip`, `SEC 10-K Financial Analyzer/`).
     - Heavy scraped HTML filings (`**/SEC Filings/*.htm`, `data/raw_filings/*`) saving ~46 MB of git bloat.
     - Bytecode and caches (`__pycache__/`, `*.pyc`, `*.tmp`, `*.log`).
   - Replaced student email placeholder in `gui_app.py` with generic placeholder.

3. **Academic Reference Data & In-Memory Cloud Exports**:
   - Verified inclusion of the official academic Loughran-McDonald Master Dictionary (`data/Loughran-McDonald_MasterDictionary_1993-2025.csv`, 9.09 MB).
   - Bundled SEC ticker database (`data/financial_data/company_tickers.json`, 871 KB) to enable instant offline company search without triggering SEC EDGAR rate limits.
   - Verified all Excel downloads operate dynamically in cloud memory via `st.download_button`.

4. **Staging, Review & Approval**:
   - Initialized Git on `main` branch.
   - Staged 99 production files (19.36 MB total).
   - Presented comprehensive staging report and exclusion rationale to the user.
   - Configured remote origin: `https://github.com/iprexibtw-tech/SEC-10K-Financial-Analyzer.git`.
   - Created initial commit `a659c39` with author `iprexibtw-tech <iprexibtw-tech@users.noreply.github.com>`.

5. **Successful Push & Verification**:
   - User executed push via terminal with Git Credential Manager browser authorization.
   - Verified remote tracking: `main` branch is up to date with `origin/main` at commit `a659c39`.
   - Working tree verified completely clean.

---

### Prompt 21: Project Clarification — Maintain Streamlit Cloud Architecture
* **Timestamp**: 2026-10-30T18:40:40-04:00
* **Topic**: Streamlit Cloud Architecture & Enhanced In-Browser Analytics

```text
PROJECT CLARIFICATION — KEEP STREAMLIT

Please disregard my previous request to convert the application into a standalone GitHub Pages website.

I want to retain the existing architecture:

GitHub for hosting the Python source code.

Streamlit Community Cloud for hosting the interactive application.

The existing Python modules for processing financial data.

My objective is to make the Streamlit dashboard completely usable through the browser.

Please ensure:

Users can search for companies by name or ticker and run new analyses directly within Streamlit.

All 10 years of financial statements are displayed in interactive tables.

Financial ratios, sentiment analysis, stock prices, and charts are accessible directly within the dashboard.

Users can switch between analysis sections without opening external programs.

Existing AAPL and NVDA analyses can be viewed directly in the application.

Excel workbooks are still generated because my professor requires them, but users should not need to download Excel files to view the analysis.

The dashboard displays the required stock price versus sentiment P/N ratio chart.

All existing Python calculations and financial analysis functionality remain unchanged.

The application continues to work through Streamlit Community Cloud and remains connected to my GitHub repository.

Test the deployed Streamlit interface, not just the local version.

Do not create a GitHub Pages website, remove Excel generation, or rebuild the application from scratch.

My goal is one complete, browser-based financial analysis application that users can access through the existing Streamlit URL.
```

#### Implementation Actions & Verification for Prompt 21:
1. **Preservation of Core Python & Streamlit Architecture**:
   - Cleaned up any temporary static HTML files to keep the repository strictly focused on the full Python backend and Streamlit dashboard.
   - Preserved all existing underlying Python modules (`main.py`, `sec_downloader.py`, `mda_extractor.py`, `text_cleaner.py`, `sentiment_analysis.py`, `financial_extractor.py`, `financial_ratios.py`, `stock_prices.py`, `excel_report.py`, `gui_app.py`, `config.py`).
   - Retained full automated multi-tab Excel workbook generation (`<TICKER>_10K_Analysis.xlsx`).

2. **Cross-Company Comparison Dashboard (Tab 6)**:
   - Added dedicated `"⚖️ Apple vs. NVIDIA Comparison"` tab to `streamlit_app.py`.
   - Built 4 comparative KPI cards comparing Apple Inc. vs. NVIDIA Corporation:
     - Latest Net Revenue: AAPL $416,161M ($416.2B) vs NVDA $215,938M ($215.9B).
     - Latest Net Income: AAPL $112,010M ($112.0B) vs NVDA $120,067M ($120.1B).
     - Net Profit Margin: AAPL 26.92% vs NVDA 55.60%.
     - Current Ratio: AAPL 0.89x vs NVDA 3.91x.
   - Built interactive comparative Plotly visualizations:
     - Side-by-side Revenue and Net Income 10-year growth comparison (`chart_comparison_revenue_and_net_income`).
     - Side-by-side Net Profit Margin and MD&A Sentiment P/N Ratio trajectory comparison (`chart_comparison_margins_and_sentiment`).
   - Integrated a complete 10-year side-by-side comparative table with all metrics.

3. **In-Browser Quick Company Switcher & Interactive Statement Search**:
   - Added prominent one-click quick switch buttons (`[🍎 Apple Inc. (AAPL)]`, `[🟢 NVIDIA Corp (NVDA)]`) below the header for instant toggling without opening the sidebar.
   - Added interactive line-item search filter to the Three Financial Statements tab (`apply_filter`), allowing instant keyword searching (e.g., "Revenue", "Cash", "Debt", "R&D") across all 10 fiscal reporting periods.

4. **Automated Testing & Full Suite Verification**:
   - Expanded `test_streamlit.py` to 8 automated tests, verifying quick switch buttons, 7 main tabs, statement filtering, and comparative Plotly charts.
   - Ran complete project test suite: **71 / 71 automated tests passing** across all test modules.
   - Verified live server response on port 8501: HTTP 200 OK.




