# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller specification file for portable single-file SEC 10-K Financial Analyzer.

Packages the CustomTkinter GUI application, underlying financial extraction
and analysis modules, and required data files into a single, standalone, portable
executable: SEC10KAnalyzer.exe.
"""

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

block_cipher = None

# Collect CustomTkinter assets, themes, and submodules
ctk_datas, ctk_binaries, ctk_hiddenimports = collect_all('customtkinter')

# Application bundled assets
project_datas = [
    ('app_icon.ico', '.'),
    ('app_icon.png', '.'),
    ('.sec_settings', '.'),
    ('Company Analyses', 'Company Analyses'),
    ('data/Loughran-McDonald_MasterDictionary_1993-2025.csv', 'data'),
    ('data/financial_data/company_tickers.json', 'data/financial_data'),
    ('data/financial_data/companyfacts_0000320193.json', 'data/financial_data'),
    ('data/financial_data/companyfacts_0001045810.json', 'data/financial_data'),
    ('data/financial_data/AAPL_financial_ratios.csv', 'data/financial_data'),
    ('data/financial_data/AAPL_sentiment_analysis.csv', 'data/financial_data'),
    ('data/financial_data/AAPL_stock_prices.csv', 'data/financial_data'),
    ('data/financial_data/NVDA_financial_ratios.csv', 'data/financial_data'),
    ('data/financial_data/NVDA_sentiment_analysis.csv', 'data/financial_data'),
    ('data/financial_data/NVDA_stock_prices.csv', 'data/financial_data'),
    ('data/excel_outputs/AAPL_10K_Analysis.xlsx', 'data/excel_outputs'),
    ('data/excel_outputs/NVDA_10K_Analysis.xlsx', 'data/excel_outputs'),
    ('data/excel_outputs/AAPL', 'data/excel_outputs/AAPL'),
    ('data/excel_outputs/NVDA', 'data/excel_outputs/NVDA'),
]

all_datas = project_datas + ctk_datas

hidden_imports = [
    'config',
    'sec_downloader',
    'mda_extractor',
    'financial_extractor',
    'text_cleaner',
    'sentiment_analysis',
    'stock_prices',
    'financial_ratios',
    'excel_report',
    'main',
    'gui_app',
    'openpyxl',
    'openpyxl.chart',
    'openpyxl.styles',
    'openpyxl.utils',
    'pandas',
    'yfinance',
    'multitasking',
    'PIL',
    'PIL._tkinter_finder',
    'darkdetect',
] + ctk_hiddenimports

a = Analysis(
    ['gui_app.py'],
    pathex=['.'],
    binaries=ctk_binaries,
    datas=all_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter.test', 'unittest', 'pytest'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SEC10KAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
)
