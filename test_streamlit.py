"""Unit and Integration Tests for Streamlit Financial Analytics Dashboard.

Validates the Streamlit web dashboard (streamlit_app.py), session state transitions,
Plotly chart generation, data loaders, statement formatting, and Excel downloads.
"""

import io
import unittest
import openpyxl
import pandas as pd
from streamlit.testing.v1 import AppTest

import config
from streamlit_app import (
    chart_comparison_margins_and_sentiment,
    chart_comparison_revenue_and_net_income,
    chart_ratio_trends,
    chart_revenue_and_net_income,
    chart_sentiment_breakdown,
    chart_stock_price_vs_sentiment,
    format_statement_display,
    load_company_dataset,
)


class TestStreamlitDashboard(unittest.TestCase):
    """Test suite verifying Streamlit application rendering, state, and outputs."""

    def test_01_app_default_aapl_rendering(self) -> None:
        """Verify dashboard starts with AAPL, zero exceptions, and renders 8 metrics."""
        at = AppTest.from_file("streamlit_app.py", default_timeout=30)
        at.run()
        self.assertEqual(len(at.exception), 0, f"App execution raised exceptions: {at.exception}")

        # Verify metrics count and content
        self.assertGreaterEqual(len(at.metric), 8, f"Expected at least 8 metrics, found {len(at.metric)}")
        labels = [m.label for m in at.metric]
        self.assertIn("Annual Net Revenue", labels)
        self.assertIn("Annual Net Income", labels)
        self.assertIn("Current Ratio (Liquidity)", labels)
        self.assertIn("Debt-to-Equity (Solvency)", labels)
        self.assertIn("Total Asset Turnover", labels)
        self.assertIn("Net Profit Margin", labels)
        self.assertIn("Sentiment P/N Ratio", labels)
        self.assertIn("Closing Stock Price", labels)

        # Verify AAPL metric values
        metric_dict = {m.label: m.value for m in at.metric}
        self.assertEqual(metric_dict["Annual Net Revenue"], "$416,161M")
        self.assertEqual(metric_dict["Annual Net Income"], "$112,010M")
        self.assertEqual(metric_dict["Current Ratio (Liquidity)"], "0.89x")
        self.assertEqual(metric_dict["Debt-to-Equity (Solvency)"], "1.34x")
        self.assertEqual(metric_dict["Total Asset Turnover"], "1.15x")
        self.assertEqual(metric_dict["Net Profit Margin"], "26.92%")
        self.assertEqual(metric_dict["Sentiment P/N Ratio"], "0.2500")
        self.assertEqual(metric_dict["Closing Stock Price"], "$255.46")

    def test_02_app_switch_to_nvda(self) -> None:
        """Verify switching session state to NVDA loads NVDA results without errors."""
        at = AppTest.from_file("streamlit_app.py", default_timeout=30)
        at.run()
        at.session_state["current_ticker"] = "NVDA"
        at.run()
        self.assertEqual(len(at.exception), 0, f"NVDA switch raised exceptions: {at.exception}")

        metric_dict = {m.label: m.value for m in at.metric}
        self.assertEqual(metric_dict["Annual Net Revenue"], "$215,938M")
        self.assertEqual(metric_dict["Annual Net Income"], "$120,067M")
        self.assertEqual(metric_dict["Current Ratio (Liquidity)"], "3.91x")
        self.assertEqual(metric_dict["Debt-to-Equity (Solvency)"], "0.05x")
        self.assertEqual(metric_dict["Total Asset Turnover"], "1.36x")
        self.assertEqual(metric_dict["Net Profit Margin"], "55.60%")
        self.assertEqual(metric_dict["Sentiment P/N Ratio"], "0.3684")
        self.assertEqual(metric_dict["Closing Stock Price"], "$187.67")

    def test_03_load_company_dataset(self) -> None:
        """Verify load_company_dataset extracts valid datasets and Excel bytes for AAPL and NVDA."""
        for ticker in ["AAPL", "NVDA"]:
            data = load_company_dataset(ticker)
            self.assertIsNotNone(data, f"load_company_dataset returned None for {ticker}")
            self.assertEqual(data["ticker"], ticker)
            self.assertEqual(data["ratios_df"].shape[0], 10)
            self.assertEqual(data["sentiment_df"].shape[0], 10)
            self.assertEqual(data["prices_df"].shape[0], 10)
            self.assertEqual(data["income_df"].shape[0], 14)
            self.assertEqual(data["balance_df"].shape[0], 29)
            self.assertEqual(data["cashflow_df"].shape[0], 22)

            # Test Excel workbook bytes integrity
            wb_bytes = data["wb_bytes"]
            self.assertGreater(len(wb_bytes), 20000, "Workbook bytes too small")
            wb = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=True)
            self.assertEqual(len(wb.sheetnames), 8)
            self.assertIn("Income Statement", wb.sheetnames)
            self.assertIn("Charts and Analysis", wb.sheetnames)

    def test_04_dual_axis_stock_sentiment_chart(self) -> None:
        """Verify Professor's mandatory dual-axis chart has two traces and secondary axis."""
        data = load_company_dataset("AAPL")
        self.assertIsNotNone(data)
        fig = chart_stock_price_vs_sentiment(data["prices_df"], data["sentiment_df"], "AAPL")

        self.assertEqual(len(fig.data), 2, "Expected 2 traces on dual-axis chart")
        self.assertEqual(fig.data[0].name, "Stock Price ($ USD)")
        self.assertEqual(fig.data[1].name, "Sentiment P/N Ratio")
        self.assertIn("Longitudinal Correlation", fig.layout.title.text)

    def test_05_additional_plotly_charts(self) -> None:
        """Verify auxiliary charts (Revenue/NI, Ratios 4-panel, Sentiment breakdown)."""
        data = load_company_dataset("NVDA")
        self.assertIsNotNone(data)

        # Revenue & NI
        fig_rev = chart_revenue_and_net_income(data["ratios_df"], "NVDA")
        self.assertEqual(len(fig_rev.data), 2)

        # Ratio Trends (4 subplots)
        fig_ratios = chart_ratio_trends(data["ratios_df"], "NVDA")
        self.assertEqual(len(fig_ratios.data), 4)

        # Sentiment breakdown
        fig_sent = chart_sentiment_breakdown(data["sentiment_df"], "NVDA")
        self.assertEqual(len(fig_sent.data), 2)

    def test_06_statement_formatting(self) -> None:
        """Verify format_statement_display preserves dates and formats dollar values with $ and M."""
        data = load_company_dataset("AAPL")
        self.assertIsNotNone(data)
        formatted = format_statement_display(data["income_df"])
        self.assertEqual(formatted.shape, data["income_df"].shape)

        # First row is Fiscal Period-End Date
        self.assertEqual(formatted.iloc[0]["FY 2025"], "2025-09-27")
        # Revenue row has $416,161M
        self.assertEqual(formatted.iloc[1]["FY 2025"], "$416,161M")

    def test_07_comparison_charts(self) -> None:
        """Verify comparative charts generate valid multi-series traces for AAPL vs NVDA."""
        aapl = load_company_dataset("AAPL")
        nvda = load_company_dataset("NVDA")
        self.assertIsNotNone(aapl)
        self.assertIsNotNone(nvda)

        # Revenue & NI comparison
        fig_rev = chart_comparison_revenue_and_net_income(aapl["ratios_df"], nvda["ratios_df"])
        self.assertEqual(len(fig_rev.data), 4, "Expected 4 traces in comparative rev/income chart")
        names = [t.name for t in fig_rev.data]
        self.assertIn("Apple (AAPL) Revenue", names)
        self.assertIn("NVIDIA (NVDA) Revenue", names)

        # Margin & sentiment comparison
        fig_margins = chart_comparison_margins_and_sentiment(
            aapl["ratios_df"], nvda["ratios_df"],
            aapl["sentiment_df"], nvda["sentiment_df"]
        )
        self.assertEqual(len(fig_margins.data), 4, "Expected 4 traces in margins/sentiment chart")
        names = [t.name for t in fig_margins.data]
        self.assertIn("AAPL Margin %", names)
        self.assertIn("NVDA Margin %", names)
        self.assertIn("AAPL Sentiment P/N", names)
        self.assertIn("NVDA Sentiment P/N", names)

    def test_08_app_quick_switch_and_tabs(self) -> None:
        """Verify quick switch buttons, 7 main tabs, and statement filtering in AppTest."""
        at = AppTest.from_file("streamlit_app.py", default_timeout=30)
        at.run()
        self.assertEqual(len(at.exception), 0)

        # Check tabs exist (7 tabs)
        self.assertGreaterEqual(len(at.tabs), 7, f"Expected at least 7 tabs, found {len(at.tabs)}")

        # Check quick switch buttons
        btn_labels = [b.label for b in at.button]
        self.assertTrue(any("Apple" in b for b in btn_labels), "Expected Apple quick switch button")
        self.assertTrue(any("NVIDIA" in b for b in btn_labels), "Expected NVIDIA quick switch button")

    def test_09_custom_ticker_selection_msft(self) -> None:
        """Verify selecting a custom ticker (MSFT) preserves selection and renders data."""
        at = AppTest.from_file("streamlit_app.py", default_timeout=30)
        at.run()
        self.assertEqual(len(at.exception), 0)

        # Switch to custom ticker MSFT
        at.session_state["selected_ticker"] = "MSFT"
        at.session_state["current_ticker"] = "MSFT"
        at.session_state["custom_ticker_input"] = "MSFT"
        at.run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["selected_ticker"], "MSFT")

        # Verify MSFT metrics
        metric_dict = {m.label: m.value for m in at.metric}
        self.assertEqual(metric_dict["Annual Net Revenue"], "$331,839M")
        self.assertEqual(metric_dict["Annual Net Income"], "$133,749M")
        self.assertEqual(metric_dict["Current Ratio (Liquidity)"], "1.23x")

    def test_10_invalid_ticker_handling(self) -> None:
        """Verify entering an invalid ticker displays clear error without crashing or falling back."""
        at = AppTest.from_file("streamlit_app.py", default_timeout=30)
        at.run()
        self.assertEqual(len(at.exception), 0)

        # Set invalid ticker
        at.session_state["selected_ticker"] = "INVALID123"
        at.session_state["current_ticker"] = "INVALID123"
        at.session_state["custom_ticker_input"] = "INVALID123"
        at.run()
        self.assertEqual(len(at.exception), 0)

        # Verify clear error message is displayed
        error_texts = [e.value for e in at.error]
        self.assertTrue(any("Invalid Ticker Symbol" in t for t in error_texts))
        self.assertTrue(any("INVALID123" in t for t in error_texts))

        # Verify no financial statement metrics are displayed under the invalid ticker
        self.assertEqual(len(at.metric), 0)

    def test_11_multi_company_switching_and_refresh(self) -> None:
        """Verify seamless switching between AAPL, NVDA, and MSFT with state preservation."""
        at = AppTest.from_file("streamlit_app.py", default_timeout=30)
        at.run()

        # Step 1: Default is AAPL
        metric_dict = {m.label: m.value for m in at.metric}
        self.assertEqual(metric_dict["Annual Net Revenue"], "$416,161M")

        # Step 2: Switch to NVDA
        at.session_state["selected_ticker"] = "NVDA"
        at.session_state["current_ticker"] = "NVDA"
        at.session_state["custom_ticker_input"] = "NVDA"
        at.run()
        metric_dict = {m.label: m.value for m in at.metric}
        self.assertEqual(metric_dict["Annual Net Revenue"], "$215,938M")

        # Step 3: Switch to MSFT
        at.session_state["selected_ticker"] = "MSFT"
        at.session_state["current_ticker"] = "MSFT"
        at.session_state["custom_ticker_input"] = "MSFT"
        at.run()
        metric_dict = {m.label: m.value for m in at.metric}
        self.assertEqual(metric_dict["Annual Net Revenue"], "$331,839M")

        # Step 4: Refresh (rerun without changing state)
        at.run()
        self.assertEqual(at.session_state["selected_ticker"], "MSFT")
        metric_dict = {m.label: m.value for m in at.metric}
        self.assertEqual(metric_dict["Annual Net Revenue"], "$331,839M")

        # Step 5: Switch back to AAPL
        at.session_state["selected_ticker"] = "AAPL"
        at.session_state["current_ticker"] = "AAPL"
        at.session_state["custom_ticker_input"] = "AAPL"
        at.run()
        metric_dict = {m.label: m.value for m in at.metric}
        self.assertEqual(metric_dict["Annual Net Revenue"], "$416,161M")

    def test_12_unanalyzed_valid_company_prompt(self) -> None:
        """Verify unanalyzed valid ticker (e.g. TSLA) prompts analysis without crashing."""
        at = AppTest.from_file("streamlit_app.py", default_timeout=30)
        at.run()

        at.session_state["selected_ticker"] = "TSLA"
        at.session_state["current_ticker"] = "TSLA"
        at.session_state["custom_ticker_input"] = "TSLA"
        at.run()
        self.assertEqual(len(at.exception), 0)

        # Verify no statement metrics yet
        self.assertEqual(len(at.metric), 0)

        # Verify info prompt exists
        info_texts = [i.value for i in at.info]
        self.assertTrue(any("No existing analytical models found" in t for t in info_texts))
        self.assertTrue(any("TSLA" in t for t in info_texts))

    def test_13_amd_custom_ticker_regression_workflow(self) -> None:
        """Verify the exact user workflow (AAPL -> AMD -> NVDA -> AMD -> MSFT) without session-state exceptions."""
        at = AppTest.from_file("streamlit_app.py", default_timeout=30)
        at.run()
        self.assertEqual(len(at.exception), 0)

        # 1. AAPL initially loaded
        m_aapl = {m.label: m.value for m in at.metric}
        self.assertEqual(m_aapl.get("Annual Net Revenue"), "$416,161M")

        # 2. Switch to AMD
        at.session_state["active_ticker"] = "AMD"
        at.run()
        self.assertEqual(len(at.exception), 0, f"AMD transition raised: {at.exception}")
        m_amd = {m.label: m.value for m in at.metric}
        self.assertEqual(m_amd.get("Annual Net Revenue"), "$34,639M")

        # 3. Switch to NVDA
        at.session_state["active_ticker"] = "NVDA"
        at.run()
        self.assertEqual(len(at.exception), 0)
        m_nvda = {m.label: m.value for m in at.metric}
        self.assertEqual(m_nvda.get("Annual Net Revenue"), "$215,938M")

        # 4. Return to AMD from history
        at.session_state["active_ticker"] = "AMD"
        at.run()
        self.assertEqual(len(at.exception), 0)
        m_amd2 = {m.label: m.value for m in at.metric}
        self.assertEqual(m_amd2.get("Annual Net Revenue"), "$34,639M")

        # 5. Switch to MSFT
        at.session_state["active_ticker"] = "MSFT"
        at.run()
        self.assertEqual(len(at.exception), 0)
        m_msft = {m.label: m.value for m in at.metric}
        self.assertEqual(m_msft.get("Annual Net Revenue"), "$331,839M")


if __name__ == "__main__":
    unittest.main()
