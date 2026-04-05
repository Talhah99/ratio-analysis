import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np
from typing import List, Optional, Dict
import html
import logging

logger = logging.getLogger("ratio_engine")

# ---------------------------------------------------------------------------
# Metric classification helpers (used by multiple methods)
# ---------------------------------------------------------------------------
_PERCENTAGE_METRICS = {
    "Gross Margin", "Operating Margin", "Net Margin", "EBITDA Margin",
    "Pretax Margin", "ROIC", "ROE", "ROA", "ROE (Normalized)", "Cash ROIC",
    "OCF to Sales", "FCF to Sales",
    # Quality of Income (OCF/Net Income) and FCF Conversion (FCF/EBITDA) are
    # multiples (e.g. 1.20×), NOT percentages. Removed from this set so
    # _format_value displays them as "1.20" not "120.00%".
    "FCF Yield", "Dividend Payout",
    "Revenue Growth", "Gross Profit Growth", "Operating Income Growth",
    "EBITDA Growth", "Net Income Growth", "EPS Growth", "FCF Growth",
    "DuPont: Net Margin", "NWC to Assets", "Debt to Assets", "Debt to Capital",
}

_LOWER_IS_BETTER = {
    "Debt to Equity", "Debt to Assets", "Debt to Capital",
    "P/E Ratio", "PEG Ratio", "Price to Book (P/B)", "Price to Sales",
    "EV / EBITDA", "EV / Revenue",
    "Net Debt to EBITDA", "Financial Leverage",
    "Days Sales Outstanding (DSO)", "Days Inventory (DIO)",
    "Days Payables (DPO)", "Cash Conversion Cycle",
    "Dividend Payout",
}


class DashboardGenerator:

    def __init__(self, engine):
        """
        Initialize with a validated RatioEngine instance.

        Args:
            engine: RatioEngine with calculated results.

        Raises:
            ValueError: If engine has no results or no year columns.
        """
        if not hasattr(engine, 'results') or not engine.results:
            raise ValueError("Engine has no results. Run calculation first.")

        self.engine = engine
        self.results = engine.results
        self.years = engine.years

        if not self.years:
            raise ValueError("No years found in engine")

    # ------------------------------------------------------------------
    # Low-level helpers
    # ------------------------------------------------------------------

    def _clean_for_plot(self, values) -> List:
        """Convert numpy array / list to a Plotly-safe list (None for NaN/Inf)."""
        if values is None:
            return []
        if not isinstance(values, np.ndarray):
            try:
                values = np.array(values, dtype=float)
            except (ValueError, TypeError):
                return []
        if values.size == 0:
            return []
        result = []
        for val in values:
            if np.isnan(val) or np.isinf(val):
                result.append(None)
            else:
                try:
                    result.append(float(val))
                except (ValueError, TypeError):
                    result.append(None)
        return result

    def _get_metric(self, data: dict, metric_name: str) -> List:
        raw = data.get(metric_name, np.array([]))
        return self._clean_for_plot(raw)

    def _has_data(self, values: List) -> bool:
        return any(v is not None for v in values)

    def _format_value(self, value, metric_name: str) -> str:
        """Format a scalar consistently according to metric type."""
        if value is None or (isinstance(value, float) and (np.isnan(value) or np.isinf(value))):
            return "N/A"
        try:
            f = float(value)
            if metric_name in _PERCENTAGE_METRICS:
                return f"{f:.2%}"
            return f"{f:,.2f}"
        except (ValueError, TypeError):
            return "N/A"

    def _format_percentage(self, value) -> str:
        if value is None or pd.isna(value):
            return "N/A"
        try:
            return f"{float(value):.2%}"
        except (ValueError, TypeError):
            return "N/A"

    def _format_number(self, value, decimals: int = 2) -> str:
        if value is None or pd.isna(value):
            return "N/A"
        try:
            return f"{float(value):.{decimals}f}"
        except (ValueError, TypeError):
            return "N/A"

    def _trend_direction(self, values: List, n: int = 3) -> str:
        """
        Return a human-readable trend label based on the last n valid values.
        Uses linear regression slope so a single outlier period does not mislead.

        Normalises the slope against the data range (max - min) rather than
        the first-point value, which is unstable when ys[0] ≈ 0 (e.g. a company
        that had near-zero profit in the earliest period).
        """
        valid = [(i, v) for i, v in enumerate(values) if v is not None]
        if len(valid) < 2:
            return "→ Insufficient data"
        subset = valid[-n:]
        xs = np.array([p[0] for p in subset], dtype=float)
        ys = np.array([p[1] for p in subset], dtype=float)
        if len(xs) < 2:
            return "→ Stable"
        slope = np.polyfit(xs, ys, 1)[0]
        # Normalise against data range to avoid instability near zero.
        data_range = np.ptp(ys)  # max - min
        if data_range < 1e-10:
            return "→ Stable"
        relative_change = abs(slope) / data_range
        if relative_change < 0.05:   # slope < 5% of the value range per period
            return "→ Stable"
        return "↑ Improving" if slope > 0 else "↓ Declining"

    def _get_insight_emoji(self, level: str) -> str:
        return {
            "excellent": "✅", "strong": "✅", "good": "✅",
            "moderate": "⚠️", "weak": "⚠️", "concerning": "⚠️",
            "critical": "🔴", "poor": "🔴"
        }.get(level, "ℹ️")

    # ------------------------------------------------------------------
    # Chart utility
    # ------------------------------------------------------------------

    def _no_data_fig(self, title: str, msg: str, height: int = 400) -> go.Figure:
        """Blank figure with a centred 'no data' annotation."""
        fig = go.Figure()
        fig.add_annotation(
            text=msg, xref="paper", yref="paper",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color="gray")
        )
        fig.update_layout(title=title, template="plotly_white", height=height)
        return fig

    def _std_legend(self) -> dict:
        return dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)

    # ------------------------------------------------------------------
    # Chart builders
    # ------------------------------------------------------------------

    # 1. Profitability Margins
    def _create_margin_chart(self, data: dict) -> go.Figure:
        fig = go.Figure()
        metrics = [
            ("Gross Margin",     "#1f77b4"),
            ("EBITDA Margin",    "#ff7f0e"),
            ("Operating Margin", "#2ca02c"),
            ("Net Margin",       "#d62728"),
        ]
        has_data = False
        for metric, color in metrics:
            values = self._get_metric(data, metric)
            if self._has_data(values):
                has_data = True
                fig.add_trace(go.Scatter(
                    x=self.years, y=values, name=metric,
                    mode='lines+markers',
                    line=dict(width=2, color=color), marker=dict(size=8),
                    connectgaps=False,
                    hovertemplate='%{y:.1%}<extra></extra>'
                ))
        if not has_data:
            return self._no_data_fig(
                "Profitability Margins",
                "No margin data available<br>Check Income Statement field mappings"
            )
        fig.update_layout(
            title="Profitability Margins Over Time", template="plotly_white",
            height=400, xaxis_title="Year", xaxis_type='category',
            yaxis_title="Margin", yaxis_tickformat='.1%',
            hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 2. Return Ratios
    def _create_returns_chart(self, data: dict) -> go.Figure:
        fig = go.Figure()
        metrics = ["ROIC", "ROE", "ROA", "ROE (Normalized)"]
        has_data = False
        for metric in metrics:
            values = self._get_metric(data, metric)
            if self._has_data(values):
                has_data = True
                fig.add_trace(go.Bar(
                    x=self.years, y=values, name=metric,
                    text=[self._format_percentage(v) for v in values],
                    textposition='outside',
                    hovertemplate='%{y:.1%}<extra></extra>'
                ))
        if not has_data:
            return self._no_data_fig(
                "Return on Investment Metrics",
                "No return data available<br>Check Balance Sheet field mappings"
            )
        fig.add_hline(
            y=0.10, line_dash="dot", line_color="navy", line_width=1,
            annotation_text="10% Hurdle Rate", annotation_position="right"
        )
        fig.update_layout(
            title="Return on Investment Metrics", barmode='group',
            template="plotly_white", height=400,
            xaxis_title="Year", xaxis_type='category',
            yaxis_title="Return", yaxis_tickformat='.0%',
            hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 3. DuPont Decomposition (NEW)
    def _create_dupont_chart(self, data: dict) -> go.Figure:
        """
        DuPont: ROE = Net Margin x Asset Turnover x Equity Multiplier.
        Bars (primary axis) = multiples (AT, EM).
        Lines (secondary axis) = rates (Net Margin %, reported ROE %).
        Overlaying reported ROE lets the analyst immediately see if the
        product of the three components reconciles to the reported figure.
        """
        nm_vals  = self._get_metric(data, "DuPont: Net Margin")
        at_vals  = self._get_metric(data, "DuPont: Asset Turnover")
        em_vals  = self._get_metric(data, "DuPont: Equity Multiplier")
        roe_vals = self._get_metric(data, "ROE")

        has_dupont = any(self._has_data(v) for v in [nm_vals, at_vals, em_vals])
        if not has_dupont:
            return self._no_data_fig(
                "DuPont ROE Decomposition",
                "DuPont data not available<br>Map Net Income, Revenue, Total Assets, Total Equity"
            )

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        if self._has_data(nm_vals):
            fig.add_trace(go.Scatter(
                x=self.years, y=nm_vals, name="Net Margin",
                mode='lines+markers', line=dict(width=2, color='#e74c3c', dash='dot'),
                marker=dict(size=7), connectgaps=False,
                hovertemplate='Net Margin: %{y:.1%}<extra></extra>'
            ), secondary_y=True)

        if self._has_data(roe_vals):
            fig.add_trace(go.Scatter(
                x=self.years, y=roe_vals, name="ROE (Reported)",
                mode='lines+markers', line=dict(width=2.5, color='#8e44ad'),
                marker=dict(size=9, symbol='diamond'), connectgaps=False,
                hovertemplate='ROE: %{y:.1%}<extra></extra>'
            ), secondary_y=True)

        if self._has_data(at_vals):
            fig.add_trace(go.Bar(
                x=self.years, y=at_vals, name="Asset Turnover (×)",
                marker_color='rgba(52,152,219,0.7)',
                hovertemplate='Asset Turnover: %{y:.2f}×<extra></extra>'
            ), secondary_y=False)

        if self._has_data(em_vals):
            fig.add_trace(go.Bar(
                x=self.years, y=em_vals, name="Equity Multiplier (×)",
                marker_color='rgba(46,204,113,0.7)',
                hovertemplate='Equity Multiplier: %{y:.2f}×<extra></extra>'
            ), secondary_y=False)

        fig.update_xaxes(title_text="Year", type='category')
        fig.update_yaxes(title_text="Multiple (×)", secondary_y=False)
        fig.update_yaxes(title_text="Rate (%)", tickformat='.0%', secondary_y=True)

        fig.update_layout(
            title="DuPont ROE Decomposition (Net Margin × Asset Turnover × Equity Multiplier)",
            template="plotly_white", height=420, barmode='group',
            hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 4. Valuation Multiples
    def _create_valuation_chart(self, data: dict) -> go.Figure:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        pe_values        = self._get_metric(data, "P/E Ratio")
        ev_ebitda_values = self._get_metric(data, "EV / EBITDA")
        pb_values        = self._get_metric(data, "Price to Book (P/B)")

        has_data = False
        if self._has_data(pe_values):
            has_data = True
            fig.add_trace(go.Scatter(
                x=self.years, y=pe_values, name="P/E Ratio",
                mode='lines+markers', line=dict(width=2, color='#1f77b4'),
                marker=dict(size=8), connectgaps=False,
                hovertemplate='P/E: %{y:.1f}×<extra></extra>'
            ), secondary_y=False)

        if self._has_data(ev_ebitda_values):
            has_data = True
            fig.add_trace(go.Scatter(
                x=self.years, y=ev_ebitda_values, name="EV / EBITDA",
                mode='lines+markers', line=dict(width=2, dash='dot', color='#ff7f0e'),
                marker=dict(size=8), connectgaps=False,
                hovertemplate='EV/EBITDA: %{y:.1f}×<extra></extra>'
            ), secondary_y=False)

        if self._has_data(pb_values):
            has_data = True
            fig.add_trace(go.Scatter(
                x=self.years, y=pb_values, name="Price / Book",
                mode='lines+markers', line=dict(width=2, color='#9b59b6'),
                marker=dict(size=8), connectgaps=False,
                hovertemplate='P/B: %{y:.2f}×<extra></extra>'
            ), secondary_y=True)

        if not has_data:
            return self._no_data_fig(
                "Valuation Multiples",
                "No valuation data available<br>Check Market Data field mappings"
            )

        fig.update_xaxes(title_text="Year", type='category')
        fig.update_yaxes(title_text="P/E and EV/EBITDA (×)", secondary_y=False)
        fig.update_yaxes(title_text="Price / Book (×)", secondary_y=True)
        fig.update_layout(
            title="Valuation Multiples", template="plotly_white",
            height=400, hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 5. Liquidity
    def _create_liquidity_chart(self, data: dict) -> go.Figure:
        fig = go.Figure()
        metrics = [
            ("Current Ratio", "#1f77b4", 'solid'),
            ("Quick Ratio",   "#ff7f0e", 'dash'),
            ("Cash Ratio",    "#2ca02c", 'dot'),
        ]
        has_data = False
        for metric, color, dash in metrics:
            values = self._get_metric(data, metric)
            if self._has_data(values):
                has_data = True
                fig.add_trace(go.Scatter(
                    x=self.years, y=values, name=metric,
                    mode='lines+markers',
                    line=dict(width=2, color=color, dash=dash),
                    marker=dict(size=8), connectgaps=False,
                    hovertemplate=f'{metric}: %{{y:.2f}}×<extra></extra>'
                ))
        if not has_data:
            return self._no_data_fig(
                "Liquidity Analysis",
                "No liquidity data available<br>Check Current Assets / Current Liabilities mappings"
            )
        fig.add_hrect(y0=0, y1=1.0,  fillcolor="red",    opacity=0.05, layer="below", line_width=0)
        fig.add_hrect(y0=1.0, y1=1.5, fillcolor="orange", opacity=0.05, layer="below", line_width=0)
        fig.add_hline(y=1.0, line_dash="dot", line_color="red",    line_width=1,
                      annotation_text="Minimum (1.0×)", annotation_position="right")
        fig.add_hline(y=1.5, line_dash="dot", line_color="orange", line_width=1,
                      annotation_text="Sound (1.5×)", annotation_position="right")
        fig.update_layout(
            title="Liquidity Analysis", template="plotly_white",
            height=400, xaxis_title="Year", xaxis_type='category',
            yaxis_title="Ratio (×)", hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 6. Leverage & Solvency
    def _create_leverage_chart(self, data: dict) -> go.Figure:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        de_values  = self._get_metric(data, "Debt to Equity")
        da_values  = self._get_metric(data, "Debt to Assets")
        nd_ebitda  = self._get_metric(data, "Net Debt to EBITDA")
        ic_values  = self._get_metric(data, "Interest Coverage")

        has_data = False

        for values, label, color, dash in [
            (de_values, "Debt / Equity",     '#e74c3c', 'solid'),
            (da_values, "Debt / Assets",     '#f39c12', 'dash'),
            (nd_ebitda, "Net Debt / EBITDA", '#c0392b', 'dot'),
        ]:
            if self._has_data(values):
                has_data = True
                fig.add_trace(go.Scatter(
                    x=self.years, y=values, name=label,
                    mode='lines+markers',
                    line=dict(width=2, color=color, dash=dash),
                    marker=dict(size=8), connectgaps=False,
                    hovertemplate=f'{label}: %{{y:.2f}}×<extra></extra>'
                ), secondary_y=False)

        if self._has_data(ic_values):
            has_data = True
            fig.add_trace(go.Scatter(
                x=self.years, y=ic_values, name="Interest Coverage",
                mode='lines+markers', line=dict(width=2, color='#27ae60'),
                marker=dict(size=8), connectgaps=False,
                hovertemplate='Int. Cov: %{y:.1f}×<extra></extra>'
            ), secondary_y=True)

        if not has_data:
            return self._no_data_fig(
                "Leverage & Solvency Metrics",
                "No leverage data available<br>Check Debt and Interest Expense field mappings"
            )

        fig.add_hline(y=1.0, line_dash="dot", line_color="orange", line_width=1,
                      annotation_text="D/E = 1.0 (Elevated)", annotation_position="right",
                      secondary_y=False)
        # CFA standard minimum Interest Coverage = 3.0x
        fig.add_hline(y=3.0, line_dash="dot", line_color="red", line_width=1,
                      annotation_text="Int. Cov ≥ 3.0× (CFA min)", annotation_position="right",
                      secondary_y=True)

        fig.update_xaxes(title_text="Year", type='category')
        fig.update_yaxes(title_text="Debt Ratios (×)", secondary_y=False)
        fig.update_yaxes(title_text="Interest Coverage (×)", secondary_y=True)
        fig.update_layout(
            title="Leverage & Solvency Metrics", template="plotly_white",
            height=400, hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 7. Operating Efficiency
    def _create_efficiency_chart(self, data: dict) -> go.Figure:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        day_traces = [
            (self._get_metric(data, "Days Sales Outstanding (DSO)"), "DSO", '#3498db', 'solid'),
            (self._get_metric(data, "Days Inventory (DIO)"),         "DIO", '#9b59b6', 'dash'),
            (self._get_metric(data, "Days Payables (DPO)"),          "DPO", '#16a085', 'dot'),
            (self._get_metric(data, "Cash Conversion Cycle"),        "CCC", '#e67e22', 'dashdot'),
        ]
        at_values = self._get_metric(data, "Total Asset Turnover")
        has_data = False

        for values, label, color, dash in day_traces:
            if self._has_data(values):
                has_data = True
                fig.add_trace(go.Scatter(
                    x=self.years, y=values, name=f"{label} (days)",
                    mode='lines+markers',
                    line=dict(width=2, color=color, dash=dash),
                    marker=dict(size=8), connectgaps=False,
                    hovertemplate=f'{label}: %{{y:.0f}} days<extra></extra>'
                ), secondary_y=False)

        if self._has_data(at_values):
            has_data = True
            fig.add_trace(go.Scatter(
                x=self.years, y=at_values, name="Asset Turnover (×)",
                mode='lines+markers', line=dict(width=2, color='#e67e22'),
                marker=dict(size=8), connectgaps=False,
                hovertemplate='Asset Turnover: %{y:.2f}×<extra></extra>'
            ), secondary_y=True)

        if not has_data:
            return self._no_data_fig(
                "Operating Efficiency Metrics",
                "No efficiency data available<br>Check Receivables and Inventory field mappings"
            )

        fig.update_xaxes(title_text="Year", type='category')
        fig.update_yaxes(title_text="Days", secondary_y=False)
        fig.update_yaxes(title_text="Asset Turnover (×)", secondary_y=True)
        fig.update_layout(
            title="Operating Efficiency — Working Capital Cycle",
            template="plotly_white", height=400,
            hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 8. Cash Flow Quality (NEW)
    def _create_cashflow_chart(self, data: dict) -> go.Figure:
        """
        Cash flow panel: OCF/Sales and FCF/Sales (left axis, %), plus
        Quality of Income and Capex Coverage (right axis, multiples).

        Quality of Income = OCF / Net Income. A value >= 1.0 means every
        dollar of reported earnings is fully backed by operating cash flow.
        Values persistently below 0.6 are a red flag for accruals management.

        Capex Coverage = OCF / CapEx. Values >= 1.5x indicate the business
        can fund its capex entirely from operations with headroom to spare.
        """
        fig = make_subplots(specs=[[{"secondary_y": True}]])

        ocf_sales = self._get_metric(data, "OCF to Sales")
        fcf_sales = self._get_metric(data, "FCF to Sales")
        qoi       = self._get_metric(data, "Quality of Income")
        capex_cov = self._get_metric(data, "Capex Coverage")

        has_data = False

        if self._has_data(ocf_sales):
            has_data = True
            fig.add_trace(go.Scatter(
                x=self.years, y=ocf_sales, name="OCF / Sales",
                mode='lines+markers', line=dict(width=2, color='#27ae60'),
                marker=dict(size=8), connectgaps=False,
                hovertemplate='OCF/Sales: %{y:.1%}<extra></extra>'
            ), secondary_y=False)

        if self._has_data(fcf_sales):
            has_data = True
            fig.add_trace(go.Scatter(
                x=self.years, y=fcf_sales, name="FCF / Sales",
                mode='lines+markers', line=dict(width=2, dash='dash', color='#2980b9'),
                marker=dict(size=8), connectgaps=False,
                hovertemplate='FCF/Sales: %{y:.1%}<extra></extra>'
            ), secondary_y=False)

        if self._has_data(qoi):
            has_data = True
            fig.add_trace(go.Scatter(
                x=self.years, y=qoi, name="Quality of Income (×)",
                mode='lines+markers', line=dict(width=2, color='#8e44ad'),
                marker=dict(size=8), connectgaps=False,
                hovertemplate='Quality of Income: %{y:.2f}×<extra></extra>'
            ), secondary_y=True)

        if self._has_data(capex_cov):
            has_data = True
            fig.add_trace(go.Bar(
                x=self.years, y=capex_cov, name="Capex Coverage (×)",
                marker_color='rgba(231,76,60,0.45)',
                hovertemplate='Capex Coverage: %{y:.2f}×<extra></extra>'
            ), secondary_y=True)

        if not has_data:
            return self._no_data_fig(
                "Cash Flow Quality",
                "No cash flow data available<br>Check Operating Cash Flow and CapEx field mappings"
            )

        # Reference lines
        fig.add_hline(y=1.0, line_dash="dot", line_color="purple", line_width=1,
                      annotation_text="QoI = 1.0 (full earnings backed by cash)",
                      annotation_position="right", secondary_y=True)
        fig.add_hline(y=0.10, line_dash="dot", line_color="green", line_width=1,
                      annotation_text="OCF/Sales ≥ 10% benchmark",
                      annotation_position="right", secondary_y=False)

        fig.update_xaxes(title_text="Year", type='category')
        fig.update_yaxes(title_text="OCF & FCF / Sales", tickformat='.0%', secondary_y=False)
        fig.update_yaxes(title_text="Coverage Multiples (×)", secondary_y=True)
        fig.update_layout(
            title="Cash Flow Quality & Coverage",
            template="plotly_white", height=420,
            hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 9. Growth Rates
    def _create_growth_chart(self, data: dict) -> go.Figure:
        fig = go.Figure()
        metrics = [
            ("Revenue Growth",         "#3498db"),
            ("EBITDA Growth",          "#2ecc71"),
            ("Operating Income Growth","#e67e22"),
            ("Net Income Growth",      "#e74c3c"),
            ("EPS Growth",             "#f39c12"),
        ]
        has_data = False
        for metric, color in metrics:
            values = self._get_metric(data, metric)
            if self._has_data(values):
                has_data = True
                fig.add_trace(go.Bar(
                    x=self.years, y=values, name=metric,
                    text=[self._format_percentage(v) if v is not None else "" for v in values],
                    textposition='outside',
                    marker=dict(color=color),
                    hovertemplate='%{y:.1%}<extra></extra>'
                ))
        if not has_data:
            return self._no_data_fig(
                "Year-over-Year Growth Rates",
                "No growth data available<br>Check Income Statement field mappings"
            )
        fig.add_hline(y=0.0, line_dash="dash", line_color="gray", line_width=1,
                      annotation_text="Zero Growth", annotation_position="right")
        fig.update_layout(
            title="Year-over-Year Growth Rates", barmode='group',
            template="plotly_white", height=400,
            xaxis_title="Year", xaxis_type='category',
            yaxis_title="Growth Rate", yaxis_tickformat='.0%',
            hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # 10. Bankruptcy Risk — BOTH Z-Scores
    def _create_risk_dashboard(self, data: dict) -> go.Figure:
        """
        Plots both the public 5-factor Z-Score and the 4-factor EM Score
        on one chart with their respective threshold bands and lines.

        Public model (Altman 1968): Safe > 2.99, Grey 1.81–2.99, Distress < 1.81.
        EM model (Altman 1995):     Safe > 2.60, Grey 1.10–2.60, Distress < 1.10.
        EM Score does not require market cap — it uses book equity.
        """
        fig = go.Figure()
        z_public = self._get_metric(data, "Altman Z-Score")
        z_em     = self._get_metric(data, "Altman Z-Score (EM Score)")

        has_data = self._has_data(z_public) or self._has_data(z_em)
        if not has_data:
            return self._no_data_fig(
                "Bankruptcy Risk Assessment",
                "Z-Score data not available<br>Public model requires Market Cap; EM model uses book equity"
            )

        # Background zones — use public model thresholds (conservative)
        fig.add_hrect(y0=-15, y1=1.81,  fillcolor="red",    opacity=0.07, layer="below", line_width=0,
                      annotation_text="Distress (Public <1.81)", annotation_position="top left",
                      annotation_font_color="#c0392b")
        fig.add_hrect(y0=1.81, y1=2.99, fillcolor="#FFD700", opacity=0.07, layer="below", line_width=0,
                      annotation_text="Grey Zone (Public 1.81–2.99)", annotation_position="top left",
                      annotation_font_color="goldenrod")
        fig.add_hrect(y0=2.99, y1=50,   fillcolor="green",  opacity=0.07, layer="below", line_width=0,
                      annotation_text="Safe Zone (Public >2.99)", annotation_position="top left",
                      annotation_font_color="#27ae60")

        if self._has_data(z_public):
            fig.add_trace(go.Scatter(
                x=self.years, y=z_public, name="Z-Score (Public 5-factor)",
                mode='lines+markers', line=dict(width=3, color='#2c3e50'),
                marker=dict(size=10), connectgaps=False,
                hovertemplate='Z-Score (Public): %{y:.2f}<extra></extra>'
            ))
            fig.add_hline(y=1.81, line_dash="dash", line_color="red",   line_width=1.5,
                          annotation_text="1.81 Distress", annotation_position="right")
            fig.add_hline(y=2.99, line_dash="dash", line_color="green", line_width=1.5,
                          annotation_text="2.99 Safe",     annotation_position="right")

        if self._has_data(z_em):
            fig.add_trace(go.Scatter(
                x=self.years, y=z_em, name="EM Score (4-factor, book equity)",
                mode='lines+markers', line=dict(width=2.5, color='#2980b9', dash='dot'),
                marker=dict(size=9, symbol='diamond'), connectgaps=False,
                hovertemplate='EM Score: %{y:.2f}<extra></extra>'
            ))
            fig.add_hline(y=1.10, line_dash="dot", line_color="#c0392b", line_width=1,
                          annotation_text="1.10 EM Distress", annotation_position="right")
            fig.add_hline(y=2.60, line_dash="dot", line_color="#1abc9c", line_width=1,
                          annotation_text="2.60 EM Safe",     annotation_position="right")

        all_vals = [v for lst in [z_public, z_em] for v in lst if v is not None]
        y_min = min(all_vals) - 0.5 if all_vals else -1
        y_max = max(all_vals) + 0.5 if all_vals else 5

        fig.update_layout(
            title="Bankruptcy Risk — Altman Z-Score (Public) & EM Score",
            template="plotly_white", height=440,
            xaxis_title="Year", xaxis_type='category',
            yaxis_title="Score",
            yaxis_range=[min(y_min, -1), max(y_max, 10)],
            hovermode='x unified', legend=self._std_legend()
        )
        return fig

    # ------------------------------------------------------------------
    # Automated Insights
    # ------------------------------------------------------------------

    def _generate_automated_insights(self, data: dict) -> str:
        """
        Narrative insights with:
        - Absolute level (latest value vs institutional thresholds)
        - Trend direction (linear slope over last 3 periods)
        - Z-Score distress signal (both models)
        - Cash flow quality (Quality of Income)
        - ROIC vs hurdle rate
        """
        insights = []

        def _add(emoji, bold, body, trend=""):
            t = (f" <span style='color:#888;font-size:0.9em;'>{trend}</span>"
                 if trend else "")
            insights.append(f"{emoji} <b>{bold}:</b> {body}{t}")

        # 1. Net Margin
        nm_vals = self._get_metric(data, "Net Margin")
        if nm_vals and nm_vals[-1] is not None:
            nm = nm_vals[-1]
            tr = self._trend_direction(nm_vals)
            if nm > 0.20:   _add("✅", "Net Margin", f"{nm:.1%} — Exceptional profitability", tr)
            elif nm > 0.15: _add("✅", "Net Margin", f"{nm:.1%} — Strong", tr)
            elif nm > 0.10: _add("⚠️", "Net Margin", f"{nm:.1%} — Moderate (benchmark ≥10%)", tr)
            elif nm > 0.05: _add("⚠️", "Net Margin", f"{nm:.1%} — Below benchmark", tr)
            elif nm > 0:    _add("🔴", "Net Margin", f"{nm:.1%} — Very thin; limited buffer", tr)
            else:           _add("🔴", "Net Margin", f"{nm:.1%} — Negative; company is loss-making", tr)

        # 2. ROIC vs cost of capital
        roic_vals = self._get_metric(data, "ROIC")
        if roic_vals and roic_vals[-1] is not None:
            rv = roic_vals[-1]
            tr = self._trend_direction(roic_vals)
            if rv > 0.20:   _add("✅", "ROIC", f"{rv:.1%} — Exceptional value creation (>20%)", tr)
            elif rv > 0.15: _add("✅", "ROIC", f"{rv:.1%} — Well above typical WACC", tr)
            elif rv > 0.10: _add("✅", "ROIC", f"{rv:.1%} — Above 10% hurdle rate", tr)
            elif rv > 0.05: _add("⚠️", "ROIC", f"{rv:.1%} — Marginal; may not cover cost of capital", tr)
            else:           _add("🔴", "ROIC", f"{rv:.1%} — Likely destroying shareholder value", tr)

        # 3. Liquidity
        cr_vals = self._get_metric(data, "Current Ratio")
        if cr_vals and cr_vals[-1] is not None:
            cr = cr_vals[-1]
            tr = self._trend_direction(cr_vals)
            if cr >= 2.0:   _add("✅", "Liquidity", f"Current ratio {cr:.2f}× — Comfortable buffer", tr)
            elif cr >= 1.5: _add("✅", "Liquidity", f"Current ratio {cr:.2f}× — Sound", tr)
            elif cr >= 1.0: _add("⚠️", "Liquidity", f"Current ratio {cr:.2f}× — Adequate but thin (target ≥1.5×)", tr)
            else:           _add("🔴", "Liquidity", f"Current ratio {cr:.2f}× — Below 1.0×; short-term obligations exceed current assets", tr)

        # 4. Leverage
        de_vals = self._get_metric(data, "Debt to Equity")
        if de_vals and de_vals[-1] is not None:
            de = de_vals[-1]
            tr = self._trend_direction(de_vals)
            if de < 0.5:   _add("✅", "Leverage", f"D/E {de:.2f}× — Conservative capital structure", tr)
            elif de < 1.0: _add("✅", "Leverage", f"D/E {de:.2f}× — Moderate leverage", tr)
            elif de < 2.0: _add("⚠️", "Leverage", f"D/E {de:.2f}× — Elevated; monitor refinancing risk", tr)
            else:          _add("🔴", "Leverage", f"D/E {de:.2f}× — High leverage; significant financial risk", tr)

        # 5. Interest Coverage (CFA min 3.0×)
        ic_vals = self._get_metric(data, "Interest Coverage")
        if ic_vals and ic_vals[-1] is not None:
            ic = ic_vals[-1]
            tr = self._trend_direction(ic_vals)
            if ic > 5.0:    _add("✅", "Interest Coverage", f"{ic:.1f}× — Very comfortable", tr)
            elif ic >= 3.0: _add("✅", "Interest Coverage", f"{ic:.1f}× — Meets CFA threshold (≥3.0×)", tr)
            elif ic > 1.5:  _add("⚠️", "Interest Coverage", f"{ic:.1f}× — Below CFA threshold; limited headroom", tr)
            elif ic > 1.0:  _add("🔴", "Interest Coverage", f"{ic:.1f}× — Barely covering interest", tr)
            else:           _add("🔴", "Interest Coverage", f"{ic:.1f}× — Cannot cover interest from operations", tr)

        # 6. Revenue Growth
        rg_vals = self._get_metric(data, "Revenue Growth")
        if rg_vals and len(rg_vals) > 1 and rg_vals[-1] is not None:
            rg = rg_vals[-1]
            tr = self._trend_direction(rg_vals)
            if rg > 0.15:   _add("✅", "Revenue Growth", f"{rg:.1%} YoY — Strong expansion", tr)
            elif rg > 0.05: _add("✅", "Revenue Growth", f"{rg:.1%} YoY — Healthy growth", tr)
            elif rg > 0:    _add("⚠️", "Revenue Growth", f"{rg:.1%} YoY — Modest; watch competitive dynamics", tr)
            else:           _add("🔴", "Revenue Growth", f"{rg:.1%} YoY — Revenue contraction", tr)

        # 7. Cash flow quality
        qoi_vals = self._get_metric(data, "Quality of Income")
        if qoi_vals and qoi_vals[-1] is not None:
            qv = qoi_vals[-1]
            tr = self._trend_direction(qoi_vals)
            if qv >= 1.2:   _add("✅", "Earnings Quality", f"Quality of Income {qv:.2f}× — Cash earnings significantly exceed reported profit", tr)
            elif qv >= 0.9: _add("✅", "Earnings Quality", f"Quality of Income {qv:.2f}× — Reported profits well-backed by cash", tr)
            elif qv >= 0.6: _add("⚠️", "Earnings Quality", f"Quality of Income {qv:.2f}× — Partial cash backing; check accruals", tr)
            else:           _add("🔴", "Earnings Quality", f"Quality of Income {qv:.2f}× — Profits significantly exceed operating cash — low quality signal", tr)

        # 8. Z-Score distress signals
        for label, vals, safe, grey in [
            ("Z-Score (Public)",  self._get_metric(data, "Altman Z-Score"),          2.99, 1.81),
            ("EM Score",          self._get_metric(data, "Altman Z-Score (EM Score)"), 2.60, 1.10),
        ]:
            if vals and vals[-1] is not None:
                zv = vals[-1]
                tr = self._trend_direction(vals)
                if zv > safe:              _add("✅", label, f"{zv:.2f} — Safe zone (>{safe})", tr)
                elif zv > grey:            _add("⚠️", label, f"{zv:.2f} — Grey zone ({grey}–{safe}); monitor closely", tr)
                else:                      _add("🔴", label, f"{zv:.2f} — Distress zone (<{grey}); elevated bankruptcy probability", tr)

        if not insights:
            return "📊 <i>Insufficient data for automated insights. Please check field mappings.</i>"
        return "<br><br>".join(insights)

    # ------------------------------------------------------------------
    # Peer Comparison (now wired into generate_html)
    # ------------------------------------------------------------------

    def _create_peer_comparison(self, focus_company: str, metric_name: str) -> str:
        """
        HTML table: focus company vs peer median / mean / percentile.
        Direction-aware colouring via _LOWER_IS_BETTER.
        Formatting via _format_value (percentage vs number auto-detected).
        """
        peer_data = {}
        for company, data in self.results.items():
            values = self._get_metric(data, metric_name)
            if values and values[-1] is not None:
                peer_data[company] = values[-1]

        if not peer_data or focus_company not in peer_data:
            return (f"<p style='color:gray;font-style:italic;'>"
                    f"Peer data not available for {html.escape(metric_name)}</p>")

        other_values = [v for k, v in peer_data.items() if k != focus_company]
        if not other_values:
            return "<p style='color:gray;'>No peer data — only one company in dataset.</p>"

        focus_value = peer_data[focus_company]
        peer_median = float(np.median(other_values))
        peer_mean   = float(np.mean(other_values))

        all_values = sorted(other_values + [focus_value])
        percentile = (len([v for v in all_values if v <= focus_value]) / len(all_values)) * 100

        is_above = focus_value > peer_median
        if metric_name in _LOWER_IS_BETTER:
            color     = "#27ae60" if not is_above else "#e74c3c"
            rank_word = "Below" if not is_above else "Above"
        else:
            color     = "#27ae60" if is_above else "#e74c3c"
            rank_word = "Above" if is_above else "Below"

        fmt = self._format_value

        return f"""
        <table style="width:100%;border-collapse:collapse;margin:8px 0;font-size:0.91em;">
            <tr style="background:#f0f4f8;">
                <td style="padding:7px 10px;border:1px solid #dde;"><b>{html.escape(metric_name)}</b></td>
                <td style="padding:7px 10px;border:1px solid #dde;color:{color};font-weight:700;">
                    {fmt(focus_value, metric_name)}</td>
            </tr>
            <tr>
                <td style="padding:7px 10px;border:1px solid #dde;">Peer Median</td>
                <td style="padding:7px 10px;border:1px solid #dde;">{fmt(peer_median, metric_name)}</td>
            </tr>
            <tr>
                <td style="padding:7px 10px;border:1px solid #dde;">Peer Average</td>
                <td style="padding:7px 10px;border:1px solid #dde;">{fmt(peer_mean, metric_name)}</td>
            </tr>
            <tr style="background:#f0f4f8;">
                <td style="padding:7px 10px;border:1px solid #dde;font-weight:600;">vs Peers</td>
                <td style="padding:7px 10px;border:1px solid #dde;font-weight:700;">
                    {rank_word} peer median — {percentile:.0f}th percentile
                </td>
            </tr>
        </table>"""

    # ------------------------------------------------------------------
    # Summary table
    # ------------------------------------------------------------------

    def _create_summary_table(self, data: dict) -> str:
        """
        Summary table with corrected CFA-standard thresholds:
        - Current Ratio ≥ 1.5 = green (not 1.0)
        - Interest Coverage ≥ 3.0x = green (not 2.5x)
        - Growth averages use geometric mean (CAGR)
        - Trend column uses 3-period linear regression
        - Cash Flow Quality section added
        """
        metric_categories = {
            "Profitability": [
                ("Gross Margin",     "percentage"),
                ("Operating Margin", "percentage"),
                ("Net Margin",       "percentage"),
                ("EBITDA Margin",    "percentage"),
            ],
            "Returns": [
                ("ROIC",             "percentage"),
                ("ROE",              "percentage"),
                ("ROA",              "percentage"),
                ("ROE (Normalized)", "percentage"),
            ],
            "Leverage": [
                ("Debt to Equity",    "number"),
                ("Debt to Assets",    "number"),
                ("Net Debt to EBITDA","number"),
                ("Interest Coverage", "number"),
            ],
            "Liquidity": [
                ("Current Ratio",     "number"),
                ("Quick Ratio",       "number"),
                ("Cash Ratio",        "number"),
            ],
            "Cash Flow Quality": [
                ("OCF to Sales",      "percentage"),
                ("FCF to Sales",      "percentage"),
                ("Quality of Income", "number"),
                ("Capex Coverage",    "number"),
            ],
            "Valuation": [
                ("P/E Ratio",          "number"),
                ("EV / EBITDA",        "number"),
                ("Price to Book (P/B)","number"),
            ],
            "Growth (Latest YoY)": [
                ("Revenue Growth",         "growth_rate"),
                ("EBITDA Growth",          "growth_rate"),
                ("Net Income Growth",      "growth_rate"),
                ("EPS Growth",             "growth_rate"),
            ],
        }

        def _color(metric, latest):
            if latest is None:
                return "black"
            # Leverage
            if metric == "Debt to Equity":
                return "#27ae60" if latest < 1.0 else ("#f39c12" if latest < 2.0 else "#e74c3c")
            if metric == "Debt to Assets":
                return "#27ae60" if latest < 0.5 else ("#f39c12" if latest < 0.7 else "#e74c3c")
            if metric == "Net Debt to EBITDA":
                return "#27ae60" if latest < 2.0 else ("#f39c12" if latest < 3.5 else "#e74c3c")
            if metric == "Interest Coverage":
                # CFA standard: ≥3.0 healthy, 1.5–3.0 caution, <1.5 distress
                return "#27ae60" if latest >= 3.0 else ("#f39c12" if latest >= 1.5 else "#e74c3c")
            # Liquidity — CFA standard: ≥1.5 sound
            if metric == "Current Ratio":
                return "#27ae60" if latest >= 1.5 else ("#f39c12" if latest >= 1.0 else "#e74c3c")
            if metric in ("Quick Ratio", "Cash Ratio"):
                return "#27ae60" if latest >= 1.0 else "#e74c3c"
            # Cash flow quality
            if metric == "Quality of Income":
                return "#27ae60" if latest >= 0.9 else ("#f39c12" if latest >= 0.6 else "#e74c3c")
            if metric == "Capex Coverage":
                return "#27ae60" if latest >= 1.5 else ("#f39c12" if latest >= 1.0 else "#e74c3c")
            # Valuation: context-dependent → neutral
            if metric in ("P/E Ratio", "EV / EBITDA", "Price to Book (P/B)"):
                return "black"
            # Default: positive = green
            return "#27ae60" if latest > 0 else "#e74c3c"

        html_parts = ['<table class="summary-table">']
        html_parts.append(
            '<thead><tr>'
            '<th>Metric</th><th>Latest</th><th>Period Avg</th><th>Trend (3-period)</th>'
            '</tr></thead><tbody>'
        )

        for category, metrics in metric_categories.items():
            html_parts.append(
                f'<tr style="background:#f0f4f8;">'
                f'<td colspan="4"><b>{html.escape(category)}</b></td></tr>'
            )
            for metric, mtype in metrics:
                values = self._get_metric(data, metric)
                valid  = [v for v in values if v is not None]

                if not valid:
                    html_parts.append(
                        f'<tr><td>{html.escape(metric)}</td>'
                        f'<td colspan="3" style="color:#aaa;">N/A</td></tr>'
                    )
                    continue

                latest = values[-1]
                color  = _color(metric, latest)

                if mtype in ("percentage", "growth_rate"):
                    latest_str = self._format_percentage(latest)
                else:
                    latest_str = self._format_number(latest)

                try:
                    if mtype == "growth_rate":
                        factors = [1 + v for v in valid if v > -0.99]
                        if len(factors) >= 2 and all(g > 0 for g in factors):
                            product = np.prod(factors)
                            if np.isfinite(product) and product <= 1e308:
                                geo = (product ** (1 / len(factors))) - 1
                                avg_str = self._format_percentage(geo)
                            else:
                                avg_str = self._format_percentage(np.mean(valid))
                        else:
                            avg_str = self._format_percentage(np.mean(valid))
                    elif mtype == "percentage":
                        avg_str = self._format_percentage(np.mean(valid))
                    else:
                        avg_str = self._format_number(np.mean(valid))
                except (OverflowError, ValueError, TypeError):
                    avg_str = "N/A"

                trend = self._trend_direction(values)

                html_parts.append(
                    f'<tr>'
                    f'<td>{html.escape(metric)}</td>'
                    f'<td style="color:{color};font-weight:700;">{latest_str}</td>'
                    f'<td>{avg_str}</td>'
                    f'<td style="color:#666;font-size:0.9em;">{trend}</td>'
                    f'</tr>'
                )

        html_parts.append('</tbody></table>')
        return "".join(html_parts)

    # ------------------------------------------------------------------
    # Main HTML generation
    # ------------------------------------------------------------------

    def generate_html(self, focus_company: str) -> str:
        """
        Generate a fully self-contained HTML dashboard.

        Sections (in order):
          1.  Key Metrics Summary table
          2.  Peer Comparison (Net Margin, ROIC, D/E, Current Ratio)
          3.  Automated Insights narrative
          4.  Charts:
              Row 1 (full-width): Profitability Margins
              Row 2 (2-col):      Returns | DuPont Decomposition
              Row 3 (2-col):      Cash Flow Quality | Liquidity
              Row 4 (2-col):      Leverage & Solvency | Efficiency
              Row 5 (2-col):      Growth Rates | Bankruptcy Risk
        """
        if focus_company not in self.results:
            available = ", ".join(list(self.results.keys())[:5])
            raise ValueError(
                f"Company '{focus_company}' not found. Available: {available}..."
            )

        safe_name = html.escape(focus_company)
        data = self.results[focus_company]

        try:
            fig_margins    = self._create_margin_chart(data)
            fig_returns    = self._create_returns_chart(data)
            fig_dupont     = self._create_dupont_chart(data)
            fig_cashflow   = self._create_cashflow_chart(data)
            fig_liquidity  = self._create_liquidity_chart(data)
            fig_leverage   = self._create_leverage_chart(data)
            fig_efficiency = self._create_efficiency_chart(data)
            fig_growth     = self._create_growth_chart(data)
            fig_risk       = self._create_risk_dashboard(data)

            for fig, name in [
                (fig_margins,   "margins"),   (fig_returns,    "returns"),
                (fig_dupont,    "dupont"),     (fig_cashflow,   "cashflow"),
                (fig_liquidity, "liquidity"),  (fig_leverage,   "leverage"),
                (fig_efficiency,"efficiency"), (fig_growth,     "growth"),
                (fig_risk,      "risk"),
            ]:
                if not hasattr(fig, 'to_html'):
                    raise ValueError(f"'{name}' is not a valid Figure: {type(fig)}")
        except Exception as exc:
            logger.warning(f"Dashboard generation failed: {exc}")
            raise ValueError(f"Failed to generate dashboard: {exc}")

        insights      = self._generate_automated_insights(data)
        summary_table = self._create_summary_table(data)

        # Peer comparison for 4 key decision metrics
        peer_sections = ""
        for pm in ["Net Margin", "ROIC", "Debt to Equity", "Current Ratio"]:
            peer_sections += (
                f"<h4 style='margin:14px 0 4px;color:#0F2841;'>{html.escape(pm)}</h4>"
                + self._create_peer_comparison(focus_company, pm)
            )

        _cfg = {'responsive': True, 'displayModeBar': False}
        def _fig(f):
            return f.to_html(full_html=False, include_plotlyjs=False, config=_cfg)

        out = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Financial Dashboard: {safe_name}</title>
    <style>
        *{{margin:0;padding:0;box-sizing:border-box;}}
        body{{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;
              background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
              padding:20px;min-height:100vh;}}
        .container{{max-width:1440px;margin:0 auto;}}
        .header{{background:linear-gradient(135deg,#0F2841 0%,#1a3a52 100%);
                 color:white;padding:30px;border-radius:12px;margin-bottom:25px;
                 box-shadow:0 10px 30px rgba(0,0,0,.3);}}
        .header h1{{font-size:2.2em;margin-bottom:8px;font-weight:700;}}
        .header p{{font-size:1.05em;opacity:.9;margin-top:4px;}}
        .card{{background:white;padding:22px;border-radius:12px;
               box-shadow:0 4px 15px rgba(0,0,0,.1);margin-bottom:25px;
               transition:transform .25s ease,box-shadow .25s ease;}}
        .card:hover{{transform:translateY(-4px);box-shadow:0 8px 25px rgba(0,0,0,.15);}}
        .card h2{{color:#0F2841;margin-bottom:16px;font-size:1.4em;
                  border-bottom:3px solid #667eea;padding-bottom:10px;}}
        .card h3{{color:#334;margin-bottom:10px;font-size:1.05em;font-weight:600;}}
        .grid-2{{display:grid;grid-template-columns:repeat(auto-fit,minmax(580px,1fr));gap:25px;}}
        .full-width{{grid-column:1/-1;}}
        .summary-table{{width:100%;border-collapse:collapse;font-size:0.92em;}}
        .summary-table th{{background:#f8f9fa;padding:10px 12px;text-align:left;
                           font-weight:600;color:#0F2841;border-bottom:2px solid #dee2e6;}}
        .summary-table td{{padding:9px 12px;border-bottom:1px solid #eee;}}
        .summary-table tr:hover{{background:#fafafa;}}
        .insights-box{{line-height:2.1;color:#333;font-size:0.95em;}}
        .footer{{text-align:center;color:white;padding:20px;margin-top:5px;opacity:.85;}}
        @media(max-width:768px){{
            .grid-2{{grid-template-columns:1fr;}}
            .header h1{{font-size:1.5em;}}
        }}
        @media print{{
            body{{background:white;padding:0;
                  -webkit-print-color-adjust:exact;print-color-adjust:exact;}}
            .header{{background:#0F2841!important;
                     -webkit-print-color-adjust:exact;print-color-adjust:exact;}}
            .card{{box-shadow:none;border:1px solid #ddd;
                   page-break-inside:avoid;break-inside:avoid;}}
            .card:hover{{transform:none;box-shadow:none;}}
            .grid-2{{display:block;}}
            .grid-2>.card{{margin-bottom:20px;}}
        }}
    </style>
    <script src="https://cdn.plot.ly/plotly-2.26.0.min.js"></script>
    <script>
        window.addEventListener('load',function(){{
            if(typeof Plotly==='undefined'){{
                document.body.innerHTML='<div style="text-align:center;padding:60px;font-family:sans-serif;">'
                +'<h2 style="color:#e74c3c;">&#9888; Plotly Not Loaded</h2>'
                +'<p>An internet connection is required to load the charting library.</p>'
                +'<button onclick="location.reload()" style="margin-top:20px;padding:10px 24px;'
                +'font-size:1em;cursor:pointer;">&#x1F504; Retry</button></div>';
            }}
        }});
    </script>
</head>
<body>
<div class="container">

  <div class="header">
    <h1>&#128202; Financial Dashboard: {safe_name}</h1>
    <p>Institutional-Grade Ratio Analysis</p>
    <p style="font-size:0.88em;margin-top:6px;opacity:.8;">
      Period: {self.years[0] if self.years else 'N/A'} &ndash; {self.years[-1] if self.years else 'N/A'}
      &nbsp;|&nbsp; Peer group: {len(self.engine.companies)} companies
    </p>
  </div>

  <!-- 1. SUMMARY TABLE -->
  <div class="card">
    <h2>&#128200; Key Metrics Summary</h2>
    {summary_table}
  </div>

  <!-- 2. PEER COMPARISON -->
  <div class="card">
    <h2>&#127970; Peer Comparison (Latest Period)</h2>
    <p style="color:#666;font-size:0.88em;margin-bottom:12px;">
      Focus company vs all other companies in the dataset.
      Green = favourable relative to peer median. Red = unfavourable.
      Direction-aware: for debt / cost metrics lower is better.
    </p>
    <div class="grid-2">
      <div>{peer_sections}</div>
    </div>
  </div>

  <!-- 3. INSIGHTS -->
  <div class="card">
    <h2>&#128161; Automated Insights</h2>
    <div class="insights-box">{insights}</div>
    <p style="font-size:0.78em;color:#aaa;margin-top:14px;">
      Trend labels use the linear slope of the last 3 available periods.
      Thresholds follow CFA Institute standards where applicable.
      All figures are estimates &mdash; verify independently before decision-making.
    </p>
  </div>

  <!-- 4a. MARGINS (full-width) -->
  <div class="card full-width">
    <h3>Profitability Margins</h3>
    {_fig(fig_margins)}
  </div>

  <!-- 4b. Returns | DuPont -->
  <div class="grid-2">
    <div class="card">
      <h3>Return on Investment (ROIC / ROE / ROA)</h3>
      {_fig(fig_returns)}
    </div>
    <div class="card">
      <h3>DuPont ROE Decomposition</h3>
      {_fig(fig_dupont)}
    </div>
  </div>

  <!-- 4c. Cash Flow | Liquidity -->
  <div class="grid-2">
    <div class="card">
      <h3>Cash Flow Quality &amp; Coverage</h3>
      {_fig(fig_cashflow)}
    </div>
    <div class="card">
      <h3>Liquidity Analysis</h3>
      {_fig(fig_liquidity)}
    </div>
  </div>

  <!-- 4d. Leverage | Efficiency -->
  <div class="grid-2">
    <div class="card">
      <h3>Leverage &amp; Solvency</h3>
      {_fig(fig_leverage)}
    </div>
    <div class="card">
      <h3>Operating Efficiency &mdash; Working Capital Cycle</h3>
      {_fig(fig_efficiency)}
    </div>
  </div>

  <!-- 4e. Growth | Risk -->
  <div class="grid-2">
    <div class="card">
      <h3>Year-over-Year Growth Rates</h3>
      {_fig(fig_growth)}
    </div>
    <div class="card">
      <h3>Bankruptcy Risk (Altman Z-Score &amp; EM Score)</h3>
      {_fig(fig_risk)}
    </div>
  </div>

  <div class="footer">
    <p>Generated {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}
       &nbsp;|&nbsp; {len(self.engine.companies)} companies
       &nbsp;|&nbsp; {len(self.years)} periods</p>
    <p style="font-size:0.82em;margin-top:6px;opacity:.75;">
      For educational and analytical purposes only. Not financial advice.
      All calculations are estimates &mdash; verify independently before decisions.
    </p>
  </div>

</div>
</body>
</html>"""
        return out
