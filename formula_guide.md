# Financial Ratio Analysis — Formula Guide

**Engine Version:** v3.0  
**Ratios Covered:** 60+  
**Last Updated:** March 2026

---

## How to Read This Guide

Each ratio entry includes:

- **Formula** — the exact calculation used
- **Why** — the analytical purpose
- **Nuances** — caps, edge cases, and special handling in the engine

**Notation conventions:**

| Symbol | Meaning |
|--------|---------|
| Avg(X) | (X_current + X_prior) ÷ 2. First period uses current value only. |
| NaN / N/A | Not computed — denominator is zero, negative, or data is missing. |
| ×(1 − t) | After-tax adjustment using your configured corporate tax rate. |
| Trailing Sum | For quarterly/monthly data, sums the last 4/12 periods to annualize flow items. |

---

## Table of Contents

1. [Liquidity Ratios](#1-liquidity-ratios)
2. [Efficiency / Activity Ratios](#2-efficiency--activity-ratios)
3. [Solvency / Leverage Ratios](#3-solvency--leverage-ratios)
4. [Profitability Ratios](#4-profitability-ratios)
5. [Return Ratios](#5-return-ratios)
6. [Cash Flow Ratios](#6-cash-flow-ratios)
7. [Valuation Ratios](#7-valuation-ratios)
8. [Growth Ratios](#8-growth-ratios)
9. [Risk & Quality Models](#9-risk--quality-models)
10. [Global Engine Nuances](#10-global-engine-nuances)

---

## 1. Liquidity Ratios

### Current Ratio

**Formula:** Total Current Assets ÷ Total Current Liabilities

**Why:** Measures the company's ability to meet short-term obligations with short-term assets. A ratio above 1.0 means current assets exceed current liabilities.

**Nuances:**
- Returns N/A when Total Current Liabilities = 0 (no short-term obligations to measure against).
- Standard healthy range: 1.5–3.0. Below 1.0 signals potential liquidity stress.

---

### Quick Ratio (Acid Test)

**Formula:** (Cash + Short-Term Investments + Accounts Receivable) ÷ Total Current Liabilities

**Fallback:** If cash, short-term investments, and receivables are not all available individually, the engine uses: (Total Current Assets − Inventory − Prepaid Expenses) ÷ Total Current Liabilities.

**Why:** Stricter than the Current Ratio — excludes inventory and prepaid expenses because they cannot be quickly converted to cash to meet obligations.

**Nuances:**
- Prepaid expenses are excluded from the fallback calculation (per strict CFA definition), not just inventory.
- Returns N/A when Total Current Liabilities = 0.

---

### Cash Ratio

**Formula:** (Cash & Equivalents + Short-Term Investments) ÷ Total Current Liabilities

**Why:** The most conservative liquidity measure. Only counts assets that are already cash or near-cash.

**Nuances:**
- Returns N/A when Total Current Liabilities = 0.
- Typical values are well below 1.0; a very high Cash Ratio may indicate inefficient use of cash.

---

### Defensive Interval (Days)

**Formula:** (Cash + Short-Term Investments + Accounts Receivable) ÷ Daily Cash Operating Expenses

Where: Daily Cash Expenses = (Cost of Revenue + Operating Expenses − D&A) ÷ 365

**Why:** Estimates how many days the company can operate using only its liquid assets, without generating any additional revenue. Useful for assessing runway.

**Nuances:**
- D&A is subtracted from expenses because it is non-cash.
- If cash expenses come out ≤ 0 (e.g., D&A exceeds COGS + OpEx), the result is N/A — negative burn rate is meaningless.
- Uses the same quick assets as the Quick Ratio numerator.

---

### NWC to Assets

**Formula:** (Total Current Assets − Total Current Liabilities) ÷ Total Assets

**Why:** Shows what proportion of total assets is funded by net working capital. Indicates the liquidity cushion relative to company size.

**Nuances:**
- Can be negative if current liabilities exceed current assets.
- Returns N/A when Total Assets = 0.

---

## 2. Efficiency / Activity Ratios

### Inventory Turnover

**Formula:** Cost of Revenue ÷ Avg(Inventory)

**Why:** Measures how many times inventory is sold and replenished during the period. Higher values indicate faster-moving inventory and better capital efficiency.

**Nuances:**
- **Manufacturing industry only.** Not computed for Service industry.
- Capped at 100× to prevent extreme outliers from distorting charts.
- Uses Avg(Inventory) = average of current and prior period inventory.

---

### Days Inventory Outstanding (DIO)

**Formula:** Avg(Inventory) ÷ Cost of Revenue × 365

**Why:** The inverse of Inventory Turnover expressed in days. Shows how long inventory sits before being sold.

**Nuances:**
- Manufacturing industry only.
- For quarterly data, the engine uses the period's actual days (~91.25); for monthly, ~30.42 days.

---

### Receivables Turnover

**Formula:** Revenue ÷ Avg(Accounts Receivable)

**Why:** Measures how efficiently the company collects payment from customers. Higher = faster collection.

**Nuances:**
- Capped at 100×.
- Uses average receivables to smooth period-end spikes.

---

### Days Sales Outstanding (DSO)

**Formula:** Avg(Accounts Receivable) ÷ Revenue × 365

**Why:** Average number of days to collect payment after a sale. Standard target: 30–60 days.

**Nuances:**
- Adjusts for data frequency (quarterly uses ~91.25 days, monthly uses ~30.42 days).

---

### Payables Turnover

**Formula (Manufacturing):** Cost of Revenue ÷ Avg(Accounts Payable)

**Formula (Service):** When Cost of Revenue ≈ 0, uses Operating Expenses ÷ Avg(Accounts Payable) instead.

**Why:** Measures how quickly the company pays its suppliers.

**Nuances:**
- Capped at 100×.
- The Service industry adjustment exists because service firms often have near-zero COGS, which would make the standard formula meaningless.

---

### Days Payables Outstanding (DPO)

**Formula:** Avg(Accounts Payable) ÷ Cost of Revenue × 365

(Service: uses Operating Expenses instead of COGS)

**Why:** Average number of days the company takes to pay suppliers.

---

### Cash Conversion Cycle

**Formula (Manufacturing):** DIO + DSO − DPO

**Formula (Service):** DSO − DPO (no inventory component)

**Why:** Measures the total time (in days) between paying for inventory and receiving cash from sales. A shorter cycle means faster cash recovery. Negative CCC means the company collects cash before paying suppliers — a strong working capital position.

---

### Total Asset Turnover

**Formula:** Revenue ÷ Avg(Total Assets)

**Why:** Measures revenue generated per unit of total assets. A DuPont decomposition component.

**Nuances:**
- For quarterly/monthly data, revenue is annualized using trailing sum before dividing.

---

### Fixed Asset Turnover

**Formula:** Revenue ÷ Avg(PP&E Net)

**Why:** Measures how productively the company uses its fixed assets (property, plant, equipment).

**Nuances:**
- For quarterly/monthly, revenue is annualized via trailing sum.

---

### Working Capital Turnover

**Formula:** Revenue ÷ Net Working Capital

Where: NWC = Total Current Assets − Total Current Liabilities

**Why:** Measures revenue generated per unit of working capital.

**Nuances:**
- **Returns N/A when NWC ≤ 0.** A negative working capital makes the ratio's sign misleading — it would imply you generate less revenue with more efficient capital, which is nonsensical. The N/A protects against misinterpretation.

---

### CapEx to D&A

**Formula:** Capital Expenditures ÷ Depreciation & Amortisation

**Why:** Shows whether the company is investing more than it depreciates.
- \>1 = net investment (growing asset base)
- ≈1 = maintenance mode
- <1 = asset harvesting (under-investing)

---

### Degree of Operating Leverage (DOL)

**Formula:** (%Δ Operating Income) ÷ (%Δ Revenue)

**Why:** Measures fixed-cost intensity. High DOL (>3) means small revenue changes cause large profit swings — high risk in downturns but high gearing in upturns.

**Nuances:**
- First period is always N/A (no prior period for growth calculation).
- Growth rates capped at ±1000% to prevent extreme outliers.

---

### Reinvestment Rate

**Formula:** (CapEx − D&A + ΔNWC) ÷ NOPAT

Where: NOPAT = Operating Income × (1 − Tax Rate)

**Why:** What fraction of after-tax operating profit is reinvested into the business. Pairs with ROIC: high reinvestment + high ROIC = fast compounding.

**Nuances:**
- Undefined (N/A) when NOPAT ≤ 0.

---

### Sustainable Growth Rate

**Formula:** ROIC × Reinvestment Rate

**Why:** Maximum growth achievable without external financing, assuming current returns and reinvestment rates hold constant. Theoretical ceiling on organic growth.

---

## 3. Solvency / Leverage Ratios

### Debt to Equity

**Formula:** Total Debt ÷ Total Equity

**Why:** Measures financial leverage. Conservative threshold: <1.0.

**Nuances:**
- **Returns N/A when equity ≤ 0.** Negative equity flips the ratio's sign, making a negative D/E look like the company has low leverage when it is actually in deep distress. The distress signal is visible via Interest Coverage and Net Debt/EBITDA instead.

---

### Net Debt to Equity

**Formula:** (Total Debt − Cash) ÷ Total Equity

**Why:** Unlike D/E, explicitly nets off cash holdings. Negative = net cash company (more cash than debt). More informative for capital allocation analysis.

---

### Debt to Assets

**Formula:** Total Debt ÷ Total Assets

**Why:** Shows what proportion of total assets is financed by debt.

---

### Debt to Capital

**Formula:** Total Debt ÷ (Total Debt + Total Equity)

**Why:** Debt as a proportion of total capital. Bounded between 0 and 1 (unless equity is negative).

---

### Interest Coverage

**Formula:** Operating Income (EBIT) ÷ Gross Interest Expense

**Why:** Measures ability to service debt obligations from operating earnings. Threshold: ≥3.0×.

**Nuances:**
- **Interest Expense is used as an absolute (gross) value** — per CFA/IFRS convention, coverage ratios use gross interest, not net of interest income.
- **When Interest Expense = 0 and Operating Income > 0:** the company is self-financed. The ratio is capped at **100×** (the `int_cov_cap` setting) rather than showing infinity.
- **When Interest Expense = 0 and Operating Income ≤ 0:** returns N/A.
- **Negative coverage (operating loss with debt):** configurable via "Show Negative Interest Coverage" toggle:
  - **Enabled (default):** shows the negative value (e.g., −2.5×) for risk visibility.
  - **Disabled:** returns N/A for negative coverage — original conservative behavior.
- All values clipped to the range [−100, +100].

---

### Cash Interest Coverage

**Formula:** EBITDA ÷ Gross Interest Expense

**Why:** Cash-based variant. EBITDA better approximates cash available for debt service since it adds back non-cash D&A.

**Nuances:**
- Same zero-interest and negative-coverage handling as Interest Coverage (see above).
- For quarterly/monthly data, EBITDA is annualized via trailing sum.
- Capped at ±100×.

---

### Net Debt to EBITDA

**Formula:** (Total Debt − Cash) ÷ EBITDA

**Why:** Key leverage metric for credit analysis. Investment-grade threshold: <3.0×.

**Nuances:**
- For non-annual data, EBITDA is annualized via trailing sum.
- Negative EBITDA makes this ratio misleading; the engine allows it to flow through for visibility.

---

### Financial Leverage (Equity Multiplier)

**Formula:** Avg(Total Assets) ÷ Avg(Total Equity)

**Why:** DuPont decomposition component. Shows how many dollars of assets are supported by each dollar of equity.

**Nuances:**
- **Returns N/A when average equity ≤ 0.** Same rationale as D/E — sign inversion is misleading.

---

### Interest Expense Ratio

**Formula:** Interest Expense ÷ Revenue

**Why:** Shows the fraction of every revenue dollar consumed by debt service. Values above 5% in most industries warrant attention.

---

## 4. Profitability Ratios

All profitability ratios follow the pattern: **Numerator ÷ Revenue**. Returns N/A when Revenue = 0.

### Gross Margin

**Formula:** Gross Profit ÷ Revenue

**Fallback:** If Gross Profit is not directly mapped, the engine derives it: Revenue − Cost of Revenue. This derivation **only applies when both Revenue and Cost of Revenue have data** for a given period — if COGS is missing, the engine does NOT assume COGS = 0.

---

### Operating Margin

**Formula:** Operating Income ÷ Revenue

---

### Net Margin

**Formula:** Net Income ÷ Revenue

---

### EBITDA Margin

**Formula:** EBITDA ÷ Revenue

**Fallback:** If EBITDA is not directly mapped, derived as: Operating Income + D&A. Only computed when both have data.

---

### Pretax Margin

**Formula:** Pretax Income ÷ Revenue

**Fallback:** If Pretax Income is not mapped, derived as: Operating Income − Interest Expense. Only computed when both have data.

---

## 5. Return Ratios

### ROA (Return on Assets)

**Formula:** Net Income ÷ Avg(Total Assets)

**Why:** Measures how efficiently all assets generate profit. Benchmark: ≥5%.

**Nuances:**
- For quarterly/monthly data, Net Income is annualized via trailing sum.

---

### ROE (Return on Equity)

**Formula:** Net Income ÷ Avg(Total Equity)

**Why:** Measures return to shareholders. Benchmark: ≥15%.

**Nuances:**
- **Returns N/A when average equity ≤ 0.** Negative equity would invert the ratio (a loss would appear as a positive return).
- For non-annual data, Net Income is annualized.

---

### ROE (Normalized)

**Formula:** NOPAT ÷ Avg(Total Equity)

Where: NOPAT = Operating Income × (1 − Tax Rate)

**Why:** Removes the distortion of interest tax shields. Useful for comparing companies with different capital structures — one levered, one not.

**Nuances:**
- Same negative-equity handling as ROE.

---

### ROIC (Return on Invested Capital)

**Formula:** NOPAT ÷ Avg(Invested Capital)

Where:
- NOPAT = Operating Income × (1 − Tax Rate)
- Invested Capital = Total Equity + Total Debt + Minority Interest − Excess Cash
- Excess Cash = max(Cash − Operating Cash Requirement, 0)
- Operating Cash Requirement = Operating Cash % × Revenue (default: 2%)

**Why:** The gold-standard return metric. Should exceed the company's cost of capital for value creation.

**Nuances:**
- **Excess cash adjustment:** The engine subtracts cash beyond what's needed for operations from invested capital, since that cash is a financial asset, not an operating investment. The 2% default is an estimate — adjust per industry.
- **Invested Capital ≤ 0 → N/A.** Negative IC is economically meaningless.
- **NOPAT tax benefit on losses:** configurable:
  - **Standard (default):** NOPAT = OpInc × (1−t) even when OpInc is negative (assumes future tax benefit).
  - **Conservative:** No tax benefit on losses — NOPAT = OpInc when OpInc < 0.

---

### Cash ROIC

**Formula:** Free Cash Flow ÷ Avg(Invested Capital)

**Why:** Cash-based variant of ROIC. Less susceptible to accounting manipulation.

---

### ROIC Spread (vs 10% Hurdle)

**Formula:** ROIC − 10%

**Why:** Positive = value creation above a rough universal hurdle rate; negative = value destruction. The 10% is a proxy — interpret with judgment for capital-intensive sectors.

---

### Earnings Yield (EBIT/EV)

**Formula:** Operating Income ÷ Enterprise Value

**Why:** Greenblatt's Magic Formula numerator. The earnings-based alternative to E/P that strips out capital structure effects. Higher = cheaper relative to earnings power.

---

### DuPont Decomposition

The DuPont identity decomposes ROE into three drivers:

**ROE = Net Margin × Asset Turnover × Equity Multiplier**

| Component | Formula |
|-----------|---------|
| DuPont: Net Margin | Net Income ÷ Revenue |
| DuPont: Asset Turnover | Revenue ÷ Avg(Total Assets) |
| DuPont: Equity Multiplier | Avg(Total Assets) ÷ Avg(Total Equity) |

**Nuances:**
- Equity Multiplier returns N/A when average equity ≤ 0.
- All three components use consistently annualized figures for non-annual data.
- The engine verifies the identity holds: NM × AT × EM ≈ ROE.

---

## 6. Cash Flow Ratios

### OCF to Sales

**Formula:** Operating Cash Flow ÷ Revenue

**Why:** Cash-generation efficiency from operations. Benchmark: ≥10%.

---

### FCF to Sales

**Formula:** Free Cash Flow ÷ Revenue

**Fallback:** If FCF is not mapped, derived as: Operating Cash Flow − Capital Expenditures. Only computed when both have data.

---

### Quality of Income

**Formula:** Operating Cash Flow ÷ Net Income

**Why:** Values >1.0 indicate high earnings quality (cash profits exceed accrual profits). Values <1.0 suggest aggressive accrual accounting.

---

### Capex Coverage

**Formula:** Operating Cash Flow ÷ Capital Expenditures

**Why:** Values >1.5 indicate the company can fund its capital investments internally without external financing.

---

### Dividend Payout

**Formula:** |Dividends Paid| ÷ Net Income

**Why:** What fraction of earnings is returned to shareholders. Sustainable range: 30–50%.

**Nuances:**
- **Returns N/A when Net Income ≤ 0.** A payout ratio on a loss is meaningless.
- Dividends are used as absolute values (cash flow statements report them as negative).

---

### FCF Conversion

**Formula:** Free Cash Flow ÷ EBITDA

**Why:** How much of EBITDA converts into actual free cash. Benchmark: ≥0.8×.

**Nuances:**
- **Returns N/A when EBITDA ≤ 0.** The ratio direction inverts with negative EBITDA.

---

### Sloan Accrual Ratio

**Formula:** (Net Income − Operating Cash Flow) ÷ Avg(Total Assets)

**Why:** Positive = accounting profits exceed cash (low earnings quality). Negative = cash exceeds profits (high quality). Values above +5% are a red flag. Source: Sloan (1996).

---

## 7. Valuation Ratios

### Earnings Per Share (EPS)

**Formula:** (Net Income − Preferred Dividends) ÷ Diluted Shares Outstanding

**Why:** The earnings attributable to each common share using the diluted share count.

**Nuances:**
- Uses **diluted** shares (not basic) for the denominator.
- If diluted shares are not mapped, falls back to basic shares.
- For non-annual data, Net Income and Preferred Dividends are annualized via trailing sum.

---

### P/E Ratio

**Formula:** Share Price ÷ EPS

**Why:** Market price relative to earnings. Market average range: 15–20×.

**Nuances:**
- **Returns N/A when EPS ≤ 0** (negative P/E is not meaningful).
- **Capped at 500×** (`pe_ceiling` config). Extremely high P/E ratios for near-zero earnings distort analysis.

---

### PEG Ratio

**Formula:** P/E ÷ (EPS Growth × 100)

**Why:** P/E adjusted for growth. Values <1.0 may indicate undervaluation relative to growth.

**Nuances:**
- **Returns N/A when EPS growth ≤ 0.** The ratio is only meaningful for growing earnings.
- EPS growth is multiplied by 100 to convert from decimal to percentage for the standard PEG formula.

---

### Book Value Per Share

**Formula:** Total Equity ÷ Basic Shares Outstanding

**Why:** Net asset value per share.

---

### Price to Book (P/B)

**Formula:** Share Price ÷ Book Value Per Share

**Why:** Compares market valuation to net asset value.

**Nuances:**
- **Returns N/A when BVPS ≤ 0** (negative book value inverts the ratio).

---

### Price to Sales

**Formula:** Market Capitalisation ÷ Revenue

**Fallback for Market Cap:** If not directly mapped, derived as Share Price × Basic Shares Outstanding.

---

### EV / EBITDA

**Formula:** Enterprise Value ÷ EBITDA

Where: EV = Market Cap + Total Debt + Minority Interest + Preferred Stock − Cash

**Why:** The standard enterprise-level valuation multiple. Range: 8–12×.

**Nuances:**
- For non-annual data, EBITDA is annualized via trailing sum.

---

### EV / Revenue

**Formula:** Enterprise Value ÷ Revenue

**Why:** Enterprise-level revenue multiple. Useful for unprofitable companies where EBITDA-based multiples fail.

---

### FCF Yield

**Formula:** Free Cash Flow ÷ Market Capitalisation

**Why:** Cash return on market price. Benchmark: ≥5%.

---

## 8. Growth Ratios

All growth ratios use the same formula:

**Formula:** (Current Period − Prior Period) ÷ |Prior Period|

**Nuances (apply to all growth ratios):**
- **First period is always N/A** — there is no prior period to compare against.
- **Division by absolute value of prior period** — this correctly handles turnarounds (negative-to-positive transitions) by always using a positive denominator.
- **Capped at ±1000%** (`growth_cap = 10.0`). Extreme spikes from near-zero base values would otherwise distort charts and averages.
- **Prior period = 0 → N/A.** Cannot compute growth from zero.
- **Single-period data → N/A.** No growth calculable.

| Ratio | Numerator Flow |
|-------|---------------|
| Revenue Growth | Revenue |
| Gross Profit Growth | Gross Profit |
| Operating Income Growth | Operating Income |
| EBITDA Growth | EBITDA |
| Net Income Growth | Net Income |
| EPS Growth | Annualized EPS (trailing for non-annual) |
| FCF Growth | Free Cash Flow |

---

## 9. Risk & Quality Models

### Altman Z-Score (Public Company)

**Formula (Altman 1968, 5-factor):**

Z = 1.2×A + 1.4×B + 3.3×C + 0.6×D + 1.0×E

| Factor | Formula |
|--------|---------|
| A | Working Capital ÷ Total Assets |
| B | Retained Earnings ÷ Total Assets |
| C | EBIT (Operating Income) ÷ Total Assets |
| D | Market Cap ÷ Total Liabilities |
| E | Revenue ÷ Total Assets |

**Interpretation:**
- Z > 2.99 → Safe zone
- 1.81 ≤ Z ≤ 2.99 → Grey zone
- Z < 1.81 → Distress zone

**Nuances:**
- Factor D returns N/A when Total Liabilities ≤ 0.
- Entire Z-Score is N/A when Market Cap is unavailable (required for factor D).
- For non-annual data, Revenue is annualized for factor E.

---

### Altman Z-Score (EM / Private)

**Formula (4-factor emerging market model):**

Z = 6.56×A + 3.26×B + 6.72×C + 1.05×D

| Factor | Formula |
|--------|---------|
| A | Working Capital ÷ Total Assets |
| B | Retained Earnings ÷ Total Assets |
| C | EBIT ÷ Total Assets |
| D | Total Equity ÷ Total Liabilities |

**Interpretation:**
- Z > 2.6 → Safe
- 1.1 ≤ Z ≤ 2.6 → Grey
- Z < 1.1 → Distress

**Nuances:**
- **Does not require market price data** — uses equity/liabilities instead. Suitable for private companies and emerging markets.
- Factor D returns N/A when Total Liabilities ≤ 0.

---

### Piotroski F-Score

**Formula (Piotroski 2000, 9-point binary composite):**

Each criterion scores 1 if favourable, 0 if not. Total: 0–9.

**Profitability (4 points):**
1. ROA > 0
2. Operating Cash Flow > 0
3. ΔROA > 0 (improving)
4. Quality of Income > 1 (OCF > NI — accrual quality)

**Leverage / Liquidity (3 points):**
5. ΔLeverage < 0 (declining D/A)
6. ΔCurrent Ratio > 0 (improving)
7. No new shares issued (diluted shares not increased)

**Efficiency (2 points):**
8. ΔGross Margin > 0
9. ΔAsset Turnover > 0

Score 8–9 = Strong buy signal; 0–2 = Short signal.

---

### Beneish M-Score

**Formula (Beneish 1999, 8-variable logistic model):**

M = −4.84 + 0.920×DSRI + 0.528×GMI + 0.404×AQI + 0.892×SGI + 0.115×DEPI − 0.172×SGAI + 4.679×TATA − 0.327×LVGI

| Variable | Measures |
|----------|----------|
| DSRI | Days Sales in Receivables Index |
| GMI | Gross Margin Index |
| AQI | Asset Quality Index |
| SGI | Sales Growth Index |
| DEPI | Depreciation Index |
| SGAI | SG&A Index |
| TATA | Total Accruals to Total Assets |
| LVGI | Leverage Index |

**Interpretation:**
- M > −2.22 → Potential manipulation
- M < −2.22 → Manipulation unlikely

**Nuances:** This is a screening flag, not a guarantee. Use as an input for deeper investigation.

---

### Ohlson Bankruptcy Probability

**Formula (Ohlson 1980, O-Score logistic model):**

Converts an O-Score into a probability: P = 1 / (1 + e^(−O))

**Interpretation:**
- 0% = minimal risk
- Values above 50% indicate serious distress

**Nuances:** Does not require market price data.

---

## 10. Global Engine Nuances

These apply across all ratio calculations:

### Derived Fields

When a field is not directly mapped by the user, the engine attempts derivation. **Critically, derivations only occur when ALL required inputs have data** — the engine never assumes a missing field is zero:

| Field | Derivation | Guard |
|-------|-----------|-------|
| Gross Profit | Revenue − Cost of Revenue | Both must have data |
| EBITDA | Operating Income + D&A | Both must have data |
| Pretax Income | Operating Income − Interest Expense | Both must have data |
| Free Cash Flow | Operating Cash Flow − Capital Expenditures | Both must have data |
| Total Debt | Short-Term Debt + Long-Term Debt (SUM mode) | At least one must have data |
| Market Cap | Share Price × Basic Shares Outstanding | Both must have data |

### Absolute Value Handling

Several fields are converted to absolute values before use, because different data sources report them with inconsistent signs:

- Cost of Revenue, Operating Expenses, D&A → absolute (always positive)
- Interest Expense → absolute (gross interest for coverage ratios)
- Capital Expenditures → absolute
- Total Debt, Short-Term Debt, Long-Term Debt → absolute
- Dividends Paid → absolute
- Preferred Dividends → absolute

### Average Calculations

For balance sheet items (stocks, not flows), the engine uses **two-period averages**: Avg(X) = (X_current + X_prior) ÷ 2.

For the **first period**, where no prior period exists, the engine uses the current period value directly. This means first-period return ratios (ROA, ROE, ROIC, etc.) use a point-in-time denominator rather than a true average — interpret with slight caution.

### Frequency Adjustments

For quarterly and monthly data:

- **Flow items** (Revenue, Net Income, EBITDA, OCF, FCF, CapEx, Dividends) are converted to trailing 12-month sums before use in annual-equivalent ratios.
- **Days calculations** use period-appropriate divisors: 91.25 days for quarterly, 30.42 for monthly (instead of 365).
- Trailing sum requires sufficient history — if fewer periods exist than the trailing window (4 for quarterly, 12 for monthly), those periods show N/A.

### Outlier Capping

| Metric Type | Cap Value | Rationale |
|-------------|-----------|-----------|
| Turnover ratios | ±100× | Extremely low denominators create unrealistic spikes |
| Interest/Cash Coverage | ±100× | Self-financed companies would show infinity |
| P/E Ratio | 500× (and no negatives) | Near-zero earnings create meaningless multiples |
| Growth rates | ±1000% | Near-zero base values create misleading spikes |

Values are **clipped** (not removed) — a company with 150× inventory turnover will show as 100×, preserving the signal that it is extremely high.

### All-NaN Filtering

After calculation, any ratio that yields N/A for **all periods** for a given company is excluded from that company's results. This prevents the results table from being cluttered with empty rows for ratios that simply lack the required input data.

---

*This guide reflects the formulas and logic in Financial Ratio Analysis Engine v3.0. All calculations are estimates — verify independently before making financial decisions.*
