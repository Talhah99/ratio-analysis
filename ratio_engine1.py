import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Optional
from dataclasses import dataclass
from enum import Enum

# Console-only logging — safe for Streamlit Cloud and any read-only filesystem.
# RotatingFileHandler was removed: it attempted to write 'ratio_engine.log' at
# module-import time, which raises PermissionError on cloud deployments before
# the app even renders a screen.
logger = logging.getLogger("ratio_engine")
logger.setLevel(logging.INFO)

if not logger.handlers:
    _console_handler = logging.StreamHandler()
    _console_handler.setLevel(logging.INFO)
    _console_handler.setFormatter(logging.Formatter('%(levelname)s | %(funcName)s | %(message)s'))
    logger.addHandler(_console_handler)

class Frequency(Enum):
    ANNUAL = "Annual"
    QUARTERLY = "Quarterly"
    MONTHLY = "Monthly"

@dataclass
class EngineConfig:
    """
    Configuration for ratio calculations.
    
    Attributes:
        pe_ceiling: Maximum P/E ratio before capping to NaN (default: 500.0)
        int_cov_cap: Maximum interest coverage ratio (default: 100.0)
        turnover_ceiling: Maximum turnover ratio before capping (default: 100.0)
        growth_cap: Maximum growth rate as decimal (default: 10.0 = 1000%)
        operating_cash_pct: Percentage of revenue to treat as operating cash for ROIC (default: 0.02 = 2%)
        use_excess_cash_adjustment: If True, subtract excess cash from invested capital (default: True)
        allow_negative_interest_coverage: If True, show negative coverage values for distressed companies (default: True)
        nopat_tax_benefit_on_losses: If True, NOPAT = OpInc×(1-t) even for losses — standard
            valuation. If False, no tax benefit on losses — conservative approach (default: True)
    """
    pe_ceiling: float = 500.0
    int_cov_cap: float = 100.0
    turnover_ceiling: float = 100.0
    growth_cap: float = 10.0
    # ROIC adjustments
    operating_cash_pct: float = 0.02
    use_excess_cash_adjustment: bool = True
    # Risk visibility
    allow_negative_interest_coverage: bool = True
    # NOPAT convention
    nopat_tax_benefit_on_losses: bool = True
    
    def __post_init__(self):
        """Validate configuration parameters"""
        if not (0 <= self.operating_cash_pct <= 1):
            raise ValueError(f"operating_cash_pct must be between 0 and 1, got {self.operating_cash_pct}")
        if self.pe_ceiling <= 0:
            raise ValueError(f"pe_ceiling must be positive, got {self.pe_ceiling}")
        if self.int_cov_cap <= 0:
            raise ValueError(f"int_cov_cap must be positive, got {self.int_cov_cap}")
        if self.turnover_ceiling <= 0:
            raise ValueError(f"turnover_ceiling must be positive, got {self.turnover_ceiling}")
        if self.growth_cap <= 0:
            raise ValueError(f"growth_cap must be positive, got {self.growth_cap}")


@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[str]
    warnings: List[str]
    info: List[str]

FIELD_MAP = {
    "Revenue": ["Revenue", "Net Sales", "Turnover", "Top Line", "Sales", "Total Revenue"],
    "Cost of Revenue": ["Cost of Revenue", "Cost of Sales", "COGS", "Direct Costs", "Cost of Goods Sold"],
    "Gross Profit": ["Gross Profit", "Gross Margin"],
    "Operating Expenses": ["Operating Expenses", "Opex", "SG&A", "Selling, General & Admin", "Operating Costs"],
    "Operating Income": ["Operating Income", "EBIT", "Operating Profit", "EBIT (Operating Income)"],
    "Interest Expense": ["Interest Expense", "Finance Cost", "Finance Costs", "Interest Cost"],
    "Interest Income": ["Interest Income", "Interest & Investment Income", "Finance Income"],
    "Pretax Income": ["Pretax Income", "EBT", "Profit Before Tax", "Income Before Tax"],
    "Income Tax Expense": ["Income Tax Expense", "Income Tax", "Tax Provision"],
    "Net Income": ["Net Income", "Net Profit", "Profit After Tax", "Net Earnings"],
    "EBITDA": ["EBITDA", "Adjusted EBITDA"],
    "D&A": ["Depreciation & Amortization", "D&A", "D&A For EBITDA", "Depreciation", "Depreciation and Amortization"],
    "Preferred Dividends": ["Preferred Dividends", "Dividends on Preferred Stock"],
    "Cash & Equivalents": ["Cash & Equivalents", "Cash & Cash Equivalents", "Cash", "Bank Balances"],
    "Short-Term Investments": ["Short-Term Investments", "Short Term Investments", "Marketable Securities", "ST Investments"],
    "Accounts Receivable": ["Accounts Receivable", "Trade Debtors", "Receivables", "Trade Receivables", "AR"],
    "Inventory": ["Inventory", "Stock-in-trade", "Stock", "Inventories"],
    "Prepaid Expenses": ["Prepaid Expenses", "Prepayments"],
    "Total Current Assets": ["Total Current Assets", "Current Assets"],
    "PP&E (Net)": ["Property, Plant & Equipment", "Fixed Assets", "Net Fixed Assets", "PP&E", "PPE"],
    "Intangible Assets": ["Intangible Assets", "Goodwill", "Patents"],
    "Total Assets": ["Total Assets", "Assets"],
    "Accounts Payable": ["Accounts Payable", "Trade Creditors", "Payables", "Trade Payables", "AP"],
    "Accrued Expenses": ["Accrued Expenses", "Accrued Liabilities"],
    "Short-Term Debt": {"synonyms": ["Short-Term Debt", "Current Portion of Long-Term Debt", "Current Portion of Leases", "Short Term Borrowings", "ST Debt", "Current Debt"], "mode": "SUM"},
    "Total Current Liabilities": ["Total Current Liabilities", "Current Liabilities"],
    "Long-Term Debt": {"synonyms": ["Long-Term Debt", "Long-Term Leases", "Non-Current Debt", "LT Debt"], "mode": "SUM"},
    "Total Debt": ["Total Debt", "Total Borrowings", "Debt"],
    "Total Liabilities": ["Total Liabilities", "Liabilities"],
    "Total Equity": ["Total Equity", "Shareholder's Equity", "Total Shareholders Equity", "Equity", "Shareholders Equity"],
    "Retained Earnings": ["Retained Earnings", "Accumulated Profit", "Accumulated Earnings"],
    "Minority Interest": ["Minority Interest", "Non-Controlling Interest", "NCI"],
    "Operating Cash Flow": ["Operating Cash Flow", "Net Cash from Operations", "CFO", "Cash from Operations"],
    "Capital Expenditures": ["Capital Expenditures", "CapEx", "Purchase of PP&E", "CAPEX"],
    "Free Cash Flow": ["Free Cash Flow", "FCF"],
    "Dividends Paid": ["Dividends Paid", "Cash Dividends", "Common Dividends Paid"],
    "Share Price": ["Share Price", "Last Close Price", "Price", "Stock Price"],
    "Shares Outstanding (Basic)": ["Shares Outstanding", "Total Common Shares Outstanding", "Basic Shares", "Basic Shares Outstanding"],
    "Shares Outstanding (Diluted)": ["Diluted Shares Outstanding", "Diluted Shares"],
    "Market Cap": ["Market Capitalization", "Market Cap", "Mkt Cap"],
    "Preferred Stock": ["Preferred Stock", "Preference Shares", "Pref Equity", "Preferred Equity"]
}

REQUIRED_FIELDS = {k: ", ".join(v["synonyms"] if isinstance(v, dict) else v) for k, v in FIELD_MAP.items()}
CRITICAL_FIELDS = ["Revenue", "Total Assets", "Total Equity"]

# Ratio requirements map: each ratio lists its required input fields.
# Used by the validation layer to determine which ratios can be computed.
RATIO_REQUIREMENTS = {
    "Current Ratio": ["Total Current Assets", "Total Current Liabilities"],
    "Quick Ratio": ["Total Current Assets", "Total Current Liabilities"],
    "Cash Ratio": ["Cash & Equivalents", "Total Current Liabilities"],
    "Defensive Interval (Days)": ["Cash & Equivalents", "Cost of Revenue", "Operating Expenses"],
    "NWC to Assets": ["Total Current Assets", "Total Current Liabilities", "Total Assets"],
    "Inventory Turnover": ["Cost of Revenue", "Inventory"],
    "Days Inventory (DIO)": ["Cost of Revenue", "Inventory"],
    "Receivables Turnover": ["Revenue", "Accounts Receivable"],
    "Days Sales Outstanding (DSO)": ["Revenue", "Accounts Receivable"],
    "Payables Turnover": ["Cost of Revenue", "Accounts Payable"],
    "Days Payables (DPO)": ["Cost of Revenue", "Accounts Payable"],
    "Cash Conversion Cycle": ["Revenue", "Accounts Receivable", "Cost of Revenue", "Accounts Payable"],
    "Total Asset Turnover": ["Revenue", "Total Assets"],
    "Fixed Asset Turnover": ["Revenue", "PP&E (Net)"],
    "Working Capital Turnover": ["Revenue", "Total Current Assets", "Total Current Liabilities"],
    "Debt to Equity": ["Total Debt", "Total Equity"],
    "Debt to Assets": ["Total Debt", "Total Assets"],
    "Debt to Capital": ["Total Debt", "Total Equity"],
    "Interest Coverage": ["Operating Income", "Interest Expense"],
    "Cash Interest Coverage": ["EBITDA", "Interest Expense"],
    "Net Debt to EBITDA": ["Total Debt", "Cash & Equivalents", "EBITDA"],
    "Financial Leverage": ["Total Assets", "Total Equity"],
    "Gross Margin": ["Gross Profit", "Revenue"],
    "Operating Margin": ["Operating Income", "Revenue"],
    "Net Margin": ["Net Income", "Revenue"],
    "EBITDA Margin": ["EBITDA", "Revenue"],
    "Pretax Margin": ["Pretax Income", "Revenue"],
    "ROA": ["Net Income", "Total Assets"],
    "ROE": ["Net Income", "Total Equity"],
    "ROIC": ["Operating Income", "Total Equity", "Total Debt"],
    "Cash ROIC": ["Free Cash Flow", "Total Equity", "Total Debt"],
    "DuPont: Net Margin": ["Net Income", "Revenue"],
    "DuPont: Asset Turnover": ["Revenue", "Total Assets"],
    "DuPont: Equity Multiplier": ["Total Assets", "Total Equity"],
    "OCF to Sales": ["Operating Cash Flow", "Revenue"],
    "FCF to Sales": ["Free Cash Flow", "Revenue"],
    "Quality of Income": ["Operating Cash Flow", "Net Income"],
    "Capex Coverage": ["Operating Cash Flow", "Capital Expenditures"],
    "Dividend Payout": ["Dividends Paid", "Net Income"],
    "FCF Conversion": ["Free Cash Flow", "EBITDA"],
    "EPS": ["Net Income", "Shares Outstanding (Diluted)"],
    "P/E Ratio": ["Share Price", "Net Income", "Shares Outstanding (Diluted)"],
    "PEG Ratio": ["Share Price", "Net Income", "Shares Outstanding (Diluted)"],
    "Book Value Per Share": ["Total Equity", "Shares Outstanding (Basic)"],
    "Price to Book (P/B)": ["Share Price", "Total Equity", "Shares Outstanding (Basic)"],
    "Price to Sales": ["Market Cap", "Revenue"],
    "EV / EBITDA": ["Market Cap", "Total Debt", "Cash & Equivalents", "EBITDA"],
    "EV / Revenue": ["Market Cap", "Total Debt", "Cash & Equivalents", "Revenue"],
    "FCF Yield": ["Free Cash Flow", "Market Cap"],
    "Revenue Growth": ["Revenue"],
    "Gross Profit Growth": ["Gross Profit"],
    "Operating Income Growth": ["Operating Income"],
    "EBITDA Growth": ["EBITDA"],
    "Net Income Growth": ["Net Income"],
    "EPS Growth": ["Net Income", "Shares Outstanding (Diluted)"],
    "FCF Growth": ["Free Cash Flow"],
    "Altman Z-Score": ["Total Current Assets", "Total Current Liabilities", "Total Assets", "Retained Earnings", "Operating Income", "Market Cap", "Total Liabilities", "Revenue"],
    "Altman Z-Score (EM Score)": ["Total Current Assets", "Total Current Liabilities", "Total Assets", "Retained Earnings", "Operating Income", "Total Equity", "Total Liabilities"],
    "Piotroski F-Score": ["Net Income", "Total Assets", "Operating Cash Flow", "Total Debt", "Total Current Assets", "Total Current Liabilities", "Revenue", "Gross Profit", "Shares Outstanding (Basic)"],
    "Beneish M-Score": ["Accounts Receivable", "Revenue", "Gross Profit", "Total Assets", "PP&E (Net)", "Total Current Assets", "D&A", "Operating Expenses", "Total Debt", "Net Income", "Operating Cash Flow"],
    "Ohlson Bankruptcy Prob": ["Total Assets", "Total Liabilities", "Total Current Assets", "Total Current Liabilities", "Net Income", "Operating Cash Flow", "Retained Earnings"],
    "Sloan Accrual Ratio": ["Net Income", "Operating Cash Flow", "Total Assets"],
    "CapEx to D&A": ["Capital Expenditures", "D&A"],
    "Degree of Operating Leverage": ["Operating Income", "Revenue"],
    "Net Debt to Equity": ["Total Debt", "Cash & Equivalents", "Total Equity"],
    "Earnings Yield (EBIT/EV)": ["Operating Income", "Market Cap", "Total Debt", "Cash & Equivalents"],
    "ROIC Spread (vs 10% Hurdle)": ["Operating Income", "Total Equity", "Total Debt"],
    "Interest Expense Ratio": ["Interest Expense", "Revenue"],
    "Reinvestment Rate": ["Capital Expenditures", "D&A", "Total Current Assets", "Total Current Liabilities", "Operating Income"],
    "Sustainable Growth Rate": ["Operating Income", "Total Equity", "Total Debt", "Capital Expenditures", "D&A"],
}


class RatioEngine:
    def __init__(self, df: pd.DataFrame, mapping: Dict[str, str] = None, industry: str = "Manufacturing", 
                 tax_rate: float = 0.29, frequency: str = "Annual", config: Optional[EngineConfig] = None):
        self.raw_df = df
        self.gui_mapping = mapping if mapping else {}
        self.industry = industry
        self.config = config or EngineConfig()
        
        try:
            self.frequency = Frequency(frequency)
        except ValueError:
            logger.warning(f"Invalid frequency '{frequency}', defaulting to Annual")
            self.frequency = Frequency.ANNUAL
        
        self._set_period_parameters()
        self._validate_tax_rate(tax_rate)
        
        self.validation_result = None
        self.data_quality = {}
        self.results = {}
        self._cache = {}
        self.calculation_audit = {}  # company -> {ratio_name: status_string}
        
        self._initialize_data()
        logger.info(f"Engine initialized: {len(self.companies)} companies, {self.num_years} periods, {self.frequency.value}")
    
    def _set_period_parameters(self):
        if self.frequency == Frequency.QUARTERLY:
            self.period_days = 91.25
            self.trailing_periods = 4
        elif self.frequency == Frequency.MONTHLY:
            self.period_days = 30.42
            self.trailing_periods = 12
        else:
            self.period_days = 365.0
            self.trailing_periods = 1
    
    def _validate_tax_rate(self, tax_rate: float):
        if not (0.0 <= tax_rate <= 1.0):
            clamped = max(0.0, min(1.0, tax_rate))
            logger.warning(f"Tax rate {tax_rate:.2f} outside [0.0, 1.0], clamped to {clamped:.2f}")
            self.tax_rate = clamped
        else:
            self.tax_rate = tax_rate
            if tax_rate == 0.0:
                logger.info("Tax rate 0%: NOPAT = Operating Income")
    
    def _sort_years(self, years: list) -> list:
        """Sort year/period columns chronologically.
        
        Handles:
        - Pure numeric years: 2019, 2017, 2018 → 2017, 2018, 2019
        - Already sorted: returns as-is
        - Non-numeric periods (e.g. Q1 2023): falls back to string sort
        
        Returns:
            Sorted list of year strings
        """
        if len(years) <= 1:
            return years
        
        # Try parsing as integers (most common: "2017", "2018", etc.)
        try:
            parsed = [(y, int(y)) for y in years]
            parsed.sort(key=lambda x: x[1])
            sorted_years = [y for y, _ in parsed]
            if sorted_years != years:
                logger.info(f"Year columns reordered chronologically: {years} → {sorted_years}")
            return sorted_years
        except (ValueError, TypeError):
            pass
        
        # Try parsing as floats (e.g. "2017.0" from Excel)
        try:
            parsed = [(y, float(y)) for y in years]
            parsed.sort(key=lambda x: x[1])
            sorted_years = [y for y, _ in parsed]
            if sorted_years != years:
                logger.info(f"Year columns reordered chronologically: {years} → {sorted_years}")
            return sorted_years
        except (ValueError, TypeError):
            pass
        
        # Fallback: string sort (handles "Q1 2023", "Q2 2023", etc.)
        sorted_years = sorted(years)
        if sorted_years != years:
            logger.info(f"Year columns reordered (string sort): {years} → {sorted_years}")
        return sorted_years

    def _initialize_data(self):
        if self.raw_df is None or self.raw_df.empty:
            raise ValueError("DataFrame is empty")
        if self.raw_df.shape[1] < 3:
            raise ValueError(f"Expected ≥3 columns, got {self.raw_df.shape[1]}")
        
        self.raw_df.columns = [str(x).strip() for x in self.raw_df.columns]
        year_cols = self.raw_df.columns[2:].tolist()
        # Filter out whitespace-only column names
        raw_years = [str(x).strip() for x in year_cols if str(x).strip()]
        
        if len(raw_years) == 0:
            raise ValueError("No year columns found")
        
        # Sort years chronologically and reorder DataFrame columns to match
        self.years = self._sort_years(raw_years)
        if self.years != raw_years:
            # Reorder DataFrame columns: keep first 2 columns, then sorted year columns
            base_cols = list(self.raw_df.columns[:2])
            self.raw_df = self.raw_df[base_cols + self.years]
        
        self.num_years = len(self.years)
        
        company_col = self.raw_df.iloc[:, 0]
        self.companies = [str(x).strip() for x in company_col.unique() 
                         if pd.notna(x) and str(x).strip() not in ['nan', '', 'None']]
        
        if not self.companies:
            raise ValueError("No valid companies found")
        
        self.company_dfs = {}
        for c in self.companies:
            mask = self.raw_df.iloc[:, 0].astype(str).str.strip() == c
            self.company_dfs[c] = self.raw_df[mask].copy()
    
    def _clean_number(self, value) -> float:
        if pd.isna(value):
            return np.nan
        if isinstance(value, (int, float)):
            return float(value)
        
        s = str(value).strip()
        non_numeric = ['-', '', 'nan', 'None', 'N/A', 'NA', '#N/A', '#DIV/0!', '#REF!', '#VALUE!', '#NUM!', '#NAME?', '#NULL!']
        if s.upper() in [x.upper() for x in non_numeric]:
            return np.nan
        
        s = s.replace(',', '').replace('%', '').replace('$', '').replace('€', '').replace('£', '').replace('¥', '').replace('₹', '').replace('₨', '')
        is_negative = '(' in s and ')' in s
        if is_negative:
            s = s.replace('(', '').replace(')', '')
        
        # Handle scientific notation (e.g., 1.5E+09, 2.3e-4)
        s = s.strip()
        
        try:
            f = float(s)
            return -f if is_negative else f
        except (ValueError, TypeError):
            return np.nan
    
    def _extract_series(self, company: str, field_key: str) -> np.ndarray:
        cache_key = f"{company}_{field_key}"
        if cache_key in self._cache:
            return self._cache[cache_key]
        
        config = FIELD_MAP.get(field_key)
        if not config:
            result = np.full(self.num_years, np.nan)
            self._cache[cache_key] = result
            return result
        
        # Parse configuration
        if isinstance(config, dict):
            synonyms = config["synonyms"]
            mode = config.get("mode", "SINGLE")
        else:
            synonyms = config
            mode = "SINGLE"
        
        # Prioritize GUI mapping
        if field_key in self.gui_mapping and self.gui_mapping[field_key]:
            gui_mapped = self.gui_mapping[field_key]
            if gui_mapped and str(gui_mapped).strip() and str(gui_mapped) != "(Not Mapped)":
                synonyms = [gui_mapped] + list(synonyms)
        
        comp_df = self.company_dfs.get(company)
        if comp_df is None or comp_df.empty:
            result = np.full(self.num_years, np.nan)
            self._cache[cache_key] = result
            return result
        
        comp_df = comp_df.copy()
        comp_df['_lookup'] = comp_df.iloc[:, 1].astype(str).str.strip().str.lower()
        
        # SUM mode: aggregate multiple line items
        if mode == "SUM":
            result = np.zeros(self.num_years, dtype=float)
            found_any = False
            periods_with_data = np.zeros(self.num_years, dtype=bool)
            
            for syn in synonyms:
                target = syn.lower()
                matches = comp_df[comp_df['_lookup'] == target]
                
                if not matches.empty:
                    raw_row = matches.iloc[0, 2:2+self.num_years].values
                    clean_row = np.array([self._clean_number(x) for x in raw_row])
                    valid_mask = ~np.isnan(clean_row)
                    
                    if np.any(valid_mask):
                        result[valid_mask] += clean_row[valid_mask]
                        periods_with_data |= valid_mask
                        found_any = True
            
            # Warn if multiple items summed (potential double-counting)
            if found_any:
                matched_count = sum(1 for syn in synonyms if not comp_df[comp_df['_lookup'] == syn.lower()].empty)
                if matched_count > 2:
                    logger.warning(f"{company} - {field_key}: {matched_count} items summed. Check for double-counting.")
            
            # Only mark as NaN if truly no data found
            result = np.where(periods_with_data, result, np.nan)
            
            if not found_any:
                result = np.full(self.num_years, np.nan)
            
            self._cache[cache_key] = result
            return result
        
        # SINGLE mode: find first matching synonym
        for syn in synonyms:
            target = syn.lower()
            matches = comp_df[comp_df['_lookup'] == target]
            
            if not matches.empty:
                raw_row = matches.iloc[0, 2:2+self.num_years].values
                clean_row = np.array([self._clean_number(x) for x in raw_row])
                self._cache[cache_key] = clean_row
                return clean_row
        
        # No match found
        result = np.full(self.num_years, np.nan)
        self._cache[cache_key] = result
        return result
    
    def _safe_div(self, numerator, denominator, fill_value=np.nan, max_ratio=1e12) -> np.ndarray:
        """Safe division with comprehensive edge case handling"""
        num = np.asarray(numerator, dtype=float)
        den = np.asarray(denominator, dtype=float)
        
        with np.errstate(divide='ignore', invalid='ignore'):
            result = np.divide(num, den)
            
            # Handle infinities and invalid values
            result = np.where(np.isfinite(result), result, fill_value)
            
            # Cap extreme ratios
            result = np.where(np.abs(result) > max_ratio, fill_value, result)
            
            # Handle negative denominator edge cases
            result = np.where((den == 0) | np.isnan(den), fill_value, result)
            
        return result

    def _safe_add(self, *arrays) -> np.ndarray:
        """Safe addition handling NaN propagation intelligently"""
        if not arrays:
            return np.array([])
        
        arrays = [np.asarray(arr, dtype=float) for arr in arrays]
        
        if len(arrays) == 0:
            return np.array([])
        
        result = np.zeros_like(arrays[0], dtype=float)
        has_any_value = np.zeros_like(arrays[0], dtype=bool)
        
        for arr in arrays:
            valid_mask = ~np.isnan(arr)
            result = np.where(valid_mask, result + arr, result)
            has_any_value |= valid_mask
        
        return np.where(has_any_value, result, np.nan)

    def _avg(self, arr: np.ndarray) -> np.ndarray:
        """Compute period average, handling edge cases"""
        arr = np.asarray(arr, dtype=float)
        result = np.full_like(arr, np.nan, dtype=float)
        
        for i in range(len(arr)):
            if i == 0:
                # First period: use current value if available
                result[i] = arr[i] if not np.isnan(arr[i]) else np.nan
            else:
                curr, prev = arr[i], arr[i-1]
                both_valid = not np.isnan(curr) and not np.isnan(prev)
                
                if both_valid:
                    # Normal case: average of current and previous
                    result[i] = (curr + prev) / 2.0
                elif not np.isnan(curr):
                    # Only current available: use it
                    result[i] = curr
                elif not np.isnan(prev):
                    # Only previous available: use it (carry forward)
                    result[i] = prev
                else:
                    # Neither available
                    result[i] = np.nan
        
        return result

    def _growth(self, arr: np.ndarray) -> np.ndarray:
        """Calculate YoY growth with robust edge case handling
        
        Handles turnaround scenarios (negative to positive) by showing capped growth
        rather than suppressing the signal entirely.
        """
        arr = np.asarray(arr, dtype=float)
        
        if len(arr) < 2:
            return arr
        
        result = np.full_like(arr, np.nan, dtype=float)
        prev = arr[:-1]
        curr = arr[1:]
        
        # Only calculate where both values are valid
        valid_mask = ~np.isnan(prev) & ~np.isnan(curr)
        
        # Handle zero and near-zero denominators
        non_zero_mask = valid_mask & (np.abs(prev) > 1e-10)
        
        growth = np.full_like(curr, np.nan, dtype=float)
        growth[non_zero_mask] = (curr[non_zero_mask] - prev[non_zero_mask]) / np.abs(prev[non_zero_mask])
        
        # Cap extreme growth rates (but preserve sign for visibility)
        growth = np.where(np.abs(growth) > self.config.growth_cap, 
                         np.sign(growth) * self.config.growth_cap, growth)
        
        # For sign changes (turnarounds): show as capped value rather than NaN
        # This is important for distressed companies that return to profitability
        # Previously this was suppressed to NaN, hiding important turnaround signals
        
        result[1:] = growth
        
        return result

    def _cap_outliers(self, arr: np.ndarray, ceiling: float) -> np.ndarray:
        """Cap extreme outliers while preserving valid data.
        
        Returns values clipped within [-ceiling, ceiling] instead of nullifying them,
        so that extremely healthy companies remain visible on charts.
        """
        arr = np.asarray(arr, dtype=float)
        with np.errstate(invalid='ignore'):
            # Clip between -ceiling and +ceiling, preserving NaNs
            return np.clip(arr, -ceiling, ceiling)

    def _trailing_sum(self, arr: np.ndarray) -> np.ndarray:
        """Calculate trailing sum for frequency adjustments"""
        if self.frequency == Frequency.ANNUAL:
            return arr
        
        arr = np.asarray(arr, dtype=float)
        result = np.full_like(arr, np.nan, dtype=float)
        
        for i in range(len(arr)):
            if i < self.trailing_periods - 1:
                # Not enough history yet
                result[i] = np.nan
            else:
                window = arr[i - self.trailing_periods + 1 : i + 1]
                if not np.any(np.isnan(window)):
                    total = np.sum(window)
                    # Guard against overflow - set to NaN if result is not finite
                    result[i] = total if np.isfinite(total) else np.nan
                else:
                    result[i] = np.nan
        
        return result
    
    def validate_data(self) -> ValidationResult:
        errors, warnings, info = [], [], []
        mapped_fields = [k for k, v in self.gui_mapping.items() if v]
        missing_critical = [f for f in CRITICAL_FIELDS if f not in mapped_fields]
        
        # Critical field mapping check
        if missing_critical:
            errors.append(f"CRITICAL FIELDS NOT MAPPED: {', '.join(missing_critical)}")
            return ValidationResult(False, errors, warnings, info)
        
        if not mapped_fields:
            errors.append("No fields mapped")
            return ValidationResult(False, errors, warnings, info)
        
        info.append(f"✓ {len(mapped_fields)} fields mapped")
        
        # Company-level validation - check critical fields have data
        if mapped_fields:
            for company in self.companies:
                # Check critical fields first
                critical_data_missing = []
                for critical_field in CRITICAL_FIELDS:
                    if critical_field in mapped_fields:
                        data = self._extract_series(company, critical_field)
                        if np.all(np.isnan(data)):
                            critical_data_missing.append(critical_field)
                
                if critical_data_missing:
                    errors.append(f"Company '{company}': Missing critical data for {', '.join(critical_data_missing)}")
                    continue
                
                # Calculate overall completeness for info/warnings
                total_points = available_points = 0
                for field in mapped_fields:
                    data = self._extract_series(company, field)
                    total_points += len(data)
                    available_points += np.sum(~np.isnan(data))
                
                completeness_pct = (available_points / total_points * 100) if total_points > 0 else 0
                
                if total_points > 0 and available_points == 0:
                    errors.append(f"Company '{company}': No data available")
                elif completeness_pct < 30:
                    warnings.append(f"Company '{company}': Very sparse data ({completeness_pct:.0f}%) - results may be limited")
                elif completeness_pct < 50:
                    warnings.append(f"Company '{company}': Partial data ({completeness_pct:.0f}%) - some ratios unavailable")
                elif completeness_pct < 70:
                    info.append(f"Company '{company}': Adequate data ({completeness_pct:.0f}%)")
                else:
                    info.append(f"✓ Company '{company}': Strong data ({completeness_pct:.0f}%)")
        
        # Frequency check
        if self.frequency != Frequency.ANNUAL and self.num_years < self.trailing_periods:
            warnings.append(f"{self.frequency.value} needs {self.trailing_periods} periods for trailing calculations, you have {self.num_years}")
        
        # Summary
        if not errors:
            info.append("✓ Validation passed - ready for calculation")
        
        self.validation_result = ValidationResult(len(errors) == 0, errors, warnings, info)
        return self.validation_result
    
    def _calc_liquidity(self, data: dict) -> dict:
        r = {}
        tca, tcl = data['tca'], data['tcl']
        r["Current Ratio"] = self._safe_div(tca, tcl)
        quick_assets = self._safe_add(data['cash'], data['st_inv'], data['receivables'])
        if np.all(np.isnan(quick_assets)):
            # Fallback: TCA minus non-liquid items.
            # Per strict quick ratio definition, both inventory AND prepaid expenses
            # are excluded — prepaid cannot be converted to cash to meet obligations.
            quick_assets = self._safe_add(tca, -data['inventory'], -data['prepaid'])
        r["Quick Ratio"] = self._safe_div(quick_assets, tcl)
        cash_equivalents = self._safe_add(data['cash'], data['st_inv'])
        r["Cash Ratio"] = self._safe_div(cash_equivalents, tcl)
        # Defensive Interval: guard against negative cash expenses (e.g., D&A > COGS+OpEx)
        cash_expenses = self._safe_add(data['cogs'], data['opex'], -data['da'])
        cash_expenses = np.maximum(cash_expenses, 0)  # Prevent negative expenses
        daily_burn = self._safe_div(cash_expenses, self.period_days)
        daily_burn = np.where(daily_burn <= 0, np.nan, daily_burn)  # No burn = infinite interval
        r["Defensive Interval (Days)"] = self._safe_div(quick_assets, daily_burn)
        r["NWC to Assets"] = self._safe_div(data['wc'], data['ta'])
        return r
    
    def _calc_activity(self, data: dict) -> dict:
        r = {}
        rev, cogs = data['rev'], data['cogs']
        trailing_rev = self._trailing_sum(rev)
        trailing_cogs = self._trailing_sum(cogs)
        
        if self.industry == "Manufacturing":
            inv_turn = self._safe_div(cogs, data['avg_inv'])
            r["Inventory Turnover"] = self._cap_outliers(inv_turn, self.config.turnover_ceiling)
            r["Days Inventory (DIO)"] = self._safe_div(data['avg_inv'], cogs) * self.period_days
            if self.frequency != Frequency.ANNUAL:
                inv_turn_annual = self._safe_div(trailing_cogs, data['avg_inv'])
                r["Inventory Turnover (Annual)"] = self._cap_outliers(inv_turn_annual, self.config.turnover_ceiling)
        
        ar_turn = self._safe_div(rev, data['avg_ar'])
        r["Receivables Turnover"] = self._cap_outliers(ar_turn, self.config.turnover_ceiling)
        r["Days Sales Outstanding (DSO)"] = self._safe_div(data['avg_ar'], rev) * self.period_days
        # For service industries with near-zero COGS, use operating expenses as denominator for DPO
        if self.industry != "Manufacturing":
            # Service firms: use OpEx as proxy if COGS is negligible
            cogs_for_payables = np.where(
                np.abs(cogs) < 1e-10,
                data.get('opex', cogs),  # Fallback to opex
                cogs
            )
        else:
            cogs_for_payables = cogs
        ap_turn = self._safe_div(cogs_for_payables, data['avg_ap'])
        r["Payables Turnover"] = self._cap_outliers(ap_turn, self.config.turnover_ceiling)
        r["Days Payables (DPO)"] = self._safe_div(data['avg_ap'], cogs_for_payables) * self.period_days
        
        dso = r.get("Days Sales Outstanding (DSO)")
        dpo = r.get("Days Payables (DPO)")
        if self.industry == "Manufacturing":
            dio = r.get("Days Inventory (DIO)")
            if dso is not None and dio is not None and dpo is not None:
                r["Cash Conversion Cycle"] = self._safe_add(dso, dio, -dpo)
        else:
            if dso is not None and dpo is not None:
                r["Cash Conversion Cycle"] = self._safe_add(dso, -dpo)
        
        asset_turn_rev = trailing_rev if self.frequency != Frequency.ANNUAL else rev
        r["Total Asset Turnover"] = self._safe_div(asset_turn_rev, data['avg_ta'])
        r["Fixed Asset Turnover"] = self._safe_div(asset_turn_rev, data['avg_ppe'])
        wc = data['wc']
        wc_safe = np.where(wc > 0, wc, np.nan)
        r["Working Capital Turnover"] = self._safe_div(asset_turn_rev, wc_safe)
        return r
    
    def _calc_solvency(self, data: dict) -> dict:
        r = {}
        td, eq = data['td'], data['equity']
        de = self._safe_div(td, eq)
        # Suppress D/E when equity is negative: the ratio sign flips (negative D/E
        # implies debt is being "offset" by negative equity), which is misleading.
        # Consistent with Financial Leverage, which also suppresses negative equity.
        # The distress signal itself is visible via Interest Coverage and Net Debt/EBITDA.
        de = np.where(eq <= 0, np.nan, de)
        r["Debt to Equity"] = de
        r["Debt to Assets"] = self._safe_div(td, data['ta'])
        total_capital = self._safe_add(td, eq)
        # Suppress Debt/Capital when equity is negative: total capital becomes less than
        # total debt (or negative), making the ratio > 1 or negative — both are misleading
        # as a leverage metric. Distress is better captured by D/E (already suppressed) and
        # Net Debt/EBITDA. Guard: capital must be positive and >= debt.
        dtc = self._safe_div(td, total_capital)
        dtc = np.where(
            (eq <= 0) | (total_capital <= 0) | (total_capital < td),
            np.nan, dtc
        )
        r["Debt to Capital"] = dtc
        # Interest Coverage: handle zero interest (no debt - company is self-financed)
        int_cov = np.where(
            data['int_exp'] == 0,
            np.where(data['op_inc'] > 0, self.config.int_cov_cap, np.nan),  # Cap if positive income, 0 interest
            self._safe_div(data['op_inc'], data['int_exp'])
        )
        # Handle negative operating income with interest expense
        if not self.config.allow_negative_interest_coverage:
            # Original behavior: suppress negative coverage to NaN
            int_cov = np.where((data['op_inc'] < 0) & (data['int_exp'] > 0), np.nan, int_cov)
        # else: allow negative values to flow through for risk visibility
        r["Interest Coverage"] = self._cap_outliers(int_cov, self.config.int_cov_cap)
        ebitda_use = self._trailing_sum(data['ebitda']) if self.frequency != Frequency.ANNUAL else data['ebitda']
        # Cash Interest Coverage: same zero-interest handling as regular Interest Coverage
        cash_int_cov = np.where(
            data['int_exp'] == 0,
            np.where(ebitda_use > 0, self.config.int_cov_cap, np.nan),  # Cap if positive EBITDA, 0 interest
            self._safe_div(ebitda_use, data['int_exp'])
        )
        # Handle negative EBITDA with interest expense
        if not self.config.allow_negative_interest_coverage:
            cash_int_cov = np.where((ebitda_use < 0) & (data['int_exp'] > 0), np.nan, cash_int_cov)
        r["Cash Interest Coverage"] = self._cap_outliers(cash_int_cov, self.config.int_cov_cap)
        r["Net Debt to EBITDA"] = self._safe_div(data['net_debt'], ebitda_use)
        fin_lev = self._safe_div(data['avg_ta'], data['avg_equity'])
        fin_lev = np.where(data['avg_equity'] <= 0, np.nan, fin_lev)
        r["Financial Leverage"] = fin_lev
        return r
    
    def _calc_profitability(self, data: dict) -> dict:
        r = {}
        rev = data['rev']
        r["Gross Margin"] = self._safe_div(data['gp'], rev)
        r["Operating Margin"] = self._safe_div(data['op_inc'], rev)
        r["Net Margin"] = self._safe_div(data['net_inc'], rev)
        r["EBITDA Margin"] = self._safe_div(data['ebitda'], rev)
        r["Pretax Margin"] = self._safe_div(data['pretax'], rev)
        return r
    
    def _calc_returns(self, data: dict) -> dict:
        r = {}
        if self.frequency == Frequency.ANNUAL:
            net_inc_use = data['net_inc']
            nopat_use = data['nopat']
            fcf_use = data['fcf']
            rev_use = data['rev']
        else:
            net_inc_use = self._trailing_sum(data['net_inc'])
            nopat_use = self._trailing_sum(data['nopat'])
            fcf_use = self._trailing_sum(data['fcf'])
            rev_use = self._trailing_sum(data['rev'])
        
        avg_eq = data['avg_equity']
        r["ROA"] = self._safe_div(net_inc_use, data['avg_ta'])
        roe_gaap = self._safe_div(net_inc_use, avg_eq)
        roe_gaap = np.where(avg_eq <= 0, np.nan, roe_gaap)
        r["ROE"] = roe_gaap
        roe_norm = self._safe_div(nopat_use, avg_eq)
        roe_norm = np.where(avg_eq <= 0, np.nan, roe_norm)
        r["ROE (Normalized)"] = roe_norm
        r["ROIC"] = self._safe_div(nopat_use, data['avg_ic'])
        r["Cash ROIC"] = self._safe_div(fcf_use, data['avg_ic'])
        
        # DuPont Decomposition: ALL components must use consistent annualized figures
        # ROE = Net Margin × Asset Turnover × Equity Multiplier
        r["DuPont: Net Margin"] = self._safe_div(net_inc_use, rev_use)
        r["DuPont: Asset Turnover"] = self._safe_div(rev_use, data['avg_ta'])
        # Guard against None or empty avg_eq array
        if avg_eq is None or len(avg_eq) == 0:
            dupont_em = np.full(self.num_years, np.nan)
        else:
            dupont_em = self._safe_div(data['avg_ta'], avg_eq)
            dupont_em = np.where(avg_eq <= 0, np.nan, dupont_em)
        r["DuPont: Equity Multiplier"] = dupont_em
        return r
    
    def _calc_cashflow(self, data: dict) -> dict:
        r = {}
        if self.frequency != Frequency.ANNUAL:
            ocf = self._trailing_sum(data['ocf'])
            fcf = self._trailing_sum(data['fcf'])
            ebitda = self._trailing_sum(data['ebitda'])
            net_inc = self._trailing_sum(data['net_inc'])
            rev = self._trailing_sum(data['rev'])
            capex = self._trailing_sum(data['capex'])  # FIX #11: trailing-sum capex for quarterly
            divs = self._trailing_sum(data['divs'])    # FIX #3: trailing-sum dividends for quarterly
        else:
            ocf = data['ocf']
            fcf = data['fcf']
            ebitda = data['ebitda']
            net_inc = data['net_inc']
            rev = data['rev']
            capex = data['capex']
            divs = data['divs']
        r["OCF to Sales"] = self._safe_div(ocf, rev)
        r["FCF to Sales"] = self._safe_div(fcf, rev)
        r["Quality of Income"] = self._safe_div(ocf, net_inc)
        r["Capex Coverage"] = self._safe_div(ocf, capex)
        payout = self._safe_div(divs, net_inc)
        payout = np.where(net_inc <= 0, np.nan, payout)
        r["Dividend Payout"] = payout
        # FIX #15: FCF Conversion meaningless when EBITDA <= 0
        fcf_conv = self._safe_div(fcf, ebitda)
        fcf_conv = np.where(ebitda <= 0, np.nan, fcf_conv)
        r["FCF Conversion"] = fcf_conv
        return r
    
    def _calc_valuation(self, data: dict) -> dict:
        r = {}
        if self.frequency != Frequency.ANNUAL:
            net_inc_use = self._trailing_sum(data['net_inc'])
            rev_use = self._trailing_sum(data['rev'])
            ebitda_use = self._trailing_sum(data['ebitda'])
            fcf_use = self._trailing_sum(data['fcf'])
        else:
            net_inc_use = data['net_inc']
            rev_use = data['rev']
            ebitda_use = data['ebitda']
            fcf_use = data['fcf']
        
        # FIX #2: EPS uses trailing (annualized) income consistently
        pref_div_use = self._trailing_sum(data['pref_div']) if self.frequency != Frequency.ANNUAL else data['pref_div']
        eps = self._safe_div(self._safe_add(net_inc_use, -pref_div_use), data['shares_diluted'])
        r["EPS"] = eps
        pe = self._safe_div(data['price'], eps)
        pe = np.where((pe < 0) | (pe > self.config.pe_ceiling), np.nan, pe)
        r["P/E Ratio"] = pe
        eps_growth = self._growth(eps)
        peg = self._safe_div(pe, eps_growth * 100)
        peg = np.where(eps_growth <= 0, np.nan, peg)
        r["PEG Ratio"] = peg
        bvps = self._safe_div(data['equity'], data['shares'])
        r["Book Value Per Share"] = bvps
        pb = self._safe_div(data['price'], bvps)
        pb = np.where(bvps <= 0, np.nan, pb)
        r["Price to Book (P/B)"] = pb
        r["Price to Sales"] = self._safe_div(data['mc'], rev_use)
        r["EV / EBITDA"] = self._safe_div(data['ev'], ebitda_use)
        r["EV / Revenue"] = self._safe_div(data['ev'], rev_use)
        r["FCF Yield"] = self._safe_div(fcf_use, data['mc'])
        return r
    
    def _calc_growth(self, data: dict) -> dict:
        # FIX #2: EPS Growth uses the annualized EPS from valuation (stored in data)
        # For other metrics, growth is on raw period data (correct - shows period-over-period change)
        return {
            "Revenue Growth": self._growth(data['rev']),
            "Gross Profit Growth": self._growth(data['gp']),
            "Operating Income Growth": self._growth(data['op_inc']),
            "EBITDA Growth": self._growth(data['ebitda']),
            "Net Income Growth": self._growth(data['net_inc']),
            "EPS Growth": self._growth(data['eps_annualized']),
            "FCF Growth": self._growth(data['fcf'])
        }
    
    def _calc_zscore_public(self, data: dict) -> np.ndarray:
        rev_use = self._trailing_sum(data['rev']) if self.frequency != Frequency.ANNUAL else data['rev']
        A = self._safe_div(data['wc'], data['ta'])
        B = self._safe_div(data['retained_earnings'], data['ta'])
        C = self._safe_div(data['op_inc'], data['ta'])
        D = self._safe_div(data['mc'], data['total_liab'])
        D = np.where(data['total_liab'] <= 0, np.nan, D)  # Invalid if no/negative liabilities
        E = self._safe_div(rev_use, data['ta'])
        z_components = [1.2 * A, 1.4 * B, 3.3 * C, 0.6 * D, 1.0 * E]
        z_score = self._safe_add(*z_components)
        z_score = np.where(np.isnan(data['mc']), np.nan, z_score)
        return z_score
    
    def _calc_zscore_private(self, data: dict) -> np.ndarray:
        A = self._safe_div(data['wc'], data['ta'])
        B = self._safe_div(data['retained_earnings'], data['ta'])
        C = self._safe_div(data['op_inc'], data['ta'])
        # Guard against zero/negative liabilities
        D = self._safe_div(data['equity'], data['total_liab'])
        D = np.where(data['total_liab'] <= 0, np.nan, D)  # Invalid if no/negative liabilities
        z_components = [6.56 * A, 3.26 * B, 6.72 * C, 1.05 * D]
        z_score = self._safe_add(*z_components)
        return z_score
    
    def _calc_quality_scores(self, data: dict) -> dict:
        """
        Piotroski F-Score, Beneish M-Score, Sloan Accrual Ratio, Ohlson O-Score.

        Piotroski F-Score (0–9)
        -----------------------
        Binary scoring system across three dimensions. Each signal scores 0 or 1.
        Profitability (4): ROA>0, ΔROA>0, OCF/TA>0, Accruals<0 (OCF/TA > ROA)
        Leverage/Liquidity (3): ΔLeverage<0, ΔCurrent Ratio>0, No new shares issued
        Efficiency (2): ΔGross Margin>0, ΔAsset Turnover>0
        Score 8–9 = Strong; 5–7 = Neutral; 0–2 = Weak

        Beneish M-Score
        ---------------
        8-variable logistic model detecting earnings manipulation.
        M < −2.22: Manipulation unlikely  |  M > −2.22: Manipulation possible
        Coefficients: Beneish (1999) "The Detection of Earnings Manipulation"

        Sloan Accrual Ratio
        --------------------
        = (Net Income − Operating Cash Flow) / Average Total Assets
        High positive accruals (> 5%) signal low earnings quality.
        Source: Sloan (1996) "Do Stock Prices Fully Reflect Information in Accruals?"

        Ohlson O-Score (Simplified)
        ----------------------------
        Logistic bankruptcy predictor. Probability = 1 / (1 + exp(−O)).
        Source: Ohlson (1980) "Financial Ratios and the Probabilistic Prediction of Bankruptcy"
        Note: GNP price-level normalisation omitted; TA used at book value directly.
        """
        r = {}
        rev     = data['rev']
        gp      = data['gp']
        op_inc  = data['op_inc']
        net_inc = data['net_inc']
        ocf     = data['ocf']
        ta      = data['ta']
        tca     = data['tca']
        tcl     = data['tcl']
        td      = data['td']
        total_liab = data['total_liab']
        equity  = data['equity']
        ppe     = data['ppe']
        da      = data['da']
        capex   = data['capex']
        wc      = data['wc']
        avg_ta  = data['avg_ta']
        shares  = data['shares']
        opex    = data['opex']
        retained_earnings = data['retained_earnings']

        n = len(rev)

        # ── Sloan Accrual Ratio ────────────────────────────────────────────
        # = (Net Income - OCF) / avg_TA
        # Positive = accruals exceed cash earnings (low quality)
        # Negative = cash earnings exceed accounting profits (high quality)
        accrual_num = np.where(
            ~np.isnan(net_inc) & ~np.isnan(ocf),
            net_inc - ocf, np.nan
        )
        r["Sloan Accrual Ratio"] = self._safe_div(accrual_num, avg_ta)

        # ── CapEx to D&A ───────────────────────────────────────────────────
        # > 1: Net investment (growing assets)  ~1: Maintenance  < 1: Asset harvesting
        r["CapEx to D&A"] = np.where(
            (da > 0) & ~np.isnan(da) & ~np.isnan(capex),
            capex / da, np.nan
        )

        # ── Degree of Operating Leverage (DOL) ────────────────────────────
        # DOL = %ΔEBIT / %ΔRevenue = (ΔEBIT / EBIT_{t-1}) / (ΔRevenue / Revenue_{t-1})
        # Values > 3 indicate high fixed-cost structure; < 1 is unusual (variable cost biz)
        # First year is always NaN (no prior year)
        dol = np.full(n, np.nan)
        for i in range(1, n):
            prev_rev  = rev[i-1];  curr_rev  = rev[i]
            prev_ebit = op_inc[i-1]; curr_ebit = op_inc[i]
            if (not np.isnan(prev_rev) and not np.isnan(curr_rev) and
                not np.isnan(prev_ebit) and not np.isnan(curr_ebit) and
                abs(prev_rev) > 1e-10 and abs(prev_ebit) > 1e-10):
                pct_rev  = (curr_rev  - prev_rev)  / abs(prev_rev)
                pct_ebit = (curr_ebit - prev_ebit) / abs(prev_ebit)
                if abs(pct_rev) > 1e-6:
                    dol_val = pct_ebit / pct_rev
                    # Cap to avoid division-near-zero artifacts
                    dol[i] = np.clip(dol_val, -20.0, 20.0)
        r["Degree of Operating Leverage"] = dol

        # ── Piotroski F-Score ─────────────────────────────────────────────
        # Binary signals: 1 if favourable, 0 otherwise, NaN if data missing
        # Year 0 is always NaN (no prior year for delta signals)
        fscore = np.full(n, np.nan)
        for i in range(1, n):

            def _get_val(arr, idx):
                v = arr[idx] if idx < len(arr) else np.nan
                return v if not np.isnan(v) else np.nan

            # Profitability signals
            roa_curr  = (_get_val(net_inc, i)   / _get_val(ta, i)
                         if not np.isnan(_get_val(ta, i)) and abs(_get_val(ta, i)) > 1e-10
                         else np.nan)
            roa_prev  = (_get_val(net_inc, i-1) / _get_val(ta, i-1)
                         if not np.isnan(_get_val(ta, i-1)) and abs(_get_val(ta, i-1)) > 1e-10
                         else np.nan)
            ocf_ta    = (_get_val(ocf, i) / _get_val(ta, i)
                         if not np.isnan(_get_val(ta, i)) and abs(_get_val(ta, i)) > 1e-10
                         else np.nan)

            f1 = float(roa_curr > 0)   if not np.isnan(roa_curr)  else np.nan  # ROA positive
            f2 = float(ocf_ta > 0)     if not np.isnan(ocf_ta)    else np.nan  # OCF/TA positive
            f3 = (float(roa_curr > roa_prev)                                    # ROA improving
                  if not np.isnan(roa_curr) and not np.isnan(roa_prev) else np.nan)
            f4 = (float(ocf_ta > roa_curr)                                      # Accruals quality
                  if not np.isnan(ocf_ta) and not np.isnan(roa_curr) else np.nan)

            # Leverage / Liquidity signals
            lev_curr = (_get_val(td, i)   / _get_val(ta, i)   if abs(_get_val(ta, i)) > 1e-10 else np.nan)
            lev_prev = (_get_val(td, i-1) / _get_val(ta, i-1) if abs(_get_val(ta, i-1)) > 1e-10 else np.nan)
            cr_curr  = (_get_val(tca, i)  / _get_val(tcl, i)  if abs(_get_val(tcl, i)) > 1e-10 else np.nan)
            cr_prev  = (_get_val(tca, i-1)/_get_val(tcl, i-1) if abs(_get_val(tcl, i-1)) > 1e-10 else np.nan)
            sh_curr  = _get_val(shares, i)
            sh_prev  = _get_val(shares, i-1)

            f5 = (float(lev_curr < lev_prev)                  # Leverage decreased
                  if not np.isnan(lev_curr) and not np.isnan(lev_prev) else np.nan)
            f6 = (float(cr_curr > cr_prev)                    # Current ratio improved
                  if not np.isnan(cr_curr) and not np.isnan(cr_prev) else np.nan)
            f7 = (float(sh_curr <= sh_prev * 1.02)            # No meaningful share dilution (>2%)
                  if not np.isnan(sh_curr) and not np.isnan(sh_prev) and sh_prev > 0 else np.nan)

            # Operating Efficiency signals
            gm_curr  = (_get_val(gp, i)   / _get_val(rev, i)   if abs(_get_val(rev, i)) > 1e-10 else np.nan)
            gm_prev  = (_get_val(gp, i-1) / _get_val(rev, i-1) if abs(_get_val(rev, i-1)) > 1e-10 else np.nan)
            at_curr  = (_get_val(rev, i)   / _get_val(ta, i)   if abs(_get_val(ta, i)) > 1e-10 else np.nan)
            at_prev  = (_get_val(rev, i-1) / _get_val(ta, i-1) if abs(_get_val(ta, i-1)) > 1e-10 else np.nan)

            f8 = (float(gm_curr > gm_prev)                    # Gross margin improved
                  if not np.isnan(gm_curr) and not np.isnan(gm_prev) else np.nan)
            f9 = (float(at_curr > at_prev)                    # Asset turnover improved
                  if not np.isnan(at_curr) and not np.isnan(at_prev) else np.nan)

            signals = [f1, f2, f3, f4, f5, f6, f7, f8, f9]
            valid_signals = [s for s in signals if not np.isnan(s)]
            # Only compute score if at least 7 of 9 signals available
            if len(valid_signals) >= 7:
                fscore[i] = sum(valid_signals)

        r["Piotroski F-Score"] = fscore

        # ── Beneish M-Score ───────────────────────────────────────────────
        # Logistic model. Requires at least 2 years of data.
        # Variables are year-over-year indices.
        # M = −4.84 + 0.920×DSRI + 0.528×GMI + 0.404×AQI + 0.892×SGI
        #         + 0.115×DEPI − 0.172×SGAI + 4.679×TATA − 0.327×LVGI
        mscore = np.full(n, np.nan)
        for i in range(1, n):
            def _v(arr, idx):
                val = arr[idx] if idx < len(arr) else np.nan
                return float(val) if not np.isnan(val) else None

            ar_c  = _v(data['receivables'], i);  ar_p  = _v(data['receivables'], i-1)
            rev_c = _v(rev, i);                  rev_p = _v(rev, i-1)
            gp_c  = _v(gp, i);                   gp_p  = _v(gp, i-1)
            ta_c  = _v(ta, i);                   ta_p  = _v(ta, i-1)
            ppe_c = _v(ppe, i);                  ppe_p = _v(ppe, i-1)
            tca_c = _v(tca, i);                  tca_p = _v(tca, i-1)
            da_c  = _v(da, i);                   da_p  = _v(da, i-1)
            opex_c= _v(opex, i);                 opex_p= _v(opex, i-1)
            td_c  = _v(td, i);                   td_p  = _v(td, i-1)
            ni_c  = _v(net_inc, i)
            ocf_c = _v(ocf, i)

            # DSRI: Days Sales Receivable Index = (AR_t/Sales_t) / (AR_{t-1}/Sales_{t-1})
            dsri = (None if None in [ar_c, rev_c, ar_p, rev_p] or
                    rev_c == 0 or rev_p == 0
                    else (ar_c / rev_c) / (ar_p / rev_p))

            # GMI: Gross Margin Index = GM_{t-1} / GM_t
            gmi = (None if None in [gp_c, gp_p, rev_c, rev_p] or
                   rev_c == 0 or rev_p == 0 or gp_c / rev_c == 0
                   else (gp_p / rev_p) / (gp_c / rev_c))

            # AQI: Asset Quality Index = (1 - (CA_t + PPE_t)/TA_t) / (1 - (CA_{t-1} + PPE_{t-1})/TA_{t-1})
            if None not in [tca_c, ppe_c, ta_c, tca_p, ppe_p, ta_p] and ta_c != 0 and ta_p != 0:
                aq_c = 1 - (tca_c + ppe_c) / ta_c
                aq_p = 1 - (tca_p + ppe_p) / ta_p
                aqi = aq_c / aq_p if aq_p != 0 else None
            else:
                aqi = None

            # SGI: Sales Growth Index = Sales_t / Sales_{t-1}
            sgi = (None if None in [rev_c, rev_p] or rev_p == 0
                   else rev_c / rev_p)

            # DEPI: Depreciation Index = (DA_{t-1}/PPE_{t-1}) / (DA_t/PPE_t)
            if None not in [da_c, da_p, ppe_c, ppe_p] and ppe_c != 0 and da_c != 0:
                depi = (da_p / ppe_p) / (da_c / ppe_c) if ppe_p != 0 else None
            else:
                depi = None

            # SGAI: SGA Index = (SGA_t/Sales_t) / (SGA_{t-1}/Sales_{t-1})
            if None not in [opex_c, opex_p, rev_c, rev_p] and rev_c != 0 and rev_p != 0 and opex_p / rev_p != 0:
                sgai = (opex_c / rev_c) / (opex_p / rev_p)
            else:
                sgai = None

            # TATA: Total Accruals to Total Assets = (NI - OCF) / TA
            tata = (None if None in [ni_c, ocf_c, ta_c] or ta_c == 0
                    else (ni_c - ocf_c) / ta_c)

            # LVGI: Leverage Growth Index = (TD_t/TA_t) / (TD_{t-1}/TA_{t-1})
            if None not in [td_c, td_p, ta_c, ta_p] and ta_c != 0 and ta_p != 0 and td_p / ta_p != 0:
                lvgi = (td_c / ta_c) / (td_p / ta_p)
            else:
                lvgi = None

            # Need at least 6 of 8 variables to compute a meaningful M-Score
            vars_list = [dsri, gmi, aqi, sgi, depi, sgai, tata, lvgi]
            available = [v for v in vars_list if v is not None]
            if len(available) < 6:
                continue

            # Replace None with median of available values as conservative imputation
            # (rather than omitting entirely, which would overstate/understate the score)
            med = float(np.median(available))
            dsri_  = dsri  if dsri  is not None else med
            gmi_   = gmi   if gmi   is not None else med
            aqi_   = aqi   if aqi   is not None else med
            sgi_   = sgi   if sgi   is not None else med
            depi_  = depi  if depi  is not None else med
            sgai_  = sgai  if sgai  is not None else med
            tata_  = tata  if tata  is not None else 0.0
            lvgi_  = lvgi  if lvgi  is not None else med

            # Cap individual variables to prevent extreme outliers dominating the score
            def _bm_clip(v, lo=-5, hi=5): return float(np.clip(v, lo, hi))

            m = (-4.84
                 + 0.920  * _bm_clip(dsri_,  0, 5)
                 + 0.528  * _bm_clip(gmi_,   0, 5)
                 + 0.404  * _bm_clip(aqi_,   0, 5)
                 + 0.892  * _bm_clip(sgi_,   0, 5)
                 + 0.115  * _bm_clip(depi_,  0, 5)
                 - 0.172  * _bm_clip(sgai_,  0, 5)
                 + 4.679  * _bm_clip(tata_, -1, 1)
                 - 0.327  * _bm_clip(lvgi_,  0, 5))
            mscore[i] = float(np.clip(m, -10, 5))

        r["Beneish M-Score"] = mscore

        # ── Ohlson O-Score → Bankruptcy Probability ────────────────────────
        # Simplified model (without GNP deflator normalisation).
        # P(bankruptcy) = 1 / (1 + exp(-O-Score))
        # O < 0.5: low risk  |  O > 0.5: elevated risk
        oscore_prob = np.full(n, np.nan)
        for i in range(n):
            ta_v  = ta[i]; tl_v = total_liab[i]; wc_v = wc[i]
            tca_v = tca[i]; tcl_v = tcl[i]; ni_v = net_inc[i]; ocf_v = ocf[i]
            re_v  = retained_earnings[i]

            if any(np.isnan([ta_v, tl_v, ni_v])):
                continue
            if ta_v <= 0:
                continue

            oeneg  = 1.0 if tl_v > ta_v else 0.0  # negative equity signal
            wc_ta  = wc_v / ta_v  if not np.isnan(wc_v)  else 0.0
            cl_ca  = (tcl_v / tca_v if not np.isnan(tca_v) and not np.isnan(tcl_v)
                      and tca_v > 0 else 0.0)
            ni_ta  = ni_v / ta_v
            ffo_tl = (ocf_v / tl_v if not np.isnan(ocf_v) and tl_v > 0 else 0.0)
            intwo  = 0.0  # requires 2-year history; handled below

            if i >= 1 and not np.isnan(net_inc[i-1]):
                intwo = 1.0 if (ni_v < 0 and net_inc[i-1] < 0) else 0.0

            if i >= 1 and not np.isnan(net_inc[i-1]):
                denom = abs(ni_v) + abs(net_inc[i-1])
                chin  = (ni_v - net_inc[i-1]) / denom if denom > 1e-10 else 0.0
            else:
                chin = 0.0

            log_ta = np.log(max(ta_v, 1.0))  # simplified: no GNP deflation
            tl_ta  = tl_v / ta_v

            o = (-1.32 - 0.407 * log_ta + 6.03 * tl_ta - 1.43 * wc_ta
                 + 0.076 * cl_ca - 1.72 * oeneg - 2.37 * ni_ta
                 - 1.83 * ffo_tl + 0.285 * intwo - 0.521 * chin)

            oscore_prob[i] = 1.0 / (1.0 + np.exp(-o))  # probability of bankruptcy

        r["Ohlson Bankruptcy Prob"] = oscore_prob
        return r

    def _calc_advanced_metrics(self, data: dict) -> dict:
        """
        Additional institutional metrics not in the base set.

        Net Debt to Equity: (Total Debt - Cash) / Total Equity
          Different from D/E: explicitly captures cash position. Negative = net cash.
          More informative for capital allocation analysis.

        Earnings Yield: EBIT / EV (inverse of EV/EBIT)
          Greenblatt's "Magic Formula" numerator. Earnings-based alternative to E/P.

        ROIC Spread: ROIC - WACC proxy (using sector average cost of capital 10%)
          Positive spread = value creation; negative = value destruction.
          We use 10% as a rough universal cost-of-capital proxy when WACC is unknown.

        Interest Expense Ratio: Interest Expense / Total Revenue
          Shows what fraction of revenue is consumed by debt service before tax.

        Reinvestment Rate: (CapEx - D&A + ΔNWC) / NOPAT
          What fraction of NOPAT is reinvested. Pairs with ROIC for growth analysis.
          Sustainable Growth Rate = ROIC × Reinvestment Rate
        """
        r = {}
        td      = data['td']
        cash    = data['cash']
        equity  = data['equity']
        ev      = data['ev']
        op_inc  = data['op_inc']
        rev     = data['rev']
        int_exp = data['int_exp']
        nopat   = data['nopat']
        capex   = data['capex']
        da      = data['da']
        tca     = data['tca']
        tcl     = data['tcl']
        avg_ic  = data['avg_ic']
        roic    = self._safe_div(nopat, avg_ic)

        # Net Debt to Equity
        net_debt = self._safe_add(td, -cash)
        nde = self._safe_div(net_debt, equity)
        # Suppress when equity <= 0 (same logic as D/E)
        r["Net Debt to Equity"] = np.where(equity <= 0, np.nan, nde)

        # Earnings Yield = EBIT / EV (only when EV > 0)
        ey = self._safe_div(op_inc, ev)
        r["Earnings Yield (EBIT/EV)"] = np.where(ev <= 0, np.nan, ey)

        # ROIC Spread vs 10% hurdle (proxy for value creation without requiring WACC)
        roic_spread = roic - 0.10
        r["ROIC Spread (vs 10% Hurdle)"] = np.where(np.isnan(roic), np.nan, roic_spread)

        # Interest Expense / Revenue
        r["Interest Expense Ratio"] = self._safe_div(int_exp, rev)

        # Reinvestment Rate = (CapEx − D&A + ΔNWC) / NOPAT
        # Positive means NOPAT is being reinvested; negative means milking assets
        wc = self._safe_add(tca, -tcl)
        # ΔNWC needs prior period; compute pairwise
        delta_nwc = np.full(len(rev), np.nan)
        for i in range(1, len(rev)):
            if not np.isnan(wc[i]) and not np.isnan(wc[i-1]):
                delta_nwc[i] = wc[i] - wc[i-1]
        net_reinvest = self._safe_add(capex, -da, delta_nwc)
        rr = self._safe_div(net_reinvest, nopat)
        # Suppress when NOPAT <= 0 (reinvestment rate undefined under losses)
        r["Reinvestment Rate"] = np.where(nopat <= 0, np.nan, rr)

        # Sustainable Growth Rate = ROIC × Reinvestment Rate
        # Only meaningful when both are defined and NOPAT > 0
        sgr = np.where(
            ~np.isnan(roic) & ~np.isnan(rr) & (nopat > 0),
            roic * rr, np.nan
        )
        r["Sustainable Growth Rate"] = sgr

        return r

    def run_calculation(self):
        validation = self.validate_data()
        
        if not validation.is_valid:
            error_msg = '; '.join(validation.errors)
            raise ValueError(f"Validation failed: {error_msg}")
        
        if validation.warnings:
            for w in validation.warnings:
                logger.warning(w)
        
        self._cache.clear()
        self.calculation_audit = {}
        success_count = failed_count = 0
        failed_companies = []
        
        for company in self.companies:
            try:
                ratios = self._calculate_company_ratios(company)
                
                # Post-calculation filter: remove all-NaN ratio arrays
                company_audit = {}
                filtered_ratios = {}
                for ratio_name, ratio_values in ratios.items():
                    if np.all(np.isnan(ratio_values)):
                        company_audit[ratio_name] = "excluded: all values are missing or invalid"
                    else:
                        filtered_ratios[ratio_name] = ratio_values
                        company_audit[ratio_name] = "available"
                
                self.calculation_audit[company] = company_audit
                
                excluded_count = sum(1 for v in company_audit.values() if v.startswith("excluded"))
                if excluded_count > 0:
                    logger.info(f"{company}: {len(filtered_ratios)} ratios computed, {excluded_count} excluded (insufficient data)")
                
                # Verify we got some valid data
                if not filtered_ratios:
                    logger.warning(f"{company}: No ratios calculated - insufficient data")
                    self.results[company] = {k: np.full(self.num_years, np.nan) for k in ["Current Ratio"]}
                    failed_count += 1
                    failed_companies.append(company)
                else:
                    self.results[company] = filtered_ratios
                    success_count += 1
                    
            except Exception as e:
                logger.error(f"Failed for {company}: {str(e)}")
                self.results[company] = {k: np.full(self.num_years, np.nan) for k in ["Current Ratio"]}
                self.calculation_audit[company] = {"_error": str(e)}
                failed_count += 1
                failed_companies.append(company)
        
        if success_count == 0:
            raise ValueError("No companies successfully calculated. Check data mappings and quality.")
        
        total_ratios = len(next(iter(self.results.values())).keys()) if self.results else 0
        
        if failed_count > 0:
            logger.warning(f"Partial completion: {success_count}/{len(self.companies)} companies succeeded, {failed_count} had insufficient data")
            if len(failed_companies) <= 5:
                logger.info(f"Failed companies: {', '.join(failed_companies)}")
        else:
            logger.info(f"Complete: {success_count}/{len(self.companies)} companies, {total_ratios} ratios calculated")
    
    def _calculate_company_ratios(self, comp: str) -> Dict[str, np.ndarray]:
        rev = self._extract_series(comp, "Revenue")
        cogs = np.abs(self._extract_series(comp, "Cost of Revenue"))
        gp = self._extract_series(comp, "Gross Profit")
        if np.all(np.isnan(gp)):
            # Use explicit per-period mask: only compute gp where BOTH rev AND cogs
            # are available. _safe_add must NOT be used here — its any-value logic
            # treats a missing cogs as 0, producing gp = rev (100% gross margin).
            gp = np.where(~np.isnan(rev) & ~np.isnan(cogs), rev - cogs, np.nan)
        opex = np.abs(self._extract_series(comp, "Operating Expenses"))
        op_inc = self._extract_series(comp, "Operating Income")
        # Interest Expense: use gross (absolute) value for standard coverage ratios
        # Per CFA/IFRS convention, Interest Coverage = EBIT / Gross Interest Expense
        int_exp_raw = self._extract_series(comp, "Interest Expense")
        int_exp = np.abs(int_exp_raw)  # Gross absolute for coverage and debt calculations
        net_inc = self._extract_series(comp, "Net Income")
        pref_div = np.abs(self._extract_series(comp, "Preferred Dividends"))
        pretax = self._extract_series(comp, "Pretax Income")
        if np.all(np.isnan(pretax)):
            # Use explicit mask: only compute where BOTH op_inc AND int_exp are available.
            # _safe_add would treat missing interest expense as 0, overstating pretax
            # income for levered companies whose interest data is simply unmapped.
            pretax = np.where(~np.isnan(op_inc) & ~np.isnan(int_exp), op_inc - int_exp, np.nan)
        ebitda = self._extract_series(comp, "EBITDA")
        da = np.abs(self._extract_series(comp, "D&A"))
        if np.all(np.isnan(ebitda)):
            # Use explicit mask: only compute where BOTH op_inc AND da are available.
            # _safe_add would treat missing D&A as 0, understating EBITDA = EBIT.
            ebitda = np.where(~np.isnan(op_inc) & ~np.isnan(da), op_inc + da, np.nan)
        
        cash = self._extract_series(comp, "Cash & Equivalents")
        st_inv = self._extract_series(comp, "Short-Term Investments")
        receivables = self._extract_series(comp, "Accounts Receivable")
        inventory = self._extract_series(comp, "Inventory")
        prepaid = self._extract_series(comp, "Prepaid Expenses")
        tca = self._extract_series(comp, "Total Current Assets")
        ppe = self._extract_series(comp, "PP&E (Net)")
        ta = self._extract_series(comp, "Total Assets")
        
        payables = self._extract_series(comp, "Accounts Payable")
        st_debt = np.abs(self._extract_series(comp, "Short-Term Debt"))
        tcl = self._extract_series(comp, "Total Current Liabilities")
        lt_debt = np.abs(self._extract_series(comp, "Long-Term Debt"))
        td = np.abs(self._extract_series(comp, "Total Debt"))
        if np.all(np.isnan(td)):
            td = self._safe_add(st_debt, lt_debt)
        total_liab = np.abs(self._extract_series(comp, "Total Liabilities"))
        equity = self._extract_series(comp, "Total Equity")
        retained_earnings = self._extract_series(comp, "Retained Earnings")
        min_int = self._extract_series(comp, "Minority Interest")
        
        ocf = self._extract_series(comp, "Operating Cash Flow")
        capex = np.abs(self._extract_series(comp, "Capital Expenditures"))
        fcf = self._extract_series(comp, "Free Cash Flow")
        if np.all(np.isnan(fcf)):
            # Use explicit mask: only compute where BOTH ocf AND capex are available.
            # _safe_add would treat missing capex as 0, overstating FCF (fcf = ocf).
            fcf = np.where(~np.isnan(ocf) & ~np.isnan(capex), ocf - capex, np.nan)
        divs = np.abs(self._extract_series(comp, "Dividends Paid"))
        
        price = self._extract_series(comp, "Share Price")
        shares = self._extract_series(comp, "Shares Outstanding (Basic)")
        shares_diluted = self._extract_series(comp, "Shares Outstanding (Diluted)")
        if np.all(np.isnan(shares_diluted)):
            shares_diluted = shares.copy()
        mc = self._extract_series(comp, "Market Cap")
        if np.all(np.isnan(mc)):
            # Market Cap = Price × Basic Shares Outstanding (not diluted)
            mc = price * shares
        
        avg_ta = self._avg(ta)
        avg_equity = self._avg(equity)
        avg_inv = self._avg(inventory)
        avg_ar = self._avg(receivables)
        avg_ap = self._avg(payables)
        avg_ppe = self._avg(ppe)
        wc = self._safe_add(tca, -tcl)
        # FIX #4: NOPAT configurable - standard valuation vs conservative
        if self.config.nopat_tax_benefit_on_losses:
            # Standard valuation: NOPAT = OpInc × (1 - t) always
            nopat = op_inc * (1 - self.tax_rate)
        else:
            # Conservative: no tax benefit on losses
            nopat = np.where(op_inc > 0,
                             op_inc * (1 - self.tax_rate),
                             op_inc)
        # Invested Capital with documented operating cash assumption
        # Operating cash % is an estimate. Default 2% of revenue.
        # This materially impacts ROIC. Review for your industry.
        if self.config.use_excess_cash_adjustment:
            operating_cash_estimate = self.config.operating_cash_pct * rev
            excess_cash = np.maximum(cash - operating_cash_estimate, 0)
        else:
            excess_cash = np.zeros_like(cash)
        ic = self._safe_add(equity, td, min_int, -excess_cash)
        ic = np.where(ic <= 0, np.nan, ic)
        avg_ic = self._avg(ic)
        avg_ic = np.where(avg_ic <= 0, np.nan, avg_ic)
        net_debt = self._safe_add(td, -cash)
        preferred_stock = np.abs(self._extract_series(comp, "Preferred Stock"))
        ev = self._safe_add(mc, td, min_int, preferred_stock, -cash)
        # FIX #2: Compute annualized EPS for valuation/growth (trailing for non-annual)
        if self.frequency != Frequency.ANNUAL:
            trailing_net_inc = self._trailing_sum(net_inc)
            trailing_pref_div = self._trailing_sum(pref_div)
            eps_annualized = self._safe_div(self._safe_add(trailing_net_inc, -trailing_pref_div), shares_diluted)
        else:
            eps_annualized = self._safe_div(self._safe_add(net_inc, -pref_div), shares_diluted)
        # Raw period EPS (for display, not growth)
        eps = self._safe_div(self._safe_add(net_inc, -pref_div), shares_diluted)
        
        # Accounting identity consistency checks (validation layer)
        self._check_accounting_identities(comp, rev, cogs, gp, op_inc, da, ebitda, ta, total_liab, equity)
        
        data = {
            'rev': rev, 'cogs': cogs, 'gp': gp, 'opex': opex, 'op_inc': op_inc,
            'int_exp': int_exp,  # Gross interest expense for standard coverage ratios
            'net_inc': net_inc, 'pretax': pretax, 'ebitda': ebitda, 'da': da, 'pref_div': pref_div,
            'cash': cash, 'st_inv': st_inv, 'receivables': receivables, 'inventory': inventory, 'prepaid': prepaid, 'tca': tca, 
            'ppe': ppe, 'ta': ta, 'payables': payables, 'tcl': tcl, 'td': td, 'total_liab': total_liab, 
            'equity': equity, 'retained_earnings': retained_earnings, 'min_int': min_int, 'ocf': ocf, 
            'capex': capex, 'fcf': fcf, 'divs': divs, 'price': price, 'shares': shares, 
            'shares_diluted': shares_diluted, 'mc': mc, 'avg_ta': avg_ta, 'avg_equity': avg_equity, 
            'avg_inv': avg_inv, 'avg_ar': avg_ar, 'avg_ap': avg_ap, 'avg_ppe': avg_ppe, 'wc': wc, 
            'nopat': nopat, 'ic': ic, 'avg_ic': avg_ic, 'net_debt': net_debt, 'ev': ev,
            'eps': eps, 'eps_annualized': eps_annualized
        }
        
        ratios = {}
        ratios.update(self._calc_liquidity(data))
        ratios.update(self._calc_activity(data))
        ratios.update(self._calc_solvency(data))
        ratios.update(self._calc_profitability(data))
        ratios.update(self._calc_returns(data))
        ratios.update(self._calc_cashflow(data))
        ratios.update(self._calc_valuation(data))
        ratios.update(self._calc_growth(data))
        ratios["Altman Z-Score"] = self._calc_zscore_public(data)
        ratios["Altman Z-Score (EM Score)"] = self._calc_zscore_private(data)
        ratios.update(self._calc_quality_scores(data))
        ratios.update(self._calc_advanced_metrics(data))
        return ratios
    
    def _check_accounting_identities(self, company: str, rev, cogs, gp, op_inc, da, ebitda, ta, total_liab, equity):
        """Validate accounting identities and log warnings for inconsistencies."""
        tolerance = 0.01  # 1% tolerance for rounding
        
        # Check: Gross Profit ≈ Revenue - COGS
        if not np.all(np.isnan(gp)) and not np.all(np.isnan(rev)) and not np.all(np.isnan(cogs)):
            expected_gp = rev - cogs
            mask = ~np.isnan(gp) & ~np.isnan(expected_gp) & (np.abs(expected_gp) > 1e-6)
            if np.any(mask):
                deviation = np.abs((gp[mask] - expected_gp[mask]) / expected_gp[mask])
                if np.any(deviation > tolerance):
                    logger.warning(f"{company}: Gross Profit does not equal Revenue minus Cost of Revenue (deviation > {tolerance:.0%}). Verify source data.")
        
        # Check: EBITDA ≈ Operating Income + D&A
        if not np.all(np.isnan(ebitda)) and not np.all(np.isnan(op_inc)) and not np.all(np.isnan(da)):
            expected_ebitda = op_inc + da
            mask = ~np.isnan(ebitda) & ~np.isnan(expected_ebitda) & (np.abs(expected_ebitda) > 1e-6)
            if np.any(mask):
                deviation = np.abs((ebitda[mask] - expected_ebitda[mask]) / expected_ebitda[mask])
                if np.any(deviation > tolerance):
                    logger.warning(f"{company}: EBITDA does not equal Operating Income plus Depreciation and Amortisation (deviation > {tolerance:.0%}). Verify source data.")
        
        # Check: Total Assets ≈ Total Liabilities + Total Equity
        if not np.all(np.isnan(ta)) and not np.all(np.isnan(total_liab)) and not np.all(np.isnan(equity)):
            expected_ta = total_liab + equity
            mask = ~np.isnan(ta) & ~np.isnan(expected_ta) & (np.abs(expected_ta) > 1e-6)
            if np.any(mask):
                deviation = np.abs((ta[mask] - expected_ta[mask]) / expected_ta[mask])
                if np.any(deviation > tolerance):
                    logger.warning(f"{company}: Total Assets does not equal Total Liabilities plus Total Equity (deviation > {tolerance:.0%}). Verify source data.")
    
    def generate_peer_matrix(self) -> Dict[str, pd.DataFrame]:
        if not self.results or not self.companies:
            return {}
        first_comp = self.companies[0]
        all_ratios = list(self.results[first_comp].keys())
        ratio_matrices = {}
        for ratio_name in all_ratios:
            df = pd.DataFrame(index=self.companies, columns=self.years, dtype=float)
            for comp in self.companies:
                if comp in self.results:
                    df.loc[comp] = self.results[comp].get(ratio_name, np.full(self.num_years, np.nan))
            try:
                numeric_df = df.apply(pd.to_numeric, errors='coerce')
                df.loc["Mean"] = numeric_df.mean(axis=0, skipna=True)
                df.loc["Median"] = numeric_df.median(axis=0, skipna=True)
                df.loc["Std Dev"] = numeric_df.std(axis=0, skipna=True)
            except (ValueError, TypeError, KeyError) as e:
                logger.warning(f"Failed to calculate statistics for {ratio_name}: {e}")
            ratio_matrices[ratio_name] = df
        return ratio_matrices
    
    def export_excel(self, buffer):
        if not self.results:
            raise ValueError("No results to export")
        
        matrices = self.generate_peer_matrix()
        
        # 1. Define Categories for Sheets
        # Any ratio not matched here will go to "Other Ratios"
        # FIX #16: DuPont components belong with Returns, not Profitability
        CATEGORIES = {
            "Liquidity": ["Current Ratio", "Quick Ratio", "Cash Ratio", "Defensive", "NWC"],
            "Efficiency": ["Turnover", "Days", "Cycle", "DSO", "DIO", "DPO"],
            "Solvency": ["Debt", "Coverage", "Leverage"],
            "Profitability": ["Gross Margin", "Operating Margin", "Net Margin", "EBITDA Margin", "Pretax Margin"],
            "Returns": ["ROA", "ROE", "ROIC", "Return on", "DuPont", "Equity Multiplier"],
            "Valuation": ["P/E", "EV /", "Price to", "Yield", "PEG", "EPS", "Book Value"],
            "Cash Flow": ["OCF", "FCF", "Capex", "Dividend", "Quality"],
            "Growth": ["Growth"],
            "Risk": ["Z-Score"]
        }

        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Formats
            header_fmt = workbook.add_format({
                'font_name': 'Calibri', 'font_size': 11, 'bold': True, 
                'bg_color': '#4472C4', 'font_color': 'white', 
                'border': 1, 'align': 'center', 'valign': 'vcenter'
            })
            ratio_fmt = workbook.add_format({
                'font_name': 'Calibri', 'font_size': 11, 'bold': True, 
                'bg_color': '#D9E1F2', 'border': 1
            })
            num_fmt = workbook.add_format({
                'font_name': 'Calibri', 'font_size': 11, 
                'num_format': '#,##0.00', 'border': 1
            })
            
            # Helper to determine sheet name
            def get_category(ratio_name):
                for cat, keywords in CATEGORIES.items():
                    if any(k in ratio_name for k in keywords):
                        return cat
                return "Other Ratios"

            # Group matrices by category
            grouped_data = {}
            for ratio_name, df in matrices.items():
                cat = get_category(ratio_name)
                if cat not in grouped_data:
                    grouped_data[cat] = []
                grouped_data[cat].append((ratio_name, df))
            
            # Sort categories order (Key ones first)
            priority_order = ["Liquidity", "Profitability", "Returns", "Solvency", 
                              "Efficiency", "Valuation", "Growth", "Cash Flow", "Risk", "Other Ratios"]
            
            sorted_cats = [c for c in priority_order if c in grouped_data]
            
            # Create a sheet for each category
            for category in sorted_cats:
                ratios_in_cat = grouped_data[category]
                sheet_name = category[:31]  #Truncate to Excel's 31-char limit
                worksheet = workbook.add_worksheet(sheet_name)
                
                # Column widths
                worksheet.set_column(0, 0, 30) # Ratio/Company Name
                worksheet.set_column(1, len(self.years) + 1, 12) # Years
                
                current_row = 0
                
                for ratio_name, df in ratios_in_cat:
                    # Write Ratio Title
                    worksheet.merge_range(current_row, 0, current_row, len(self.years), ratio_name, ratio_fmt)
                    current_row += 1
                    
                    # Write Header (Years)
                    worksheet.write(current_row, 0, "Company / Metric", header_fmt)
                    for col_idx, year in enumerate(self.years):
                        worksheet.write(current_row, col_idx + 1, str(year), header_fmt)
                    current_row += 1
                    
                    # Write Data
                    # 1. Companies
                    for company in self.companies:
                        if company in df.index:
                            worksheet.write(current_row, 0, company, num_fmt) # Reuse num_fmt for border or create generic border
                            for col_idx, year in enumerate(self.years):
                                val = df.loc[company, year]
                                worksheet.write(current_row, col_idx + 1, val if pd.notna(val) else "", num_fmt)
                            current_row += 1
                            
                    # 2. Stats (Mean/Median)
                    for stat in ['Mean', 'Median']:
                        if stat in df.index:
                            worksheet.write(current_row, 0, f"{stat} (Peer Group)", header_fmt) # Highlight stats
                            for col_idx, year in enumerate(self.years):
                                val = df.loc[stat, year]
                                worksheet.write(current_row, col_idx + 1, val if pd.notna(val) else "", num_fmt)
                            current_row += 1
                    
                    current_row += 2  # Gap between ratios

        buffer.seek(0)
    
    def get_data_quality_report(self) -> Dict:
        active_mappings = {k: v for k, v in self.gui_mapping.items() if v and str(v).strip() and str(v) != "(Not Mapped)"}
        report = {'total_companies': len(self.companies), 'total_years': self.num_years, 'fields_mapped': len(active_mappings), 'frequency': self.frequency.value, 'company_completeness': {}, 'field_availability': {}}
        if not active_mappings:
            return report
        for company in self.companies:
            total_points = available_points = 0
            for field_key in active_mappings.keys():
                data = self._extract_series(company, field_key)
                total_points += len(data)
                available_points += np.sum(~np.isnan(data))
            score = (available_points / total_points * 100) if total_points > 0 else 0
            report['company_completeness'][company] = {'score': round(score, 1), 'available_points': available_points, 'total_points': total_points}
        for field_key in active_mappings.keys():
            companies_with_data = sum(1 for company in self.companies if not np.all(np.isnan(self._extract_series(company, field_key))))
            avail = (companies_with_data / len(self.companies) * 100) if self.companies else 0
            report['field_availability'][field_key] = round(avail, 1)
        return report


# Backward-compatible alias
AdvancedRatioEngine = RatioEngine