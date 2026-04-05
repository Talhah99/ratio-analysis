# 📊 Financial Ratio Analysis Engine

A professional-grade, multi-company financial ratio analysis tool built with Python and Streamlit. Upload your own Excel data, map financial fields, and instantly compute **60+ institutional-quality ratios** across liquidity, profitability, solvency, returns, valuation, growth, and bankruptcy risk.

🚀 **[Live Demo →](https://your-app-name.streamlit.app)**  
*(Replace with your Streamlit Community Cloud URL after deployment)*

---

## Features

- **60+ Financial Ratios** — Liquidity, Efficiency, Solvency, Profitability, Returns, Cash Flow, Valuation, Growth, and Risk
- **Multi-Company Peer Comparison** — side-by-side benchmarking with peer mean/median
- **Advanced Quantitative Models** — Altman Z-Score (Public + EM), Piotroski F-Score, Beneish M-Score, Ohlson Bankruptcy Probability, Sloan Accrual Ratio
- **Frequency Support** — Annual, Quarterly, and Monthly data with correct trailing-sum annualisation
- **Interactive Charts** — Line, Bar, and Area charts with Plotly
- **HTML Dashboard Export** — Offline standalone dashboard per company (9 charts + automated insights)
- **Excel Export** — All ratios organized into category sheets with peer statistics
- **PDF Formula Guide** — Downloadable reference of every formula, edge case, and scholarly citation
- **Dark & Light Mode** — Fully adaptive theme using CSS custom properties

---

## Data Format

Your Excel file must follow this layout (one row per company-metric):

| Company     | Field         | 2021  | 2022  | 2023  |
|-------------|---------------|-------|-------|-------|
| ABC Ltd     | Revenue       | 1000  | 1200  | 1500  |
| ABC Ltd     | Net Income    | 100   | 120   | 150   |
| ABC Ltd     | Total Assets  | 2000  | 2200  | 2500  |
| XYZ Corp    | Revenue       | 5000  | 5500  | 6000  |

**Rules:**
- Column 1: Company name (repeat for every metric row)
- Column 2: Financial field name
- Column 3+: Period data (numeric only)
- No merged cells, no trailing empty columns

Download sample datasets directly inside the app (Cement & Pharma sectors provided).

---

## Running Locally

```bash
# 1. Clone the repository
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name

# 2. Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate       # Windows
# source .venv/bin/activate  # macOS/Linux

# 3. Install dependencies
pip install -r requirements.txt

# 4. Launch the app
streamlit run App.py
```

---

## Tech Stack

| Library | Purpose |
|---|---|
| `streamlit` | Web framework & UI |
| `pandas` | Data manipulation |
| `numpy` | Vectorised numerical computation |
| `plotly` | Interactive charts |
| `openpyxl` / `xlrd` | Excel reading |
| `xlsxwriter` | Excel export |
| `reportlab` | PDF formula guide generation |
| `scipy` | Statistical utilities |

---

## Ratio Categories

| Category | Key Ratios |
|---|---|
| 💧 Liquidity | Current Ratio, Quick Ratio, Cash Ratio, Defensive Interval, NWC to Assets |
| ⚙️ Efficiency | Inventory/Receivables/Payables Turnover, CCC, DOL, Reinvestment Rate |
| 🏛️ Solvency | Debt/Equity, Interest Coverage, Net Debt/EBITDA, Financial Leverage |
| 📈 Profitability | Gross/Operating/Net/EBITDA/Pretax Margins |
| 💰 Returns | ROA, ROE, ROIC, Cash ROIC, DuPont Decomposition |
| 💵 Cash Flow | OCF to Sales, FCF Conversion, Quality of Income, Capex Coverage |
| 📊 Valuation | P/E, EV/EBITDA, P/B, EPS, PEG, FCF Yield |
| 📉 Growth | Revenue, EBITDA, Net Income, EPS, FCF Growth |
| ⚠️ Risk | Altman Z-Score, Piotroski F-Score, Beneish M-Score, Ohlson O-Score |

---

## Disclaimer

This tool is for **educational and analytical purposes only**. It is not financial advice. All calculations are estimates based on the data provided. Always verify independently before making investment decisions.

---

## License

MIT License — see [LICENSE](LICENSE) for details.
