# 🧮 Portfolio Auto Updater

A lightweight, self-contained system to keep your investment portfolio in Excel automatically updated with live market data.  
Built for personal finance tracking — no external database or paid APIs required.

---

## 🚀 Features
- ✅ Pulls **real-time stock / ETF / FX prices** from Yahoo Finance via `yfinance`
- ✅ Supports **intraday intervals** (`1m`, `5m`, `60m`, `1d`) with automatic fallback
- ✅ Converts all prices to **EUR**
- ✅ Integrates directly with your **Excel portfolio**
- ✅ Optional `--eur-only` flag to skip FX conversion (if portfolio is constrcuted by EUR only assets)
- ✅ Can be scheduled via **.bat** or Windows Task Scheduler
- ✅ Minimal dependencies (pandas, openpyxl, yfinance, requests)

---

## 📁 Folder Structure
update_portfolio_unified.py # Main Python script
Portfolio_template.xlsx # Excel file with 'Map', 'Transactions', and 'Analytics' sheets
update_portfolio.bat # One-click updater (runs Python, then opens Excel)
requirements.txt # pip dependencies


---

## ⚙️ Installation

```bash
git clone https://github.com/<yourname>/investments-auto-updater.git
cd investments-auto-updater

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

Optionally, set your OpenFIGI API key (for ISIN→Ticker mapping):

setx OPENFIGI_API_KEY "your_key_here"
