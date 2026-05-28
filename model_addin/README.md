# MODEL() Excel Online Add-in

Live stock data in Excel Online via `=MODEL()`, `=MODELPRICE()`, `=MODELFUNDS()`, `=MODELOPTIONS()`

---

## Deployment Overview

```
Excel Online  ←──  addin/ files (GitHub Pages)
                       │
                       └──→  backend/ API (Render)
                                    │
                                    └──→  yfinance (Yahoo Finance)
```

---

## Step 1 — Deploy Backend to Render

1. Create a free account at **render.com**
2. Click **New → Web Service**
3. Connect your GitHub account and select your repo
   - If you don't have a repo yet: create one, push this whole folder to it
4. Set these values:
   - **Root Directory:** `backend`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn server:app --workers 2 --timeout 60`
   - **Instance Type:** Free
5. Click **Deploy**
6. Wait ~2 minutes. Your API URL will look like:
   `https://model-addin-api.onrender.com`
7. Visit `https://your-app.onrender.com/health` — should return `{"status":"ok"}`

---

## Step 2 — Update the Render URL in the add-in files

Open `addin/functions.js` and replace line 4:
```js
// BEFORE
const API = "https://YOUR-APP-NAME.onrender.com";

// AFTER (example)
const API = "https://model-addin-api.onrender.com";
```

Also update `addin/functions.json` — replace `YOUR-APP-NAME` in the helpUrl fields.

---

## Step 3 — Host Add-in Files on GitHub Pages

1. Push your repo to GitHub (the whole project)
2. Go to your repo → **Settings → Pages**
3. Set **Source** to `main` branch, folder `/addin`
4. Click **Save**
5. GitHub Pages URL will be:
   `https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/`

---

## Step 4 — Update the Manifest

Open `addin/manifest.xml` and replace ALL occurrences of:
- `YOUR-GITHUB-USERNAME` → your GitHub username
- `YOUR-REPO-NAME` → your repo name
- `YOUR-APP-NAME` → your Render app name

Example — if your GitHub username is `jsmith`, repo is `model-addin`, Render app is `model-addin-api`:
```xml
<!-- Replace this -->
https://YOUR-GITHUB-USERNAME.github.io/YOUR-REPO-NAME/functions.js

<!-- With this -->
https://jsmith.github.io/model-addin/functions.js
```

---

## Step 5 — Upload Manifest to Excel Online

1. Open **Excel Online** (office.com → Excel)
2. Open any workbook
3. Go to **Insert → Office Add-ins**
4. Click **Upload My Add-in** tab
5. Browse to `addin/manifest.xml` → **Upload**
6. Done! The MODEL functions are now available.

---

## Usage

```excel
=MODEL("AAPL", "Revenue", "TTM")
=MODEL("AAPL", "Revenue", "LY",, 1000000)        → in millions
=MODEL("TSLA", "EPS", "LQ")                      → latest quarter EPS
=MODEL("MSFT", "Free Cash Flow", "LY-1")
=MODEL("AAPL", "PE Ratio", "TTM")
=MODEL("AAPL", "Dividend", "LQ")
=MODEL("AAPL", "Sector", "LY")
=MODEL("AAPL", "Target Price", "LY")
=MODEL("AAPL", "Recommendation", "LY")

=MODELPRICE("AAPL", "Price")
=MODELPRICE("AAPL", "Close", 30)                  → 30-day table
=MODELPRICE("AAPL", "Close",, "01/01/2024", "06/01/2024")
=MODELPRICE("BTCUSD", "Price")                    → Bitcoin
=MODELPRICE("ETHUSD", "Close", 10)                → Ethereum 10 days
=MODELPRICE("EURUSD", "Price")                    → EUR/USD live
=MODELPRICE("^GSPC", "Close", 10)                 → S&P 500
=MODELPRICE("GC=F", "Price")                      → Gold futures
=MODELPRICE("AAPL", "Dividend")                   → dividend history

=MODELFUNDS("SPY", "Expense Ratio")
=MODELFUNDS("QQQ", "AUM")
=MODELFUNDS("VTI", "YTD Return")
=MODELFUNDS("SPY", "NAV")

=MODELOPTIONS("AAPL", "Call", "Strike")
=MODELOPTIONS("AAPL", "Put", "IV")
=MODELOPTIONS("AAPL", "Call", "Strike", "2025-06-20")
```

---

## All Supported Fields

### MODEL — periods: TTM | LY | LY-1 | LY-2 | LQ | LQ-1 | 2024 | 2023 ...

**Income Statement:** Revenue · Cost of Revenue · Gross Profit · R&D · SGA · Operating Income · EBIT · Interest Expense · Net Income · EBITDA · EPS · Diluted EPS · Shares Outstanding

**Balance Sheet:** Cash · Total Cash · Current Assets · Total Assets · Current Liabilities · Total Debt · Long Term Debt · Total Liabilities · Shareholders Equity · Book Value · Retained Earnings

**Cash Flow:** Operating Cash Flow · Capex · Free Cash Flow · Investing Cash Flow · Financing Cash Flow · Dividends Paid · Share Buybacks

**Key Metrics:** PE Ratio · Forward PE · PEG Ratio · Price to Book · Price to Sales · EV/EBITDA · EV/Revenue · Enterprise Value · Market Cap · Profit Margin · Operating Margin · Gross Margin · ROE · ROA · Current Ratio · Quick Ratio · Debt to Equity · Beta

**Dividends:** Dividend · Adjusted Dividend *(use LY / LQ / 2023 etc.)*

**Estimates:** Estimated Revenue Avg · Estimated EPS Avg · Target Price · Target High · Target Low · Recommendation · Analyst Count

**Profile:** Name · Sector · Industry · Country · Employees · Website · Exchange · Description

### MODELPRICE — asset classes

| Asset | Ticker format | Example |
|---|---|---|
| Stocks/ETFs | Normal | `AAPL`, `SPY` |
| Crypto | No hyphen, no dash | `BTCUSD`, `ETHUSD` |
| Forex | No =X suffix | `EURUSD`, `GBPJPY` |
| Commodities | Yahoo format | `GC=F` (Gold), `CL=F` (Oil) |
| Indices | Yahoo format | `^GSPC`, `^IXIC`, `^DJI` |

**Live params:** Price · Open · High · Low · Prev Close · Volume · 52 Week High · 52 Week Low · 50 Day MA · 200 Day MA · Market Cap

**Historical params:** Open · High · Low · Close · Volume *(add days or start/end)*

### MODELFUNDS
NAV · Expense Ratio · AUM · Yield · YTD Return · 3 Year Return · 5 Year Return · Category · Fund Family · Morningstar Rating

### MODELOPTIONS
Strike · Last Price · Bid · Ask · Volume · Open Interest · IV · Implied Volatility · In The Money · Contract · Change · Percent Change

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `#ERROR: Cannot reach API` | Check your Render URL in `functions.js`; visit `/health` |
| Render is slow first request | Free tier sleeps after 15 min of inactivity — first call wakes it up (30s) |
| `#ERROR: Unknown parameter` | Visit `https://your-app.onrender.com/fields` for full list |
| Manifest rejected by Excel Online | Make sure ALL `YOUR-*` placeholders are replaced |
| Functions show `#GETTING_DATA` | Excel is fetching — wait a few seconds |
| GitHub Pages shows 404 | Check Pages is set to `/addin` folder, not root |

---

## Keeping Render Awake (Optional)

Render's free tier sleeps after 15 minutes. To avoid slow first-call wakeup, use a free uptime monitor like **UptimeRobot** to ping `https://your-app.onrender.com/health` every 14 minutes.
