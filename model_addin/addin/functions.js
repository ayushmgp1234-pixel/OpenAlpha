/* global CustomFunctions */

// ⚠️  Replace this with your actual Render URL after deploying the backend
// e.g. "https://model-addin-api.onrender.com"
const API = "https://YOUR-APP-NAME.onrender.com";

async function callAPI(path, params) {
  // Remove empty params
  Object.keys(params).forEach(k => {
    if (params[k] === "" || params[k] === undefined || params[k] === null) {
      delete params[k];
    }
  });
  const url = `${API}/${path}?${new URLSearchParams(params)}`;
  try {
    const res = await fetch(url);
    if (!res.ok) {
      const json = await res.json().catch(() => ({}));
      return `#ERROR: ${json.error || res.statusText}`;
    }
    return await res.json();
  } catch (e) {
    return `#ERROR: Cannot reach API – check ${API}/health`;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// =MODEL("AAPL", "Revenue", "LY")
// =MODEL("AAPL", "Revenue", 2022)
// =MODEL("AAPL", "Revenue", "LY", "Q2")
// =MODEL("AAPL", "Revenue", "LY",, 1000000)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Get financials, key metrics, dividends or estimates for any stock.
 * @customfunction
 * @param {string} ticker    Stock symbol e.g. "AAPL"
 * @param {string} param     Field: Revenue, EPS, ROE, PE Ratio, Free Cash Flow, Dividend, Sector...
 * @param {string} [period]  TTM | LY | LY-1 | LY-2 | LQ | LQ-1 | 2024 | 2023 ...
 * @param {string} [quarter] Q1, Q2, Q3, or Q4 for quarterly data
 * @param {number} [divisor] e.g. 1000000 for millions
 * @returns {string|number}
 */
async function MODEL(ticker, param, period, quarter, divisor) {
  const res = await callAPI("model", {
    ticker,
    param,
    period:  period  || "LY",
    quarter: quarter || "",
    divisor: divisor || 1,
  });
  if (typeof res === "string") return res;
  return res.value === null || res.value === undefined ? "#N/A" : res.value;
}
CustomFunctions.associate("MODEL", MODEL);

// ─────────────────────────────────────────────────────────────────────────────
// =MODELPRICE("AAPL", "Price")
// =MODELPRICE("AAPL", "Close", 30)
// =MODELPRICE("AAPL", "Close",, "01/01/2024", "06/01/2024")
// =MODELPRICE("BTCUSD", "Price")   → crypto
// =MODELPRICE("EURUSD", "Price")   → forex
// =MODELPRICE("^GSPC", "Close", 10) → S&P 500
// =MODELPRICE("GC=F", "Price")     → Gold
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Get live or historical prices for stocks, crypto, forex, commodities and indices.
 * @customfunction
 * @param {string} ticker    Symbol: AAPL, BTCUSD, EURUSD, ^GSPC, GC=F ...
 * @param {string} param     Price | Open | High | Low | Close | Volume | Dividend | 52 Week High ...
 * @param {number} [days]    Last N trading days (returns a table)
 * @param {string} [start]   Start date mm/dd/yyyy
 * @param {string} [end]     End date mm/dd/yyyy
 * @returns {string|number|any[][]}
 */
async function MODELPRICE(ticker, param, days, start, end) {
  const res = await callAPI("modelprice", {
    ticker, param,
    days:  days  || "",
    start: start || "",
    end:   end   || "",
  });
  if (typeof res === "string") return res;

  // Single live value
  if (res.value !== undefined) {
    return res.value === null ? "#N/A" : res.value;
  }

  // Table / matrix
  if (res.data && Array.isArray(res.data)) {
    if (res.data.length === 0) return "#N/A";
    if ("dividend" in res.data[0]) {
      return [["Date", "Dividend"]].concat(
        res.data.map(r => [r.date, r.dividend])
      );
    }
    return [["Date", param]].concat(res.data.map(r => [r.date, r.value]));
  }
  return "#N/A";
}
CustomFunctions.associate("MODELPRICE", MODELPRICE);

// ─────────────────────────────────────────────────────────────────────────────
// =MODELFUNDS("SPY", "Expense Ratio")
// =MODELFUNDS("QQQ", "AUM")
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Get ETF/Fund data: NAV, Expense Ratio, AUM, Yield, YTD Return, Category ...
 * @customfunction
 * @param {string} ticker  ETF symbol e.g. "SPY", "QQQ"
 * @param {string} param   NAV | Expense Ratio | AUM | Yield | YTD Return | Category | Fund Family ...
 * @returns {string|number}
 */
async function MODELFUNDS(ticker, param) {
  const res = await callAPI("modelfunds", { ticker, param });
  if (typeof res === "string") return res;
  return res.value === null ? "#N/A" : res.value;
}
CustomFunctions.associate("MODELFUNDS", MODELFUNDS);

// ─────────────────────────────────────────────────────────────────────────────
// =MODELOPTIONS("AAPL", "Call", "Strike")
// =MODELOPTIONS("AAPL", "Put", "IV", "2025-06-20")
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Get real-time or historical options chain data.
 * @customfunction
 * @param {string} ticker      Stock symbol e.g. "AAPL"
 * @param {string} optionType  "Call" or "Put"
 * @param {string} param       Strike | Bid | Ask | IV | Volume | Open Interest | In The Money ...
 * @param {string} [date]      Expiration date YYYY-MM-DD (omit for nearest expiry)
 * @returns {any[][]}
 */
async function MODELOPTIONS(ticker, optionType, param, date) {
  const res = await callAPI("modeloptions", {
    ticker, type: optionType, param, date: date || "",
  });
  if (typeof res === "string") return res;
  if (res.data && Array.isArray(res.data)) {
    return [[`${optionType} ${param} — exp: ${res.expiration}`],
            ...res.data.map(v => [v])];
  }
  return "#N/A";
}
CustomFunctions.associate("MODELOPTIONS", MODELOPTIONS);
