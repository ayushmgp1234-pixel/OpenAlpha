"""
MODEL() Excel Add-in Backend — Render Deployment
Full WiseSheet feature parity via yfinance.
"""

from flask import Flask, jsonify, request
from flask_cors import CORS
import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

app = Flask(__name__)

# Allow requests from GitHub Pages and Excel Online
CORS(app, origins=[
    "https://*.github.io",
    "https://*.officeapps.live.com",
    "https://*.office.com",
    "https://*.microsoft.com",
    "*"   # loosen during dev; tighten to your GitHub Pages URL after deploy
])

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def err(msg, code=400):
    return jsonify({"error": str(msg)}), code

def safe(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if hasattr(val, "timestamp"):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        return round(val, 6)
    return str(val)

def ticker_obj(sym):
    return yf.Ticker(sym.upper().strip())

def parse_date(s):
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    raise ValueError(f"Cannot parse date: {s}")

# ─────────────────────────────────────────────────────────────────────────────
# FIELD MAPS
# ─────────────────────────────────────────────────────────────────────────────

INCOME_FIELDS = {
    "revenue": "Total Revenue", "total revenue": "Total Revenue",
    "cost of revenue": "Cost Of Revenue", "gross profit": "Gross Profit",
    "research and development": "Research And Development", "r&d": "Research And Development",
    "selling general administrative": "Selling General And Administrative", "sga": "Selling General And Administrative",
    "operating income": "Operating Income", "ebit": "Operating Income",
    "interest expense": "Interest Expense", "pretax income": "Pretax Income",
    "income tax": "Tax Provision", "net income": "Net Income",
    "ebitda": "EBITDA", "eps": "Basic EPS", "basic eps": "Basic EPS",
    "diluted eps": "Diluted EPS", "shares outstanding": "Diluted Average Shares",
}

BALANCE_FIELDS = {
    "cash": "Cash And Cash Equivalents", "total cash": "Cash And Cash Equivalents",
    "short term investments": "Other Short Term Investments",
    "total current assets": "Current Assets", "current assets": "Current Assets",
    "total assets": "Total Assets",
    "total current liabilities": "Current Liabilities", "current liabilities": "Current Liabilities",
    "total debt": "Total Debt", "long term debt": "Long Term Debt",
    "total liabilities": "Total Liabilities Net Minority Interest",
    "shareholders equity": "Stockholders Equity", "book value": "Stockholders Equity",
    "retained earnings": "Retained Earnings",
}

CASHFLOW_FIELDS = {
    "operating cash flow": "Operating Cash Flow",
    "capital expenditures": "Capital Expenditure", "capex": "Capital Expenditure",
    "free cash flow": "Free Cash Flow",
    "investing cash flow": "Investing Cash Flow", "financing cash flow": "Financing Cash Flow",
    "dividends paid": "Common Stock Dividend Paid", "share buybacks": "Repurchase Of Capital Stock",
}

KEY_METRICS = {
    "pe ratio": "trailingPE", "p/e ratio": "trailingPE", "trailing pe": "trailingPE",
    "forward pe": "forwardPE", "peg ratio": "pegRatio",
    "price to book": "priceToBook", "p/b ratio": "priceToBook",
    "price to sales": "priceToSalesTrailing12Months",
    "ev/ebitda": "enterpriseToEbitda", "ev/revenue": "enterpriseToRevenue",
    "enterprise value": "enterpriseValue", "market cap": "marketCap",
    "profit margin": "profitMargins", "operating margin": "operatingMargins",
    "gross margin": "grossMargins", "ebitda margin": "ebitdaMargins",
    "roe": "returnOnEquity", "return on equity": "returnOnEquity",
    "roa": "returnOnAssets", "return on assets": "returnOnAssets",
    "current ratio": "currentRatio", "quick ratio": "quickRatio",
    "debt to equity": "debtToEquity",
    "dividend rate": "dividendRate", "dividend yield": "dividendYield",
    "payout ratio": "payoutRatio", "ex dividend date": "exDividendDate",
    "target price": "targetMeanPrice", "last month avg price target": "targetMeanPrice",
    "last quarter avg price target": "targetMeanPrice",
    "target high": "targetHighPrice", "target low": "targetLowPrice",
    "recommendation": "recommendationKey", "analyst count": "numberOfAnalystOpinions",
    "earnings growth": "earningsGrowth", "revenue growth": "revenueGrowth",
    "beta": "beta", "shares short": "sharesShort",
    "short ratio": "shortRatio", "short percent": "shortPercentOfFloat",
    "estimated revenue avg": "revenueEstimate", "estimated eps avg": "earningsEstimate",
}

LIVE_PRICE_FIELDS = {
    "price": "last_price", "open": "open", "high": "day_high", "low": "day_low",
    "volume": "last_volume", "prev close": "previous_close", "previous close": "previous_close",
    "52 week high": "fifty_two_week_high", "52 week low": "fifty_two_week_low",
    "52w high": "fifty_two_week_high", "52w low": "fifty_two_week_low",
    "50 day ma": "fifty_day_average", "200 day ma": "two_hundred_day_average",
    "market cap": "market_cap",
}

HIST_COL = {
    "open": "Open", "high": "High", "low": "Low",
    "close": "Close", "volume": "Volume", "price": "Close",
}

GROWTH_METRICS = {
    "revenue growth yoy": ("Total Revenue", "income_stmt"),
    "net income growth": ("Net Income", "income_stmt"),
    "eps growth": ("Basic EPS", "income_stmt"),
    "gross profit growth": ("Gross Profit", "income_stmt"),
    "free cash flow growth": ("Free Cash Flow", "cashflow"),
}

ETF_FIELDS = {
    "nav": "navPrice", "net asset value": "navPrice",
    "expense ratio": "annualReportExpenseRatio",
    "aum": "totalAssets", "assets under management": "totalAssets",
    "yield": "yield", "ytd return": "ytdReturn",
    "3 year return": "threeYearAverageReturn", "5 year return": "fiveYearAverageReturn",
    "category": "category", "fund family": "fundFamily",
    "morningstar rating": "morningStarOverallRating",
}

OPTION_COLS = {
    "strike": "strike", "last price": "lastPrice", "bid": "bid", "ask": "ask",
    "volume": "volume", "open interest": "openInterest",
    "iv": "impliedVolatility", "implied volatility": "impliedVolatility",
    "in the money": "inTheMoney", "expiration": "expiration",
    "contract": "contractSymbol", "change": "change",
    "percent change": "percentChange", "last trade date": "lastTradeDate",
}

PROFILE_MAP = {
    "name": "longName", "company name": "longName", "sector": "sector",
    "industry": "industry", "country": "country", "employees": "fullTimeEmployees",
    "website": "website", "exchange": "exchange", "currency": "currency",
    "description": "longBusinessSummary",
}

# ─────────────────────────────────────────────────────────────────────────────
# PERIOD RESOLUTION
# ─────────────────────────────────────────────────────────────────────────────

def resolve_period(period_str, df_cols):
    s = str(period_str).strip().upper()
    cols = list(df_cols)
    if s in ("TTM", "LY", ""):
        return 0
    if s.startswith("LY"):
        n = abs(int(s[2:])) if s[2:] else 0
        return min(n, len(cols) - 1)
    if s.startswith("LQ"):
        n = abs(int(s[2:])) if s[2:] else 0
        return min(n, len(cols) - 1)
    try:
        year = int(s)
        for i, c in enumerate(cols):
            if hasattr(c, "year") and c.year == year:
                return i
        return None
    except ValueError:
        return None

def get_stmt(t, stmt_name, quarterly):
    if stmt_name == "income_stmt":
        return t.quarterly_income_stmt if quarterly else t.income_stmt
    if stmt_name == "balance_sheet":
        return t.quarterly_balance_sheet if quarterly else t.balance_sheet
    if stmt_name == "cashflow":
        return t.quarterly_cashflow if quarterly else t.cashflow
    return None

# ─────────────────────────────────────────────────────────────────────────────
# /model  — =MODEL(ticker, param, period, quarter, divisor)
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/model")
def model():
    ticker  = request.args.get("ticker", "").strip()
    param   = request.args.get("param",  "").strip()
    period  = request.args.get("period", "LY").strip()
    quarter = request.args.get("quarter", "").strip().upper()
    divisor = float(request.args.get("divisor", 1) or 1)

    if not ticker or not param:
        return err("ticker and param are required")

    key = param.strip().lower()
    t   = ticker_obj(ticker)

    # Key metrics
    if key in KEY_METRICS:
        try:
            val = t.info.get(KEY_METRICS[key])
            return jsonify({"value": safe(val)})
        except Exception as e:
            return err(e)

    # Live price
    if key in LIVE_PRICE_FIELDS:
        try:
            val = getattr(t.fast_info, LIVE_PRICE_FIELDS[key], None)
            return jsonify({"value": safe(val)})
        except Exception as e:
            return err(e)

    # Profile
    if key in PROFILE_MAP:
        try:
            val = t.info.get(PROFILE_MAP[key])
            return jsonify({"value": safe(val)})
        except Exception as e:
            return err(e)

    # Financial statements
    use_quarterly = bool(quarter) or period.upper().startswith("LQ")
    for field_map, stmt_name in [
        (INCOME_FIELDS, "income_stmt"),
        (BALANCE_FIELDS, "balance_sheet"),
        (CASHFLOW_FIELDS, "cashflow"),
    ]:
        if key in field_map:
            row_label = field_map[key]
            df = get_stmt(t, stmt_name, use_quarterly)
            if df is None or df.empty:
                return err(f"No data for {ticker}")
            col_idx = resolve_period(period, df.columns)
            if col_idx is None:
                return err(f"Period '{period}' not found")
            try:
                raw = df.loc[row_label].iloc[col_idx]
                val = float(raw) / divisor if raw is not None and not pd.isna(raw) else None
                return jsonify({"value": safe(val)})
            except KeyError:
                return err(f"Row '{row_label}' not in {stmt_name}")
            except Exception as e:
                return err(e)

    # Growth metrics
    if key in GROWTH_METRICS:
        row_label, stmt_name = GROWTH_METRICS[key]
        df = get_stmt(t, stmt_name, False)
        if df is None or df.empty:
            return err("No data")
        try:
            row = df.loc[row_label].dropna()
            pct = (row.iloc[0] - row.iloc[1]) / abs(row.iloc[1])
            return jsonify({"value": safe(pct)})
        except Exception as e:
            return err(e)

    # Dividends
    if key in ("dividend", "dividends", "adjusted dividend"):
        try:
            df = t.dividends
            if df.empty:
                return jsonify({"value": 0})
            df.index = pd.to_datetime(df.index)
            p = period.strip().upper()
            if p.startswith("LQ"):
                n = abs(int(p[2:])) if p[2:] else 0
                grouped = df.resample("QE").sum()
                val = float(grouped.iloc[-(n + 1)]) if len(grouped) > n else 0
            elif p.startswith("LY"):
                n = abs(int(p[2:])) if p[2:] else 0
                grouped = df.resample("YE").sum()
                val = float(grouped.iloc[-(n + 1)]) if len(grouped) > n else 0
            else:
                try:
                    year = int(p)
                    val = float(df[df.index.year == year].sum())
                except ValueError:
                    val = float(df.iloc[-1])
            return jsonify({"value": round(val / divisor, 6)})
        except Exception as e:
            return err(e)

    return err(f"Unknown parameter: '{param}'. Visit /fields for the full list.")


# ─────────────────────────────────────────────────────────────────────────────
# /modelprice  — =MODELPRICE(ticker, param, days, start, end)
# ─────────────────────────────────────────────────────────────────────────────

def normalise_ticker(sym):
    s = sym.upper().strip()
    forex = ["USD","EUR","GBP","JPY","CAD","AUD","CHF","CNY","HKD","SGD","INR","MXN","BRL","KRW","SEK","NOK"]
    if len(s) == 6 and s.isalpha():
        base, quote = s[:3], s[3:]
        if base in forex and quote in forex:
            return s + "=X"
    crypto = ["BTC","ETH","BNB","XRP","ADA","SOL","DOGE","DOT","AVAX","MATIC","LTC","LINK","UNI","ATOM"]
    for c in crypto:
        if s == c + "USD":
            return c + "-USD"
    return s

@app.route("/modelprice")
def modelprice():
    ticker = request.args.get("ticker", "").strip()
    param  = request.args.get("param",  "Close").strip()
    days   = request.args.get("days",   "").strip()
    start  = request.args.get("start",  "").strip()
    end    = request.args.get("end",    "").strip()

    if not ticker:
        return err("ticker is required")

    key = param.strip().lower()
    sym = normalise_ticker(ticker)
    t   = yf.Ticker(sym)

    # Live single value
    if not days and not start and not end:
        if key == "dividend":
            try:
                df = t.dividends
                if df.empty:
                    return jsonify({"data": []})
                df.index = df.index.strftime("%Y-%m-%d")
                rows = [{"date": d, "dividend": round(v, 6)} for d, v in df.items()]
                return jsonify({"data": rows})
            except Exception as e:
                return err(e)

        if key in LIVE_PRICE_FIELDS:
            try:
                val = getattr(t.fast_info, LIVE_PRICE_FIELDS[key], None)
                return jsonify({"value": safe(val)})
            except Exception as e:
                return err(e)

        # fallback to 1d history
        try:
            col = HIST_COL.get(key, "Close")
            h = t.history(period="1d")
            if not h.empty:
                return jsonify({"value": safe(float(h[col].iloc[-1]))})
        except Exception:
            pass
        return err(f"Unknown param: {param}")

    # Historical
    try:
        col = HIST_COL.get(key, "Close")
        if days:
            df = t.history(period="max").tail(int(days))
        elif start and end:
            s = parse_date(start)
            e = parse_date(end)
            df = t.history(start=s.strftime("%Y-%m-%d"),
                           end=(e + timedelta(days=1)).strftime("%Y-%m-%d"))
        elif start:
            s = parse_date(start)
            df = t.history(start=s.strftime("%Y-%m-%d"),
                           end=(s + timedelta(days=5)).strftime("%Y-%m-%d"))
        else:
            return err("Provide days, start, or start+end")

        if df.empty:
            return err(f"No historical data for {ticker}")

        df.index = df.index.strftime("%Y-%m-%d")
        rows = [{"date": d, "value": safe(float(row[col]))} for d, row in df.iterrows()]
        return jsonify({"data": rows, "ticker": ticker, "param": param})
    except Exception as e:
        return err(e)


# ─────────────────────────────────────────────────────────────────────────────
# /modelfunds  — =MODELFUNDS(ticker, param)
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/modelfunds")
def modelfunds():
    ticker = request.args.get("ticker", "").strip()
    param  = request.args.get("param",  "").strip()
    if not ticker or not param:
        return err("ticker and param are required")
    key = param.strip().lower()
    yf_key = ETF_FIELDS.get(key)
    if not yf_key:
        return err(f"Unknown ETF param: '{param}'. Supported: {', '.join(ETF_FIELDS)}")
    try:
        val = ticker_obj(ticker).info.get(yf_key)
        return jsonify({"value": safe(val)})
    except Exception as e:
        return err(e)


# ─────────────────────────────────────────────────────────────────────────────
# /modeloptions  — =MODELOPTIONS(ticker, type, param, date)
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/modeloptions")
def modeloptions():
    ticker   = request.args.get("ticker", "").strip()
    opt_type = request.args.get("type",   "Call").strip().capitalize()
    param    = request.args.get("param",  "Strike").strip()
    exp_date = request.args.get("date",   "").strip()

    if not ticker:
        return err("ticker is required")
    key = param.strip().lower()
    col = OPTION_COLS.get(key)
    if not col:
        return err(f"Unknown option param: '{param}'")
    try:
        t = ticker_obj(ticker)
        expirations = t.options
        if not expirations:
            return err(f"No options data for {ticker}")
        if exp_date:
            target = exp_date if "-" in exp_date else parse_date(exp_date).strftime("%Y-%m-%d")
            exp = next((e for e in expirations if e == target), expirations[0])
        else:
            exp = expirations[0]
        chain = t.option_chain(exp)
        df = chain.calls if opt_type == "Call" else chain.puts
        if col not in df.columns:
            return err(f"Column '{col}' not in options data")
        vals = [safe(v) for v in df[col].tolist()]
        return jsonify({"data": vals, "expiration": exp, "type": opt_type})
    except Exception as e:
        return err(e)


# ─────────────────────────────────────────────────────────────────────────────
# /dump  — Statement Dump
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/dump")
def dump():
    ticker = request.args.get("ticker", "").strip()
    stype  = request.args.get("type",   "income").lower()
    freq   = request.args.get("freq",   "annual").lower()
    if not ticker:
        return err("ticker is required")
    t = ticker_obj(ticker)
    try:
        if stype == "income":
            df = t.quarterly_income_stmt if freq == "quarterly" else t.income_stmt
        elif stype == "balance":
            df = t.quarterly_balance_sheet if freq == "quarterly" else t.balance_sheet
        elif stype == "cashflow":
            df = t.quarterly_cashflow if freq == "quarterly" else t.cashflow
        elif stype == "metrics":
            info = t.info
            return jsonify({"ticker": ticker, "metrics": {k: safe(info.get(v)) for k, v in KEY_METRICS.items()}})
        else:
            return err("type must be income|balance|cashflow|metrics")
        if df is None or df.empty:
            return err(f"No {stype} data for {ticker}")
        df.columns = [c.strftime("%Y-%m-%d") if hasattr(c, "strftime") else str(c) for c in df.columns]
        out = {str(row): {col: safe(val) for col, val in df.loc[row].items()} for row in df.index}
        return jsonify({"ticker": ticker, "statement": stype, "frequency": freq, "data": out})
    except Exception as e:
        return err(e)


# ─────────────────────────────────────────────────────────────────────────────
# /screener
# ─────────────────────────────────────────────────────────────────────────────

SCREENER_UNIVERSE = [
    "AAPL","MSFT","GOOGL","AMZN","META","NVDA","TSLA","BRK-B","JPM","JNJ",
    "V","PG","UNH","HD","MA","ABBV","MRK","CVX","PEP","KO","AVGO","COST",
    "WMT","MCD","BAC","LLY","TMO","ABT","DIS","ADBE","CRM","NFLX","ORCL",
    "INTC","AMD","QCOM","TXN","AMAT","MU","CSCO","GE","CAT","DE","BA","HON",
    "XOM","COP","OXY","SLB","GS","MS","C","WFC","PFE","BMY","GILD","REGN",
]

@app.route("/screener")
def screener():
    sector        = request.args.get("sector", "").strip()
    industry      = request.args.get("industry", "").strip()
    min_mktcap    = request.args.get("min_market_cap", "")
    max_mktcap    = request.args.get("max_market_cap", "")
    min_div_yield = request.args.get("min_dividend_yield", "")
    max_pe        = request.args.get("max_pe", "")
    min_pe        = request.args.get("min_pe", "")
    exchange      = request.args.get("exchange", "").strip()
    min_beta      = request.args.get("min_beta", "")
    max_beta      = request.args.get("max_beta", "")
    limit         = int(request.args.get("limit", 20))

    results = []
    for sym in SCREENER_UNIVERSE:
        if len(results) >= limit:
            break
        try:
            info = yf.Ticker(sym).info
            mc  = info.get("marketCap", 0) or 0
            pe  = info.get("trailingPE")
            dy  = info.get("dividendYield", 0) or 0
            b   = info.get("beta")
            sec = info.get("sector", "")
            ind = info.get("industry", "")
            exc = info.get("exchange", "")
            if sector      and sector.lower()   not in sec.lower(): continue
            if industry    and industry.lower() not in ind.lower(): continue
            if exchange    and exchange.upper() not in exc.upper(): continue
            if min_mktcap  and mc < float(min_mktcap): continue
            if max_mktcap  and mc > float(max_mktcap): continue
            if min_div_yield and dy < float(min_div_yield): continue
            if max_pe and pe and pe > float(max_pe): continue
            if min_pe and pe and pe < float(min_pe): continue
            if min_beta and b and b < float(min_beta): continue
            if max_beta and b and b > float(max_beta): continue
            results.append({
                "ticker": sym, "name": info.get("longName", sym),
                "sector": sec, "industry": ind, "exchange": exc,
                "market_cap": safe(mc), "price": safe(info.get("currentPrice")),
                "pe_ratio": safe(pe), "dividend_yield": safe(dy), "beta": safe(b),
            })
        except Exception:
            continue
    return jsonify({"count": len(results), "results": results})


# ─────────────────────────────────────────────────────────────────────────────
# /trending
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/trending")
def trending():
    category = request.args.get("category", "gainers").lower()
    try:
        screen_map = {"gainers": "day_gainers", "losers": "day_losers", "active": "most_actives"}
        key = screen_map.get(category, "day_gainers")
        df = yf.screen(key) or {}
        quotes = df.get("quotes", []) if isinstance(df, dict) else []
        out = [{
            "ticker":     q.get("symbol"),
            "name":       q.get("longName") or q.get("shortName"),
            "price":      safe(q.get("regularMarketPrice")),
            "change":     safe(q.get("regularMarketChange")),
            "change_pct": safe(q.get("regularMarketChangePercent")),
            "volume":     safe(q.get("regularMarketVolume")),
            "market_cap": safe(q.get("marketCap")),
        } for q in quotes[:20]]
        return jsonify({"category": category, "stocks": out})
    except Exception as e:
        return err(e)


# ─────────────────────────────────────────────────────────────────────────────
# DISCOVERY & HEALTH
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/fields")
def fields():
    return jsonify({
        "MODEL": {
            "income_statement": list(INCOME_FIELDS),
            "balance_sheet":    list(BALANCE_FIELDS),
            "cash_flow":        list(CASHFLOW_FIELDS),
            "key_metrics":      list(KEY_METRICS),
            "growth_metrics":   list(GROWTH_METRICS),
            "dividends":        ["dividend", "adjusted dividend"],
            "profile":          list(PROFILE_MAP),
        },
        "MODELPRICE": {
            "live":       list(LIVE_PRICE_FIELDS),
            "historical": list(HIST_COL),
            "special":    ["dividend"],
        },
        "MODELFUNDS": list(ETF_FIELDS),
        "MODELOPTIONS": list(OPTION_COLS),
        "periods": ["TTM","LY","LY-1","LY-2","LQ","LQ-1","LQ-2","2024","2023","..."],
    })

@app.route("/")
@app.route("/health")
def health():
    return jsonify({"status": "ok", "version": "2.0"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)
