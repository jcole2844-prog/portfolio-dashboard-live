"""
data_feed.py — All data fetching for the portfolio dashboard.
Uses Yahoo Finance (yfinance) — no API key required.
"""

import warnings
warnings.filterwarnings("ignore")

import io
import logging
import time
import numpy as np
import pandas as pd
import requests
import streamlit as st
import yfinance as yf
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor

# Silence yfinance's noisy "possibly delisted / no price data" console messages.
# These cases are already handled gracefully (the affected cells show "—").
logging.getLogger("yfinance").setLevel(logging.CRITICAL)

_HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

_MAX_WORKERS = 8


def _parallel_map(fn, items):
    """Run fn over items in a thread pool, returning {item: result}."""
    out = {}
    with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as ex:
        futures = {ex.submit(fn, it): it for it in items}
        for fut in futures:
            it = futures[fut]
            try:
                out[it] = fut.result()
            except Exception:
                out[it] = None
    return out

EXCEL_PATH = r"C:\Users\Julie\Downloads\Code Holdings Query\Holdings Listing 06172026.xlsx"
CD_EXCEL_PATH = r"C:\Users\Julie\Downloads\Code Holdings Query\CD_Holdings.xlsx"

# Dow Jones 30 constituents (hardcoded for stability), current as of 2025-2026.
# 2024 changes: AMZN replaced WBA; NVDA replaced INTC; SHW replaced DOW.
DOW_30 = (
    "AAPL", "AMGN", "AMZN", "AXP", "BA", "CAT", "CRM", "CSCO", "CVX", "DIS",
    "GS", "HD", "HON", "IBM", "JNJ", "JPM", "KO", "MCD", "MMM", "MRK",
    "MSFT", "NKE", "NVDA", "PG", "SHW", "TRV", "UNH", "V", "VZ", "WMT"
)

# Watchlist tickers (not held — for monitoring). Yahoo uses '-' not '.', so
# BRK.B → BRK-B. Edit this list to change the watchlist.
WATCHLIST = (
    "SPCX", "UBER", "LYFT", "INTC", "ANET", "LRCX", "KLAC", "CRWD", "PANW",
    "NOW", "GE", "GEV", "CAT", "HII", "ETN", "HON", "BKNG", "ABNB", "LOW",
    "MCD", "NKE", "JNJ", "ABBV", "MRK", "THM", "UNH", "VRTX", "REGN", "BRK-B",
    "PSUS", "WFC", "V", "MS", "AXP", "VST", "CEG", "NEE", "SO", "ETR", "NRG",
    "BE", "NVT", "D", "FCX", "SCCO", "NEM", "LIN", "SHW", "MLM", "VMC",
    "TMUS", "VZ", "NFLX", "DIS", "AMBA", "CRSP", "INOD", "WLDN", "LUNR",
    "IRDM", "MDGL", "EOSE", "MU", "AMD",
)


# Fallback large-cap S&P 500 names, used only if the Wikipedia fetch fails.
_SP500_FALLBACK = (
    "AAPL", "MSFT", "NVDA", "AMZN", "GOOGL", "GOOG", "META", "BRK-B", "LLY", "AVGO",
    "TSLA", "JPM", "WMT", "UNH", "XOM", "V", "MA", "PG", "JNJ", "COST",
    "HD", "MRK", "ABBV", "CVX", "CRM", "BAC", "NFLX", "KO", "AMD", "PEP",
    "TMO", "ADBE", "LIN", "MCD", "CSCO", "ACN", "WFC", "ABT", "GE", "DHR",
    "TXN", "PM", "IBM", "QCOM", "DIS", "CAT", "VZ", "INTU", "AMGN", "NOW",
    "ISRG", "CMCSA", "SPGI", "UBER", "GS", "PFE", "RTX", "T", "NEE", "HON",
    "LOW", "AXP", "BKNG", "ETN", "BLK", "UNP", "PGR", "SYK", "C", "TJX",
    "MS", "BSX", "VRTX", "ADP", "MDT", "GILD", "LMT", "CB", "MMC", "SCHW",
    "DE", "ADI", "PLD", "BMY", "AMAT", "FI", "MDLZ", "SBUX", "CI", "MO",
    "SO", "DUK", "BX", "REGN", "ICE", "ELV", "ZTS", "SHW", "EQIX", "KLAC",
)

# ── Static / slow-changing data ───────────────────────────────────────────────

def _sp500_from_grokipedia():
    """Parse the S&P 500 constituents table from Grokipedia."""
    resp = requests.get(
        "https://grokipedia.com/page/List_of_S%26P_500_companies",
        headers=_HTTP_HEADERS,
        timeout=20,
    )
    resp.raise_for_status()
    tables = pd.read_html(io.StringIO(resp.text))
    for t in tables:
        if "Ticker" in t.columns:
            return (
                t["Ticker"].astype(str)
                .str.replace(".", "-", regex=False)
                .str.strip()
                .tolist()
            )
    return []


def _sp500_from_wikipedia():
    """Parse the S&P 500 constituents table from Wikipedia (fallback source)."""
    resp = requests.get(
        "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies",
        headers=_HTTP_HEADERS,
        timeout=15,
    )
    resp.raise_for_status()
    tables = pd.read_html(io.StringIO(resp.text), attrs={"id": "constituents"})
    return (
        tables[0]["Symbol"].str.replace(".", "-", regex=False).str.strip().tolist()
    )


@st.cache_data(ttl=3600, show_spinner=False)
def get_sp500_tickers():
    """Fetch S&P 500 ticker list (cached 1 hour).

    Primary source is Wikipedia (most current); falls back to Grokipedia, then a
    hardcoded large-cap list if both network fetches fail.
    """
    for source in (_sp500_from_wikipedia, _sp500_from_grokipedia):
        try:
            syms = source()
            if syms:
                return tuple(syms)
        except Exception:
            continue
    return _SP500_FALLBACK


# ── Excel source: Google Drive (cloud) with local-file fallback ───────────────

# Maps each workbook to its file-id key in the [google_drive] secrets section.
_DRIVE_FILE_KEYS = {"holdings": "workbook1_id", "fixed_income": "workbook2_id"}


def _drive_credentials(sa_value):
    """Build read-only Drive credentials from the [google_drive].service_account
    secret, which may be a JSON string or an already-parsed mapping."""
    import json
    from google.oauth2 import service_account
    info = json.loads(sa_value) if isinstance(sa_value, str) else dict(sa_value)
    return service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )


def _read_drive_xlsx(file_id: str, creds) -> pd.DataFrame:
    """Download a private .xlsx from Google Drive into a DataFrame."""
    import io as _io
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload

    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    request = service.files().get_media(fileId=file_id)
    buf = _io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return pd.read_excel(buf, engine="openpyxl")


def _read_excel(local_path: str, which: str) -> pd.DataFrame:
    """Read an Excel workbook from Google Drive when configured (cloud
    deployment), otherwise from the local file path (running on your PC).

    Cloud config lives in st.secrets under a [google_drive] section with:
        service_account = '''<JSON key>'''
        workbook1_id    = "<holdings file id>"
        workbook2_id    = "<fixed income file id>"
    """
    try:
        gd = st.secrets["google_drive"]
        file_id = gd.get(_DRIVE_FILE_KEYS[which])
        if file_id:
            creds = _drive_credentials(gd["service_account"])
            return _read_drive_xlsx(file_id, creds)
    except Exception:
        pass
    return pd.read_excel(local_path, engine="openpyxl")


# ── Google Sheets source (editable, persistent store) ─────────────────────────

# The "Portfolio Dashboard" spreadsheet (3 tabs: Holdings, Watchlist, Fixed
# Income). The id is not secret — access is controlled by sharing the sheet with
# the service account.
SHEET_ID = "1CwH9jfOWJg4pejyAn4VlRrohURSHoKzdrux9FWuzPd0"


def _sheet_enabled() -> bool:
    """True when a Google service account is configured in st.secrets."""
    try:
        return "service_account" in st.secrets["google_drive"]
    except Exception:
        return False


def _sheets_service(write: bool = False):
    """Build a Google Sheets API client from the service-account secret."""
    import json
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    scope = ("https://www.googleapis.com/auth/spreadsheets" if write
             else "https://www.googleapis.com/auth/spreadsheets.readonly")
    sa = st.secrets["google_drive"]["service_account"]
    info = json.loads(sa) if isinstance(sa, str) else dict(sa)
    creds = service_account.Credentials.from_service_account_info(info, scopes=[scope])
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _read_sheet_df(tab: str) -> pd.DataFrame:
    """Read a worksheet tab into a DataFrame (all cells as strings)."""
    svc = _sheets_service(write=False)
    res = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"'{tab}'"
    ).execute()
    vals = res.get("values", [])
    if not vals:
        return pd.DataFrame()
    header = [h.strip() for h in vals[0]]
    width = len(header)
    rows = [(r + [None] * width)[:width] for r in vals[1:]]
    return pd.DataFrame(rows, columns=header)


def write_sheet_df(tab: str, df: pd.DataFrame):
    """Overwrite a worksheet tab with the given DataFrame (header + rows).
    Blank/NaN cells are written as empty strings. Values are USER_ENTERED so
    Sheets interprets numbers, dates and percentages naturally."""
    svc = _sheets_service(write=True)

    def _cell(v):
        if v is None:
            return ""
        if isinstance(v, float) and pd.isna(v):
            return ""
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        return v

    values = [list(map(str, df.columns))]
    for _, row in df.iterrows():
        values.append([_cell(v) for v in row.tolist()])

    svc.spreadsheets().values().clear(
        spreadsheetId=SHEET_ID, range=f"'{tab}'"
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"'{tab}'!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()


# ── Change log (audit trail of edits) ─────────────────────────────────────────

LOG_TAB = "Change Log"
_LOG_HEADER = ["Date", "Time", "Source", "Ticker", "Name",
               "Qty Change", "Cost Basis Change"]


def _ensure_log_tab(svc):
    """Create the Change Log tab (with header) if it doesn't exist yet."""
    meta = svc.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if LOG_TAB not in titles:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": LOG_TAB}}}]},
        ).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=SHEET_ID, range=f"'{LOG_TAB}'!A1",
            valueInputOption="USER_ENTERED", body={"values": [_LOG_HEADER]},
        ).execute()


def append_change_log(entries: list):
    """Append change-log rows. Each entry is a dict with keys:
    date, time, source, ticker, name, qty_change, cost_change."""
    if not entries:
        return
    svc = _sheets_service(write=True)
    _ensure_log_tab(svc)
    values = [[e["date"], e["time"], e["source"], e["ticker"], e["name"],
               e["qty_change"], e["cost_change"]] for e in entries]
    svc.spreadsheets().values().append(
        spreadsheetId=SHEET_ID, range=f"'{LOG_TAB}'!A1",
        valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS",
        body={"values": values},
    ).execute()


@st.cache_data(ttl=30, show_spinner=False)
def load_change_log() -> pd.DataFrame:
    """Read the Change Log tab. Returns empty DataFrame if it doesn't exist."""
    if not _sheet_enabled():
        return pd.DataFrame(columns=_LOG_HEADER)
    try:
        df = _read_sheet_df(LOG_TAB)
        if df.empty:
            return pd.DataFrame(columns=_LOG_HEADER)
        for c in ("Qty Change", "Cost Basis Change"):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        return df
    except Exception:
        return pd.DataFrame(columns=_LOG_HEADER)


def _parse_num(v):
    """Parse a possibly-text number ('10,000', '$1,234.50') to float."""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.replace(",", "").replace("$", "").strip()
        if s == "":
            return None
        try:
            return float(s)
        except ValueError:
            return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _parse_rate(v):
    """Parse a rate to a decimal ratio. '4.150%' → 0.0415; 0.0415 → 0.0415."""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.replace("%", "").replace(",", "").strip()
        if s == "":
            return None
        try:
            return float(s) / 100.0
        except ValueError:
            return None
    try:
        return float(v)          # already a decimal ratio from Excel
    except (TypeError, ValueError):
        return None


@st.cache_data(ttl=60, show_spinner=False)
def load_portfolio():
    """Load portfolio holdings. Source order: Google Sheet → Drive/local Excel.
    Columns: Ticker, Name, Total Quantity, Total Cost Basis, Avg Basis/Sh."""
    df = None
    if _sheet_enabled():
        try:
            df = _read_sheet_df("Holdings")
        except Exception:
            df = None
    if df is None or df.empty:
        df = _read_excel(EXCEL_PATH, "holdings")

    df.columns = df.columns.str.strip()
    df = df[df.iloc[:, 0].notna() & (df.iloc[:, 0].astype(str).str.strip() != "")]
    df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.strip().str.upper()
    # Sheets returns text — coerce the numeric columns.
    for c in ("Total Quantity", "Total Cost Basis", "Avg Basis/Sh"):
        if c in df.columns:
            df[c] = df[c].map(_parse_num)
    return df.reset_index(drop=True)


@st.cache_data(ttl=300, show_spinner=False)
def load_fixed_income():
    """Load fixed-income holdings. Source order: Google Sheet → Drive/local Excel.
    Columns: Symbol, Description, Type, Quantity, Acquisition Date,
    Maturity Date, Coupon, YTM."""
    df = None
    if _sheet_enabled():
        try:
            df = _read_sheet_df("Fixed Income")
        except Exception:
            df = None
    if df is None or df.empty:
        df = _read_excel(CD_EXCEL_PATH, "fixed_income")

    df.columns = df.columns.str.strip()
    df = df[df["Symbol"].notna() & (df["Symbol"].astype(str).str.strip() != "")]
    # Coerce text → numbers / rates / dates (works for both Sheet and Excel).
    if "Quantity" in df.columns:
        df["Quantity"] = df["Quantity"].map(_parse_num)
    for c in ("Coupon", "YTM"):
        if c in df.columns:
            df[c] = df[c].map(_parse_rate)
    for c in ("Acquisition Date", "Maturity Date"):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df.reset_index(drop=True)


@st.cache_data(ttl=60, show_spinner=False)
def get_watchlist():
    """Watchlist tickers from the Google Sheet's Watchlist tab (Ticker column),
    falling back to the hardcoded WATCHLIST list. Names are pulled live from
    Yahoo elsewhere, so only the Ticker column is used here."""
    if _sheet_enabled():
        try:
            df = _read_sheet_df("Watchlist")
            if "Ticker" in df.columns:
                seen, out = set(), []
                for t in df["Ticker"]:
                    s = str(t).strip().upper().replace(".", "-")
                    if s and s != "NONE" and s not in seen:
                        seen.add(s)
                        out.append(s)
                if out:
                    return tuple(out)
        except Exception:
            pass
    return WATCHLIST


# ── Price / quote data ────────────────────────────────────────────────────────

def _parse_field(data, tickers, field="Close"):
    """Extract a per-ticker OHLC field series from a yf.download result."""
    out = {}
    if data.empty:
        return out
    if isinstance(data.columns, pd.MultiIndex):
        if field in data.columns.get_level_values(0):
            fdf = data[field]
            for t in tickers:
                if t in fdf.columns:
                    out[t] = fdf[t].dropna()
    else:
        if field in data.columns and len(tickers) == 1:
            out[tickers[0]] = data[field].dropna()
    return out


def _parse_closes(data, tickers):
    """Extract per-ticker close series from yf.download result."""
    return _parse_field(data, tickers, "Close")


@st.cache_data(ttl=30, show_spinner=False)
def get_quotes(tickers: tuple) -> dict:
    """
    Batch-fetch current price + day change for a tuple of tickers.
    Returns dict keyed by ticker →
        {current, prev_close, open, day_change, day_change_pct}
    'open' is the latest session's opening price. Cached 30 seconds.
    """
    if not tickers:
        return {}

    result = {}
    try:
        data = yf.download(
            list(tickers),
            period="5d",
            auto_adjust=True,
            progress=False,
            threads=True,
        )
        closes = _parse_closes(data, list(tickers))
        opens  = _parse_field(data, list(tickers), "Open")

        for t in tickers:
            s = closes.get(t, pd.Series(dtype=float)).dropna()
            o = opens.get(t, pd.Series(dtype=float)).dropna()
            opn = float(o.iloc[-1]) if len(o) else None
            if len(s) >= 2:
                cur = float(s.iloc[-1])
                prev = float(s.iloc[-2])
                chg = cur - prev
                pct = chg / prev * 100
            elif len(s) == 1:
                cur = float(s.iloc[-1])
                prev = cur
                chg = 0.0
                pct = 0.0
            else:
                cur = prev = chg = pct = None

            result[t] = {
                "current": cur,
                "prev_close": prev,
                "open": opn,
                "day_change": chg,
                "day_change_pct": pct,
            }
    except Exception:
        for t in tickers:
            result[t] = {"current": None, "prev_close": None, "open": None,
                         "day_change": None, "day_change_pct": None}

    return result


@st.cache_data(ttl=120, show_spinner=False)
def get_premarket_quotes(tickers: tuple) -> dict:
    """
    Fetch pre/after-market prices for tickers using 1-minute intraday data.
    Returns dict keyed by ticker → {pre_price, pre_change_pct}
    Cached 2 minutes.
    """
    if not tickers:
        return {}

    result = {}
    try:
        data = yf.download(
            list(tickers),
            period="1d",
            interval="1m",
            prepost=True,
            auto_adjust=True,
            progress=False,
            threads=True,
        )
        closes = _parse_closes(data, list(tickers))

        # Regular session open: 09:30 ET. Pre-market: before 09:30.
        for t in tickers:
            s = closes.get(t, pd.Series(dtype=float))
            if s.empty:
                result[t] = {"pre_price": None, "pre_change_pct": None}
                continue

            # Localize to ET if tz-aware
            try:
                s_et = s.tz_convert("America/New_York") if s.index.tz else s
            except Exception:
                s_et = s

            # Pre-market rows: hour < 9 or (hour == 9 and minute < 30)
            pre = s_et[
                (s_et.index.hour < 9)
                | ((s_et.index.hour == 9) & (s_et.index.minute < 30))
            ]

            if not pre.empty:
                pre_price = float(pre.iloc[-1])
                # Compare to yesterday's close (first regular-session price as reference)
                first_reg = s_et[
                    (s_et.index.hour == 9) & (s_et.index.minute >= 30)
                    | (s_et.index.hour > 9)
                ]
                ref = float(first_reg.iloc[0]) if not first_reg.empty else float(s_et.iloc[0])
                pre_pct = (pre_price - ref) / ref * 100 if ref else None
            else:
                # Market already open — use most recent price
                pre_price = float(s_et.iloc[-1])
                pre_pct = None

            result[t] = {"pre_price": pre_price, "pre_change_pct": pre_pct}
    except Exception:
        for t in tickers:
            result[t] = {"pre_price": None, "pre_change_pct": None}

    return result


@st.cache_data(ttl=120, show_spinner=False)
def get_ext_hours_prices(tickers: tuple) -> dict:
    """Latest extended-hours (pre- or post-market) price per ticker, if the most
    recent print is outside the 9:30 AM–4:00 PM ET regular session. Returns
    {ticker: price or None}; None when the market is open (no separate
    extended-hours print) or no data is available. Cached 2 minutes."""
    if not tickers:
        return {}
    result = {t: None for t in tickers}
    try:
        data = yf.download(
            list(tickers),
            period="1d",
            interval="1m",
            prepost=True,
            auto_adjust=True,
            progress=False,
            threads=True,
        )
        closes = _parse_closes(data, list(tickers))
        for t in tickers:
            s = closes.get(t, pd.Series(dtype=float)).dropna()
            if s.empty:
                continue
            try:
                s_et = s.tz_convert("America/New_York") if s.index.tz else s
            except Exception:
                s_et = s
            last = s_et.index[-1]
            h, m = last.hour, last.minute
            in_regular = ((h > 9) or (h == 9 and m >= 30)) and (h < 16)
            if not in_regular:
                result[t] = float(s_et.iloc[-1])
    except Exception:
        pass
    return result


# ── Market indices ────────────────────────────────────────────────────────────

# Index definitions: (Yahoo symbol, display name, CNBC quote URL).
# Nasdaq card tracks the Nasdaq 100 (^NDX), not the Composite (^IXIC).
_INDICES = (
    ("^DJI",  "Dow Jones",        "https://www.cnbc.com/quotes/.DJI"),
    ("^GSPC", "S&P 500",          "https://www.cnbc.com/quotes/.SPX"),
    ("^IXIC", "Nasdaq Composite", "https://www.cnbc.com/quotes/.IXIC"),
    ("^NDX",  "Nasdaq 100",       "https://www.cnbc.com/quotes/.NDX"),
)


@st.cache_data(ttl=30, show_spinner=False)
def get_index_quotes() -> list:
    """
    Current level + day change for the Dow 30, S&P 500 and Nasdaq 100.
    Returns a list of {symbol, name, level, change, change_pct, url}, in
    display order. Cached 30 seconds.
    """
    syms = tuple(s for s, _, _ in _INDICES)
    quotes = get_quotes(syms)
    out = []
    for sym, name, url in _INDICES:
        q = quotes.get(sym, {})
        out.append({
            "symbol":     sym,
            "name":       name,
            "level":      q.get("current"),
            "change":     q.get("day_change"),
            "change_pct": q.get("day_change_pct"),
            "url":        url,
        })
    return out


# ── Futures & commodities ─────────────────────────────────────────────────────

# (Yahoo symbol, display name, CNBC quote URL)
_FUTURES = (
    ("YM=F", "Dow",    "https://www.cnbc.com/quotes/@DJ.1"),
    ("ES=F", "S&P",    "https://www.cnbc.com/quotes/@SP.1"),
    ("NQ=F", "Nasdaq", "https://www.cnbc.com/quotes/@ND.1"),
)
_COMMODITIES = (
    ("GC=F",    "Gold",    "https://www.cnbc.com/quotes/@GC.1"),
    ("BTC-USD", "Bitcoin", "https://www.cnbc.com/quotes/BTC.CM="),
    ("CL=F",    "WTI",     "https://www.cnbc.com/quotes/@CL.1"),
    ("BZ=F",    "Brent",   "https://www.cnbc.com/quotes/@LCO.1"),
)


@st.cache_data(ttl=30, show_spinner=False)
def get_market_extras() -> dict:
    """Current value + day change for futures and commodities. Returns
    {'futures': [...], 'commodities': [...]} with name/value/change/change_pct/url.
    Cached 30 seconds."""
    syms = tuple(s for s, _, _ in (_FUTURES + _COMMODITIES))
    q = get_quotes(syms)

    def _pack(group):
        rows = []
        for sym, name, url in group:
            v = q.get(sym, {})
            rows.append({
                "name":       name,
                "value":      v.get("current"),
                "change":     v.get("day_change"),
                "change_pct": v.get("day_change_pct"),
                "url":        url,
            })
        return rows

    return {"futures": _pack(_FUTURES), "commodities": _pack(_COMMODITIES)}


# ── Fundamental / static info ────────────────────────────────────────────────

def _norm_div_yield(raw):
    """Normalize dividend yield to a percent number.
    yfinance >= 0.2.40 returns dividendYield already as a percent
    (e.g. 0.36 = 0.36%, 1.97 = 1.97%), so no scaling is applied."""
    if not raw:
        return 0.0
    return float(raw)


@st.cache_data(ttl=300, show_spinner=False)
def get_fundamentals(tickers: tuple) -> dict:
    """
    Fetch P/E, forward P/E, dividend yield, 52-week high/low, long name.
    Cached 5 minutes.
    """
    def _one(t):
        # Yahoo intermittently returns an empty dict / "Invalid Crumb" 401 under
        # concurrency. Retry a few times so a single hiccup doesn't blank a row.
        info = {}
        for attempt in range(4):
            try:
                info = yf.Ticker(t).info or {}
            except Exception:
                info = {}
            if any(info.get(k) is not None for k in
                   ("fiftyTwoWeekHigh", "trailingPE", "regularMarketPrice", "dayHigh")):
                break
            time.sleep(0.5)
        return {
            "name":         info.get("longName", t),
            "sector":       info.get("sector"),
            "pe":           info.get("trailingPE"),
            "forward_pe":   info.get("forwardPE"),
            # yfinance >= 0.2.x returns dividendYield already as a percent (e.g. 1.97 = 1.97%).
            # Older versions returned a ratio (0.0197). Normalize: ratios are < 1.
            "div_yield":    _norm_div_yield(info.get("dividendYield")),
            "day_high":     info.get("dayHigh") or info.get("regularMarketDayHigh"),
            "day_low":      info.get("dayLow") or info.get("regularMarketDayLow"),
            "high_52w":     info.get("fiftyTwoWeekHigh"),
            "low_52w":      info.get("fiftyTwoWeekLow"),
            "market_cap":   info.get("marketCap"),
        }

    raw = _parallel_map(_one, tickers)
    return {t: (v if v is not None else {}) for t, v in raw.items()}


@st.cache_data(ttl=3600, show_spinner=False)
def get_earnings_dates(tickers: tuple) -> dict:
    """Next scheduled earnings date per ticker (or None). Cached 1 hour."""
    def _one(t):
        try:
            cal = yf.Ticker(t).calendar
            ed = cal.get("Earnings Date") if isinstance(cal, dict) else None
            if ed:
                dates = [d for d in ed if d is not None]
                if dates:
                    return min(dates)
        except Exception:
            return None
        return None

    return _parallel_map(_one, tickers)


# ── Analyst data ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=300, show_spinner=False)
def get_analyst_actions(tickers: tuple) -> pd.DataFrame:
    """
    Fetch recent analyst upgrades/downgrades for tickers (last 60 days).
    Returns a DataFrame sorted by date descending.
    Cached 5 minutes.
    """
    cutoff = pd.Timestamp.now(tz="UTC") - pd.Timedelta(days=60)

    def _one(t):
        ud = yf.Ticker(t).upgrades_downgrades
        if ud is None or ud.empty:
            return []
        ud = ud.copy().reset_index()
        date_col = next((c for c in ud.columns if "date" in c.lower()), None)
        if date_col is None:
            return []
        ud[date_col] = pd.to_datetime(ud[date_col], utc=True, errors="coerce")
        recent = ud[ud[date_col] >= cutoff]
        out = []
        for _, r in recent.iterrows():
            action = str(r.get("Action", "")).strip().lower()
            out.append({
                "Date":   r[date_col].strftime("%Y-%m-%d"),
                "Ticker": t,
                "Firm":   r.get("Firm", ""),
                "Action": r.get("Action", ""),
                "From":   r.get("FromGrade", ""),
                "To":     r.get("ToGrade", ""),
                "_type":  "upgrade" if "up" in action else ("downgrade" if "down" in action else "maintain"),
            })
        return out

    raw = _parallel_map(_one, tickers)
    rows = []
    for lst in raw.values():
        if lst:
            rows.extend(lst)

    if not rows:
        return pd.DataFrame()
    return (
        pd.DataFrame(rows)
        .sort_values("Date", ascending=False)
        .reset_index(drop=True)
    )


@st.cache_data(ttl=300, show_spinner=False)
def get_recommendations(tickers: tuple) -> dict:
    """
    Fetch analyst recommendation summary (buy/hold/sell counts) for each ticker.
    Returns dict: ticker → {opinion, strong_buy, buy, hold, sell, strong_sell,
                             total, buy_pct, sell_pct}
    Cached 5 minutes.
    """
    def _one(t):
        rec = yf.Ticker(t).recommendations_summary
        if rec is None or rec.empty:
            return None
        row = rec.iloc[0]
        sb  = int(row.get("strongBuy",  0) or 0)
        b   = int(row.get("buy",        0) or 0)
        h   = int(row.get("hold",       0) or 0)
        s   = int(row.get("sell",       0) or 0)
        ss  = int(row.get("strongSell", 0) or 0)
        total = sb + b + h + s + ss
        if total == 0:
            return None

        buy_pct  = (sb + b) / total * 100
        sell_pct = (s + ss) / total * 100

        if buy_pct >= 70:
            opinion = "Strong Buy"
        elif buy_pct >= 50:
            opinion = "Buy"
        elif sell_pct >= 60:
            opinion = "Strong Sell"
        elif sell_pct >= 40:
            opinion = "Sell"
        else:
            opinion = "Hold"

        return {
            "opinion":     opinion,
            "strong_buy":  sb,
            "buy":         b,
            "hold":        h,
            "sell":        s,
            "strong_sell": ss,
            "total":       total,
            "buy_pct":     buy_pct,
            "sell_pct":    sell_pct,
        }

    raw = _parallel_map(_one, tickers)
    return {t: v for t, v in raw.items() if v is not None}


# ── Technical indicators ─────────────────────────────────────────────────────

def _rsi(prices: pd.Series, period: int = 14) -> pd.Series:
    delta = prices.diff()
    gain  = delta.clip(lower=0).rolling(period).mean()
    loss  = (-delta.clip(upper=0)).rolling(period).mean()
    rs    = gain / loss.replace(0, np.nan)
    return 100 - 100 / (1 + rs)


def _macd(prices: pd.Series):
    ema12  = prices.ewm(span=12, adjust=False).mean()
    ema26  = prices.ewm(span=26, adjust=False).mean()
    macd   = ema12 - ema26
    signal = macd.ewm(span=9, adjust=False).mean()
    return macd, signal


@st.cache_data(ttl=300, show_spinner=False)
def get_technical_signal(ticker: str) -> str:
    """
    Return a short string describing any technical signals triggered in the
    last trading session (RSI, MACD crossover, Golden/Death cross, 52W high).
    Cached 5 minutes per ticker.
    """
    try:
        hist = yf.Ticker(ticker).history(period="200d", auto_adjust=True)
        if len(hist) < 30:
            return ""
        closes = hist["Close"].dropna()
        signals = []

        # RSI
        rsi = _rsi(closes)
        if not rsi.empty:
            r = rsi.iloc[-1]
            if r < 30:
                signals.append(f"RSI Oversold ({r:.0f})")
            elif r > 70:
                signals.append(f"RSI Overbought ({r:.0f})")

        # MACD cross (last 2 bars)
        if len(closes) >= 35:
            macd, signal = _macd(closes)
            if len(macd) >= 2:
                if macd.iloc[-2] < signal.iloc[-2] and macd.iloc[-1] > signal.iloc[-1]:
                    signals.append("MACD Bull Cross")
                elif macd.iloc[-2] > signal.iloc[-2] and macd.iloc[-1] < signal.iloc[-1]:
                    signals.append("MACD Bear Cross")

        # Golden / Death cross (50 vs 200 MA)
        if len(closes) >= 200:
            ma50  = closes.rolling(50).mean()
            ma200 = closes.rolling(200).mean()
            if ma50.iloc[-2] < ma200.iloc[-2] and ma50.iloc[-1] > ma200.iloc[-1]:
                signals.append("Golden Cross")
            elif ma50.iloc[-2] > ma200.iloc[-2] and ma50.iloc[-1] < ma200.iloc[-1]:
                signals.append("Death Cross")

        # Near 52-week high (within 1%)
        high52 = closes.tail(252).max()
        if closes.iloc[-1] >= high52 * 0.99:
            signals.append("At 52W High")

        return " | ".join(signals)
    except Exception:
        return ""


@st.cache_data(ttl=300, show_spinner=False)
def get_technical_signals_batch(tickers: tuple) -> dict:
    """Get technical signals for multiple tickers in parallel. Cached 5 minutes."""
    raw = _parallel_map(lambda t: get_technical_signal(t), tickers)
    return {t: (v or "") for t, v in raw.items()}


# ── Market mover helpers ──────────────────────────────────────────────────────

@st.cache_data(ttl=120, show_spinner=False)
def get_top_movers(tickers: tuple, n: int = 10) -> pd.DataFrame:
    """
    Return top N gainers and losers from the given universe.
    Cached 2 minutes.
    """
    if not tickers:
        return pd.DataFrame()
    quotes = get_quotes(tickers)
    rows = [
        {
            "Ticker":    t,
            "Price":     q["current"],
            "Change $":  q["day_change"],
            "Change %":  q["day_change_pct"],
        }
        for t, q in quotes.items()
        if q.get("current") is not None and q.get("day_change_pct") is not None
    ]
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).sort_values("Change %", ascending=False)
    return df


@st.cache_data(ttl=120, show_spinner=False)
def get_top_premarket_movers(tickers: tuple, n: int = 10) -> pd.DataFrame:
    """
    Return top N pre-market movers from the given universe.
    Cached 2 minutes.
    """
    if not tickers:
        return pd.DataFrame()
    pq = get_premarket_quotes(tickers)
    rows = [
        {
            "Ticker":      t,
            "Pre Price":   v["pre_price"],
            "Pre Chg %":   v["pre_change_pct"],
        }
        for t, v in pq.items()
        if v.get("pre_change_pct") is not None
    ]
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).sort_values("Pre Chg %", ascending=False)
    return df
