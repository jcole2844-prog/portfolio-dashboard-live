"""
app.py — Portfolio Dashboard (Streamlit)
Refreshes automatically every 60 seconds.
"""

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from datetime import datetime, date, time as dtime, timedelta
from zoneinfo import ZoneInfo

from data_feed import (
    EXCEL_PATH, DOW_30,
    load_portfolio,
    load_fixed_income,
    get_watchlist,
    write_sheet_df,
    _read_sheet_df,
    _sheet_enabled,
    append_change_log,
    load_change_log,
    get_quotes,
    get_premarket_quotes,
    get_ext_hours_prices,
    get_fundamentals,
    get_analyst_actions,
    get_recommendations,
    get_technical_signals_batch,
    get_top_movers,
    get_top_premarket_movers,
    get_sp500_tickers,
    get_index_quotes,
    get_market_extras,
    get_earnings_dates,
)

# ─────────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Portfolio Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Page routing state (needed before deciding whether to auto-refresh)
if "page" not in st.session_state:
    st.session_state.page = "main"

# Auto-refresh every 30 seconds on every page. The ONLY time it pauses is while
# an "Edit mode" panel is open, so an in-progress edit can't be interrupted.
_editing = any(
    st.session_state.get(k, False)
    for k in ("hold_edit_mode", "wl_edit_mode", "fi_edit_mode")
)
if not _editing:
    st_autorefresh(interval=60_000, key="autorefresh")

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Tighten metric cards */
[data-testid="metric-container"] {
    background: #1a1a2e;
    border: 1px solid #2d2d4e;
    border-radius: 8px;
    padding: 8px 14px 8px 14px;
}
[data-testid="stMetricValue"] { font-size: 1.1rem; }
[data-testid="stMetricLabel"] { font-size: 0.75rem; color: #aaa; }

/* Scrolling ticker banners */
.ticker-bar { display:flex; align-items:center; background:#1a1a2e;
    border:1px solid #2d2d4e; border-radius:8px; overflow:hidden;
    margin-bottom:6px; }
.ticker-label { flex:0 0 auto; padding:6px 12px; font-size:0.72rem;
    font-weight:700; letter-spacing:0.5px; color:#0b0b14; white-space:nowrap; }
.ticker-track { flex:1 1 auto; overflow:hidden; }
.ticker-move { display:inline-block; white-space:nowrap; padding-left:100%;
    animation: ticker-scroll 40s linear infinite; font-size:0.9rem;
    font-weight:600; }
.ticker-move:hover { animation-play-state: paused; }
@keyframes ticker-scroll { 0% { transform: translateX(0); }
    100% { transform: translateX(-100%); } }

/* Flag columns */
.flag-red  { color: #ff4b4b; font-weight: bold; font-size: 1.1rem; }
.flag-green { color: #00d488; font-weight: bold; font-size: 1.1rem; }

/* Subheader spacing */
h3 { margin-top: 0.4rem !important; margin-bottom: 0.2rem !important; }

/* Slightly smaller dataframe text */
[data-testid="stDataFrame"] { font-size: 0.82rem; }

/* Blue Watchlist button (targets the widget by its key) */
.st-key-watchlist_btn button {
    background-color: #1f6feb !important;
    border-color: #1f6feb !important;
    color: #ffffff !important;
}
.st-key-watchlist_btn button:hover {
    background-color: #1a5fd0 !important;
    border-color: #1a5fd0 !important;
}

/* Teal Fixed Income button */
.st-key-fixedincome_btn button {
    background-color: #0f9d8f !important;
    border-color: #0f9d8f !important;
    color: #ffffff !important;
}
.st-key-fixedincome_btn button:hover {
    background-color: #0c8377 !important;
    border-color: #0c8377 !important;
}

/* Purple Change Log button */
.st-key-changelog_btn button {
    background-color: #7c5cff !important;
    border-color: #7c5cff !important;
    color: #ffffff !important;
}
.st-key-changelog_btn button:hover {
    background-color: #6a4ae0 !important;
    border-color: #6a4ae0 !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Formatting helpers
# ─────────────────────────────────────────────────────────────────────────────

def _na(v):
    return v is None or (isinstance(v, float) and np.isnan(v))

def _now_ct():
    """(date_str, time_str) in Central Time for change-log entries."""
    from zoneinfo import ZoneInfo
    now = datetime.now(ZoneInfo("America/Chicago"))
    return now.strftime("%m/%d/%Y"), now.strftime("%I:%M:%S %p")

def _num(v):
    """Coerce to float, treating blank/None as 0."""
    try:
        f = float(v)
        return 0.0 if pd.isna(f) else f
    except (TypeError, ValueError):
        return 0.0

def diff_positions(old_map, new_map, source):
    """Compare {key: (qty, cost, name)} dicts and return change-log entries for
    any added / removed / changed positions. Rounds to avoid float noise."""
    d, t = _now_ct()
    entries = []
    for key in sorted(set(old_map) | set(new_map)):
        oq, oc, on = old_map.get(key, (0.0, 0.0, ""))
        nq, nc, nn = new_map.get(key, (0.0, 0.0, on))
        dq, dc = round(nq - oq, 4), round(nc - oc, 2)
        if dq == 0 and dc == 0:
            continue
        entries.append({
            "date": d, "time": t, "source": source,
            "ticker": key, "name": nn or on,
            "qty_change": dq, "cost_change": dc,
        })
    return entries

def near_52w_low(tickers, threshold=10.0):
    """Return [{ticker, price, pct}] for tickers whose price is within
    `threshold`% above their 52-week low. Price = extended-hours price when the
    market is in pre/post-market, otherwise the current/last price. Closest to
    the low first."""
    q = get_quotes(tuple(tickers))
    f = get_fundamentals(tuple(tickers))
    ext = get_ext_hours_prices(tuple(tickers))
    out = []
    for t in tickers:
        cur = q.get(t, {}).get("current")
        ev  = ext.get(t)
        price = ev if ev is not None else cur   # ext-hours price when available
        lo  = f.get(t, {}).get("low_52w")
        if price and lo and lo > 0:
            pct = (price - lo) / lo * 100
            if 0 <= pct <= threshold:
                out.append({"ticker": t, "price": price, "pct": pct})
    out.sort(key=lambda x: x["pct"])
    return out

def near_52w_high(tickers, threshold=10.0):
    """Return [{ticker, price, pct}] for tickers whose price is within
    `threshold`% below their 52-week high. Price = extended-hours price when the
    market is in pre/post-market, otherwise the current/last price. Closest to
    the high first."""
    q = get_quotes(tuple(tickers))
    f = get_fundamentals(tuple(tickers))
    ext = get_ext_hours_prices(tuple(tickers))
    out = []
    for t in tickers:
        cur = q.get(t, {}).get("current")
        ev  = ext.get(t)
        price = ev if ev is not None else cur
        hi  = f.get(t, {}).get("high_52w")
        if price and hi and hi > 0:
            pct_below = (hi - price) / hi * 100
            if 0 <= pct_below <= threshold:
                out.append({"ticker": t, "price": price, "pct": pct_below})
    out.sort(key=lambda x: x["pct"])
    return out

def upcoming_earnings(tickers, business_days=10):
    """[{ticker, date}] for tickers with earnings from today through
    `business_days` business days ahead (inclusive). Soonest first."""
    from datetime import timedelta
    today = datetime.now(ZoneInfo("America/New_York")).date()
    end, added = today, 0
    while added < business_days:
        end += timedelta(days=1)
        if end.weekday() < 5:
            added += 1
    ed = get_earnings_dates(tuple(tickers))
    items = [{"ticker": t, "date": d} for t, d in ed.items()
             if d is not None and today <= d <= end]
    items.sort(key=lambda x: x["date"])
    return items

def earnings_banner(label, items, label_bg="#1f9bff"):
    """Scrolling marquee banner of upcoming earnings (Ticker + date)."""
    if items:
        parts = [f'<span style="color:#fff;">{it["ticker"]}</span> '
                 f'<span style="color:#9fd0ff;">{it["date"].strftime("%m/%d")}</span>'
                 for it in items]
        content = ' &nbsp;&nbsp;•&nbsp;&nbsp; '.join(parts)
    else:
        content = '<span style="color:#888;">no earnings in the next 10 business days</span>'
    st.markdown(
        f'<div class="ticker-bar">'
        f'<div class="ticker-label" style="background:{label_bg};">{label}</div>'
        f'<div class="ticker-track"><span class="ticker-move">{content}</span></div>'
        f'</div>',
        unsafe_allow_html=True,
    )

def ticker_banner(label, items, label_bg="#ff4b4b"):
    """Render a scrolling marquee banner (Ticker + price only)."""
    if items:
        parts = [f'<span style="color:#fff;">{it["ticker"]}</span> '
                 f'<span style="color:#ffd166;">{it["price"]:,.2f}</span>'
                 for it in items]
        content = ' &nbsp;&nbsp;•&nbsp;&nbsp; '.join(parts)
    else:
        content = '<span style="color:#888;">none within 10% of the 52-week low</span>'
    st.markdown(
        f'<div class="ticker-bar">'
        f'<div class="ticker-label" style="background:{label_bg};">{label}</div>'
        f'<div class="ticker-track"><span class="ticker-move">{content}</span></div>'
        f'</div>',
        unsafe_allow_html=True,
    )

# CNBC quote pages — every ticker becomes a clickable link.
CNBC_QUOTE = "https://www.cnbc.com/quotes/{}"

def cnbc(t):
    return CNBC_QUOTE.format(str(t).strip().upper())

def ticker_link_col():
    """A column_config that renders the Ticker column as a CNBC link,
    displaying just the symbol (extracted from the URL)."""
    return st.column_config.LinkColumn("Ticker", display_text=r"quotes/([^/]+)$")

def cap_size(mc):
    """Classify a market cap into Mega/Large/Mid/Small/Micro."""
    if _na(mc) or not mc:
        return ""
    if mc >= 200e9:
        return "Mega"
    if mc >= 10e9:
        return "Large"
    if mc >= 2e9:
        return "Mid"
    if mc >= 300e6:
        return "Small"
    return "Micro"

_DOLLAR_COLS = (
    "Price", "Pre Price", "Change $", "Last Close", "Last Closing Price",
    "Open", "Current Price",
)
_PCT_COLS = ("Chg %", "Change %", "Pre Chg %", "Current Change %")

def mover_num_config(df):
    """Build column_config so price/percent columns are numeric (right-aligned)."""
    cfg = {}
    for c in df.columns:
        if c in _DOLLAR_COLS:
            cfg[c] = st.column_config.NumberColumn(format="%,.2f")
        elif c in _PCT_COLS:
            cfg[c] = st.column_config.NumberColumn(format="%+.2f%%")
        elif c == "Fwd P/E":
            cfg[c] = st.column_config.NumberColumn(format="%.1f")
    return cfg

def add_name_col(df):
    """Insert a company-name column immediately after the Ticker column."""
    if df.empty or "Ticker" not in df.columns:
        return df
    df = df.copy()
    f = get_fundamentals(tuple(sorted(set(df["Ticker"]))))
    names = df["Ticker"].map(lambda t: (f.get(t, {}).get("name") or t))
    df.insert(df.columns.get_loc("Ticker") + 1, "Name", names)
    return df

def build_ohlc_movers(tickers, n=10):
    """Top n movers with Last Close, Open, Change $/%, Current Price, Current Change %.
    Change $/% = Open vs prior Close; Current Change % = Current vs Open.
    Ranked by the full-day move (Current vs prior Close)."""
    q = get_quotes(tuple(tickers))
    rows = []
    for t in tickers:
        d = q.get(t, {})
        cur, prev, opn = d.get("current"), d.get("prev_close"), d.get("open")
        if cur is None or prev is None:
            continue
        chg_d = (opn - prev) if opn is not None else None
        chg_p = (chg_d / prev * 100) if (chg_d is not None and prev) else None
        cur_p = ((cur - opn) / opn * 100) if opn else None
        rows.append({
            "Ticker":           t,
            "Last Close":       prev,
            "Open":             opn,
            "Change $":         chg_d,
            "Change %":         chg_p,
            "Current Price":    cur,
            "Current Change %": cur_p,
            "_sort":            (cur - prev) / prev * 100 if prev else 0,
        })
    if not rows:
        return pd.DataFrame()
    return (
        pd.DataFrame(rows)
        .sort_values("_sort", ascending=False)
        .head(n)
        .drop(columns="_sort")
        .reset_index(drop=True)
    )

def build_premarket(tickers, n=10):
    """Top n pre-market movers: Ticker, Last Closing Price, Pre Price, Pre Chg %."""
    pq = get_premarket_quotes(tuple(tickers))
    q  = get_quotes(tuple(tickers))
    rows = []
    for t in tickers:
        v = pq.get(t, {})
        if v.get("pre_change_pct") is None:
            continue
        rows.append({
            "Ticker":             t,
            "Last Closing Price": q.get(t, {}).get("prev_close"),
            "Pre Price":          v.get("pre_price"),
            "Pre Chg %":          v.get("pre_change_pct"),
            "_sort":              v.get("pre_change_pct") or 0,
        })
    if not rows:
        return pd.DataFrame()
    return (
        pd.DataFrame(rows)
        .sort_values("_sort", ascending=False)
        .head(n)
        .drop(columns="_sort")
        .reset_index(drop=True)
    )

def show_ticker_df(df, extra_config=None, **kwargs):
    """Render a dataframe with the Ticker column linked to CNBC quote pages."""
    df = df.copy()
    cfg = dict(extra_config) if extra_config else {}
    if "Ticker" in df.columns:
        df["Ticker"] = df["Ticker"].map(cnbc)
        cfg["Ticker"] = ticker_link_col()
    st.dataframe(df, column_config=cfg, **kwargs)

def fc(v, decimals=2):
    """Format as currency, e.g. $1,234.56"""
    if _na(v):
        return "—"
    return f"${v:,.{decimals}f}"

def fp(v, decimals=2):
    """Format as signed percentage, e.g. +3.45%"""
    if _na(v):
        return "—"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.{decimals}f}%"

def fn(v, decimals=2):
    """Format plain number"""
    if _na(v):
        return "—"
    return f"{v:,.{decimals}f}"

_ET = ZoneInfo("America/New_York")
_CT = ZoneInfo("America/Chicago")
_MKT_OPEN        = dtime(9, 30)
_MKT_CLOSE       = dtime(16, 0)
_MKT_EARLY_CLOSE = dtime(13, 0)   # 1:00 PM ET on half-days

# NYSE full-day market holidays (market closed). These follow the exchange
# calendar, which differs from federal holidays (markets stay open on Columbus
# Day and Veterans Day). Observed dates shown when a holiday falls on a weekend.
_MARKET_HOLIDAYS = {
    # 2025
    date(2025, 1, 1),  date(2025, 1, 20), date(2025, 2, 17), date(2025, 4, 18),
    date(2025, 5, 26), date(2025, 6, 19), date(2025, 7, 4),  date(2025, 9, 1),
    date(2025, 11, 27), date(2025, 12, 25),
    # 2026
    date(2026, 1, 1),  date(2026, 1, 19), date(2026, 2, 16), date(2026, 4, 3),
    date(2026, 5, 25), date(2026, 6, 19), date(2026, 7, 3),  date(2026, 9, 7),
    date(2026, 11, 26), date(2026, 12, 25),
    # 2027
    date(2027, 1, 1),  date(2027, 1, 18), date(2027, 2, 15), date(2027, 3, 26),
    date(2027, 5, 31), date(2027, 6, 18), date(2027, 7, 5),  date(2027, 9, 6),
    date(2027, 11, 25), date(2027, 12, 24),
    # 2028  (New Year's Day falls on Saturday → not observed by NYSE)
    date(2028, 1, 17), date(2028, 2, 21), date(2028, 4, 14), date(2028, 5, 29),
    date(2028, 6, 19), date(2028, 7, 4),  date(2028, 9, 4),  date(2028, 11, 23),
    date(2028, 12, 25),
}

# NYSE early-close days — market closes at 1:00 PM ET (the day before
# Independence Day when it's a weekday, Black Friday, and Christmas Eve).
_HALF_DAYS = {
    # 2025
    date(2025, 7, 3),  date(2025, 11, 28), date(2025, 12, 24),
    # 2026  (no July half-day: July 3 is a full holiday)
    date(2026, 11, 27), date(2026, 12, 24),
    # 2027  (no July half-day: July 5 observed; Dec 24 is a full holiday)
    date(2027, 11, 26),
    # 2028  (Dec 24 falls on Sunday → no Christmas Eve half-day)
    date(2028, 7, 3),  date(2028, 11, 24),
}

def _is_trading_day(d):
    """True if d is a weekday and not a NYSE market holiday."""
    return d.weekday() < 5 and d not in _MARKET_HOLIDAYS

def _next_market_open(now_et):
    """Datetime of the next 9:30 AM ET market open from now (skips weekends
    and NYSE holidays)."""
    candidate = now_et.replace(hour=_MKT_OPEN.hour, minute=_MKT_OPEN.minute,
                               second=0, microsecond=0)
    if now_et >= candidate:                 # today's open already passed
        candidate += timedelta(days=1)
    while not _is_trading_day(candidate.date()):   # skip weekends + holidays
        candidate += timedelta(days=1)
    return candidate


def market_clock():
    """Return (status, countdown_hhmm, et_str, ct_str).

    status 'open'  → countdown is hours:minutes until the 4:00 PM ET close.
    status 'closed'→ countdown is hours:minutes until the next 9:30 AM ET open.
    """
    now_et = datetime.now(_ET)
    now_ct = datetime.now(_CT)
    et_str = now_et.strftime("%I:%M %p %Z")   # %Z → EDT/EST automatically
    ct_str = now_ct.strftime("%I:%M %p %Z")   # %Z → CDT/CST automatically

    is_trading = _is_trading_day(now_et.date())
    close_t = _MKT_EARLY_CLOSE if now_et.date() in _HALF_DAYS else _MKT_CLOSE
    open_dt  = now_et.replace(hour=_MKT_OPEN.hour, minute=_MKT_OPEN.minute, second=0, microsecond=0)
    close_dt = now_et.replace(hour=close_t.hour,   minute=close_t.minute,   second=0, microsecond=0)

    if is_trading and open_dt <= now_et < close_dt:
        remaining = close_dt - now_et
        status = "open"
    else:
        remaining = _next_market_open(now_et) - now_et
        status = "closed"

    total_min = int(remaining.total_seconds() // 60)
    hh, mm = divmod(total_min, 60)
    return status, f"{hh:02d}:{mm:02d}", et_str, ct_str


def _color_pct(v):
    """Return HTML-colored percentage string."""
    if _na(v):
        return "—"
    color = "#00d488" if v >= 0 else "#ff4b4b"
    return f'<span style="color:{color}">{fp(v)}</span>'

def _movers_display(df: pd.DataFrame, pct_col: str, price_col: str = "Price") -> pd.DataFrame:
    """Return the top 10 rows with numeric price/percent columns intact, so the
    grid right-aligns them (formatting is applied via mover_num_config)."""
    if df.empty:
        return df
    return df.copy().head(10)


# ─────────────────────────────────────────────────────────────────────────────
# ██████  HOLDINGS DETAIL PAGE
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.page == "holdings":

    col_back, col_title = st.columns([1, 9])
    with col_back:
        if st.button("← Dashboard"):
            st.session_state.page = "main"
            st.rerun()
    with col_title:
        st.title("All Holdings — Detail View")

    st.markdown(
        '<a href="https://www.cnbc.com/markets/sectors/" target="_blank" '
        'style="color:#1f9bff;text-decoration:none;font-weight:700;font-size:1.05rem;">'
        '📊 Sectors (CNBC) ↗</a>',
        unsafe_allow_html=True,
    )
    st.caption(f"Last updated: {datetime.now().strftime('%A %B %d, %Y  %I:%M:%S %p')}")

    # ── Edit holdings (writes to the Google Sheet) ──
    if _sheet_enabled():
        if st.checkbox("✏️  Edit holdings — add/remove tickers, update shares & cost basis "
                       "(pauses auto-refresh; saves to Google Sheet)", key="hold_edit_mode"):
            src = load_portfolio()[["Ticker", "Name", "Total Quantity", "Total Cost Basis"]].copy()
            edited = st.data_editor(
                src, num_rows="dynamic", width="stretch", hide_index=True,
                key="hold_editor",
                column_config={
                    "Ticker":           st.column_config.TextColumn("Ticker", required=True),
                    "Name":             st.column_config.TextColumn("Name"),
                    "Total Quantity":   st.column_config.NumberColumn("Total Quantity", min_value=0.0, format="%.4f"),
                    "Total Cost Basis": st.column_config.NumberColumn("Total Cost Basis", min_value=0.0, format="%.2f"),
                },
            )
            if st.button("💾  Save holdings to Google Sheet", key="save_holdings", type="primary"):
                out = edited.copy()
                out = out[out["Ticker"].astype(str).str.strip() != ""]
                out["Ticker"] = out["Ticker"].astype(str).str.strip().str.upper()
                qty  = pd.to_numeric(out["Total Quantity"], errors="coerce")
                cost = pd.to_numeric(out["Total Cost Basis"], errors="coerce")
                out["Avg Basis/Sh"] = (cost / qty).where(qty.ne(0))
                out = out[["Ticker", "Name", "Total Quantity", "Total Cost Basis", "Avg Basis/Sh"]]
                # Change log: compare pre-edit holdings to the new ones.
                _old = load_portfolio()
                old_map = {str(r["Ticker"]).upper(): (_num(r["Total Quantity"]),
                           _num(r["Total Cost Basis"]), str(r.get("Name", "")))
                           for _, r in _old.iterrows()}
                new_map = {str(r["Ticker"]).upper(): (_num(r["Total Quantity"]),
                           _num(r["Total Cost Basis"]), str(r.get("Name", "")))
                           for _, r in out.iterrows()}
                try:
                    write_sheet_df("Holdings", out)
                    append_change_log(diff_positions(old_map, new_map, "Holdings"))
                    load_portfolio.clear()
                    load_change_log.clear()
                    st.success("Holdings saved to Google Sheet ✓ — uncheck Edit to resume auto-refresh.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Save failed: {e}")

    portfolio = load_portfolio()
    tickers = tuple(portfolio.iloc[:, 0].tolist())

    with st.spinner("Fetching live data for all holdings…"):
        quotes    = get_quotes(tickers)
        funds     = get_fundamentals(tickers)
        recs      = get_recommendations(tickers)
        signals   = get_technical_signals_batch(tickers)
        ext_hours = get_ext_hours_prices(tickers)

    rows = []
    for _, row in portfolio.iterrows():
        t        = str(row.iloc[0]).strip().upper()
        name     = str(row.iloc[1]) if len(row) > 1 else t
        qty      = float(row.iloc[2]) if len(row) > 2 else 0
        tot_cost = float(row.iloc[3]) if len(row) > 3 else 0
        avg_bas  = float(row.iloc[4]) if len(row) > 4 else (tot_cost / qty if qty else 0)

        q  = quotes.get(t, {})
        f  = funds.get(t, {})
        rc = recs.get(t, {})

        cur   = q.get("current")
        cur_val  = cur * qty if cur else None
        day_chg  = (q.get("day_change") or 0) * qty if cur else None
        day_pct  = q.get("day_change_pct")
        unr      = (cur_val - tot_cost) if cur_val is not None else None
        unr_pct  = (unr / tot_cost * 100) if (unr is not None and tot_cost) else None

        opinion  = rc.get("opinion", "—")
        buy_pct  = rc.get("buy_pct",  0)
        sell_pct = rc.get("sell_pct", 0)

        red_flag   = "🚩" if sell_pct >= 60 else ""
        green_flag = "✅" if buy_pct  >= 70 else ""

        ext_price = ext_hours.get(t)
        ext_chg   = (ext_price - cur) if (ext_price is not None and cur is not None) else None

        # Keep raw numeric values so the grid's column headers sort numerically.
        # Display formatting is applied via column_config below.
        rows.append({
            "Ticker":       t,
            "Name":         f.get("name", name),
            "Ext. Hrs Price": ext_price,
            "Ext. Hrs Chg $": ext_chg,
            "Sector":       f.get("sector") or "",
            "Shares":       qty,
            "Avg Basis":    avg_bas,
            "Cur Price":    cur,
            "Cur Value":    cur_val,
            "P/E":          f.get("pe"),
            "Fwd P/E":      f.get("forward_pe"),
            "Div %":        f.get("div_yield") or None,
            "Unr Gain $":   unr,
            "Unr Gain %":   unr_pct,
            "Day Gain $":   day_chg,
            "Day Gain %":   day_pct,
            "Day High":     f.get("day_high"),
            "Day Low":      f.get("day_low"),
            "52W High":     f.get("high_52w"),
            "52W Low":      f.get("low_52w"),
            "Tech Signal":  signals.get(t, ""),
            "Analyst":      opinion,
            "🚩 Sell":      red_flag,
            "✅ Buy":       green_flag,
        })

    detail_df = pd.DataFrame(rows)

    # Keep numeric columns as real numbers so the column-header sort works
    # numerically (the formatted-string values sorted alphabetically before).
    _numeric_cols = [
        "Ext. Hrs Price", "Ext. Hrs Chg $", "Shares", "Avg Basis", "Cur Price",
        "Cur Value", "P/E", "Fwd P/E", "Div %", "Unr Gain $", "Unr Gain %",
        "Day Gain $", "Day Gain %", "Day High", "Day Low", "52W High", "52W Low",
    ]
    for _c in _numeric_cols:
        detail_df[_c] = pd.to_numeric(detail_df[_c], errors="coerce")

    # Display formatting via column_config — values stay numeric so the column
    # header sort works numerically, and missing values (NaN) render blank.
    detail_config = {
        "Ext. Hrs Price": st.column_config.NumberColumn(format="%,.2f"),
        "Ext. Hrs Chg $": st.column_config.NumberColumn(format="%,.2f"),
        "Shares":     st.column_config.NumberColumn(format="%.2f"),
        "Avg Basis":  st.column_config.NumberColumn(format="%,.2f"),
        "Cur Price":  st.column_config.NumberColumn(format="%,.2f"),
        "Cur Value":  st.column_config.NumberColumn(format="%,.2f"),
        "P/E":        st.column_config.NumberColumn(format="%.1f"),
        "Fwd P/E":    st.column_config.NumberColumn(format="%.1f"),
        "Div %":      st.column_config.NumberColumn(format="%.2f%%"),
        "Unr Gain $": st.column_config.NumberColumn(format="%,.2f"),
        "Unr Gain %": st.column_config.NumberColumn(format="%.2f%%"),
        "Day Gain $": st.column_config.NumberColumn(format="%,.2f"),
        "Day Gain %": st.column_config.NumberColumn(format="%.2f%%"),
        "Day High":   st.column_config.NumberColumn(format="%,.2f"),
        "Day Low":    st.column_config.NumberColumn(format="%,.2f"),
        "52W High":   st.column_config.NumberColumn(format="%,.2f"),
        "52W Low":    st.column_config.NumberColumn(format="%,.2f"),
    }

    # Link tickers to CNBC and color the gain/loss columns by sign.
    detail_df["Ticker"] = detail_df["Ticker"].map(cnbc)
    detail_config["Ticker"] = ticker_link_col()

    def _sign_color(v):
        if pd.isna(v):
            return ""
        return "color: #00d488" if v >= 0 else "color: #ff4b4b"

    gain_cols = ["Ext. Hrs Chg $", "Unr Gain $", "Unr Gain %", "Day Gain $", "Day Gain %"]
    styled_detail = detail_df.style.map(_sign_color, subset=gain_cols)

    st.dataframe(
        styled_detail,
        column_config=detail_config,
        width="stretch",
        height=620,
        hide_index=True,
    )
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# ██████  WATCHLIST PAGE
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.page == "watchlist":

    col_back, col_title = st.columns([1, 9])
    with col_back:
        if st.button("← Dashboard"):
            st.session_state.page = "main"
            st.rerun()
    with col_title:
        st.title("👁 Watchlist — Detail View")

    st.caption(f"Last updated: {datetime.now().strftime('%A %B %d, %Y  %I:%M:%S %p')}")

    # ── Edit watchlist (writes to the Google Sheet) ──
    if _sheet_enabled():
        if st.checkbox("✏️  Edit watchlist — add or remove tickers "
                       "(pauses auto-refresh; saves to Google Sheet)", key="wl_edit_mode"):
            wsrc = _read_sheet_df("Watchlist")
            if "Ticker" not in wsrc.columns:
                wsrc = pd.DataFrame({"Ticker": [], "Name": []})
            if "Name" not in wsrc.columns:
                wsrc["Name"] = ""
            wsrc = wsrc[["Ticker", "Name"]]
            wedited = st.data_editor(
                wsrc, num_rows="dynamic", width="stretch", hide_index=True,
                key="wl_editor",
                column_config={
                    "Ticker": st.column_config.TextColumn("Ticker", required=True,
                              help="Use Yahoo format, e.g. BRK-B not BRK.B"),
                    "Name":   st.column_config.TextColumn("Name (optional)"),
                },
            )
            if st.button("💾  Save watchlist to Google Sheet", key="save_watchlist", type="primary"):
                out = wedited.copy()
                out = out[out["Ticker"].astype(str).str.strip() != ""]
                out["Ticker"] = out["Ticker"].astype(str).str.strip().str.upper().str.replace(".", "-", regex=False)
                out = out[["Ticker", "Name"]]
                try:
                    write_sheet_df("Watchlist", out)
                    get_watchlist.clear()
                    st.success("Watchlist saved to Google Sheet ✓ — uncheck Edit to resume auto-refresh.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Save failed: {e}")

    wl_tickers = get_watchlist()

    with st.spinner("Fetching live data for watchlist…"):
        wl_quotes  = get_quotes(wl_tickers)
        wl_funds   = get_fundamentals(wl_tickers)
        wl_recs    = get_recommendations(wl_tickers)
        wl_signals = get_technical_signals_batch(wl_tickers)
        wl_ext     = get_ext_hours_prices(wl_tickers)

    wl_rows = []
    for t in wl_tickers:
        q  = wl_quotes.get(t, {})
        f  = wl_funds.get(t, {})
        rc = wl_recs.get(t, {})

        buy_pct  = rc.get("buy_pct",  0)
        sell_pct = rc.get("sell_pct", 0)

        wl_ext_price = wl_ext.get(t)
        wl_cur       = q.get("current")
        wl_ext_chg   = (wl_ext_price - wl_cur) if (wl_ext_price is not None and wl_cur is not None) else None

        wl_hi = f.get("high_52w")
        wl_lo = f.get("low_52w")
        pct_from_high = ((wl_cur - wl_hi) / wl_hi * 100) if (wl_cur and wl_hi) else None
        pct_from_low  = ((wl_cur - wl_lo) / wl_lo * 100) if (wl_cur and wl_lo) else None

        wl_rows.append({
            "Ticker":      t,
            "Name":        f.get("name", t),
            "Ext. Hrs Price": wl_ext_price,
            "Ext. Hrs Chg $": wl_ext_chg,
            "Sector":      f.get("sector") or "",
            "Size":        cap_size(f.get("market_cap")),
            "Cur Price":   q.get("current"),
            "% from 52W High": pct_from_high,
            "% from 52W Low":  pct_from_low,
            "P/E":         f.get("pe"),
            "Fwd P/E":     f.get("forward_pe"),
            "Div %":       f.get("div_yield") or None,
            "Day High":    f.get("day_high"),
            "Day Low":     f.get("day_low"),
            "52W High":    f.get("high_52w"),
            "52W Low":     f.get("low_52w"),
            "Tech Signal": wl_signals.get(t, ""),
            "Analyst":     rc.get("opinion", "—"),
            "🚩 Sell":     "🚩" if sell_pct >= 60 else "",
            "✅ Buy":      "✅" if buy_pct  >= 70 else "",
        })

    wl_df = pd.DataFrame(wl_rows)

    _wl_numeric = [
        "Ext. Hrs Price", "Ext. Hrs Chg $", "Cur Price",
        "% from 52W High", "% from 52W Low", "P/E", "Fwd P/E", "Div %",
        "Day High", "Day Low", "52W High", "52W Low",
    ]
    for _c in _wl_numeric:
        wl_df[_c] = pd.to_numeric(wl_df[_c], errors="coerce")

    wl_config = {
        "Ext. Hrs Price": st.column_config.NumberColumn(format="%,.2f"),
        "Ext. Hrs Chg $": st.column_config.NumberColumn(format="%,.2f"),
        "Cur Price": st.column_config.NumberColumn(format="%,.2f"),
        "% from 52W High": st.column_config.NumberColumn(format="%+.2f%%"),
        "% from 52W Low":  st.column_config.NumberColumn(format="%+.2f%%"),
        "P/E":       st.column_config.NumberColumn(format="%.1f"),
        "Fwd P/E":   st.column_config.NumberColumn(format="%.1f"),
        "Div %":     st.column_config.NumberColumn(format="%.2f%%"),
        "Day High":  st.column_config.NumberColumn(format="%,.2f"),
        "Day Low":   st.column_config.NumberColumn(format="%,.2f"),
        "52W High":  st.column_config.NumberColumn(format="%,.2f"),
        "52W Low":   st.column_config.NumberColumn(format="%,.2f"),
    }

    # Link tickers to CNBC and color the extended-hours change by sign.
    wl_df["Ticker"] = wl_df["Ticker"].map(cnbc)
    wl_config["Ticker"] = ticker_link_col()

    def _wl_sign_color(v):
        if pd.isna(v):
            return ""
        return "color: #00d488" if v >= 0 else "color: #ff4b4b"

    wl_styled = wl_df.style.map(_wl_sign_color, subset=["Ext. Hrs Chg $"])

    st.dataframe(
        wl_styled,
        column_config=wl_config,
        width="stretch",
        height=620,
        hide_index=True,
    )
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# ██████  FIXED INCOME PAGE
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.page == "fixedincome":

    col_back, col_title = st.columns([1, 9])
    with col_back:
        if st.button("← Dashboard"):
            st.session_state.page = "main"
            st.rerun()
    with col_title:
        st.title("🏦 Fixed Income — Detail View")

    st.caption(f"Last updated: {datetime.now().strftime('%A %B %d, %Y  %I:%M:%S %p')}")

    # ── Edit fixed income (writes to the Google Sheet) ──
    if _sheet_enabled():
        if st.checkbox("✏️  Edit fixed income — add/remove securities, edit details "
                       "(pauses auto-refresh; saves to Google Sheet)", key="fi_edit_mode"):
            fsrc = load_fixed_income().copy()
            fsrc["Coupon"] = pd.to_numeric(fsrc["Coupon"], errors="coerce") * 100
            fsrc["YTM"]    = pd.to_numeric(fsrc["YTM"],    errors="coerce") * 100
            fsrc["Acquisition Date"] = pd.to_datetime(fsrc["Acquisition Date"], errors="coerce").dt.date
            fsrc["Maturity Date"]    = pd.to_datetime(fsrc["Maturity Date"], errors="coerce").dt.date
            fsrc = fsrc[["Symbol", "Description", "Type", "Quantity",
                         "Acquisition Date", "Maturity Date", "Coupon", "YTM"]]
            fedited = st.data_editor(
                fsrc, num_rows="dynamic", width="stretch", hide_index=True,
                key="fi_editor",
                column_config={
                    "Symbol":           st.column_config.TextColumn("Symbol", required=True),
                    "Description":      st.column_config.TextColumn("Description"),
                    "Type":             st.column_config.TextColumn("Type"),
                    "Quantity":         st.column_config.NumberColumn("Quantity", min_value=0.0, format="%.2f"),
                    "Acquisition Date": st.column_config.DateColumn("Acquisition Date", format="MM/DD/YYYY"),
                    "Maturity Date":    st.column_config.DateColumn("Maturity Date", format="MM/DD/YYYY"),
                    "Coupon":           st.column_config.NumberColumn("Coupon %", format="%.3f"),
                    "YTM":              st.column_config.NumberColumn("YTM %", format="%.3f"),
                },
            )
            if st.button("💾  Save fixed income to Google Sheet", key="save_fi", type="primary"):
                out = fedited.copy()
                out = out[out["Symbol"].astype(str).str.strip() != ""]

                def _pct_str(v):
                    return "" if pd.isna(v) else f"{float(v):.3f}%"

                def _date_str(v):
                    if v is None or pd.isna(v):
                        return ""
                    d = pd.to_datetime(v)
                    return f"{d.month}/{d.day}/{d.year}"

                # Change log: compare pre-edit fixed income (Quantity = principal;
                # used for both Qty and Cost Basis change) before reformatting.
                _fi_old = load_fixed_income()
                old_map = {str(r["Symbol"]).upper(): (_num(r["Quantity"]),
                           _num(r["Quantity"]), str(r.get("Description", "")))
                           for _, r in _fi_old.iterrows()}
                _clean = out[out["Symbol"].astype(str).str.strip() != ""]
                new_map = {str(r["Symbol"]).upper(): (_num(r["Quantity"]),
                           _num(r["Quantity"]), str(r.get("Description", "")))
                           for _, r in _clean.iterrows()}

                out["Coupon"] = out["Coupon"].map(_pct_str)
                out["YTM"]    = out["YTM"].map(_pct_str)
                out["Acquisition Date"] = out["Acquisition Date"].map(_date_str)
                out["Maturity Date"]    = out["Maturity Date"].map(_date_str)
                out = out[["Symbol", "Description", "Type", "Quantity",
                           "Acquisition Date", "Maturity Date", "Coupon", "YTM"]]
                try:
                    write_sheet_df("Fixed Income", out)
                    append_change_log(diff_positions(old_map, new_map, "Fixed Income"))
                    load_fixed_income.clear()
                    load_change_log.clear()
                    st.success("Fixed income saved to Google Sheet ✓ — uncheck Edit to resume auto-refresh.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Save failed: {e}")

    fi = load_fixed_income().copy()
    today = pd.Timestamp(datetime.now().date())

    fi["Acquisition Date"] = pd.to_datetime(fi["Acquisition Date"])
    fi["Maturity Date"]    = pd.to_datetime(fi["Maturity Date"])

    # Estimate accrued interest: face value × coupon × (days held / 365),
    # accrued from acquisition up to today (capped at maturity).
    def _accrued(row):
        acq = row["Acquisition Date"]
        mat = row["Maturity Date"]
        end = min(today, mat)
        days = max((end - acq).days, 0)
        return row["Quantity"] * row["Coupon"] * days / 365.0

    fi["Accrued Interest"] = fi.apply(_accrued, axis=1)

    # ── Summary metrics ──
    total_principal = fi["Quantity"].sum()
    total_accrued   = fi["Accrued Interest"].sum()
    m1, m2, m3 = st.columns(3)
    m1.metric("Securities", len(fi))
    m2.metric("Total Principal", fc(total_principal))
    m3.metric("Total Interest Earned", fc(total_accrued))

    st.markdown("---")

    # ── Detail table (Excel columns + Accrued Interest) ──
    fi_display = fi.copy()
    fi_display["Acquisition Date"] = fi_display["Acquisition Date"].dt.date
    fi_display["Maturity Date"]    = fi_display["Maturity Date"].dt.date
    # Coupon & YTM stored as decimals → percent for display
    fi_display["Coupon"] = fi_display["Coupon"] * 100
    fi_display["YTM"]    = fi_display["YTM"] * 100

    fi_config = {
        "Quantity":         st.column_config.NumberColumn("Quantity", format="%,.2f"),
        "Coupon":           st.column_config.NumberColumn("Coupon",   format="%.3f%%"),
        "YTM":              st.column_config.NumberColumn("YTM",      format="%.3f%%"),
        "Accrued Interest": st.column_config.NumberColumn("Accrued Interest", format="%,.2f"),
        "Acquisition Date": st.column_config.DateColumn("Acquisition Date", format="MM/DD/YYYY"),
        "Maturity Date":    st.column_config.DateColumn("Maturity Date", format="MM/DD/YYYY"),
    }

    st.dataframe(
        fi_display,
        width="stretch",
        height=420,
        hide_index=True,
        column_config=fi_config,
    )

    st.markdown("---")

    # ── Principal maturing by month (bar chart) ──
    st.subheader("Principal Maturing by Month")

    fut = fi[fi["Maturity Date"] >= today].copy()
    if fut.empty:
        st.info("No securities maturing in the future.")
    else:
        fut["MatMonth"] = fut["Maturity Date"].dt.to_period("M").dt.to_timestamp()
        by_month = (
            fut.groupby("MatMonth", as_index=False)["Quantity"].sum()
            .sort_values("MatMonth")
        )
        by_month["Label"]    = by_month["MatMonth"].dt.strftime("%b %Y")
        by_month["AmtLabel"] = by_month["Quantity"].map(lambda v: f"${v:,.0f}")
        order = by_month["Label"].tolist()

        base = alt.Chart(by_month).encode(
            x=alt.X("Label:N", sort=order, title="Maturity Month",
                    axis=alt.Axis(labelAngle=-45)),
            y=alt.Y("Quantity:Q", title="Principal Maturing ($)"),
        )
        bars = base.mark_bar(color="#0f9d8f", size=28)
        text = base.mark_text(dy=-8, color="#ffffff", fontWeight="bold").encode(
            text=alt.Text("AmtLabel:N")
        )
        st.altair_chart((bars + text).properties(height=380), use_container_width=True)

    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# ██████  CHANGE LOG PAGE
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.page == "changelog":

    col_back, col_title = st.columns([1, 9])
    with col_back:
        if st.button("← Dashboard"):
            st.session_state.page = "main"
            st.rerun()
    with col_title:
        st.title("📝 Change Log")

    st.caption("Every edit you save to Holdings or Fixed Income is recorded here. "
               "Qty changes: green = added, red = reduced.")

    log = load_change_log()

    def _render_log(df, source):
        st.subheader(f"{source} Change Log")
        sub = df[df.get("Source", "") == source].copy() if not df.empty else df
        if sub is None or sub.empty:
            st.info(f"No {source.lower()} changes recorded yet.")
            return
        # Newest first (rows are appended chronologically).
        sub = sub.iloc[::-1].reset_index(drop=True)
        cols = ["Ticker", "Name", "Qty Change", "Cost Basis Change", "Date", "Time"]
        sub = sub[[c for c in cols if c in sub.columns]]
        sub["Qty Change"] = pd.to_numeric(sub["Qty Change"], errors="coerce")
        sub["Cost Basis Change"] = pd.to_numeric(sub["Cost Basis Change"], errors="coerce")

        def _sign_color(v):
            if pd.isna(v):
                return ""
            return "color: #00d488" if v >= 0 else "color: #ff4b4b"

        styled = sub.style.map(_sign_color, subset=["Qty Change"])
        st.dataframe(
            styled,
            width="stretch",
            hide_index=True,
            column_config={
                "Qty Change":        st.column_config.NumberColumn(format="%+,.2f"),
                "Cost Basis Change": st.column_config.NumberColumn(format="%+,.2f"),
            },
        )

    _render_log(log, "Holdings")
    st.markdown("---")
    _render_log(log, "Fixed Income")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# ██████  MAIN DASHBOARD PAGE
# ─────────────────────────────────────────────────────────────────────────────

# Header row
mkt_status, mkt_countdown, et_time, ct_time = market_clock()
if mkt_status == "open":
    _clock_color = "#ff4b4b"                 # red while open
    _clock_label = "Countdown to Market Close"
else:
    _clock_color = "#00d488"                 # green while closed
    _clock_label = "Countdown to Market Open"

def _mini_stack(title, items, suppress=False, suppress_text="Market Open"):
    """Compact stacked quote box (name left, value + ▲/▼ %chg right).
    When suppress=True, each row shows suppress_text instead of a value
    (used for futures while the regular market is open)."""
    rows = ""
    for it in items:
        v, p, url = it["value"], it["change_pct"], it.get("url")
        nm = it["name"]
        name_html = (f'<a href="{url}" target="_blank" style="color:#aaa;'
                     f'text-decoration:none;white-space:nowrap;">{nm}</a>' if url else
                     f'<span style="color:#aaa;white-space:nowrap;">{nm}</span>')
        if suppress:
            right = f'<span style="color:#777;white-space:nowrap;">{suppress_text}</span>'
        elif v is None or p is None:
            right = '<span style="color:#777;">—</span>'
        else:
            c = "#00d488" if p >= 0 else "#ff4b4b"
            ar = "▲" if p >= 0 else "▼"
            dec = 2 if abs(v) < 1000 else 0
            right = (f'<span style="color:{c};font-weight:600;white-space:nowrap;">'
                     f'{v:,.{dec}f} {ar}{abs(p):.2f}%</span>')
        rows += (f'<div style="display:flex;justify-content:space-between;gap:10px;'
                 f'font-size:0.74rem;line-height:1.65;">{name_html}{right}</div>')
    # Only the box title (FUTURES / COMMODITIES) is enlarged and yellow.
    return (f'<div style="background:#1a1a2e;border:1px solid #2d2d4e;border-radius:8px;'
            f'padding:6px 10px;">'
            f'<div style="font-size:0.95rem;font-weight:700;color:#ffd400;'
            f'letter-spacing:0.3px;margin-bottom:3px;white-space:nowrap;">{title}</div>{rows}</div>')


extras = get_market_extras()

# Row 1 — title (full width)
st.markdown(
    '<div style="font-size:2.6rem;font-weight:800;line-height:1.1;">'
    '📈 Portfolio Dashboard</div>',
    unsafe_allow_html=True,
)

# Row 2 — countdown clock | (spacer) | futures | commodities | CNBC + time zones
hdr_clock, _hdr_spacer, hdr_fut, hdr_com, hdr_r = st.columns([2.5, 1.4, 2.1, 2.4, 2.6])
with hdr_clock:
    st.markdown(
        f"""<div style="text-align:center;line-height:1.0;padding-top:2px;">
             <div style="font-size:3.0rem;font-weight:800;color:{_clock_color};
                  font-variant-numeric:tabular-nums;white-space:nowrap;">{mkt_countdown}</div>
             <div style="font-size:0.8rem;color:#aaa;">{_clock_label}</div></div>""",
        unsafe_allow_html=True,
    )
with hdr_fut:
    st.markdown(_mini_stack("FUTURES", extras["futures"],
                            suppress=(mkt_status == "open")), unsafe_allow_html=True)
with hdr_com:
    st.markdown(_mini_stack("COMMODITIES", extras["commodities"]), unsafe_allow_html=True)
with hdr_r:
    st.markdown(
        f"""<div style="text-align:right;line-height:1.3;">
             <div style="font-size:0.95rem;margin-bottom:3px;">
                 <a href="https://www.cnbc.com" target="_blank"
                    style="color:#1f9bff;text-decoration:none;font-weight:700;">📰 CNBC.com ↗</a></div>
             <div style="font-size:0.9rem;color:#ddd;"><b>ET</b> {et_time}</div>
             <div style="font-size:0.9rem;color:#ddd;"><b>CT</b> {ct_time}</div>
             <div style="font-size:0.72rem;color:#888;">Auto-refreshes every 60 sec</div></div>""",
        unsafe_allow_html=True,
    )

# ── Load portfolio ───────────────────────────────────────────────────────────
portfolio = load_portfolio()
port_tickers = tuple(portfolio.iloc[:, 0].tolist())

# ── Fetch portfolio data ─────────────────────────────────────────────────────
with st.spinner("Loading portfolio…"):
    quotes = get_quotes(port_tickers)
    funds  = get_fundamentals(port_tickers)
    recs   = get_recommendations(port_tickers)

# ── Compute portfolio totals ──────────────────────────────────────────────────
total_cost  = 0.0
total_value = 0.0
prev_value  = 0.0
port_rows   = []

missing_cost = []
for _, row in portfolio.iterrows():
    t        = str(row.iloc[0]).strip().upper()
    qty      = float(row.iloc[2]) if (len(row) > 2 and pd.notna(row.iloc[2])) else 0
    if len(row) > 3 and pd.notna(row.iloc[3]):
        tot_cost = float(row.iloc[3])
    else:
        tot_cost = 0                      # blank cost basis → 0, don't blank totals
        missing_cost.append(t)
    q        = quotes.get(t, {})
    cur      = q.get("current")
    prev     = q.get("prev_close")
    f        = funds.get(t, {})

    total_cost += tot_cost
    if cur  is not None:
        total_value += cur  * qty
    if prev is not None:
        prev_value  += prev * qty

    port_rows.append({
        "Ticker":    t,
        "Name":      f.get("name", str(row.iloc[1]) if len(row) > 1 else t)[:28],
        "Qty":       qty,
        "Price":     cur,
        "Chg %":     q.get("day_change_pct"),
        "_chg_sort": q.get("day_change_pct") or 0,
        "Fwd P/E":   f.get("forward_pe"),
        "Analyst":   recs.get(t, {}).get("opinion", "—"),
    })

total_day_gain     = total_value - prev_value
total_day_gain_pct = total_day_gain / prev_value * 100 if prev_value else 0
total_unr          = total_value - total_cost
total_unr_pct      = total_unr / total_cost * 100 if total_cost else 0

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — Portfolio Summary Metrics
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("Portfolio Summary")

if missing_cost:
    st.warning(
        "⚠️ Missing cost basis for: **" + ", ".join(missing_cost) + "** — "
        "Total Cost and Unrealized Gain exclude these. Add their cost basis in "
        "View Holdings → Edit (or the Google Sheet) to include them."
    )

# Fixed-income principal (face value, excluding accrued interest)
fixed_inc_total = float(load_fixed_income()["Quantity"].sum())

mc1, mc2, mc3, mc4, mc5, mc6 = st.columns(6)
mc1.metric("Holdings",      len(port_tickers))
mc2.metric("Total Value",   fc(total_value))
mc3.metric("Total Cost",    fc(total_cost))
mc4.metric("Unr Gain $",    fc(total_unr),  delta=fp(total_unr_pct))
mc5.metric("Today Gain $",  fc(total_day_gain), delta=fp(total_day_gain_pct))
mc6.metric("Fixed Inc",     fc(fixed_inc_total))

# ── Index performance bar (Dow / S&P 500 / Nasdaq), green up / red down ────────
indices = get_index_quotes()
idx_cols = st.columns(len(indices))
for col, idx in zip(idx_cols, indices):
    lvl = idx["level"]
    pct = idx["change_pct"]
    chg = idx["change"]
    url = idx["url"]
    if _na(lvl) or _na(pct):
        col.markdown(
            f"""<a href="{url}" target="_blank" style="text-decoration:none;">
                 <div style="background:#1a1a2e;border:1px solid #2d2d4e;border-radius:8px;
                 padding:10px 14px;text-align:center;cursor:pointer;">
                 <div style="font-size:0.8rem;color:#aaa;">{idx['name']} ↗</div>
                 <div style="font-size:1.2rem;font-weight:700;color:#888;">—</div>
                 <div style="font-size:0.85rem;color:#888;">market closed</div></div></a>""",
            unsafe_allow_html=True,
        )
        continue
    up = pct >= 0
    color = "#00d488" if up else "#ff4b4b"
    arrow = "▲" if up else "▼"
    col.markdown(
        f"""<a href="{url}" target="_blank" style="text-decoration:none;">
             <div style="background:#1a1a2e;border:1px solid {color}55;border-radius:8px;
             padding:10px 14px;text-align:center;cursor:pointer;">
             <div style="font-size:0.8rem;color:#aaa;">{idx['name']} ↗</div>
             <div style="font-size:1.4rem;font-weight:700;color:{color};">{lvl:,.2f}</div>
             <div style="font-size:0.9rem;font-weight:600;color:{color};">
                 {arrow} {fn(abs(chg))} ({fp(pct)})</div></div></a>""",
        unsafe_allow_html=True,
    )

st.markdown("---")

# ── Scrolling banners: earnings + names near their 52-week high / low ─────────
earnings_banner("📅 UPCOMING EARNINGS",
                upcoming_earnings(port_tickers, business_days=10), label_bg="#1f9bff")
ticker_banner("🔺 HOLDINGS NEAR 52W HIGH",
              near_52w_high(port_tickers, threshold=10.0), label_bg="#00d488")
ticker_banner("🔻 HOLDINGS NEAR 52W LOW",
              near_52w_low(port_tickers, threshold=10.0), label_bg="#ff4b4b")
ticker_banner("🔻 WATCHLIST NEAR 52W LOW",
              near_52w_low(get_watchlist(), threshold=10.0), label_bg="#ff8c1a")

st.markdown("---")

# ── Holdings, Watchlist & Fixed Income navigation cards
nav_1, nav_2, nav_3, nav_4 = st.columns(4)
with nav_1:
    if st.button("📋  View Holdings", type="primary",
                 width="stretch", key="holdings_btn"):
        st.session_state.page = "holdings"
        st.rerun()
with nav_2:
    if st.button("👁  View Watchlist", type="primary",
                 width="stretch", key="watchlist_btn"):
        st.session_state.page = "watchlist"
        st.rerun()
with nav_3:
    if st.button("🏦  Fixed Income", type="primary",
                 width="stretch", key="fixedincome_btn"):
        st.session_state.page = "fixedincome"
        st.rerun()
with nav_4:
    if st.button("📝  Change Log", type="primary",
                 width="stretch", key="changelog_btn"):
        st.session_state.page = "changelog"
        st.rerun()

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — Portfolio Movers (Top 10 + Bottom 10)
# ─────────────────────────────────────────────────────────────────────────────
def _movers_table(rows, ascending):
    """Build a Top/Bottom movers table: Ticker, Name, Price, Chg %, Analyst."""
    return (
        pd.DataFrame(rows)
        .sort_values("_chg_sort", ascending=ascending)
        .drop(columns=["_chg_sort", "Qty"], errors="ignore")
        .reset_index(drop=True)
        .head(10)
    )

col_top, col_bot = st.columns(2)

with col_top:
    st.subheader("Top 10 Portfolio Movers Today")
    _mt = _movers_table(port_rows, ascending=False)
    show_ticker_df(_mt, extra_config=mover_num_config(_mt),
                   width="stretch", hide_index=True, height=370)

with col_bot:
    st.subheader("Bottom 10 Portfolio Movers Today")
    _mb = _movers_table(port_rows, ascending=True)
    show_ticker_df(_mb, extra_config=mover_num_config(_mb),
                   width="stretch", hide_index=True, height=370)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2b — Watchlist Movers (Top 10 + Bottom 10)
# ─────────────────────────────────────────────────────────────────────────────
_watchlist = get_watchlist()
with st.spinner("Loading watchlist…"):
    wl_quotes = get_quotes(_watchlist)
    wl_funds  = get_fundamentals(_watchlist)
    wl_recs   = get_recommendations(_watchlist)

wl_mover_rows = []
for t in _watchlist:
    q = wl_quotes.get(t, {})
    if q.get("current") is None:
        continue
    wl_mover_rows.append({
        "Ticker":    t,
        "Name":      (wl_funds.get(t, {}).get("name", t) or t)[:28],
        "Price":     q.get("current"),
        "Chg %":     q.get("day_change_pct"),
        "_chg_sort": q.get("day_change_pct") or 0,
        "Fwd P/E":   wl_funds.get(t, {}).get("forward_pe"),
        "Analyst":   wl_recs.get(t, {}).get("opinion", "—"),
    })

col_wtop, col_wbot = st.columns(2)

with col_wtop:
    st.subheader("Top 10 Watchlist Movers Today")
    if wl_mover_rows:
        _wmt = _movers_table(wl_mover_rows, ascending=False)
        show_ticker_df(_wmt, extra_config=mover_num_config(_wmt),
                       width="stretch", hide_index=True, height=370)
    else:
        st.info("Watchlist data loading…")

with col_wbot:
    st.subheader("Bottom 10 Watchlist Movers Today")
    if wl_mover_rows:
        _wmb = _movers_table(wl_mover_rows, ascending=True)
        show_ticker_df(_wmb, extra_config=mover_num_config(_wmb),
                       width="stretch", hide_index=True, height=370)
    else:
        st.info("Watchlist data loading…")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — Portfolio Analyst Upgrades/Downgrades
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("Portfolio — Analyst Upgrades/Downgrades (Last 60 Days)")
with st.spinner(""):
    port_actions = get_analyst_actions(port_tickers)

if not port_actions.empty:
    display_cols = ["Date", "Ticker", "Firm", "Action", "From", "To"]
    show_ticker_df(
        add_name_col(port_actions[display_cols]),
        width="stretch",
        hide_index=True,
        height=280,
    )
else:
    st.info("No recent analyst actions found for your portfolio holdings.")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — Market Movers (S&P 500 + Pre-Market)
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("Market Movers")

with st.spinner("Fetching market data…"):
    sp500 = get_sp500_tickers()

col_pre, col_open = st.columns(2)

with col_pre:
    st.markdown("**Top 10 Pre-Market Movers (S&P 500 + Dow)**")
    # Use Dow 30 + first 70 S&P 500 for speed
    pm_universe = tuple(set(DOW_30) | set(sp500[:70]))
    pm_df = build_premarket(pm_universe, n=10)
    if not pm_df.empty:
        pm_show = add_name_col(pm_df)
        show_ticker_df(pm_show, extra_config=mover_num_config(pm_show),
                       width="stretch", hide_index=True, height=370)
    else:
        st.info("Pre-market data not available at this time.")

with col_open:
    st.markdown("**Top 10 Open Market Movers (S&P 500)**")
    if sp500:
        sp_df = build_ohlc_movers(sp500, n=10)
        if not sp_df.empty:
            sp_show = add_name_col(sp_df)
            show_ticker_df(sp_show, extra_config=mover_num_config(sp_show),
                           width="stretch", hide_index=True, height=370)
        else:
            st.info("Market data loading…")
    else:
        st.warning("Could not load S&P 500 ticker list.")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — Dow 30 Movers
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("Dow Jones 30 — Today's Movers")
dow_df = build_ohlc_movers(DOW_30, n=10)
if not dow_df.empty:
    dow_show = add_name_col(dow_df)
    show_ticker_df(dow_show, extra_config=mover_num_config(dow_show),
                   width="stretch", hide_index=True, height=370)
else:
    st.info("Dow data loading…")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 — Market Analyst Actions (S&P 500 + Dow 30)
# ─────────────────────────────────────────────────────────────────────────────
col_sp_ana, col_dow_ana = st.columns(2)

# Use a curated subset of large-cap S&P 500 stocks for analyst actions
# (fetching all 500 would be very slow on free data)
SP500_LARGECAP_SAMPLE = (
    "AAPL","MSFT","AMZN","NVDA","GOOGL","META","TSLA","BRK-B","JPM","UNH",
    "V","XOM","MA","LLY","JNJ","PG","AVGO","MRK","HD","CVX",
    "ABBV","PEP","KO","COST","WMT","BAC","MCD","TMO","CSCO","ABT",
    "ACN","CRM","DHR","NEE","TXN","NKE","QCOM","PM","RTX","HON",
    "AMGN","LOW","IBM","SPGI","GS","CAT","BLK","ELV","SYK","GILD",
    "MDT","MMM","DE","ADP","BKNG","SBUX","ZTS","MO","CB","ISRG",
    "AMD","REGN","TJX","CI","PLD","SCHW","SO","DUK","AON","BSX",
    "CME","ADI","VRTX","EOG","SLB","NOC","ITW","GD","HCA","ICE",
    "CCI","WM","FISV","KLAC","LRCX","MCK","PSA","TGT","APH","NSC",
    "USB","PNC","ORLY","AZO","F","GM","D","EW","IDXX","SHW",
)

with col_sp_ana:
    st.subheader("S&P 500 — Analyst Upgrades/Downgrades (60 Days)")
    with st.spinner(""):
        sp_actions = get_analyst_actions(SP500_LARGECAP_SAMPLE)
    if not sp_actions.empty:
        display_cols = ["Date", "Ticker", "Firm", "Action", "From", "To"]
        show_ticker_df(add_name_col(sp_actions[display_cols].head(30)),
                       width="stretch", hide_index=True, height=400)
    else:
        st.info("No recent S&P 500 analyst actions found.")

with col_dow_ana:
    st.subheader("Dow 30 — Analyst Upgrades/Downgrades (60 Days)")
    with st.spinner(""):
        dow_actions = get_analyst_actions(DOW_30)
    if not dow_actions.empty:
        display_cols = ["Date", "Ticker", "Firm", "Action", "From", "To"]
        show_ticker_df(add_name_col(dow_actions[display_cols].head(30)),
                       width="stretch", hide_index=True, height=400)
    else:
        st.info("No recent Dow analyst actions found.")

st.markdown("---")
st.caption(
    f"Data provided by Yahoo Finance via yfinance.  "
    f"Dashboard auto-refreshes every 60 seconds.  "
    f"Last load: {datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')}"
)
