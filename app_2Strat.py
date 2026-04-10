"""
Portfolio Simulator
================================
Reads SPX_DATA.xlsx from the same directory as app.py.
No file upload needed — just place the workbook next to app.py and run.

Changes in this version:
  1. Auto-loads workbook from same dir — no upload widget
  2. Start-year / End-year range selectors in sidebar
  3. Monthly-return heat map (years × months)
  4. Quarterly holdings visual (ticker presence grid)
  5. Tax report — sells table, bar chart, CSV download
"""

import os, subprocess, tempfile, warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import date
import yfinance as yf

# ─── Page config ─────────────────────────────────────────────
st.set_page_config(page_title="Portfolio Simulator", page_icon="📈",
                   layout="wide", initial_sidebar_state="expanded")

# ─── CSS ─────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

/* ── Base ── */
html,body,[class*="css"]{font-family:'Sora',sans-serif;font-size:15px;color:#0A0A0A;}
.stApp{background:#F5F3EF;}
p,li,span,div{color:#0A0A0A;}

/* ── Sidebar ── */
section[data-testid="stSidebar"]{background:#FEFEFE;border-right:2px solid #E8E4DC;}
section[data-testid="stSidebar"] h3{
    font-size:11px!important;letter-spacing:2px;text-transform:uppercase;
    color:#00B4A6!important;margin:1.5rem 0 0.5rem!important;font-weight:800!important;}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stNumberInput label,
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stDateInput label{
    font-size:13px!important;font-weight:600!important;color:#0A0A0A!important;}

/* ── Banner ── */
.banner{
    background:linear-gradient(135deg,#0A3060 0%,#0952A0 60%,#1A6FD4 100%);
    border-radius:16px;padding:28px 36px;margin-bottom:24px;
    display:flex;align-items:center;}
.banner h1{color:#FFFFF0 !important;font-size:26px;font-weight:800;margin:0;letter-spacing:-0.3px;
           font-family:"Sora",sans-serif !important;
           text-shadow:0 2px 10px rgba(0,0,0,0.7);}

/* ── KPI cards ── */
.kpi{
    background:#FFFFFF;border:1.5px solid #E8E4DC;
    border-radius:14px;padding:20px 24px;
    border-top:3px solid #00B4A6;}
.kpi-lbl{
    font-size:11px;letter-spacing:1.6px;text-transform:uppercase;
    color:#555555;font-weight:700;margin-bottom:6px;}
.kpi-val{
    font-size:27px;font-weight:800;color:#0A0A0A;
    font-family:'JetBrains Mono',monospace;line-height:1.1;}
.kpi-d{font-size:13px;font-weight:700;margin-top:4px;}
.g{color:#00A878;}.r{color:#E03535;}.b{color:#5B5FEF;}.gr{color:#777777;}

/* ── Section headers ── */
.sec{
    font-size:13px;font-weight:800;color:#0A0A0A;
    text-transform:uppercase;letter-spacing:1.4px;
    border-left:4px solid #FF6B35;padding-left:12px;margin:0 0 16px;}

/* ── Cards ── */
.card{
    background:#FFFFFF;border:1.5px solid #E8E4DC;
    border-radius:14px;padding:22px;margin-bottom:20px;}

/* ── Expanders ── */
div[data-testid="stExpander"]{
    background:#FFFFFF!important;border:1.5px solid #E8E4DC!important;
    border-radius:14px!important;}
div[data-testid="stExpander"] summary{
    font-size:14px!important;font-weight:700!important;color:#0A0A0A!important;}

/* ── Dataframe text ── */
.dataframe td,.dataframe th{font-size:13px!important;}

/* ── Streamlit native widget text size ── */
.stSelectbox>div,[data-baseweb="select"] span{font-size:14px!important;}
.stNumberInput input,.stDateInput input{font-size:14px!important;color:#0A0A0A!important;}
.stButton>button{font-size:14px!important;font-weight:700!important;}
.stToggle label{font-size:14px!important;}
.stCaption,.element-container .stMarkdown p{font-size:13px!important;color:#333333!important;}
</style>
""", unsafe_allow_html=True)

# ─── Workbook path ────────────────────────────────────────────
# Robust resolution for Streamlit Cloud where __file__ may resolve
# to a runner directory rather than the repo root.
_FNAME = "SPX_DATA.xlsx"

def _find_xlsx(fname):
    candidates = [
        # 1. Same dir as this script (local & most Streamlit Cloud cases)
        os.path.join(os.path.dirname(os.path.abspath(__file__)), fname),
        # 2. Current working directory (Streamlit Cloud sometimes sets cwd to repo root)
        os.path.join(os.getcwd(), fname),
        # 3. Streamlit Cloud canonical repo mount path
        os.path.join("/mount/src/tcmequitygrowth", fname),
        # 4. One level up from script (monorepo layouts)
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), fname),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p, os.path.dirname(p)
    return None, os.path.dirname(os.path.abspath(__file__))

XLSX_PATH, APP_DIR = _find_xlsx(_FNAME)

# ═════════════════════════════════════════════════════════════
# PARSING
# ═════════════════════════════════════════════════════════════

def _parse_sheet(df, year):
    result = {}
    hdr_idx = next((i for i, row in df.iterrows()
                    if any(str(v).strip().lower() == "ticker" for v in row if pd.notna(v))), 3)
    hdr    = df.iloc[hdr_idx]
    tk_col = next((j for j,v in enumerate(hdr)
                   if pd.notna(v) and str(v).strip().lower() == "ticker"), 2)
    q1_cap = next((j for j,v in enumerate(hdr)
                   if pd.notna(v) and "q1" in str(v).lower() and "cap" in str(v).lower()),
                  tk_col + 2)
    for q in range(1, 5):
        cap_col = q1_cap + (q - 1)
        tickers = []
        for _, row in df.iloc[hdr_idx + 1:].iterrows():
            rank = row.iloc[0]
            if pd.isna(rank): break
            if isinstance(rank, str) and "total" in rank.lower(): break
            tk = str(row.iloc[tk_col]).strip().upper() if pd.notna(row.iloc[tk_col]) else ""
            if not tk or not (1 <= len(tk) <= 6 and tk.replace(".", "").isalpha()): continue
            cap = row.iloc[cap_col] if cap_col < len(row) else None
            if pd.notna(cap) and str(cap) not in ("—", "-", ""):
                try:
                    float(cap); tickers.append(tk)
                except (TypeError, ValueError):
                    pass
        if tickers:
            result[f"{year}-Q{q}"] = tickers[:15]
    return result


@st.cache_data(show_spinner=False)
def parse_workbook(path):
    quarters = {}

    def _try(fpath, engine=None):
        kw = {} if engine is None else {"engine": engine}
        try:
            xf = pd.ExcelFile(fpath, **kw)
            for s in [x for x in xf.sheet_names if x.strip().isdigit()]:
                quarters.update(_parse_sheet(xf.parse(s, header=None), int(s)))
            return bool(quarters)
        except Exception:
            return False

    if _try(path, "openpyxl"): return quarters
    if _try(path, "odf"):      return quarters
    try:
        tmp = tempfile.mkdtemp()
        subprocess.run(["libreoffice","--headless","--convert-to","ods",path,"--outdir",tmp],
                       capture_output=True, timeout=90)
        ods = os.path.join(tmp, os.path.splitext(os.path.basename(path))[0]+".ods")
        if os.path.exists(ods): _try(ods, "odf")
    except Exception:
        pass
    return quarters


# ═════════════════════════════════════════════════════════════
# PRICE FETCHING
# ═════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False, ttl=3600)
def fetch_prices(tickers, start, end):
    if not tickers: return pd.DataFrame()
    try:
        raw = yf.download(list(tickers), start=start, end=end,
                          auto_adjust=True, progress=False, group_by="ticker")
        if raw.empty: return pd.DataFrame()
        if isinstance(raw.columns, pd.MultiIndex):
            lvl0   = raw.columns.get_level_values(0).unique().tolist()
            fields = {"Close","Open","High","Low","Volume","Adj Close"}
            if any(v in fields for v in lvl0):
                px = raw["Close"]
            else:
                cc = [(t,"Close") for t in lvl0 if (t,"Close") in raw.columns]
                px = raw[cc]; px.columns = [c[0] for c in px.columns]
        else:
            px = raw[["Close"]] if "Close" in raw.columns else raw
        if isinstance(px, pd.Series):
            px = px.to_frame(name=str(list(tickers)[0]))
        px.columns = [c if isinstance(c,str) else str(c) for c in px.columns]
        px.index = pd.to_datetime(px.index)
        return px.resample("MS").first()
    except Exception as e:
        st.warning(f"Price fetch issue: {e}"); return pd.DataFrame()


# ═════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════

def q2date(label):
    y,q = label.split("-Q")
    return date(int(y),{1:1,2:4,3:7,4:10}[int(q)],1)

def px_at(prices, ticker, dt):
    if ticker not in prices.columns: return None
    col = prices[ticker].dropna()
    fut = col[col.index >= pd.Timestamp(dt)]
    return float(fut.iloc[0]) if not fut.empty else None

def max_dd(s):
    peak = s.expanding().max()
    return float(((s-peak)/peak).min())

def cagr(v0, v1, yrs):
    return (v1/v0)**(1/yrs)-1 if v0>0 and yrs>0 else 0.0

def fmt_usd(v):
    if abs(v)>=1e9: return f"${v/1e9:.2f}B"
    if abs(v)>=1e6: return f"${v/1e6:.2f}M"
    return f"${v:,.0f}"

def fmt_pct(v, sign=True):
    return f"{'+'if sign and v>=0 else ''}{v:.1f}%"


# ═════════════════════════════════════════════════════════════
# SIMULATION (returns sell_log for tax report)
# ═════════════════════════════════════════════════════════════

def simulate(quarters, prices, start_amt, sim_start, dca_on, dca_monthly, dca_start,
             hedge_pct=5.0):
    active = {k:v for k,v in quarters.items() if q2date(k)>=sim_start}
    if not active or prices.empty: return pd.DataFrame(), {}, [], []

    sorted_q = sorted(active)
    qd       = {k: q2date(k) for k in sorted_q}
    first_dt = q2date(sorted_q[0])
    last_qd  = q2date(sorted_q[-1])
    end_dt   = min(date(last_qd.year+(last_qd.month>9),(last_qd.month+2)%12+1,1), date.today())

    holdings   = {}          # ticker → shares
    share_cost = {}          # ticker → avg cost per share
    cash       = start_amt
    tot_inv    = start_amt
    cur_tickers= []
    cost_basis = {}
    rows       = []
    sell_log   = []
    fee_log    = []
    monthly_hedge_rate = (hedge_pct / 100.0) / 12.0
    yr_start_val = start_amt

    for mts in pd.date_range(first_dt, end_dt, freq="MS"):
        md    = mts.date()
        if md.month == 1 and rows:
            yr_start_val = float(rows[-1]["Portfolio"])
        rebal = next((k for k,d in qd.items() if d.year==md.year and d.month==md.month), None)

        if rebal:
            new_tk    = active[rebal]
            qtr_label = rebal

            # 1. Sell fully dropped tickers
            for t in list(holdings):
                if t not in new_tk and holdings[t] > 0:
                    p        = px_at(prices, t, md)
                    sh       = holdings[t]
                    proceeds = sh * p if p else 0
                    avg_c    = share_cost.get(t, 0)
                    gain     = proceeds - sh * avg_c
                    cash    += proceeds
                    sell_log.append({"Year":md.year,"Quarter":qtr_label,"Ticker":t,
                                     "Shares":round(sh,4),"Price":round(p,2) if p else 0,
                                     "Proceeds":round(proceeds,2),
                                     "Cost Basis":round(sh*avg_c,2),
                                     "Gain / Loss":round(gain,2)})
                    holdings[t]=0.0; share_cost[t]=0.0

            cur_tickers = new_tk

            # 2. DCA
            if dca_on and md >= dca_start:
                cash += dca_monthly; tot_inv += dca_monthly

            # 3. Portfolio value
            port_v = cash + sum(holdings.get(t,0)*(px_at(prices,t,md) or 0) for t in cur_tickers)

            # 4. Equal-weight rebalance
            n = len(cur_tickers)
            if n > 0:
                target = port_v / n
                for t in cur_tickers:
                    p = px_at(prices, t, md)
                    if not p or p <= 0: continue
                    cur_v = holdings.get(t,0) * p
                    dv    = target - cur_v
                    if dv < 0:          # partial sell
                        sh_sold  = abs(dv)/p
                        avg_c    = share_cost.get(t,0)
                        gain     = sh_sold*(p-avg_c)
                        sell_log.append({"Year":md.year,"Quarter":qtr_label,"Ticker":t,
                                         "Shares":round(sh_sold,4),"Price":round(p,2),
                                         "Proceeds":round(sh_sold*p,2),
                                         "Cost Basis":round(sh_sold*avg_c,2),
                                         "Gain / Loss":round(gain,2)})
                    holdings[t] = holdings.get(t,0) + dv/p
                    cash       -= dv
                    if dv > 0:
                        prev_sh   = holdings[t] - dv/p
                        prev_cost = share_cost.get(t,0)*prev_sh
                        share_cost[t] = (prev_cost+dv)/holdings[t] if holdings[t]>0 else 0
                        cost_basis[t] = cost_basis.get(t,0)+dv

        else:
            if dca_on and md >= dca_start and cur_tickers:
                per_t    = dca_monthly/len(cur_tickers)
                tot_inv += dca_monthly
                for t in cur_tickers:
                    p = px_at(prices,t,md)
                    if p and p>0:
                        new_sh       = per_t/p
                        prev_sh      = holdings.get(t,0)
                        prev_cost    = share_cost.get(t,0)*prev_sh
                        holdings[t]  = prev_sh+new_sh
                        share_cost[t]= (prev_cost+per_t)/holdings[t]
                        cost_basis[t]= cost_basis.get(t,0)+per_t

        pv = max(cash, 0)
        for t,sh in holdings.items():
            if sh>0:
                p = px_at(prices,t,md)
                if p: pv += sh*p

        # Monthly hedge cost (equity portfolio only)
        hedge_cost = pv * monthly_hedge_rate
        cash      -= hedge_cost
        pv        -= hedge_cost

        fee_log.append({
            "Year":          md.year,
            "Month":         mts.strftime("%b"),
            "Equity AUM":    round(max(pv, 0), 2),
            "Hedge Cost":    round(hedge_cost, 2),
        })
        rows.append({"Date":mts,"Portfolio":max(pv,0),"Invested":tot_inv})

    return pd.DataFrame(rows).set_index("Date"), cost_basis, sell_log, fee_log


@st.cache_data(show_spinner=False, ttl=3600)
def benchmark(ticker, start, initial, dca_on, dca_monthly, dca_start_str):
    try:
        raw = yf.download(ticker, start=start, auto_adjust=True, progress=False)
        if raw.empty: return pd.DataFrame()
        if isinstance(raw.columns, pd.MultiIndex):
            lvl0   = raw.columns.get_level_values(0).unique().tolist()
            fields = {"Close","Open","High","Low","Volume","Adj Close"}
            col    = raw["Close"] if any(v in fields for v in lvl0) else raw.iloc[:,0]
        else:
            col = raw["Close"] if "Close" in raw.columns else raw.iloc[:,0]
        if isinstance(col, pd.DataFrame): col = col.squeeze()
        col = col.resample("MS").first().dropna()
        col.index = pd.to_datetime(col.index)
        if col.empty: return pd.DataFrame()
        dca_ts  = pd.Timestamp(dca_start_str)
        shares  = float(initial)/float(col.iloc[0])
        tot_inv = float(initial)
        rows    = []
        for ts, price in col.items():
            p = float(price)
            if dca_on and ts >= dca_ts:
                shares += float(dca_monthly)/p; tot_inv += float(dca_monthly)
            rows.append({"Date":ts,"Benchmark":shares*p,"Invested":tot_inv})
        return pd.DataFrame(rows).set_index("Date")
    except Exception as e:
        st.warning(f"Benchmark issue ({ticker}): {e}"); return pd.DataFrame()


def simulate_option_income(start_amt, sim_start_date, cagr_pct,
                            dca_on, dca_monthly, dca_start,
                            port_index):
    """
    Simulate OptionIncome strategy: pure compounding at fixed CAGR.
    Returns a DataFrame with same Date index as port_df (EquityGrowth).
    Monthly CAGR factor = (1 + annual_cagr)^(1/12) - 1
    DCA additions are invested immediately at current value and compound forward.
    """
    monthly_r = (1.0 + cagr_pct / 100.0) ** (1.0 / 12.0) - 1.0
    pv        = float(start_amt)
    tot_inv   = float(start_amt)
    rows      = []
    for mts in port_index:
        md = mts.date()
        if dca_on and md >= dca_start:
            pv      += float(dca_monthly)
            tot_inv += float(dca_monthly)
        pv *= (1.0 + monthly_r)
        rows.append({"Date": mts, "OI_Portfolio": round(pv, 2), "OI_Invested": round(tot_inv, 2)})
    return pd.DataFrame(rows).set_index("Date")


# ═════════════════════════════════════════════════════════════
# SIDEBAR
# ═════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### ⚙️ Simulation")
    start_amount = st.number_input("Starting Amount ($)", 1_000, 10_000_000, 100_000, 5_000, format="%d")
    sim_start    = st.date_input("Simulation Start", value=date(2000,1,1),
                                 min_value=date(2000,1,1), max_value=date(2024,1,1))

    st.markdown("### 💰 Dollar-Cost Averaging")
    dca_on  = st.toggle("Enable DCA", value=False)
    dca_amt = st.number_input("DCA Amount ($/month)", 100, 50_000, 1_000, 100, format="%d", disabled=not dca_on)
    dca_dt  = st.date_input("DCA Start Date", value=date(2000,1,1), min_value=date(2000,1,1), disabled=not dca_on)

    st.markdown("### 🛡️ Costs & Fees")
    hedge_pct = st.number_input("Hedging Cost (% AUM/yr)", 0.0, 20.0, 5.0, 0.5, format="%.1f",
                                 help="Deducted monthly (1/12 per month) from AUM")
    mgmt_fee  = st.number_input("Management Fee (% AUM/yr)", 0.0, 10.0, 2.0, 0.25, format="%.2f",
                                 help="Deducted at December year-end on AUM")
    perf_fee  = st.number_input("Performance Fee (% of yearly gain)", 0.0, 50.0, 20.0, 1.0, format="%.1f",
                                 help="Charged at December year-end on positive gain only")

    st.markdown("### 📊 Strategy Allocation")
    eq_pct = st.slider(
        "EquityGrowth allocation (%)", 0, 100, 50, 5,
        help="Remaining % goes to OptionIncome. Default 50/50 split."
    )
    oi_pct = 100 - eq_pct
    st.caption(f"EquityGrowth: {eq_pct}%  ·  OptionIncome: {oi_pct}%")
    oi_cagr = st.number_input(
        "OptionIncome CAGR (%)", 5.0, 100.0, 30.0, 1.0, format="%.1f",
        help="Compounded annual growth rate applied monthly to the OptionIncome allocation"
    )

    st.markdown("### 📅 Report Year Range")
    yr_start = st.selectbox("Start Year", list(range(2000,2026)), index=0)
    yr_end   = st.selectbox("End Year",   list(range(2000,2026)), index=25)
    if yr_end < yr_start: yr_end = yr_start

    st.markdown("### 📈 Add Tickers to Growth Chart")
    st.caption("Select any historical ticker to overlay on the growth chart")
    # Populated after workbook loads; use session_state to persist across reruns
    _available = sorted(st.session_state.get("all_tickers_list", []))
    overlay_tickers = st.multiselect(
        "Choose tickers",
        options=_available,
        default=[],
        placeholder="e.g. AAPL, MSFT, NVDA …",
        help="Each ticker will be simulated as a standalone $100K lump-sum investment "
             "starting from the same simulation start date, for comparison.",
    )

    st.markdown("---")
    run_btn = st.button("▶  Run Simulation", type="primary", use_container_width=True)
    st.caption("Prices via yfinance · split & div adjusted")


# ═════════════════════════════════════════════════════════════
# BANNER
# ═════════════════════════════════════════════════════════════

st.markdown("""
<div class="banner">
  <div>
    <h1 style="color:#FFFFF0 !important;font-family:Sora,sans-serif !important;font-size:26px;font-weight:800;margin:0;letter-spacing:-0.3px;text-shadow:0 2px 10px rgba(0,0,0,0.7);">📈 Portfolio Simulator</h1>
  </div>
</div>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════
# LOAD WORKBOOK FROM SAME DIR
# ═════════════════════════════════════════════════════════════

if not XLSX_PATH or not os.path.exists(XLSX_PATH):
    searched = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), _FNAME),
        os.path.join(os.getcwd(), _FNAME),
        os.path.join("/mount/src/tcmequitygrowth", _FNAME),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), _FNAME),
    ]
    st.error(
        "❌ **SPX_DATA.xlsx not found.** Searched:\n\n"
        + "\n".join(f"- `{p}`" for p in searched)
        + f"\n\n**cwd:** `{os.getcwd()}`\n\n**__file__:** `{os.path.abspath(__file__)}`"
    )
    st.stop()

with st.spinner("Parsing workbook…"):
    quarters = parse_workbook(XLSX_PATH)

if not quarters:
    st.error("❌ Could not parse the workbook.")
    st.stop()

all_tickers = sorted({t for v in quarters.values() for t in v})
st.success(f"✅ Loaded **{len(quarters)} quarters** · **{len(all_tickers)} tickers** ({min(quarters)} → {max(quarters)})")
# Expose tickers list for sidebar multiselect
st.session_state["all_tickers_list"] = all_tickers


# ═════════════════════════════════════════════════════════════
# RUN
# ═════════════════════════════════════════════════════════════

if run_btn:
    start_str = str(sim_start); end_str = str(date.today())

    # Split capital by allocation percentages
    _eq_amt = float(start_amount) * eq_pct / 100.0
    _oi_amt = float(start_amount) * oi_pct / 100.0
    _eq_dca = float(dca_amt) * eq_pct / 100.0 if dca_on else 0.0
    _oi_dca = float(dca_amt) * oi_pct / 100.0 if dca_on else 0.0

    with st.spinner(f"⬇ Fetching {len(all_tickers)} tickers + SPY + QQQ…"):
        prices_df = fetch_prices(tuple(sorted(set(all_tickers+["SPY","QQQ"]))), start_str, end_str)
    if prices_df.empty:
        st.error("Could not fetch price data."); st.stop()

    with st.spinner("⚙️ Simulating EquityGrowth…"):
        port_df, contrib, sell_log, eq_fee_log = simulate(
            quarters, prices_df, _eq_amt, sim_start,
            dca_on, _eq_dca, dca_dt,
            hedge_pct=float(hedge_pct))
    if port_df.empty:
        st.error("Simulation returned no data."); st.stop()

    with st.spinner("⚙️ Simulating OptionIncome…"):
        oi_df = simulate_option_income(
            _oi_amt, sim_start, float(oi_cagr),
            dca_on, _oi_dca, dca_dt, port_df.index)

    # ── Apply mgmt (2%) + perf (20%) fees to COMBINED total at year-end ──────
    # Merge equity + OI into combined monthly series, then deduct fees
    _comb_raw   = port_df["Portfolio"].values.copy().astype(float)
    _oi_raw     = oi_df["OI_Portfolio"].values.copy().astype(float) if not oi_df.empty else np.zeros(len(_comb_raw))
    _total_raw  = _comb_raw + _oi_raw
    _total_inv  = (port_df["Invested"].values + (oi_df["OI_Invested"].values if not oi_df.empty else 0)).astype(float)

    _comb_net   = _total_raw.copy()   # will accumulate fee deductions
    _fee_log    = []                  # combined fee log
    _yr_sv      = float(_comb_net[0])

    for _idx, _mts in enumerate(port_df.index):
        _md = _mts.date()
        if _md.month == 1 and _idx > 0:
            _yr_sv = float(_comb_net[_idx - 1])
        _gross = float(_comb_net[_idx])
        _hedge_eq = float(eq_fee_log[_idx]["Hedge Cost"]) if _idx < len(eq_fee_log) else 0.0

        _yr_mgmt = 0.0; _yr_perf = 0.0
        if _md.month == 12:
            _yr_mgmt        = _gross * (float(mgmt_fee) / 100.0)
            # Perf fee on net gain AFTER mgmt fee — avoids double-counting
            _net_after_mgmt = _gross - _yr_mgmt
            _yr_gain        = _net_after_mgmt - _yr_sv
            _yr_perf        = max(_yr_gain, 0) * (float(perf_fee) / 100.0)
            # Deduct proportionally from EQ and OI
            _total_fees = _yr_mgmt + _yr_perf
            _eq_share   = _comb_raw[_idx] / _gross if _gross > 0 else 0.5
            _oi_share   = 1.0 - _eq_share
            _comb_raw[_idx] = max(_comb_raw[_idx] - _total_fees * _eq_share, 0)
            _oi_raw[_idx]   = max(_oi_raw[_idx]   - _total_fees * _oi_share, 0)
            _comb_net[_idx] = _comb_raw[_idx] + _oi_raw[_idx]

        _fee_log.append({
            "Year":          _md.year,
            "Month":         _mts.strftime("%b"),
            "Equity AUM":    round(float(eq_fee_log[_idx]["Equity AUM"]) if _idx < len(eq_fee_log) else 0, 2),
            "OI AUM":        round(float(_oi_raw[_idx]), 2),
            "Combined AUM":  round(float(_comb_net[_idx]), 2),
            "Hedge Cost":    round(_hedge_eq, 2),
            "Mgmt Fee":      round(_yr_mgmt, 2),
            "Perf Fee":      round(_yr_perf, 2),
            "Total Deducted":round(_hedge_eq + _yr_mgmt + _yr_perf, 2),
        })

    # Rebuild port_df and oi_df with fee-adjusted values
    port_df = port_df.copy()
    port_df["Portfolio"] = _comb_raw
    if not oi_df.empty:
        oi_df = oi_df.copy()
        oi_df["OI_Portfolio"] = _oi_raw
    fee_log = _fee_log

    with st.spinner("⚙️ Benchmarks…"):
        spy_df = benchmark("SPY", start_str, float(start_amount), dca_on, float(dca_amt), str(dca_dt))
        qqq_df = benchmark("QQQ", start_str, float(start_amount), dca_on, float(dca_amt), str(dca_dt))

    # Fetch overlay ticker prices (same date range)
    overlay_px = {}
    if overlay_tickers:
        with st.spinner(f"⬇ Fetching {len(overlay_tickers)} overlay ticker(s)…"):
            ov_raw = fetch_prices(tuple(sorted(overlay_tickers)), start_str, end_str)
            for tk in overlay_tickers:
                if tk in ov_raw.columns:
                    col = ov_raw[tk].dropna()
                    if not col.empty:
                        init_price = float(col.iloc[0])
                        shares     = float(start_amount) / init_price
                        overlay_px[tk] = (col * shares).rename(tk)

    st.session_state.update(dict(port_df=port_df, oi_df=oi_df,
                                  spy_df=spy_df, qqq_df=qqq_df,
                                  contrib=contrib, sell_log=sell_log,
                                  fee_log=fee_log,
                                  quarters=quarters, prices_df=prices_df,
                                  overlay_px=overlay_px, overlay_tickers=overlay_tickers,
                                  eq_pct=eq_pct, oi_pct=oi_pct, oi_cagr=oi_cagr))

if "port_df" not in st.session_state:
    st.info("👈 Click **Run Simulation** to begin."); st.stop()

port_df  = st.session_state["port_df"]
oi_df    = st.session_state.get("oi_df", pd.DataFrame())
spy_df   = st.session_state["spy_df"]
qqq_df   = st.session_state["qqq_df"]
contrib  = st.session_state["contrib"]
sell_log = st.session_state["sell_log"]
fee_log  = st.session_state.get("fee_log", [])
qs_data  = st.session_state["quarters"]
_eq_pct_ss = st.session_state.get("eq_pct", eq_pct)
_oi_pct_ss = st.session_state.get("oi_pct", oi_pct)
_oi_cagr_ss= st.session_state.get("oi_cagr", oi_cagr)

# Combined series for financial reporting (EquityGrowth + OptionIncome)
_comb_for_report     = port_df["Portfolio"] + (oi_df["OI_Portfolio"] if not oi_df.empty else 0)
_comb_inv_for_report = port_df["Invested"]  + (oi_df["OI_Invested"]  if not oi_df.empty else 0)
mdf_all = pd.DataFrame({
    "Date":      port_df.index,
    "Portfolio": _comb_for_report.values,
    "Invested":  _comb_inv_for_report.values,
})
mdf_all["Year"]   = mdf_all["Date"].dt.year
mdf_all["Month"]  = mdf_all["Date"].dt.strftime("%b")
mdf_all["MoM %"]  = mdf_all["Portfolio"].pct_change()*100
mdf_all["Gain $"] = mdf_all["Portfolio"] - mdf_all["Invested"]
mdf_all["Gain %"] = mdf_all["Gain $"]/mdf_all["Invested"]*100
mdf = mdf_all[(mdf_all["Year"]>=yr_start)&(mdf_all["Year"]<=yr_end)].copy()


# ═════════════════════════════════════════════════════════════
# ① KPI CARDS
# ═════════════════════════════════════════════════════════════

# EquityGrowth metrics
eq_final  = float(port_df["Portfolio"].iloc[-1])
eq_inv    = float(port_df["Invested"].iloc[-1])
# OptionIncome metrics
oi_final  = float(oi_df["OI_Portfolio"].iloc[-1]) if not oi_df.empty else 0.0
oi_inv    = float(oi_df["OI_Invested"].iloc[-1])  if not oi_df.empty else 0.0
# Combined total
final_v   = eq_final + oi_final
invested  = eq_inv + oi_inv
tot_ret   = (final_v - invested) / invested * 100 if invested else 0
yrs       = (port_df.index[-1] - port_df.index[0]).days / 365.25
p_cagr    = cagr(float(port_df["Portfolio"].iloc[0]) + (float(oi_df["OI_Portfolio"].iloc[0]) if not oi_df.empty else 0),
                 final_v, yrs) * 100
# Combined portfolio series for drawdown
_comb_series = port_df["Portfolio"] + (oi_df["OI_Portfolio"] if not oi_df.empty else 0)
p_mdd     = max_dd(_comb_series) * 100
final_v   = eq_final + oi_final   # alias for rest of app
spy_v    = float(spy_df["Benchmark"].iloc[-1]) if not spy_df.empty else 0
qqq_v    = float(qqq_df["Benchmark"].iloc[-1]) if not qqq_df.empty else 0
spy_i    = float(spy_df["Invested"].iloc[-1])  if not spy_df.empty else invested
qqq_i    = float(qqq_df["Invested"].iloc[-1])  if not qqq_df.empty else invested
a_spy    = tot_ret-((spy_v-spy_i)/spy_i*100 if spy_i else 0)
a_qqq    = tot_ret-((qqq_v-qqq_i)/qqq_i*100 if qqq_i else 0)

st.markdown('<p class="sec">① Summary KPIs — Combined Portfolio</p>', unsafe_allow_html=True)
# Mini allocation bar
_alloc_html = (f'<div style="display:flex;gap:4px;margin-bottom:14px;border-radius:8px;overflow:hidden;height:8px;">'
               f'<div style="flex:{_eq_pct_ss};background:#5B5FEF;"></div>'
               f'<div style="flex:{_oi_pct_ss};background:#14B864;"></div></div>'
               f'<div style="font-size:11px;color:#555;margin-bottom:16px;">'
               f'<span style="color:#5B5FEF;font-weight:700;">■</span> EquityGrowth {_eq_pct_ss}% &nbsp;|&nbsp; '
               f'<span style="color:#14B864;font-weight:700;">■</span> OptionIncome {_oi_pct_ss}% ({_oi_cagr_ss:.0f}% CAGR)</div>')
st.markdown(_alloc_html, unsafe_allow_html=True)
k1,k2,k3,k4,k5,k6 = st.columns(6)
for col,lbl,val,delta,cls in [
    (k1,"Combined Final", fmt_usd(final_v),"",""),
    (k2,"Total Invested", fmt_usd(invested),"",""),
    (k3,"Total Return",  fmt_pct(tot_ret),fmt_pct(a_spy)+" vs SPY","g"if a_spy>=0 else "r"),
    (k4,"CAGR",          fmt_pct(p_cagr),f"{yrs:.1f} yrs","b"),
    (k5,"Max Drawdown",  fmt_pct(p_mdd,False),"from peak","r"),
    (k6,"vs QQQ Alpha",  fmt_pct(a_qqq),"excess return","g"if a_qqq>=0 else "r"),
]:
    with col:
        st.markdown(f'<div class="kpi"><div class="kpi-lbl">{lbl}</div>'
                    f'<div class="kpi-val">{val}</div>'
                    f'{"" if not delta else f"<div class=kpi-d {cls}>{delta}</div>"}'
                    f'</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════
# ② GROWTH CHART
# ═════════════════════════════════════════════════════════════

st.markdown('<p class="sec">② Growth Charts</p>', unsafe_allow_html=True)

_LO = dict(plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", hovermode="x unified",
           margin=dict(l=0,r=0,t=36,b=0),
           legend=dict(orientation="h",y=1.06,x=0,font=dict(size=11,family="Sora",color="#0A0A0A")),
           xaxis=dict(showgrid=True,gridcolor="#EDE9E0",tickfont=dict(size=10,family="Sora",color="#0A0A0A")),
           yaxis=dict(showgrid=True,gridcolor="#EDE9E0",tickformat="$,.0f",
                      tickfont=dict(size=10,family="JetBrains Mono",color="#0A0A0A")),
           font=dict(family="Sora",size=11,color="#0A0A0A"))

# ── Row 1: side-by-side individual strategies ─────────────────────────────────
_col_eq, _col_oi = st.columns(2)

with _col_eq:
    st.markdown('<div class="card" style="padding:14px 16px;">', unsafe_allow_html=True)
    fig_eq = go.Figure()
    fig_eq.add_trace(go.Scatter(x=port_df.index, y=port_df["Portfolio"],
        name="EquityGrowth",
        line=dict(color="#5B5FEF",width=2.5),
        fill="tozeroy", fillcolor="rgba(91,95,239,.06)",
        hovertemplate="<b>EquityGrowth</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
    fig_eq.add_trace(go.Scatter(x=port_df.index, y=port_df["Invested"],
        name="Invested", line=dict(color="#AAAAAA",width=1.2,dash="dot"),
        hovertemplate="<b>Invested</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
    if not spy_df.empty:
        fig_eq.add_trace(go.Scatter(x=spy_df.index, y=spy_df["Benchmark"],
            name="SPY", line=dict(color="#FF6B35",width=1.8),
            hovertemplate="<b>SPY</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
    if not qqq_df.empty:
        fig_eq.add_trace(go.Scatter(x=qqq_df.index, y=qqq_df["Benchmark"],
            name="QQQ", line=dict(color="#00B4A6",width=1.8),
            hovertemplate="<b>QQQ</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
    # Ticker overlays
    _ov_px  = st.session_state.get("overlay_px", {})
    _ov_tks = st.session_state.get("overlay_tickers", [])
    _OVC    = ["#9333EA","#C026D3","#DB2777","#DC2626","#D97706","#65A30D","#0891B2"]
    for _i,_tk in enumerate(_ov_tks):
        if _tk in _ov_px:
            fig_eq.add_trace(go.Scatter(x=_ov_px[_tk].index, y=_ov_px[_tk].values,
                name=f"{_tk}", line=dict(color=_OVC[_i%len(_OVC)],width=1.5,dash="dashdot"),
                hovertemplate=f"<b>{_tk}</b> %{{x|%b %Y}} · $%{{y:,.0f}}<extra></extra>"))
    fig_eq.update_layout(height=360, title=dict(text=f"EquityGrowth ({_eq_pct_ss}%)",
        font=dict(size=13,color="#5B5FEF",family="Sora"),x=0), **_LO)
    st.plotly_chart(fig_eq, use_container_width=True, config={"displayModeBar":False})
    st.markdown("</div>", unsafe_allow_html=True)

with _col_oi:
    st.markdown('<div class="card" style="padding:14px 16px;">', unsafe_allow_html=True)
    fig_oi = go.Figure()
    if not oi_df.empty:
        fig_oi.add_trace(go.Scatter(x=oi_df.index, y=oi_df["OI_Portfolio"],
            name="OptionIncome",
            line=dict(color="#14B864",width=2.5),
            fill="tozeroy", fillcolor="rgba(20,184,100,.06)",
            hovertemplate="<b>OptionIncome</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
        fig_oi.add_trace(go.Scatter(x=oi_df.index, y=oi_df["OI_Invested"],
            name="Invested", line=dict(color="#AAAAAA",width=1.2,dash="dot"),
            hovertemplate="<b>Invested</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
    fig_oi.update_layout(height=360, title=dict(
        text=f"OptionIncome ({_oi_pct_ss}% · {_oi_cagr_ss:.0f}% CAGR)",
        font=dict(size=13,color="#14B864",family="Sora"),x=0), **_LO)
    st.plotly_chart(fig_oi, use_container_width=True, config={"displayModeBar":False})
    st.markdown("</div>", unsafe_allow_html=True)

# ── Row 2: combined total ─────────────────────────────────────────────────────
st.markdown('<div class="card" style="padding:14px 16px;margin-top:4px;">', unsafe_allow_html=True)
_comb = port_df["Portfolio"] + (oi_df["OI_Portfolio"] if not oi_df.empty else 0)
_comb_inv = port_df["Invested"] + (oi_df["OI_Invested"] if not oi_df.empty else 0)
fig_comb = go.Figure()
fig_comb.add_trace(go.Scatter(x=_comb.index, y=_comb.values,
    name="Total Combined",
    line=dict(color="#0A4DA8",width=3.5),
    fill="tozeroy", fillcolor="rgba(10,77,168,.05)",
    hovertemplate="<b>Total</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
fig_comb.add_trace(go.Scatter(x=port_df.index, y=port_df["Portfolio"],
    name="EquityGrowth", line=dict(color="#5B5FEF",width=2,dash="dash"),
    hovertemplate="<b>EquityGrowth</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
if not oi_df.empty:
    fig_comb.add_trace(go.Scatter(x=oi_df.index, y=oi_df["OI_Portfolio"],
        name="OptionIncome", line=dict(color="#14B864",width=2,dash="dash"),
        hovertemplate="<b>OptionIncome</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
fig_comb.add_trace(go.Scatter(x=_comb_inv.index, y=_comb_inv.values,
    name="Total Invested", line=dict(color="#AAAAAA",width=1.5,dash="dot"),
    hovertemplate="<b>Invested</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
if not spy_df.empty:
    fig_comb.add_trace(go.Scatter(x=spy_df.index, y=spy_df["Benchmark"],
        name="SPY (full capital)", line=dict(color="#FF6B35",width=2),
        hovertemplate="<b>SPY</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
if not qqq_df.empty:
    fig_comb.add_trace(go.Scatter(x=qqq_df.index, y=qqq_df["Benchmark"],
        name="QQQ (full capital)", line=dict(color="#00B4A6",width=2),
        hovertemplate="<b>QQQ</b> %{x|%b %Y} · $%{y:,.0f}<extra></extra>"))
fig_comb.update_layout(height=460, title=dict(text="Combined Total Portfolio",
    font=dict(size=14,color="#0A4DA8",family="Sora"),x=0), **_LO)
st.plotly_chart(fig_comb, use_container_width=True, config={"displayModeBar":False})
st.markdown("</div>", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════
# ③ MONTHLY RETURN HEAT MAP
# ═════════════════════════════════════════════════════════════

st.markdown('<p class="sec">③ Monthly Return Heat Map</p>', unsafe_allow_html=True)
MONTHS     = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
MONTH_NUMS = {m:i+1 for i,m in enumerate(MONTHS)}

# Build heat map data from the full monthly series (no year-range filter —
# Cumul % is always from simulation start; display filter applied via y-axis)
_hm_src = mdf_all[["Year","Month","MoM %"]].copy()
_hm_src = _hm_src[(_hm_src["Year"]>=yr_start)&(_hm_src["Year"]<=yr_end)]
_hm_src = _hm_src.dropna(subset=["MoM %"])
# Add numeric month for correct ordering
_hm_src["MonthN"] = _hm_src["Month"].map(MONTH_NUMS)
_hm_src = _hm_src.sort_values(["Year","MonthN"])

if not _hm_src.empty:
    # Build pivot using numeric month index, then rename columns to abbrev
    _piv_num = (_hm_src.pivot_table(index="Year", columns="MonthN",
                                     values="MoM %", aggfunc="mean")
                .reindex(columns=range(1,13)))
    _piv_num.columns = MONTHS

    # YTD: compound all non-NaN months in each row
    def _ytd(row):
        vals = [v for v in row if pd.notna(v)]
        return (float(np.prod([1 + v/100.0 for v in vals])) - 1) * 100.0 if vals else np.nan
    _ytd_vals = _piv_num.apply(_ytd, axis=1)

    # Cumul: running compound product year-over-year
    _cum = 1.0; _cvs = []
    for _y in _ytd_vals:
        if pd.notna(_y): _cum *= (1.0 + _y / 100.0)
        _cvs.append((_cum - 1.0) * 100.0)

    # ── Monthly heatmap (12 cols) ─────────────────────────────────────────────
    _CS_MOM = [[0.00,"#7B1D1D"],[0.15,"#C0392B"],[0.30,"#E74C3C"],
               [0.45,"#FADBD8"],[0.50,"#F8F9FA"],
               [0.55,"#D5F5E3"],[0.70,"#27AE60"],[0.85,"#1E8449"],[1.00,"#145A32"]]

    _z_mom  = _piv_num.values.tolist()
    _txt_m  = [[f"{v:+.1f}%" if pd.notna(v) else "" for v in row] for row in _z_mom]
    _y_lbls = [str(y) for y in _piv_num.index.tolist()]

    # ── YTD column ────────────────────────────────────────────────────────────
    _z_ytd  = [[v] for v in _ytd_vals.tolist()]
    _txt_ytd= [[f"{v:+.1f}%" if pd.notna(v) else ""] for v in _ytd_vals.tolist()]

    # ── Cumul column ──────────────────────────────────────────────────────────
    _z_cum  = [[v] for v in _cvs]
    _txt_cum= [[f"{v:+.1f}%"] for v in _cvs]

    _n_rows  = len(_y_lbls)
    _row_h   = 28
    _fig_h   = max(320, _n_rows * _row_h + 80)

    # Create figure with 3 side-by-side subplots (shared y-axis)
    from plotly.subplots import make_subplots
    fig_hm = make_subplots(
        rows=1, cols=3,
        column_widths=[0.78, 0.11, 0.11],
        shared_yaxes=True,
        horizontal_spacing=0.01,
    )

    fig_hm.add_trace(go.Heatmap(
        z=_z_mom, x=MONTHS, y=_y_lbls,
        text=_txt_m, texttemplate="%{text}",
        textfont=dict(size=10, family="JetBrains Mono"),
        colorscale=_CS_MOM, zmid=0, showscale=False,
        hovertemplate="<b>%{y} %{x}</b><br>%{text}<extra></extra>",
        xgap=1, ygap=1,
    ), row=1, col=1)

    fig_hm.add_trace(go.Heatmap(
        z=_z_ytd, x=["YTD %"], y=_y_lbls,
        text=_txt_ytd, texttemplate="%{text}",
        textfont=dict(size=10, family="JetBrains Mono"),
        colorscale=_CS_MOM, zmid=0, showscale=False,
        hovertemplate="<b>%{y} YTD</b><br>%{text}<extra></extra>",
        xgap=1, ygap=1,
    ), row=1, col=2)

    fig_hm.add_trace(go.Heatmap(
        z=_z_cum, x=["Cumul %"], y=_y_lbls,
        text=_txt_cum, texttemplate="%{text}",
        textfont=dict(size=10, family="JetBrains Mono"),
        colorscale=[[0,"#EAE0FF"],[0.3,"#9B59B6"],[0.6,"#6C3483"],[1.0,"#3B1F7A"]],
        showscale=False,
        hovertemplate="<b>%{y} Cumulative</b><br>%{text}<extra></extra>",
        xgap=1, ygap=1,
    ), row=1, col=3)

    fig_hm.update_layout(
        height=_fig_h,
        margin=dict(l=0, r=10, t=36, b=10),
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        font=dict(family="Sora", size=11, color="#0A0A0A"),
        showlegend=False,
    )
    # All x-axes on top, all y-axes reversed
    for _ax in ["xaxis","xaxis2","xaxis3"]:
        fig_hm.update_layout(**{_ax: dict(side="top", tickfont=dict(size=10, family="Sora"),
                                          fixedrange=True, showgrid=False)})
    for _ax in ["yaxis","yaxis2","yaxis3"]:
        fig_hm.update_layout(**{_ax: dict(autorange="reversed", fixedrange=True,
                                          tickfont=dict(size=11, family="JetBrains Mono"),
                                          showgrid=False)})

    st.markdown('<div class="card" style="padding:16px;">', unsafe_allow_html=True)
    st.caption("Jan–Dec: monthly return  ·  YTD: compounded annual return  ·  Cumul: running total from start (purple scale)")
    st.plotly_chart(fig_hm, use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════
# ④ MONTHLY REPORT TABLE + ANNUAL BAR
# ═════════════════════════════════════════════════════════════

st.markdown('<p class="sec">④ Monthly Portfolio Report</p>', unsafe_allow_html=True)

def cpct(v):
    if pd.isna(v): return ""
    return "color:#12B36B;font-weight:600" if v>=0 else "color:#E03535;font-weight:600"

disp = mdf[["Year","Month","Portfolio","Invested","Gain $","Gain %","MoM %"]].copy()
styled = (disp.style
    .format({"Portfolio":"${:,.0f}","Invested":"${:,.0f}",
             "Gain $":"${:+,.0f}","Gain %":"{:+.1f}%","MoM %":"{:+.1f}%"},na_rep="—")
    .applymap(cpct,subset=["Gain %","MoM %"])
    .set_properties(**{"font-family":"JetBrains Mono,monospace","font-size":"13px","color":"#0A0A0A"})
    .set_table_styles([{"selector":"th","props":[("background","#F2F4F8"),("font-size","10px"),
        ("text-transform","uppercase"),("letter-spacing","1px"),("color","#8A95A8"),
        ("font-family","Sora,sans-serif")]}]))
st.dataframe(styled, use_container_width=True, height=400)

yr_df = (mdf.groupby("Year").agg(E=("Portfolio","last"),S=("Portfolio","first"))
            .assign(Ret=lambda d:(d.E/d.S-1)*100).reset_index())
fig_yr = go.Figure(go.Bar(
    x=yr_df["Year"].astype(str), y=yr_df["Ret"],
    marker_color=["#00A878" if r>=0 else "#E03535" for r in yr_df["Ret"]],
    text=yr_df["Ret"].apply(fmt_pct), textposition="outside",
    hovertemplate="<b>%{x}</b> · %{y:.1f}%<extra></extra>"))
fig_yr.update_layout(
    title=dict(text="Annual Portfolio Returns",font=dict(size=15,family="Sora",color="#0A0A0A"),x=0),
    height=280,margin=dict(l=0,r=0,t=36,b=0),
    plot_bgcolor="#FFFFFF",paper_bgcolor="#FFFFFF",
    xaxis=dict(showgrid=False,tickfont=dict(size=12,family="Sora",color="#0A0A0A")),
    yaxis=dict(showgrid=True,gridcolor="#EDE9E0",ticksuffix="%",
               tickfont=dict(size=12,family="JetBrains Mono",color="#0A0A0A")),
    font=dict(family="Sora",size=13,color="#0A0A0A"),bargap=0.28)
st.markdown('<div class="card" style="padding:16px;margin-top:10px;">', unsafe_allow_html=True)
st.plotly_chart(fig_yr, use_container_width=True, config={"displayModeBar":False})
st.markdown("</div>", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════
# ⑤ QUARTERLY HOLDINGS VISUAL
# ═════════════════════════════════════════════════════════════

st.markdown('<p class="sec">⑤ Quarterly Holdings Map</p>', unsafe_allow_html=True)

fqs = {k:v for k,v in qs_data.items()
       if yr_start<=int(k.split("-")[0])<=yr_end}

with st.expander("🗓  Show / Hide Quarterly Holdings Map", expanded=False):
    if fqs:
        all_q_labels  = sorted(fqs.keys())
        all_q_tickers = sorted({t for v in fqs.values() for t in v})

        presence = pd.DataFrame(0, index=all_q_tickers, columns=all_q_labels)
        rank_mat  = pd.DataFrame(0.0, index=all_q_tickers, columns=all_q_labels)
        for ql,tks in fqs.items():
            for rank,t in enumerate(tks,1):
                presence.loc[t,ql] = 1
                rank_mat.loc[t,ql] = rank

        presence = presence.loc[presence.sum(axis=1).sort_values(ascending=True).index]
        rank_mat  = rank_mat.loc[presence.index]
        txt_mat   = rank_mat.applymap(lambda v: str(int(v)) if v>0 else "")

        q1_labels = [q for q in all_q_labels if q.endswith("-Q1")]
        q1_years  = [q.split("-")[0] for q in q1_labels]

        fig_hold = go.Figure(go.Heatmap(
            z=presence.values, x=all_q_labels, y=presence.index.tolist(),
            text=txt_mat.values, texttemplate="%{text}",
            textfont=dict(size=10,family="JetBrains Mono",color="#0A0A0A"),
            colorscale=[[0,"#F0FBF9"],[1,"#007A72"]], showscale=False,
            hovertemplate="<b>%{y}</b> · %{x}<br>Rank #%{text}<extra></extra>",
        ))
        fig_hold.update_layout(
            height=max(300, len(all_q_tickers)*22+80),
            margin=dict(l=0,r=0,t=12,b=0),
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            xaxis=dict(tickangle=-45, tickfont=dict(size=11,family="Sora",color="#0A0A0A"), side="top",
                       fixedrange=True, tickmode="array",
                       tickvals=q1_labels, ticktext=q1_years),
            yaxis=dict(tickfont=dict(size=12,family="JetBrains Mono",color="#0A0A0A"),
                       autorange="reversed", fixedrange=True),
            font=dict(family="Sora",size=13,color="#0A0A0A"))
        st.caption("Blue = in top holdings that quarter · number = market-cap rank · sorted by total appearances")
        st.plotly_chart(fig_hold, use_container_width=True, config={"displayModeBar":False})

        with st.expander("Quarterly snapshots table", expanded=False):
            qrows = [{"Quarter":ql,"# Tickers":len(tks),"Holdings":"  ·  ".join(tks)}
                     for ql,tks in sorted(fqs.items())]
            st.dataframe(pd.DataFrame(qrows)
                         .style.set_properties(**{"font-family":"JetBrains Mono","font-size":"13px","color":"#0A0A0A"}),
                         use_container_width=True, height=400, hide_index=True)
    else:
        st.info("No holdings data for the selected year range.")


# ═════════════════════════════════════════════════════════════
# ⑥ TICKER CONTRIBUTION
# ═════════════════════════════════════════════════════════════

st.markdown('<p class="sec">⑥ Ticker Contribution (Cost Basis)</p>', unsafe_allow_html=True)
if contrib:
    cdf = (pd.DataFrame({"Ticker":k,"Cost Basis":v} for k,v in contrib.items())
           .sort_values("Cost Basis",ascending=False).head(20).reset_index(drop=True))
    fig_c = go.Figure(go.Bar(
        x=cdf["Ticker"], y=cdf["Cost Basis"],
        marker=dict(color=cdf["Cost Basis"],
                    colorscale=[[0,"#C7F2EE"],[0.5,"#00B4A6"],[1,"#007A72"]],showscale=False),
        text=cdf["Cost Basis"].apply(fmt_usd), textposition="outside",
        hovertemplate="<b>%{x}</b><br>$%{y:,.0f}<extra></extra>"))
    fig_c.update_layout(height=340,margin=dict(l=0,r=0,t=12,b=0),
        plot_bgcolor="#FFFFFF",paper_bgcolor="#FFFFFF",
        xaxis=dict(showgrid=False,tickfont=dict(size=13,family="JetBrains Mono",color="#0A0A0A")),
        yaxis=dict(showgrid=True,gridcolor="#EDE9E0",tickformat="$,.0f",
                   tickfont=dict(size=12,family="JetBrains Mono",color="#0A0A0A")),
        font=dict(family="Sora",size=13,color="#0A0A0A"),bargap=0.3)
    cb,ct = st.columns([3,2])
    with cb:
        st.markdown('<div class="card" style="padding:16px;">', unsafe_allow_html=True)
        st.plotly_chart(fig_c, use_container_width=True, config={"displayModeBar":False})
        st.markdown("</div>", unsafe_allow_html=True)
    with ct:
        cdf["% of Total"]  = (cdf["Cost Basis"]/cdf["Cost Basis"].sum()*100).map("{:.1f}%".format)
        cdf["Cost Basis $"]= cdf["Cost Basis"].apply(fmt_usd)
        st.dataframe(cdf[["Ticker","Cost Basis $","% of Total"]]
                     .style.set_properties(**{"font-family":"JetBrains Mono","font-size":"13px","color":"#0A0A0A"}),
                     use_container_width=True, height=360, hide_index=True)


# ═════════════════════════════════════════════════════════════
# ⑥b FINAL VALUE CONTRIBUTION — PIE CHART
# ═════════════════════════════════════════════════════════════

st.markdown('<p class="sec">⑥b Final Value Contribution by Ticker</p>', unsafe_allow_html=True)

if "port_df" in st.session_state and "prices_df" in st.session_state and contrib:
    _prices = st.session_state.get("prices_df", pd.DataFrame())
    _port   = st.session_state["port_df"]

    # Estimate current market value per ticker from latest available price × shares
    # We use cost_basis as proxy weighted by latest price appreciation
    # More accurate: recompute final holdings value from last simulation state
    # For display, apportion final portfolio value by cost_basis weight
    total_cb   = sum(contrib.values())
    final_port = float(_port["Portfolio"].iloc[-1])

    pie_df = (pd.DataFrame({"Ticker":k,"Cost Basis":v} for k,v in contrib.items())
              .sort_values("Cost Basis", ascending=False))
    # Apportion final portfolio value proportionally to cost basis weight
    pie_df["Final Value $"] = pie_df["Cost Basis"] / total_cb * final_port
    pie_df["Pct"]           = pie_df["Final Value $"] / final_port * 100

    # Top 15, bundle rest as "Other"
    top15_pie = pie_df.head(15).copy()
    other_val = pie_df.iloc[15:]["Final Value $"].sum() if len(pie_df) > 15 else 0
    if other_val > 0:
        top15_pie = pd.concat([top15_pie,
                               pd.DataFrame([{"Ticker":"Other","Cost Basis":0,
                                              "Final Value $":other_val,
                                              "Pct":other_val/final_port*100}])],
                              ignore_index=True)

    COLORS = [
        "#5B5FEF","#00B4A6","#FF6B35","#F59E0B","#00A878","#E03535",
        "#EC4899","#7C3AED","#0EA5E9","#14532D","#B45309","#1E3A5F",
        "#6B21A8","#0F766E","#C2410C","#374151",
    ]

    fig_pie = go.Figure(go.Pie(
        labels=top15_pie["Ticker"],
        values=top15_pie["Final Value $"].round(2),
        hole=0.42,
        textinfo="label+percent",
        textfont=dict(size=13, family="JetBrains Mono", color="#0A0A0A"),
        marker=dict(colors=COLORS[:len(top15_pie)],
                    line=dict(color="#FFFFFF", width=2)),
        hovertemplate="<b>%{label}</b><br>Final Value: $%{value:,.0f}<br>Share: %{percent}<extra></extra>",
        pull=[0.03 if i == 0 else 0 for i in range(len(top15_pie))],
    ))
    fig_pie.update_layout(
        height=480,
        margin=dict(l=0, r=0, t=20, b=0),
        paper_bgcolor="#FFFFFF",
        showlegend=True,
        legend=dict(orientation="v", x=1.02, y=0.5,
                    font=dict(size=11, family="JetBrains Mono"),
                    bgcolor="rgba(0,0,0,0)"),
        annotations=[dict(
            text=f"<b>{fmt_usd(final_port)}</b><br><span style='font-size:13px'>Final Value</span>",
            x=0.5, y=0.5, font=dict(size=15, family="Sora", color="#0A0A0A"),
            showarrow=False,
        )],
        font=dict(family="Sora",size=13,color="#0A0A0A"),
    )

    pc1, pc2 = st.columns([3, 2])
    with pc1:
        st.markdown('<div class="card" style="padding:16px;">', unsafe_allow_html=True)
        st.caption("Final portfolio value apportioned by each ticker's cost-basis weight")
        st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})
        st.markdown("</div>", unsafe_allow_html=True)
    with pc2:
        disp_pie = top15_pie[["Ticker","Final Value $","Pct"]].copy()
        disp_pie["Final Value"] = disp_pie["Final Value $"].apply(fmt_usd)
        disp_pie["Share %"]     = disp_pie["Pct"].map("{:.1f}%".format)
        st.dataframe(
            disp_pie[["Ticker","Final Value","Share %"]]
            .style.set_properties(**{"font-family":"JetBrains Mono","font-size":"13px","color":"#0A0A0A"})
            .set_table_styles([{"selector":"th","props":[("background","#F2F4F8"),
                ("font-size","10px"),("text-transform","uppercase"),("letter-spacing","1px"),
                ("color","#8A95A8"),("font-family","Sora,sans-serif")]}]),
            use_container_width=True, height=450, hide_index=True,
        )


# ═════════════════════════════════════════════════════════════
# ⑦ TAX REPORT — REALISED SELLS
# ═════════════════════════════════════════════════════════════

st.markdown('<p class="sec">⑦ Tax Report — Realised Sells</p>', unsafe_allow_html=True)

if sell_log:
    tax_df = pd.DataFrame(sell_log)
    tax_df = tax_df[(tax_df["Year"]>=yr_start)&(tax_df["Year"]<=yr_end)]

    if not tax_df.empty:
        # Annual bar chart
        yr_tax = (tax_df.groupby("Year")
                  .agg(Proceeds=("Proceeds","sum"),
                       CostBasis=("Cost Basis","sum"),
                       GainLoss=("Gain / Loss","sum")).reset_index())
        fig_tax = go.Figure()
        fig_tax.add_trace(go.Bar(name="Proceeds",
            x=yr_tax["Year"].astype(str),y=yr_tax["Proceeds"],marker_color="#5B5FEF",
            hovertemplate="<b>%{x}</b><br>Proceeds: $%{y:,.0f}<extra></extra>"))
        fig_tax.add_trace(go.Bar(name="Cost Basis",
            x=yr_tax["Year"].astype(str),y=yr_tax["CostBasis"],marker_color="#CCCCCC",
            hovertemplate="<b>%{x}</b><br>Cost Basis: $%{y:,.0f}<extra></extra>"))
        fig_tax.add_trace(go.Scatter(name="Net Gain / Loss",
            x=yr_tax["Year"].astype(str),y=yr_tax["GainLoss"],
            mode="lines+markers",line=dict(color="#FF6B35",width=2.5),
            marker=dict(size=7,color=["#00A878"if v>=0 else "#E03535" for v in yr_tax["GainLoss"]]),
            hovertemplate="<b>%{x}</b><br>Net G/L: $%{y:,.0f}<extra></extra>"))
        fig_tax.update_layout(barmode="group",height=320,margin=dict(l=0,r=0,t=12,b=0),
            plot_bgcolor="#FFFFFF",paper_bgcolor="#FFFFFF",
            legend=dict(orientation="h",y=1.02,x=0,font=dict(size=13,family="Sora",color="#0A0A0A")),
            xaxis=dict(showgrid=False,tickfont=dict(size=12,family="Sora",color="#0A0A0A")),
            yaxis=dict(showgrid=True,gridcolor="#EDE9E0",tickformat="$,.0f",
                       tickfont=dict(size=12,family="JetBrains Mono",color="#0A0A0A")),
            font=dict(family="Sora",size=13,color="#0A0A0A"))
        st.markdown('<div class="card" style="padding:16px;">', unsafe_allow_html=True)
        st.caption("Grouped bars = Proceeds vs Cost Basis each year  ·  line = net realised gain/loss")
        st.plotly_chart(fig_tax, use_container_width=True, config={"displayModeBar":False})
        st.markdown("</div>", unsafe_allow_html=True)

        # Summary KPIs
        total_proceeds = yr_tax["Proceeds"].sum()
        total_cost     = yr_tax["CostBasis"].sum()
        total_gain     = yr_tax["GainLoss"].sum()
        best_yr        = yr_tax.loc[yr_tax["GainLoss"].idxmax()]
        worst_yr       = yr_tax.loc[yr_tax["GainLoss"].idxmin()]
        ta,tb,tc,td,te = st.columns(5)
        for col,lbl,val,cls in [
            (ta,"Total Proceeds",  fmt_usd(total_proceeds),"b"),
            (tb,"Total Cost Basis",fmt_usd(total_cost),"gr"),
            (tc,"Net Realised G/L",fmt_usd(total_gain),"g"if total_gain>=0 else "r"),
            (td,f"Best Year ({int(best_yr['Year'])})",fmt_usd(best_yr['GainLoss']),"g"),
            (te,f"Worst Year ({int(worst_yr['Year'])})",fmt_usd(worst_yr['GainLoss']),"r"),
        ]:
            with col:
                st.markdown(f'<div class="kpi"><div class="kpi-lbl">{lbl}</div>'
                            f'<div class="kpi-val {cls}" style="font-size:20px;">{val}</div></div>',
                            unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # Detailed table
        def cgl(v):
            if pd.isna(v): return ""
            return "color:#12B36B;font-weight:600" if v>=0 else "color:#E03535;font-weight:600"

        tax_styled = (tax_df.sort_values(["Year","Quarter","Ticker"]).style
            .format({"Shares":"{:.4f}","Price":"${:,.2f}","Proceeds":"${:,.2f}",
                     "Cost Basis":"${:,.2f}","Gain / Loss":"${:+,.2f}"})
            .applymap(cgl,subset=["Gain / Loss"])
            .set_properties(**{"font-family":"JetBrains Mono,monospace","font-size":"13px","color":"#0A0A0A"})
            .set_table_styles([{"selector":"th","props":[("background","#F2F4F8"),
                ("font-size","10px"),("text-transform","uppercase"),("letter-spacing","1px"),
                ("color","#8A95A8"),("font-family","Sora,sans-serif")]}]))
        st.dataframe(tax_styled, use_container_width=True, height=420, hide_index=True)
        st.download_button("⬇  Download Tax Report CSV",
                           data=tax_df.to_csv(index=False),
                           file_name=f"spx15_tax_report_{yr_start}_{yr_end}.csv",
                           mime="text/csv")
    else:
        st.info("No sell transactions in the selected year range.")
else:
    st.info("Run the simulation to generate the tax report.")


# ═════════════════════════════════════════════════════════════
# ⑧ BENCHMARK COMPARISON
# ═════════════════════════════════════════════════════════════

st.markdown('<p class="sec">⑧ Strategy vs Benchmark Comparison</p>', unsafe_allow_html=True)

def stats(label, s, inv_s):
    v1=float(s.iloc[-1]); v0=float(s.iloc[0]); inv=float(inv_s.iloc[-1])
    yrs=(s.index[-1]-s.index[0]).days/365.25; ch=s.pct_change().dropna()
    tot_ret = (v1-inv)/inv*100 if inv > 0 else 0.0
    _cagr   = cagr(v0,v1,yrs)*100 if v0 > 0 else 0.0
    _mdd    = max_dd(s)*100 if v0 > 0 else 0.0
    _best   = float(ch.max()*100) if not ch.empty else 0.0
    _worst  = float(ch.min()*100) if not ch.empty else 0.0
    return {"Strategy":label,"Final Value":v1,"Invested":inv,
            "Total Return %":tot_ret,"CAGR %":_cagr,
            "Max Drawdown %":_mdd,
            "Best Month %":_best,"Worst Month %":_worst}

_eq_label  = f"EquityGrowth ({_eq_pct_ss}%)"
_oi_label  = f"OptionIncome ({_oi_pct_ss}% · {_oi_cagr_ss:.0f}% CAGR)"
_tot_label = "Total Combined (EQ + OI)"
rows_c = [stats(_eq_label, port_df["Portfolio"], port_df["Invested"])]
if not oi_df.empty and _oi_pct_ss > 0:
    rows_c.append(stats(_oi_label, oi_df["OI_Portfolio"], oi_df["OI_Invested"]))
_comb_s     = port_df["Portfolio"] + (oi_df["OI_Portfolio"] if not oi_df.empty else 0)
_comb_inv_s = port_df["Invested"]  + (oi_df["OI_Invested"]  if not oi_df.empty else 0)
rows_c.append(stats(_tot_label, _comb_s, _comb_inv_s))
if not spy_df.empty: rows_c.append(stats("SPY (S&P 500 ETF)", spy_df["Benchmark"], spy_df["Invested"]))
if not qqq_df.empty: rows_c.append(stats("QQQ (Nasdaq-100 ETF)", qqq_df["Benchmark"], qqq_df["Invested"]))
cmp = pd.DataFrame(rows_c)
st.dataframe(cmp.style
    .format({"Final Value":"${:,.0f}","Invested":"${:,.0f}",
             "Total Return %":"{:+.1f}%","CAGR %":"{:+.2f}%",
             "Max Drawdown %":"{:.1f}%","Best Month %":"{:+.1f}%","Worst Month %":"{:+.1f}%"})
    .highlight_max(subset=["Final Value","Total Return %","CAGR %","Best Month %"],color="#C7F2EE")
    .highlight_min(subset=["Max Drawdown %","Worst Month %"],color="#C7F2EE")
    .apply(lambda row: ["background-color:#E8F4FD;font-weight:700"]*len(row)
           if row.get("Strategy","") == _tot_label else [""]*len(row), axis=1)
    .set_properties(**{"font-family":"JetBrains Mono,monospace","font-size":"13px","color":"#0A0A0A"})
    .set_table_styles([{"selector":"th","props":[("background","#F2F4F8"),("font-size","10px"),
        ("text-transform","uppercase"),("letter-spacing","1px"),("color","#8A95A8"),
        ("font-family","Sora,sans-serif")]}]),
    use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════════════════
# ⑨ FEES & HEDGE COST REPORT
# ═══════════════════════════════════════════════════════════

st.markdown('<p class="sec">⑨ Fees & Hedge Cost Report</p>', unsafe_allow_html=True)

if fee_log:
    fl_df  = pd.DataFrame(fee_log)
    fl_fil = fl_df[(fl_df["Year"]>=yr_start)&(fl_df["Year"]<=yr_end)].copy()
    if not fl_fil.empty:
        yr_f = (fl_fil.groupby("Year")
                .agg(HC=("Hedge Cost","sum"),MF=("Mgmt Fee","sum"),
                     PF=("Perf Fee","sum"),TD=("Total Deducted","sum"),
                     AA=("Combined AUM","mean")).reset_index())
        fc1,fc2,fc3,fc4,fc5 = st.columns(5)
        _fin = float(port_df["Portfolio"].iloc[-1])
        for _col,_lbl,_val,_cls in [
            (fc1,"Total Hedge",   fmt_usd(fl_fil["Hedge Cost"].sum()),"r"),
            (fc2,"Total Mgmt",    fmt_usd(fl_fil["Mgmt Fee"].sum()),"r"),
            (fc3,"Total Perf",    fmt_usd(fl_fil["Perf Fee"].sum()),"r"),
            (fc4,"All Costs",     fmt_usd(fl_fil["Total Deducted"].sum()),"r"),
            (fc5,"Cost % Gross",  f'{fl_fil["Total Deducted"].sum()/(fl_fil["Total Deducted"].sum()+_fin)*100:.1f}%',"gr"),
        ]:
            with _col:
                st.markdown(f'<div class="kpi"><div class="kpi-lbl">{_lbl}</div>'
                            f'<div class="kpi-val {_cls}" style="font-size:18px;">{_val}</div></div>',
                            unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        fig_fee = go.Figure()
        fig_fee.add_trace(go.Bar(name="Hedge Cost",x=yr_f["Year"].astype(str),y=yr_f["HC"],
            marker_color="#E03535",hovertemplate="<b>%{x}</b><br>Hedge: $%{y:,.0f}<extra></extra>"))
        fig_fee.add_trace(go.Bar(name="Mgmt Fee",x=yr_f["Year"].astype(str),y=yr_f["MF"],
            marker_color="#F27820",hovertemplate="<b>%{x}</b><br>Mgmt: $%{y:,.0f}<extra></extra>"))
        fig_fee.add_trace(go.Bar(name="Perf Fee",x=yr_f["Year"].astype(str),y=yr_f["PF"],
            marker_color="#9333EA",hovertemplate="<b>%{x}</b><br>Perf: $%{y:,.0f}<extra></extra>"))
        fig_fee.add_trace(go.Scatter(name="Avg AUM",x=yr_f["Year"].astype(str),y=yr_f["AA"],
            mode="lines+markers",yaxis="y2",line=dict(color="#0A4DA8",width=2,dash="dot"),
            marker=dict(size=6),hovertemplate="<b>%{x}</b><br>Avg AUM: $%{y:,.0f}<extra></extra>"))
        fig_fee.update_layout(barmode="stack",height=360,margin=dict(l=0,r=70,t=12,b=0),
            plot_bgcolor="#FAFBFC",paper_bgcolor="white",
            legend=dict(orientation="h",y=1.02,x=0,font=dict(size=11,family="Sora")),
            xaxis=dict(showgrid=False,tickfont=dict(size=10,family="Sora")),
            yaxis=dict(title="Annual Cost ($)",showgrid=True,gridcolor="#EAEDF2",
                       tickformat="$,.0f",tickfont=dict(size=10,family="JetBrains Mono")),
            yaxis2=dict(title="Avg AUM ($)",overlaying="y",side="right",showgrid=False,
                        tickformat="$,.0f",tickfont=dict(size=10,family="JetBrains Mono")),
            font=dict(family="Sora"))
        st.markdown('<div class="card" style="padding:16px;">', unsafe_allow_html=True)
        st.caption("Stacked = annual costs · dotted line = average AUM (right axis)")
        st.plotly_chart(fig_fee, use_container_width=True, config={"displayModeBar":False})
        st.markdown("</div>", unsafe_allow_html=True)
        yr_d = yr_f.copy()
        yr_d.columns=["Year","Hedge Cost","Mgmt Fee","Perf Fee","Total Deducted","Avg Combined AUM"]
        for _c in ["Hedge Cost","Mgmt Fee","Perf Fee","Total Deducted","Avg Combined AUM"]:
            yr_d[_c] = yr_d[_c].apply(fmt_usd)
        st.dataframe(yr_d.style.set_properties(**{"font-family":"JetBrains Mono","font-size":"12px"})
            .set_table_styles([{"selector":"th","props":[("background","#F2F4F8"),("font-size","10px"),
                ("text-transform","uppercase"),("letter-spacing","1px"),("color","#8A95A8"),
                ("font-family","Sora,sans-serif")]}]),
            use_container_width=True, height=380, hide_index=True)
        with st.expander("📋 Monthly fee detail", expanded=False):
            fl_d2 = fl_fil.copy()
            def _hi(v): return "background-color:#FFF3CD;font-weight:600" if str(v) not in ("$0","") else ""
            for _c in ["Equity AUM","OI AUM","Combined AUM","Hedge Cost","Mgmt Fee","Perf Fee","Total Deducted"]:
                if _c in fl_d2.columns:
                    fl_d2[_c] = fl_d2[_c].apply(fmt_usd)
            st.dataframe(fl_d2.style.applymap(_hi,subset=["Mgmt Fee","Perf Fee"])
                .set_properties(**{"font-family":"JetBrains Mono","font-size":"12px"})
                .set_table_styles([{"selector":"th","props":[("background","#F2F4F8"),("font-size","10px"),
                    ("text-transform","uppercase"),("letter-spacing","1px"),("color","#8A95A8"),
                    ("font-family","Sora,sans-serif")]}]),
                use_container_width=True, height=420, hide_index=True)
        st.download_button("⬇  Download Fee Report CSV",
            data=fl_fil.to_csv(index=False),
            file_name=f"fee_report_{yr_start}_{yr_end}.csv", mime="text/csv")
    else:
        st.info("No fee data for the selected year range.")
else:
    st.info("Run the simulation to generate fee reporting.")


st.markdown("---")
st.markdown('<p style="text-align:center;color:#555555;font-size:13px;font-family:Sora,sans-serif;">'
            'Portfolio Simulator · Market data via yfinance · '
            'Past performance is not indicative of future results · For informational purposes only</p>',
            unsafe_allow_html=True)
