
import base64
import io
import json
import math
import os
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from PIL import Image


# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="Otogard | Modern Trade Sales Monitoring",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_DB = APP_DIR / "MT - Daily Sales Database(1).xlsx"
LOGO_PATH = APP_DIR / "otogard_logo.png"

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_NUM = {m: i + 1 for i, m in enumerate(MONTH_ORDER)}

CUSTOMER_FIXED = [
    ("MENSA BINASUKSES", "Mensa Binasukses"),
    ("KEBAYORAN PHARMA", "Kebayoran Pharma"),
    ("PRIMA MEDITAMA", "Prima Meditama"),
    ("VARIA INDAH PARAMITA", "Varia Indah Paramita"),
    ("PANCA PILAR PERKASA", "Panca Pilar Perkasa Jaya"),
    ("HARI-HARI GROUP", "Hari-Hari"),
    ("HYPERMART", "Hypermart"),
    ("KEDAI MART", "Kedai Mart"),
    ("SAGA SUPERMARKET", "Saga Supermarket"),
]

ACCOUNT_FIXED = [
    ("INDOMARET", "Indomaret"),
    ("ALFAMART", "Alfamart"),
    ("ALFAMIDI", "Alfamidi"),
    ("LION SUPERINDO", "Superindo"),
    ("__MTI__", "MTI"),
]

PRODUCT_DISPLAY = {
    "GLO WASH & SHINE POUCH 700ML": "GLO Wash & Shine",
    "GLO WASH & WAX POUCH 720ML": "GLO Wash & Wax",
    "GLO HYPER BLACK TIRE SHINE 165ML": "Hyper Black Tire Shine",
    "GLO CRYSTAL CLEAR WIPER FLUID 400ML": "Wiper Fluid 400 ml",
    "EVO INJECTOR CLEANER 80ML": "EVO Injector Cleaner",
    "EVO ENGINE FLUSH 80ML": "EVO Engine Flush",
}


# =========================================================
# STYLE
# =========================================================
st.markdown(
    """
    <style>
    :root {
        --bg:#031124;
        --bg2:#061a33;
        --card:#07284a;
        --card2:#0a3158;
        --cyan:#10c9ff;
        --blue:#1677ff;
        --green:#23e6b1;
        --yellow:#ffd75e;
        --orange:#ff944d;
        --red:#ff575f;
        --purple:#8c66ff;
        --text:#f7fbff;
        --muted:#a9bfd6;
        --stroke:rgba(16,201,255,.80);
    }

    html, body, [class*="css"] {
        font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont,
                     "Segoe UI", sans-serif;
    }

    .stApp {
        background:
          radial-gradient(circle at 50% -15%, rgba(18,100,190,.22), transparent 35%),
          linear-gradient(180deg, #031124 0%, #04162a 50%, #031124 100%);
        color: var(--text);
    }

    .block-container {
        max-width: 900px;
        padding-top: 14px;
        padding-bottom: 90px;
        padding-left: 14px;
        padding-right: 14px;
    }

    /* hide Streamlit chrome */
    #MainMenu, footer, header {visibility:hidden;}

    .headline-wrap {
        display:grid;
        grid-template-columns: 190px 1fr auto;
        align-items:center;
        gap:18px;
        margin-bottom:14px;
    }
    .brand-logo img {
        width:100%;
        max-width:190px;
        max-height:78px;
        object-fit:contain;
    }
    .headline-copy {
        border-left:1px solid rgba(255,255,255,.4);
        padding-left:20px;
        min-width:0;
    }
    .headline-main {
        font-size:32px;
        line-height:1.0;
        font-weight:800;
        letter-spacing:.02em;
        color:white;
    }
    .headline-sub {
        font-size:27px;
        line-height:1.1;
        font-weight:300;
        color:#e2eef8;
        margin-top:4px;
    }
    .data-badge {
        border:1px solid rgba(16,201,255,.9);
        border-radius:10px;
        padding:8px 12px;
        background:linear-gradient(180deg, rgba(10,49,88,.94), rgba(4,25,48,.9));
        min-width:120px;
        box-shadow:0 0 18px rgba(16,201,255,.10);
        text-align:center;
    }
    .data-badge .tiny {font-size:11px;color:var(--muted);}
    .data-badge .big {font-size:15px;font-weight:700;color:#fff;}

    div[data-testid="stSelectbox"] > div {
        background:linear-gradient(180deg, rgba(9,45,82,.95), rgba(5,30,57,.96));
        border-radius:10px;
    }
    div[data-testid="stSelectbox"] label {
        color:#d9e9f6 !important;
        font-size:13px !important;
        font-weight:650 !important;
    }

    .card {
        border:1px solid var(--stroke);
        border-radius:12px;
        background:
          linear-gradient(135deg, rgba(14,57,99,.98), rgba(4,26,50,.98));
        box-shadow:
          inset 0 0 26px rgba(14,115,190,.08),
          0 0 18px rgba(0,168,255,.05);
        overflow:hidden;
    }

    .hero-card {
        min-height:338px;
        padding:18px 20px;
        background:
          radial-gradient(circle at 80% 5%, rgba(16,201,255,.24), transparent 34%),
          linear-gradient(135deg, rgba(8,76,145,.97), rgba(3,28,55,.98));
    }
    .side-card {
        padding:15px 16px;
        min-height:162px;
    }
    .flow-card {
        padding:15px 16px;
        min-height:162px;
    }
    .section-card {
        padding:15px 16px;
        margin-top:16px;
    }
    .section-title {
        font-size:20px;
        font-weight:750;
        line-height:1.1;
        color:#fff;
        margin:0;
    }
    .section-title-row {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:10px;
        margin-bottom:10px;
    }
    .period-label {
        color:#d6e7f6;
        font-size:13px;
        white-space:nowrap;
    }
    .muted {color:var(--muted);}
    .hero-value {
        font-size:50px;
        line-height:1.02;
        font-weight:800;
        color:#fff;
        margin:24px 0 10px;
        letter-spacing:-.02em;
    }
    .hero-growth {
        font-size:34px;
        line-height:1;
        font-weight:800;
        color:var(--green);
        display:flex;
        align-items:center;
        gap:12px;
    }
    .hero-growth span.note {
        color:#d8e5f1;
        font-size:14px;
        font-weight:450;
        line-height:1.25;
    }
    .date-line {
        font-size:14px;
        color:#e2edf7;
        margin-top:4px;
    }
    .kpi-big {
        font-size:39px;
        font-weight:800;
        line-height:1;
        margin:8px 0 6px;
    }
    .kpi-grid3, .kpi-grid4, .flow-grid {
        display:grid;
        gap:0;
        margin-top:11px;
        border:1px solid rgba(62,139,208,.38);
        border-radius:9px;
        background:rgba(2,23,44,.25);
    }
    .kpi-grid3 {grid-template-columns:repeat(3,1fr);}
    .kpi-grid4 {grid-template-columns:repeat(4,1fr);}
    .flow-grid {grid-template-columns:repeat(2,1fr);}
    .kpi-cell, .flow-cell {
        padding:10px 12px;
        border-right:1px solid rgba(142,191,230,.32);
    }
    .kpi-cell:last-child, .flow-cell:last-child {border-right:0;}
    .kpi-label {font-size:12px;color:#dce8f2;}
    .kpi-value {font-size:18px;font-weight:750;color:#fff;margin-top:2px;}
    .kpi-value.red {color:var(--red);}
    .kpi-value.green {color:var(--green);}
    .flow-title {font-size:13px;color:#dce8f2;}
    .flow-value {font-size:21px;font-weight:750;margin-top:3px;}
    .flow-share {font-size:15px;font-weight:700;margin-top:7px;color:#fff;}
    .mini-bar {
        height:8px;
        border-radius:999px;
        background:#173b60;
        overflow:hidden;
        margin-top:5px;
    }
    .mini-fill-blue {height:100%;background:linear-gradient(90deg,#108cff,#18c9ff);}
    .mini-fill-green {height:100%;background:linear-gradient(90deg,#16c99b,#27efbb);}

    .ytd-card {padding:14px 16px;}
    .ytd-metrics {
        display:grid;
        grid-template-columns:repeat(4,1fr);
        margin-top:9px;
        border-top:1px solid rgba(87,150,204,.22);
    }
    .ytd-item {
        padding:10px 14px 2px;
        border-right:1px solid rgba(142,191,230,.35);
    }
    .ytd-item:first-child {padding-left:4px;}
    .ytd-item:last-child {border-right:0;}
    .ytd-label {font-size:12px;color:#d7e6f3;}
    .ytd-value {font-size:21px;font-weight:750;margin-top:3px;white-space:nowrap;}

    .table-head, .bar-row {
        display:grid;
        grid-template-columns: 2.2fr 3.2fr 1.55fr .7fr;
        align-items:center;
        gap:8px;
    }
    .table-head {
        color:#dceaf6;
        font-size:12px;
        border-bottom:1px solid rgba(108,173,225,.28);
        padding:0 0 6px;
        margin-bottom:5px;
    }
    .bar-row {min-height:28px;font-size:13px;}
    .bar-name {
        color:#f3f8fc;
        overflow:hidden;
        text-overflow:ellipsis;
        white-space:nowrap;
    }
    .bar-track {
        height:14px;
        border-radius:4px;
        background:#173759;
        overflow:hidden;
    }
    .bar-fill {
        height:100%;
        border-radius:4px;
        background:linear-gradient(90deg,#1497ff,#22c8ff);
    }
    .bar-fill.g2 {background:linear-gradient(90deg,#16c99b,#31ebbc);}
    .bar-fill.g3 {background:linear-gradient(90deg,#ffd65a,#ffca42);}
    .bar-fill.g4 {background:linear-gradient(90deg,#ff9854,#ff8444);}
    .bar-fill.g5 {background:linear-gradient(90deg,#ff6768,#ff4e59);}
    .bar-fill.g6 {background:linear-gradient(90deg,#ae6fff,#9458f5);}
    .bar-amount {text-align:right;color:#f1f7fc;white-space:nowrap;}
    .bar-share {text-align:right;color:#f1f7fc;white-space:nowrap;}

    .product-head, .product-row {
        display:grid;
        grid-template-columns: 2.0fr 3.2fr 1.55fr;
        align-items:center;
        gap:8px;
    }
    .product-row {min-height:29px;font-size:13px;}

    .account-head, .account-row {
        display:grid;
        grid-template-columns: 1.5fr 3.3fr 1.55fr;
        align-items:center;
        gap:8px;
    }
    .account-row {min-height:30px;font-size:13px;}

    .insight-grid {
        display:grid;
        grid-template-columns:repeat(4,1fr);
        gap:8px;
        margin-top:10px;
    }
    .insight {
        border:1px solid rgba(54,147,217,.55);
        background:rgba(8,48,86,.82);
        border-radius:9px;
        padding:11px;
        min-height:98px;
        display:flex;
        gap:8px;
        align-items:flex-start;
        font-size:12px;
        line-height:1.35;
    }
    .insight-no {
        width:27px;height:27px;min-width:27px;
        border-radius:50%;
        background:#96ccff;
        color:#07264a;
        display:flex;align-items:center;justify-content:center;
        font-weight:800;font-size:14px;
    }

    .admin-fab {
        position:fixed;
        right:18px;
        bottom:18px;
        z-index:999;
    }


    /* Compact segmented controls */
    div[data-testid="stSegmentedControl"] {
        margin-top: 6px;
        margin-bottom: 2px;
    }
    div[data-testid="stSegmentedControl"] > div {
        justify-content:flex-end;
    }
    div[data-testid="stSegmentedControl"] button {
        min-height:34px !important;
        border-color:rgba(62,139,208,.55) !important;
        background:rgba(6,35,65,.85) !important;
        color:#dcebf7 !important;
        font-size:12px !important;
        padding:5px 14px !important;
    }
    div[data-testid="stSegmentedControl"] button[aria-pressed="true"] {
        background:linear-gradient(180deg,#1e7cff,#1767e5) !important;
        color:white !important;
        box-shadow:0 0 12px rgba(30,124,255,.28);
    }

    .admin-note {
        border:1px solid rgba(16,201,255,.55);
        border-radius:10px;
        padding:12px;
        background:rgba(4,28,53,.95);
        color:#dbe9f5;
        font-size:13px;
    }

    .divider-space {height:16px;}

    @media (max-width: 720px) {
        .block-container {padding-left:10px;padding-right:10px;padding-top:8px;}
        .headline-wrap {
            grid-template-columns: 105px 1fr 90px;
            gap:9px;
            margin-bottom:10px;
        }
        .brand-logo img {max-width:105px;max-height:54px;}
        .headline-copy {padding-left:9px;}
        .headline-main {font-size:18px;}
        .headline-sub {font-size:17px;}
        .data-badge {min-width:0;padding:6px 6px;}
        .data-badge .tiny {font-size:9px;}
        .data-badge .big {font-size:11px;}
        .hero-card {min-height:225px;padding:15px;}
        .hero-value {font-size:42px;}
        .hero-growth {font-size:29px;}
        .side-card,.flow-card {min-height:145px;padding:12px;}
        .section-title {font-size:17px;}
        .kpi-big {font-size:33px;}
        .kpi-grid3 {margin-top:8px;}
        .kpi-cell,.flow-cell {padding:8px;}
        .kpi-label,.flow-title {font-size:10px;}
        .kpi-value {font-size:14px;}
        .flow-value {font-size:17px;}
        .ytd-item {padding:9px 8px 2px;}
        .ytd-label {font-size:10px;}
        .ytd-value {font-size:15px;white-space:nowrap;}
        .table-head, .bar-row {grid-template-columns: 2.05fr 2.7fr 1.45fr .6fr; gap:5px;}
        .product-head, .product-row {grid-template-columns: 1.8fr 2.5fr 1.25fr; gap:5px;}
        .account-head, .account-row {grid-template-columns: 1.3fr 2.6fr 1.3fr; gap:5px;}
        .bar-row,.product-row,.account-row {font-size:11px;}
        .bar-track {height:12px;}
        .insight-grid {grid-template-columns:repeat(2,1fr);}
        .insight {min-height:78px;font-size:11px;}
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# HELPERS
# =========================================================
def fmt_rp(v):
    v = 0 if pd.isna(v) else float(v)
    av = abs(v)
    sign = "-" if v < 0 else ""
    if av >= 1_000_000_000:
        return f"{sign}Rp {av/1_000_000_000:.2f} B"
    if av >= 1_000_000:
        return f"{sign}Rp {av/1_000_000:.1f} M"
    if av >= 1_000:
        return f"{sign}Rp {av/1_000:.1f} K"
    return f"{sign}Rp {av:,.0f}"


def fmt_qty(v):
    v = 0 if pd.isna(v) else float(v)
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.2f} M pcs"
    if abs(v) >= 1_000:
        return f"{v/1_000:.1f} K pcs"
    return f"{v:,.0f} pcs"


def month_label(month, year):
    return f"{month} {year}"


def safe_pct(num, den):
    if den is None or den == 0 or pd.isna(den):
        return 0.0
    return float(num) / float(den) * 100.0


def read_logo_b64():
    if not LOGO_PATH.exists():
        return ""
    return base64.b64encode(LOGO_PATH.read_bytes()).decode("utf-8")


def normalise_sales(df):
    df = df.copy()
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df.columns = [str(c).strip().upper() for c in df.columns]
    required = {
        "YEAR", "MONTH", "DATE", "CATEGORY", "DISTRIBUTION TYPE",
        "ACCOUNT TYPE", "CUSTOMER CHAIN", "SKU NAME",
        "SALES VALUE", "SALES QUANTITY"
    }
    missing = required - set(df.columns)
    if missing:
        raise ValueError("Missing columns in Sales Database: " + ", ".join(sorted(missing)))

    for col in ["CATEGORY", "DISTRIBUTION TYPE", "ACCOUNT TYPE",
                "CUSTOMER CHAIN", "SKU NAME", "MONTH"]:
        df[col] = df[col].astype(str).str.strip().str.upper()

    # Restore expected month casing after normalization
    reverse_month = {m.upper(): m for m in MONTH_ORDER}
    df["MONTH"] = df["MONTH"].map(reverse_month).fillna(df["MONTH"].str.title())

    df["YEAR"] = pd.to_numeric(df["YEAR"], errors="coerce").astype("Int64")
    df["DATE"] = pd.to_numeric(df["DATE"], errors="coerce").fillna(1).astype(int)
    df["SALES VALUE"] = pd.to_numeric(df["SALES VALUE"], errors="coerce").fillna(0.0)
    df["SALES QUANTITY"] = pd.to_numeric(df["SALES QUANTITY"], errors="coerce").fillna(0.0)

    df["MONTH NUM"] = df["MONTH"].map(MONTH_NUM)
    df["FULL DATE"] = pd.to_datetime(
        dict(
            year=df["YEAR"].astype("Int64"),
            month=df["MONTH NUM"],
            day=df["DATE"].clip(1, 31)
        ),
        errors="coerce",
    )
    return df


def normalise_target(df):
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    if not {"MONTH", "SALES TARGET"}.issubset(df.columns):
        raise ValueError("MONTHLY TARGET must contain 'Month' and 'Sales Target'.")
    df["MONTH"] = pd.to_datetime(df["MONTH"], errors="coerce")
    df["SALES TARGET"] = pd.to_numeric(df["SALES TARGET"], errors="coerce").fillna(0.0)
    df["YEAR"] = df["MONTH"].dt.year
    df["MONTH NUM"] = df["MONTH"].dt.month
    df["MONTH LABEL"] = df["MONTH"].dt.strftime("%b")
    return df


def validate_workbook(file_like):
    try:
        xls = pd.ExcelFile(file_like)
        if "Sales Database" not in xls.sheet_names:
            return False, "Sheet 'Sales Database' tidak ditemukan."
        if "MONTHLY TARGET" not in xls.sheet_names:
            return False, "Sheet 'MONTHLY TARGET' tidak ditemukan."
        sales = pd.read_excel(xls, "Sales Database", header=1)
        target = pd.read_excel(xls, "MONTHLY TARGET")
        sales = normalise_sales(sales)
        target = normalise_target(target)
        if sales.empty:
            return False, "Sales Database kosong."
        return True, (sales, target)
    except Exception as e:
        return False, str(e)


def load_db_bytes():
    # Uploaded file in this browser session has priority.
    if "uploaded_db_bytes" in st.session_state:
        return st.session_state["uploaded_db_bytes"]
    if DEFAULT_DB.exists():
        return DEFAULT_DB.read_bytes()
    raise FileNotFoundError(f"Database not found: {DEFAULT_DB}")


@st.cache_data(show_spinner=False)
def parse_db_bytes(raw_bytes):
    bio = io.BytesIO(raw_bytes)
    xls = pd.ExcelFile(bio)
    sales = pd.read_excel(xls, "Sales Database", header=1)
    target = pd.read_excel(xls, "MONTHLY TARGET")
    return normalise_sales(sales), normalise_target(target)


def github_persist(raw_bytes):
    """
    Optional persistent updater.
    Configure Streamlit secrets:
      GITHUB_TOKEN
      REPO            e.g. username/repository
      BRANCH          e.g. main
      FILE_PATH       e.g. "MT - Daily Sales Database(1).xlsx"
    """
    required = ["GITHUB_TOKEN", "REPO", "BRANCH", "FILE_PATH"]
    if not all(k in st.secrets for k in required):
        return False, "GitHub persistence belum dikonfigurasi. Data aktif untuk session ini saja."

    token = st.secrets["GITHUB_TOKEN"]
    repo = st.secrets["REPO"]
    branch = st.secrets["BRANCH"]
    file_path = st.secrets["FILE_PATH"]

    api = f"https://api.github.com/repos/{repo}/contents/{file_path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    get_resp = requests.get(api, headers=headers, params={"ref": branch}, timeout=30)
    sha = None
    if get_resp.status_code == 200:
        sha = get_resp.json().get("sha")
    elif get_resp.status_code not in (404,):
        return False, f"GitHub read failed: {get_resp.status_code} {get_resp.text[:160]}"

    payload = {
        "message": "Update Modern Trade daily sales database from Streamlit Admin Mode",
        "content": base64.b64encode(raw_bytes).decode("utf-8"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    put_resp = requests.put(api, headers=headers, json=payload, timeout=45)
    if put_resp.status_code not in (200, 201):
        return False, f"GitHub update failed: {put_resp.status_code} {put_resp.text[:180]}"
    return True, "Database berhasil disimpan ke GitHub."


def filter_account(df, selected):
    if selected == "All Account":
        return df
    d = df.copy()
    if selected == "MTI":
        return d[d["ACCOUNT TYPE"].eq("MTI")]
    mapping = {
        "Indomaret": "INDOMARET",
        "Alfamart": "ALFAMART",
        "Alfamidi": "ALFAMIDI",
        "Superindo": "LION SUPERINDO",
    }
    val = mapping.get(selected)
    if val:
        return d[d["CUSTOMER CHAIN"].eq(val)]
    return d


def direct_base(df, selected_account):
    d = df[df["DISTRIBUTION TYPE"].eq("DIRECT")].copy()
    return filter_account(d, selected_account)


def account_perf_base(df, selected_account):
    # FINAL RULE:
    # Distribution Type = ALL
    # Account Type = EXCLUDE DISTRIBUTOR
    d = df[~df["ACCOUNT TYPE"].eq("DISTRIBUTOR")].copy()
    return filter_account(d, selected_account)


def selected_month_df(df, year, month):
    return df[(df["YEAR"].eq(year)) & (df["MONTH"].eq(month))].copy()


def target_value(target_df, year, month_num):
    x = target_df[(target_df["YEAR"].eq(year)) & (target_df["MONTH NUM"].eq(month_num))]
    return float(x["SALES TARGET"].sum()) if len(x) else 0.0


def make_monthly_chart(direct_df, target_df, year, selected_month_num):
    months = MONTH_ORDER[:selected_month_num]
    vals = []
    targets = []
    for m in months:
        vals.append(float(direct_df[
            (direct_df["YEAR"].eq(year)) & (direct_df["MONTH"].eq(m))
        ]["SALES VALUE"].sum()))
        targets.append(target_value(target_df, year, MONTH_NUM[m]))

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=months,
        y=vals,
        name="Actual",
        marker=dict(
            color=vals,
            colorscale=[[0, "#0d78ff"], [1, "#17caff"]],
            line=dict(color="#24c8ff", width=1.2),
        ),
        hovertemplate="%{x}<br>Actual: Rp %{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=months,
        y=targets,
        name="Target",
        mode="lines+markers",
        line=dict(color="#eef7ff", width=2, dash="dash"),
        marker=dict(size=7, color="#eef7ff"),
        hovertemplate="%{x}<br>Target: Rp %{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        height=310,
        margin=dict(l=8, r=8, t=10, b=4),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#d8e8f5", size=11),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1,
            bgcolor="rgba(0,0,0,0)"
        ),
        bargap=.45,
        xaxis=dict(showgrid=False, zeroline=False),
        yaxis=dict(
            gridcolor="rgba(119,168,210,.17)",
            zeroline=False,
            tickformat="~s",
            title="Sales",
        ),
        hoverlabel=dict(bgcolor="#071d36", font_color="#fff"),
    )
    return fig


def horizontal_rows(rows, value_key, amount_formatter, name_key="name"):
    maxv = max([float(r[value_key]) for r in rows] + [1])
    parts = []
    for i, r in enumerate(rows):
        pct = max(0, min(100, float(r[value_key]) / maxv * 100))
        cls = f"g{min(i+1,6)}" if i > 0 else ""
        parts.append(
            f'<div class="product-row">'
            f'<div class="bar-name">{r[name_key]}</div>'
            f'<div class="bar-track"><div class="bar-fill {cls}" style="width:{pct:.1f}%"></div></div>'
            f'<div class="bar-amount">{amount_formatter(r[value_key])} ({r.get("share",0):.0f}%)</div>'
            f'</div>'
        )
    return "".join(parts)


# =========================================================
# DATA
# =========================================================
raw = load_db_bytes()
sales_df, target_df = parse_db_bytes(raw)

available_years = sorted([int(x) for x in sales_df["YEAR"].dropna().unique()])
if not available_years:
    st.error("No valid YEAR values found in Sales Database.")
    st.stop()

latest_row = sales_df.dropna(subset=["FULL DATE"]).sort_values("FULL DATE").tail(1)
if len(latest_row):
    latest_date = latest_row.iloc[0]["FULL DATE"]
else:
    latest_date = pd.Timestamp.today()

selected_year = int(latest_date.year)
available_months = [
    m for m in MONTH_ORDER
    if ((sales_df["YEAR"].eq(selected_year)) & (sales_df["MONTH"].eq(m))).any()
]
if not available_months:
    available_months = [MONTH_ORDER[int(latest_date.month)-1]]

default_month = available_months[-1]

if "month_filter" not in st.session_state:
    st.session_state["month_filter"] = default_month
if st.session_state["month_filter"] not in available_months:
    st.session_state["month_filter"] = default_month


# =========================================================
# HEADER
# =========================================================
logo_b64 = read_logo_b64()
latest_label = pd.Timestamp(latest_date).strftime("%-d %b %Y") if os.name != "nt" else pd.Timestamp(latest_date).strftime("%d %b %Y").lstrip("0")

st.markdown(
    f"""
    <div class="headline-wrap">
      <div class="brand-logo">
        {'<img src="data:image/png;base64,' + logo_b64 + '">' if logo_b64 else '<div style="font-size:24px;font-weight:800">otogard</div>'}
      </div>
      <div class="headline-copy">
        <div class="headline-main">MODERN TRADE</div>
        <div class="headline-sub">Sales Monitoring</div>
      </div>
      <div class="data-badge">
        <div class="tiny">Data per</div>
        <div class="big">{latest_label}</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

f1, f2 = st.columns([1, 1], gap="small")
with f1:
    selected_month = st.selectbox(
        "📅 Month",
        available_months,
        index=available_months.index(st.session_state["month_filter"]),
        key="month_filter_select",
    )
    st.session_state["month_filter"] = selected_month
with f2:
    selected_account = st.selectbox(
        "🏬 Account",
        ["All Account", "Indomaret", "Alfamart", "Alfamidi", "Superindo", "MTI"],
        index=0,
        key="account_filter",
    )

selected_month_num = MONTH_NUM[selected_month]

direct_all = direct_base(sales_df, selected_account)
direct_month = selected_month_df(direct_all, selected_year, selected_month)
account_all = account_perf_base(sales_df, selected_account)
account_month = selected_month_df(account_all, selected_year, selected_month)

# Date context for same-period comparison
month_max_day = int(direct_month["DATE"].max()) if len(direct_month) else 0
month_max_day = month_max_day or 1
selected_start = pd.Timestamp(selected_year, selected_month_num, 1)
selected_end = pd.Timestamp(selected_year, selected_month_num, min(month_max_day, pd.Period(selected_start, freq="M").days_in_month))
prev_start = selected_start - pd.DateOffset(months=1)
prev_month_num = int(prev_start.month)
prev_month_label = MONTH_ORDER[prev_month_num - 1]
prev_year = int(prev_start.year)
prev_end_day = min(month_max_day, pd.Period(prev_start, freq="M").days_in_month)

prev_direct = direct_all[
    (direct_all["YEAR"].eq(prev_year))
    & (direct_all["MONTH"].eq(prev_month_label))
    & (direct_all["DATE"].le(prev_end_day))
]
cur_direct_same = direct_month[direct_month["DATE"].le(month_max_day)]

mtd_actual = float(cur_direct_same["SALES VALUE"].sum())
prev_same = float(prev_direct["SALES VALUE"].sum())
growth = ((mtd_actual / prev_same) - 1) * 100 if prev_same else 0.0

mtd_target = target_value(target_df, selected_year, selected_month_num)
mtd_gap = mtd_actual - mtd_target
mtd_ach = safe_pct(mtd_actual, mtd_target)

sales_to_dist = float(direct_month[direct_month["CATEGORY"].eq("SALES TO DISTRIBUTOR")]["SALES VALUE"].sum())
sales_to_store = float(direct_month[direct_month["CATEGORY"].eq("SALES TO MARKET")]["SALES VALUE"].sum())
flow_total = sales_to_dist + sales_to_store
share_dist = safe_pct(sales_to_dist, flow_total)
share_store = safe_pct(sales_to_store, flow_total)

# YTD
ytd_direct = direct_all[
    (direct_all["YEAR"].eq(selected_year))
    & (direct_all["MONTH NUM"].le(selected_month_num))
]
ytd_actual = float(ytd_direct["SALES VALUE"].sum())
ytd_target = float(target_df[
    (target_df["YEAR"].eq(selected_year))
    & (target_df["MONTH NUM"].le(selected_month_num))
]["SALES TARGET"].sum())
ytd_gap = ytd_actual - ytd_target
ytd_ach = safe_pct(ytd_actual, ytd_target)


# =========================================================
# TOP KPI AREA
# =========================================================
left, right = st.columns([1.08, 1], gap="small")

with left:
    growth_arrow = "▲" if growth >= 0 else "▼"
    growth_color = "#23e6b1" if growth >= 0 else "#ff575f"
    st.markdown(
        f"""
        <div class="card hero-card">
          <div class="section-title">Sales – Month to Date</div>
          <div class="date-line">1 – {month_max_day} {selected_month} {selected_year}
            &nbsp; | &nbsp; vs 1 – {prev_end_day} {prev_month_label} {prev_year}</div>
          <div class="hero-value">{fmt_rp(mtd_actual)}</div>
          <div class="hero-growth" style="color:{growth_color}">
            {growth_arrow} {growth:+.1f}%
            <span class="note">vs same period<br>last month</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with right:
    gap_pct = safe_pct(mtd_gap, mtd_target)
    st.markdown(
        f"""
        <div class="card side-card">
          <div class="section-title">🎯 &nbsp;Achievement</div>
          <div style="display:flex;align-items:center;justify-content:center;gap:20px">
            <div class="kpi-big">{mtd_ach:.0f}%</div>
            <div style="font-size:16px;color:{'#23e6b1' if mtd_gap>=0 else '#ff6a70'}">
              {'▲' if mtd_gap>=0 else '▼'} {gap_pct:+.1f}% (Gap)
            </div>
          </div>
          <div class="kpi-grid3">
            <div class="kpi-cell"><div class="kpi-label">Actual MTD</div><div class="kpi-value">{fmt_rp(mtd_actual)}</div></div>
            <div class="kpi-cell"><div class="kpi-label">Target MTD</div><div class="kpi-value">{fmt_rp(mtd_target)}</div></div>
            <div class="kpi-cell"><div class="kpi-label">Gap</div><div class="kpi-value {'green' if mtd_gap>=0 else 'red'}">{fmt_rp(mtd_gap)}</div></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="card flow-card">
          <div class="section-title">⇄ &nbsp;Sales Flow Breakdown</div>
          <div class="flow-grid">
            <div class="flow-cell">
              <div class="flow-title">Sales to Distributor</div>
              <div class="flow-value">{fmt_rp(sales_to_dist)}</div>
              <div class="flow-share">{share_dist:.0f}%</div>
              <div class="mini-bar"><div class="mini-fill-blue" style="width:{share_dist:.1f}%"></div></div>
            </div>
            <div class="flow-cell">
              <div class="flow-title">Sales to Stores</div>
              <div class="flow-value">{fmt_rp(sales_to_store)}</div>
              <div class="flow-share">{share_store:.0f}%</div>
              <div class="mini-bar"><div class="mini-fill-green" style="width:{share_store:.1f}%"></div></div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# YTD
# =========================================================
st.markdown(
    f"""
    <div class="card section-card ytd-card">
      <div class="section-title-row">
        <div class="section-title">▥ &nbsp;YTD Sales</div>
        <div class="period-label">Jan – {selected_month} {selected_year}{' (MTD)' if month_max_day < pd.Period(selected_start, freq='M').days_in_month else ''}</div>
      </div>
      <div class="ytd-metrics">
        <div class="ytd-item"><div class="ytd-label">Actual Sales</div><div class="ytd-value">{fmt_rp(ytd_actual)}</div></div>
        <div class="ytd-item"><div class="ytd-label">Sales Target</div><div class="ytd-value">{fmt_rp(ytd_target)}</div></div>
        <div class="ytd-item"><div class="ytd-label">Gap</div><div class="ytd-value" style="color:{'#23e6b1' if ytd_gap>=0 else '#ff575f'}">{fmt_rp(ytd_gap)}</div></div>
        <div class="ytd-item"><div class="ytd-label">Achievement</div><div class="ytd-value" style="color:#23e6b1">{ytd_ach:.0f}%</div></div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# MONTHLY SALES TREND
# =========================================================
st.markdown(
    """
    <div class="card section-card" style="padding-bottom:2px">
      <div class="section-title">↗ &nbsp;Monthly Sales Trend</div>
    """,
    unsafe_allow_html=True,
)
fig = make_monthly_chart(direct_all, target_df, selected_year, selected_month_num)
st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# SALES BY CUSTOMER
# =========================================================
cust_rows = []
cust_total = float(direct_month["SALES VALUE"].sum())
for key, display in CUSTOMER_FIXED:
    val = float(direct_month[direct_month["CUSTOMER CHAIN"].eq(key)]["SALES VALUE"].sum())
    cust_rows.append({"name": display, "value": val, "share": safe_pct(val, cust_total)})

max_cust = max([r["value"] for r in cust_rows] + [1])
customer_html = ""
for i, r in enumerate(cust_rows):
    pct = r["value"] / max_cust * 100
    cls = f"g{min(i+1,6)}" if i else ""
    customer_html += f"""
    <div class="bar-row">
      <div class="bar-name">{r['name']}</div>
      <div class="bar-track"><div class="bar-fill {cls}" style="width:{pct:.1f}%"></div></div>
      <div class="bar-amount">{fmt_rp(r['value'])}</div>
      <div class="bar-share">{r['share']:.0f}%</div>
    </div>
    """

st.markdown(
    f"""
    <div class="card section-card">
      <div class="section-title-row">
        <div class="section-title">▦ &nbsp;Sales by Customer <span style="font-size:14px;font-weight:400">(Direct Channel)</span></div>
        <div class="period-label">{selected_month} {selected_year}</div>
      </div>
      <div class="table-head">
        <div>Customer</div><div></div><div style="text-align:right">Sales Amount</div><div style="text-align:right">Share</div>
      </div>
      {customer_html}
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# SALES BY PRODUCT
# =========================================================
st.markdown('<div style="height:2px"></div>', unsafe_allow_html=True)

if hasattr(st, "segmented_control"):
    product_mode = st.segmented_control(
        "Sales by Product metric",
        options=["By Value", "By Qty"],
        default="By Value",
        label_visibility="collapsed",
        key="product_mode",
    )
else:
    product_mode = st.radio(
        "Sales by Product metric",
        ["By Value", "By Qty"],
        horizontal=True,
        label_visibility="collapsed",
        key="product_mode",
    )

metric_col = "SALES VALUE" if product_mode == "By Value" else "SALES QUANTITY"
prod = (
    direct_month.groupby("SKU NAME", as_index=False)[metric_col]
    .sum()
    .sort_values(metric_col, ascending=False)
    .head(5)
)
prod_total = float(direct_month[metric_col].sum())

prod_rows = []
for _, row in prod.iterrows():
    raw_name = str(row["SKU NAME"]).strip().upper()
    name = PRODUCT_DISPLAY.get(raw_name, raw_name.title())
    val = float(row[metric_col])
    prod_rows.append({
        "name": name,
        "value": val,
        "share": safe_pct(val, prod_total)
    })

product_rows_html = horizontal_rows(
    prod_rows,
    "value",
    fmt_rp if product_mode == "By Value" else fmt_qty,
)

product_card_html = (
    '<div class="card section-card" style="margin-top:4px">'
    '<div class="section-title-row">'
    '<div class="section-title">◇ &nbsp;Sales by Product</div>'
    f'<div class="period-label">{selected_month} {selected_year}</div>'
    '</div>'
    + (product_rows_html if product_rows_html else '<div class="muted">No data.</div>')
    + '</div>'
)
st.markdown(product_card_html, unsafe_allow_html=True)


# =========================================================
# ACCOUNT PERFORMANCE
# =========================================================
if hasattr(st, "segmented_control"):
    account_mode = st.segmented_control(
        "Account Performance metric",
        options=["Sales Amount", "Sales Quantity"],
        default="Sales Amount",
        label_visibility="collapsed",
        key="account_mode",
    )
else:
    account_mode = st.radio(
        "Account Performance metric",
        ["Sales Amount", "Sales Quantity"],
        horizontal=True,
        label_visibility="collapsed",
        key="account_mode",
    )

acc_metric = "SALES VALUE" if account_mode == "Sales Amount" else "SALES QUANTITY"

account_rows = []
for key, display in ACCOUNT_FIXED:
    if key == "__MTI__":
        val = float(
            account_month[account_month["ACCOUNT TYPE"].eq("MTI")][acc_metric].sum()
        )
    else:
        val = float(
            account_month[account_month["CUSTOMER CHAIN"].eq(key)][acc_metric].sum()
        )
    account_rows.append({"name": display, "value": val})

acc_total = sum(r["value"] for r in account_rows)
for r in account_rows:
    r["share"] = safe_pct(r["value"], acc_total)

account_rows_html = horizontal_rows(
    account_rows,
    "value",
    fmt_rp if account_mode == "Sales Amount" else fmt_qty,
)

account_card_html = (
    '<div class="card section-card" style="margin-top:4px">'
    '<div class="section-title-row">'
    '<div class="section-title">▦ &nbsp;Account Performance</div>'
    f'<div class="period-label">{account_mode}</div>'
    '</div>'
    + account_rows_html
    + '</div>'
)
st.markdown(account_card_html, unsafe_allow_html=True)


# =========================================================
# INSIGHTS
# =========================================================
top_cust = max(cust_rows, key=lambda x: x["value"]) if cust_rows else {"name":"-", "share":0}
largest_product = max(prod_rows, key=lambda x: x["value"]) if prod_rows else {"name":"-", "share":0}
acc_top = max(account_rows, key=lambda x: x["value"]) if account_rows else {"name":"-", "share":0}

insights = [
    f"{top_cust['name']} contributes {top_cust['share']:.0f}% of Direct sales in {selected_month}.",
    f"Sales to Stores contributes {share_store:.0f}% of Direct MTD sales.",
    f"Achievement is {mtd_ach:.0f}% with a gap of {fmt_rp(mtd_gap)} against target.",
    f"{acc_top['name']} is the largest non-distributor account contributor for the selected month.",
]
insight_html = "".join(
    f'<div class="insight"><div class="insight-no">{i+1}</div><div>{txt}</div></div>'
    for i, txt in enumerate(insights)
)
st.markdown(
    f"""
    <div class="card section-card">
      <div class="section-title">💡 &nbsp;Insights</div>
      <div class="insight-grid">{insight_html}</div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# ADMIN MODE - low profile / floating
# =========================================================
with st.popover("🔐 Admin", use_container_width=False):
    st.markdown("### Admin Mode")
    st.caption("Upload latest daily sales database without cluttering the main dashboard.")

    admin_password = st.text_input("Admin password", type="password", key="admin_password")
    configured_password = st.secrets.get("ADMIN_PASSWORD", "otogard-admin")

    if admin_password == configured_password:
        st.success("Admin access granted.")
        uploaded = st.file_uploader("Upload Excel database", type=["xlsx"], key="admin_upload")
        if uploaded is not None:
            up_bytes = uploaded.getvalue()
            ok, result = validate_workbook(io.BytesIO(up_bytes))
            if not ok:
                st.error(result)
            else:
                up_sales, up_target = result
                max_date = up_sales["FULL DATE"].max()
                st.info(
                    f"Valid file • {len(up_sales):,} sales rows • "
                    f"latest data {pd.Timestamp(max_date).strftime('%d %b %Y')}"
                )
                if st.button("Update Data", type="primary", use_container_width=True):
                    # Keep active immediately.
                    st.session_state["uploaded_db_bytes"] = up_bytes
                    parse_db_bytes.clear()

                    persisted, message = github_persist(up_bytes)
                    if persisted:
                        st.success(message)
                    else:
                        # Also update local file when possible (works in local dev;
                        # Streamlit Cloud local storage may be ephemeral).
                        try:
                            DEFAULT_DB.write_bytes(up_bytes)
                            st.warning(message + " Local app file was updated where writable.")
                        except Exception:
                            st.warning(message)
                    st.rerun()
    elif admin_password:
        st.error("Incorrect password.")

st.markdown(
    """
    <div style="text-align:center;color:#5f7f9b;font-size:10px;margin-top:10px;margin-bottom:4px">
      BUILD V3 — 8 SEP 2026
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div style="text-align:center;color:#849bb2;font-size:10px;letter-spacing:.16em;margin-top:28px">
      OTOGARD &nbsp; | &nbsp; MODERN TRADE SALES MONITORING
    </div>
    """,
    unsafe_allow_html=True,
)
