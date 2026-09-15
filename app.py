
import base64
import io
import json
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Otogard | Modern Trade Sales Monitoring",
    page_icon="📊",
    layout="centered",
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

st.markdown('''
<style>
.stApp {background:#031124;}
.block-container {max-width:430px!important;padding:0!important;margin:0 auto!important;}
#MainMenu,header,footer{visibility:hidden;}
div[data-testid="stVerticalBlock"]{gap:8px!important;}
div[data-testid="stPopover"]{margin:0 10px 20px 10px!important;}
</style>
''', unsafe_allow_html=True)


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
            day=df["DATE"].clip(1, 31),
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
    return df


def validate_workbook(file_like):
    try:
        xls = pd.ExcelFile(file_like)
        if "Sales Database" not in xls.sheet_names:
            return False, "Sheet 'Sales Database' tidak ditemukan."
        if "MONTHLY TARGET" not in xls.sheet_names:
            return False, "Sheet 'MONTHLY TARGET' tidak ditemukan."
        sales = normalise_sales(pd.read_excel(xls, "Sales Database", header=1))
        target = normalise_target(pd.read_excel(xls, "MONTHLY TARGET"))
        if sales.empty:
            return False, "Sales Database kosong."
        return True, (sales, target)
    except Exception as e:
        return False, str(e)


def load_db_bytes():
    if "uploaded_db_bytes" in st.session_state:
        return st.session_state["uploaded_db_bytes"]
    if DEFAULT_DB.exists():
        return DEFAULT_DB.read_bytes()
    raise FileNotFoundError(f"Database not found: {DEFAULT_DB}")


@st.cache_data(show_spinner=False)
def parse_db_bytes(raw_bytes):
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    sales = normalise_sales(pd.read_excel(xls, "Sales Database", header=1))
    target = normalise_target(pd.read_excel(xls, "MONTHLY TARGET"))
    return sales, target


def github_persist(raw_bytes):
    required = ["GITHUB_TOKEN", "REPO", "BRANCH", "FILE_PATH"]
    if not all(k in st.secrets for k in required):
        return False, "GitHub persistence belum dikonfigurasi."

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
    sha = get_resp.json().get("sha") if get_resp.status_code == 200 else None

    payload = {
        "message": "Update Modern Trade sales database from Admin Mode",
        "content": base64.b64encode(raw_bytes).decode("utf-8"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    put_resp = requests.put(api, headers=headers, json=payload, timeout=45)
    if put_resp.status_code not in (200, 201):
        return False, f"GitHub update failed: {put_resp.status_code}"
    return True, "Database berhasil disimpan ke GitHub."


def target_value(target_df, year, month_num):
    x = target_df[(target_df["YEAR"].eq(year)) & (target_df["MONTH NUM"].eq(month_num))]
    return float(x["SALES TARGET"].sum()) if len(x) else 0.0


def safe_pct(num, den):
    return 0.0 if not den else float(num) / float(den) * 100.0


def fmt_rp(v):
    v = float(v or 0)
    sign = "-" if v < 0 else ""
    a = abs(v)
    if a >= 1_000_000_000:
        return f"{sign}Rp {a/1_000_000_000:.2f} B"
    if a >= 1_000_000:
        return f"{sign}Rp {a/1_000_000:.1f} M"
    if a >= 1_000:
        return f"{sign}Rp {a/1_000:.1f} K"
    return f"{sign}Rp {a:,.0f}"


def account_perf_base(df):
    return df[~df["ACCOUNT TYPE"].eq("DISTRIBUTOR")].copy()


def build_month_payload(sales_df, target_df, year, month):
    mnum = MONTH_NUM[month]
    direct_all = sales_df[sales_df["DISTRIBUTION TYPE"].eq("DIRECT")].copy()
    direct_month = direct_all[(direct_all["YEAR"].eq(year)) & (direct_all["MONTH"].eq(month))].copy()

    month_max_day = int(direct_month["DATE"].max()) if len(direct_month) else 1
    selected_start = pd.Timestamp(year, mnum, 1)
    month_days = pd.Period(selected_start, freq="M").days_in_month
    month_max_day = min(max(month_max_day, 1), month_days)

    prev_start = selected_start - pd.DateOffset(months=1)
    prev_year = int(prev_start.year)
    prev_mnum = int(prev_start.month)
    prev_month = MONTH_ORDER[prev_mnum - 1]
    prev_days = pd.Period(prev_start, freq="M").days_in_month
    prev_end_day = min(month_max_day, prev_days)

    current_same = direct_month[direct_month["DATE"].le(month_max_day)]
    previous_same = direct_all[
        (direct_all["YEAR"].eq(prev_year))
        & (direct_all["MONTH"].eq(prev_month))
        & (direct_all["DATE"].le(prev_end_day))
    ]

    mtd_actual = float(current_same["SALES VALUE"].sum())
    prev_actual = float(previous_same["SALES VALUE"].sum())
    growth = ((mtd_actual / prev_actual) - 1) * 100 if prev_actual else 0.0

    mtd_target = target_value(target_df, year, mnum)
    mtd_gap = mtd_actual - mtd_target
    mtd_ach = safe_pct(mtd_actual, mtd_target)

    sales_to_dist = float(
        direct_month[direct_month["CATEGORY"].eq("SALES TO DISTRIBUTOR")]["SALES VALUE"].sum()
    )
    sales_to_store = float(
        direct_month[direct_month["CATEGORY"].eq("SALES TO MARKET")]["SALES VALUE"].sum()
    )
    flow_total = sales_to_dist + sales_to_store
    dist_share = safe_pct(sales_to_dist, flow_total)
    store_share = safe_pct(sales_to_store, flow_total)

    ytd_direct = direct_all[
        (direct_all["YEAR"].eq(year)) & (direct_all["MONTH NUM"].le(mnum))
    ]
    ytd_actual = float(ytd_direct["SALES VALUE"].sum())
    ytd_target = float(
        target_df[
            (target_df["YEAR"].eq(year)) & (target_df["MONTH NUM"].le(mnum))
        ]["SALES TARGET"].sum()
    )
    ytd_gap = ytd_actual - ytd_target
    ytd_ach = safe_pct(ytd_actual, ytd_target)

    trend_months = MONTH_ORDER[:mnum]
    trend_actual = []
    trend_target = []
    for m in trend_months:
        trend_actual.append(
            float(
                direct_all[
                    (direct_all["YEAR"].eq(year)) & (direct_all["MONTH"].eq(m))
                ]["SALES VALUE"].sum()
            )
        )
        trend_target.append(target_value(target_df, year, MONTH_NUM[m]))

    daily = (
        current_same.groupby("DATE")["SALES VALUE"]
        .sum()
        .reindex(range(1, month_max_day + 1), fill_value=0)
        .cumsum()
        .tolist()
    )

    customer_total = float(direct_month["SALES VALUE"].sum())
    customers = []
    for key, display in CUSTOMER_FIXED:
        val = float(direct_month[direct_month["CUSTOMER CHAIN"].eq(key)]["SALES VALUE"].sum())
        customers.append({"name": display, "value": val, "share": safe_pct(val, customer_total)})

    products = {}
    for metric_name, col in [("value", "SALES VALUE"), ("quantity", "SALES QUANTITY")]:
        grouped = (
            direct_month.groupby("SKU NAME", as_index=False)[col]
            .sum()
            .sort_values(col, ascending=False)
            .head(5)
        )
        total = float(direct_month[col].sum())
        rows = []
        for _, r in grouped.iterrows():
            raw = str(r["SKU NAME"]).strip().upper()
            rows.append({
                "name": PRODUCT_DISPLAY.get(raw, raw.title()),
                "value": float(r[col]),
                "share": safe_pct(float(r[col]), total),
            })
        products[metric_name] = rows

    acc_month = account_perf_base(sales_df)
    acc_month = acc_month[(acc_month["YEAR"].eq(year)) & (acc_month["MONTH"].eq(month))]
    accounts = {}
    for metric_name, col in [("value", "SALES VALUE"), ("quantity", "SALES QUANTITY")]:
        rows = []
        for key, display in ACCOUNT_FIXED:
            if key == "__MTI__":
                val = float(acc_month[acc_month["ACCOUNT TYPE"].eq("MTI")][col].sum())
            else:
                val = float(acc_month[acc_month["CUSTOMER CHAIN"].eq(key)][col].sum())
            rows.append({"name": display, "value": val})
        total = sum(r["value"] for r in rows)
        for r in rows:
            r["share"] = safe_pct(r["value"], total)
        accounts[metric_name] = rows

    top_customer = max(customers, key=lambda x: x["value"]) if customers else {"name":"-", "share":0}
    top_account = max(accounts["value"], key=lambda x: x["value"]) if accounts["value"] else {"name":"-"}

    return {
        "month": month, "year": int(year), "latest_day": int(month_max_day),
        "previous_month": prev_month, "previous_year": prev_year, "previous_day": int(prev_end_day),
        "mtd_actual": mtd_actual, "growth": growth, "mtd_target": mtd_target,
        "mtd_gap": mtd_gap, "mtd_achievement": mtd_ach,
        "sales_to_distributor": sales_to_dist, "sales_to_stores": sales_to_store,
        "distributor_share": dist_share, "stores_share": store_share,
        "ytd_actual": ytd_actual, "ytd_target": ytd_target,
        "ytd_gap": ytd_gap, "ytd_achievement": ytd_ach,
        "sparkline": daily, "trend_months": trend_months,
        "trend_actual": trend_actual, "trend_target": trend_target,
        "customers": customers, "products": products, "accounts": accounts,
        "insights": [
            f"{top_customer['name']} contributes {top_customer['share']:.0f}% of Direct sales in {month}.",
            f"Sales to Stores contributes {store_share:.0f}% of Direct MTD sales.",
            f"Achievement is {mtd_ach:.0f}% with a gap of {fmt_rp(mtd_gap)} against target.",
            f"{top_account['name']} is the largest non-distributor account contributor for the selected month.",
        ],
    }


raw = load_db_bytes()
sales_df, target_df = parse_db_bytes(raw)
latest_date = sales_df["FULL DATE"].dropna().max()
if pd.isna(latest_date):
    latest_date = pd.Timestamp.today()

selected_year = int(latest_date.year)
available_months = [
    m for m in MONTH_ORDER
    if ((sales_df["YEAR"].eq(selected_year)) & (sales_df["MONTH"].eq(m))).any()
]
if not available_months:
    available_months = [MONTH_ORDER[int(latest_date.month)-1]]

dashboard_data = {m: build_month_payload(sales_df, target_df, selected_year, m) for m in available_months}

logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("utf-8") if LOGO_PATH.exists() else ""
data_json = json.dumps(dashboard_data, ensure_ascii=False)
months_json = json.dumps(available_months)
default_month = available_months[-1]
latest_label = pd.Timestamp(latest_date).strftime("%d %b %Y").lstrip("0")

html = r'''
<!doctype html><html><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<style>
*{box-sizing:border-box}
html,body{margin:0;padding:0;background:#031124;color:#f7fbff;font-family:Inter,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
body{display:flex;justify-content:center}
#app{width:100%;max-width:390px;padding:10px 10px 18px}
:root{--gap:10px;--radius:11px;--stroke:#14c5ff;--muted:#a9bfd6;--green:#20e5b0;--red:#ff5d66}
.header{display:grid;grid-template-columns:112px 1fr 96px;gap:9px;align-items:center;margin-bottom:10px}
.logo{width:108px;height:48px;object-fit:contain}
.headcopy{border-left:1px solid rgba(255,255,255,.35);padding-left:9px;min-width:0}
.headline{font-size:17px;font-weight:800;line-height:1;white-space:nowrap}
.subhead{font-size:14px;font-weight:300;color:#dce9f5;margin-top:4px;line-height:1}
.datebadge{border:1px solid var(--stroke);border-radius:9px;padding:6px;text-align:center;background:linear-gradient(180deg,#0b3159,#061e39)}
.datebadge .tiny{font-size:7.5px;color:var(--muted)} .datebadge .big{font-size:10px;font-weight:800;margin-top:2px}
.filter-label{font-size:12px;margin:0}.month-row{display:grid;grid-template-columns:82px 1fr;gap:10px;align-items:center;margin:0 0 12px 8px;width:calc(100% - 8px)}
select{width:100%;height:44px;border-radius:9px;border:1px solid #343946;background:#242731;color:#fff;padding:0 16px;font-size:12.5px;outline:none;margin-bottom:0}
.grid-top{display:grid;grid-template-columns:44% 1fr;gap:10px;align-items:stretch}
.stack-right{display:grid;grid-template-rows:117px 99px;gap:10px}
.card{border:1px solid var(--stroke);border-radius:var(--radius);background:linear-gradient(135deg,#0c3b69,#061e3a);box-shadow:inset 0 0 20px rgba(14,115,190,.08);overflow:hidden}
.hero{height:226px;padding:11px 10px;background:radial-gradient(circle at 80% 5%,rgba(16,201,255,.22),transparent 34%),linear-gradient(135deg,#095396,#061f3f)}
.achievement{height:117px;padding:9px}.flow{height:99px;padding:9px}
.title{font-size:13px;font-weight:750;line-height:1.08}
.dates{font-size:7.5px;color:#dce8f2;margin-top:4px;line-height:1.15}
.hero-value{font-size:25px;font-weight:800;margin:18px 0 9px;line-height:1;white-space:nowrap;letter-spacing:-.02em}
.growth{font-size:19px;font-weight:800;color:var(--green);display:flex;gap:6px;align-items:center}
.growth small{font-size:7.5px;color:#dce8f2;font-weight:400;line-height:1.15}
.compare-period{font-size:7.2px;color:#dce8f2;margin-top:7px;padding-top:3px;border-top:1px solid rgba(142,191,230,.18);white-space:nowrap}.hero-illustration{height:58px;margin-top:12px;border-radius:8px;background:linear-gradient(135deg,rgba(22,126,210,.14),rgba(6,31,58,.16));display:grid;grid-template-columns:48px 1fr 40px;align-items:end;gap:8px;padding:8px 10px;opacity:.92}.store-icon{position:relative;width:43px;height:38px;border:1.5px solid rgba(39,202,255,.85);border-radius:4px 4px 2px 2px}.store-icon:before{content:"";position:absolute;left:-2px;top:-8px;width:45px;height:9px;border:1.5px solid rgba(39,202,255,.85);border-bottom:none;border-radius:4px 4px 0 0;background:repeating-linear-gradient(90deg,rgba(39,202,255,.75) 0 5px,transparent 5px 9px)}.store-icon:after{content:"";position:absolute;left:7px;bottom:0;width:10px;height:14px;border:1.5px solid rgba(39,202,255,.70);border-bottom:none}.store-lines{display:flex;flex-direction:column;gap:5px;align-self:center}.store-lines span{display:block;height:4px;border-radius:999px;background:linear-gradient(90deg,rgba(26,192,255,.72),rgba(26,192,255,.08))}.store-lines span:nth-child(1){width:92%}.store-lines span:nth-child(2){width:74%}.store-lines span:nth-child(3){width:55%}.carton-icon{position:relative;width:34px;height:29px;border:1.5px solid rgba(39,202,255,.72);transform:skewY(-8deg);margin-bottom:2px}.carton-icon:before{content:"";position:absolute;left:6px;top:-7px;width:22px;height:8px;border-left:1.5px solid rgba(39,202,255,.72);border-top:1.5px solid rgba(39,202,255,.72);transform:skewY(22deg)}
.ach-row{display:flex;align-items:center;justify-content:center;gap:10px;margin-top:12px}
.ach-big{font-size:24px;font-weight:800;line-height:1;margin:4px 0 2px}.gap-pct{font-size:9px;color:var(--red);white-space:nowrap}
.kpi3{display:grid;grid-template-columns:1fr 1fr 1.16fr;border:1px solid rgba(92,160,214,.35);border-radius:7px;margin-top:10px;background:rgba(1,18,35,.18)}
.kcell{padding:5px 4px;border-right:1px solid rgba(142,191,230,.26)}.kcell:last-child{border-right:none}.kcell:last-child{padding-right:6px}.kcell:last-child .kval{font-size:8.8px;letter-spacing:-.02em;white-space:nowrap}
.klab{font-size:7px;color:#dce8f2}.kval{font-size:9.5px;font-weight:750;margin-top:2px;white-space:nowrap}
.red{color:#ff5d66}.green{color:#20e5b0}
.flowgrid{display:grid;grid-template-columns:1fr 1fr;border:1px solid rgba(92,160,214,.35);border-radius:7px;margin-top:7px}
.flowcell{padding:4px 4px 3px;border-right:1px solid rgba(142,191,230,.26)}.flowcell:last-child{border-right:none}
.flowlab{font-size:7px;color:#dce8f2}.flowval{font-size:11.5px;font-weight:750;margin-top:2px}
.flowshare{font-size:9.5px;font-weight:750;margin-top:4px}.track{height:5px;background:#163b61;border-radius:999px;overflow:hidden;margin-top:3px}
.fill-blue{height:100%;background:linear-gradient(90deg,#188eff,#22c9ff)}.fill-green{height:100%;background:linear-gradient(90deg,#16c99b,#2cebb9)}
.section{margin-top:10px;padding:9px 11px}.section-head{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:10px}
.period{font-size:8px;color:#dbe8f3;white-space:nowrap}
.ytd{min-height:80px;padding-bottom:7px}.ytdgrid{display:grid;grid-template-columns:repeat(4,1fr);border-top:1px solid rgba(102,164,213,.22);padding-top:5px}
.ycell{padding:5px 6px 0;border-right:1px solid rgba(142,191,230,.26)}.ycell:last-child{border-right:none}
.ylab{font-size:7px;color:#dce8f2}.yval{font-size:11px;font-weight:750;margin-top:3px;white-space:nowrap}
.trend{height:189px}.chart-wrap{height:151px;margin-top:0}.chart-wrap svg{width:100%;height:100%;display:block}
.rows-head,.customer-row{display:grid;grid-template-columns:2.05fr 2.55fr 1.42fr .55fr;gap:4px;align-items:center}
.rows-head{font-size:7px;color:#dce8f2;border-bottom:1px solid rgba(110,169,216,.25);padding-bottom:4px}
.customer-row{min-height:22px;font-size:8.2px}.name{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar{height:9px;background:#163959;border-radius:4px;overflow:hidden}.barfill{height:100%;border-radius:4px}
.amount,.share{text-align:right;white-space:nowrap}
.c1{background:linear-gradient(90deg,#1497ff,#22c8ff)}.c2{background:linear-gradient(90deg,#16c99b,#31ebbc)}.c3{background:linear-gradient(90deg,#ffd65a,#ffca42)}.c4{background:linear-gradient(90deg,#ff9854,#ff8444)}.c5{background:linear-gradient(90deg,#ff6768,#ff4e59)}.c6{background:linear-gradient(90deg,#ae6fff,#9458f5)}
.ctrl-card{margin-top:10px;padding:9px 10px}.ctrl-head{display:grid;grid-template-columns:1.75fr 1.2fr .55fr;gap:5px;align-items:center;margin-bottom:6px}
.toggle{display:grid;grid-template-columns:1fr 1fr;border:1px solid rgba(73,128,177,.55);border-radius:5px;overflow:hidden;height:22px}
.toggle button{border:0;background:#081d35;color:#dbe9f5;font-size:7px;padding:0 3px;cursor:pointer}.toggle button.active{background:linear-gradient(180deg,#1d7cff,#1768e6);color:#fff}
.product-row{display:grid;grid-template-columns:1.78fr 2.48fr 1.42fr;gap:4px;align-items:center;min-height:23px;font-size:8.2px}
.account-row{display:grid;grid-template-columns:1.45fr 2.62fr 1.45fr;gap:4px;align-items:center;min-height:23px;font-size:8.2px}
.insights{min-height:136px}.insight-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-top:7px}
.insight{min-height:92px;border:1px solid rgba(54,147,217,.55);background:rgba(8,48,86,.82);border-radius:8px;padding:7px 6px;display:flex;gap:4px;font-size:7.5px;line-height:1.22}
.no{width:20px;height:20px;min-width:20px;border-radius:50%;background:#96ccff;color:#07264a;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:8.5px}
.footer{color:#6f879d;font-size:8px;text-align:center;letter-spacing:.16em;margin-top:24px}.build{color:#56738d;font-size:7.5px;text-align:center;margin-top:12px}
</style></head><body><div id="app">
<div class="header"><img class="logo" src="data:image/png;base64,__LOGO__"/><div class="headcopy"><div class="headline">MODERN TRADE</div><div class="subhead">Sales Monitoring</div></div><div class="datebadge"><div class="tiny">Data per</div><div class="big">__LATEST__</div></div></div>
<div class="month-row"><div class="filter-label">📅 Month</div><select id="monthSelect"></select></div>
<div class="grid-top">
<div class="card hero"><div class="title">Sales – Month to Date</div><div class="dates" id="mtdCurrentDate"></div><div class="hero-value" id="mtdActual"></div><div class="growth"><span id="growthValue"></span><small>vs same period<br/>last month</small></div><div class="compare-period" id="mtdCompare"></div><div class="hero-illustration"><div class="store-icon"></div><div class="store-lines"><span></span><span></span><span></span></div><div class="carton-icon"></div></div></div>
<div class="stack-right">
<div class="card achievement"><div class="title">🎯 &nbsp;Achievement</div><div class="ach-row"><div class="ach-big" id="mtdAch"></div><div class="gap-pct" id="mtdGapPct"></div></div><div class="kpi3"><div class="kcell"><div class="klab">Actual MTD</div><div class="kval" id="achActual"></div></div><div class="kcell"><div class="klab">Target MTD</div><div class="kval" id="achTarget"></div></div><div class="kcell"><div class="klab">Gap</div><div class="kval red" id="achGap"></div></div></div></div>
<div class="card flow"><div class="title">⇄ &nbsp;Sales Flow Breakdown</div><div class="flowgrid"><div class="flowcell"><div class="flowlab">Sales to Distributor</div><div class="flowval" id="distVal"></div><div class="flowshare" id="distShare"></div><div class="track"><div class="fill-blue" id="distBar"></div></div></div><div class="flowcell"><div class="flowlab">Sales to Stores</div><div class="flowval" id="storeVal"></div><div class="flowshare" id="storeShare"></div><div class="track"><div class="fill-green" id="storeBar"></div></div></div></div></div>
</div></div>
<div class="card section ytd"><div class="section-head"><div class="title">▥ &nbsp;YTD Sales</div><div class="period" id="ytdPeriod"></div></div><div class="ytdgrid"><div class="ycell"><div class="ylab">Actual Sales</div><div class="yval" id="ytdActual"></div></div><div class="ycell"><div class="ylab">Sales Target</div><div class="yval" id="ytdTarget"></div></div><div class="ycell"><div class="ylab">Gap</div><div class="yval red" id="ytdGap"></div></div><div class="ycell"><div class="ylab">Achievement</div><div class="yval green" id="ytdAch"></div></div></div></div>
<div class="card section trend"><div class="section-head"><div class="title">↗ &nbsp;Monthly Sales Trend</div><div class="period"><span style="color:#18bfff">■</span> Actual &nbsp;&nbsp; ─●─ Target</div></div><div class="chart-wrap" id="trendChart"></div></div>
<div class="card section"><div class="section-head"><div class="title">▦ &nbsp;Sales by Customer <span style="font-size:8px;font-weight:400">(Direct Channel)</span></div><div class="period" id="customerPeriod"></div></div><div class="rows-head"><div>Customer</div><div></div><div style="text-align:right">Sales Amount</div><div style="text-align:right">Share</div></div><div id="customerRows"></div></div>
<div class="card ctrl-card"><div class="ctrl-head"><div class="title">◇ &nbsp;Sales by Product</div><div class="toggle"><button id="prodValue" class="active">Value</button><button id="prodQty">Quantity</button></div><div class="period" id="productPeriod"></div></div><div id="productRows"></div></div>
<div class="card ctrl-card"><div class="ctrl-head"><div class="title">▦ &nbsp;Account Performance</div><div class="toggle"><button id="accValue" class="active">Value</button><button id="accQty">Quantity</button></div><div class="period" id="accountPeriod"></div></div><div id="accountRows"></div></div>
<div class="card section insights"><div class="title">💡 &nbsp;Insights</div><div class="insight-grid" id="insightRows"></div></div>
<div class="build">BUILD V16 YTD TREND REVISION — 15 SEP 2026</div><div class="footer">OTOGARD &nbsp; | &nbsp; MODERN TRADE SALES MONITORING</div>
</div>
<script>
const DATA=__DATA__,MONTHS=__MONTHS__;let selectedMonth="__DEFAULT__",productMetric="value",accountMetric="value";
const colors=["c1","c2","c3","c4","c5","c6"];
function rp(v){const s=v<0?"-":"",a=Math.abs(v||0);if(a>=1e9)return `${s}Rp ${(a/1e9).toFixed(2)} B`;if(a>=1e6)return `${s}Rp ${(a/1e6).toFixed(1)} M`;if(a>=1e3)return `${s}Rp ${(a/1e3).toFixed(1)} K`;return `${s}Rp ${Math.round(a).toLocaleString()}`;}
function qty(v){const a=Math.abs(v||0);if(a>=1e6)return (v/1e6).toFixed(2)+" M";if(a>=1e3)return (v/1e3).toFixed(1)+" K";return Math.round(v).toLocaleString();}
function pct(v){return `${Math.round(v||0)}%`;}
function sparkline(vals){const w=150,h=62,p=4;if(!vals||!vals.length)vals=[0,0];let lo=Math.min(...vals),hi=Math.max(...vals);if(hi===lo)hi=lo+1;const pts=vals.map((v,i)=>{const x=p+(w-p*2)*(i/Math.max(1,vals.length-1));const y=h-p-(h-p*2)*((v-lo)/(hi-lo));return[x,y]});const path=pts.map(x=>x.join(",")).join(" "),last=pts[pts.length-1];return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none"><polyline points="${path}" fill="none" stroke="#1bc7ff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/><circle cx="${last[0]}" cy="${last[1]}" r="3.2" fill="#21dbff"/></svg>`;}
function trendSvg(d){const labels=d.trend_months,actual=d.trend_actual,target=d.trend_target,w=350,h=220,left=36,right=8,top=14,bottom=28,cw=w-left-right,ch=h-top-bottom,maxv=Math.max(1,...actual,...target)*1.12,n=labels.length,step=cw/Math.max(n,1),barW=Math.min(24,step*.55);let bars="",xl="",grid="",lp="",cir="";[0,.25,.5,.75,1].forEach(fr=>{const y=top+ch*(1-fr),val=maxv*fr;grid+=`<line x1="${left}" y1="${y}" x2="${w-right}" y2="${y}" stroke="rgba(111,164,207,.18)" stroke-width="1"/><text x="${left-5}" y="${y+3}" fill="#b8cce0" font-size="8" text-anchor="end">${val>=1e9?(val/1e9).toFixed(1)+"B":Math.round(val/1e6)+"M"}</text>`});labels.forEach((lab,i)=>{const cx=left+step*(i+.5),bh=ch*(actual[i]/maxv),y=top+ch-bh,ty=top+ch-ch*(target[i]/maxv);bars+=`<rect x="${cx-barW/2}" y="${y}" width="${barW}" height="${bh}" rx="1.5" fill="#18bfff" stroke="#22c9ff" stroke-width=".7"/>`;xl+=`<text x="${cx}" y="${h-9}" fill="#d6e4f1" font-size="8" text-anchor="middle">${lab}</text>`;lp+=`${cx},${ty} `;cir+=`<circle cx="${cx}" cy="${ty}" r="3" fill="#edf7ff"/>`});return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">${grid}${bars}<polyline points="${lp}" fill="none" stroke="#edf7ff" stroke-width="2" stroke-dasharray="6 5"/>${cir}${xl}</svg>`;}
function customerRows(rows){const max=Math.max(1,...rows.map(r=>r.value));return rows.map((r,i)=>`<div class="customer-row"><div class="name">${r.name}</div><div class="bar"><div class="barfill ${colors[Math.min(i,5)]}" style="width:${Math.max(0,r.value/max*100)}%"></div></div><div class="amount">${rp(r.value)}</div><div class="share">${Math.round(r.share)}%</div></div>`).join("");}
function metricRows(rows,metric,type){const max=Math.max(1,...rows.map(r=>r.value)),cls=type==="product"?"product-row":"account-row";return rows.map((r,i)=>`<div class="${cls}"><div class="name">${r.name}</div><div class="bar"><div class="barfill ${colors[Math.min(i,5)]}" style="width:${Math.max(0,r.value/max*100)}%"></div></div><div class="amount">${metric==="value"?rp(r.value):qty(r.value)} (${Math.round(r.share)}%)</div></div>`).join("");}
function render(){const d=DATA[selectedMonth];document.getElementById("mtdCurrentDate").innerText=`1 – ${d.latest_day} ${d.month} ${d.year}`;document.getElementById("mtdCompare").innerText=`1 – ${d.latest_day} ${d.month} ${d.year} | vs 1 – ${d.previous_day} ${d.previous_month} ${d.previous_year}`;document.getElementById("mtdActual").innerText=rp(d.mtd_actual);const g=d.growth||0,gv=document.getElementById("growthValue");gv.innerText=`${g>=0?"▲":"▼"} ${g>=0?"+":""}${g.toFixed(1)}%`;gv.style.color=g>=0?"#20e5b0":"#ff5d66";document.getElementById("mtdAch").innerText=pct(d.mtd_achievement);const gp=d.mtd_target?d.mtd_gap/d.mtd_target*100:0;document.getElementById("mtdGapPct").innerText=`${gp>=0?"▲":"▼"} ${gp>=0?"+":""}${gp.toFixed(1)}% (Gap)`;document.getElementById("achActual").innerText=rp(d.mtd_actual);document.getElementById("achTarget").innerText=rp(d.mtd_target);document.getElementById("achGap").innerText=rp(d.mtd_gap);document.getElementById("distVal").innerText=rp(d.sales_to_distributor);document.getElementById("storeVal").innerText=rp(d.sales_to_stores);document.getElementById("distShare").innerText=pct(d.distributor_share);document.getElementById("storeShare").innerText=pct(d.stores_share);document.getElementById("distBar").style.width=`${d.distributor_share}%`;document.getElementById("storeBar").style.width=`${d.stores_share}%`;document.getElementById("ytdPeriod").innerText=`Jan – ${d.month} ${d.year}`;document.getElementById("ytdActual").innerText=rp(d.ytd_actual);document.getElementById("ytdTarget").innerText=rp(d.ytd_target);document.getElementById("ytdGap").innerText=rp(d.ytd_gap);document.getElementById("ytdAch").innerText=pct(d.ytd_achievement);document.getElementById("trendChart").innerHTML=trendSvg(d);document.getElementById("customerPeriod").innerText=`${d.month} ${d.year}`;document.getElementById("customerRows").innerHTML=customerRows(d.customers);document.getElementById("productPeriod").innerText=`${d.month} ${d.year}`;document.getElementById("productRows").innerHTML=metricRows(d.products[productMetric],productMetric,"product");document.getElementById("accountPeriod").innerText=`${d.month} ${d.year}`;document.getElementById("accountRows").innerHTML=metricRows(d.accounts[accountMetric],accountMetric,"account");document.getElementById("insightRows").innerHTML=d.insights.map((x,i)=>`<div class="insight"><div class="no">${i+1}</div><div>${x}</div></div>`).join("");}
const sel=document.getElementById("monthSelect");MONTHS.forEach(m=>{const o=document.createElement("option");o.value=m;o.textContent=m;if(m===selectedMonth)o.selected=true;sel.appendChild(o)});sel.addEventListener("change",e=>{selectedMonth=e.target.value;render()});
document.getElementById("prodValue").onclick=()=>{productMetric="value";prodValue.classList.add("active");prodQty.classList.remove("active");render()};document.getElementById("prodQty").onclick=()=>{productMetric="quantity";prodQty.classList.add("active");prodValue.classList.remove("active");render()};document.getElementById("accValue").onclick=()=>{accountMetric="value";accValue.classList.add("active");accQty.classList.remove("active");render()};document.getElementById("accQty").onclick=()=>{accountMetric="quantity";accQty.classList.add("active");accValue.classList.remove("active");render()};render();
</script></body></html>
'''

html = (html
        .replace("__LOGO__", logo_b64)
        .replace("__LATEST__", latest_label)
        .replace("__DATA__", data_json)
        .replace("__MONTHS__", months_json)
        .replace("__DEFAULT__", default_month))

components.html(html, height=1900, scrolling=False)

with st.popover("🔐 Admin"):
    st.markdown("### Admin Mode")
    st.caption("Upload database terbaru untuk refresh dashboard.")
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
                up_sales, _ = result
                max_date = up_sales["FULL DATE"].max()
                st.info(f"Valid file • {len(up_sales):,} sales rows • latest data {pd.Timestamp(max_date).strftime('%d %b %Y')}")
                if st.button("Update Data", type="primary", use_container_width=True):
                    st.session_state["uploaded_db_bytes"] = up_bytes
                    parse_db_bytes.clear()
                    persisted, message = github_persist(up_bytes)
                    if persisted:
                        st.success(message)
                    else:
                        try:
                            DEFAULT_DB.write_bytes(up_bytes)
                            st.warning(message + " Local app file was updated where writable.")
                        except Exception:
                            st.warning(message)
                    st.rerun()
    elif admin_password:
        st.error("Incorrect password.")
