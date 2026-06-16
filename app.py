"""
Meta Ads Multi-Region Dashboard (EMEA / India / NA)
Run: streamlit run app.py
Install: pip install streamlit pandas plotly openpyxl python-calamine numpy gspread google-auth

Data source: single Excel / Google Sheet with tabs named {REGION}_{Type}
e.g. EMEA_Daily, EMEA_Weekly, EMEA_Manufacture, NA_Daily, INDIA_Smartlink ...

Tab naming rules:
  - Prefix before first underscore = region  (EMEA / NA / INDIA)
  - Suffix after first underscore  = type    (Daily / Weekly / campaign name)
  - Tabs named {REGION}_Daily are treated as the daily total for that region
  - Tabs named {REGION}_Weekly are treated as weekly summaries
  - All other {REGION}_* tabs are treated as campaign / ad-set detail sheets
"""

import re
import io
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Meta Ads Dashboard", layout="wide")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
STANDARD_COLS = [
    "Date", "Spend", "Reach", "Impressions", "Link Clicks", "Leads",
    "Frequency", "CTR", "CPC", "CPL", "Lead CVR", "3-Day Avg CPL",
    "Status", "Daily New Leads", "Daily CPL Active", "Key actions",
]

NUMERIC_COLS = [
    "Spend", "Reach", "Impressions", "Link Clicks", "Leads", "Frequency",
    "CTR", "CPC", "CPL", "Lead CVR", "3-Day Avg CPL",
    "Daily New Leads", "Daily CPL Active",
]

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

SALES_SHEET_CANDIDATES = ["销售线索数据"]
SALES_ADSET_COLS = [
    "Lead Source Drill-Down 3", "Lead Source Drill-Down 3(Old)",
    "营销活动", "市场活动（text）", "市场活动",
]
SALES_STAGE_COLS = ["线索阶段", "Lifecycle Stage", "生命状态"]
SALES_CREATED_TIME_COLS = ["创建时间", "创建日期", "网站注册用户的提交日期"]
SALES_CONVERTED_TIME_COLS = ["转换时间", "线索阶段变更时间"]
SALES_MQL_TIME_COLS = ["转MQL时间"]
SALES_COUNTRY_COLS = ["Country / Region of Lead", "Country （文本）", "国家", "手机归属国家"]
SALES_DUPLICATE_COLS = ["是否存在重复数据"]
SALES_ID_COLS = ["线索编号（必填）", "唯一标识", "客户ID", "Email", "邮件", "电话", "Telephone", "Mobile", "手机"]

CAMPAIGN_ALIAS_RULES = {
    "Manufacture": ["manufacture", "manufacturing", "manufacturer"],
    "CA": ["customer acquisition", "customers acquisition", "sales acquisition", "acquisition"],
    "AI Instant": ["ai instant", "ai design", "ai render", "ai rendering"],
    "AI Layout": ["ai layout"],
    "KC Tools": ["kc tools", "kc tool", "kctools", "kc_tool"],
    "SmartLinkSuite": ["smartlinksuite", "smart link suite", "smartlink suite"],
    "Single Image": ["single image", "single-image", "singleimage", "static image", "image ad"],
    "Carousel": ["carousel", "carrousel", "carousal"],
    "SA": ["sa", "search ads", "search acquisition"],
    "Smartlink": ["smartlink", "smart link"],
}

EMEA_COUNTRIES = {
    "egypt", "saudi arabia", "united arab emirates", "uae", "south africa",
    "italy", "namibia", "united kingdom", "uk", "kenya", "sri lanka",
    "turkey", "germany", "france", "spain", "qatar", "kuwait", "oman", "bahrain",
}
INDIA_COUNTRIES = {"india"}
NA_COUNTRIES = {"united states", "usa", "us", "canada", "mexico"}


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------
def clean_numeric(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype="float64")
    return pd.to_numeric(series.replace("#REF!", np.nan), errors="coerce")


def normalize_key(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s\-_]+", " ", text)
    text = re.sub(r"[^a-z0-9 ]+", "", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_region(value) -> str:
    key = normalize_key(value)
    if not key:
        return "UNKNOWN"
    if key in {"india", "in"}:
        return "INDIA"
    if key in {"na", "north america", "us", "usa", "united states", "america"}:
        return "NA"
    if key in {"emea", "europe middle east africa", "middle east", "gcc"}:
        return "EMEA"
    return str(value).strip().upper()


def parse_chinese_date(date_str, reference_year=2026):
    """
    Parse Chinese date format like '4月1日' or '12月31日'
    Returns a datetime object with year 2026
    """
    if pd.isna(date_str) or date_str == "":
        return pd.NaT

    date_str = str(date_str).strip()

    # Try standard datetime parsing first
    try:
        result = pd.to_datetime(date_str, errors="coerce")
        if pd.notna(result):
            return result
    except:
        pass

    # Try Chinese format: X月Y日
    match = re.match(r'(\d+)月(\d+)日', date_str)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))

        try:
            return pd.Timestamp(year=reference_year, month=month, day=day)
        except:
            return pd.NaT

    return pd.NaT


def first_existing_col(df: pd.DataFrame, candidates: Iterable[str]):
    for col in candidates:
        if col in df.columns:
            return col
    return None


def coalesce_columns(df: pd.DataFrame, candidates: Iterable[str], default=np.nan) -> pd.Series:
    existing = [c for c in candidates if c in df.columns]
    if not existing:
        return pd.Series([default] * len(df), index=df.index)
    result = df[existing[0]].copy()
    for col in existing[1:]:
        result = result.where(result.notna() & (result.astype(str).str.strip() != ""), df[col])
    return result


def fmt_number(value, decimals=0, prefix="", suffix="") -> str:
    if pd.isna(value) or value in (np.inf, -np.inf):
        return "-"
    return f"{prefix}{value:,.{decimals}f}{suffix}"


def fmt_percent(value, decimals=2) -> str:
    if pd.isna(value) or value in (np.inf, -np.inf):
        return "-"
    return f"{value:.{decimals}%}"


def open_excel_bytes(file_bytes: bytes) -> pd.ExcelFile:
    try:
        return pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception:
        return pd.ExcelFile(io.BytesIO(file_bytes), engine="calamine")


# ---------------------------------------------------------------------------
# Tab parsing: extract region + type from tab name
# ---------------------------------------------------------------------------
def parse_tab_name(sheet_name: str):
    """
    'EMEA_Daily'       -> ('EMEA', 'Daily',       'daily_total')
    'EMEA_Weekly'      -> ('EMEA', 'Weekly',       'weekly')
    'EMEA_Manufacture' -> ('EMEA', 'Manufacture',  'campaign')
    'NA_Single'        -> ('NA',   'Single',        'campaign')
    Returns (region, label, tab_type) where tab_type in
        {'daily_total', 'weekly', 'campaign', 'unknown'}
    """
    parts = sheet_name.split("_", 1)
    if len(parts) != 2:
        return None, sheet_name, "unknown"

    region_raw, label = parts[0].strip().upper(), parts[1].strip()
    region = normalize_region(region_raw)

    label_lower = label.lower()
    if label_lower == "daily":
        tab_type = "daily_total"
    elif label_lower == "weekly":
        tab_type = "weekly"
    else:
        tab_type = "campaign"

    return region, label, tab_type


def classify_sheets(sheet_names):
    """
    Returns dict keyed by region:
      { 'EMEA': {'daily': 'EMEA_Daily', 'weekly': 'EMEA_Weekly',
                 'campaigns': ['EMEA_Manufacture', 'EMEA_CA', ...]},
        'NA':   {...}, ... }
    """
    result = {}
    for name in sheet_names:
        region, label, tab_type = parse_tab_name(name)
        if region is None or region == "UNKNOWN":
            continue
        if region not in result:
            result[region] = {"daily": None, "weekly": None, "campaigns": []}
        if tab_type == "daily_total":
            result[region]["daily"] = name
        elif tab_type == "weekly":
            result[region]["weekly"] = name
        elif tab_type == "campaign":
            result[region]["campaigns"].append(name)
    return result


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------
def load_campaign_sheet(xls: pd.ExcelFile, sheet_name: str, region: str, campaign: str) -> pd.DataFrame:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    for col in STANDARD_COLS:
        if col not in df.columns:
            df[col] = np.nan
    df = df[STANDARD_COLS].copy()

    for col in NUMERIC_COLS:
        df[col] = clean_numeric(df[col])

    mask_cpc = df["CPC"].isna() & df["Spend"].notna() & df["Link Clicks"].notna() & (df["Link Clicks"] != 0)
    df.loc[mask_cpc, "CPC"] = df.loc[mask_cpc, "Spend"] / df.loc[mask_cpc, "Link Clicks"]

    mask_ctr = df["Impressions"].notna() & (df["Impressions"] != 0) & df["Link Clicks"].notna()
    df.loc[mask_ctr, "CTR"] = df.loc[mask_ctr, "Link Clicks"] / df.loc[mask_ctr, "Impressions"]

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df = df[df["Spend"].notna()]

    df["Campaign"] = campaign
    df["Region"] = region
    df["RegionKey"] = normalize_region(region)
    return df.reset_index(drop=True)


def load_weekly_sheet(xls: pd.ExcelFile, sheet_name: str, region: str) -> pd.DataFrame:
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    records = []
    current_month = None

    for _, row in raw.iterrows():
        first = row[0]
        if pd.isna(first):
            continue
        first_str = str(first).strip()
        if first_str in MONTH_NAMES:
            current_month = first_str
            continue
        if first_str == "Week":
            continue
        if first_str.startswith("Week"):
            spend = row[1] if len(row) > 1 else np.nan
            leads = row[2] if len(row) > 2 else np.nan
            avg_cpl = row[3] if len(row) > 3 else np.nan
            records.append({
                "Month": current_month,
                "Week": first_str,
                "Spend": spend,
                "Leads": leads,
                "Avg CPL": avg_cpl,
            })

    df = pd.DataFrame(records)
    if df.empty:
        df["Region"] = region
        df["RegionKey"] = normalize_region(region)
        df["Period"] = None
        return df

    for col in ["Spend", "Leads", "Avg CPL"]:
        df[col] = clean_numeric(df[col])

    df = df.dropna(subset=["Spend", "Leads", "Avg CPL"], how="all").reset_index(drop=True)
    if df.empty:
        df["Region"] = region
        df["RegionKey"] = normalize_region(region)
        df["Period"] = None
        return df

    month_order = {m: i for i, m in enumerate(MONTH_NAMES)}
    df["MonthOrder"] = df["Month"].map(month_order)
    df["WeekNum"] = df["Week"].str.extract(r"(\d+)").astype(int)
    df = df.sort_values(["MonthOrder", "WeekNum"]).reset_index(drop=True)
    df["Period"] = df["Month"].str[:3] + " " + df["Week"]
    df["Region"] = region
    df["RegionKey"] = normalize_region(region)
    return df


@st.cache_data(show_spinner=False)
def load_single_file(file_bytes: bytes):
    """Load the unified overview Excel and return (campaigns_df, daily_total_df, weekly_df)."""
    xls = open_excel_bytes(file_bytes)
    sheet_map = classify_sheets(xls.sheet_names)

    all_campaigns, all_daily, all_weekly = [], [], []

    for region, sheets in sheet_map.items():
        # Campaign / ad-set sheets
        for sheet_name in sheets["campaigns"]:
            _, label, _ = parse_tab_name(sheet_name)
            try:
                df = load_campaign_sheet(xls, sheet_name, region, label)
                all_campaigns.append(df)
            except Exception as e:
                st.warning(f"Could not load sheet '{sheet_name}': {e}")

        # Daily total sheet
        if sheets["daily"]:
            try:
                df = load_campaign_sheet(xls, sheets["daily"], region, "Total")
                all_daily.append(df)
            except Exception as e:
                st.warning(f"Could not load daily sheet '{sheets['daily']}': {e}")

        # Weekly sheet
        if sheets["weekly"]:
            try:
                df = load_weekly_sheet(xls, sheets["weekly"], region)
                if not df.empty and "Period" in df.columns and df["Period"].notna().any():
                    all_weekly.append(df)
            except Exception as e:
                st.warning(f"Could not load weekly sheet '{sheets['weekly']}': {e}")

    campaigns_df = pd.concat(all_campaigns, ignore_index=True) if all_campaigns else pd.DataFrame()
    daily_total_df = pd.concat(all_daily, ignore_index=True) if all_daily else pd.DataFrame()
    weekly_df = pd.concat(all_weekly, ignore_index=True) if all_weekly else pd.DataFrame()
    return campaigns_df, daily_total_df, weekly_df


# ---------------------------------------------------------------------------
# Sales / SQL helpers (unchanged logic, English labels)
# ---------------------------------------------------------------------------
def infer_region_from_sales_row(raw_adset, country) -> str:
    text = normalize_key(raw_adset)
    country_key = normalize_key(country)
    if "emea" in text:
        return "EMEA"
    if "india" in text or country_key in INDIA_COUNTRIES:
        return "INDIA"
    if re.search(r"(^|\s)(us|usa|na)(\s|$)", text) or country_key in NA_COUNTRIES:
        return "NA"
    if country_key in EMEA_COUNTRIES:
        return "EMEA"
    if country_key in INDIA_COUNTRIES:
        return "INDIA"
    if country_key in NA_COUNTRIES:
        return "NA"
    return "UNKNOWN"


def map_raw_adset_to_campaign(raw_adset, loaded_campaigns=None) -> str:
    key = normalize_key(raw_adset)
    if not key:
        return "Unmapped"
    for campaign in (loaded_campaigns or []):
        ckey = normalize_key(campaign)
        if ckey and (ckey == key or ckey in key):
            return campaign
    for campaign, aliases in CAMPAIGN_ALIAS_RULES.items():
        if any(alias in key for alias in aliases):
            return campaign
    return "Unmapped"


def build_lead_id(df: pd.DataFrame) -> pd.Series:
    lead_id = coalesce_columns(df, SALES_ID_COLS)
    lead_id = lead_id.astype(str).replace({"nan": "", "NaN": "", "None": ""}).str.strip()
    fallback = pd.Series([f"row_{i}" for i in range(len(df))], index=df.index)
    return lead_id.where(lead_id != "", fallback)


@st.cache_data(show_spinner=False)
def load_sales_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    xls = open_excel_bytes(file_bytes)
    sheet = next((s for s in SALES_SHEET_CANDIDATES if s in xls.sheet_names), xls.sheet_names[0])
    raw = pd.read_excel(xls, sheet_name=sheet)
    raw.columns = [str(c).strip() for c in raw.columns]

    adset = coalesce_columns(raw, SALES_ADSET_COLS)
    stage = coalesce_columns(raw, SALES_STAGE_COLS)
    created_time = coalesce_columns(raw, SALES_CREATED_TIME_COLS)
    converted_time = coalesce_columns(raw, SALES_CONVERTED_TIME_COLS)
    mql_time = coalesce_columns(raw, SALES_MQL_TIME_COLS)
    country = coalesce_columns(raw, SALES_COUNTRY_COLS)
    duplicate_flag = coalesce_columns(raw, SALES_DUPLICATE_COLS, default="")

    df = pd.DataFrame({
        "Source File": file_name,
        "Source Sheet": sheet,
        "Lead ID": build_lead_id(raw),
        "Raw Adset": adset,
        "Lead Stage": stage,
        "Created Time": pd.to_datetime(created_time, errors="coerce"),
        "Converted Time": pd.to_datetime(converted_time, errors="coerce"),
        "MQL Time": pd.to_datetime(mql_time, errors="coerce"),
        "Country": country,
        "Duplicate Flag": duplicate_flag,
    })

    stage_text = df["Lead Stage"].fillna("").astype(str).str.lower()
    file_is_sql = "sql" in file_name.lower()
    has_stage_signal = stage_text.str.contains("sql|mql", regex=True).any()
    df["Is SQL"] = stage_text.str.contains("sql", regex=True)
    if file_is_sql and not has_stage_signal:
        df["Is SQL"] = True
    df["Is CRM MQL"] = stage_text.str.contains("mql|sql", regex=True)
    df["SQL Date"] = df["Converted Time"].fillna(df["Created Time"])
    df["MQL Cohort Date"] = df["Created Time"].fillna(df["MQL Time"]).fillna(df["Converted Time"])
    df["Inferred Region"] = [infer_region_from_sales_row(a, c) for a, c in zip(df["Raw Adset"], df["Country"])]
    df["RegionKey"] = df["Inferred Region"].apply(normalize_region)
    df["Raw Adset Clean"] = df["Raw Adset"].fillna("Unmapped").astype(str).replace({"nan": "Unmapped"})
    return df


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
st.sidebar.title("Data Sources")

# --- Meta Ads file ---
st.sidebar.subheader("Meta Ads Overview")
data_source = st.sidebar.radio(
    "Load data from",
    ["Upload Excel file", "Connect Google Sheets"],
    index=0,
)

campaigns_df = pd.DataFrame()
daily_total_df = pd.DataFrame()
weekly_df = pd.DataFrame()

if data_source == "Upload Excel file":
    uploaded_file = st.sidebar.file_uploader(
        "Upload the unified Meta Ads overview (.xlsx)",
        type=["xlsx"],
        key="meta_file",
    )
    if uploaded_file:
        with st.spinner("Loading data..."):
            campaigns_df, daily_total_df, weekly_df = load_single_file(uploaded_file.getvalue())

else:
    st.sidebar.info(
        "Connect to the Meta Ads Google Sheet"
    )

    # Embedded Google Sheet URL
    GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1_a1Ddh1Pe09GpC4r1l_vMOqCSADKZRF4OBaGSS0w84o/"

    # Direct load on button click
    if st.sidebar.button("Connect to Google Sheets", type="primary"):
        sheet_url = GOOGLE_SHEET_URL
        try:
            import gspread
            from google.oauth2.service_account import Credentials

            scopes = [
                "https://www.googleapis.com/auth/spreadsheets.readonly",
                "https://www.googleapis.com/auth/drive.readonly",
            ]
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"], scopes=scopes
            )
            gc = gspread.authorize(creds)
            sh = gc.open_by_url(sheet_url)

            with st.spinner("Loading data from Google Sheets..."):
                sheet_map = classify_sheets([ws.title for ws in sh.worksheets()])
                all_campaigns, all_daily, all_weekly = [], [], []

                for region, sheets in sheet_map.items():
                    # Campaign sheets
                    for sheet_name in sheets["campaigns"]:
                        try:
                            ws = sh.worksheet(sheet_name)
                            records = ws.get_all_records(expected_headers=[], numericise_ignore=["all"])
                            if not records:
                                continue
                            df = pd.DataFrame(records)
                            _, label, _ = parse_tab_name(sheet_name)
                            # Numeric conversion
                            for col in STANDARD_COLS:
                                if col not in df.columns:
                                    df[col] = np.nan
                            df = df[STANDARD_COLS].copy()
                            for col in NUMERIC_COLS:
                                df[col] = clean_numeric(df[col])
                            mask_cpc = df["CPC"].isna() & df["Spend"].notna() & df["Link Clicks"].notna() & (df["Link Clicks"] != 0)
                            df.loc[mask_cpc, "CPC"] = df.loc[mask_cpc, "Spend"] / df.loc[mask_cpc, "Link Clicks"]
                            mask_ctr = df["Impressions"].notna() & (df["Impressions"] != 0) & df["Link Clicks"].notna()
                            df.loc[mask_ctr, "CTR"] = df.loc[mask_ctr, "Link Clicks"] / df.loc[mask_ctr, "Impressions"]
                            # Parse dates with Chinese format support
                            df["Date"] = df["Date"].apply(parse_chinese_date)
                            df = df.dropna(subset=["Date"])

                            # Debug: Check NA sheets
                            if region == "NA" and not df.empty:
                                st.sidebar.write(f"DEBUG NA: {sheet_name}, rows before Spend filter: {len(df)}, Spend notna: {df['Spend'].notna().sum()}")

                            df = df[df["Spend"].notna()]
                            df["Campaign"] = label
                            df["Region"] = region
                            df["RegionKey"] = normalize_region(region)
                            if not df.empty:
                                all_campaigns.append(df.reset_index(drop=True))
                        except Exception as e:
                            st.warning(f"Could not load sheet '{sheet_name}': {e}")

                    # Daily total
                    if sheets["daily"]:
                        try:
                            ws = sh.worksheet(sheets["daily"])
                            records = ws.get_all_records(expected_headers=[], numericise_ignore=["all"])
                            if records:
                                df = pd.DataFrame(records)
                                for col in STANDARD_COLS:
                                    if col not in df.columns:
                                        df[col] = np.nan
                                df = df[STANDARD_COLS].copy()
                                for col in NUMERIC_COLS:
                                    df[col] = clean_numeric(df[col])
                                # Parse dates with Chinese format support
                                df["Date"] = df["Date"].apply(parse_chinese_date)
                                df = df.dropna(subset=["Date"])
                                df = df[df["Spend"].notna()]
                                df["Campaign"] = "Total"
                                df["Region"] = region
                                df["RegionKey"] = normalize_region(region)
                                all_daily.append(df.reset_index(drop=True))
                        except Exception as e:
                            st.warning(f"Could not load daily sheet '{sheets['daily']}': {e}")

                    # Weekly
                    if sheets["weekly"]:
                        try:
                            ws = sh.worksheet(sheets["weekly"])
                            raw_vals = ws.get_all_values()
                            if raw_vals:
                                raw = pd.DataFrame(raw_vals)
                                records_w = []
                                current_month = None
                                for _, row in raw.iterrows():
                                    first = row[0]
                                    if not first or str(first).strip() == "":
                                        continue
                                    first_str = str(first).strip()
                                    if first_str in MONTH_NAMES:
                                        current_month = first_str
                                        continue
                                    if first_str == "Week":
                                        continue
                                    if first_str.startswith("Week"):
                                        records_w.append({
                                            "Month": current_month,
                                            "Week": first_str,
                                            "Spend": row[1] if len(row) > 1 else np.nan,
                                            "Leads": row[2] if len(row) > 2 else np.nan,
                                            "Avg CPL": row[3] if len(row) > 3 else np.nan,
                                        })
                                if records_w:
                                    dfw = pd.DataFrame(records_w)
                                    for col in ["Spend", "Leads", "Avg CPL"]:
                                        dfw[col] = clean_numeric(dfw[col])
                                    dfw = dfw.dropna(subset=["Spend", "Leads", "Avg CPL"], how="all")
                                    if not dfw.empty:
                                        month_order = {m: i for i, m in enumerate(MONTH_NAMES)}
                                        dfw["MonthOrder"] = dfw["Month"].map(month_order)
                                        dfw["WeekNum"] = dfw["Week"].str.extract(r"(\d+)").astype(int)
                                        dfw = dfw.sort_values(["MonthOrder", "WeekNum"]).reset_index(drop=True)
                                        dfw["Period"] = dfw["Month"].str[:3] + " " + dfw["Week"]
                                        dfw["Region"] = region
                                        dfw["RegionKey"] = normalize_region(region)
                                        all_weekly.append(dfw)
                        except Exception as e:
                            st.warning(f"Could not load weekly sheet '{sheets['weekly']}': {e}")

                temp_campaigns_df = pd.concat(all_campaigns, ignore_index=True) if all_campaigns else pd.DataFrame()
                temp_daily_total_df = pd.concat(all_daily, ignore_index=True) if all_daily else pd.DataFrame()
                temp_weekly_df = pd.concat(all_weekly, ignore_index=True) if all_weekly else pd.DataFrame()

                st.session_state["gs_campaigns"] = temp_campaigns_df
                st.session_state["gs_daily"] = temp_daily_total_df
                st.session_state["gs_weekly"] = temp_weekly_df
                st.session_state["gs_url"] = sheet_url

            if st.session_state["gs_campaigns"].empty:
                st.sidebar.warning("Connected to Google Sheets, but no campaign data was loaded. Check sheet names.")
            else:
                st.sidebar.success(f"Connected to Google Sheets. Loaded {len(st.session_state['gs_campaigns'])} rows.")
        except Exception as e:
            st.sidebar.error(f"Connection failed: {e}")

# Restore from session state on every rerun (after both if/else blocks)
# Only restore if user is still on Google Sheets mode
if data_source == "Connect Google Sheets" and "gs_campaigns" in st.session_state:
    if not st.session_state["gs_campaigns"].empty:
        campaigns_df = st.session_state["gs_campaigns"]
        daily_total_df = st.session_state["gs_daily"]
        weekly_df = st.session_state["gs_weekly"]
        st.sidebar.success(f"✓ Data loaded from Google Sheets ({len(campaigns_df)} rows)")



# --- SQL file ---
st.sidebar.markdown("---")
st.sidebar.subheader("SQL / Sales Leads (optional)")
uploaded_sql_files = st.sidebar.file_uploader(
    "Upload SQL export file(s) (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="sql_files",
)

if campaigns_df.empty:
    st.info("Upload the unified Meta Ads overview Excel file in the sidebar to get started.")
    st.stop()

# ---------------------------------------------------------------------------
# Sidebar filters
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.title("Filters")

regions = sorted(campaigns_df["Region"].unique())
sel_regions = st.sidebar.multiselect("Region", regions, default=regions)
sel_region_keys = {normalize_region(r) for r in sel_regions}

campaigns_list = sorted(campaigns_df["Campaign"].unique())
sel_campaigns = st.sidebar.multiselect("Campaign / Ad set", campaigns_list, default=campaigns_list)

min_date, max_date = campaigns_df["Date"].min(), campaigns_df["Date"].max()
date_range = st.sidebar.date_input("Date range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

# SQL settings
raw_sales_df = pd.DataFrame()
if uploaded_sql_files:
    sales_dfs = []
    for f in uploaded_sql_files:
        try:
            sales_dfs.append(load_sales_file(f.getvalue(), f.name))
        except Exception as exc:
            st.sidebar.error(f"Failed to parse SQL file '{f.name}': {exc}")
    raw_sales_df = pd.concat(sales_dfs, ignore_index=True) if sales_dfs else pd.DataFrame()

if not raw_sales_df.empty:
    st.sidebar.markdown("---")
    st.sidebar.subheader("SQL Settings")
    sql_date_basis = st.sidebar.selectbox(
        "SQL date basis",
        ["MQL Cohort Date", "SQL Date", "Created Time", "Converted Time"],
        index=0,
    )
    exclude_duplicates = st.sidebar.checkbox("Exclude leads flagged as duplicates", value=True)
    show_unmapped_sql = st.sidebar.checkbox("Show unmapped SQL leads", value=True)
else:
    sql_date_basis = "MQL Cohort Date"
    exclude_duplicates = True
    show_unmapped_sql = True

# ---------------------------------------------------------------------------
# Filter main data
# ---------------------------------------------------------------------------
mask = (
    campaigns_df["Region"].isin(sel_regions)
    & campaigns_df["Campaign"].isin(sel_campaigns)
    & (campaigns_df["Date"] >= pd.Timestamp(start_date))
    & (campaigns_df["Date"] <= pd.Timestamp(end_date))
)
fdf = campaigns_df[mask].copy()

if fdf.empty:
    st.warning("No data for the current filters. Adjust the sidebar selections.")
    st.stop()

# Map SQL leads
sales_df = pd.DataFrame()
sales_fdf = pd.DataFrame()
if not raw_sales_df.empty:
    sales_df = raw_sales_df.copy()
    sales_df["Campaign"] = sales_df["Raw Adset"].apply(lambda x: map_raw_adset_to_campaign(x, campaigns_list))
    sales_df["SQL Analysis Date"] = sales_df[sql_date_basis]
    sales_mask = (
        sales_df["RegionKey"].isin(sel_region_keys)
        & sales_df["Campaign"].isin(sel_campaigns)
        & (sales_df["SQL Analysis Date"] >= pd.Timestamp(start_date))
        & (sales_df["SQL Analysis Date"] <= pd.Timestamp(end_date))
    )
    if exclude_duplicates and "Duplicate Flag" in sales_df.columns:
        sales_mask &= sales_df["Duplicate Flag"].fillna("").astype(str).str.strip().ne("已重复")
    sales_fdf = sales_df[sales_mask].copy()

# ---------------------------------------------------------------------------
# KPI calculation (daily deltas from cumulative data)
# ---------------------------------------------------------------------------
kpi_data = fdf.copy().sort_values(["Region", "Campaign", "Date"])
kpi_data["Daily Spend"] = kpi_data.groupby(["Region", "Campaign"])["Spend"].diff()
kpi_data["Daily Leads"] = kpi_data.groupby(["Region", "Campaign"])["Leads"].diff()
kpi_data["Daily Leads"] = kpi_data["Daily New Leads"].fillna(kpi_data["Daily Leads"])

total_spend = kpi_data["Daily Spend"].sum()
total_leads = kpi_data["Daily Leads"].sum()
blended_cpl = total_spend / total_leads if total_leads else np.nan
avg_ctr = fdf["CTR"].mean()
avg_cvr = fdf["Lead CVR"].mean()

# ---------------------------------------------------------------------------
# Header + KPIs
# ---------------------------------------------------------------------------
st.title("Meta Ads Dashboard")

if not sales_fdf.empty:
    total_sql = sales_fdf.loc[sales_fdf["Is SQL"], "Lead ID"].nunique()
    sql_mql_cvr = total_sql / total_leads if total_leads else np.nan
    cost_per_sql = total_spend / total_sql if total_sql else np.nan
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("Total Spend", f"${total_spend:,.0f}")
    c2.metric("MQL Leads", f"{total_leads:,.0f}")
    c3.metric("SQL", f"{total_sql:,.0f}")
    c4.metric("SQL / MQL", fmt_percent(sql_mql_cvr))
    c5.metric("Cost / SQL", fmt_number(cost_per_sql, 2, prefix="$"))
    c6.metric("Blended CPL", fmt_number(blended_cpl, 2, prefix="$"))
    c7.metric("Avg CTR", fmt_percent(avg_ctr))
else:
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Spend", f"${total_spend:,.0f}")
    c2.metric("Total Leads / MQL", f"{total_leads:,.0f}")
    c3.metric("Blended CPL", fmt_number(blended_cpl, 2, prefix="$"))
    c4.metric("Avg CTR", fmt_percent(avg_ctr))
    c5.metric("Avg Lead CVR", fmt_percent(avg_cvr))

st.markdown("---")

# ---------------------------------------------------------------------------
# 1. Daily trend
# ---------------------------------------------------------------------------
st.subheader("1. Daily Trend")

col_metric, col_view = st.columns([2, 1])
with col_metric:
    trend_metric = st.selectbox(
        "Metric",
        ["Spend", "Leads", "CPL", "CTR", "Lead CVR", "Frequency", "3-Day Avg CPL"],
        key="trend_metric",
    )
with col_view:
    trend_view = st.selectbox(
        "View",
        ["By region (aggregated)", "By campaign within region"],
        key="trend_view",
    )

if trend_view == "By region (aggregated)":
    trend_data = (
        fdf.groupby(["Date", "Region"])
        .agg(Spend=("Spend", "sum"), Leads=("Leads", "sum"),
             Daily_New_Leads=("Daily New Leads", "sum"),
             Impressions=("Impressions", "sum"), Link_Clicks=("Link Clicks", "sum"),
             CTR=("CTR", "mean"), Lead_CVR=("Lead CVR", "mean"),
             Frequency=("Frequency", "mean"), Avg_CPL=("3-Day Avg CPL", "mean"))
        .reset_index()
    )
    trend_data = trend_data.sort_values(["Region", "Date"])
    trend_data["Daily Spend"] = trend_data.groupby("Region")["Spend"].diff()
    trend_data["Daily Leads"] = trend_data.groupby("Region")["Leads"].diff()
    trend_data["Daily Leads"] = trend_data["Daily_New_Leads"].fillna(trend_data["Daily Leads"])
    trend_data["CPL"] = trend_data["Daily Spend"] / trend_data["Daily Leads"].replace(0, np.nan)
    trend_data["Lead CVR"] = trend_data["Lead_CVR"]
    trend_data["3-Day Avg CPL"] = trend_data["Avg_CPL"]
    trend_data["Series"] = trend_data["Region"]
    title_suffix = "(all campaigns aggregated per region)"
else:
    trend_data = fdf.copy().sort_values(["Region", "Campaign", "Date"])
    trend_data["Daily Spend"] = trend_data.groupby(["Region", "Campaign"])["Spend"].diff()
    trend_data["Daily Leads"] = trend_data.groupby(["Region", "Campaign"])["Leads"].diff()
    trend_data["Daily Leads"] = trend_data["Daily New Leads"].fillna(trend_data["Daily Leads"])
    trend_data["CPL"] = trend_data["Daily Spend"] / trend_data["Daily Leads"].replace(0, np.nan)
    trend_data["Series"] = trend_data["Region"] + " – " + trend_data["Campaign"]
    title_suffix = "(campaign detail)"

metric_col_map = {
    "Spend": "Daily Spend", "Leads": "Daily Leads", "CPL": "CPL",
    "CTR": "CTR", "Lead CVR": "Lead CVR", "Frequency": "Frequency",
    "3-Day Avg CPL": "3-Day Avg CPL",
}
display_col = metric_col_map.get(trend_metric, trend_metric)

fig_trend = px.line(
    trend_data.sort_values("Date"), x="Date", y=display_col, color="Series",
    title=f"{trend_metric} – daily trend {title_suffix}",
)
st.plotly_chart(fig_trend, use_container_width=True)

# ---------------------------------------------------------------------------
# 2. Campaign comparison
# ---------------------------------------------------------------------------
st.subheader("2. Campaign / Ad set Comparison")

agg = (
    kpi_data.groupby(["Region", "RegionKey", "Campaign"])
    .agg(Spend=("Daily Spend", "sum"), Leads=("Daily Leads", "sum"),
         Impressions=("Impressions", "sum"), Link_Clicks=("Link Clicks", "sum"),
         CTR=("CTR", "mean"), Lead_CVR=("Lead CVR", "mean"))
    .reset_index()
)
agg["Blended CPL"] = agg["Spend"] / agg["Leads"].replace(0, np.nan)
agg["CPC"] = agg["Spend"] / agg["Link_Clicks"].replace(0, np.nan)

colA, colB = st.columns(2)
with colA:
    fig_cpl = px.bar(agg, x="Campaign", y="Blended CPL", color="Region", barmode="group",
                     title="CPL by campaign (Spend ÷ MQL Leads)", text_auto=".2f")
    st.plotly_chart(fig_cpl, use_container_width=True)
with colB:
    fig_spend = px.bar(agg, x="Campaign", y="Spend", color="Region", barmode="group",
                       title="Spend by campaign", text_auto=".0f")
    st.plotly_chart(fig_spend, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    fig_ctr = px.bar(agg, x="Campaign", y="CTR", color="Region", barmode="group",
                     title="Avg CTR by campaign", text_auto=".2%")
    fig_ctr.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig_ctr, use_container_width=True)
with colD:
    fig_cvr = px.bar(agg, x="Campaign", y="Lead_CVR", color="Region", barmode="group",
                     title="Avg Lead CVR by campaign", text_auto=".2%")
    fig_cvr.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig_cvr, use_container_width=True)

# ---------------------------------------------------------------------------
# 3. SQL / MQL conversion
# ---------------------------------------------------------------------------
st.subheader("3. SQL / MQL Conversion")

if raw_sales_df.empty:
    st.info("Upload a SQL sales export file in the sidebar to see conversion analysis.")
elif sales_fdf.empty:
    st.warning("SQL file loaded, but no leads match the current region / campaign / date filters.")
else:
    sql_only = sales_fdf[sales_fdf["Is SQL"]].copy()
    sql_agg = (
        sql_only.groupby(["RegionKey", "Campaign"])
        .agg(SQL=("Lead ID", "nunique"),
             Raw_Adsets=("Raw Adset Clean", lambda x: " | ".join(sorted(set(str(v) for v in x if str(v).strip()))[:5])),
             Countries=("Country", lambda x: " | ".join(sorted(set(str(v) for v in x if str(v).strip() and str(v) != "nan"))[:5])))
        .reset_index()
    )
    crm_mql_agg = (
        sales_fdf[sales_fdf["Is CRM MQL"]]
        .groupby(["RegionKey", "Campaign"])
        .agg(CRM_MQL=("Lead ID", "nunique"))
        .reset_index()
    )
    conversion = agg.merge(sql_agg, on=["RegionKey", "Campaign"], how="left")
    conversion = conversion.merge(crm_mql_agg, on=["RegionKey", "Campaign"], how="left")
    conversion["SQL"] = conversion["SQL"].fillna(0).astype(int)
    conversion["CRM_MQL"] = conversion["CRM_MQL"].fillna(0).astype(int)
    conversion["SQL / MQL"] = conversion["SQL"] / conversion["Leads"].replace(0, np.nan)
    conversion["Cost / SQL"] = conversion["Spend"] / conversion["SQL"].replace(0, np.nan)
    conversion["MQL CPL"] = conversion["Spend"] / conversion["Leads"].replace(0, np.nan)
    conversion["Raw_Adsets"] = conversion["Raw_Adsets"].fillna("")
    conversion["Countries"] = conversion["Countries"].fillna("")

    col_sql_a, col_sql_b = st.columns(2)
    with col_sql_a:
        chart_df = conversion.copy()
        chart_df["Series"] = chart_df["Region"] + " – " + chart_df["Campaign"]
        fig_sql_rate = px.bar(chart_df.sort_values("SQL / MQL", ascending=False),
                              x="Series", y="SQL / MQL", color="Region",
                              title="SQL / MQL rate by campaign", text_auto=".2%")
        fig_sql_rate.update_yaxes(tickformat=".1%")
        fig_sql_rate.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig_sql_rate, use_container_width=True)
    with col_sql_b:
        fig_sql_count = px.bar(chart_df.sort_values("SQL", ascending=False),
                               x="Series", y="SQL", color="Region",
                               title="SQL count by campaign", text_auto=".0f")
        fig_sql_count.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig_sql_count, use_container_width=True)

    display_cols = conversion[[
        "Region", "Campaign", "Spend", "Leads", "SQL", "SQL / MQL",
        "MQL CPL", "Cost / SQL", "CTR", "Lead_CVR", "Raw_Adsets", "Countries",
    ]].rename(columns={
        "Leads": "MQL Leads (Meta)", "Lead_CVR": "Lead CVR",
        "Raw_Adsets": "SQL Source Adsets", "Countries": "SQL Countries",
    })
    display_cols = display_cols.sort_values(["Region", "SQL / MQL"], ascending=[True, False])
    st.dataframe(
        display_cols.style.format({
            "Spend": "${:,.0f}", "MQL Leads (Meta)": "{:,.0f}", "SQL": "{:,.0f}",
            "SQL / MQL": "{:.2%}", "MQL CPL": "${:,.2f}", "Cost / SQL": "${:,.2f}",
            "CTR": "{:.2%}", "Lead CVR": "{:.2%}",
        }),
        use_container_width=True, hide_index=True,
    )
    st.download_button("Download SQL/MQL table (CSV)",
                       display_cols.to_csv(index=False).encode("utf-8-sig"),
                       file_name="sql_mql_by_campaign.csv", mime="text/csv")

    daily_sql = (
        sql_only.dropna(subset=["SQL Analysis Date"])
        .assign(Date=lambda x: x["SQL Analysis Date"].dt.date)
        .groupby(["Date", "Inferred Region", "Campaign"])
        .agg(SQL=("Lead ID", "nunique"))
        .reset_index()
    )
    if not daily_sql.empty:
        daily_sql["Series"] = daily_sql["Inferred Region"] + " – " + daily_sql["Campaign"]
        fig_daily_sql = px.line(daily_sql.sort_values("Date"), x="Date", y="SQL", color="Series",
                                title=f"Daily SQL trend ({sql_date_basis})", markers=True)
        st.plotly_chart(fig_daily_sql, use_container_width=True)

    if show_unmapped_sql:
        all_sql_mask = sales_df["Is SQL"].fillna(False)
        unmapped = sales_df[all_sql_mask & (sales_df["Campaign"].eq("Unmapped") | sales_df["RegionKey"].eq("UNKNOWN"))].copy()
        if exclude_duplicates and "Duplicate Flag" in unmapped.columns:
            unmapped = unmapped[unmapped["Duplicate Flag"].fillna("").astype(str).str.strip().ne("已重复")]
        if not unmapped.empty:
            st.markdown("**Unmapped SQL leads (check adset naming):**")
            st.dataframe(
                unmapped[["Source File", "Lead ID", "Raw Adset", "Campaign", "Inferred Region",
                           "Country", "Lead Stage", "Created Time", "Converted Time", "Duplicate Flag"]]
                .sort_values(["Created Time", "Converted Time"], ascending=False),
                use_container_width=True, hide_index=True,
            )

# ---------------------------------------------------------------------------
# 4. Region comparison
# ---------------------------------------------------------------------------
st.subheader("4. Region Comparison")

region_agg = (
    kpi_data.groupby("Region")
    .agg(Spend=("Daily Spend", "sum"), Leads=("Daily Leads", "sum"))
    .reset_index()
)
region_agg["Blended CPL"] = region_agg["Spend"] / region_agg["Leads"].replace(0, np.nan)

if not sales_fdf.empty:
    region_sql = (
        sales_fdf[sales_fdf["Is SQL"]]
        .groupby("RegionKey")
        .agg(SQL=("Lead ID", "nunique"))
        .reset_index()
    )
    region_agg["RegionKey"] = region_agg["Region"].apply(normalize_region)
    region_agg = region_agg.merge(region_sql, on="RegionKey", how="left")
    region_agg["SQL"] = region_agg["SQL"].fillna(0).astype(int)
    region_agg["SQL / MQL"] = region_agg["SQL"] / region_agg["Leads"].replace(0, np.nan)

colE, colF, colG = st.columns(3)
with colE:
    st.plotly_chart(px.pie(region_agg, names="Region", values="Spend", title="Spend share by region"),
                    use_container_width=True)
with colF:
    st.plotly_chart(px.pie(region_agg, names="Region", values="Leads", title="MQL Leads share by region"),
                    use_container_width=True)
with colG:
    if not sales_fdf.empty and "SQL / MQL" in region_agg.columns:
        fig_r_sql = px.bar(region_agg, x="Region", y="SQL / MQL", title="SQL / MQL by region", text_auto=".2%")
        fig_r_sql.update_yaxes(tickformat=".1%")
        st.plotly_chart(fig_r_sql, use_container_width=True)
    else:
        fig_r_cpl = px.bar(region_agg, x="Region", y="Blended CPL", title="Blended CPL by region", text_auto=".2f")
        st.plotly_chart(fig_r_cpl, use_container_width=True)

# ---------------------------------------------------------------------------
# 5. Daily total trend
# ---------------------------------------------------------------------------
if not daily_total_df.empty:
    st.subheader("5. Daily Total Trend (per region)")
    dt_mask = (
        daily_total_df["Region"].isin(sel_regions)
        & (daily_total_df["Date"] >= pd.Timestamp(start_date))
        & (daily_total_df["Date"] <= pd.Timestamp(end_date))
    )
    dt = daily_total_df[dt_mask].copy().sort_values(["Region", "Date"])
    dt["Daily Spend"] = dt.groupby("Region")["Spend"].diff()
    dt["Daily Leads"] = dt.groupby("Region")["Leads"].diff()
    dt["Daily CPL"] = dt["Daily Spend"] / dt["Daily Leads"].replace(0, np.nan)

    st.plotly_chart(px.line(dt.sort_values("Date"), x="Date", y="Daily Spend", color="Region",
                            title="Daily spend (actual daily, not cumulative)", markers=True),
                    use_container_width=True)
    col_dt_a, col_dt_b = st.columns(2)
    with col_dt_a:
        st.plotly_chart(px.line(dt.sort_values("Date"), x="Date", y="Daily Leads", color="Region",
                                title="Daily new leads", markers=True), use_container_width=True)
    with col_dt_b:
        st.plotly_chart(px.line(dt.sort_values("Date"), x="Date", y="Daily CPL", color="Region",
                                title="Daily CPL", markers=True), use_container_width=True)

# ---------------------------------------------------------------------------
# 6. Weekly summary
# ---------------------------------------------------------------------------
if not weekly_df.empty:
    st.subheader("6. Weekly Summary")
    wdf = weekly_df[weekly_df["Region"].isin(sel_regions)]
    if not wdf.empty:
        fig_weekly = go.Figure()
        for region, g in wdf.groupby("Region"):
            fig_weekly.add_trace(go.Bar(x=g["Period"], y=g["Spend"], name=f"{region} Spend"))
            fig_weekly.add_trace(go.Scatter(x=g["Period"], y=g["Avg CPL"],
                                            name=f"{region} Avg CPL", yaxis="y2"))
        fig_weekly.update_layout(
            title="Weekly Spend & Avg CPL",
            yaxis=dict(title="Spend"),
            yaxis2=dict(title="Avg CPL", overlaying="y", side="right"),
        )
        st.plotly_chart(fig_weekly, use_container_width=True)
        st.dataframe(wdf[["Region", "Period", "Spend", "Leads", "Avg CPL"]], use_container_width=True)

# ---------------------------------------------------------------------------
# 7. Key actions & alerts
# ---------------------------------------------------------------------------
st.subheader("7. Key Actions & Status Alerts")

notes = fdf[fdf["Key actions"].notna() & (fdf["Key actions"].astype(str).str.strip() != "")]
st.markdown("**Key actions log:**")
st.dataframe(notes[["Date", "Region", "Campaign", "Status", "CPL", "Key actions"]]
             .sort_values("Date", ascending=False), use_container_width=True, hide_index=True)

alert = fdf[fdf["Status"].str.contains("Fatigue", na=False)]
if not alert.empty:
    st.markdown("**Audience Fatigue alerts (latest 5):**")
    st.dataframe(alert.sort_values("Date", ascending=False)
                 [["Date", "Region", "Campaign", "Spend", "Leads", "CPL"]].head(5),
                 use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# 8. Raw data export
# ---------------------------------------------------------------------------
with st.expander("Raw data / Export"):
    st.dataframe(fdf, use_container_width=True)
    st.download_button("Download filtered Meta data (CSV)",
                       fdf.to_csv(index=False).encode("utf-8-sig"),
                       file_name="meta_ads_filtered.csv", mime="text/csv")

if not sales_fdf.empty:
    with st.expander("SQL leads detail / Export"):
        detail_cols = ["Source File", "Lead ID", "Raw Adset", "Campaign", "Inferred Region",
                       "Country", "Lead Stage", "Created Time", "Converted Time",
                       "SQL Analysis Date", "Duplicate Flag"]
        st.dataframe(sales_fdf[detail_cols].sort_values("SQL Analysis Date", ascending=False),
                     use_container_width=True, hide_index=True)
        st.download_button("Download SQL leads (CSV)",
                           sales_fdf[detail_cols].to_csv(index=False).encode("utf-8-sig"),
                           file_name="sql_leads_filtered.csv", mime="text/csv")
