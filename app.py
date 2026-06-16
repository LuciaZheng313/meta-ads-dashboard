"""
Meta Ads 多地区看板 (EMEA / India / North America ...)
运行方式: streamlit run app.py
依赖: pip install streamlit pandas plotly openpyxl python-calamine numpy --break-system-packages

新增：支持上传销售线索 SQL 导出文件，并按 Region + Campaign / Ad set 计算 SQL / MQL 转化。
SQL 文件建议命名：SQL_销售线索_YYYYMMDD.xlsx，例如 SQL_销售线索_20260612.xlsx

新增 (NA 看板兼容)：
- 部分新导出的 Excel 文件 (例如 NA 看板) 的内部样式信息有损坏, openpyxl 会在打开时报错
  (TypeError: expected <class 'openpyxl.styles.fills.Fill'>)。代码会自动尝试 openpyxl,
  失败时回退到 calamine 引擎 (pip install python-calamine)。
- "Daily Total" / "Daily total" 等大小写写法都会被识别为汇总 sheet。
- 新增 "Single Image" / "Carousel" 等 ad set sheet 会被自动识别为 campaign 明细 sheet
  (只要包含 Date / Spend 列), 并补充了对应的销售线索别名规则。
"""

import re
import io
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Meta Ads 多地区看板", layout="wide")

# ---------------------------------------------------------------------------
# 配置: 各 campaign sheet 名称 + 统一字段
# ---------------------------------------------------------------------------
# 现有看板常用 sheet。为了兼容后续 NA 看板，代码会额外自动读取其他包含 Date/Spend 的 sheet。
PREFERRED_CAMPAIGN_SHEETS = ["Manufacture", "CA", "AI Instant", "AI Layout", "Single Image", "Carousel"]
DAILY_TOTAL_SHEET = "Daily total"
# NA 看板等文件里 sheet 名写作 "Daily Total" (大写 T)，统一用下面的候选列表做大小写无关匹配。
DAILY_TOTAL_SHEET_CANDIDATES = ["Daily total", "Daily Total", "Daily Totals", "daily total"]
WEEKLY_SHEET = "Weekly"
EXCLUDED_SHEETS = {DAILY_TOTAL_SHEET, WEEKLY_SHEET, "hidden0", "隐藏", "说明", "README", "readme"}

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

# 销售线索导出表常见字段。若后续字段名有变化，只需在这里补 candidate。
SALES_SHEET_CANDIDATES = ["销售线索数据"]
SALES_ADSET_COLS = [
    "Lead Source Drill-Down 3",
    "Lead Source Drill-Down 3(Old)",
    "营销活动",
    "市场活动（text）",
    "市场活动",
]
SALES_STAGE_COLS = ["线索阶段", "Lifecycle Stage", "生命状态"]
SALES_CREATED_TIME_COLS = ["创建时间", "创建日期", "网站注册用户的提交日期"]
SALES_CONVERTED_TIME_COLS = ["转换时间", "线索阶段变更时间"]
SALES_MQL_TIME_COLS = ["转MQL时间"]
SALES_COUNTRY_COLS = ["Country / Region of Lead", "Country （文本）", "国家", "手机归属国家"]
SALES_DUPLICATE_COLS = ["是否存在重复数据"]
SALES_ID_COLS = ["线索编号（必填）", "唯一标识", "客户ID", "Email", "邮件", "电话", "Telephone", "Mobile", "手机"]

# 用销售侧 Raw Ad set 名称归类到 Meta 看板里的 Campaign / Ad set。
# 如果后续 NA 看板有更准确的 sheet 名称，代码会优先用 sheet 名称做 exact/contains 匹配。
CAMPAIGN_ALIAS_RULES = {
    "Manufacture": ["manufacture", "manufacturing", "manufacturer"],
    "CA": ["customer acquisition", "customers acquisition", "sales acquisition", "acquisition"],
    "AI Instant": ["ai instant", "ai design", "ai render", "ai rendering"],
    "AI Layout": ["ai layout"],
    "KC Tools": ["kc tools", "kc tool", "kctools", "kc_tool"],
    "SmartLinkSuite": ["smartlinksuite", "smart link suite", "smartlink suite"],
    "Single Image": ["single image", "single-image", "singleimage", "static image", "image ad"],
    "Carousel": ["carousel", "carrousel", "carousal"],
}

EMEA_COUNTRIES = {
    "egypt", "saudi arabia", "united arab emirates", "uae", "south africa",
    "italy", "namibia", "united kingdom", "uk", "kenya", "sri lanka",
    "turkey", "germany", "france", "spain", "qatar", "kuwait", "oman", "bahrain",
    "埃及", "沙特阿拉伯", "阿联酋", "南非", "意大利", "纳米比亚", "英国", "肯尼亚",
}

INDIA_COUNTRIES = {"india", "印度"}
NA_COUNTRIES = {"united states", "usa", "us", "canada", "mexico", "美国", "加拿大", "墨西哥"}


# ---------------------------------------------------------------------------
# 通用清洗函数
# ---------------------------------------------------------------------------
def clean_numeric(series: pd.Series) -> pd.Series:
    """把 #REF! 等异常值统一转成 NaN, 并转成数值型。"""
    if series is None:
        return pd.Series(dtype="float64")
    return pd.to_numeric(series.replace("#REF!", np.nan), errors="coerce")


def normalize_key(value) -> str:
    """用于宽松匹配的字符串标准化。"""
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s\-_]+", " ", text)
    text = re.sub(r"[^a-z0-9\u4e00-\u9fff ]+", "", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_region(value) -> str:
    """统一 Region key，避免 EMEA / emea / North America / US 无法 join。"""
    key = normalize_key(value)
    if not key:
        return "UNKNOWN"
    if key in {"india", "in", "印度"}:
        return "INDIA"
    if key in {"na", "north america", "us", "usa", "united states", "united states of america", "america", "美国"}:
        return "NA"
    if key in {"emea", "europe middle east africa", "middle east", "gcc"}:
        return "EMEA"
    return str(value).strip().upper()


def first_existing_col(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    """返回 df 中第一个存在的候选字段。"""
    for col in candidates:
        if col in df.columns:
            return col
    return None


def coalesce_columns(df: pd.DataFrame, candidates: Iterable[str], default=np.nan) -> pd.Series:
    """按候选字段顺序合并非空值。"""
    existing = [c for c in candidates if c in df.columns]
    if not existing:
        return pd.Series([default] * len(df), index=df.index)

    result = df[existing[0]].copy()
    for col in existing[1:]:
        result = result.where(result.notna() & (result.astype(str).str.strip() != ""), df[col])
    return result


def safe_div(numerator, denominator):
    denominator = denominator.replace(0, np.nan) if isinstance(denominator, pd.Series) else denominator
    return numerator / denominator


def fmt_number(value, decimals=0, prefix="", suffix="") -> str:
    if pd.isna(value) or value == np.inf or value == -np.inf:
        return "-"
    return f"{prefix}{value:,.{decimals}f}{suffix}"


def fmt_percent(value, decimals=2) -> str:
    if pd.isna(value) or value == np.inf or value == -np.inf:
        return "-"
    return f"{value:.{decimals}%}"


def open_excel_file(file_bytes: bytes) -> pd.ExcelFile:
    """打开 Excel 文件。

    部分新导出的 Excel (例如 NA 看板) 内部样式信息有损坏，openpyxl 在打开时会报错
    (TypeError: expected <class 'openpyxl.styles.fills.Fill'>)。这里优先尝试默认引擎
    (openpyxl)，失败后自动回退到 calamine 引擎，对样式问题更宽容。
    """
    try:
        return pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception:
        return pd.ExcelFile(io.BytesIO(file_bytes), engine="calamine")


def find_sheet_name(xls: pd.ExcelFile, *candidates: str) -> str | None:
    """按候选名称做大小写/前后空格无关的 sheet 名称匹配，返回文件中实际的 sheet 名。"""
    normalized = {str(name).strip().lower(): name for name in xls.sheet_names}
    for candidate in candidates:
        key = str(candidate).strip().lower()
        if key in normalized:
            return normalized[key]
    return None


# ---------------------------------------------------------------------------
# Meta Ads 看板加载函数
# ---------------------------------------------------------------------------
def get_non_campaign_sheet_names(xls: pd.ExcelFile) -> set[str]:
    """计算应排除在 campaign/ad set 之外的 sheet 名称(汇总表、说明等)。"""
    excluded = set(EXCLUDED_SHEETS)
    daily_total_sheet = find_sheet_name(xls, *DAILY_TOTAL_SHEET_CANDIDATES)
    if daily_total_sheet:
        excluded.add(daily_total_sheet)
    weekly_sheet = find_sheet_name(xls, WEEKLY_SHEET)
    if weekly_sheet:
        excluded.add(weekly_sheet)
    return excluded


def looks_like_campaign_sheet(xls: pd.ExcelFile, sheet_name: str, excluded_sheets: set[str]) -> bool:
    """判断一个 sheet 是否像 campaign/ad set 明细 sheet。"""
    if sheet_name in excluded_sheets:
        return False
    if sheet_name.lower().startswith("hidden"):
        return False
    try:
        preview = pd.read_excel(xls, sheet_name=sheet_name, nrows=1)
    except Exception:
        return False
    return "Date" in preview.columns and "Spend" in preview.columns


def get_campaign_sheet_names(xls: pd.ExcelFile) -> list[str]:
    """优先读取固定 campaign sheet，同时自动兼容后续新增的 NA/ad set sheet。"""
    excluded_sheets = get_non_campaign_sheet_names(xls)
    sheets = []
    for sheet in PREFERRED_CAMPAIGN_SHEETS:
        if sheet in xls.sheet_names and sheet not in excluded_sheets:
            sheets.append(sheet)
    for sheet in xls.sheet_names:
        if sheet not in sheets and looks_like_campaign_sheet(xls, sheet, excluded_sheets):
            sheets.append(sheet)
    return sheets


def load_campaign_sheet(xls: pd.ExcelFile, sheet_name: str, region: str, campaign: str) -> pd.DataFrame:
    """读取单个 campaign / Daily total sheet, 统一字段、清洗数据。"""
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # 补齐缺失列(例如 CA 缺 CPC), 多余列忽略
    for col in STANDARD_COLS:
        if col not in df.columns:
            df[col] = np.nan
    df = df[STANDARD_COLS].copy()

    for col in NUMERIC_COLS:
        df[col] = clean_numeric(df[col])

    # CPC 缺失但有 Spend / Link Clicks 时, 用 Spend/Clicks 补算
    mask = df["CPC"].isna() & df["Spend"].notna() & df["Link Clicks"].notna() & (df["Link Clicks"] != 0)
    df.loc[mask, "CPC"] = df.loc[mask, "Spend"] / df.loc[mask, "Link Clicks"]

    # 源表里部分 sheet 的 CTR 可能是公式引用错误；统一按 Link Clicks / Impressions 重算。
    clicks_imp_mask = df["Impressions"].notna() & (df["Impressions"] != 0) & df["Link Clicks"].notna()
    df.loc[clicks_imp_mask, "CTR"] = df.loc[clicks_imp_mask, "Link Clicks"] / df.loc[clicks_imp_mask, "Impressions"]

    # 去掉日期为空 / Spend 为空的尾部空行
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df = df[df["Spend"].notna()]

    df["Campaign"] = campaign
    df["Region"] = region
    df["RegionKey"] = normalize_region(region)
    return df.reset_index(drop=True)


def load_weekly_sheet(xls: pd.ExcelFile, region: str, sheet_name: str = WEEKLY_SHEET) -> pd.DataFrame:
    """解析多月份分块的 Weekly sheet, 自动跳过 #REF! / 空白的早期周。"""
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
            records.append({
                "Month": current_month,
                "Week": first_str,
                "Spend": row[1],
                "Leads": row[2],
                "Avg CPL": row[3],
            })

    df = pd.DataFrame(records)
    if df.empty:
        df["Region"] = region
        df["RegionKey"] = normalize_region(region)
        df["Period"] = None
        return df

    for col in ["Spend", "Leads", "Avg CPL"]:
        df[col] = clean_numeric(df[col])

    # 自动剔除 #REF!/空白周
    df = df.dropna(subset=["Spend", "Leads", "Avg CPL"], how="all").reset_index(drop=True)

    month_order = {m: i for i, m in enumerate(MONTH_NAMES)}
    df["MonthOrder"] = df["Month"].map(month_order)
    df["WeekNum"] = df["Week"].str.extract(r"(\d+)").astype(int)
    df = df.sort_values(["MonthOrder", "WeekNum"]).reset_index(drop=True)

    df["Period"] = df["Month"].str[:3] + " " + df["Week"]
    df["Region"] = region
    df["RegionKey"] = normalize_region(region)
    return df


def guess_region_name(filename: str) -> str:
    """从文件名猜地区名, 例如 'EMEA地区_Meta_ads_看板.xlsx' -> 'EMEA'。"""
    m = re.match(r"^([A-Za-z]+)", filename)
    if m:
        return m.group(1).upper()
    return filename


@st.cache_data(show_spinner=False)
def load_region_file(file_bytes: bytes, region_name: str):
    xls = open_excel_file(file_bytes)

    campaign_dfs = []
    for sheet in get_campaign_sheet_names(xls):
        campaign_dfs.append(load_campaign_sheet(xls, sheet, region_name, sheet))
    campaigns = pd.concat(campaign_dfs, ignore_index=True) if campaign_dfs else pd.DataFrame()

    daily_total = None
    daily_total_sheet = find_sheet_name(xls, *DAILY_TOTAL_SHEET_CANDIDATES)
    if daily_total_sheet:
        daily_total = load_campaign_sheet(xls, daily_total_sheet, region_name, "Total")

    weekly = None
    weekly_sheet = find_sheet_name(xls, WEEKLY_SHEET)
    if weekly_sheet:
        weekly = load_weekly_sheet(xls, region_name, weekly_sheet)

    return campaigns, daily_total, weekly


# ---------------------------------------------------------------------------
# 销售线索 SQL 文件加载函数
# ---------------------------------------------------------------------------
def pick_sales_sheet(xls: pd.ExcelFile) -> str:
    for sheet in SALES_SHEET_CANDIDATES:
        if sheet in xls.sheet_names:
            return sheet
    for sheet in xls.sheet_names:
        if not sheet.lower().startswith("hidden"):
            return sheet
    return xls.sheet_names[0]


def infer_region_from_sales_row(raw_adset, country) -> str:
    text = normalize_key(raw_adset)
    country_key = normalize_key(country)

    if "emea" in text:
        return "EMEA"
    if "india" in text or country_key in INDIA_COUNTRIES:
        return "INDIA"
    if re.search(r"(^|\s)(us|usa|na)(\s|$)", text) or "united states" in country_key or country_key in NA_COUNTRIES:
        return "NA"
    if country_key in EMEA_COUNTRIES:
        return "EMEA"
    if country_key in INDIA_COUNTRIES:
        return "INDIA"
    if country_key in NA_COUNTRIES:
        return "NA"
    return "UNKNOWN"


def map_raw_adset_to_campaign(raw_adset, loaded_campaigns: Iterable[str] | None = None) -> str:
    """把销售侧 Raw Ad set 映射到看板 Campaign / Ad set。"""
    key = normalize_key(raw_adset)
    if not key:
        return "Unmapped"

    # 1) 优先用已上传 Meta 看板中的 sheet 名称匹配，方便后续 NA 新 sheet 直接生效。
    loaded_campaigns = list(loaded_campaigns or [])
    exact_matches = []
    for campaign in loaded_campaigns:
        ckey = normalize_key(campaign)
        if ckey and (ckey == key or ckey in key):
            exact_matches.append(campaign)
    if exact_matches:
        return sorted(exact_matches, key=lambda x: len(str(x)), reverse=True)[0]

    # 2) 再用别名规则归类。
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
    """读取销售线索导出文件，统一成 SQL/MQL 分析需要的字段。"""
    xls = open_excel_file(file_bytes)
    sheet = pick_sales_sheet(xls)
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
    file_is_sql = "sql" in file_name.lower() or "销售认可" in file_name

    # 如果上传的是 SQL 专用导出，且没有可靠阶段字段，则默认每行都是 SQL。
    has_stage_signal = stage_text.str.contains("sql|mql|销售认可|市场认可", regex=True).any()
    df["Is SQL"] = stage_text.str.contains("sql|销售认可", regex=True)
    if file_is_sql and not has_stage_signal:
        df["Is SQL"] = True

    # CRM MQL：若销售文件未来包含 MQL 阶段，也能展示；SQL 也视为已通过 MQL。
    df["Is CRM MQL"] = stage_text.str.contains("mql|市场认可|sql|销售认可", regex=True)

    df["SQL Date"] = df["Converted Time"].fillna(df["Created Time"])
    df["MQL Cohort Date"] = df["Created Time"].fillna(df["MQL Time"]).fillna(df["Converted Time"])
    df["Inferred Region"] = [infer_region_from_sales_row(a, c) for a, c in zip(df["Raw Adset"], df["Country"])]
    df["RegionKey"] = df["Inferred Region"].apply(normalize_region)
    df["Raw Adset Clean"] = df["Raw Adset"].fillna("Unmapped").astype(str).replace({"nan": "Unmapped"})
    return df


# ---------------------------------------------------------------------------
# Sidebar: 上传文件(支持多个地区 + SQL 文件)
# ---------------------------------------------------------------------------
st.sidebar.title("📂 数据源")
uploaded_files = st.sidebar.file_uploader(
    "上传各地区的 Meta Ads 看板表 (.xlsx)，可多选",
    type=["xlsx"],
    accept_multiple_files=True,
    key="meta_files",
)

uploaded_sql_files = st.sidebar.file_uploader(
    "上传 SQL 销售线索导出文件 (.xlsx，可多选)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="sql_files",
    help="建议命名：SQL_销售线索_YYYYMMDD.xlsx，例如 SQL_销售线索_20260612.xlsx。字段优先读取 Lead Source Drill-Down 3、线索阶段、创建时间、转换时间。",
)

st.sidebar.caption("SQL 文件建议命名：`SQL_销售线索_YYYYMMDD.xlsx`。命名不是强制，代码主要按字段识别。")

if not uploaded_files:
    st.info("⬅️ 请在左侧上传一个或多个地区的 Excel 看板文件(EMEA / India / North America ...)")
    st.stop()

all_campaigns, all_daily_total, all_weekly = [], [], []
for f in uploaded_files:
    default_region = guess_region_name(f.name)
    region_name = st.sidebar.text_input(f"「{f.name}」对应的地区名称", value=default_region, key=f"region_{f.name}")

    campaigns, daily_total, weekly = load_region_file(f.getvalue(), region_name)
    if not campaigns.empty:
        all_campaigns.append(campaigns)
    if daily_total is not None and not daily_total.empty:
        all_daily_total.append(daily_total)
    if weekly is not None and not weekly.empty:
        all_weekly.append(weekly)

campaigns_df = pd.concat(all_campaigns, ignore_index=True) if all_campaigns else pd.DataFrame()
daily_total_df = pd.concat(all_daily_total, ignore_index=True) if all_daily_total else pd.DataFrame()
weekly_df = pd.concat(all_weekly, ignore_index=True) if all_weekly else pd.DataFrame()

if campaigns_df.empty:
    st.error("没有解析到任何 campaign/ad set 数据，请检查表格 sheet 是否包含 Date / Spend 字段。")
    st.stop()

raw_sales_df = pd.DataFrame()
if uploaded_sql_files:
    sales_dfs = []
    for f in uploaded_sql_files:
        try:
            sales_dfs.append(load_sales_file(f.getvalue(), f.name))
        except Exception as exc:
            st.sidebar.error(f"SQL 文件解析失败：{f.name}，原因：{exc}")
    raw_sales_df = pd.concat(sales_dfs, ignore_index=True) if sales_dfs else pd.DataFrame()


# ---------------------------------------------------------------------------
# Sidebar: 筛选器
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.title("🔍 筛选")

regions = sorted(campaigns_df["Region"].unique())
sel_regions = st.sidebar.multiselect("地区 Region", regions, default=regions)
sel_region_keys = {normalize_region(r) for r in sel_regions}

campaigns_list = sorted(campaigns_df["Campaign"].unique())
sel_campaigns = st.sidebar.multiselect("Campaign / Ad set", campaigns_list, default=campaigns_list)

min_date, max_date = campaigns_df["Date"].min(), campaigns_df["Date"].max()
date_range = st.sidebar.date_input(
    "日期范围",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

if not raw_sales_df.empty:
    st.sidebar.markdown("---")
    st.sidebar.title("🧲 SQL 匹配设置")
    sql_date_basis = st.sidebar.selectbox(
        "SQL 日期口径",
        ["MQL Cohort Date", "SQL Date", "Created Time", "Converted Time"],
        index=0,
        help="建议默认用 MQL Cohort Date，即按线索创建日期归因到广告 MQL cohort；若想看销售实际转 SQL 的日期，可切换 SQL Date。",
    )
    exclude_duplicates = st.sidebar.checkbox("排除销售侧标记为「已重复」的线索", value=True)
    show_unmapped_sql = st.sidebar.checkbox("显示未匹配到 Campaign / Ad set 的 SQL", value=True)
else:
    sql_date_basis = "MQL Cohort Date"
    exclude_duplicates = True
    show_unmapped_sql = True

mask = (
    campaigns_df["Region"].isin(sel_regions)
    & campaigns_df["Campaign"].isin(sel_campaigns)
    & (campaigns_df["Date"] >= pd.Timestamp(start_date))
    & (campaigns_df["Date"] <= pd.Timestamp(end_date))
)
fdf = campaigns_df[mask].copy()

if fdf.empty:
    st.warning("当前筛选条件下没有数据，请调整左侧筛选项。")
    st.stop()

# 销售侧：在知道 Meta campaign list 后，再做 Raw Ad set -> Campaign 映射。
sales_df = pd.DataFrame()
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
else:
    sales_fdf = pd.DataFrame()


# ---------------------------------------------------------------------------
# 顶部 KPI
# ---------------------------------------------------------------------------
st.title("📊 Meta Ads 多地区效果看板")

total_spend = fdf["Spend"].sum()
total_leads = fdf["Leads"].sum()
blended_cpl = total_spend / total_leads if total_leads else np.nan
avg_ctr = fdf["CTR"].mean()
avg_cvr = fdf["Lead CVR"].mean()

if not sales_fdf.empty:
    total_sql = sales_fdf.loc[sales_fdf["Is SQL"], "Lead ID"].nunique()
    sql_mql_cvr = total_sql / total_leads if total_leads else np.nan
    cost_per_sql = total_spend / total_sql if total_sql else np.nan

    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("总花费 Spend", f"${total_spend:,.0f}")
    c2.metric("MQL Leads", f"{total_leads:,.0f}")
    c3.metric("SQL", f"{total_sql:,.0f}")
    c4.metric("SQL / MQL", fmt_percent(sql_mql_cvr, 2))
    c5.metric("Cost / SQL", fmt_number(cost_per_sql, 2, prefix="$"))
    c6.metric("综合 CPL", fmt_number(blended_cpl, 2, prefix="$"))
    c7.metric("平均 CTR", fmt_percent(avg_ctr, 2))
else:
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("总花费 Spend", f"${total_spend:,.0f}")
    c2.metric("总线索 Leads / MQL", f"{total_leads:,.0f}")
    c3.metric("综合 CPL", fmt_number(blended_cpl, 2, prefix="$"))
    c4.metric("平均 CTR", fmt_percent(avg_ctr, 2))
    c5.metric("平均 Lead CVR", fmt_percent(avg_cvr, 2))

st.markdown("---")


# ---------------------------------------------------------------------------
# 1. 时间趋势: Spend / Leads / CPL
# ---------------------------------------------------------------------------
st.subheader("1️⃣ 每日趋势 (按地区 + Campaign / Ad set)")

col_trend_metric, col_trend_view = st.columns([2, 1])
with col_trend_metric:
    trend_metric = st.selectbox(
        "选择指标",
        ["Spend", "Leads", "CPL", "CTR", "Lead CVR", "Frequency", "3-Day Avg CPL"],
        key="trend_metric",
    )
with col_trend_view:
    trend_view_mode = st.selectbox(
        "查看模式",
        ["按区域汇总 (Region Total)", "按区域内Campaign对比 (Campaign by Region)"],
        key="trend_view_mode",
    )

# Prepare data based on view mode
if trend_view_mode == "按区域汇总 (Region Total)":
    # Aggregate all campaigns within each region by date
    trend_data = (
        fdf.groupby(["Date", "Region"])
        .agg({
            "Spend": "sum",
            "Leads": "sum",
            "Impressions": "sum",
            "Link Clicks": "sum",
            "CTR": "mean",
            "Lead CVR": "mean",
            "Frequency": "mean",
            "3-Day Avg CPL": "mean",
        })
        .reset_index()
    )
    # Calculate daily delta for Spend (since Excel shows cumulative)
    # For each Region, calculate the difference from previous day
    trend_data = trend_data.sort_values(["Region", "Date"])
    trend_data["Daily Spend"] = trend_data.groupby("Region")["Spend"].diff().fillna(trend_data["Spend"])

    # Recalculate CPL based on aggregated Spend and Leads
    trend_data["CPL"] = trend_data["Spend"] / trend_data["Leads"].replace(0, np.nan)
    trend_data["Series"] = trend_data["Region"]
    title_suffix = "(各区域所有Campaign汇总)"
else:
    # Show individual campaigns within each region
    trend_data = fdf.copy()
    trend_data = trend_data.sort_values(["Region", "Campaign", "Date"])
    # Calculate daily delta for Spend for each Campaign
    trend_data["Daily Spend"] = trend_data.groupby(["Region", "Campaign"])["Spend"].diff().fillna(trend_data["Spend"])

    trend_data["Series"] = trend_data["Region"] + " - " + trend_data["Campaign"]
    title_suffix = "(各区域Campaign明细)"

# Use Daily Spend if user selects Spend metric
display_metric = "Daily Spend" if trend_metric == "Spend" else trend_metric

fig_trend = px.line(
    trend_data.sort_values("Date"),
    x="Date", y=display_metric, color="Series",
    markers=False,
    title=f"{trend_metric} 每日趋势 {title_suffix}",
)
st.plotly_chart(fig_trend, use_container_width=True)


# ---------------------------------------------------------------------------
# 2. Campaign / Ad set 横向对比
# ---------------------------------------------------------------------------
st.subheader("2️⃣ Campaign / Ad set 横向对比")

agg = (
    fdf.groupby(["Region", "RegionKey", "Campaign"])
    .agg(
        Spend=("Spend", "sum"),
        Leads=("Leads", "sum"),
        Impressions=("Impressions", "sum"),
        Link_Clicks=("Link Clicks", "sum"),
        CTR=("CTR", "mean"),
        Lead_CVR=("Lead CVR", "mean"),
    )
    .reset_index()
)
agg["Blended CPL"] = agg["Spend"] / agg["Leads"].replace(0, np.nan)
agg["CPC"] = agg["Spend"] / agg["Link_Clicks"].replace(0, np.nan)

colA, colB = st.columns(2)

with colA:
    fig_cpl = px.bar(
        agg, x="Campaign", y="Blended CPL", color="Region", barmode="group",
        title="各 Campaign / Ad set 综合 CPL 对比 (Spend ÷ MQL Leads)",
        text_auto=".2f",
    )
    st.plotly_chart(fig_cpl, use_container_width=True)

with colB:
    fig_spend = px.bar(
        agg, x="Campaign", y="Spend", color="Region", barmode="group",
        title="各 Campaign / Ad set 花费对比",
        text_auto=".0f",
    )
    st.plotly_chart(fig_spend, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    fig_ctr = px.bar(
        agg, x="Campaign", y="CTR", color="Region", barmode="group",
        title="各 Campaign / Ad set 平均 CTR", text_auto=".2%",
    )
    fig_ctr.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig_ctr, use_container_width=True)

with colD:
    fig_cvr = px.bar(
        agg, x="Campaign", y="Lead_CVR", color="Region", barmode="group",
        title="各 Campaign / Ad set 平均 Lead CVR", text_auto=".2%",
    )
    fig_cvr.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig_cvr, use_container_width=True)


# ---------------------------------------------------------------------------
# 3. SQL / MQL 转化分析
# ---------------------------------------------------------------------------
st.subheader("3️⃣ SQL / MQL 转化分析 (按 Campaign / Ad set)")

if raw_sales_df.empty:
    st.info(
        "上传 SQL 销售线索导出文件后，这里会展示每个 Campaign / Ad set 的 SQL / MQL 转化。"
        "建议文件名：`SQL_销售线索_YYYYMMDD.xlsx`；关键字段：`Lead Source Drill-Down 3`、`线索阶段`、`创建时间`、`转换时间`。"
    )
elif sales_fdf.empty:
    st.warning(
        "已读取 SQL 文件，但当前 Region / Campaign / 日期 / 去重筛选下没有匹配到 SQL 线索。"
        "可以尝试切换 SQL 日期口径，或打开下方未匹配明细检查 Raw Ad set 命名。"
    )
else:
    sql_only = sales_fdf[sales_fdf["Is SQL"]].copy()

    sql_agg = (
        sql_only.groupby(["RegionKey", "Campaign"])
        .agg(
            SQL=("Lead ID", "nunique"),
            Raw_Adsets=("Raw Adset Clean", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()]))[:5])),
            Countries=("Country", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != "nan"]))[:5])),
        )
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

    conversion_display = conversion[[
        "Region", "Campaign", "Spend", "Leads", "SQL", "SQL / MQL", "MQL CPL", "Cost / SQL",
        "CTR", "Lead_CVR", "Raw_Adsets", "Countries",
    ]].rename(columns={
        "Leads": "MQL Leads (Meta)",
        "Lead_CVR": "Lead CVR",
        "Raw_Adsets": "SQL Source Adset Examples",
        "Countries": "SQL Countries",
    })

    conversion_display = conversion_display.sort_values(["Region", "SQL / MQL", "SQL"], ascending=[True, False, False])

    col_sql_a, col_sql_b = st.columns(2)
    with col_sql_a:
        chart_df = conversion.copy()
        chart_df["Series"] = chart_df["Region"] + " - " + chart_df["Campaign"]
        fig_sql_rate = px.bar(
            chart_df.sort_values("SQL / MQL", ascending=False),
            x="Series", y="SQL / MQL", color="Region",
            title="SQL / MQL 转化率 by Campaign / Ad set",
            text_auto=".2%",
        )
        fig_sql_rate.update_yaxes(tickformat=".1%")
        fig_sql_rate.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig_sql_rate, use_container_width=True)

    with col_sql_b:
        fig_sql_count = px.bar(
            chart_df.sort_values("SQL", ascending=False),
            x="Series", y="SQL", color="Region",
            title="SQL 数量 by Campaign / Ad set",
            text_auto=".0f",
        )
        fig_sql_count.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig_sql_count, use_container_width=True)

    st.dataframe(
        conversion_display.style.format({
            "Spend": "${:,.0f}",
            "MQL Leads (Meta)": "{:,.0f}",
            "SQL": "{:,.0f}",
            "SQL / MQL": "{:.2%}",
            "MQL CPL": "${:,.2f}",
            "Cost / SQL": "${:,.2f}",
            "CTR": "{:.2%}",
            "Lead CVR": "{:.2%}",
        }),
        use_container_width=True,
        hide_index=True,
    )

    csv_bytes = conversion_display.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "下载 SQL/MQL 转化表 CSV",
        csv_bytes,
        file_name="sql_mql_by_adset.csv",
        mime="text/csv",
    )

    # SQL 日趋势
    daily_sql = (
        sql_only.dropna(subset=["SQL Analysis Date"])
        .assign(Date=lambda x: x["SQL Analysis Date"].dt.date)
        .groupby(["Date", "Inferred Region", "Campaign"])
        .agg(SQL=("Lead ID", "nunique"))
        .reset_index()
    )
    if not daily_sql.empty:
        daily_sql["Series"] = daily_sql["Inferred Region"] + " - " + daily_sql["Campaign"]
        fig_daily_sql = px.line(
            daily_sql.sort_values("Date"),
            x="Date", y="SQL", color="Series",
            title=f"每日 SQL 趋势 ({sql_date_basis})",
            markers=True,
        )
        st.plotly_chart(fig_daily_sql, use_container_width=True)

    if show_unmapped_sql:
        all_sql_mask = sales_df["Is SQL"].fillna(False)
        unmapped = sales_df[
            all_sql_mask
            & (sales_df["Campaign"].eq("Unmapped") | sales_df["RegionKey"].eq("UNKNOWN"))
        ].copy()
        if exclude_duplicates and "Duplicate Flag" in unmapped.columns:
            unmapped = unmapped[unmapped["Duplicate Flag"].fillna("").astype(str).str.strip().ne("已重复")]
        if not unmapped.empty:
            st.markdown("**未匹配到 Campaign / Ad set 或 Region 的 SQL 线索（用于检查命名/字段）：**")
            st.dataframe(
                unmapped[[
                    "Source File", "Lead ID", "Raw Adset", "Campaign", "Inferred Region",
                    "Country", "Lead Stage", "Created Time", "Converted Time", "Duplicate Flag",
                ]].sort_values(["Created Time", "Converted Time"], ascending=False),
                use_container_width=True,
                hide_index=True,
            )


# ---------------------------------------------------------------------------
# 4. 地区对比
# ---------------------------------------------------------------------------
st.subheader("4️⃣ 地区对比")

region_agg = (
    fdf.groupby("Region")
    .agg(Spend=("Spend", "sum"), Leads=("Leads", "sum"))
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
    fig_r_spend = px.pie(region_agg, names="Region", values="Spend", title="花费占比 by Region")
    st.plotly_chart(fig_r_spend, use_container_width=True)
with colF:
    fig_r_leads = px.pie(region_agg, names="Region", values="Leads", title="MQL Leads 占比 by Region")
    st.plotly_chart(fig_r_leads, use_container_width=True)
with colG:
    if not sales_fdf.empty and "SQL / MQL" in region_agg.columns:
        fig_r_sql = px.bar(region_agg, x="Region", y="SQL / MQL", title="SQL / MQL by Region", text_auto=".2%")
        fig_r_sql.update_yaxes(tickformat=".1%")
        st.plotly_chart(fig_r_sql, use_container_width=True)
    else:
        fig_r_cpl = px.bar(region_agg, x="Region", y="Blended CPL", title="综合 CPL by Region", text_auto=".2f")
        st.plotly_chart(fig_r_cpl, use_container_width=True)


# ---------------------------------------------------------------------------
# 5. 全地区 Daily total 趋势 (如果有该 sheet)
# ---------------------------------------------------------------------------
if not daily_total_df.empty:
    st.subheader("5️⃣ Daily Total 趋势 (各地区汇总)")
    dt_mask = (
        daily_total_df["Region"].isin(sel_regions)
        & (daily_total_df["Date"] >= pd.Timestamp(start_date))
        & (daily_total_df["Date"] <= pd.Timestamp(end_date))
    )
    dt = daily_total_df[dt_mask]
    fig_dt = px.line(
        dt.sort_values("Date"), x="Date", y="Spend", color="Region",
        title="每日总花费 (Daily Total)",
    )
    st.plotly_chart(fig_dt, use_container_width=True)


# ---------------------------------------------------------------------------
# 6. Weekly 汇总
# ---------------------------------------------------------------------------
if not weekly_df.empty:
    st.subheader("6️⃣ Weekly 汇总")
    w_mask = weekly_df["Region"].isin(sel_regions)
    wdf = weekly_df[w_mask]

    fig_weekly = go.Figure()
    for region, g in wdf.groupby("Region"):
        fig_weekly.add_trace(go.Bar(x=g["Period"], y=g["Spend"], name=f"{region} Spend"))
        fig_weekly.add_trace(go.Scatter(x=g["Period"], y=g["Avg CPL"], name=f"{region} Avg CPL", yaxis="y2"))

    fig_weekly.update_layout(
        title="周维度 Spend & Avg CPL",
        yaxis=dict(title="Spend"),
        yaxis2=dict(title="Avg CPL", overlaying="y", side="right"),
        xaxis=dict(title="Week"),
    )
    st.plotly_chart(fig_weekly, use_container_width=True)
    st.dataframe(wdf[["Region", "Period", "Spend", "Leads", "Avg CPL"]], use_container_width=True)


# ---------------------------------------------------------------------------
# 7. 动态/异常记录 (Key actions + Status)
# ---------------------------------------------------------------------------
st.subheader("7️⃣ 关键动态 & 状态预警")

notes = fdf[fdf["Key actions"].notna() & (fdf["Key actions"].astype(str).str.strip() != "")]
notes = notes.sort_values("Date", ascending=False)
st.markdown("**Key actions 记录:**")
st.dataframe(
    notes[["Date", "Region", "Campaign", "Status", "CPL", "Key actions"]],
    use_container_width=True,
    hide_index=True,
)

alert = fdf[fdf["Status"].isin(["Audience Fatigue"])]
if not alert.empty:
    st.markdown("**⚠️ 当前 Audience Fatigue 状态的记录(最近5条):**")
    st.dataframe(
        alert.sort_values("Date", ascending=False)
        [["Date", "Region", "Campaign", "Spend", "Leads", "CPL"]]
        .head(5),
        use_container_width=True,
        hide_index=True,
    )


# ---------------------------------------------------------------------------
# 8. 原始数据(可下载)
# ---------------------------------------------------------------------------
with st.expander("📋 查看 / 下载筛选后的 Meta 明细数据"):
    drop_cols = [c for c in ["Series"] if c in fdf.columns]
    st.dataframe(fdf.drop(columns=drop_cols), use_container_width=True)
    st.download_button(
        "下载 Meta CSV",
        fdf.drop(columns=drop_cols).to_csv(index=False).encode("utf-8-sig"),
        file_name="meta_ads_filtered.csv",
        mime="text/csv",
    )

if not sales_fdf.empty:
    with st.expander("📋 查看 / 下载筛选后的 SQL 销售线索明细"):
        sales_detail_cols = [
            "Source File", "Lead ID", "Raw Adset", "Campaign", "Inferred Region", "Country",
            "Lead Stage", "Created Time", "Converted Time", "SQL Analysis Date", "Duplicate Flag",
        ]
        st.dataframe(
            sales_fdf[sales_detail_cols].sort_values("SQL Analysis Date", ascending=False),
            use_container_width=True,
            hide_index=True,
        )
        st.download_button(
            "下载 SQL 明细 CSV",
            sales_fdf[sales_detail_cols].to_csv(index=False).encode("utf-8-sig"),
            file_name="sql_leads_filtered.csv",
            mime="text/csv",
        )