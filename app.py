"""
Meta Ads 多地区看板 (EMEA / India / North America ...)
运行方式: streamlit run app.py
依赖: pip install streamlit pandas plotly openpyxl numpy --break-system-packages
"""

import re
import io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Meta Ads 多地区看板", layout="wide")

# ---------------------------------------------------------------------------
# 配置: 统一字段(各地区的campaign sheet名称不强制一致，自动识别)
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


# ---------------------------------------------------------------------------
# 数据清洗 / 加载函数
# ---------------------------------------------------------------------------
def clean_numeric(series: pd.Series) -> pd.Series:
    """把 #REF! 等异常值统一转成 NaN, 并转成数值型"""
    return pd.to_numeric(series.replace("#REF!", np.nan), errors="coerce")


def load_campaign_sheet(xls: pd.ExcelFile, sheet_name: str, region: str, campaign: str) -> pd.DataFrame:
    """读取单个 campaign / Daily total sheet, 统一字段、清洗数据"""
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

    # 注意: 源表里 CA / AI Instant / AI Layout / Daily total 的 "CTR" 列实际存的是
    # CPC 的数值(公式引用错误), 这里统一按 Link Clicks / Impressions 重新计算 CTR,
    # 以保证跨 campaign / 跨地区对比的口径一致。
    clicks_imp_mask = df["Impressions"].notna() & (df["Impressions"] != 0) & df["Link Clicks"].notna()
    df.loc[clicks_imp_mask, "CTR"] = df.loc[clicks_imp_mask, "Link Clicks"] / df.loc[clicks_imp_mask, "Impressions"]

    # 去掉日期为空 / Spend 为空的尾部空行(比如表格末尾还没填的6/12, 6/13)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df = df[df["Spend"].notna()]

    df["Campaign"] = campaign
    df["Region"] = region
    return df.reset_index(drop=True)


def load_weekly_sheet(xls: pd.ExcelFile, sheet_name: str, region: str) -> pd.DataFrame:
    """解析多月份分块的 Weekly sheet, 自动跳过 #REF! / 空白的早期周"""
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
            continue  # 表头行
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
        df["Period"] = None
        return df

    for col in ["Spend", "Leads", "Avg CPL"]:
        df[col] = clean_numeric(df[col])

    # 自动剔除 #REF!/空白周(对应 user 说的 "之前的 #REF! 不用管")
    df = df.dropna(subset=["Spend", "Leads", "Avg CPL"], how="all").reset_index(drop=True)

    # 给一个排序用的序号, 方便按时间顺序展示(月份内 week 1..n)
    month_order = {m: i for i, m in enumerate(MONTH_NAMES)}
    df["MonthOrder"] = df["Month"].map(month_order)
    df["WeekNum"] = df["Week"].str.extract(r"(\d+)").astype(int)
    df = df.sort_values(["MonthOrder", "WeekNum"]).reset_index(drop=True)

    df["Period"] = df["Month"].str[:3] + " " + df["Week"]
    df["Region"] = region
    return df


def guess_region_name(filename: str) -> str:
    """从文件名猜地区名, 例如 'EMEA地区_Meta_ads_看板.xlsx' -> 'EMEA'"""
    m = re.match(r"^([A-Za-z]+)", filename)
    if m:
        return m.group(1).upper()
    return filename


@st.cache_data(show_spinner=False)
def load_region_file(file_bytes: bytes, region_name: str):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    campaign_dfs = []
    daily_total = None
    weekly = None

    for sheet in xls.sheet_names:
        norm = sheet.strip().lower()

        if norm == "weekly":
            weekly = load_weekly_sheet(xls, sheet, region_name)
            continue

        # 先看看这个sheet是不是"每日数据"格式(有Date和Spend列)
        header_df = pd.read_excel(xls, sheet_name=sheet, nrows=0)
        cols = set(header_df.columns)
        if not ({"Date", "Spend"} <= cols):
            continue  # 跳过不认识的sheet(比如SQL线索表)

        if norm in ("daily total", "daily_total"):
            daily_total = load_campaign_sheet(xls, sheet, region_name, "Total")
        else:
            campaign_dfs.append(load_campaign_sheet(xls, sheet, region_name, sheet))

    campaigns = pd.concat(campaign_dfs, ignore_index=True) if campaign_dfs else pd.DataFrame()
    return campaigns, daily_total, weekly


# ---------------------------------------------------------------------------
# Sidebar: 上传文件(支持多个地区)
# ---------------------------------------------------------------------------
st.sidebar.title("📂 数据源")
uploaded_files = st.sidebar.file_uploader(
    "上传各地区的 Meta Ads 看板表 (.xlsx)，可多选",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("⬅️ 请在左侧上传一个或多个地区的 Excel 看板文件(EMEA / India / North America ...)")
    st.stop()

all_campaigns, all_daily_total, all_weekly = [], [], []
for f in uploaded_files:
    default_region = guess_region_name(f.name)
    region_name = st.sidebar.text_input(f"「{f.name}」对应的地区名称", value=default_region, key=f.name)

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
    st.error("没有解析到任何 campaign 数据, 请检查表格里是否有包含 Date 和 Spend 列的 sheet")
    st.stop()


# ---------------------------------------------------------------------------
# Sidebar: 筛选器
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.title("🔍 筛选")

regions = sorted(campaigns_df["Region"].unique())
sel_regions = st.sidebar.multiselect("地区 Region", regions, default=regions)

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

mask = (
    campaigns_df["Region"].isin(sel_regions)
    & campaigns_df["Campaign"].isin(sel_campaigns)
    & (campaigns_df["Date"] >= pd.Timestamp(start_date))
    & (campaigns_df["Date"] <= pd.Timestamp(end_date))
)
fdf = campaigns_df[mask].copy()

if fdf.empty:
    st.warning("当前筛选条件下没有数据, 请调整左侧筛选项")
    st.stop()


# ---------------------------------------------------------------------------
# 顶部 KPI
# ---------------------------------------------------------------------------
st.title("📊 Meta Ads 多地区效果看板")

total_spend = fdf["Spend"].sum()
total_leads = fdf["Leads"].sum()
blended_cpl = total_spend / total_leads if total_leads else np.nan
avg_ctr = fdf["CTR"].mean()
avg_cvr = fdf["Lead CVR"].mean()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("总花费 Spend", f"${total_spend:,.0f}")
c2.metric("总线索 Leads", f"{total_leads:,.0f}")
c3.metric("综合 CPL", f"${blended_cpl:,.2f}")
c4.metric("平均 CTR", f"{avg_ctr:.2%}")
c5.metric("平均 Lead CVR", f"{avg_cvr:.2%}")

st.markdown("---")


# ---------------------------------------------------------------------------
# 1. 时间趋势: Spend / Leads / CPL
# ---------------------------------------------------------------------------
st.subheader("1️⃣ 每日趋势 (按地区+campaign)")

trend_metric = st.selectbox(
    "选择指标",
    ["Spend", "Leads", "CPL", "CTR", "Lead CVR", "Frequency", "3-Day Avg CPL"],
    key="trend_metric",
)

fdf["Series"] = fdf["Region"] + " - " + fdf["Campaign"]
fig_trend = px.line(
    fdf.sort_values("Date"),
    x="Date", y=trend_metric, color="Series",
    markers=False,
    title=f"{trend_metric} 每日趋势",
)
st.plotly_chart(fig_trend, use_container_width=True)


# ---------------------------------------------------------------------------
# 2. Campaign / Ad set 横向对比
# ---------------------------------------------------------------------------
st.subheader("2️⃣ Campaign / Ad set 横向对比")

agg = (
    fdf.groupby(["Region", "Campaign"])
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
        title="各 Campaign 综合 CPL 对比 (Spend ÷ Leads)",
        text_auto=".2f",
    )
    st.plotly_chart(fig_cpl, use_container_width=True)

with colB:
    fig_spend = px.bar(
        agg, x="Campaign", y="Spend", color="Region", barmode="group",
        title="各 Campaign 花费占比",
        text_auto=".0f",
    )
    st.plotly_chart(fig_spend, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    fig_ctr = px.bar(
        agg, x="Campaign", y="CTR", color="Region", barmode="group",
        title="各 Campaign 平均 CTR", text_auto=".2%",
    )
    fig_ctr.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig_ctr, use_container_width=True)

with colD:
    fig_cvr = px.bar(
        agg, x="Campaign", y="Lead_CVR", color="Region", barmode="group",
        title="各 Campaign 平均 Lead CVR", text_auto=".2%",
    )
    fig_cvr.update_yaxes(tickformat=".1%")
    st.plotly_chart(fig_cvr, use_container_width=True)


# ---------------------------------------------------------------------------
# 3. 地区对比
# ---------------------------------------------------------------------------
st.subheader("3️⃣ 地区对比")

region_agg = (
    fdf.groupby("Region")
    .agg(Spend=("Spend", "sum"), Leads=("Leads", "sum"))
    .reset_index()
)
region_agg["Blended CPL"] = region_agg["Spend"] / region_agg["Leads"].replace(0, np.nan)

colE, colF, colG = st.columns(3)
with colE:
    fig_r_spend = px.pie(region_agg, names="Region", values="Spend", title="花费占比 by Region")
    st.plotly_chart(fig_r_spend, use_container_width=True)
with colF:
    fig_r_leads = px.pie(region_agg, names="Region", values="Leads", title="线索占比 by Region")
    st.plotly_chart(fig_r_leads, use_container_width=True)
with colG:
    fig_r_cpl = px.bar(region_agg, x="Region", y="Blended CPL", title="综合 CPL by Region", text_auto=".2f")
    st.plotly_chart(fig_r_cpl, use_container_width=True)


# ---------------------------------------------------------------------------
# 4. 全地区 Daily total 趋势 (如果有该 sheet)
# ---------------------------------------------------------------------------
if not daily_total_df.empty:
    st.subheader("4️⃣ Daily Total 趋势 (各地区汇总)")
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
# 5. Weekly 汇总 (从 April Week2 开始)
# ---------------------------------------------------------------------------
if not weekly_df.empty:
    st.subheader("5️⃣ Weekly 汇总")
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
# 6. 动态/异常记录 (Key actions + Status)
# ---------------------------------------------------------------------------
st.subheader("6️⃣ 关键动态 & 状态预警")

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
# 7. 原始数据(可下载)
# ---------------------------------------------------------------------------
with st.expander("📋 查看 / 下载筛选后的明细数据"):
    st.dataframe(fdf.drop(columns=["Series"]), use_container_width=True)
    st.download_button(
        "下载 CSV",
        fdf.drop(columns=["Series"]).to_csv(index=False).encode("utf-8-sig"),
        file_name="meta_ads_filtered.csv",
        mime="text/csv",
    )