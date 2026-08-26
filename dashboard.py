from datetime import datetime, timedelta, timezone
import io
import os
import re
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

# =================【1. 页面基本配置与全局移动端响应式 UI】=================
st.set_page_config(
    page_title="Tim 卡台数据看板",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 强制定义北京时间 (UTC+8)
BEIJING_TZ = timezone(timedelta(hours=8))

st.markdown(
    """
    <style>
    /* 全局字体与版心微调 */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    .main .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2.5rem;
        max-width: 98%;
    }

    /* 顶部 KPI Metric 指标卡片立体化与自适应 */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 12px 14px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.82rem !important;
        font-weight: 600 !important;
        color: #64748b !important;
        white-space: normal !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.35rem !important;
        font-weight: 700 !important;
        color: #0f172a !important;
    }

    /* 标签页 Tabs 现代圆角胶囊样式 (支持手机端横向自然滑动) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background-color: #f1f5f9;
        padding: 5px;
        border-radius: 10px;
        overflow-x: auto;
        flex-wrap: nowrap;
        -webkit-overflow-scrolling: touch;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 6px 14px;
        font-size: 0.85rem;
        font-weight: 600;
        border: none !important;
        background-color: transparent;
        white-space: nowrap;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06) !important;
        color: #2563eb !important;
    }

    /* 折叠面板平滑化 */
    [data-testid="stExpander"] {
        border: 1px solid #e2e8f0 !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }

    /* 手机端专属响应式补丁 (屏幕宽度 < 768px) */
    @media (max-width: 768px) {
        .main .block-container {
            padding-left: 0.6rem !important;
            padding-right: 0.6rem !important;
            padding-top: 0.8rem !important;
            max-width: 100% !important;
        }
        [data-testid="stMetric"] {
            padding: 8px 10px !important;
            margin-bottom: 6px !important;
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.72rem !important;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.08rem !important;
            word-break: break-all;
        }
        div[data-testid="column"] {
            min-width: 100% !important;
            margin-bottom: 6px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =================【2. 仓库与数据源核心配置】=================
GITHUB_OWNER = "White55-web"
GITHUB_REPO = "overseas-card-bi"
FOLDER_NAME = "Tim_Data"
DICT_FILE_NAME = "Tim字典.xlsx"
FALLBACK_LOCAL_FILE = "Tim 清洗操作.xlsx"


# =================【3. 纯算法时间解析与卡号向量化】=================
def extract_file_datetime(filepath_or_name):
    if not filepath_or_name:
        return datetime.min

    filename = os.path.basename(str(filepath_or_name))

    match_sec = re.search(
        r"(\d{4})[-_](\d{2})[-_](\d{2})[\s_]+(\d{2})[-_:](\d{2})[-_:](\d{2})",
        filename,
    )
    if match_sec:
        try:
            y, m, d, h, mi, s = map(int, match_sec.groups())
            return datetime(y, m, d, h, mi, s)
        except Exception:
            pass

    match_day = re.search(r"(\d{4})[-_]?(\d{2})[-_]?(\d{2})", filename)
    if match_day:
        try:
            y, m, d = map(int, match_day.groups())
            return datetime(y, m, d, 0, 0, 0)
        except Exception:
            pass

    return datetime.min


def normalize_card_series(series):
    if series is None or series.empty:
        return pd.Series(dtype=str)
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    return s.str[-4:].str.zfill(4)


# =================【4. GitHub API 在线直读 + 本地双轨引擎 (带缓存)】=================
@st.cache_data(ttl=20, show_spinner=False)
def fetch_latest_dataset(owner, repo, folder_name, dict_name, fallback_file):
    headers = {
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "Streamlit-BI-App",
    }
    api_url = (
        f"https://api.github.com/repos/{owner}/{repo}/contents/{folder_name}"
    )

    try:
        resp = requests.get(api_url, headers=headers, timeout=6)
        if resp.status_code == 200:
            files_data = resp.json()
            xlsx_items = [
                f
                for f in files_data
                if isinstance(f, dict)
                and f.get("name", "").endswith(".xlsx")
                and not f.get("name", "").startswith("~$")
            ]
            if xlsx_items:
                latest_item = max(
                    xlsx_items,
                    key=lambda x: extract_file_datetime(x.get("name", "")),
                )
                file_name = latest_item["name"]
                dl_url = latest_item["download_url"]

                file_resp = requests.get(dl_url, headers=headers, timeout=12)
                if file_resp.status_code == 200:
                    stream = io.BytesIO(file_resp.content)
                    dt_obj = extract_file_datetime(file_name)
                    return stream, file_name, dt_obj, "GitHub API 在线直连"
    except Exception:
        pass

    if os.path.exists(folder_name):
        local_files = [
            os.path.join(folder_name, f)
            for f in os.listdir(folder_name)
            if f.endswith(".xlsx") and not f.startswith("~$")
        ]
        if local_files:
            latest_local = max(local_files, key=extract_file_datetime)
            file_name = os.path.basename(latest_local)
            dt_obj = extract_file_datetime(latest_local)
            return latest_local, file_name, dt_obj, "本地容器缓存"

    if fallback_file and os.path.exists(fallback_file):
        return (
            fallback_file,
            os.path.basename(fallback_file),
            datetime.min,
            "备用文件",
        )

    return None, "未找到数据", datetime.min, "无可用源"


# =================【5. 高性能数据清洗与标准化引擎 (内存级缓存)】=================
@st.cache_data(show_spinner=False, ttl=300)
def load_and_clean_raw_cached(
    raw_source, dict_file_path, file_identifier, timestamp_val
):
    df_raw = pd.DataFrame()
    try:
        excel_file = pd.ExcelFile(raw_source)
        candidate_dfs = []
        for s_name in excel_file.sheet_names:
            temp_df = excel_file.parse(s_name)
            if not temp_df.empty:
                temp_cols = [str(c).strip() for c in temp_df.columns]
                if any("卡号" in c for c in temp_cols) or any(
                    "金额" in c for c in temp_cols
                ):
                    df_raw = temp_df
                    break
                candidate_dfs.append(temp_df)

        if df_raw.empty and candidate_dfs:
            df_raw = max(candidate_dfs, key=len)
        elif df_raw.empty and excel_file.sheet_names:
            df_raw = excel_file.parse(excel_file.sheet_names[0])
    except Exception:
        return pd.DataFrame()

    if df_raw.empty:
        return df_raw

    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    if "卡号" in df_raw.columns:
        df_raw["卡号"] = df_raw["卡号"].astype(str).str.strip()
        df_raw["match_key"] = normalize_card_series(df_raw["卡号"])
    else:
        df_raw["match_key"] = ""

    if dict_file_path and os.path.exists(dict_file_path):
        try:
            df_dict = pd.read_excel(dict_file_path)
            df_dict.columns = [str(c).strip() for c in df_dict.columns]

            if "UA名字" in df_dict.columns and "卡号" in df_dict.columns:
                df_dict["UA名字"] = df_dict["UA名字"].astype(str).str.strip()
                df_dict["match_key"] = normalize_card_series(df_dict["卡号"])
                df_dict = df_dict.drop_duplicates(
                    subset=["match_key"], keep="last"
                )

                if "UA名字" in df_raw.columns:
                    df_raw = df_raw.drop(columns=["UA名字"])

                if "match_key" in df_raw.columns:
                    df_raw = pd.merge(
                        df_raw,
                        df_dict[["match_key", "UA名字"]],
                        on="match_key",
                        how="left",
                    )
                    df_raw["UA名字"] = df_raw["UA名字"].fillna("未分配")
        except Exception:
            pass

    if "UA名字" not in df_raw.columns:
        df_raw["UA名字"] = "未分配"

    if "match_key" in df_raw.columns:
        df_raw = df_raw.drop(columns=["match_key"])

    if "交易金额" in df_raw.columns:
        df_raw["交易金额"] = pd.to_numeric(
            df_raw["交易金额"], errors="coerce"
        ).fillna(0.0)

    if "交易时间" in df_raw.columns:
        df_raw["交易时间"] = pd.to_datetime(df_raw["交易时间"], errors="coerce")
        df_raw["交易日期"] = df_raw["交易时间"].dt.date
        df_raw = df_raw.sort_values(by="交易时间", ascending=False)
    elif "交易日期" in df_raw.columns:
        df_raw["交易日期"] = pd.to_datetime(
            df_raw["交易日期"], errors="coerce"
        ).dt.date

    for col in ["卡号", "UA名字", "交易状态"]:
        if col in df_raw.columns:
            df_raw[col] = (
                df_raw[col].astype(str).str.strip().replace("nan", "")
            )

    return df_raw


# =================【6. 数据载入与状态初始化】=================
raw_source, current_filename, latest_dt_obj, source_channel = (
    fetch_latest_dataset(
        GITHUB_OWNER,
        GITHUB_REPO,
        FOLDER_NAME,
        DICT_FILE_NAME,
        FALLBACK_LOCAL_FILE,
    )
)

if raw_source is None:
    st.error(
        f"❌ 未能从 GitHub 或本地找到有效流水数据。请检查仓库路径 `{FOLDER_NAME}/`。"
    )
    st.stop()

raw_timestamp = (
    latest_dt_obj.timestamp() if latest_dt_obj != datetime.min else 0
)

if st.session_state.get("_last_raw_timestamp") != raw_timestamp:
    st.session_state["_last_raw_timestamp"] = raw_timestamp
    for k in [
        "main_date_range_tim",
        "pivot_date_range_tim",
        "raw_date_range_tim",
    ]:
        if k in st.session_state:
            del st.session_state[k]

df_raw = load_and_clean_raw_cached(
    raw_source, DICT_FILE_NAME, current_filename, raw_timestamp
)

if df_raw.empty:
    st.warning(f"⚠️ `{current_filename}` 中暂无有效流水数据。")
    st.stop()


# =================【7. 侧边栏控制面板】=================
with st.sidebar:
    st.header("⚙️ 中台系统控制")
    if st.button("🔄 立即刷新 / 获取最新数据", use_container_width=True):
        st.cache_data.clear()
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.success("✅ 缓存已清空，正在重连 GitHub API...")
        st.rerun()

    auto_refresh = st.checkbox("⏱️ 开启 60 秒自动静默轮询", value=False)
    if auto_refresh:
        st.markdown(
            """
            <script>
            setTimeout(function(){
                window.location.reload();
            }, 60000);
            </script>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.caption("🛠️ Tim 买量消耗大盘与流水对账系统 ｜ White制作")


# =================【8. 主页面渲染】=================
mtime_str = (
    latest_dt_obj.strftime("%Y-%m-%d %H:%M:%S")
    if latest_dt_obj != datetime.min
    else "历史备份"
)
dict_status = (
    f"✅ 已关联 `{DICT_FILE_NAME}`"
    if os.path.exists(DICT_FILE_NAME)
    else f"⚠️ 缺少字典 (显示为未分配)"
)

st.caption(
    f"📁 载入流水: `{current_filename}` ｜ 🕒 导出: **{mtime_str}** ｜ 📡 `{source_channel}` ｜ 📖 {dict_status}"
)

# ----------------------------------------------------
# 板块 0：【Tim】主筛选条件（展开/折叠框）
# ----------------------------------------------------
with st.expander("🔍 展开 / 折叠【Tim】主筛选条件", expanded=True):
    f_col1, f_col2, f_col3 = st.columns([1.2, 1.2, 1.6])

    with f_col1:
        if "交易日期" in df_raw.columns and not df_raw["交易日期"].dropna().empty:
            m_min_date = df_raw["交易日期"].min()
            m_max_date = df_raw["交易日期"].max()
            main_date_range = st.date_input(
                "📅 交易日期范围",
                value=(m_min_date, m_max_date),
                min_value=m_min_date,
                max_value=m_max_date,
                key="main_date_range_tim",
            )
        else:
            main_date_range = None

    with f_col2:
        if "UA名字" in df_raw.columns:
            main_ua_options = sorted(
                [
                    str(x)
                    for x in df_raw["UA名字"].dropna().unique()
                    if str(x).strip() and str(x) != "nan"
                ]
            )
            selected_main_ua = st.multiselect(
                "👤 UA 名字",
                options=main_ua_options,
                default=[],
                placeholder="留空默认展示所有人",
                key="main_ua_tim",
            )
        else:
            selected_main_ua = []

    with f_col3:
        if "交易状态" in df_raw.columns:
            main_status_options = sorted(
                [
                    x
                    for x in df_raw["交易状态"].dropna().unique()
                    if str(x).strip()
                ]
            )
            selected_main_status = st.multiselect(
                "📌 交易状态",
                options=main_status_options,
                default=main_status_options,
                key="main_status_tim",
            )
        else:
            selected_main_status = []

# 应用主筛选
df_main_filtered = df_raw.copy()
if (
    main_date_range
    and isinstance(main_date_range, (tuple, list))
    and len(main_date_range) == 2
):
    df_main_filtered = df_main_filtered[
        (df_main_filtered["交易日期"] >= main_date_range[0])
        & (df_main_filtered["交易日期"] <= main_date_range[1])
    ]
elif (
    main_date_range
    and isinstance(main_date_range, (tuple, list))
    and len(main_date_range) == 1
):
    df_main_filtered = df_main_filtered[
        df_main_filtered["交易日期"] == main_date_range[0]
    ]

if selected_main_ua and "UA名字" in df_main_filtered.columns:
    df_main_filtered = df_main_filtered[
        df_main_filtered["UA名字"].isin(selected_main_ua)
    ]

if selected_main_status and "交易状态" in df_main_filtered.columns:
    df_main_filtered = df_main_filtered[
        df_main_filtered["交易状态"].isin(selected_main_status)
    ]

# ----------------------------------------------------
# 顶部核心 KPI 指标卡
# ----------------------------------------------------
latest_global_date = (
    df_raw["交易日期"].max() if "交易日期" in df_raw.columns else None
)
today_pending_sum = (
    df_raw[
        (df_raw["交易日期"] == latest_global_date)
        & (df_raw["交易状态"] == "PENDING")
    ]["交易金额"].sum()
    if (latest_global_date and "交易状态" in df_raw.columns)
    else 0.0
)

pending_spend = (
    df_main_filtered[df_main_filtered["交易状态"] == "PENDING"][
        "交易金额"
    ].sum()
    if "交易状态" in df_main_filtered.columns
    else 0.0
)
complete_spend = (
    df_main_filtered[df_main_filtered["交易状态"] == "COMPLETE"][
        "交易金额"
    ].sum()
    if "交易状态" in df_main_filtered.columns
    else 0.0
)
active_card_cnt = (
    df_main_filtered[df_main_filtered["交易状态"] == "PENDING"]["卡号"].nunique()
    if "交易状态" in df_main_filtered.columns
    else df_main_filtered["卡号"].nunique()
)

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.metric(
        label="💰 筛选总消耗 (PENDING)",
        value=f"${pending_spend:,.2f}",
        help="当前筛选日期与 UA 范围内的 PENDING 消耗总计",
    )
with k2:
    st.metric(
        label=f"🔥 今日全盘消耗 ({latest_global_date})",
        value=f"${today_pending_sum:,.2f}",
        help=f"最新数据日 ({latest_global_date}) 截至目前的 PENDING 扣款总额",
    )
with k3:
    st.metric(
        label="💳 PENDING 活跃卡数",
        value=f"{active_card_cnt} 张",
        help="当前筛选范围内产生扣费的独立活跃卡数",
    )
with k4:
    st.metric(
        label="✅ COMPLETE 历史结算",
        value=f"${complete_spend:,.2f}",
        help="当前筛选范围内已入账结算的 COMPLETE 金额",
    )

# ----------------------------------------------------
# 板块 1：中部趋势图表 (全量锁定坐标轴，彻底杜绝手机误触缩放)
# ----------------------------------------------------
chart1, chart2 = st.columns(2)
with chart1:
    if (
        "交易日期" in df_main_filtered.columns
        and "交易金额" in df_main_filtered.columns
        and not df_main_filtered.empty
    ):
        daily_data = (
            df_main_filtered.groupby(["交易日期", "交易状态"])["交易金额"]
            .sum()
            .reset_index()
        )
        fig_trend = px.bar(
            daily_data,
            x="交易日期",
            y="交易金额",
            color="交易状态",
            title="📅 Tim - 每日消耗与状态趋势",
            barmode="stack",
            template="plotly_white",
            color_discrete_map={
                "PENDING": "#2563eb",
                "COMPLETE": "#10b981",
                "REVERSED": "#f59e0b",
                "DECLINED": "#ef4444",
                "FAILED": "#ef4444",
            },
        )
        fig_trend.update_layout(
            font_family="-apple-system, BlinkMacSystemFont, Segoe UI",
            title=dict(text="📅 Tim - 每日消耗与状态趋势", x=0, xanchor="left"),
            title_font_size=14,
            legend_title_text="",
            dragmode=False,  # 彻底禁止拖拽框选
            xaxis=dict(fixedrange=True),  # 锁定 X 轴禁止缩放
            yaxis=dict(fixedrange=True),  # 锁定 Y 轴禁止缩放
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.18,
                xanchor="center",
                x=0.5,
            ),
            margin=dict(l=10, r=10, t=35, b=45),
            hovermode="x unified",
        )
        st.plotly_chart(
            fig_trend,
            use_container_width=True,
            config={
                "displayModeBar": False,
                "scrollZoom": False,
                "doubleClick": False,
                "showAxisDragHandles": False,
            },
        )

with chart2:
    if (
        "UA名字" in df_main_filtered.columns
        and "交易金额" in df_main_filtered.columns
        and not df_main_filtered.empty
    ):
        df_pending_chart = (
            df_main_filtered[df_main_filtered["交易状态"] == "PENDING"]
            if "交易状态" in df_main_filtered.columns
            else df_main_filtered
        )
        ua_data = (
            df_pending_chart.groupby("UA名字")["交易金额"]
            .sum()
            .reset_index()
            .sort_values(by="交易金额", ascending=False)
        )
        fig_ua = px.pie(
            ua_data,
            names="UA名字",
            values="交易金额",
            title="🎯 Tim - 投放团队 PENDING 消耗占比",
            hole=0.45,
            template="plotly_white",
            color_discrete_sequence=px.colors.qualitative.Pastel,
        )
        fig_ua.update_traces(textposition="inside", textinfo="percent+label")
        fig_ua.update_layout(
            font_family="-apple-system, BlinkMacSystemFont, Segoe UI",
            title=dict(text="🎯 Tim - 投放团队 PENDING 消耗占比", x=0, xanchor="left"),
            title_font_size=14,
            dragmode=False,
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.15,
                xanchor="center",
                x=0.5,
            ),
            margin=dict(l=10, r=10, t=35, b=40),
        )
        st.plotly_chart(
            fig_ua,
            use_container_width=True,
            config={
                "displayModeBar": False,
                "scrollZoom": False,
                "doubleClick": False,
            },
        )

st.markdown("---")

# ----------------------------------------------------
# 板块 2：单卡每日消耗透视 (仅统计 PENDING)
# ----------------------------------------------------
st.subheader("💳 Tim 单卡每日消耗透视 (仅统计 PENDING)")

df_pending = (
    df_raw[df_raw["交易状态"] == "PENDING"]
    if "交易状态" in df_raw.columns
    else df_raw
)

if (
    "卡号" in df_pending.columns
    and "交易日期" in df_pending.columns
    and "交易金额" in df_pending.columns
    and not df_pending.empty
):
    p_col1, p_col2, p_col3 = st.columns(3)

    with p_col1:
        card_options = sorted(list(df_pending["卡号"].dropna().unique()))
        selected_cards = st.multiselect(
            "💳 筛选卡号",
            options=card_options,
            default=[],
            placeholder="支持搜索/勾选卡号（留空全选）",
            key="pivot_card_tim",
        )

    with p_col2:
        if "UA名字" in df_pending.columns:
            pivot_ua_options = sorted(
                [
                    str(x)
                    for x in df_pending["UA名字"].dropna().unique()
                    if str(x).strip() and str(x) != "nan"
                ]
            )
            selected_pivot_ua = st.multiselect(
                "👤 筛选 UA",
                options=pivot_ua_options,
                default=[],
                placeholder="按 UA 筛选（留空全选）",
                key="pivot_ua_tim",
            )
        else:
            selected_pivot_ua = []

    with p_col3:
        p_min_date = df_pending["交易日期"].min()
        p_max_date = df_pending["交易日期"].max()
        selected_pivot_date_range = st.date_input(
            "📅 交易日期范围",
            value=(p_min_date, p_max_date)
            if (p_min_date and p_max_date)
            else datetime.now().date(),
            min_value=p_min_date,
            max_value=p_max_date,
            key="pivot_date_range_tim",
        )

    df_pivot_filtered = df_pending.copy()
    if selected_cards:
        df_pivot_filtered = df_pivot_filtered[
            df_pivot_filtered["卡号"].isin(selected_cards)
        ]
    if selected_pivot_ua and "UA名字" in df_pivot_filtered.columns:
        df_pivot_filtered = df_pivot_filtered[
            df_pivot_filtered["UA名字"].isin(selected_pivot_ua)
        ]
    if (
        isinstance(selected_pivot_date_range, (tuple, list))
        and len(selected_pivot_date_range) == 2
    ):
        df_pivot_filtered = df_pivot_filtered[
            (
                df_pivot_filtered["交易日期"]
                >= selected_pivot_date_range[0]
            )
            & (
                df_pivot_filtered["交易日期"]
                <= selected_pivot_date_range[1]
            )
        ]
    elif (
        isinstance(selected_pivot_date_range, (tuple, list))
        and len(selected_pivot_date_range) == 1
    ):
        df_pivot_filtered = df_pivot_filtered[
            df_pivot_filtered["交易日期"] == selected_pivot_date_range[0]
        ]

    if not df_pivot_filtered.empty:
        sum_spend = df_pivot_filtered["交易金额"].sum()
        sum_cards = df_pivot_filtered["卡号"].nunique()
        sum_tx = len(df_pivot_filtered)

        s_c1, s_c2, s_c3 = st.columns(3)
        with s_c1:
            st.metric("💰 当前筛选总消耗 (SUM)", f"${sum_spend:,.2f}")
        with s_c2:
            st.metric("💳 涉及活跃卡号数", f"{sum_cards} 张")
        with s_c3:
            st.metric("📝 累计扣费总笔数", f"{sum_tx} 笔")

        view_tab1, view_tab2, view_tab3 = st.tabs(
            [
                "📋 单卡每日消耗明细表",
                "👥 UA 每日汇总表 (SUM)",
                "📊 卡号 × 日期 透视大表 (Pivot)",
            ]
        )

        with view_tab1:
            group_cols = ["卡号", "交易日期"]
            if "UA名字" in df_pivot_filtered.columns:
                group_cols.insert(1, "UA名字")

            card_daily_df = (
                df_pivot_filtered.groupby(group_cols)
                .agg(
                    当日总消耗=("交易金额", "sum"),
                    交易笔数=("交易金额", "count"),
                )
                .reset_index()
            )
            card_daily_df = card_daily_df.sort_values(
                by=["交易日期", "当日总消耗"], ascending=[False, False]
            )

            top_sum_dict = {
                "卡号": "🔥 【当前筛选合计 SUM】",
                "当日总消耗": sum_spend,
                "交易笔数": sum_tx,
            }
            if "UA名字" in df_pivot_filtered.columns:
                top_sum_dict["UA名字"] = (
                    selected_pivot_ua[0]
                    if len(selected_pivot_ua) == 1
                    else f"共 {df_pivot_filtered['UA名字'].nunique()} 人"
                )
            if "交易日期" in df_pivot_filtered.columns:
                if (
                    isinstance(selected_pivot_date_range, (tuple, list))
                    and len(selected_pivot_date_range) == 2
                ):
                    if (
                        selected_pivot_date_range[0]
                        == selected_pivot_date_range[1]
                    ):
                        top_sum_dict["交易日期"] = str(
                            selected_pivot_date_range[0]
                        )
                    else:
                        top_sum_dict["交易日期"] = (
                            f"{selected_pivot_date_range[0]} ~ {selected_pivot_date_range[1]}"
                        )
                else:
                    top_sum_dict["交易日期"] = (
                        f"共 {df_pivot_filtered['交易日期'].nunique()} 天"
                    )

            card_daily_display = pd.concat(
                [pd.DataFrame([top_sum_dict]), card_daily_df],
                ignore_index=True,
            )

            st.dataframe(
                card_daily_display,
                column_config={
                    "当日总消耗": st.column_config.NumberColumn(
                        "当日总消耗 (PENDING)", format="$%.2f"
                    ),
                    "交易笔数": st.column_config.NumberColumn(
                        "交易笔数", format="%d 笔"
                    ),
                },
                hide_index=True,
                use_container_width=True,
            )

        with view_tab2:
            if "UA名字" in df_pivot_filtered.columns:
                ua_daily_df = (
                    df_pivot_filtered.groupby(["UA名字", "交易日期"])
                    .agg(
                        当日UA总消耗=("交易金额", "sum"),
                        消耗卡数=("卡号", "nunique"),
                        扣费笔数=("交易金额", "count"),
                    )
                    .reset_index()
                    .sort_values(
                        by=["交易日期", "当日UA总消耗"],
                        ascending=[False, False],
                    )
                )
                st.dataframe(
                    ua_daily_df,
                    column_config={
                        "当日UA总消耗": st.column_config.NumberColumn(
                            "当日 UA 总消耗 (SUM)", format="$%.2f"
                        ),
                        "消耗卡数": st.column_config.NumberColumn(
                            "消耗卡数", format="%d 张"
                        ),
                        "扣费笔数": st.column_config.NumberColumn(
                            "扣费总笔数", format="%d 笔"
                        ),
                    },
                    hide_index=True,
                    use_container_width=True,
                )
            else:
                st.info("💡 当前数据源无 UA 归属信息。")

        with view_tab3:
            pivot_index = (
                ["卡号", "UA名字"]
                if "UA名字" in df_pivot_filtered.columns
                else ["卡号"]
            )
            pivot_df = df_pivot_filtered.pivot_table(
                index=pivot_index,
                columns="交易日期",
                values="交易金额",
                aggfunc="sum",
                fill_value=0,
            )
            pivot_df["期间总消耗"] = pivot_df.sum(axis=1)
            pivot_df = pivot_df.sort_values(
                by="期间总消耗", ascending=False
            )

            pivot_display = pivot_df.map(
                lambda x: f"${x:,.2f}"
                if isinstance(x, (int, float))
                else x
            )
            st.dataframe(pivot_display, use_container_width=True)
    else:
        st.info("💡 筛选条件下未找到匹配的 PENDING 流水。")
else:
    st.info("当前所选筛选条件下暂无 PENDING 交易数据。")

st.markdown("---")

# ----------------------------------------------------
# 板块 3：全量明细流水对账
# ----------------------------------------------------
st.subheader("📋 Tim 全量流水对账")

if not df_raw.empty:
    r_col1, r_col2, r_col3 = st.columns(3)

    with r_col1:
        raw_card_options = (
            sorted(list(df_raw["卡号"].dropna().unique()))
            if "卡号" in df_raw.columns
            else []
        )
        selected_raw_cards = st.multiselect(
            "💳 筛选卡号",
            options=raw_card_options,
            default=[],
            placeholder="支持搜索/勾选卡号（留空全选）",
            key="raw_card_tim",
        )

    with r_col2:
        raw_ua_options = (
            sorted(
                [
                    str(x)
                    for x in df_raw["UA名字"].dropna().unique()
                    if str(x).strip() and str(x) != "nan"
                ]
            )
            if "UA名字" in df_raw.columns
            else []
        )
        selected_raw_ua = st.multiselect(
            "👤 筛选 UA",
            options=raw_ua_options,
            default=[],
            placeholder="按 UA 筛选（留空全选）",
            key="raw_ua_tim",
        )

    with r_col3:
        if (
            "交易日期" in df_raw.columns
            and not df_raw["交易日期"].dropna().empty
        ):
            r_min_date = df_raw["交易日期"].min()
            r_max_date = df_raw["交易日期"].max()
            selected_raw_date_range = st.date_input(
                "📅 交易日期范围",
                value=(r_min_date, r_max_date),
                min_value=r_min_date,
                max_value=r_max_date,
                key="raw_date_range_tim",
            )
        else:
            selected_raw_date_range = None

    df_raw_filtered = df_raw.copy()
    if selected_raw_cards and "卡号" in df_raw_filtered.columns:
        df_raw_filtered = df_raw_filtered[
            df_raw_filtered["卡号"].isin(selected_raw_cards)
        ]
    if selected_raw_ua and "UA名字" in df_raw_filtered.columns:
        df_raw_filtered = df_raw_filtered[
            df_raw_filtered["UA名字"].isin(selected_raw_ua)
        ]
    if (
        selected_raw_date_range
        and isinstance(selected_raw_date_range, (tuple, list))
        and len(selected_raw_date_range) == 2
    ):
        df_raw_filtered = df_raw_filtered[
            (df_raw_filtered["交易日期"] >= selected_raw_date_range[0])
            & (df_raw_filtered["交易日期"] <= selected_raw_date_range[1])
        ]
    elif (
        selected_raw_date_range
        and isinstance(selected_raw_date_range, (tuple, list))
        and len(selected_raw_date_range) == 1
    ):
        df_raw_filtered = df_raw_filtered[
            df_raw_filtered["交易日期"] == selected_raw_date_range[0]
        ]

    col_custom_cfg = {}
    if "交易金额" in df_raw_filtered.columns:
        col_custom_cfg["交易金额"] = st.column_config.NumberColumn(
            "交易金额", format="$%.2f"
        )

    st.dataframe(
        df_raw_filtered,
        column_config=col_custom_cfg,
        use_container_width=True,
    )
else:
    st.info("当前筛选条件下暂无流水数据。")
