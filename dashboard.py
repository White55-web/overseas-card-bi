import glob
import os
import re
import subprocess
from datetime import datetime, timedelta, timezone
import pandas as pd
import plotly.express as px
import streamlit as st

# =================【1. 页面基本配置与时区设定】=================
st.set_page_config(
    page_title="出海多卡台数据看板 (公开版)",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 强制定义北京时间 (UTC+8)
BEIJING_TZ = timezone(timedelta(hours=8))


def get_display_file_time(file_path, raw_mtime):
    """智能解析时间：
    1. 优先从卡台导出的文件名中正则提取真实导出时间
    2. 若提取失败，则取物理修改时间并强制转换为北京时间 (UTC+8)
    """
    if file_path:
        filename = os.path.basename(file_path)
        match = re.search(
            r"(\d{4}[-_]\d{2}[-_]\d{2})[\s_]+(\d{2})[-_:](\d{2})[-_:](\d{2})",
            filename,
        )
        if match:
            date_part = match.group(1).replace("_", "-")
            return f"{date_part} {match.group(2)}:{match.group(3)}:{match.group(4)}"

    if raw_mtime > 0:
        return datetime.fromtimestamp(raw_mtime, tz=BEIJING_TZ).strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    return "未知"


# =================【全局高颜值 UI/CSS 注入】=================
st.markdown(
    """
    <style>
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    .main .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
        max-width: 96%;
    }
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 14px 18px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    [data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08), 0 4px 6px -2px rgba(0, 0, 0, 0.04);
        border-color: #cbd5e1;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.84rem !important;
        font-weight: 600 !important;
        color: #64748b !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.55rem !important;
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f5f9;
        padding: 6px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 18px;
        font-weight: 600;
        border: none !important;
        background-color: transparent;
        transition: all 0.2s ease;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06) !important;
        color: #2563eb !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# =================【卡台路由与专属字典配置】=================
PLATFORMS = {
    "Forest": {
        "folder": "Forest_Data",
        "fallback_file": "Forest 清洗操作.xlsx",
        "dict_file": "Forest字典.xlsx",
    },
    "Tim": {
        "folder": "Tim_Data",
        "fallback_file": "Tim 清洗操作.xlsx",
        "dict_file": "Tim字典.xlsx",
    },
}

# =================【侧边栏控制面板 (含云端主动 Git Pull 引擎)】=================
with st.sidebar:
    st.header("⚙️ 中台系统控制")
    if st.button("🔄 立即同步所有卡台数据", use_container_width=True):
        # 1. 强制云端服务器在后台执行 git pull 拉取 GitHub 最新推送的 Excel
        try:
            pull_res = subprocess.run(
                ["git", "pull"], capture_output=True, text=True, timeout=15
            )
            git_msg = (
                pull_res.stdout.strip()
                if pull_res.stdout
                else pull_res.stderr.strip()
            )
        except Exception as e:
            git_msg = "Git 自动拉取已跳过"

        # 2. 清除所有内存缓存与组件状态
        st.cache_data.clear()
        for key in list(st.session_state.keys()):
            del st.session_state[key]

        st.success(f"✅ 同步完成！({git_msg})")
        st.rerun()

    st.markdown("---")
    st.caption("🛠️ 多卡台买量消耗大盘与流水对账系统")


# =================【通用工具与数据清洗引擎】=================
def normalize_card_series(series):
    """向量化快速格式化卡号，剥离 .0 并左侧补零"""
    if series is None or series.empty:
        return pd.Series(dtype=str)
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    return s.str[-4:].str.zfill(4)


def get_latest_excel_path(folder_path, fallback_file):
    """自动获取指定文件夹中最新命名的 .xlsx 文件（按自然文件名升序，取最后一张）"""
    if os.path.exists(folder_path):
        excel_files = [
            f
            for f in glob.glob(os.path.join(folder_path, "*.xlsx"))
            if not os.path.basename(f).startswith("~$")
        ]
        if excel_files:
            return sorted(excel_files)[-1]

    if os.path.exists(fallback_file):
        return fallback_file

    return None


@st.cache_data(show_spinner=False)
def load_and_clean_raw_cached(
    raw_file_path, dict_file_path, raw_mtime, dict_mtime
):
    """带 Streamlit 内存级缓存的数据清洗引擎"""
    df_raw = pd.DataFrame()
    try:
        excel_file = pd.ExcelFile(raw_file_path)
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
                df_dict["UA名字"] = (
                    df_dict["UA名字"].astype(str).str.strip()
                )
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


# =================【看板核心渲染函数 (公开精简版)】=================
def render_dashboard(platform_name, folder_path, fallback_file, dict_file):
    file_path = get_latest_excel_path(folder_path, fallback_file)

    if not file_path or not os.path.exists(file_path):
        st.error(
            f"❌ 未在 `{folder_path}/` 找到任何 `.xlsx` 文件，且未找到备用文件 `{fallback_file}`。"
        )
        return

    raw_mtime = os.path.getmtime(file_path) if os.path.exists(file_path) else 0
    dict_mtime = (
        os.path.getmtime(dict_file) if os.path.exists(dict_file) else 0
    )

    df_raw = load_and_clean_raw_cached(
        file_path, dict_file, raw_mtime, dict_mtime
    )

    if df_raw.empty:
        st.warning(f"⚠️ `{file_path}` 中暂无有效流水数据。")
        return

    mtime_str = get_display_file_time(file_path, raw_mtime)
    dict_status = (
        f"✅ 已关联 `{dict_file}`"
        if os.path.exists(dict_file)
        else f"⚠️ 未找到 `{dict_file}` (显示为未分配)"
    )

    st.caption(
        f"📁 实时载入流水: `{file_path}` ｜ 🕒 生成时间: **{mtime_str}** ｜ 📖 字典状态: {dict_status} ｜ ⚡ 内存缓存已激活"
    )

    # ----------------------------------------------------
    # 板块 1：中部趋势图表（柱状图 + 环形饼图）
    # ----------------------------------------------------
    chart1, chart2 = st.columns(2)
    with chart1:
        if (
            "交易日期" in df_raw.columns
            and "交易金额" in df_raw.columns
            and not df_raw.empty
        ):
            daily_data = (
                df_raw.groupby(["交易日期", "交易状态"])["交易金额"]
                .sum()
                .reset_index()
            )
            fig_trend = px.bar(
                daily_data,
                x="交易日期",
                y="交易金额",
                color="交易状态",
                title=f"📅 {platform_name} - 每日消耗与状态趋势",
                barmode="stack",
                template="plotly_white",
                color_discrete_map={
                    "PENDING": "#2563eb",
                    "COMPLETE": "#10b981",
                    "REVERSED": "#f59e0b",
                    "FAILED": "#ef4444",
                },
            )
            fig_trend.update_layout(
                font_family="-apple-system, BlinkMacSystemFont, Segoe UI",
                title_font_size=15,
                legend_title_text="",
                margin=dict(l=10, r=10, t=40, b=10),
                hovermode="x unified",
            )
            st.plotly_chart(fig_trend, use_container_width=True)

    with chart2:
        if (
            "UA名字" in df_raw.columns
            and "交易金额" in df_raw.columns
            and not df_raw.empty
        ):
            df_pending_chart = (
                df_raw[df_raw["交易状态"] == "PENDING"]
                if "交易状态" in df_raw.columns
                else df_raw
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
                title=f"🎯 {platform_name} - 投放团队 PENDING 消耗占比",
                hole=0.45,
                template="plotly_white",
                color_discrete_sequence=px.colors.qualitative.Pastel,
            )
            fig_ua.update_traces(
                textposition="inside", textinfo="percent+label"
            )
            fig_ua.update_layout(
                font_family="-apple-system, BlinkMacSystemFont, Segoe UI",
                title_font_size=15,
                margin=dict(l=10, r=10, t=40, b=10),
            )
            st.plotly_chart(fig_ua, use_container_width=True)

    st.markdown("---")

    # ----------------------------------------------------
    # 板块 2：单卡每日消耗透视 (仅统计 PENDING)
    # ----------------------------------------------------
    st.subheader(f"💳 {platform_name} 单卡每日消耗透视 (仅统计 PENDING)")

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
                key=f"pivot_card_{platform_name}",
            )

        with p_col2:
            if "UA名字" in df_pending.columns:
                pivot_ua_options = sorted(
                    [
                        str(x)
                        for x in df_pending["UA名字"].dropna().unique()
                        if str(x).strip()
                    ]
                )
                selected_pivot_ua = st.multiselect(
                    "👤 筛选 UA",
                    options=pivot_ua_options,
                    default=[],
                    placeholder="按 UA 筛选（留空全选）",
                    key=f"pivot_ua_{platform_name}",
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
                key=f"pivot_date_range_{platform_name}",
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

            # 3 列 Metric 指标卡
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

            # Tab 1: 单卡明细（含置顶合计行）
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

            # Tab 2: UA 每日汇总表 (按 UA + 交易日期 求和)
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

            # Tab 3: Pivot 透视大表 (卡号 x 日期)
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
    st.subheader(f"📋 {platform_name} 全量流水对账")

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
                key=f"raw_card_{platform_name}",
            )

        with r_col2:
            raw_ua_options = (
                sorted(
                    [
                        str(x)
                        for x in df_raw["UA名字"].dropna().unique()
                        if str(x).strip()
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
                key=f"raw_ua_{platform_name}",
            )

        with r_col3:
            if "交易日期" in df_raw.columns and not df_raw["交易日期"].dropna().empty:
                r_min_date = df_raw["交易日期"].min()
                r_max_date = df_raw["交易日期"].max()
                selected_raw_date_range = st.date_input(
                    "📅 交易日期范围",
                    value=(r_min_date, r_max_date)
                    if (r_min_date and r_max_date)
                    else datetime.now().date(),
                    min_value=r_min_date,
                    max_value=r_max_date,
                    key=f"raw_date_range_{platform_name}",
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


# =================【3. 页面主入口】=================
tab1, tab2 = st.tabs(["🌲 Forest 清洗操作", "⚡ Tim 清洗操作"])

with tab1:
    cfg = PLATFORMS["Forest"]
    render_dashboard(
        "Forest", cfg["folder"], cfg["fallback_file"], cfg["dict_file"]
    )

with tab2:
    cfg = PLATFORMS["Tim"]
    render_dashboard(
        "Tim", cfg["folder"], cfg["fallback_file"], cfg["dict_file"]
    )
