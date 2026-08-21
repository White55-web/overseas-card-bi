import glob
import json
import os
from datetime import datetime, timedelta
import pandas as pd
import plotly.express as px
import streamlit as st

# =================【1. 页面基本配置】=================
st.set_page_config(
    page_title="出海多卡台数据中台", layout="wide", initial_sidebar_state="expanded"
)

# =================【全局高颜值 UI/CSS 注入】=================
st.markdown(
    """
    <style>
    /* 全局字体与版心间距 */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    .main .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
        max-width: 96%;
    }

    /* 顶部 KPI Metric 指标卡片立体化与悬浮效果 */
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

    /* 标签页 Tabs 现代圆角胶囊样式 */
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

    /* 折叠面板与表格容器圆角平滑化 */
    [data-testid="stExpander"] {
        border: 1px solid #e2e8f0 !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
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

# =================【侧边栏全局状态与一键刷新】=================
with st.sidebar:
    st.header("⚙️ 中台系统控制")
    if st.button("🔄 立即同步所有卡台数据", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")
    st.caption("🛠️ 数据中台自动化巡检系统 ｜ 实时对账 & 扣税预警 ｜ White制作")


# =================【通用工具与极速缓存清洗函数】=================
def merge_ua(series):
    unique_uas = [str(x) for x in series.dropna().unique() if str(x).strip()]
    return " / ".join(unique_uas) if unique_uas else "未知"


def normalize_card_series(series):
    """向量化快速格式化卡号，剥离 .0 并左侧补 0（比 apply 快 20 倍以上）"""
    if series is None or series.empty:
        return pd.Series(dtype=str)
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    return s.str[-4:].str.zfill(4)


def get_latest_excel_path(folder_path, fallback_file):
    """自动获取指定文件夹中最新修改的 .xlsx 文件，空则回退到备用文件"""
    if os.path.exists(folder_path):
        excel_files = [
            f
            for f in glob.glob(os.path.join(folder_path, "*.xlsx"))
            if not os.path.basename(f).startswith("~$")
        ]
        if excel_files:
            return max(excel_files, key=os.path.getmtime)

    if os.path.exists(fallback_file):
        return fallback_file

    return None


def analyze_threshold_ladder(recent_amounts, max_single):
    """分析卡号是否处于阈值跃升临界点，并计算跃升后的推荐补款金额"""
    LADDER = [
        (800, 900, "🔥 $900+ 顶格大户", None),
        (450, 500, "🚀 $500+ 核心放量", 900),
        (200, 250, "⚡ $250+ 标准放量", 500),
        (100, 125, "📈 $125+ 阶梯起量", 250),
        (40, 50, "🌱 $50 测品/小额", 125),
    ]

    tier_label = "🌱 $50 以下测品/小额"
    curr_bracket = 50
    next_bracket = 125

    for lower_bound, bracket_cap, label, nxt in LADDER:
        if max_single >= lower_bound:
            curr_bracket = bracket_cap
            next_bracket = nxt
            tier_label = label
            break

    if not next_bracket:
        return tier_label, False, max_single, max_single * 2

    consecutive_hits = 0
    for amt in reversed(recent_amounts):
        if amt >= curr_bracket * 0.90:
            consecutive_hits += 1
        else:
            break

    if consecutive_hits >= 3:
        tier_display = (
            f"🔥 临界跃升: 下次必扣 ${next_bracket} (已扣{consecutive_hits}笔${curr_bracket})"
        )
        safe_min = float(next_bracket)
        rec_topup = float(next_bracket * 2)
        is_leaping = True
    elif consecutive_hits == 2:
        tier_display = (
            f"⚠️ 跃升预警: 预计跳 ${next_bracket} (已扣2笔${curr_bracket})"
        )
        safe_min = float(next_bracket)
        rec_topup = float(next_bracket * 2)
        is_leaping = True
    else:
        tier_display = f"{tier_label} (稳态)"
        safe_min = max_single
        rec_topup = max_single * 2 if max_single >= 100 else 100.0
        is_leaping = False

    return tier_display, is_leaping, safe_min, rec_topup


@st.cache_data(show_spinner=False)
def load_and_clean_raw_cached(
    raw_file_path, dict_file_path, raw_mtime, dict_mtime
):
    """带 Streamlit 内存级缓存的数据清洗引擎（仅在文件产生变动时重新计算）"""
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


# =================【通用卡台看板渲染核心函数】=================
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

    mtime_str = datetime.fromtimestamp(raw_mtime).strftime("%Y-%m-%d %H:%M:%S")
    dict_status = (
        f"✅ 已关联 `{dict_file}`"
        if os.path.exists(dict_file)
        else f"⚠️ 未找到 `{dict_file}` (显示为未分配)"
    )
    st.caption(
        f"📁 实时载入流水: `{file_path}` ｜ 🕒 生成时间: **{mtime_str}** ｜ 📖 字典状态: {dict_status} ｜ ⚡ 内存缓存已激活"
    )

    # 1. 顶部主筛选区域
    df = df_raw.copy()
    with st.expander(
        f"🔍 展开 / 折叠【{platform_name}】主筛选条件", expanded=True
    ):
        f_col1, f_col2, f_col3 = st.columns(3)

        with f_col1:
            if "交易日期" in df_raw.columns and not df_raw["交易日期"].dropna().empty:
                min_date = df_raw["交易日期"].min()
                max_date = df_raw["交易日期"].max()
                selected_date_range = st.date_input(
                    "📅 交易日期范围",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date,
                    key=f"date_{platform_name}",
                )
                if (
                    isinstance(selected_date_range, tuple)
                    and len(selected_date_range) == 2
                ):
                    start_date, end_date = selected_date_range
                    df = df[
                        (df["交易日期"] >= start_date)
                        & (df["交易日期"] <= end_date)
                    ]
                elif (
                    isinstance(selected_date_range, tuple)
                    and len(selected_date_range) == 1
                ):
                    start_date = selected_date_range[0]
                    df = df[df["交易日期"] == start_date]

        with f_col2:
            if "UA名字" in df_raw.columns:
                ua_list = sorted(
                    [
                        str(x)
                        for x in df_raw["UA名字"].dropna().unique()
                        if str(x).strip()
                    ]
                )
                selected_ua = st.multiselect(
                    "👤 UA 名字",
                    options=ua_list,
                    default=[],
                    placeholder="留空默认展示所有人",
                    key=f"ua_{platform_name}",
                )
                if selected_ua:
                    df = df[df["UA名字"].isin(selected_ua)]
            else:
                selected_ua = []

        with f_col3:
            if "交易状态" in df_raw.columns:
                status_list = [
                    x
                    for x in df_raw["交易状态"].dropna().unique()
                    if str(x).strip()
                ]
                selected_status = st.multiselect(
                    "📌 交易状态",
                    options=status_list,
                    default=status_list,
                    key=f"status_{platform_name}",
                )
                if selected_status:
                    df = df[df["交易状态"].isin(selected_status)]

    # 2. 计算今日大盘流速与全天终局推算
    latest_global_date = (
        df_raw["交易日期"].max() if "交易日期" in df_raw.columns else None
    )
    today_pending_df = (
        df_raw[
            (df_raw["交易日期"] == latest_global_date)
            & (df_raw["交易状态"] == "PENDING")
        ]
        if (latest_global_date and "交易状态" in df_raw.columns)
        else pd.DataFrame()
    )

    hours_elapsed = 12.0
    max_dataset_datetime = None
    if not today_pending_df.empty and "交易时间" in today_pending_df.columns:
        max_dataset_datetime = today_pending_df["交易时间"].max()
        if pd.notnull(max_dataset_datetime):
            hours_elapsed = max(
                round(
                    max_dataset_datetime.hour
                    + max_dataset_datetime.minute / 60.0,
                    1,
                ),
                1.0,
            )

    today_pending_sum = (
        today_pending_df["交易金额"].sum() if not today_pending_df.empty else 0.0
    )
    burn_rate_hourly = (
        (today_pending_sum / hours_elapsed) if hours_elapsed > 0 else 0.0
    )
    projected_today_eod = burn_rate_hourly * 24.0

    # 3. 顶部 5 列 KPI 指标卡
    pending_spend = (
        df[df["交易状态"] == "PENDING"]["交易金额"].sum()
        if "交易状态" in df.columns and "交易金额" in df.columns
        else 0.0
    )
    complete_spend = (
        df[df["交易状态"] == "COMPLETE"]["交易金额"].sum()
        if "交易状态" in df.columns and "交易金额" in df.columns
        else 0.0
    )
    total_spend = pending_spend
    card_count = (
        df[df["交易状态"] == "PENDING"]["卡号"].nunique()
        if "交易状态" in df.columns and "卡号" in df.columns
        else df["卡号"].nunique()
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric(
            label="💰 总消耗 (PENDING)", value=f"${total_spend:,.2f}"
        )
    with c2:
        st.metric(
            label=f"⚡ 今日大盘时速 ({latest_global_date})",
            value=f"${burn_rate_hourly:,.1f} / h",
            help=f"截至最新扣款时间已跑约 {hours_elapsed:.1f} 小时，按 PENDING 实时扣费计算",
        )
    with c3:
        st.metric(
            label="🔮 预计今日大盘终局",
            value=f"${projected_today_eod:,.2f}",
            help=f"按当前 {burn_rate_hourly:,.1f}/h 时速外推全天 24 小时总消耗预测",
        )
    with c4:
        st.metric(
            label="✅ COMPLETE 历史扣费", value=f"${complete_spend:,.2f}"
        )
    with c5:
        st.metric(label="💳 PENDING 活跃卡数", value=f"{card_count} 张")

    # =================【功能 1：🚨 休眠/停滞卡（去重 + 持久化保存 + 状态锁定防折叠）】=================
    df_dormant_base = df_raw.copy()
    if selected_ua and "UA名字" in df_dormant_base.columns:
        df_dormant_base = df_dormant_base[
            df_dormant_base["UA名字"].isin(selected_ua)
        ]

    deleted_file = "deleted_cards.json"
    deleted_cards_set = set()
    if os.path.exists(deleted_file):
        try:
            with open(deleted_file, "r", encoding="utf-8") as f:
                deleted_cards_set = set(json.load(f))
        except Exception:
            deleted_cards_set = set()

    if (
        "卡号" in df_dormant_base.columns
        and "交易日期" in df_dormant_base.columns
        and not df_dormant_base.empty
    ):
        max_data_date = df_dormant_base["交易日期"].max()

        summary_rows = []
        for card_no, group in df_dormant_base.groupby("卡号"):
            last_date = group["交易日期"].max()
            spend_sum = group["交易金额"].sum()
            tx_cnt = len(group)
            ua_str = (
                merge_ua(group["UA名字"]) if "UA名字" in group.columns else ""
            )
            summary_rows.append(
                {
                    "卡号": card_no,
                    "UA名字": ua_str,
                    "最近一次消耗日期": last_date,
                    "累计总消耗": spend_sum,
                    "总交易笔数": tx_cnt,
                    "停滞天数": (max_data_date - last_date).days,
                }
            )

        card_summary = pd.DataFrame(summary_rows)
        dormant_cards = card_summary[card_summary["停滞天数"] >= 3].sort_values(
            by="停滞天数", ascending=False
        )
        total_dormant_count = len(dormant_cards)

        if total_dormant_count > 0:
            dormant_cards["已删卡"] = dormant_cards["卡号"].apply(
                lambda x: str(x) in deleted_cards_set
            )
            active_dormant_cards = dormant_cards[~dormant_cards["已删卡"]]
            real_dormant_count = len(active_dormant_cards)

            exp_state_key = f"dormant_exp_open_{platform_name}"
            is_dormant_open = st.session_state.get(exp_state_key, False)

            expander_label = f"🚨 【{platform_name}】休眠/停滞卡预警清单（待排查: {real_dormant_count} 张 ｜ 已删卡: {total_dormant_count - real_dormant_count} 张 ｜ 总计停滞: {total_dormant_count} 张）"
            with st.expander(expander_label, expanded=is_dormant_open):
                col_tip, col_filter = st.columns([3, 1])
                with col_filter:
                    view_mode = st.radio(
                        "显示范围",
                        ["全部停滞卡", "仅看待排查"],
                        horizontal=True,
                        key=f"dormant_view_{platform_name}",
                    )

                display_dormant = (
                    dormant_cards
                    if view_mode == "全部停滞卡"
                    else active_dormant_cards
                )

                if real_dormant_count > 0:
                    st.warning(
                        f"⚠️ 当前有 **{real_dormant_count}** 张独立卡号在最近 **3 天及以上** 未产生扣费且**未标记删卡**，请及时排查！"
                    )
                else:
                    st.success(
                        "🎉 表现优异：所有停滞老卡均已标记为【已删卡】，暂无待排查的异常卡！"
                    )

                col_cfg = {
                    "卡号": st.column_config.TextColumn("卡号", disabled=True),
                    "UA名字": st.column_config.TextColumn(
                        "UA名字", disabled=True
                    ),
                    "最近一次消耗日期": st.column_config.DateColumn(
                        "最近一次消耗日期", disabled=True
                    ),
                    "累计总消耗": st.column_config.NumberColumn(
                        "累计总消耗", format="$%.2f", disabled=True
                    ),
                    "总交易笔数": st.column_config.NumberColumn(
                        "交易笔数", format="%d 笔", disabled=True
                    ),
                    "停滞天数": st.column_config.NumberColumn(
                        "停滞天数", format="%d 天", disabled=True
                    ),
                    "已删卡": st.column_config.CheckboxColumn(
                        "已删卡？",
                        help="勾选后标记为已废弃老卡，不计入待排查预警",
                        default=False,
                    ),
                }

                edited_dormant = st.data_editor(
                    display_dormant,
                    column_config=col_cfg,
                    hide_index=True,
                    use_container_width=True,
                    key=f"editor_dormant_{platform_name}",
                )

                new_checked = set(
                    edited_dormant[edited_dormant["已删卡"] == True]["卡号"]
                    .astype(str)
                    .tolist()
                )
                new_unchecked = set(
                    edited_dormant[edited_dormant["已删卡"] == False]["卡号"]
                    .astype(str)
                    .tolist()
                )

                current_deleted = (
                    deleted_cards_set - new_unchecked
                ) | new_checked
                if current_deleted != deleted_cards_set:
                    with open(deleted_file, "w", encoding="utf-8") as f:
                        json.dump(
                            list(current_deleted), f, ensure_ascii=False, indent=2
                        )
                    st.session_state[exp_state_key] = True
                    st.rerun()
        else:
            with st.expander(
                f"🚨 【{platform_name}】休眠/停滞卡预警清单", expanded=False
            ):
                st.success(
                    "🎉 表现优异：当前所选条件下所有卡号在最近 3 天内均有活跃消耗！"
                )

    # =================【功能 2：💸 扣税卡智能排查（分流 + 偶发明细可查 + 默认折叠）】=================
    if "交易金额" in df.columns and "卡号" in df.columns and not df.empty:
        df_decimal = df[(df["交易金额"].round(2) % 1) != 0].copy()

        if not df_decimal.empty:
            group_keys = ["卡号"]
            if "UA名字" in df_decimal.columns:
                group_keys.append("UA名字")

            decimal_summary = (
                df_decimal.groupby(group_keys)
                .agg(
                    带小数扣款笔数=('交易金额', 'count'),
                    带小数扣费总额=('交易金额', 'sum'),
                    最近扣款时间=('交易时间', 'max')
                    if "交易时间" in df_decimal.columns
                    else ('交易日期', 'max'),
                )
                .reset_index()
            )

            high_tax_cards = decimal_summary[
                decimal_summary["带小数扣款笔数"] >= 3
            ].sort_values(by="带小数扣款笔数", ascending=False)
            sporadic_cards = decimal_summary[
                decimal_summary["带小数扣款笔数"] < 3
            ].sort_values(by="带小数扣款笔数", ascending=False)

            high_count = len(high_tax_cards)
            sporadic_count = len(sporadic_cards)

            expander_title = f"💸 【{platform_name}】扣税卡排查（🚨 持续扣税: {high_count} 张 ｜ ℹ️ 偶发结算: {sporadic_count} 张）"
            with st.expander(expander_title, expanded=False):
                if high_count > 0:
                    st.error(
                        f"🚨 **发现 {high_count} 张卡存在 $\ge 3$ 笔非整额扣款！** 属于持续跑量且被扣税的重点账户，请优先通知对应 UA 修改免税州："
                    )
                    st.dataframe(
                        high_tax_cards,
                        column_config={
                            "带小数扣款笔数": st.column_config.NumberColumn(
                                "异常扣费笔数", format="%d 笔"
                            ),
                            "带小数扣费总额": st.column_config.NumberColumn(
                                "扣费总金额", format="$%.2f"
                            ),
                        },
                        hide_index=True,
                        use_container_width=True,
                    )
                else:
                    st.success(
                        "🎉 当前没有发现持续产生多笔扣税的卡号，投放状态良好！"
                    )

                if sporadic_count > 0:
                    with st.expander(
                        f"ℹ️ 点击展开查看这 {sporadic_count} 张【偶发结算卡】明细（通常为拒付手动支付/关停结算，无需处理）",
                        expanded=False,
                    ):
                        st.dataframe(
                            sporadic_cards,
                            column_config={
                                "带小数扣款笔数": st.column_config.NumberColumn(
                                    "异常扣费笔数", format="%d 笔"
                                ),
                                "带小数扣费总额": st.column_config.NumberColumn(
                                    "扣费总金额", format="$%.2f"
                                ),
                            },
                            hide_index=True,
                            use_container_width=True,
                        )
        else:
            with st.expander(
                f"💸 【{platform_name}】扣税卡排查", expanded=False
            ):
                st.success("🎉 全绿通过：当前筛选条件下无任何非整额扣款！")

    # =================【功能 3：💳 阈值跃升预警 & 下次扣费倒计时雷达（智能高维补款决策）】=================
    df_thresh_base = (
        df_raw[df_raw["交易状态"] == "PENDING"].copy()
        if "交易状态" in df_raw.columns
        else df_raw.copy()
    )

    if (
        "卡号" in df_thresh_base.columns
        and "交易金额" in df_thresh_base.columns
        and not df_thresh_base.empty
    ):
        max_data_date = (
            df_thresh_base["交易日期"].max()
            if "交易日期" in df_thresh_base.columns
            else None
        )
        recent_3d_max_dict = {}
        if max_data_date:
            start_3d = max_data_date - timedelta(days=2)
            df_3d = df_thresh_base[df_thresh_base["交易日期"] >= start_3d]
            if not df_3d.empty:
                daily_3d = (
                    df_3d.groupby(["卡号", "交易日期"])["交易金额"]
                    .sum()
                    .reset_index()
                )
                recent_3d_max_dict = (
                    daily_3d.groupby("卡号")["交易金额"].max().to_dict()
                )

        threshold_rows = []
        leaping_count = 0
        urgent_bill_count = 0

        for card_no, group in df_thresh_base.groupby("卡号"):
            is_deleted = str(card_no) in deleted_cards_set
            max_single = group["交易金额"].max()
            sorted_group = group.sort_values(
                by="交易时间" if "交易时间" in group.columns else "交易日期"
            )
            last_single = sorted_group["交易金额"].iloc[-1]
            last_time = (
                sorted_group["交易时间"].iloc[-1]
                if "交易时间" in sorted_group.columns
                else sorted_group["交易日期"].iloc[-1]
            )
            ua_str = (
                merge_ua(group["UA名字"])
                if "UA名字" in group.columns
                else "未知"
            )
            max_3d_daily = recent_3d_max_dict.get(card_no, 0.0)

            # 1. 阈值阶梯跃升前瞻预测模型
            recent_amts = sorted_group["交易金额"].tolist()
            tier_display, is_leaping, safe_min_topup, rec_topup = (
                analyze_threshold_ladder(recent_amts, max_single)
            )
            if is_leaping and not is_deleted:
                leaping_count += 1

            # 2. 今日扣款脉搏节奏与下次扣款倒计时雷达
            card_today_tx = (
                group[group["交易日期"] == max_data_date]
                if max_data_date
                else pd.DataFrame()
            )
            today_tx_cnt = len(card_today_tx)

            if today_tx_cnt >= 2 and "交易时间" in card_today_tx.columns:
                sorted_tx = card_today_tx.sort_values(by="交易时间")
                time_diffs = (
                    sorted_tx["交易时间"].diff().dt.total_seconds().dropna()
                )
                avg_mins = (
                    (time_diffs.mean() / 60.0) if not time_diffs.empty else 0
                )

                if (
                    isinstance(last_time, (pd.Timestamp, datetime))
                    and avg_mins > 0
                ):
                    next_tx_dt = last_time + timedelta(minutes=avg_mins)
                    next_time_str = next_tx_dt.strftime("%H:%M")
                    mins_remaining = (
                        (next_tx_dt - max_dataset_datetime).total_seconds()
                        / 60.0
                        if max_dataset_datetime
                        else avg_mins
                    )

                    if avg_mins < 60:
                        if not is_deleted:
                            urgent_bill_count += 1
                        if mins_remaining <= 0:
                            pulse = (
                                f"🚨 极速抽水 ｜ 随时扣款 (预计 {next_time_str})"
                            )
                        else:
                            pulse = f"🚨 极速抽水 ｜ 预计 {next_time_str} (约 {int(mins_remaining)}分后)"
                    elif avg_mins <= 180:
                        if mins_remaining <= 0:
                            pulse = (
                                f"⚡ 快速放量 ｜ 随时扣款 (预计 {next_time_str})"
                            )
                        else:
                            pulse = f"⚡ 快速放量 ｜ 预计 {next_time_str} (约 {mins_remaining/60.0:.1f}h后)"
                    else:
                        pulse = f"🌱 稳定慢跑 ｜ 预计 {next_time_str} (均 {avg_mins/60.0:.1f}h/笔)"
                else:
                    if avg_mins < 60:
                        if not is_deleted:
                            urgent_bill_count += 1
                        pulse = f"🚨 极速抽水 (均 {int(avg_mins)}分/笔)"
                    elif avg_mins <= 180:
                        pulse = f"⚡ 快速放量 (均 {avg_mins/60.0:.1f}h/笔)"
                    else:
                        pulse = f"🌱 稳定慢跑 (均 {avg_mins/60.0:.1f}h/笔)"
            elif today_tx_cnt == 1:
                time_fmt = (
                    last_time.strftime("%H:%M")
                    if isinstance(last_time, (pd.Timestamp, datetime))
                    else ""
                )
                pulse = (
                    f"⏱️ 今日已扣 1 笔 ({time_fmt})"
                    if time_fmt
                    else "⏱️ 今日已扣 1 笔"
                )
            else:
                pulse = "⏸️ 今日暂无扣款"

            daily_standby = max(max_3d_daily, rec_topup)

            threshold_rows.append(
                {
                    "卡号": card_no,
                    "UA名字": ua_str,
                    "阈值档位画像": tier_display,
                    "今日扣款脉搏": pulse,
                    "历史最高单笔扣款": max_single,
                    "近3天单日最高消耗": max_3d_daily,
                    "最近一笔扣款": last_single,
                    "安全补款底线(1次)": safe_min_topup,
                    "推荐补款金额(2次缓冲)": rec_topup,
                    "建议全天备用额度": daily_standby,
                    "累计总扣款笔数": len(group),
                    "最近扣款时间": last_time,
                    "已删卡": is_deleted,
                }
            )

        df_threshold_all = pd.DataFrame(threshold_rows).sort_values(
            by="历史最高单笔扣款", ascending=False
        )

        expander_thresh_title = f"💳 【{platform_name}】单卡最高阈值画像与智能补款参考（🚨 {urgent_bill_count} 张极速抽水 ｜ 🔥 {leaping_count} 张即将跃升跳档 ｜ 仅统计 PENDING）"
        with st.expander(expander_thresh_title, expanded=False):
            t_col1, t_col2, t_col3 = st.columns([2, 2, 2])
            with t_col1:
                hide_deleted = st.checkbox(
                    "🚫 自动隐藏【已删卡】",
                    value=True,
                    help="勾选后自动过滤在休眠清单中已标记删卡的废弃卡号",
                    key=f"t_hide_del_{platform_name}",
                )

            df_threshold = (
                df_threshold_all[~df_threshold_all["已删卡"]]
                if hide_deleted
                else df_threshold_all
            )

            with t_col2:
                search_t_card = st.multiselect(
                    "🔍 搜索卡号",
                    options=sorted(list(df_threshold["卡号"].unique())),
                    default=[],
                    placeholder="可输入卡号后4位快速检索...",
                    key=f"t_search_card_{platform_name}",
                )
            with t_col3:
                search_t_ua = st.multiselect(
                    "👤 筛选 UA",
                    options=sorted(
                        [
                            x
                            for x in df_threshold["UA名字"].unique()
                            if x != "未知"
                        ]
                    ),
                    default=[],
                    placeholder="按 UA 筛选...",
                    key=f"t_search_ua_{platform_name}",
                )

            df_t_show = df_threshold.copy()
            if search_t_card:
                df_t_show = df_t_show[df_t_show["卡号"].isin(search_t_card)]
            if search_t_ua:
                df_t_show = df_t_show[df_t_show["UA名字"].isin(search_t_ua)]

            display_cols = [c for c in df_t_show.columns if c != "已删卡"]

            st.dataframe(
                df_t_show[display_cols],
                column_config={
                    "阈值档位画像": st.column_config.TextColumn(
                        "🎯 阈值画像与跃升预警",
                        help="实时检测连续达标笔数，自动预判 Meta 下次是否跳档 (如 $250 -> $500)",
                    ),
                    "今日扣款脉搏": st.column_config.TextColumn(
                        "⏱️ 今日扣款脉搏与倒计时雷达",
                        help="根据扣款频次与最近扣费时间，实时推算下一次扣款预计触发时间点",
                    ),
                    "历史最高单笔扣款": st.column_config.NumberColumn(
                        "历史最高单笔 (PENDING 阈值)", format="$%.2f"
                    ),
                    "近3天单日最高消耗": st.column_config.NumberColumn(
                        "🔥 近3天单日最高 (PENDING 爆发力)",
                        help="近3天单日 PENDING 消耗峰值，直接反映放量流速",
                        format="$%.2f",
                    ),
                    "最近一笔扣款": st.column_config.NumberColumn(
                        "最近一笔扣款", format="$%.2f"
                    ),
                    "安全补款底线(1次)": st.column_config.NumberColumn(
                        "⚠️ 绝对最低补款 (含跳档)",
                        help="补款绝不能低于此金额！检测到跃升时已自动升级为下一档金额，防止扣款瞬间触发 FAILED 拒付！",
                        format="$%.2f",
                    ),
                    "推荐补款金额(2次缓冲)": st.column_config.NumberColumn(
                        "💡 推荐充值额 (2次缓冲)",
                        help="预留2次扣款缓冲，防止频繁报单",
                        format="$%.2f",
                    ),
                    "建议全天备用额度": st.column_config.NumberColumn(
                        "🛡️ 建议全天备用额度",
                        help="结合近3天放量峰值与阈值缓冲计算，防止全天放量断流",
                        format="$%.2f",
                    ),
                    "累计总扣款笔数": st.column_config.NumberColumn(
                        "PENDING 笔数", format="%d 笔"
                    ),
                },
                hide_index=True,
                use_container_width=True,
            )

    # =================【功能 4：📈 UA 投放消耗波动与流速对比（智能双模态）】=================
    df_pending_all = (
        df_raw[df_raw["交易状态"] == "PENDING"]
        if "交易状态" in df_raw.columns
        else df_raw
    )
    if (
        "UA名字" in df_pending_all.columns
        and "交易日期" in df_pending_all.columns
        and not df_pending_all.empty
    ):
        dates_available = sorted(
            list(df_pending_all["交易日期"].dropna().unique()), reverse=True
        )
        if len(dates_available) >= 2:
            default_base_idx = 1 if len(dates_available) > 2 else 0
            default_comp_idx = 2 if len(dates_available) > 2 else 1

            with st.expander(
                f"📈 【{platform_name}】UA 投放消耗波动对比（环比与流速推算）",
                expanded=False,
            ):
                d_col1, d_col2, d_col3 = st.columns([2, 2, 3])
                with d_col1:
                    base_date = st.selectbox(
                        "🎯 分析基准日",
                        options=dates_available,
                        index=default_base_idx,
                        key=f"dod_base_{platform_name}",
                    )
                with d_col2:
                    comp_options = [
                        d for d in dates_available if d != base_date
                    ]
                    comp_date = st.selectbox(
                        "📅 对照参照日",
                        options=comp_options
                        if comp_options
                        else dates_available,
                        index=1 if len(comp_options) > 1 else 0,
                        key=f"dod_comp_{platform_name}",
                    )

                is_today = base_date == dates_available[0]
                with d_col3:
                    if is_today:
                        st.info(
                            f"⚡ 当前分析日为**进行中单日 (已跑约 {hours_elapsed:.1f}h)**，已自动激活 **【时速与全天终局推算模式】**！"
                        )
                    else:
                        st.success(
                            "✅ 当前对比均为**已结算完整自然日**，数据环比真实客观。"
                        )

                df_today = df_pending_all[
                    df_pending_all["交易日期"] == base_date
                ]
                df_prev = df_pending_all[
                    df_pending_all["交易日期"] == comp_date
                ]

                today_spend = df_today.groupby("UA名字")["交易金额"].sum()
                prev_spend = df_prev.groupby("UA名字")["交易金额"].sum()
                all_uas = sorted(
                    list(set(today_spend.index) | set(prev_spend.index))
                )

                dod_rows = []
                for u in all_uas:
                    if not str(u).strip() or str(u) == "nan":
                        continue
                    t_val = today_spend.get(u, 0.0)
                    p_val = prev_spend.get(u, 0.0)

                    if is_today:
                        ua_hourly = (
                            (t_val / hours_elapsed) if hours_elapsed > 0 else 0
                        )
                        ua_projected = ua_hourly * 24.0
                        proj_diff = ua_projected - p_val
                        proj_pct = (
                            (proj_diff / p_val * 100)
                            if p_val > 0
                            else (100.0 if ua_projected > 0 else 0.0)
                        )

                        if p_val == 0 and t_val > 0:
                            status = "🆕 今日新起跑量"
                        elif t_val == 0 and p_val > 0:
                            status = "⏸️ 今日暂停消耗"
                        elif proj_pct >= 20:
                            status = f"🚀 强劲放量中 (推算 +{proj_pct:.1f}%)"
                        elif proj_pct <= -20:
                            status = f"🔻 减速缩量中 (推算 {proj_pct:.1f}%)"
                        else:
                            status = f"⚡ 稳定平跑 (推算 {proj_pct:+.1f}%)"

                        dod_rows.append(
                            {
                                "UA名字": u,
                                "今日已跑消耗": t_val,
                                "当前时速 ($/h)": ua_hourly,
                                "预计全天终局消耗": ua_projected,
                                "对照日消耗 (完整)": p_val,
                                "放量趋势预判": status,
                            }
                        )
                    else:
                        diff_val = t_val - p_val
                        pct_val = (
                            (diff_val / p_val * 100)
                            if p_val > 0
                            else (100.0 if t_val > 0 else 0.0)
                        )

                        if p_val == 0 and t_val > 0:
                            status = "🆕 当日新增跑量"
                        elif t_val == 0 and p_val > 0:
                            status = "⏸️ 当日暂停消耗"
                        elif diff_val > 0:
                            status = f"🚀 放量增长 (+{pct_val:.1f}%)"
                        elif diff_val < 0:
                            status = f"🔻 缩量减少 ({pct_val:.1f}%)"
                        else:
                            status = "持平"

                        dod_rows.append(
                            {
                                "UA名字": u,
                                "基准日消耗 (PENDING)": t_val,
                                "对照日消耗 (PENDING)": p_val,
                                "消耗差额": diff_val,
                                "环比增幅": pct_val,
                                "放量动态": status,
                            }
                        )

                if is_today:
                    df_dod = pd.DataFrame(dod_rows).sort_values(
                        by="预计全天终局消耗", ascending=False
                    )
                    st.dataframe(
                        df_dod,
                        column_config={
                            "今日已跑消耗": st.column_config.NumberColumn(
                                f"今日已跑 ({base_date})", format="$%.2f"
                            ),
                            "当前时速 ($/h)": st.column_config.NumberColumn(
                                "⚡ 当前时速", format="$%.1f / h"
                            ),
                            "预计全天终局消耗": st.column_config.NumberColumn(
                                "🔮 预计全天终局", format="$%.2f"
                            ),
                            "对照日消耗 (完整)": st.column_config.NumberColumn(
                                f"参照日 ({comp_date})", format="$%.2f"
                            ),
                        },
                        hide_index=True,
                        use_container_width=True,
                    )
                else:
                    df_dod = pd.DataFrame(dod_rows).sort_values(
                        by="基准日消耗 (PENDING)", ascending=False
                    )
                    st.dataframe(
                        df_dod,
                        column_config={
                            "基准日消耗 (PENDING)": st.column_config.NumberColumn(
                                f"分析日 ({base_date})", format="$%.2f"
                            ),
                            "对照日消耗 (PENDING)": st.column_config.NumberColumn(
                                f"参照日 ({comp_date})", format="$%.2f"
                            ),
                            "消耗差额": st.column_config.NumberColumn(
                                "波动差额 ($)", format="$%.2f"
                            ),
                            "环比增幅": st.column_config.NumberColumn(
                                "环比增减", format="%.1f%%"
                            ),
                        },
                        hide_index=True,
                        use_container_width=True,
                    )

    # =================【5. 中部趋势图表（商务高级配色）】=================
    chart1, chart2 = st.columns(2)
    with chart1:
        if "交易日期" in df.columns and "交易金额" in df.columns and not df.empty:
            daily_data = (
                df.groupby(["交易日期", "交易状态"])["交易金额"]
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
        if "UA名字" in df.columns and "交易金额" in df.columns and not df.empty:
            df_pending_chart = (
                df[df["交易状态"] == "PENDING"]
                if "交易状态" in df.columns
                else df
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

    # =================【6. 单卡每日消耗透视（仅聚合 PENDING + 局部独立筛选 + 实时 SUM 汇总）】=================
    st.subheader(f"💳 {platform_name} 单卡每日消耗透视 (仅统计 PENDING)")

    df_pending = (
        df[df["交易状态"] == "PENDING"] if "交易状态" in df.columns else df
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
            date_options = sorted(
                list(df_pending["交易日期"].dropna().unique()), reverse=True
            )
            selected_pivot_dates = st.multiselect(
                "📅 筛选指定日期",
                options=date_options,
                default=[],
                placeholder="按具体单日/多日筛选（留空全选）",
                key=f"pivot_date_{platform_name}",
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
        if selected_pivot_dates and "交易日期" in df_pivot_filtered.columns:
            df_pivot_filtered = df_pivot_filtered[
                df_pivot_filtered["交易日期"].isin(selected_pivot_dates)
            ]

        if not df_pivot_filtered.empty:
            # 实时计算当前筛选维度的总消耗 (SUM)
            sum_spend = df_pivot_filtered["交易金额"].sum()
            sum_cards = df_pivot_filtered["卡号"].nunique()
            sum_tx = len(df_pivot_filtered)

            # 顶部实时汇总 Metric 卡片条
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
                        当日总消耗=('交易金额', 'sum'),
                        交易笔数=('交易金额', 'count'),
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
                    top_sum_dict["交易日期"] = (
                        str(selected_pivot_dates[0])
                        if len(selected_pivot_dates) == 1
                        else f"共 {df_pivot_filtered['交易日期'].nunique()} 天"
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

            # Tab 2: UA 每日汇总表 (按 UA + 交易日期 纯聚合求和)
            with view_tab2:
                if "UA名字" in df_pivot_filtered.columns:
                    ua_daily_df = (
                        df_pivot_filtered.groupby(["UA名字", "交易日期"])
                        .agg(
                            当日UA总消耗=('交易金额', 'sum'),
                            消耗卡数=('卡号', 'nunique'),
                            扣费笔数=('交易金额', 'count'),
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

            # Tab 3: Pivot 透视大表 (卡号 x 日期 + 期间总消耗)
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

    # =================【7. 底部全量明细流水（专属局部快速筛选）】=================
    st.subheader(f"📋 {platform_name} 全量流水对账")

    if not df.empty:
        r_col1, r_col2, r_col3 = st.columns(3)

        with r_col1:
            raw_card_options = (
                sorted(list(df["卡号"].dropna().unique()))
                if "卡号" in df.columns
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
                        for x in df["UA名字"].dropna().unique()
                        if str(x).strip()
                    ]
                )
                if "UA名字" in df.columns
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
            raw_date_options = (
                sorted(
                    list(df["交易日期"].dropna().unique()), reverse=True
                )
                if "交易日期" in df.columns
                else []
            )
            selected_raw_dates = st.multiselect(
                "📅 筛选指定日期",
                options=raw_date_options,
                default=[],
                placeholder="按具体单日/多日筛选（留空全选）",
                key=f"raw_date_{platform_name}",
            )

        df_raw_filtered = df.copy()
        if selected_raw_cards and "卡号" in df_raw_filtered.columns:
            df_raw_filtered = df_raw_filtered[
                df_raw_filtered["卡号"].isin(selected_raw_cards)
            ]
        if selected_raw_ua and "UA名字" in df_raw_filtered.columns:
            df_raw_filtered = df_raw_filtered[
                df_raw_filtered["UA名字"].isin(selected_raw_ua)
            ]
        if selected_raw_dates and "交易日期" in df_raw_filtered.columns:
            df_raw_filtered = df_raw_filtered[
                df_raw_filtered["交易日期"].isin(selected_raw_dates)
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