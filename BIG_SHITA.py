import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import re
from io import BytesIO

st.set_page_config(page_title="维修分析（多Shit增强版）", layout="wide")
st.title("🔧 维修分析（多Shit增强版）")

report_file = st.file_uploader("上传维修报告（Excel）", type=["xlsx"], key="repair_file")
sku_file = st.file_uploader("上传SKU对照表（可选，Excel/CSV）", type=["xlsx", "csv"], key="sku_map_file")


# =============================
# 工具函数
# =============================
def normalize_colname(col):
    return str(col).strip().lower()


def safe_columns(df):
    df.columns = [normalize_colname(c) for c in df.columns]
    return df


def ensure_column(df, col, default=None):
    if col not in df.columns:
        df[col] = default
    return df


def to_numeric_safe(series, default=0):
    return pd.to_numeric(series, errors="coerce").fillna(default)


def normalize_text_series(series, default="未知"):
    return (
        series.fillna(default)
        .astype(str)
        .str.strip()
        .replace("", default)
    )


def normalize_sku_value(x):
    if pd.isna(x):
        return None

    s = str(x).strip()

    if s.lower() in ["", "nan", "none", "null"]:
        return None

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    return s


def find_replaced_sku_columns(df):
    sku_cols = []
    for col in df.columns:
        col_str = str(col).strip().lower()
        if col_str.startswith("replaced sku"):
            suffix = col_str.replace("replaced sku", "").strip()
            if suffix == "" or suffix.isdigit():
                sku_cols.append(col)
    return sku_cols


def calc_tat(row):
    try:
        if pd.isna(row["received_date"]) or pd.isna(row["shipment_date"]):
            return None
        if row["shipment_date"] < row["received_date"]:
            return None
        return np.busday_count(
            row["received_date"].date(),
            row["shipment_date"].date()
        )
    except Exception:
        return None


def safe_ratio(series):
    total = series.sum()
    if total == 0 or pd.isna(total):
        return series * 0
    return series / total


# =============================
# 模糊匹配函数
# =============================
def normalize_for_match(x):
    """
    做模糊匹配前统一清洗：
    - 转字符串
    - 去首尾空格
    - 转小写
    - 把 _, -, /, \ 替换为空格
    - 压缩多余空格
    """
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = re.sub(r"[_\-/\\]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s


def contains_keyword_fuzzy(value, keywords):
    """
    只要文本中包含关键词，就算命中
    """
    text = normalize_for_match(value)
    if text == "":
        return None

    for k in keywords:
        k_norm = normalize_for_match(k)
        if k_norm and k_norm in text:
            return k.upper()
    return None


def map_agent_repair_report(sales_channel, model):
    sales_keywords = ["CESV", "Carrefour", "Feuvert", "Conforama"]
    model_keywords = ["GOLF", "BIRDIE 3X", "BIRDIE 3"]

    sales_hit = contains_keyword_fuzzy(sales_channel, sales_keywords)
    if sales_hit:
        return sales_hit

    model_text = normalize_for_match(model)
    if any(normalize_for_match(k) in model_text for k in model_keywords):
        return "GOLF"

    return "MCR"


def map_agent_additional_activity(client):
    """
    Additional Activity 的代理模糊归类
    """
    client_keywords = ["CESV", "Carrefour", "Feuvert", "Conforama"]
    client_hit = contains_keyword_fuzzy(client, client_keywords)
    if client_hit:
        return client_hit
    return "MCR"


def normalize_activity_name(activity):
    """
    Additional Activity 的 Activity 模糊识别标准化
    """
    text = normalize_for_match(activity)

    if text == "":
        return None

    # 1. COMUNICATION
    if (
        "comunication" in text or
        "communication" in text or
        "comunications" in text or
        "communications" in text
    ):
        return "COMUNICATION"

    # 2. CALLS
    if "call" in text or "calls" in text:
        return "CALLS"

    # 3. BOXES
    if "box" in text or "boxes" in text:
        return "BOXES"

    # 4. WORTEN RECEPTION
    if (
        ("worten" in text and "reception" in text) or
        "worten reception" in text
    ):
        return "WORTEN RECEPTION"

    # 5. BATTERY VOLTAGE CHECK
    if (
        ("battery" in text and "voltage" in text and "check" in text) or
        "battery voltage" in text
    ):
        return "BATTERY VOLTAGE CHECK"

    # 6. DOAS Management
    if (
        ("doa" in text and "management" in text) or
        "doas management" in text or
        "doa management" in text
    ):
        return "DOAS Management"

    # 7. ANOVO Stock Transfer
    if (
        ("anovo" in text and "stock" in text and "transfer" in text) or
        "stock transfer" in text
    ):
        return "ANOVO Stock Transfer"

    return None


# =============================
# SKU对照表
# =============================
def load_sku_mapping(uploaded_sku_file):
    if uploaded_sku_file is None:
        return None

    try:
        if uploaded_sku_file.name.endswith(".csv"):
            sku_map_df = pd.read_csv(uploaded_sku_file)
        else:
            sku_map_df = pd.read_excel(uploaded_sku_file)

        sku_map_df.columns = sku_map_df.columns.astype(str).str.strip().str.lower()

        rename_map = {
            "sku": "SKU",
            "sku code": "SKU",
            "item sku": "SKU",
            "part sku": "SKU",
            "中文名称": "中文名称",
            "名称": "中文名称",
            "name": "中文名称",
            "description": "中文名称",
            "desc": "中文名称",
            "part name": "中文名称"
        }
        sku_map_df = sku_map_df.rename(columns={k: v for k, v in rename_map.items() if k in sku_map_df.columns})

        if "SKU" not in sku_map_df.columns:
            st.warning("SKU对照表缺少 SKU 列，将忽略该对照表。")
            return None

        if "中文名称" not in sku_map_df.columns:
            sku_map_df["中文名称"] = "未提供"

        sku_map_df["SKU"] = sku_map_df["SKU"].apply(normalize_sku_value)
        sku_map_df["中文名称"] = sku_map_df["中文名称"].astype(str).str.strip()
        sku_map_df = sku_map_df.dropna(subset=["SKU"])
        sku_map_df = sku_map_df[["SKU", "中文名称"]].drop_duplicates(subset=["SKU"])

        return sku_map_df

    except Exception as e:
        st.error(f"SKU对照表读取失败: {e}")
        return None


def attach_sku_name(df_sku_result, sku_map_df):
    if df_sku_result is None or df_sku_result.empty:
        return df_sku_result

    result = df_sku_result.copy()

    if "SKU" in result.columns:
        result["SKU"] = result["SKU"].apply(normalize_sku_value)

    if sku_map_df is None:
        if "中文名称" not in result.columns:
            result["中文名称"] = "未匹配"
        return result

    merged = result.merge(sku_map_df, on="SKU", how="left")
    merged["中文名称"] = merged["中文名称"].fillna("未匹配")
    return merged


def to_excel_download(data_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


# =============================
# 读取工作簿
# =============================
@st.cache_data(show_spinner=False)
def load_workbook(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for s in xls.sheet_names:
        sheets[s] = pd.read_excel(uploaded_file, sheet_name=s)
    return sheets, xls.sheet_names


def parse_overview_sheet(df_raw):
    """
    Übersicht 的真实表头可能不在第一行
    尝试自动找到 Position / Quantity / Price 所在行
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=["position", "quantity", "price"])

    df = df_raw.copy()

    header_idx = None
    for i in range(min(len(df), 10)):
        row_vals = [str(x).strip().lower() for x in df.iloc[i].tolist()]
        if "position" in row_vals and "quantity" in row_vals and "price" in row_vals:
            header_idx = i
            break

    if header_idx is None:
        df = df.iloc[:, :3].copy()
        df.columns = ["position", "quantity", "price"]
    else:
        df.columns = [str(x).strip().lower() for x in df.iloc[header_idx].tolist()]
        df = df.iloc[header_idx + 1:].copy()

    df = safe_columns(df)
    df = df.rename(columns={
        "position": "position",
        "quantity": "quantity",
        "price": "price"
    })

    for col in ["position", "quantity", "price"]:
        df = ensure_column(df, col)

    df["position"] = normalize_text_series(df["position"], default="")
    df = df[df["position"] != ""].copy()
    df["quantity"] = to_numeric_safe(df["quantity"], default=0)
    df["price"] = to_numeric_safe(df["price"], default=0)

    return df[["position", "quantity", "price"]]


def parse_repair_report(df_raw):
    df = df_raw.copy()
    df = safe_columns(df)

    rename_map = {
        "repair order no.": "repair_id",
        "repair order": "repair_id",
        "repair id": "repair_id",
        "date of receipt": "received_date",
        "date of shipment": "shipment_date",
        "nation /state": "country",
        "nation /state ": "country",
        "nation/state": "country",
        "nation": "country",
        "sales channal": "sales_channel",
        "sales channal ": "sales_channel",
        "sales channel": "sales_channel",
        "customer name": "customer_name",
        "order id": "order_id",
        "model": "model",
        "warranty status": "repair_type",
        "problem description by customer": "customer_issue",
        "problem description by customer ": "customer_issue",
        "problem description by avono": "issue_desc",
        "problem description by avono ": "issue_desc",
        "responsible person (repair man )": "technician",
        "responsible person (repair man)": "technician",
        "sn": "sn",
        "repair fee": "repair_fee",
        "return shipment fee": "shipping_fee",
        "resend shipment fee": "resend_shipping_fee",
        "scrap fee": "scrap_fee",
        "other fee": "other_fee"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    for col in [
        "repair_id", "received_date", "shipment_date", "country", "sales_channel",
        "customer_name", "order_id", "model", "repair_type", "issue_desc",
        "technician", "sn", "repair_fee", "shipping_fee", "resend_shipping_fee",
        "scrap_fee", "other_fee"
    ]:
        df = ensure_column(df, col)

    df["received_date"] = pd.to_datetime(df["received_date"], errors="coerce")
    df["shipment_date"] = pd.to_datetime(df["shipment_date"], errors="coerce")

    type_map = {
        "iw": "保内",
        "ow": "保外",
        "doa": "DOA"
    }
    df["repair_type"] = (
        df["repair_type"]
        .astype(str)
        .str.strip()
        .str.lower()
        .map(type_map)
        .fillna(df["repair_type"])
    )

    for col in ["repair_fee", "shipping_fee", "resend_shipping_fee", "scrap_fee", "other_fee"]:
        df[col] = to_numeric_safe(df[col], default=0)

    df["total_cost"] = (
        df["repair_fee"]
        + df["shipping_fee"]
        + df["resend_shipping_fee"]
        + df["scrap_fee"]
        + df["other_fee"]
    )

    for col in ["country", "sales_channel", "customer_name", "model", "issue_desc", "technician"]:
        df[col] = normalize_text_series(df[col])

    df["sn"] = df["sn"].fillna("").astype(str).str.strip()
    df["TAT"] = df.apply(calc_tat, axis=1)

    df = df.sort_values(by=["sn", "received_date"], na_position="last")
    df["repeat"] = df.duplicated(subset=["sn"], keep="first")

    df["agent"] = df.apply(
        lambda r: map_agent_repair_report(r.get("sales_channel"), r.get("model")),
        axis=1
    )

    return df


def parse_additional_activity(df_raw):
    df = df_raw.copy()
    df = safe_columns(df)

    rename_map = {
        "activity": "activity",
        "client": "client",
        "price": "price"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    for col in ["activity", "client", "price"]:
        df = ensure_column(df, col)

    df["activity"] = normalize_text_series(df["activity"])
    df["client"] = normalize_text_series(df["client"])
    df["price"] = to_numeric_safe(df["price"], default=0)

    # 代理模糊识别
    df["agent"] = df["client"].apply(map_agent_additional_activity)

    # Activity 模糊识别并标准化
    df["activity_std"] = df["activity"].apply(normalize_activity_name)

    # 只保留能识别到目标类目的数据
    filtered = df[df["activity_std"].notna()].copy()

    return df, filtered


def parse_ow_sheet(df_raw):
    df = df_raw.copy()
    df = safe_columns(df)

    rename_map = {
        "order id": "order_id",
        "model": "model",
        "sn": "sn",
        "replaced sku": "SKU",
        "precio boom euros": "unit_price"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    for col in ["order_id", "model", "sn", "SKU", "unit_price"]:
        df = ensure_column(df, col)

    df["order_id"] = normalize_text_series(df["order_id"], default="")
    df["model"] = normalize_text_series(df["model"])
    df["sn"] = normalize_text_series(df["sn"])
    df["SKU"] = df["SKU"].apply(normalize_sku_value)
    df["unit_price"] = to_numeric_safe(df["unit_price"], default=0)

    df = df[df["order_id"] != ""].copy()
    return df


# =============================
# 主逻辑
# =============================
if report_file:
    try:
        sheets, sheet_names = load_workbook(report_file)
    except Exception as e:
        st.error(f"维修报告读取失败：{e}")
        st.stop()

    sku_map_df = load_sku_mapping(sku_file)

    if sku_map_df is not None:
        st.success("✅ 已加载SKU对照表")
    else:
        st.info("未上传有效SKU对照表，涉及SKU的地方将仅显示SKU编码")

    overview_raw = None
    repair_raw = None
    doa_raw = None
    add_raw = None
    ow_raw = None

    for s in sheet_names:
        s_std = s.strip().lower()
        if s_std == "übersicht" or s_std == "ubersicht":
            overview_raw = sheets[s]
        elif s_std == "repair report":
            repair_raw = sheets[s]
        elif "doa report" in s_std:
            doa_raw = sheets[s]
        elif s_std == "additional activity":
            add_raw = sheets[s]
        elif s_std == "ow":
            ow_raw = sheets[s]

    if repair_raw is None:
        st.error("未找到 Repair Report sheet，请检查文件。")
        st.stop()

    overview_df = parse_overview_sheet(overview_raw) if overview_raw is not None else pd.DataFrame()
    repair_df = parse_repair_report(repair_raw)
    doa_df = doa_raw.copy() if doa_raw is not None else pd.DataFrame()
    add_all_df, add_filtered_df = parse_additional_activity(add_raw) if add_raw is not None else (pd.DataFrame(), pd.DataFrame())
    ow_df = parse_ow_sheet(ow_raw) if ow_raw is not None else pd.DataFrame()

    # =============================
    # 侧边栏筛选（Repair Report）
    # =============================
    st.sidebar.header("Repair Report 筛选条件")

    country_vals = sorted(repair_df["country"].dropna().unique().tolist())
    type_vals = sorted(repair_df["repair_type"].dropna().astype(str).unique().tolist())
    model_vals = sorted(repair_df["model"].dropna().unique().tolist())
    agent_vals = sorted(repair_df["agent"].dropna().unique().tolist())

    country_filter = st.sidebar.multiselect("国家", country_vals, default=country_vals)
    type_filter = st.sidebar.multiselect("维修类型", type_vals, default=type_vals)
    model_filter = st.sidebar.multiselect("Model", model_vals, default=model_vals)
    agent_filter = st.sidebar.multiselect("代理", agent_vals, default=agent_vals)

    min_date = repair_df["received_date"].min() if repair_df["received_date"].notna().any() else pd.Timestamp.today()
    max_date = repair_df["received_date"].max() if repair_df["received_date"].notna().any() else pd.Timestamp.today()
    date_range = st.sidebar.date_input("日期范围", [min_date, max_date])

    repair_filtered = repair_df.copy()

    if country_filter:
        repair_filtered = repair_filtered[repair_filtered["country"].isin(country_filter)]
    if type_filter:
        repair_filtered = repair_filtered[repair_filtered["repair_type"].isin(type_filter)]
    if model_filter:
        repair_filtered = repair_filtered[repair_filtered["model"].isin(model_filter)]
    if agent_filter:
        repair_filtered = repair_filtered[repair_filtered["agent"].isin(agent_filter)]
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        repair_filtered = repair_filtered[
            (repair_filtered["received_date"] >= pd.to_datetime(date_range[0])) &
            (repair_filtered["received_date"] <= pd.to_datetime(date_range[1]))
        ]

    tabs = st.tabs([
        "总览 Overview",
        "Repair Report",
        "Additional Activity",
        "OW",
        "Storage",
        "DOA"
    ])

    # =============================
    # TAB 1 - 总览 Overview
    # =============================
    with tabs[0]:
        st.subheader("📋 Übersicht 总览")

        if overview_df.empty:
            st.info("未检测到 Übersicht 数据。")
        else:
            overview_show = overview_df.copy()
            overview_show.columns = ["Position", "Quantity", "Price"]
            st.dataframe(overview_show, use_container_width=True)

            def get_price_by_position(df, keyword):
                mask = df["position"].astype(str).str.strip().str.lower() == keyword.lower()
                if mask.any():
                    return df.loc[mask, "price"].sum()
                return 0

            total_price = get_price_by_position(overview_df, "Total")
            ow_parts_price = get_price_by_position(overview_df, "OW Parts")
            total_invoice = total_price - ow_parts_price

            c1, c2, c3 = st.columns(3)
            c1.metric("Total", f"{total_price:,.2f}")
            c2.metric("OW Parts", f"{ow_parts_price:,.2f}")
            c3.metric("Total Invoice", f"{total_invoice:,.2f}")

            pos_chart_df = overview_df.copy()
            pos_chart_df["position"] = pos_chart_df["position"].replace("", "未知")
            st.altair_chart(
                alt.Chart(pos_chart_df).mark_bar().encode(
                    x=alt.X("position:N", title="Position", sort="-y"),
                    y=alt.Y("price:Q", title="Price"),
                    tooltip=["position", "quantity", "price"]
                ),
                use_container_width=True
            )

    # =============================
    # TAB 2 - Repair Report
    # =============================
    with tabs[1]:
        st.subheader("🛠 Repair Report 分析")

        if repair_filtered.empty:
            st.warning("当前筛选条件下 Repair Report 没有数据。")
        else:
            df = repair_filtered.copy()

            df_iw = df[df["repair_type"] == "保内"]
            avg_tat = df_iw["TAT"].dropna().mean() if not df_iw.empty else 0
            rate_5 = (df_iw["TAT"] <= 5).mean() if not df_iw.empty else 0
            rate_10 = (df_iw["TAT"] <= 10).mean() if not df_iw.empty else 0
            repeat_rate = df["repeat"].mean() if not df.empty else 0
            doa_rate = (df["repair_type"] == "DOA").mean() if "repair_type" in df.columns else 0

            st.subheader("📊 核心指标")
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("平均TAT(保内)", round(avg_tat, 1) if not pd.isna(avg_tat) else 0)
            col2.metric("5天完成率", f"{rate_5:.1%}")
            col3.metric("10天完成率", f"{rate_10:.1%}")
            col4.metric("重复维修率", f"{repeat_rate:.1%}")
            col5.metric("DOA占比", f"{doa_rate:.1%}")

            st.subheader("📈 维修趋势")
            if df["received_date"].notna().any():
                df["month"] = df["received_date"].dt.to_period("M").astype(str)
                trend = df.groupby("month").size().reset_index(name="count")

                trend_chart = alt.Chart(trend).mark_line(point=True).encode(
                    x=alt.X("month:N", title="月份", sort=None),
                    y=alt.Y("count:Q", title="维修数量"),
                    tooltip=["month", "count"]
                )
                st.altair_chart(trend_chart, use_container_width=True)
            else:
                st.info("无有效日期数据，无法生成趋势图")

            st.subheader("📦 结构分析")
            c1, c2, c3 = st.columns(3)

            repair_type_dist = df["repair_type"].fillna("未知").value_counts(normalize=True).reset_index()
            repair_type_dist.columns = ["repair_type", "ratio"]
            c1.altair_chart(
                alt.Chart(repair_type_dist).mark_bar().encode(
                    x=alt.X("repair_type:N", title="维修类型"),
                    y=alt.Y("ratio:Q", title="占比"),
                    tooltip=["repair_type", alt.Tooltip("ratio:Q", format=".1%")]
                ),
                use_container_width=True
            )
            c1.write("维修类型占比")

            country_dist = df["country"].fillna("未知").value_counts(normalize=True).reset_index()
            country_dist.columns = ["country", "ratio"]
            c2.altair_chart(
                alt.Chart(country_dist).mark_bar().encode(
                    x=alt.X("country:N", title="国家"),
                    y=alt.Y("ratio:Q", title="占比"),
                    tooltip=["country", alt.Tooltip("ratio:Q", format=".1%")]
                ),
                use_container_width=True
            )
            c2.write("国家占比")

            cost_ratio = df.groupby("country", dropna=False)["total_cost"].sum().fillna(0)
            cost_ratio = safe_ratio(cost_ratio).reset_index()
            cost_ratio.columns = ["country", "ratio"]
            c3.altair_chart(
                alt.Chart(cost_ratio).mark_bar().encode(
                    x=alt.X("country:N", title="国家"),
                    y=alt.Y("ratio:Q", title="费用占比"),
                    tooltip=["country", alt.Tooltip("ratio:Q", format=".1%")]
                ),
                use_container_width=True
            )
            c3.write("国家费用占比")

            st.subheader("🏷 代理分析")
            a1, a2 = st.columns(2)

            agent_count = (
                df["agent"]
                .value_counts()
                .reset_index()
            )
            agent_count.columns = ["agent", "count"]

            a1.altair_chart(
                alt.Chart(agent_count).mark_bar().encode(
                    x=alt.X("agent:N", title="代理"),
                    y=alt.Y("count:Q", title="维修数量"),
                    tooltip=["agent", "count"]
                ),
                use_container_width=True
            )
            a1.write("代理维修量")

            agent_model = (
                df.groupby(["agent", "model"])
                .size()
                .reset_index(name="count")
                .sort_values(["agent", "count"], ascending=[True, False])
            )
            a2.dataframe(agent_model, use_container_width=True)

            st.subheader("💰 费用分析")
            c1, c2 = st.columns(2)

            repair_fee_by_country = (
                df.groupby("country", dropna=False)["repair_fee"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )
            c1.altair_chart(
                alt.Chart(repair_fee_by_country).mark_bar().encode(
                    x=alt.X("country:N", title="国家", sort="-y"),
                    y=alt.Y("repair_fee:Q", title="维修人工费"),
                    tooltip=["country", "repair_fee"]
                ),
                use_container_width=True
            )
            c1.write("维修人工费")

            shipping_fee_by_country = (
                df.groupby("country", dropna=False)["shipping_fee"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )
            c2.altair_chart(
                alt.Chart(shipping_fee_by_country).mark_bar().encode(
                    x=alt.X("country:N", title="国家", sort="-y"),
                    y=alt.Y("shipping_fee:Q", title="维修物流费"),
                    tooltip=["country", "shipping_fee"]
                ),
                use_container_width=True
            )
            c2.write("维修物流费")

            st.subheader("🔍 故障分析")
            c1, c2 = st.columns(2)

            issue_top = (
                df["issue_desc"]
                .value_counts()
                .head(10)
                .reset_index()
            )
            issue_top.columns = ["issue_desc", "count"]
            c1.altair_chart(
                alt.Chart(issue_top).mark_bar().encode(
                    x=alt.X("count:Q", title="数量"),
                    y=alt.Y("issue_desc:N", sort="-x", title="问题描述"),
                    tooltip=["issue_desc", "count"]
                ),
                use_container_width=True
            )
            c1.write("Top问题（基于 Problem description by AVONO）")

            tat_dist = (
                df["TAT"]
                .dropna()
                .astype(int)
                .value_counts()
                .sort_index()
                .reset_index()
            )
            if not tat_dist.empty:
                tat_dist.columns = ["TAT", "count"]
                c2.altair_chart(
                    alt.Chart(tat_dist).mark_bar().encode(
                        x=alt.X("TAT:O", title="TAT（工作日）"),
                        y=alt.Y("count:Q", title="数量"),
                        tooltip=["TAT", "count"]
                    ),
                    use_container_width=True
                )
                c2.write("TAT分布（工作日）")
            else:
                c2.info("无有效TAT数据")

            st.subheader("📦 Model分析")
            c1, c2 = st.columns(2)

            model_count = (
                df["model"]
                .value_counts()
                .head(10)
                .reset_index()
            )
            model_count.columns = ["model", "count"]
            c1.altair_chart(
                alt.Chart(model_count).mark_bar().encode(
                    x=alt.X("count:Q", title="维修数量"),
                    y=alt.Y("model:N", sort="-x", title="Model"),
                    tooltip=["model", "count"]
                ),
                use_container_width=True
            )
            c1.write("Model维修量 Top10")

            model_tat = (
                df.groupby("model", dropna=False)["TAT"]
                .mean()
                .dropna()
                .sort_values(ascending=False)
                .head(10)
                .reset_index()
            )
            if not model_tat.empty:
                model_tat.columns = ["model", "avg_tat"]
                c2.altair_chart(
                    alt.Chart(model_tat).mark_bar().encode(
                        x=alt.X("avg_tat:Q", title="平均TAT"),
                        y=alt.Y("model:N", sort="-x", title="Model"),
                        tooltip=["model", alt.Tooltip("avg_tat:Q", format=".2f")]
                    ),
                    use_container_width=True
                )
                c2.write("Model平均TAT Top10")
            else:
                c2.info("无有效Model TAT数据")

            st.subheader("🔧 更换SKU分析 Top10")
            sku_columns = find_replaced_sku_columns(df)

            if sku_columns:
                sku_long = (
                    df[["model", "agent"] + sku_columns]
                    .copy()
                    .assign(
                        model=lambda x: x["model"].fillna("未知").astype(str).str.strip().replace("", "未知"),
                        agent=lambda x: x["agent"].fillna("未知").astype(str).str.strip().replace("", "未知")
                    )
                    .melt(id_vars=["model", "agent"], value_vars=sku_columns, value_name="SKU")
                )

                sku_long["SKU"] = sku_long["SKU"].apply(normalize_sku_value)
                sku_long = sku_long.dropna(subset=["SKU"])

                if not sku_long.empty:
                    st.caption(f"已识别SKU字段：{', '.join(sku_columns)}")

                    sku_top = sku_long["SKU"].value_counts().head(10).reset_index()
                    sku_top.columns = ["SKU", "数量"]
                    sku_top = attach_sku_name(sku_top, sku_map_df)
                    sku_top = sku_top[["SKU", "中文名称", "数量"]]

                    c1, c2 = st.columns([1.2, 2])
                    c1.dataframe(sku_top, use_container_width=True)
                    c2.altair_chart(
                        alt.Chart(sku_top).mark_bar().encode(
                            x=alt.X("数量:Q", title="数量"),
                            y=alt.Y("SKU:N", sort="-x", title="SKU"),
                            tooltip=["SKU", "中文名称", "数量"]
                        ),
                        use_container_width=True
                    )

                    st.subheader("🔗 SKU + Model 关联分析")
                    model_sku_top = (
                        sku_long.groupby(["model", "SKU"])
                        .size()
                        .reset_index(name="数量")
                        .sort_values(["数量", "model", "SKU"], ascending=[False, True, True])
                    )

                    model_sku_top = attach_sku_name(model_sku_top, sku_map_df)
                    model_sku_top = model_sku_top[["model", "SKU", "中文名称", "数量"]]

                    st.write("Model 与 SKU 组合 Top20")
                    st.dataframe(model_sku_top.head(20), use_container_width=True)

                    pair_chart = alt.Chart(model_sku_top.head(20)).mark_bar().encode(
                        x=alt.X("数量:Q", title="更换数量"),
                        y=alt.Y("model:N", sort="-x", title="Model"),
                        color=alt.Color("SKU:N", title="SKU"),
                        tooltip=["model", "SKU", "中文名称", "数量"]
                    )
                    st.altair_chart(pair_chart, use_container_width=True)

                    st.subheader("🏷 代理 + SKU 分析")
                    agent_sku = (
                        sku_long.groupby(["agent", "SKU"])
                        .size()
                        .reset_index(name="数量")
                        .sort_values(["数量", "agent"], ascending=[False, True])
                    )
                    agent_sku = attach_sku_name(agent_sku, sku_map_df)
                    st.dataframe(agent_sku.head(30), use_container_width=True)

                    model_options = sorted(sku_long["model"].dropna().unique().tolist())
                    selected_model = st.selectbox("选择一个Model查看常更换SKU", options=model_options)

                    model_sku_detail = (
                        sku_long[sku_long["model"] == selected_model]["SKU"]
                        .value_counts()
                        .head(10)
                        .reset_index()
                    )
                    model_sku_detail.columns = ["SKU", "数量"]
                    model_sku_detail = attach_sku_name(model_sku_detail, sku_map_df)
                    model_sku_detail = model_sku_detail[["SKU", "中文名称", "数量"]]

                    c1, c2 = st.columns([1.2, 2])
                    c1.dataframe(model_sku_detail, use_container_width=True)
                    c2.altair_chart(
                        alt.Chart(model_sku_detail).mark_bar().encode(
                            x=alt.X("数量:Q", title="数量"),
                            y=alt.Y("SKU:N", sort="-x", title=f"{selected_model} 对应SKU"),
                            tooltip=["SKU", "中文名称", "数量"]
                        ),
                        use_container_width=True
                    )

                    with st.expander("查看SKU匹配调试信息"):
                        st.write("维修报告中的SKU示例：", sku_long["SKU"].dropna().unique()[:10])
                        if sku_map_df is not None:
                            st.write("对照表中的SKU示例：", sku_map_df["SKU"].dropna().unique()[:10])

                else:
                    st.info("已识别到 Replaced SKU 相关字段，但这些列中没有有效SKU数据")
            else:
                st.info("未识别到 Replaced SKU / Replaced SKU2 / Replaced SKU3 等字段，请检查表头是否在第一行")

            st.subheader("👨‍🔧 技术员表现")
            tech_perf = (
                df.groupby("technician", dropna=False)
                .agg(维修量=("sn", "count"), 平均TAT=("TAT", "mean"))
                .reset_index()
                .sort_values(by="维修量", ascending=False)
            )
            st.dataframe(tech_perf, use_container_width=True)

            st.subheader("📄 Repair Report 明细")
            st.dataframe(df, use_container_width=True)

    # =============================
    # TAB 3 - Additional Activity
    # =============================
    with tabs[2]:
        st.subheader("📦 Additional Activity 分析")

        if add_all_df.empty:
            st.info("未检测到 Additional Activity 数据。")
        else:
            st.write("原始数据")
            st.dataframe(add_all_df, use_container_width=True)

            if add_filtered_df.empty:
                st.warning("目标 Activity 类目下没有可分析数据。")
            else:
                st.subheader("🏷 按代理分析费用总额及占比")

                agent_activity = (
                    add_filtered_df.groupby(["agent", "activity_std"], dropna=False)["price"]
                    .sum()
                    .reset_index()
                )
                agent_totals = (
                    agent_activity.groupby("agent", dropna=False)["price"]
                    .sum()
                    .reset_index()
                    .rename(columns={"price": "agent_total"})
                )
                agent_activity = agent_activity.merge(agent_totals, on="agent", how="left")
                agent_activity["占比"] = np.where(
                    agent_activity["agent_total"] == 0,
                    0,
                    agent_activity["price"] / agent_activity["agent_total"]
                )

                show_df = agent_activity.copy()
                show_df = show_df.rename(columns={
                    "agent": "代理",
                    "activity_std": "Activity",
                    "price": "总费用",
                    "agent_total": "代理总费用"
                })
                st.dataframe(show_df, use_container_width=True)

                pivot_df = agent_activity.pivot_table(
                    index="activity_std",
                    columns="agent",
                    values="price",
                    aggfunc="sum",
                    fill_value=0
                ).reset_index()
                pivot_df = pivot_df.rename(columns={"activity_std": "Activity"})
                st.subheader("📊 Activity * 代理 费用透视表")
                st.dataframe(pivot_df, use_container_width=True)

                st.altair_chart(
                    alt.Chart(agent_activity).mark_bar().encode(
                        x=alt.X("agent:N", title="代理"),
                        y=alt.Y("price:Q", title="总费用"),
                        color=alt.Color("activity_std:N", title="Activity"),
                        tooltip=[
                            alt.Tooltip("agent:N", title="代理"),
                            alt.Tooltip("activity_std:N", title="Activity"),
                            alt.Tooltip("price:Q", title="总费用"),
                            alt.Tooltip("占比:Q", title="占比", format=".1%")
                        ]
                    ),
                    use_container_width=True
                )

            with st.expander("查看未识别的 Activity 原始值"):
                unmatched_activity = (
                    add_all_df[add_all_df["activity_std"].isna()]["activity"]
                    .value_counts()
                    .reset_index()
                )
                if not unmatched_activity.empty:
                    unmatched_activity.columns = ["未识别Activity", "出现次数"]
                    st.dataframe(unmatched_activity, use_container_width=True)
                else:
                    st.success("所有 Activity 都已成功识别。")

    # =============================
    # TAB 4 - OW
    # =============================
    with tabs[3]:
        st.subheader("💶 OW 分析")

        if ow_df.empty:
            st.info("未检测到 OW 数据。")
        else:
            ow_detail = attach_sku_name(ow_df.copy(), sku_map_df)
            ow_detail["备件金额"] = ow_detail["unit_price"]

            st.subheader("逐行明细")
            st.dataframe(
                ow_detail.rename(columns={
                    "order_id": "Order ID",
                    "model": "Model",
                    "sn": "SN",
                    "SKU": "Replaced SKU",
                    "中文名称": "中文名称",
                    "unit_price": "Precio BOOM Euros",
                    "备件金额": "金额"
                }),
                use_container_width=True
            )

            ow_grouped = (
                ow_detail.groupby(["order_id", "model", "sn"], dropna=False)
                .agg(
                    更换备件=("SKU", lambda x: " | ".join([str(v) for v in x.dropna().tolist()])),
                    中文名称=("中文名称", lambda x: " | ".join([str(v) for v in x.dropna().tolist()])),
                    总金额=("unit_price", "sum")
                )
                .reset_index()
                .sort_values("总金额", ascending=False)
            )

            st.subheader("按订单汇总（每个 Order ID 为一单）")
            st.dataframe(
                ow_grouped.rename(columns={
                    "order_id": "Order ID",
                    "model": "Model",
                    "sn": "SN"
                }),
                use_container_width=True
            )

            total_ow_amount = ow_detail["unit_price"].sum()
            st.metric("OW 总金额", f"{total_ow_amount:,.2f}")

    # =============================
    # TAB 5 - Storage
    # =============================
    with tabs[4]:
        st.subheader("📦 Storage")

        if overview_df.empty:
            st.info("未检测到 Übersicht 数据，因此无法提取 Storage。")
        else:
            storage_df = overview_df[
                overview_df["position"].astype(str).str.strip().str.lower() == "storage"
            ].copy()

            if storage_df.empty:
                st.info("Übersicht 中未找到 Storage 行。")
            else:
                storage_qty = storage_df["quantity"].sum()
                storage_price = storage_df["price"].sum()

                c1, c2 = st.columns(2)
                c1.metric("Quantity", f"{storage_qty:,.0f}")
                c2.metric("Overall Price", f"{storage_price:,.2f}")

                st.dataframe(
                    storage_df.rename(columns={
                        "position": "Position",
                        "quantity": "Quantity",
                        "price": "Price"
                    }),
                    use_container_width=True
                )

    # =============================
    # TAB 6 - DOA
    # =============================
    with tabs[5]:
        st.subheader("📄 DOA report")
        if doa_df.empty:
            st.info("未检测到 DOA report 数据。")
        else:
            st.info("按你的要求，DOA report 不做额外处理，仅展示原始数据。")
            st.dataframe(doa_df, use_container_width=True)

    # =============================
    # 导出
    # =============================
    st.subheader("⬇ 数据导出")

    export_files = {
        "Repair_Report_Filtered": repair_filtered,
        "Overview": overview_df,
        "Additional_Activity_Filtered": add_filtered_df,
        "OW": ow_df
    }

    st.download_button(
        "下载 Repair Report 清洗后CSV",
        repair_filtered.to_csv(index=False).encode("utf-8-sig"),
        file_name="repair_report_filtered.csv",
        mime="text/csv"
    )

    st.download_button(
        "下载多Sheet分析结果 Excel",
        to_excel_download(export_files),
        file_name="repair_dashboard_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("请上传维修报告 Excel 文件")
