import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import re
from io import BytesIO

st.set_page_config(page_title="统一维修分析平台", layout="wide")
st.title("🔧 统一维修分析平台 v1")


# =========================================================
# 基础工具函数
# =========================================================
def normalize_colname(col):
    return str(col).strip().lower()


def safe_columns(df):
    df = df.copy()
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


def normalize_for_match(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = re.sub(r"[_\-/\\]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s


def contains_keyword_fuzzy(value, keywords):
    text = normalize_for_match(value)
    if text == "":
        return None
    for k in keywords:
        k_norm = normalize_for_match(k)
        if k_norm and k_norm in text:
            return k.upper()
    return None


def normalize_sku_value(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if s.lower() in ["", "nan", "none", "null"]:
        return None
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def calc_tat(row, received_col="received_date", shipment_col="shipment_date"):
    try:
        if pd.isna(row[received_col]) or pd.isna(row[shipment_col]):
            return None
        if row[shipment_col] < row[received_col]:
            return None
        return np.busday_count(
            row[received_col].date(),
            row[shipment_col].date()
        )
    except Exception:
        return None


def safe_ratio(series):
    total = series.sum()
    if total == 0 or pd.isna(total):
        return series * 0
    return series / total


def find_column(df, candidates):
    cols = [normalize_colname(c) for c in df.columns]
    for cand in candidates:
        cand_n = normalize_colname(cand)
        for c in cols:
            if cand_n == c:
                return c
    for cand in candidates:
        cand_n = normalize_colname(cand)
        for c in cols:
            if cand_n in c:
                return c
    return None


def find_replaced_sku_columns(df):
    sku_cols = []
    for col in df.columns:
        col_str = str(col).strip().lower()
        if col_str.startswith("replaced sku"):
            suffix = col_str.replace("replaced sku", "").strip()
            if suffix == "" or suffix.isdigit():
                sku_cols.append(col)
    return sku_cols


def load_uploaded_file(uploaded_file):
    if uploaded_file.name.lower().endswith(".csv"):
        return {"single_csv": pd.read_csv(uploaded_file)}, ["single_csv"]
    xls = pd.ExcelFile(uploaded_file)
    sheets = {s: pd.read_excel(uploaded_file, sheet_name=s) for s in xls.sheet_names}
    return sheets, xls.sheet_names


def to_excel_download(data_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in data_dict.items():
            if df is not None and isinstance(df, pd.DataFrame):
                df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


# =========================================================
# SKU 映射
# =========================================================
def load_sku_mapping(uploaded_sku_file):
    if uploaded_sku_file is None:
        return None

    try:
        if uploaded_sku_file.name.lower().endswith(".csv"):
            sku_map_df = pd.read_csv(uploaded_sku_file)
        else:
            sku_map_df = pd.read_excel(uploaded_sku_file)

        sku_map_df.columns = sku_map_df.columns.astype(str).str.strip().str.lower()

        rename_map = {
            "sku": "SKU",
            "sku code": "SKU",
            "item sku": "SKU",
            "part sku": "SKU",
            "navee_code": "SKU",
            "中文名称": "中文名称",
            "名称": "中文名称",
            "name": "中文名称",
            "description": "中文名称",
            "desc": "中文名称",
            "part name": "中文名称",
        }
        sku_map_df = sku_map_df.rename(
            columns={k: v for k, v in rename_map.items() if k in sku_map_df.columns}
        )

        if "SKU" not in sku_map_df.columns:
            st.warning("SKU 对照表缺少 SKU 列，将忽略该对照表。")
            return None

        if "中文名称" not in sku_map_df.columns:
            sku_map_df["中文名称"] = "未提供"

        sku_map_df["SKU"] = sku_map_df["SKU"].apply(normalize_sku_value)
        sku_map_df["中文名称"] = sku_map_df["中文名称"].astype(str).str.strip()
        sku_map_df = sku_map_df.dropna(subset=["SKU"])
        sku_map_df = sku_map_df[["SKU", "中文名称"]].drop_duplicates(subset=["SKU"])
        return sku_map_df

    except Exception as e:
        st.error(f"SKU 对照表读取失败: {e}")
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


# =========================================================
# 模板识别
# =========================================================
def detect_template(sheets, sheet_names):
    normalized_sheet_names = [s.strip().lower() for s in sheet_names]

    # AVONO multisheet
    avono_keys = {"repair report", "additional activity", "ow"}
    if avono_keys.issubset(set(normalized_sheet_names)):
        return "avono_multisheet"

    # NAVEE service report
    navee_keys = {"pcs", "labor", "spareparts"}
    if navee_keys.issubset(set(normalized_sheet_names)):
        return "navee_service_report"

    # generic single/other
    return "generic_report"


# =========================================================
# AVONO 解析
# =========================================================
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
    client_keywords = ["CESV", "Carrefour", "Feuvert", "Conforama"]
    client_hit = contains_keyword_fuzzy(client, client_keywords)
    if client_hit:
        return client_hit
    return "MCR"


def normalize_activity_name(activity):
    text = normalize_for_match(activity)
    if text == "":
        return None

    if any(x in text for x in ["comunication", "communication", "comunications", "communications"]):
        return "COMUNICATION"

    if "call" in text or "calls" in text:
        return "CALLS"

    if "box" in text or "boxes" in text:
        return "BOXES"

    if ("worten" in text and "reception" in text) or "worten reception" in text:
        return "WORTEN RECEPTION"

    if ("battery" in text and "voltage" in text and "check" in text) or "battery voltage" in text:
        return "BATTERY VOLTAGE CHECK"

    if ("doa" in text and "management" in text) or "doas management" in text or "doa management" in text:
        return "DOAS Management"

    if ("anovo" in text and "stock" in text and "transfer" in text) or "stock transfer" in text:
        return "ANOVO Stock Transfer"

    return None


def parse_avono_overview(df_raw):
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
    for c in ["position", "quantity", "price"]:
        df = ensure_column(df, c)

    df["position"] = normalize_text_series(df["position"], default="")
    df = df[df["position"] != ""].copy()
    df["quantity"] = to_numeric_safe(df["quantity"], default=0)
    df["price"] = to_numeric_safe(df["price"], default=0)
    return df[["position", "quantity", "price"]]


def parse_avono_repair_report(df_raw):
    df = safe_columns(df_raw.copy())

    rename_map = {
        "repair order no.": "repair_id",
        "repair order": "repair_id",
        "repair id": "repair_id",
        "date of receipt": "received_date",
        "date of shipment": "shipment_date",
        "nation /state": "country",
        "nation/state": "country",
        "nation": "country",
        "sales channal": "sales_channel",
        "sales channel": "sales_channel",
        "customer name": "customer_name",
        "order id": "order_id",
        "model": "model",
        "warranty status": "repair_type",
        "problem description by customer": "customer_issue",
        "problem description by avono": "issue_desc",
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

    type_map = {"iw": "保内", "ow": "保外", "doa": "DOA"}
    df["repair_type"] = (
        df["repair_type"].astype(str).str.strip().str.lower().map(type_map).fillna(df["repair_type"])
    )

    for col in ["repair_fee", "shipping_fee", "resend_shipping_fee", "scrap_fee", "other_fee"]:
        df[col] = to_numeric_safe(df[col], default=0)

    for col in ["country", "sales_channel", "customer_name", "model", "issue_desc", "technician"]:
        df[col] = normalize_text_series(df[col])

    df["sn"] = df["sn"].fillna("").astype(str).str.strip()
    df["total_cost"] = df["repair_fee"] + df["shipping_fee"] + df["resend_shipping_fee"] + df["scrap_fee"] + df["other_fee"]
    df["TAT"] = df.apply(lambda r: calc_tat(r, "received_date", "shipment_date"), axis=1)
    df = df.sort_values(by=["sn", "received_date"], na_position="last")
    df["repeat"] = df.duplicated(subset=["sn"], keep="first")
    df["agent"] = df.apply(lambda r: map_agent_repair_report(r.get("sales_channel"), r.get("model")), axis=1)

    return df


def parse_avono_additional_activity(df_raw):
    df = safe_columns(df_raw.copy())
    for col in ["activity", "client", "price"]:
        df = ensure_column(df, col)

    df["activity"] = normalize_text_series(df["activity"])
    df["client"] = normalize_text_series(df["client"])
    df["price"] = to_numeric_safe(df["price"], default=0)
    df["agent"] = df["client"].apply(map_agent_additional_activity)
    df["activity_std"] = df["activity"].apply(normalize_activity_name)

    filtered = df[df["activity_std"].notna()].copy()
    return df, filtered


def parse_avono_ow(df_raw):
    df = safe_columns(df_raw.copy())
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


def parse_avono_template(sheets, sheet_names):
    sheet_map = {s.strip().lower(): s for s in sheet_names}

    overview_raw = sheets.get(sheet_map.get("übersicht")) or sheets.get(sheet_map.get("ubersicht"))
    repair_raw = sheets.get(sheet_map.get("repair report"))
    add_raw = sheets.get(sheet_map.get("additional activity"))
    ow_raw = sheets.get(sheet_map.get("ow"))
    doa_raw = None
    for s in sheet_names:
        if "doa report" in s.strip().lower():
            doa_raw = sheets[s]
            break

    overview_df = parse_avono_overview(overview_raw) if overview_raw is not None else pd.DataFrame()
    repairs_df = parse_avono_repair_report(repair_raw) if repair_raw is not None else pd.DataFrame()
    activity_all_df, activity_df = parse_avono_additional_activity(add_raw) if add_raw is not None else (pd.DataFrame(), pd.DataFrame())
    ow_df = parse_avono_ow(ow_raw) if ow_raw is not None else pd.DataFrame()
    doa_df = doa_raw.copy() if doa_raw is not None else pd.DataFrame()

    parts_df = pd.DataFrame()
    if not repairs_df.empty:
        sku_columns = find_replaced_sku_columns(repairs_df)
        if sku_columns:
            parts_df = repairs_df[["repair_id", "order_id", "model", "sn", "agent"] + sku_columns].copy()
            parts_df = parts_df.melt(
                id_vars=["repair_id", "order_id", "model", "sn", "agent"],
                value_vars=sku_columns,
                value_name="SKU"
            )
            parts_df["SKU"] = parts_df["SKU"].apply(normalize_sku_value)
            parts_df = parts_df.dropna(subset=["SKU"])
            parts_df["qty"] = 1

    meta_info = {
        "template_name": "avono_multisheet",
        "sheet_names": sheet_names,
        "notes": []
    }

    return {
        "template_name": "avono_multisheet",
        "repairs_df": repairs_df,
        "parts_df": parts_df,
        "activity_df": activity_df,
        "activity_all_df": activity_all_df,
        "finance_df": overview_df,
        "orders_df": ow_df,
        "overview_df": overview_df,
        "doa_df": doa_df,
        "meta_info": meta_info
    }


# =========================================================
# NAVEE 解析
# =========================================================
def parse_navee_pcs(df_raw):
    df = safe_columns(df_raw.copy())
    rename_map = {
        "reg_date": "reg_date",
        "rma_nr": "repair_id",
        "rma_sub_nr": "repair_sub_id",
        "equipment": "model",
        "sn": "sn",
        "battery_sn": "battery_sn",
        "awb_reception": "awb_reception",
        "date_of_reception": "received_date",
        "defect": "issue_desc",
        "state": "state",
        "ready_for_shipment": "ready_for_shipment",
        "date_of_delivery": "shipment_date",
        "awb_delivery": "awb_delivery",
        "defect_eng": "issue_desc_en",
        "defect_details": "issue_detail",
        "date_of_purchase": "purchase_date",
        "purchase_doc": "purchase_doc",
        "obs.": "remark",
        "partener": "partner",
        "pf": "pf"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    for col in [
        "repair_id", "repair_sub_id", "model", "sn", "battery_sn", "received_date",
        "shipment_date", "issue_desc", "issue_desc_en", "issue_detail", "partner",
        "state", "pf"
    ]:
        df = ensure_column(df, col)

    df["received_date"] = pd.to_datetime(df["received_date"], errors="coerce")
    df["shipment_date"] = pd.to_datetime(df["shipment_date"], errors="coerce")
    df["purchase_date"] = pd.to_datetime(df.get("purchase_date"), errors="coerce")

    for col in ["model", "sn", "battery_sn", "issue_desc", "issue_desc_en", "issue_detail", "partner", "state"]:
        df[col] = normalize_text_series(df[col])

    df["repair_type"] = np.where(df["pf"].astype(str).str.strip() == "1", "保内", "保外")
    df["country"] = "未知"
    df["customer_name"] = df["partner"]
    df["sales_channel"] = df["partner"]
    df["agent"] = df["partner"]
    df["technician"] = "未知"
    df["total_cost"] = 0
    df["TAT"] = df.apply(lambda r: calc_tat(r, "received_date", "shipment_date"), axis=1)
    df = df.sort_values(by=["sn", "received_date"], na_position="last")
    df["repeat"] = df.duplicated(subset=["sn"], keep="first")

    # 同一 RMA 可能多条 defect_eng，汇总成 repair 主表
    group_cols = [
        "repair_id", "repair_sub_id", "model", "sn", "battery_sn", "received_date", "shipment_date",
        "purchase_date", "partner", "state", "repair_type", "country", "customer_name",
        "sales_channel", "agent", "technician", "pf", "TAT", "repeat"
    ]
    agg_df = df.groupby(group_cols, dropna=False).agg(
        issue_desc=("issue_desc", lambda x: " | ".join(pd.Series(x).dropna().astype(str).unique())),
        issue_desc_en=("issue_desc_en", lambda x: " | ".join(pd.Series(x).dropna().astype(str).unique())),
        issue_detail=("issue_detail", lambda x: " | ".join(pd.Series(x).dropna().astype(str).unique())),
        total_cost=("total_cost", "sum")
    ).reset_index()

    return agg_df, df


def parse_navee_labor(df_raw):
    df = safe_columns(df_raw.copy())
    rename_map = {
        "reg_date": "reg_date",
        "rma_nr": "repair_id",
        "rma_sub_nr": "repair_sub_id",
        "equipment": "model",
        "sn": "sn",
        "date_of_reception": "received_date",
        "ready_for_shipment": "ready_for_shipment",
        "date_of_delivery": "shipment_date",
        "defect": "issue_desc",
        "labor_eng": "labor_name",
        "level": "labor_level",
        "price": "labor_price",
        "labor details": "labor_detail",
        "state": "state"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    for col in [
        "repair_id", "repair_sub_id", "model", "sn", "received_date", "shipment_date",
        "issue_desc", "labor_name", "labor_level", "labor_price", "labor_detail", "state"
    ]:
        df = ensure_column(df, col)

    df["received_date"] = pd.to_datetime(df["received_date"], errors="coerce")
    df["shipment_date"] = pd.to_datetime(df["shipment_date"], errors="coerce")
    for col in ["model", "sn", "issue_desc", "labor_name", "labor_level", "labor_detail", "state"]:
        df[col] = normalize_text_series(df[col])
    df["labor_price"] = to_numeric_safe(df["labor_price"], default=0)
    return df


def parse_navee_spareparts(df_raw):
    df = safe_columns(df_raw.copy())
    rename_map = {
        "reg_date": "reg_date",
        "rma_nr": "repair_id",
        "rma_sub_nr": "repair_sub_id",
        "sparepart": "part_name",
        "qty": "qty",
        "um": "um",
        "navee_code": "SKU",
        "state": "state"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    for col in ["repair_id", "repair_sub_id", "part_name", "qty", "um", "SKU", "state"]:
        df = ensure_column(df, col)

    df["part_name"] = normalize_text_series(df["part_name"])
    df["qty"] = to_numeric_safe(df["qty"], default=0)
    df["SKU"] = df["SKU"].apply(normalize_sku_value)
    df["state"] = normalize_text_series(df["state"])
    return df


def parse_navee_template(sheets, sheet_names):
    pcs_raw = sheets[[s for s in sheet_names if s.strip().lower() == "pcs"][0]]
    labor_raw = sheets[[s for s in sheet_names if s.strip().lower() == "labor"][0]]
    parts_raw = sheets[[s for s in sheet_names if s.strip().lower() == "spareparts"][0]]

    repairs_df, pcs_detail_df = parse_navee_pcs(pcs_raw)
    labor_df = parse_navee_labor(labor_raw)
    parts_df = parse_navee_spareparts(parts_raw)

    # labor 汇总进 repairs
    labor_summary = labor_df.groupby(["repair_id", "repair_sub_id"], dropna=False).agg(
        labor_cost=("labor_price", "sum"),
        labor_names=("labor_name", lambda x: " | ".join(pd.Series(x).dropna().astype(str).unique()))
    ).reset_index()

    repairs_df = repairs_df.merge(labor_summary, on=["repair_id", "repair_sub_id"], how="left")
    repairs_df["labor_cost"] = repairs_df["labor_cost"].fillna(0)
    repairs_df["total_cost"] = repairs_df["labor_cost"]

    # parts 关联 model/sn/agent
    parts_df = parts_df.merge(
        repairs_df[["repair_id", "repair_sub_id", "model", "sn", "agent"]],
        on=["repair_id", "repair_sub_id"],
        how="left"
    )

    # orders
    orders_df = repairs_df[[
        "repair_id", "repair_sub_id", "model", "sn", "agent",
        "received_date", "shipment_date", "repair_type", "total_cost"
    ]].copy()

    finance_df = pd.DataFrame({
        "position": ["Labor"],
        "quantity": [len(labor_df)],
        "price": [labor_df["labor_price"].sum()]
    })

    meta_info = {
        "template_name": "navee_service_report",
        "sheet_names": sheet_names,
        "notes": []
    }

    return {
        "template_name": "navee_service_report",
        "repairs_df": repairs_df,
        "parts_df": parts_df,
        "activity_df": pd.DataFrame(),
        "activity_all_df": pd.DataFrame(),
        "finance_df": finance_df,
        "orders_df": orders_df,
        "overview_df": finance_df,
        "doa_df": pd.DataFrame(),
        "labor_df": labor_df,
        "pcs_detail_df": pcs_detail_df,
        "meta_info": meta_info
    }


# =========================================================
# 通用模板解析
# =========================================================
def parse_generic_template(sheets, sheet_names):
    first_sheet_name = sheet_names[0]
    df = safe_columns(sheets[first_sheet_name].copy())

    model_col = find_column(df, ["model", "equipment"])
    sn_col = find_column(df, ["sn", "serial"])
    recv_col = find_column(df, ["date_of_reception", "receipt", "received"])
    ship_col = find_column(df, ["date_of_delivery", "shipment", "ship", "delivery"])
    issue_col = find_column(df, ["problem", "issue", "defect"])
    warranty_col = find_column(df, ["warranty", "pf"])
    client_col = find_column(df, ["client", "customer", "partner", "partener"])
    price_col = find_column(df, ["price", "fee", "cost"])
    order_col = find_column(df, ["order", "repair", "rma"])

    repairs_df = pd.DataFrame()
    repairs_df["repair_id"] = df[order_col] if order_col else np.arange(1, len(df) + 1)
    repairs_df["model"] = normalize_text_series(df[model_col]) if model_col else "未知"
    repairs_df["sn"] = normalize_text_series(df[sn_col]) if sn_col else "未知"
    repairs_df["received_date"] = pd.to_datetime(df[recv_col], errors="coerce") if recv_col else pd.NaT
    repairs_df["shipment_date"] = pd.to_datetime(df[ship_col], errors="coerce") if ship_col else pd.NaT
    repairs_df["issue_desc"] = normalize_text_series(df[issue_col]) if issue_col else "未知"
    repairs_df["customer_name"] = normalize_text_series(df[client_col]) if client_col else "未知"
    repairs_df["sales_channel"] = repairs_df["customer_name"]
    repairs_df["country"] = "未知"
    repairs_df["agent"] = repairs_df["customer_name"]
    repairs_df["technician"] = "未知"

    if warranty_col:
        wt = normalize_text_series(df[warranty_col], default="未知").astype(str).str.lower()
        repairs_df["repair_type"] = np.where(
            wt.str.contains("iw"), "保内",
            np.where(wt.str.contains("ow"), "保外", np.where(wt.str.contains("doa"), "DOA", "未知"))
        )
    else:
        repairs_df["repair_type"] = "未知"

    repairs_df["total_cost"] = to_numeric_safe(df[price_col], default=0) if price_col else 0
    repairs_df["TAT"] = repairs_df.apply(lambda r: calc_tat(r, "received_date", "shipment_date"), axis=1)
    repairs_df = repairs_df.sort_values(by=["sn", "received_date"], na_position="last")
    repairs_df["repeat"] = repairs_df.duplicated(subset=["sn"], keep="first")

    # parts
    sku_cols = [c for c in df.columns if "sku" in c]
    parts_df = pd.DataFrame()
    if sku_cols:
        tmp = df[[order_col] + sku_cols].copy() if order_col else df[sku_cols].copy()
        if order_col:
            tmp.columns = ["repair_id"] + sku_cols
        else:
            tmp.insert(0, "repair_id", np.arange(1, len(tmp) + 1))

        parts_df = tmp.melt(id_vars=["repair_id"], value_vars=sku_cols, value_name="SKU")
        parts_df["SKU"] = parts_df["SKU"].apply(normalize_sku_value)
        parts_df = parts_df.dropna(subset=["SKU"])
        parts_df["qty"] = 1
        parts_df = parts_df.merge(repairs_df[["repair_id", "model", "sn", "agent"]], on="repair_id", how="left")

    finance_df = pd.DataFrame({
        "position": ["Total"],
        "quantity": [len(repairs_df)],
        "price": [repairs_df["total_cost"].sum()]
    })

    meta_info = {
        "template_name": "generic_report",
        "sheet_names": sheet_names,
        "notes": [f"按通用模板解析：{first_sheet_name}"]
    }

    return {
        "template_name": "generic_report",
        "repairs_df": repairs_df,
        "parts_df": parts_df,
        "activity_df": pd.DataFrame(),
        "activity_all_df": pd.DataFrame(),
        "finance_df": finance_df,
        "orders_df": repairs_df.copy(),
        "overview_df": finance_df,
        "doa_df": pd.DataFrame(),
        "meta_info": meta_info
    }


# =========================================================
# 统一解析入口
# =========================================================
def parse_report(sheets, sheet_names):
    template_name = detect_template(sheets, sheet_names)
    if template_name == "avono_multisheet":
        return parse_avono_template(sheets, sheet_names)
    if template_name == "navee_service_report":
        return parse_navee_template(sheets, sheet_names)
    return parse_generic_template(sheets, sheet_names)


# =========================================================
# 展示模块
# =========================================================
def render_header_info(parsed):
    st.subheader("🧭 识别结果")
    c1, c2 = st.columns([1, 2])
    c1.metric("识别模板", parsed["template_name"])
    c2.write("检测到的 Sheets：", " / ".join(parsed["meta_info"].get("sheet_names", [])))
    notes = parsed["meta_info"].get("notes", [])
    if notes:
        for n in notes:
            st.caption(f"- {n}")


def render_sidebar_filters(repairs_df):
    st.sidebar.header("筛选条件")

    df = repairs_df.copy()
    if df.empty:
        return df

    country_vals = sorted(df["country"].dropna().astype(str).unique().tolist()) if "country" in df.columns else []
    type_vals = sorted(df["repair_type"].dropna().astype(str).unique().tolist()) if "repair_type" in df.columns else []
    model_vals = sorted(df["model"].dropna().astype(str).unique().tolist()) if "model" in df.columns else []
    agent_vals = sorted(df["agent"].dropna().astype(str).unique().tolist()) if "agent" in df.columns else []

    country_filter = st.sidebar.multiselect("国家", country_vals, default=country_vals)
    type_filter = st.sidebar.multiselect("维修类型", type_vals, default=type_vals)
    model_filter = st.sidebar.multiselect("Model", model_vals, default=model_vals)
    agent_filter = st.sidebar.multiselect("代理/客户", agent_vals, default=agent_vals)

    min_date = df["received_date"].min() if "received_date" in df.columns and df["received_date"].notna().any() else pd.Timestamp.today()
    max_date = df["received_date"].max() if "received_date" in df.columns and df["received_date"].notna().any() else pd.Timestamp.today()
    date_range = st.sidebar.date_input("日期范围", [min_date, max_date])

    if country_vals:
        df = df[df["country"].isin(country_filter)]
    if type_vals:
        df = df[df["repair_type"].isin(type_filter)]
    if model_vals:
        df = df[df["model"].isin(model_filter)]
    if agent_vals:
        df = df[df["agent"].isin(agent_filter)]
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2 and "received_date" in df.columns:
        df = df[
            (df["received_date"] >= pd.to_datetime(date_range[0])) &
            (df["received_date"] <= pd.to_datetime(date_range[1]))
        ]

    return df


def render_overview_tab(parsed, repairs_df_filtered):
    st.subheader("📊 总览")

    repairs_df = repairs_df_filtered.copy()
    finance_df = parsed.get("finance_df", pd.DataFrame()).copy()
    overview_df = parsed.get("overview_df", pd.DataFrame()).copy()

    total_repairs = len(repairs_df)
    avg_tat = repairs_df["TAT"].dropna().mean() if "TAT" in repairs_df.columns and not repairs_df.empty else 0
    total_cost = repairs_df["total_cost"].sum() if "total_cost" in repairs_df.columns and not repairs_df.empty else 0
    repeat_rate = repairs_df["repeat"].mean() if "repeat" in repairs_df.columns and not repairs_df.empty else 0
    doa_rate = (repairs_df["repair_type"] == "DOA").mean() if "repair_type" in repairs_df.columns and not repairs_df.empty else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("总维修量", f"{total_repairs:,}")
    c2.metric("平均TAT", round(avg_tat, 1) if not pd.isna(avg_tat) else 0)
    c3.metric("总费用", f"{total_cost:,.2f}")
    c4.metric("重复维修率", f"{repeat_rate:.1%}")
    c5.metric("DOA占比", f"{doa_rate:.1%}")

    if parsed["template_name"] == "avono_multisheet" and not overview_df.empty:
        def get_price(keyword):
            mask = overview_df["position"].astype(str).str.strip().str.lower() == keyword.lower()
            return overview_df.loc[mask, "price"].sum() if mask.any() else 0

        total_price = get_price("Total")
        ow_parts_price = get_price("OW Parts")
        total_invoice = total_price - ow_parts_price

        s1, s2, s3 = st.columns(3)
        s1.metric("Overview Total", f"{total_price:,.2f}")
        s2.metric("OW Parts", f"{ow_parts_price:,.2f}")
        s3.metric("Total Invoice", f"{total_invoice:,.2f}")

    st.subheader("📈 趋势")
    if "received_date" in repairs_df.columns and repairs_df["received_date"].notna().any():
        tmp = repairs_df.copy()
        tmp["month"] = tmp["received_date"].dt.to_period("M").astype(str)
        trend = tmp.groupby("month").size().reset_index(name="count")
        st.altair_chart(
            alt.Chart(trend).mark_line(point=True).encode(
                x=alt.X("month:N", title="月份", sort=None),
                y=alt.Y("count:Q", title="维修数量"),
                tooltip=["month", "count"]
            ),
            use_container_width=True
        )
    else:
        st.info("无有效日期数据，无法生成趋势图。")

    if not finance_df.empty and {"position", "price"}.issubset(finance_df.columns):
        st.subheader("💰 汇总费用")
        st.dataframe(finance_df, use_container_width=True)


def render_repairs_tab(repairs_df_filtered):
    st.subheader("🛠 维修分析")
    df = repairs_df_filtered.copy()

    if df.empty:
        st.warning("当前筛选条件下没有数据。")
        return

    st.subheader("📦 结构分析")
    c1, c2, c3 = st.columns(3)

    if "repair_type" in df.columns:
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

    if "country" in df.columns:
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

        if "total_cost" in df.columns:
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

    st.subheader("🏷 代理 / 客户分析")
    if "agent" in df.columns:
        agent_count = df["agent"].value_counts().reset_index()
        agent_count.columns = ["agent", "count"]
        a1, a2 = st.columns(2)
        a1.altair_chart(
            alt.Chart(agent_count).mark_bar().encode(
                x=alt.X("agent:N", title="代理/客户"),
                y=alt.Y("count:Q", title="维修量"),
                tooltip=["agent", "count"]
            ),
            use_container_width=True
        )

        agent_model = (
            df.groupby(["agent", "model"], dropna=False)
            .size()
            .reset_index(name="count")
            .sort_values(["agent", "count"], ascending=[True, False])
        )
        a2.dataframe(agent_model.head(50), use_container_width=True)

    st.subheader("📦 Model 分析")
    m1, m2 = st.columns(2)

    if "model" in df.columns:
        model_count = df["model"].value_counts().head(10).reset_index()
        model_count.columns = ["model", "count"]
        m1.altair_chart(
            alt.Chart(model_count).mark_bar().encode(
                x=alt.X("count:Q", title="维修数量"),
                y=alt.Y("model:N", sort="-x", title="Model"),
                tooltip=["model", "count"]
            ),
            use_container_width=True
        )

        if "TAT" in df.columns:
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
                m2.altair_chart(
                    alt.Chart(model_tat).mark_bar().encode(
                        x=alt.X("avg_tat:Q", title="平均TAT"),
                        y=alt.Y("model:N", sort="-x", title="Model"),
                        tooltip=["model", alt.Tooltip("avg_tat:Q", format=".2f")]
                    ),
                    use_container_width=True
                )

    st.subheader("🔍 故障分析")
    i1, i2 = st.columns(2)
    issue_col = "issue_desc" if "issue_desc" in df.columns else None
    if issue_col:
        issue_top = df[issue_col].value_counts().head(10).reset_index()
        issue_top.columns = ["issue_desc", "count"]
        i1.altair_chart(
            alt.Chart(issue_top).mark_bar().encode(
                x=alt.X("count:Q", title="数量"),
                y=alt.Y("issue_desc:N", sort="-x", title="问题描述"),
                tooltip=["issue_desc", "count"]
            ),
            use_container_width=True
        )

    if "TAT" in df.columns:
        tat_dist = df["TAT"].dropna()
        if not tat_dist.empty:
            tat_dist = tat_dist.astype(int).value_counts().sort_index().reset_index()
            tat_dist.columns = ["TAT", "count"]
            i2.altair_chart(
                alt.Chart(tat_dist).mark_bar().encode(
                    x=alt.X("TAT:O", title="TAT（工作日）"),
                    y=alt.Y("count:Q", title="数量"),
                    tooltip=["TAT", "count"]
                ),
                use_container_width=True
            )

    st.subheader("📄 维修主表明细")
    st.dataframe(df, use_container_width=True)


def render_parts_tab(parts_df, sku_map_df):
    st.subheader("🔧 备件 / SKU 分析")
    df = parts_df.copy()

    if df.empty:
        st.info("当前报告未识别到备件 / SKU 数据。")
        return

    if "qty" not in df.columns:
        df["qty"] = 1

    sku_top = (
        df.groupby("SKU", dropna=False)["qty"]
        .sum()
        .reset_index()
        .sort_values("qty", ascending=False)
        .head(10)
        .rename(columns={"qty": "数量"})
    )
    sku_top = attach_sku_name(sku_top, sku_map_df)

    c1, c2 = st.columns([1.2, 2])
    c1.dataframe(sku_top, use_container_width=True)
    c2.altair_chart(
        alt.Chart(sku_top).mark_bar().encode(
            x=alt.X("数量:Q", title="数量"),
            y=alt.Y("SKU:N", sort="-x", title="SKU"),
            tooltip=[c for c in sku_top.columns]
        ),
        use_container_width=True
    )

    if {"model", "SKU", "qty"}.issubset(df.columns):
        st.subheader("🔗 Model + SKU")
        model_sku = (
            df.groupby(["model", "SKU"], dropna=False)["qty"]
            .sum()
            .reset_index()
            .sort_values("qty", ascending=False)
            .head(30)
            .rename(columns={"qty": "数量"})
        )
        model_sku = attach_sku_name(model_sku, sku_map_df)
        st.dataframe(model_sku, use_container_width=True)

    if {"agent", "SKU", "qty"}.issubset(df.columns):
        st.subheader("🏷 代理/客户 + SKU")
        agent_sku = (
            df.groupby(["agent", "SKU"], dropna=False)["qty"]
            .sum()
            .reset_index()
            .sort_values("qty", ascending=False)
            .head(30)
            .rename(columns={"qty": "数量"})
        )
        agent_sku = attach_sku_name(agent_sku, sku_map_df)
        st.dataframe(agent_sku, use_container_width=True)

    st.subheader("📄 备件明细")
    detail = attach_sku_name(df, sku_map_df)
    st.dataframe(detail, use_container_width=True)


def render_activity_tab(parsed):
    st.subheader("📦 Additional Activity / 附加活动分析")
    activity_all_df = parsed.get("activity_all_df", pd.DataFrame()).copy()
    activity_df = parsed.get("activity_df", pd.DataFrame()).copy()

    if activity_all_df.empty and activity_df.empty:
        st.info("当前报告未识别到 Additional Activity 数据。")
        return

    st.write("原始数据")
    st.dataframe(activity_all_df, use_container_width=True)

    if not activity_df.empty and {"agent", "activity_std", "price"}.issubset(activity_df.columns):
        st.subheader("🏷 按代理分析费用总额及占比")

        agent_activity = (
            activity_df.groupby(["agent", "activity_std"], dropna=False)["price"]
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
            agent_activity["agent_total"] == 0, 0, agent_activity["price"] / agent_activity["agent_total"]
        )

        show_df = agent_activity.rename(columns={
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
        ).reset_index().rename(columns={"activity_std": "Activity"})
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
            unmatched = (
                activity_all_df[activity_all_df["activity_std"].isna()]["activity"]
                .value_counts()
                .reset_index()
            )
            if not unmatched.empty:
                unmatched.columns = ["未识别Activity", "出现次数"]
                st.dataframe(unmatched, use_container_width=True)
            else:
                st.success("所有 Activity 都已成功识别。")


def render_finance_tab(repairs_df_filtered, parsed):
    st.subheader("💰 费用分析")
    df = repairs_df_filtered.copy()

    if df.empty:
        st.info("无可分析数据。")
        return

    if "total_cost" in df.columns:
        c1, c2 = st.columns(2)

        if "country" in df.columns:
            by_country = (
                df.groupby("country", dropna=False)["total_cost"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )
            c1.altair_chart(
                alt.Chart(by_country).mark_bar().encode(
                    x=alt.X("country:N", title="国家", sort="-y"),
                    y=alt.Y("total_cost:Q", title="总费用"),
                    tooltip=["country", "total_cost"]
                ),
                use_container_width=True
            )

        if "agent" in df.columns:
            by_agent = (
                df.groupby("agent", dropna=False)["total_cost"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )
            c2.altair_chart(
                alt.Chart(by_agent).mark_bar().encode(
                    x=alt.X("agent:N", title="代理/客户", sort="-y"),
                    y=alt.Y("total_cost:Q", title="总费用"),
                    tooltip=["agent", "total_cost"]
                ),
                use_container_width=True
            )

    if parsed["template_name"] == "navee_service_report" and "labor_df" in parsed:
        labor_df = parsed["labor_df"].copy()
        if not labor_df.empty:
            st.subheader("🧰 Labor 分析")
            labor_summary = (
                labor_df.groupby("labor_name", dropna=False)["labor_price"]
                .sum()
                .reset_index()
                .sort_values("labor_price", ascending=False)
                .head(20)
            )
            st.dataframe(labor_summary, use_container_width=True)


def render_orders_tab(parsed, sku_map_df):
    st.subheader("📦 订单 / 工单明细")
    orders_df = parsed.get("orders_df", pd.DataFrame()).copy()

    if orders_df.empty:
        st.info("当前报告未识别到订单级明细。")
        return

    if parsed["template_name"] == "avono_multisheet":
        ow_df = orders_df.copy()
        if not ow_df.empty:
            ow_detail = attach_sku_name(ow_df.copy(), sku_map_df)
            ow_detail["备件金额"] = ow_detail["unit_price"] if "unit_price" in ow_detail.columns else 0

            st.subheader("OW 逐行明细")
            st.dataframe(ow_detail, use_container_width=True)

            if {"order_id", "model", "sn", "SKU", "unit_price"}.issubset(ow_detail.columns):
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
                st.dataframe(ow_grouped, use_container_width=True)

    else:
        st.dataframe(orders_df, use_container_width=True)


def render_quality_tab(parsed, repairs_df_filtered, parts_df):
    st.subheader("🧪 数据质量检查")
    df = repairs_df_filtered.copy()

    q = {
        "总记录数": len(df),
        "缺失 Model": int((df["model"].astype(str).str.strip().isin(["", "未知"])).sum()) if "model" in df.columns else 0,
        "缺失 SN": int((df["sn"].astype(str).str.strip().isin(["", "未知"])).sum()) if "sn" in df.columns else 0,
        "缺失 收货日期": int(df["received_date"].isna().sum()) if "received_date" in df.columns else 0,
        "缺失 发货日期": int(df["shipment_date"].isna().sum()) if "shipment_date" in df.columns else 0,
        "异常 TAT(负数)": int((df["TAT"] < 0).sum()) if "TAT" in df.columns and not df.empty else 0,
    }
    st.dataframe(pd.DataFrame(list(q.items()), columns=["检查项", "数量"]), use_container_width=True)

    if not parts_df.empty:
        with st.expander("查看缺失 SKU 的明细"):
            bad_sku = parts_df[parts_df["SKU"].isna()] if "SKU" in parts_df.columns else pd.DataFrame()
            st.dataframe(bad_sku, use_container_width=True)

    activity_all_df = parsed.get("activity_all_df", pd.DataFrame())
    if not activity_all_df.empty and "activity_std" in activity_all_df.columns:
        with st.expander("查看未识别 Activity 明细"):
            unmatched = activity_all_df[activity_all_df["activity_std"].isna()]
            st.dataframe(unmatched, use_container_width=True)


def render_export_tab(parsed, repairs_df_filtered, parts_df):
    st.subheader("⬇ 数据导出")

    export_files = {
        "repairs_filtered": repairs_df_filtered,
        "parts": parts_df,
        "activity": parsed.get("activity_df", pd.DataFrame()),
        "finance": parsed.get("finance_df", pd.DataFrame()),
        "orders": parsed.get("orders_df", pd.DataFrame()),
        "overview": parsed.get("overview_df", pd.DataFrame()),
    }

    st.download_button(
        "下载维修主表 CSV",
        repairs_df_filtered.to_csv(index=False).encode("utf-8-sig"),
        file_name="repairs_filtered.csv",
        mime="text/csv"
    )

    st.download_button(
        "下载统一分析结果 Excel",
        to_excel_download(export_files),
        file_name="repair_platform_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================================================
# 主程序
# =========================================================
report_file = st.file_uploader(
    "上传维修报告（支持 Excel / CSV）",
    type=["xlsx", "csv"],
    key="report_file"
)
sku_file = st.file_uploader(
    "上传 SKU 对照表（可选，Excel / CSV）",
    type=["xlsx", "csv"],
    key="sku_file"
)

if report_file:
    try:
        sheets, sheet_names = load_uploaded_file(report_file)
        parsed = parse_report(sheets, sheet_names)
    except Exception as e:
        st.error(f"文件读取或解析失败：{e}")
        st.stop()

    sku_map_df = load_sku_mapping(sku_file)
    if sku_map_df is not None:
        st.success("✅ 已加载 SKU 对照表")
    else:
        st.info("未上传有效 SKU 对照表，SKU 相关模块将仅显示编码。")

    render_header_info(parsed)

    repairs_df = parsed.get("repairs_df", pd.DataFrame()).copy()
    parts_df = parsed.get("parts_df", pd.DataFrame()).copy()

    if repairs_df.empty:
        st.warning("未解析出维修主表数据。")
        st.stop()

    repairs_df_filtered = render_sidebar_filters(repairs_df)

    tabs = st.tabs([
        "总览",
        "维修分析",
        "备件分析",
        "附加活动",
        "费用分析",
        "订单明细",
        "数据质量",
        "导出"
    ])

    with tabs[0]:
        render_overview_tab(parsed, repairs_df_filtered)

    with tabs[1]:
        render_repairs_tab(repairs_df_filtered)

    with tabs[2]:
        parts_filtered = parts_df.copy()
        if not parts_filtered.empty and "repair_id" in parts_filtered.columns and "repair_id" in repairs_df_filtered.columns:
            valid_ids = repairs_df_filtered["repair_id"].dropna().unique().tolist()
            parts_filtered = parts_filtered[parts_filtered["repair_id"].isin(valid_ids)]
        render_parts_tab(parts_filtered, sku_map_df)

    with tabs[3]:
        render_activity_tab(parsed)

    with tabs[4]:
        render_finance_tab(repairs_df_filtered, parsed)

    with tabs[5]:
        render_orders_tab(parsed, sku_map_df)

    with tabs[6]:
        render_quality_tab(parsed, repairs_df_filtered, parts_df)

    with tabs[7]:
        render_export_tab(parsed, repairs_df_filtered, parts_df)

else:
    st.info("请上传维修报告文件。")
