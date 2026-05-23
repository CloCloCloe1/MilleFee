from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st

from bp_generator import generate_outputs


st.set_page_config(page_title="MilleFee BP Generator", page_icon="MF", layout="wide")


TEXT = {
    "English": {
        "title": "MilleFee BP Generator",
        "caption": "Upload sales and stock files to generate a clean Excel analysis workbook and a Word business analysis report.",
        "language": "Language",
        "files": "Files",
        "sales": "Upload Sales Report",
        "stock": "Upload Stock Levels",
        "catalogue": "Optional: Upload Catalogue / Price List",
        "brand": "Brand name",
        "generate": "Generate Analysis",
        "missing_files": "Upload the Sales Report and Stock Levels Report to begin.",
        "ready": "Files are ready. Click Generate Analysis when you want to create the BP outputs.",
        "spinner": "Cleaning, merging, analyzing, and writing files...",
        "failed": "Generation failed",
        "total_skus": "Total SKUs",
        "qty_sold": "12M Qty Sold",
        "urgent": "Urgent / Stockout",
        "overstock": "Overstock SKUs",
        "download_excel": "Download MilleFee BP Data.xlsx",
        "download_word": "Download MilleFee Business Analysis.docx",
        "tab_final": "Final Analysis",
        "tab_summaries": "Summaries",
        "tab_priority": "Priority SKUs",
        "tab_detected": "Detected Columns",
        "final_analysis": "Final Analysis",
        "sabc_summary": "SABC Summary",
        "action_summary": "Action Summary",
        "inventory_summary": "Inventory Status Summary",
        "matrix": "SABC x Inventory Status Matrix",
        "replenishment": "Replenishment Priority",
        "overstock_risk": "Overstock Risk",
        "auto_detected": "Auto-detected Columns",
        "field": "Field",
        "detected_column": "Detected Column",
    },
    "中文": {
        "title": "MilleFee 商务分析生成器",
        "caption": "上传销售报表和库存报表，自动生成 Excel 分析文件和 Word 商务分析报告。",
        "language": "语言",
        "files": "上传文件",
        "sales": "上传销售报表",
        "stock": "上传库存报表",
        "catalogue": "可选：上传产品目录 / 价格表",
        "brand": "品牌名称",
        "generate": "生成分析",
        "missing_files": "请先上传销售报表和库存报表。",
        "ready": "文件已准备好。点击“生成分析”后即可创建 BP 输出文件。",
        "spinner": "正在清洗、合并、分析并生成文件...",
        "failed": "生成失败",
        "total_skus": "SKU 总数",
        "qty_sold": "12个月销量",
        "urgent": "紧急 / 缺货",
        "overstock": "库存过高 SKU",
        "download_excel": "下载 MilleFee BP Data.xlsx",
        "download_word": "下载 MilleFee Business Analysis.docx",
        "tab_final": "最终分析",
        "tab_summaries": "汇总",
        "tab_priority": "重点 SKU",
        "tab_detected": "识别列名",
        "final_analysis": "最终分析",
        "sabc_summary": "SABC 销售分类汇总",
        "action_summary": "行动建议汇总",
        "inventory_summary": "库存状态汇总",
        "matrix": "SABC x 库存状态矩阵",
        "replenishment": "补货优先级",
        "overstock_risk": "库存过高风险",
        "auto_detected": "自动识别的列名",
        "field": "字段",
        "detected_column": "识别到的列名",
    },
}


def load_file(uploaded_file):
    if uploaded_file is None:
        return None
    return BytesIO(uploaded_file.getvalue())


def metric_card(label: str, value: str, help_text: str | None = None):
    st.metric(label, value, help=help_text)


st.markdown(
    """
    <style>
    .block-container {padding-top: 2rem; max-width: 1180px;}
    h1, h2, h3 {letter-spacing: 0;}
    div[data-testid="stMetric"] {
        background: #f5f5f7;
        border: 1px solid #e8e8ed;
        border-radius: 8px;
        padding: 14px 16px;
    }
    div[data-testid="stDownloadButton"] button, div[data-testid="stButton"] button {
        border-radius: 8px;
        border: 1px solid #0071e3;
        background: #0071e3;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    language = st.selectbox(TEXT["English"]["language"], ["English", "中文"], index=0)

t = TEXT[language]

st.title(t["title"])
st.caption(t["caption"])

with st.sidebar:
    st.header(t["files"])
    sales_file = st.file_uploader(t["sales"], type=["xlsx", "xls"], key="sales")
    stock_file = st.file_uploader(t["stock"], type=["xlsx", "xls"], key="stock")
    catalogue_file = st.file_uploader(t["catalogue"], type=["xlsx", "xls"], key="catalogue")
    brand_name = st.text_input(t["brand"], value="MilleFee")
    generate = st.button(t["generate"], type="primary", use_container_width=True)

if not sales_file or not stock_file:
    st.info(t["missing_files"])
    st.stop()

if generate:
    with st.spinner(t["spinner"]):
        try:
            result, excel_bytes, word_bytes = generate_outputs(
                load_file(sales_file),
                load_file(stock_file),
                load_file(catalogue_file),
                brand_name=brand_name.strip() or "MilleFee",
            )
        except Exception as exc:
            st.error(f"{t['failed']}: {exc}")
            st.stop()

    st.session_state["result"] = result
    st.session_state["excel_bytes"] = excel_bytes
    st.session_state["word_bytes"] = word_bytes

if "result" not in st.session_state:
    st.warning(t["ready"])
    st.stop()

result = st.session_state["result"]
insights = result.insights

cols = st.columns(4)
with cols[0]:
    metric_card(t["total_skus"], f"{insights['total_skus']:,}")
with cols[1]:
    metric_card(t["qty_sold"], f"{insights['total_qty']:,.0f}")
with cols[2]:
    metric_card(t["urgent"], f"{insights['urgent_count']:,}")
with cols[3]:
    metric_card(t["overstock"], f"{insights['overstock_count']:,}")

st.divider()

download_cols = st.columns(2)
with download_cols[0]:
    st.download_button(
        t["download_excel"],
        data=st.session_state["excel_bytes"],
        file_name="MilleFee BP Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with download_cols[1]:
    st.download_button(
        t["download_word"],
        data=st.session_state["word_bytes"],
        file_name="MilleFee Business Analysis.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

tab1, tab2, tab3, tab4 = st.tabs([t["tab_final"], t["tab_summaries"], t["tab_priority"], t["tab_detected"]])

with tab1:
    st.subheader(t["final_analysis"])
    view = result.final.copy()
    for col in ["Contribution %", "Cumulative %"]:
        if col in view:
            view[col] = view[col].map(lambda x: f"{x:.1%}" if pd.notna(x) else "")
    if "Coverage" in view:
        view["Coverage"] = view["Coverage"].map(lambda x: "" if pd.isna(x) else f"{x:,.2f}")
    st.dataframe(view, use_container_width=True, height=520)

with tab2:
    left, right = st.columns(2)
    with left:
        st.subheader(t["sabc_summary"])
        st.dataframe(result.sabc_summary, use_container_width=True)
        st.subheader(t["action_summary"])
        st.dataframe(result.action_summary, use_container_width=True)
    with right:
        st.subheader(t["inventory_summary"])
        st.dataframe(result.inventory_status_summary, use_container_width=True)
        st.subheader(t["matrix"])
        st.dataframe(result.matrix, use_container_width=True)

with tab3:
    priority_cols = [
        "Product SKU",
        "Product Name",
        "Qty",
        "SABC Type",
        "Adjusted Future Inventory",
        "Coverage",
        "Inventory Status",
        "Action",
    ]
    st.subheader(t["replenishment"])
    st.dataframe(result.insights["urgent_table"][priority_cols], use_container_width=True)
    st.subheader(t["overstock_risk"])
    st.dataframe(result.insights["overstock_table"][priority_cols], use_container_width=True)

with tab4:
    st.subheader(t["auto_detected"])
    detected = pd.DataFrame([{t["field"]: key, t["detected_column"]: value} for key, value in result.detected_columns.items()])
    st.dataframe(detected, use_container_width=True, hide_index=True)
