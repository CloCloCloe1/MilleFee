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
        "column_notes": "Column explanation",
        "column_notes_help": "Open this section to understand how each column is calculated and used.",
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
        "qty_sold": "过去12个月销量",
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
        "column_notes": "字段解释",
        "column_notes_help": "点开这里可以查看每一列的计算方式和业务含义。",
    },
}


EXPLANATIONS = {
    "English": {
        "Final Analysis": {
            "Product SKU": "Unique product code used to merge Sales and Stock reports.",
            "Product Name": "Product description from the Sales report; stock name is used only if sales name is missing.",
            "Qty": "Total quantity sold in the latest available 12-month sales window.",
            "Contribution %": "SKU Qty divided by total Qty. Shows how much this SKU contributes to sales volume.",
            "Cumulative %": "Running total of Contribution % after SKUs are sorted from highest Qty to lowest Qty. Used to assign S/A/B/C type.",
            "SABC Type": "Sales priority class: S is the very top contribution, A is core, B is mid-tail, C is long-tail.",
            "Available": "Current available stock from the stock report.",
            "Incoming": "Incoming inventory or stock on order from the stock report.",
            "Future Inventory": "Available + Incoming. This estimates stock after incoming inventory arrives.",
            "Adjusted Future Inventory": "MAX(0, Future Inventory). Negative stock is treated as zero usable stock.",
            "Avg Monthly Sales": "Qty divided by 12. This represents average monthly sales based on the past 12 months.",
            "Coverage": "Adjusted Future Inventory divided by Avg Monthly Sales. It estimates how many months current/future stock can support.",
            "Inventory Status": "Stock health label based on coverage: Stockout, Urgent, Healthy, Monitor, Overstock, or No Sales / Review.",
            "Action": "Suggested business action based on Inventory Status.",
        },
        "SABC Summary": {
            "SABC Type": "Sales priority class generated from cumulative sales contribution.",
            "SKU_Count": "Number of SKUs in each SABC class.",
            "Qty": "Past 12-month sales quantity for all SKUs in that class.",
            "Avg_Coverage": "Average inventory coverage for SKUs in that class.",
            "Future_Inventory": "Total adjusted future inventory for SKUs in that class.",
            "Qty Share": "Qty in this SABC class divided by total Qty.",
        },
        "Inventory Status Summary": {
            "Inventory Status": "Stock health group based on coverage and sales demand.",
            "SKU_Count": "Number of SKUs in each inventory status.",
            "Qty": "Past 12-month sales quantity for SKUs in that status.",
            "Avg_Coverage": "Average months of inventory coverage for SKUs in that status.",
            "Future_Inventory": "Total adjusted future inventory for SKUs in that status.",
            "Qty Share": "Qty in this inventory status divided by total Qty.",
        },
        "Action Summary": {
            "Action": "Recommended business action, such as Replenish, Monitor, Review PO, Reduce PO, or Review SKU.",
            "SKU_Count": "Number of SKUs assigned to this action.",
            "Qty": "Past 12-month sales quantity for SKUs assigned to this action.",
            "Avg_Coverage": "Average inventory coverage for SKUs assigned to this action.",
            "Future_Inventory": "Total adjusted future inventory for SKUs assigned to this action.",
            "Qty Share": "Qty assigned to this action divided by total Qty.",
        },
        "Matrix": {
            "Rows": "SABC sales priority type.",
            "Columns": "Inventory Status groups.",
            "Values": "Number of SKUs in each SABC and Inventory Status combination.",
        },
        "Priority": {
            "Qty": "Past 12-month sales quantity.",
            "Adjusted Future Inventory": "Usable future inventory after preventing negative stock.",
            "Coverage": "Months of stock coverage based on average monthly sales.",
            "Inventory Status": "Current inventory risk label.",
            "Action": "Recommended next business action.",
        },
        "Detected Columns": {
            "Field": "Internal field needed by the app.",
            "Detected Column": "Column name automatically matched from the uploaded Excel file.",
        },
    },
    "中文": {
        "Final Analysis": {
            "Product SKU": "用于连接 Sales 和 Stock 两份报表的唯一产品编码。",
            "Product Name": "产品名称，优先来自销售报表；如果销售报表缺失，则使用库存报表名称。",
            "Qty": "过去 12 个月的总销售数量。",
            "Contribution %": "单个 SKU 的 Qty / 全部 SKU 的总 Qty，用来看该 SKU 对销量的贡献。",
            "Cumulative %": "按照 Qty 从高到低排序后，Contribution % 的累计值，用来判断 S/A/B/C 分类。",
            "SABC Type": "销售优先级：S 是最核心爆品，A 是核心 SKU，B 是中腰部，C 是长尾 SKU。",
            "Available": "库存报表中的当前可用库存。",
            "Incoming": "库存报表中的在途/即将入库库存。",
            "Future Inventory": "Available + Incoming，表示未来可用库存预估。",
            "Adjusted Future Inventory": "MAX(0, Future Inventory)，如果系统库存为负数，则按 0 可用库存处理。",
            "Avg Monthly Sales": "Qty / 12，基于过去 12 个月计算的平均月销量。",
            "Coverage": "Adjusted Future Inventory / Avg Monthly Sales，表示库存大约还能支持几个月销售。",
            "Inventory Status": "根据销量和 coverage 判断库存状态：缺货、紧急、健康、观察、库存过高或无销售/复核。",
            "Action": "根据库存状态自动给出的业务动作建议。",
        },
        "SABC Summary": {
            "SABC Type": "根据累计销售贡献得到的销售优先级。",
            "SKU_Count": "每个 SABC 分类下的 SKU 数量。",
            "Qty": "该分类下所有 SKU 过去 12 个月销量合计。",
            "Avg_Coverage": "该分类下 SKU 的平均库存覆盖月数。",
            "Future_Inventory": "该分类下所有 SKU 的调整后未来库存合计。",
            "Qty Share": "该分类 Qty / 全部 Qty。",
        },
        "Inventory Status Summary": {
            "Inventory Status": "根据 coverage 和销售需求得到的库存健康状态。",
            "SKU_Count": "每个库存状态下的 SKU 数量。",
            "Qty": "该库存状态下 SKU 的过去 12 个月销量合计。",
            "Avg_Coverage": "该库存状态下 SKU 的平均库存覆盖月数。",
            "Future_Inventory": "该库存状态下 SKU 的调整后未来库存合计。",
            "Qty Share": "该库存状态 Qty / 全部 Qty。",
        },
        "Action Summary": {
            "Action": "建议动作，例如 Replenish、Monitor、Review PO、Reduce PO 或 Review SKU。",
            "SKU_Count": "被分配到该动作的 SKU 数量。",
            "Qty": "该动作下 SKU 的过去 12 个月销量合计。",
            "Avg_Coverage": "该动作下 SKU 的平均库存覆盖月数。",
            "Future_Inventory": "该动作下 SKU 的调整后未来库存合计。",
            "Qty Share": "该动作下 Qty / 全部 Qty。",
        },
        "Matrix": {
            "Rows": "SABC 销售优先级分类。",
            "Columns": "库存状态分类。",
            "Values": "每一个 SABC x 库存状态组合下的 SKU 数量。",
        },
        "Priority": {
            "Qty": "过去 12 个月销售数量。",
            "Adjusted Future Inventory": "剔除负库存后的可用未来库存。",
            "Coverage": "基于平均月销量计算的库存覆盖月数。",
            "Inventory Status": "当前库存风险标签。",
            "Action": "建议下一步业务动作。",
        },
        "Detected Columns": {
            "Field": "系统内部需要识别的字段。",
            "Detected Column": "从上传 Excel 中自动匹配到的列名。",
        },
    },
}


def load_file(uploaded_file):
    if uploaded_file is None:
        return None
    return BytesIO(uploaded_file.getvalue())


def metric_card(label: str, value: str, help_text: str | None = None):
    st.metric(label, value, help=help_text)


def show_column_notes(language: str, table_key: str):
    notes = EXPLANATIONS[language][table_key]
    title = TEXT[language]["column_notes"]
    help_text = TEXT[language]["column_notes_help"]
    with st.expander(f"{title} - {help_text}", expanded=False):
        rows = [{"Column": key, "Explanation": value} for key, value in notes.items()]
        if language == "中文":
            rows = [{"列名": key, "解释": value} for key, value in notes.items()]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


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
    show_column_notes(language, "Final Analysis")
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
        show_column_notes(language, "SABC Summary")
        st.dataframe(result.sabc_summary, use_container_width=True)
        st.subheader(t["action_summary"])
        show_column_notes(language, "Action Summary")
        st.dataframe(result.action_summary, use_container_width=True)
    with right:
        st.subheader(t["inventory_summary"])
        show_column_notes(language, "Inventory Status Summary")
        st.dataframe(result.inventory_status_summary, use_container_width=True)
        st.subheader(t["matrix"])
        show_column_notes(language, "Matrix")
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
    show_column_notes(language, "Priority")
    st.dataframe(result.insights["urgent_table"][priority_cols], use_container_width=True)
    st.subheader(t["overstock_risk"])
    show_column_notes(language, "Priority")
    st.dataframe(result.insights["overstock_table"][priority_cols], use_container_width=True)

with tab4:
    st.subheader(t["auto_detected"])
    show_column_notes(language, "Detected Columns")
    detected = pd.DataFrame([{t["field"]: key, t["detected_column"]: value} for key, value in result.detected_columns.items()])
    st.dataframe(detected, use_container_width=True, hide_index=True)
