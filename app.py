from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st

from auth import authenticate, change_own_password, create_user, delete_user, list_users, reset_password
from bp_generator import generate_outputs


st.set_page_config(page_title="Beauty Brand BP Generator", page_icon="BP", layout="wide")


ZH = "\u4e2d\u6587"

TEXT = {
    "English": {
        "title": "Beauty Brand BP Generator",
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
        "download_excel": "Download BP Data.xlsx",
        "download_word": "Download Business Analysis.docx",
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
        "login_title": "Sign in",
        "username": "Username",
        "password": "Password",
        "login": "Log in",
        "logout": "Log out",
        "bad_login": "Incorrect username or password.",
        "signed_in": "Signed in as",
        "page": "Page",
        "bp_page": "BP Generator",
        "manage_users": "Manage Users",
        "add_user": "Add user",
        "delete_user": "Delete user",
        "reset_password": "Reset user password",
        "change_password": "Change my password",
        "role": "Role",
        "new_password": "New password",
        "current_password": "Current password",
        "confirm": "Confirm",
        "security_note": "Passwords are stored as hashes. Admin can reset a user's password but cannot view the original password.",
        "default_admin_note": "Default admin is admin / change-me-now. Please change it after first login.",
    },
    ZH: {
        "title": "\u7f8e\u5986\u54c1\u724c BP \u5206\u6790\u751f\u6210\u5668",
        "caption": "\u4e0a\u4f20\u9500\u552e\u62a5\u8868\u548c\u5e93\u5b58\u62a5\u8868\uff0c\u81ea\u52a8\u751f\u6210 Excel \u5206\u6790\u6587\u4ef6\u548c Word \u5546\u52a1\u5206\u6790\u62a5\u544a\u3002",
        "language": "\u8bed\u8a00",
        "files": "\u4e0a\u4f20\u6587\u4ef6",
        "sales": "\u4e0a\u4f20\u9500\u552e\u62a5\u8868",
        "stock": "\u4e0a\u4f20\u5e93\u5b58\u62a5\u8868",
        "catalogue": "\u53ef\u9009\uff1a\u4e0a\u4f20\u4ea7\u54c1\u76ee\u5f55 / \u4ef7\u683c\u8868",
        "brand": "\u54c1\u724c\u540d\u79f0",
        "generate": "\u751f\u6210\u5206\u6790",
        "missing_files": "\u8bf7\u5148\u4e0a\u4f20\u9500\u552e\u62a5\u8868\u548c\u5e93\u5b58\u62a5\u8868\u3002",
        "ready": "\u6587\u4ef6\u5df2\u51c6\u5907\u597d\u3002\u70b9\u51fb\u201c\u751f\u6210\u5206\u6790\u201d\u540e\u5373\u53ef\u521b\u5efa BP \u8f93\u51fa\u6587\u4ef6\u3002",
        "spinner": "\u6b63\u5728\u6e05\u6d17\u3001\u5408\u5e76\u3001\u5206\u6790\u5e76\u751f\u6210\u6587\u4ef6...",
        "failed": "\u751f\u6210\u5931\u8d25",
        "total_skus": "SKU \u603b\u6570",
        "qty_sold": "\u8fc7\u53bb12\u4e2a\u6708\u9500\u91cf",
        "urgent": "\u7d27\u6025 / \u7f3a\u8d27",
        "overstock": "\u5e93\u5b58\u8fc7\u9ad8 SKU",
        "download_excel": "\u4e0b\u8f7d BP Data.xlsx",
        "download_word": "\u4e0b\u8f7d Business Analysis.docx",
        "tab_final": "\u6700\u7ec8\u5206\u6790",
        "tab_summaries": "\u6c47\u603b",
        "tab_priority": "\u91cd\u70b9 SKU",
        "tab_detected": "\u8bc6\u522b\u5217\u540d",
        "final_analysis": "\u6700\u7ec8\u5206\u6790",
        "sabc_summary": "SABC \u9500\u552e\u5206\u7c7b\u6c47\u603b",
        "action_summary": "\u884c\u52a8\u5efa\u8bae\u6c47\u603b",
        "inventory_summary": "\u5e93\u5b58\u72b6\u6001\u6c47\u603b",
        "matrix": "SABC x \u5e93\u5b58\u72b6\u6001\u77e9\u9635",
        "replenishment": "\u8865\u8d27\u4f18\u5148\u7ea7",
        "overstock_risk": "\u5e93\u5b58\u8fc7\u9ad8\u98ce\u9669",
        "auto_detected": "\u81ea\u52a8\u8bc6\u522b\u7684\u5217\u540d",
        "field": "\u5b57\u6bb5",
        "detected_column": "\u8bc6\u522b\u5230\u7684\u5217\u540d",
        "column_notes": "\u5b57\u6bb5\u89e3\u91ca",
        "column_notes_help": "\u70b9\u5f00\u8fd9\u91cc\u53ef\u4ee5\u67e5\u770b\u6bcf\u4e00\u5217\u7684\u8ba1\u7b97\u65b9\u5f0f\u548c\u4e1a\u52a1\u542b\u4e49\u3002",
        "login_title": "\u767b\u5f55",
        "username": "\u7528\u6237\u540d",
        "password": "\u5bc6\u7801",
        "login": "\u767b\u5f55",
        "logout": "\u9000\u51fa\u767b\u5f55",
        "bad_login": "\u7528\u6237\u540d\u6216\u5bc6\u7801\u4e0d\u6b63\u786e\u3002",
        "signed_in": "\u5df2\u767b\u5f55",
        "page": "\u9875\u9762",
        "bp_page": "BP \u751f\u6210\u5668",
        "manage_users": "\u7528\u6237\u7ba1\u7406",
        "add_user": "\u65b0\u589e\u7528\u6237",
        "delete_user": "\u5220\u9664\u7528\u6237",
        "reset_password": "\u91cd\u7f6e\u7528\u6237\u5bc6\u7801",
        "change_password": "\u4fee\u6539\u6211\u7684\u5bc6\u7801",
        "role": "\u89d2\u8272",
        "new_password": "\u65b0\u5bc6\u7801",
        "current_password": "\u5f53\u524d\u5bc6\u7801",
        "confirm": "\u786e\u8ba4",
        "security_note": "\u5bc6\u7801\u4f1a\u52a0\u5bc6\u4fdd\u5b58\u3002Admin \u53ef\u4ee5\u91cd\u7f6e\u7528\u6237\u5bc6\u7801\uff0c\u4f46\u4e0d\u4f1a\u663e\u793a\u539f\u59cb\u660e\u6587\u5bc6\u7801\u3002",
        "default_admin_note": "\u9ed8\u8ba4 admin \u8d26\u53f7\u662f admin / change-me-now\u3002\u9996\u6b21\u767b\u5f55\u540e\u8bf7\u5c3d\u5feb\u4fee\u6539\u3002",
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
            "Inventory Status": "Stock health label based on coverage.",
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
        "Matrix": {"Rows": "SABC sales priority type.", "Columns": "Inventory Status groups.", "Values": "Number of SKUs in each combination."},
        "Priority": {
            "Qty": "Past 12-month sales quantity.",
            "Adjusted Future Inventory": "Usable future inventory after preventing negative stock.",
            "Coverage": "Months of stock coverage based on average monthly sales.",
            "Inventory Status": "Current inventory risk label.",
            "Action": "Recommended next business action.",
        },
        "Detected Columns": {"Field": "Internal field needed by the app.", "Detected Column": "Column name automatically matched from the uploaded Excel file."},
    },
    ZH: {
        "Final Analysis": {
            "Product SKU": "\u7528\u4e8e\u8fde\u63a5 Sales \u548c Stock \u4e24\u4efd\u62a5\u8868\u7684\u552f\u4e00\u4ea7\u54c1\u7f16\u7801\u3002",
            "Product Name": "\u4ea7\u54c1\u540d\u79f0\uff0c\u4f18\u5148\u6765\u81ea\u9500\u552e\u62a5\u8868\u3002",
            "Qty": "\u8fc7\u53bb 12 \u4e2a\u6708\u7684\u603b\u9500\u552e\u6570\u91cf\u3002",
            "Contribution %": "\u5355\u4e2a SKU \u7684 Qty / \u5168\u90e8 SKU \u7684\u603b Qty\u3002",
            "Cumulative %": "\u6309 Qty \u4ece\u9ad8\u5230\u4f4e\u6392\u5e8f\u540e\uff0cContribution % \u7684\u7d2f\u8ba1\u503c\u3002",
            "SABC Type": "\u9500\u552e\u4f18\u5148\u7ea7\uff1aS/A \u662f\u6838\u5fc3 SKU\uff0cB/C \u662f\u4e2d\u957f\u5c3e SKU\u3002",
            "Available": "\u5e93\u5b58\u62a5\u8868\u4e2d\u7684\u5f53\u524d\u53ef\u7528\u5e93\u5b58\u3002",
            "Incoming": "\u5e93\u5b58\u62a5\u8868\u4e2d\u7684\u5728\u9014 / \u5373\u5c06\u5165\u5e93\u5e93\u5b58\u3002",
            "Future Inventory": "Available + Incoming\uff0c\u8868\u793a\u672a\u6765\u53ef\u7528\u5e93\u5b58\u9884\u4f30\u3002",
            "Adjusted Future Inventory": "MAX(0, Future Inventory)\uff0c\u8d1f\u5e93\u5b58\u6309 0 \u5904\u7406\u3002",
            "Avg Monthly Sales": "Qty / 12\uff0c\u57fa\u4e8e\u8fc7\u53bb 12 \u4e2a\u6708\u8ba1\u7b97\u7684\u5e73\u5747\u6708\u9500\u91cf\u3002",
            "Coverage": "Adjusted Future Inventory / Avg Monthly Sales\uff0c\u8868\u793a\u5e93\u5b58\u5927\u7ea6\u8fd8\u80fd\u652f\u6301\u51e0\u4e2a\u6708\u9500\u552e\u3002",
            "Inventory Status": "\u6839\u636e\u9500\u91cf\u548c coverage \u5224\u65ad\u5e93\u5b58\u72b6\u6001\u3002",
            "Action": "\u6839\u636e\u5e93\u5b58\u72b6\u6001\u81ea\u52a8\u7ed9\u51fa\u7684\u4e1a\u52a1\u5efa\u8bae\u3002",
        },
        "SABC Summary": {
            "SABC Type": "\u6839\u636e\u7d2f\u8ba1\u9500\u552e\u8d21\u732e\u5f97\u5230\u7684\u9500\u552e\u4f18\u5148\u7ea7\u3002",
            "SKU_Count": "\u6bcf\u4e2a SABC \u5206\u7c7b\u4e0b\u7684 SKU \u6570\u91cf\u3002",
            "Qty": "\u8be5\u5206\u7c7b\u4e0b\u6240\u6709 SKU \u8fc7\u53bb 12 \u4e2a\u6708\u9500\u91cf\u5408\u8ba1\u3002",
            "Avg_Coverage": "\u8be5\u5206\u7c7b\u4e0b SKU \u7684\u5e73\u5747\u5e93\u5b58\u8986\u76d6\u6708\u6570\u3002",
            "Future_Inventory": "\u8be5\u5206\u7c7b\u4e0b\u8c03\u6574\u540e\u672a\u6765\u5e93\u5b58\u5408\u8ba1\u3002",
            "Qty Share": "\u8be5\u5206\u7c7b Qty / \u5168\u90e8 Qty\u3002",
        },
        "Inventory Status Summary": {
            "Inventory Status": "\u6839\u636e coverage \u548c\u9500\u552e\u9700\u6c42\u5f97\u5230\u7684\u5e93\u5b58\u5065\u5eb7\u72b6\u6001\u3002",
            "SKU_Count": "\u6bcf\u4e2a\u5e93\u5b58\u72b6\u6001\u4e0b\u7684 SKU \u6570\u91cf\u3002",
            "Qty": "\u8be5\u5e93\u5b58\u72b6\u6001\u4e0b SKU \u7684\u8fc7\u53bb 12 \u4e2a\u6708\u9500\u91cf\u5408\u8ba1\u3002",
            "Avg_Coverage": "\u8be5\u72b6\u6001\u4e0b SKU \u7684\u5e73\u5747\u5e93\u5b58\u8986\u76d6\u6708\u6570\u3002",
            "Future_Inventory": "\u8be5\u72b6\u6001\u4e0b SKU \u7684\u8c03\u6574\u540e\u672a\u6765\u5e93\u5b58\u5408\u8ba1\u3002",
            "Qty Share": "\u8be5\u72b6\u6001 Qty / \u5168\u90e8 Qty\u3002",
        },
        "Action Summary": {
            "Action": "\u5efa\u8bae\u52a8\u4f5c\uff0c\u4f8b\u5982 Replenish\u3001Monitor\u3001Reduce PO \u6216 Review SKU\u3002",
            "SKU_Count": "\u88ab\u5206\u914d\u5230\u8be5\u52a8\u4f5c\u7684 SKU \u6570\u91cf\u3002",
            "Qty": "\u8be5\u52a8\u4f5c\u4e0b SKU \u7684\u8fc7\u53bb 12 \u4e2a\u6708\u9500\u91cf\u5408\u8ba1\u3002",
            "Avg_Coverage": "\u8be5\u52a8\u4f5c\u4e0b SKU \u7684\u5e73\u5747\u5e93\u5b58\u8986\u76d6\u6708\u6570\u3002",
            "Future_Inventory": "\u8be5\u52a8\u4f5c\u4e0b SKU \u7684\u8c03\u6574\u540e\u672a\u6765\u5e93\u5b58\u5408\u8ba1\u3002",
            "Qty Share": "\u8be5\u52a8\u4f5c Qty / \u5168\u90e8 Qty\u3002",
        },
        "Matrix": {"Rows": "SABC \u9500\u552e\u4f18\u5148\u7ea7\u5206\u7c7b\u3002", "Columns": "\u5e93\u5b58\u72b6\u6001\u5206\u7c7b\u3002", "Values": "\u6bcf\u4e2a\u7ec4\u5408\u4e0b\u7684 SKU \u6570\u91cf\u3002"},
        "Priority": {
            "Qty": "\u8fc7\u53bb 12 \u4e2a\u6708\u9500\u552e\u6570\u91cf\u3002",
            "Adjusted Future Inventory": "\u5254\u9664\u8d1f\u5e93\u5b58\u540e\u7684\u53ef\u7528\u672a\u6765\u5e93\u5b58\u3002",
            "Coverage": "\u57fa\u4e8e\u5e73\u5747\u6708\u9500\u91cf\u8ba1\u7b97\u7684\u5e93\u5b58\u8986\u76d6\u6708\u6570\u3002",
            "Inventory Status": "\u5f53\u524d\u5e93\u5b58\u98ce\u9669\u6807\u7b7e\u3002",
            "Action": "\u5efa\u8bae\u4e0b\u4e00\u6b65\u4e1a\u52a1\u52a8\u4f5c\u3002",
        },
        "Detected Columns": {"Field": "\u7cfb\u7edf\u5185\u90e8\u9700\u8981\u8bc6\u522b\u7684\u5b57\u6bb5\u3002", "Detected Column": "\u4ece\u4e0a\u4f20 Excel \u4e2d\u81ea\u52a8\u5339\u914d\u5230\u7684\u5217\u540d\u3002"},
    },
}


def load_file(uploaded_file):
    if uploaded_file is None:
        return None
    return BytesIO(uploaded_file.getvalue())


def metric_card(label: str, value: str, help_text: str | None = None):
    st.metric(label, value, help=help_text)


def clean_brand_name(value: str) -> str:
    cleaned = "".join(ch for ch in value.strip() if ch.isalnum() or ch in {" ", "-", "_"}).strip()
    return cleaned or "Brand"


def show_column_notes(language: str, table_key: str):
    notes = EXPLANATIONS[language][table_key]
    title = TEXT[language]["column_notes"]
    help_text = TEXT[language]["column_notes_help"]
    with st.expander(f"{title} - {help_text}", expanded=False):
        if language == ZH:
            rows = [{"\u5217\u540d": key, "\u89e3\u91ca": value} for key, value in notes.items()]
        else:
            rows = [{"Column": key, "Explanation": value} for key, value in notes.items()]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def apply_styles():
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


def login_page(t: dict):
    st.title(t["login_title"])
    st.caption(t["default_admin_note"])
    with st.form("login_form"):
        username = st.text_input(t["username"])
        password = st.text_input(t["password"], type="password")
        submitted = st.form_submit_button(t["login"], use_container_width=True)
    if submitted:
        user = authenticate(username, password)
        if user:
            st.session_state["user"] = user
            st.rerun()
        st.error(t["bad_login"])


def manage_users_page(t: dict):
    st.title(t["manage_users"])
    st.info(t["security_note"])
    if st.session_state["user"].get("must_change_password"):
        st.warning(t["default_admin_note"])

    st.subheader(t["add_user"])
    with st.form("add_user_form", clear_on_submit=True):
        new_username = st.text_input(t["username"], key="new_user")
        new_password = st.text_input(t["new_password"], type="password", key="new_password")
        role = st.selectbox(t["role"], ["user", "admin"])
        if st.form_submit_button(t["add_user"]):
            ok, message = create_user(new_username, new_password, role)
            st.success(message) if ok else st.error(message)

    st.subheader(t["reset_password"])
    users_df = pd.DataFrame(list_users())
    st.dataframe(users_df, use_container_width=True, hide_index=True)
    usernames = users_df["Username"].tolist() if not users_df.empty else []
    with st.form("reset_password_form"):
        reset_username = st.selectbox(t["username"], usernames, key="reset_user")
        reset_new_password = st.text_input(t["new_password"], type="password", key="reset_password")
        if st.form_submit_button(t["reset_password"]):
            ok, message = reset_password(reset_username, reset_new_password)
            st.success(message) if ok else st.error(message)

    st.subheader(t["delete_user"])
    with st.form("delete_user_form"):
        delete_username = st.selectbox(t["username"], usernames, key="delete_user")
        if st.form_submit_button(t["delete_user"]):
            ok, message = delete_user(delete_username, st.session_state["user"]["username"])
            st.success(message) if ok else st.error(message)

    st.subheader(t["change_password"])
    with st.form("change_own_password_form"):
        current_password = st.text_input(t["current_password"], type="password")
        own_new_password = st.text_input(t["new_password"], type="password", key="own_new_password")
        if st.form_submit_button(t["change_password"]):
            ok, message = change_own_password(st.session_state["user"]["username"], current_password, own_new_password)
            if ok:
                st.session_state["user"]["must_change_password"] = False
                st.success(message)
            else:
                st.error(message)


def bp_generator_page(t: dict, language: str):
    st.title(t["title"])
    st.caption(t["caption"])

    with st.sidebar:
        st.header(t["files"])
        sales_file = st.file_uploader(t["sales"], type=["xlsx", "xls"], key="sales")
        stock_file = st.file_uploader(t["stock"], type=["xlsx", "xls"], key="stock")
        catalogue_file = st.file_uploader(t["catalogue"], type=["xlsx", "xls"], key="catalogue")
        brand_name = st.text_input(t["brand"], value="", placeholder="MilleFee / Judydoll / Joocyee")
        generate = st.button(t["generate"], type="primary", use_container_width=True)

    if not sales_file or not stock_file:
        st.info(t["missing_files"])
        return

    if generate:
        with st.spinner(t["spinner"]):
            try:
                result, excel_bytes, word_bytes = generate_outputs(
                    load_file(sales_file),
                    load_file(stock_file),
                    load_file(catalogue_file),
                    brand_name=clean_brand_name(brand_name),
                )
            except Exception as exc:
                st.error(f"{t['failed']}: {exc}")
                return

        st.session_state["result"] = result
        st.session_state["excel_bytes"] = excel_bytes
        st.session_state["word_bytes"] = word_bytes

    if "result" not in st.session_state:
        st.warning(t["ready"])
        return

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
    output_brand = clean_brand_name(brand_name)
    with download_cols[0]:
        st.download_button(t["download_excel"], data=st.session_state["excel_bytes"], file_name=f"{output_brand} BP Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with download_cols[1]:
        st.download_button(t["download_word"], data=st.session_state["word_bytes"], file_name=f"{output_brand} Business Analysis.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

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
        priority_cols = ["Product SKU", "Product Name", "Qty", "SABC Type", "Adjusted Future Inventory", "Coverage", "Inventory Status", "Action"]
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


apply_styles()

with st.sidebar:
    language = st.selectbox(TEXT["English"]["language"], ["English", ZH], index=0)

t = TEXT[language]

if "user" not in st.session_state:
    login_page(t)
    st.stop()

with st.sidebar:
    st.caption(f"{t['signed_in']}: {st.session_state['user']['username']} ({st.session_state['user']['role']})")
    if st.button(t["logout"], use_container_width=True):
        for key in ["user", "result", "excel_bytes", "word_bytes"]:
            st.session_state.pop(key, None)
        st.rerun()

if st.session_state["user"]["role"] == "admin":
    page = st.sidebar.radio(t["page"], [t["bp_page"], t["manage_users"]])
    if page == t["manage_users"]:
        manage_users_page(t)
    else:
        bp_generator_page(t, language)
else:
    bp_generator_page(t, language)
