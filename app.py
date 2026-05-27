from __future__ import annotations

from io import BytesIO

import altair as alt
import pandas as pd
import streamlit as st

from auth import authenticate, change_own_password, create_user, delete_user, list_users, reset_password
from bp_generator import generate_enriched_catalogue, generate_outputs


st.set_page_config(page_title="Beauty Brand BP Generator", page_icon="BP", layout="wide")


ZH = "\u4e2d\u6587"

TEXT = {
    "English": {
        "title": "Beauty Brand BP Generator",
        "caption": "Upload sales and stock files to generate a clean Excel analysis workbook and a Word business analysis report.",
        "language": "Language",
        "files": "Files",
        "sales": "Upload Sales Report",
        "sales_2024": "Upload 2024 Sales Report",
        "sales_2025": "Upload 2025 Sales Report",
        "sales_2026": "Upload 2026 YTD Sales Report",
        "stock": "Upload Stock Levels",
        "catalogue": "Optional: Upload Catalogue / Price List",
        "purchase": "Optional: Upload Purchase / PO History",
        "purchase_2024": "Upload 2024 PO with lines",
        "purchase_2025": "Upload 2025 PO with lines",
        "purchase_2026": "Upload 2026 YTD PO with lines",
        "purchase_keyword": "Brand / product filter keyword",
        "location_filter": "Location filter",
        "location_filter_help": "Optional. Use this when Sales is exported for one location but PO/Stock are not location-filtered.",
        "brand": "Brand name",
        "generate": "Generate Analysis",
        "missing_files": "Upload the Sales Report and Stock Levels Report to begin.",
        "ready": "Files are ready. Click Generate Analysis when you want to create the BP outputs.",
        "spinner": "Cleaning, merging, analyzing, and writing files...",
        "failed": "Generation failed",
        "total_skus": "Total SKUs",
        "qty_sold": "Sales Qty",
        "urgent": "Urgent / Stockout",
        "overstock": "Overstock SKUs",
        "download_excel": "Download BP Data.xlsx",
        "download_word": "Download Business Analysis.docx",
        "download_catalogue": "Download Updated Catalogue.xlsx",
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
        "sales_year_summary": "Sales Summary by Year",
        "sales_purchase_year_summary": "Sales vs PO by Year",
        "purchase_sku_summary": "Purchase by SKU",
        "location_year_view": "Latest Sales Year + Current Inventory",
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
        "sales_2024": "\u4e0a\u4f20 2024 Sales Report",
        "sales_2025": "\u4e0a\u4f20 2025 Sales Report",
        "sales_2026": "\u4e0a\u4f20 2026 YTD Sales Report",
        "stock": "\u4e0a\u4f20\u5e93\u5b58\u62a5\u8868",
        "catalogue": "\u53ef\u9009\uff1a\u4e0a\u4f20\u4ea7\u54c1\u76ee\u5f55 / \u4ef7\u683c\u8868",
        "purchase": "\u53ef\u9009\uff1a\u4e0a\u4f20\u91c7\u8d2d / PO \u5386\u53f2",
        "purchase_2024": "\u4e0a\u4f20 2024 PO with lines",
        "purchase_2025": "\u4e0a\u4f20 2025 PO with lines",
        "purchase_2026": "\u4e0a\u4f20 2026 YTD PO with lines",
        "purchase_keyword": "\u54c1\u724c / \u4ea7\u54c1\u7b5b\u9009\u5173\u952e\u8bcd",
        "location_filter": "Location \u7b5b\u9009",
        "location_filter_help": "\u53ef\u9009\u3002\u5982\u679c Sales \u662f\u4f60\u5df2\u7ecf\u5728\u7cfb\u7edf\u91cc\u6309\u67d0\u4e2a location \u5bfc\u51fa\u7684\uff0c\u4f46 PO/Stock \u662f\u5168\u91cf\uff0c\u5c31\u5728\u8fd9\u91cc\u586b\u540c\u4e00\u4e2a location\u3002",
        "brand": "\u54c1\u724c\u540d\u79f0",
        "generate": "\u751f\u6210\u5206\u6790",
        "missing_files": "\u8bf7\u5148\u4e0a\u4f20\u9500\u552e\u62a5\u8868\u548c\u5e93\u5b58\u62a5\u8868\u3002",
        "ready": "\u6587\u4ef6\u5df2\u51c6\u5907\u597d\u3002\u70b9\u51fb\u201c\u751f\u6210\u5206\u6790\u201d\u540e\u5373\u53ef\u521b\u5efa BP \u8f93\u51fa\u6587\u4ef6\u3002",
        "spinner": "\u6b63\u5728\u6e05\u6d17\u3001\u5408\u5e76\u3001\u5206\u6790\u5e76\u751f\u6210\u6587\u4ef6...",
        "failed": "\u751f\u6210\u5931\u8d25",
        "total_skus": "SKU \u603b\u6570",
        "qty_sold": "\u9500\u552e\u6570\u91cf",
        "urgent": "\u7d27\u6025 / \u7f3a\u8d27",
        "overstock": "\u5e93\u5b58\u8fc7\u9ad8 SKU",
        "download_excel": "\u4e0b\u8f7d BP Data.xlsx",
        "download_word": "\u4e0b\u8f7d Business Analysis.docx",
        "download_catalogue": "\u4e0b\u8f7d\u66f4\u65b0\u540e\u7684\u62a5\u4ef7\u5355.xlsx",
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
        "sales_year_summary": "\u6309\u5e74\u4efd\u6c47\u603b\u9500\u552e",
        "sales_purchase_year_summary": "\u6309\u5e74\u4efd\u5bf9\u6bd4 Sales vs PO",
        "purchase_sku_summary": "\u6309 SKU \u6c47\u603b\u91c7\u8d2d\u91d1\u989d",
        "location_year_view": "\u6700\u65b0\u5e74\u4efd Sales + \u5f53\u524d\u5e93\u5b58",
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
            "Qty": "Sales quantity used for current inventory decisions. If year-labeled sales files are uploaded, this uses only the latest uploaded year.",
            "Sales Amount ($ CAD)": "Sales revenue in the current analysis period when the Sales report includes an amount column.",
            "Profit ($ CAD)": "Profit in the current analysis period when the Sales report includes a profit column. Higher is usually better.",
            "Sales Margin %": "Profit divided by Sales Amount for the current analysis period. Higher is usually better.",
            "Contribution %": "SKU Qty divided by total Qty. Shows how much this SKU contributes to sales volume.",
            "Cumulative %": "Running total of Contribution % after SKUs are sorted from highest Qty to lowest Qty. Used to assign S/A/B/C type.",
            "SABC Type": "Sales priority class: S is the very top contribution, A is core, B is mid-tail, C is long-tail.",
            "Available": "Current available stock from the stock report.",
            "Incoming": "Incoming inventory or stock on order from the stock report.",
            "Future Inventory": "Available + Incoming. This estimates stock after incoming inventory arrives.",
            "Adjusted Future Inventory": "MAX(0, Future Inventory). Negative stock is treated as zero usable stock.",
            "Avg Monthly Sales": "Qty divided by 12. For year-labeled uploads, Qty is the latest uploaded year and this field is used as a simple coverage benchmark.",
            "Coverage": "Adjusted Future Inventory divided by Avg Monthly Sales. It estimates how many months current/future stock can support.",
            "Inventory Status": "Stock health label based on coverage.",
            "Action": "Suggested business action based on Inventory Status.",
            "2024 PO Cost ($ CAD)": "PO line cost matched to this SKU for the 2024 uploaded PO file.",
            "2025 PO Cost ($ CAD)": "PO line cost matched to this SKU for the 2025 uploaded PO file.",
            "2026 PO Cost ($ CAD)": "PO line cost matched to this SKU for the 2026 YTD uploaded PO file.",
            "Total PO Cost ($ CAD)": "Combined PO cost from 2024, 2025, and 2026 uploaded PO files.",
        },
        "SABC Summary": {
            "SABC Type": "Sales priority class generated from cumulative sales contribution.",
            "SKU_Count": "Number of SKUs in each SABC class.",
            "Qty": "Sales quantity from the current analysis period, usually the latest uploaded sales year.",
            "Avg_Coverage": "Average inventory coverage for SKUs in that class.",
            "Future_Inventory": "Total adjusted future inventory for SKUs in that class.",
            "Qty Share": "Qty in this SABC class divided by total Qty.",
        },
        "Inventory Status Summary": {
            "Inventory Status": "Stock health group based on coverage and sales demand.",
            "SKU_Count": "Number of SKUs in each inventory status.",
            "Qty": "Sales quantity from the current analysis period, usually the latest uploaded sales year.",
            "Avg_Coverage": "Average months of inventory coverage for SKUs in that status.",
            "Future_Inventory": "Total adjusted future inventory for SKUs in that status.",
            "Qty Share": "Qty in this inventory status divided by total Qty.",
        },
        "Action Summary": {
            "Action": "Recommended business action, such as Replenish, Monitor, Review PO, Reduce PO, or Review SKU.",
            "SKU_Count": "Number of SKUs assigned to this action.",
            "Qty": "Sales quantity from the current analysis period, usually the latest uploaded sales year.",
            "Avg_Coverage": "Average inventory coverage for SKUs assigned to this action.",
            "Future_Inventory": "Total adjusted future inventory for SKUs assigned to this action.",
            "Qty Share": "Qty assigned to this action divided by total Qty.",
        },
        "Matrix": {"Rows": "SABC sales priority type.", "Columns": "Inventory Status groups.", "Values": "Number of SKUs in each combination."},
        "Priority": {
            "Qty": "Sales quantity from the current analysis period, usually the latest uploaded sales year.",
            "Sales Amount ($ CAD)": "Sales revenue from the current analysis period when available.",
            "Profit ($ CAD)": "Profit from the current analysis period when available. Higher is usually better.",
            "Sales Margin %": "Profit divided by Sales Amount when available. Higher is usually better.",
            "Adjusted Future Inventory": "Usable future inventory after preventing negative stock.",
            "Coverage": "Months of stock coverage based on average monthly sales.",
            "Inventory Status": "Current inventory risk label.",
            "Action": "Recommended next business action.",
        },
        "Detected Columns": {"Field": "Internal field needed by the app.", "Detected Column": "Column name automatically matched from the uploaded Excel file."},
        "Sales Summary by Year": {
            "Year": "Year assigned by the upload slot, or detected from the sales date when using a single sales file.",
            "SKU Count": "Number of SKUs sold in that year.",
            "Sales Qty": "Total sales quantity for the uploaded sales file assigned to that year.",
            "Sales Amount ($ CAD)": "Sales revenue amount for that year when the sales report includes an amount column.",
            "Profit ($ CAD)": "Profit from the sales report when available. Higher is usually better.",
            "Sales Margin %": "Profit divided by Sales Amount. Higher is usually better because each sales dollar creates more profit.",
        },
        "Sales vs PO by Year": {
            "Year": "Year used to connect the sales file and PO file.",
            "Sales Qty": "Sales quantity from the uploaded sales file for that year.",
            "Purchase Qty": "Total purchased quantity from the uploaded PO file for that year after brand and location filters.",
            "Sales Amount ($ CAD)": "Sales revenue from the sales report for that year when available.",
            "Profit ($ CAD)": "Profit from the sales report when available. Higher is usually better.",
            "Sales Margin %": "Profit divided by Sales Amount. Higher is usually better because growth is more profitable.",
            "PO Cost ($ CAD)": "Total PO line cost for that year after brand and location filters. This is the purchasing/inbound cost, not sales revenue.",
            "PO Cost / Sales Qty ($ CAD)": "PO cost divided by units sold. Lower is usually better because less purchasing cost was needed per unit sold, but read it together with current stock and replenishment needs.",
            "PO Cost / Sales Amount": "PO cost divided by sales revenue when sales amount is available. Lower usually means better sales efficiency versus purchase cost.",
        },
        "Purchase by SKU": {
            "SKU": "SKU detected from PO line file.",
            "Product": "Product name detected from PO line file.",
            "PO Cost ($ CAD)": "Total PO line cost after applying the brand and location filters.",
            "Quantity": "Total purchased quantity when available.",
        },
        "Latest Inventory View": {
            "Year": "Latest uploaded sales year. Current stock is compared only to this year.",
            "Location": "Location detected from Sales, Stock, or PO files.",
            "Sales Qty": "Sales quantity in this year and location when Sales Report includes date and location.",
            "Sales Amount ($ CAD)": "Sales revenue in this year and location when a sales amount column is available.",
            "Profit ($ CAD)": "Profit in this year and location when available. Higher is usually better.",
            "Sales Margin %": "Profit divided by Sales Amount for the latest sales year and location. Higher is usually better.",
            "PO Cost ($ CAD)": "PO cost in this year and location after applying the brand keyword filter.",
            "Purchase Qty": "Purchased quantity in this year and location after applying the brand keyword filter.",
            "Available / Incoming / On Hand / Future Inventory": "Current stock values from today's stock report. These are not historical 2024 or 2025 inventory numbers.",
        },
    },
    ZH: {
        "Final Analysis": {
            "Product SKU": "\u7528\u4e8e\u8fde\u63a5 Sales \u548c Stock \u4e24\u4efd\u62a5\u8868\u7684\u552f\u4e00\u4ea7\u54c1\u7f16\u7801\u3002",
            "Product Name": "\u4ea7\u54c1\u540d\u79f0\uff0c\u4f18\u5148\u6765\u81ea\u9500\u552e\u62a5\u8868\u3002",
            "Qty": "\u7528\u4e8e\u5f53\u524d\u5e93\u5b58\u51b3\u7b56\u7684\u9500\u91cf\u3002\u5982\u679c\u4e0a\u4f20\u4e86\u591a\u4e2a\u5e74\u4efd sales file\uff0c\u8fd9\u91cc\u53ea\u7528\u6700\u65b0\u4e0a\u4f20\u5e74\u4efd\u7684\u9500\u91cf\u3002",
            "Sales Amount ($ CAD)": "\u5f53\u524d\u5206\u6790\u533a\u95f4\u7684\u9500\u552e\u989d\uff0c\u9700\u8981 Sales Report \u5305\u542b\u91d1\u989d\u5217\u3002",
            "Profit ($ CAD)": "\u5f53\u524d\u5206\u6790\u533a\u95f4\u7684\u5229\u6da6\uff0c\u9700\u8981 Sales Report \u5305\u542b profit \u5217\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "Sales Margin %": "Profit / Sales Amount\uff0c\u7528\u4e8e\u5224\u65ad SKU \u76c8\u5229\u6548\u7387\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "Contribution %": "\u5355\u4e2a SKU \u7684 Qty / \u5168\u90e8 SKU \u7684\u603b Qty\u3002",
            "Cumulative %": "\u6309 Qty \u4ece\u9ad8\u5230\u4f4e\u6392\u5e8f\u540e\uff0cContribution % \u7684\u7d2f\u8ba1\u503c\u3002",
            "SABC Type": "\u9500\u552e\u4f18\u5148\u7ea7\uff1aS/A \u662f\u6838\u5fc3 SKU\uff0cB/C \u662f\u4e2d\u957f\u5c3e SKU\u3002",
            "Available": "\u5e93\u5b58\u62a5\u8868\u4e2d\u7684\u5f53\u524d\u53ef\u7528\u5e93\u5b58\u3002",
            "Incoming": "\u5e93\u5b58\u62a5\u8868\u4e2d\u7684\u5728\u9014 / \u5373\u5c06\u5165\u5e93\u5e93\u5b58\u3002",
            "Future Inventory": "Available + Incoming\uff0c\u8868\u793a\u672a\u6765\u53ef\u7528\u5e93\u5b58\u9884\u4f30\u3002",
            "Adjusted Future Inventory": "MAX(0, Future Inventory)\uff0c\u8d1f\u5e93\u5b58\u6309 0 \u5904\u7406\u3002",
            "Avg Monthly Sales": "Qty / 12\u3002\u5982\u679c\u4e0a\u4f20\u591a\u5e74 sales file\uff0cQty \u4f1a\u7528\u6700\u65b0\u5e74\u4efd\uff0c\u8fd9\u4e00\u5217\u4f5c\u4e3a\u7b80\u5316\u7684 coverage \u53c2\u8003\u3002",
            "Coverage": "Adjusted Future Inventory / Avg Monthly Sales\uff0c\u8868\u793a\u5e93\u5b58\u5927\u7ea6\u8fd8\u80fd\u652f\u6301\u51e0\u4e2a\u6708\u9500\u552e\u3002",
            "Inventory Status": "\u6839\u636e\u9500\u91cf\u548c coverage \u5224\u65ad\u5e93\u5b58\u72b6\u6001\u3002",
            "Action": "\u6839\u636e\u5e93\u5b58\u72b6\u6001\u81ea\u52a8\u7ed9\u51fa\u7684\u4e1a\u52a1\u5efa\u8bae\u3002",
            "2024 PO Cost ($ CAD)": "\u6309 SKU \u5339\u914d\u5230\u8be5\u4ea7\u54c1\u7684 2024 PO line \u8fdb\u8d27\u6210\u672c\u3002",
            "2025 PO Cost ($ CAD)": "\u6309 SKU \u5339\u914d\u5230\u8be5\u4ea7\u54c1\u7684 2025 PO line \u8fdb\u8d27\u6210\u672c\u3002",
            "2026 PO Cost ($ CAD)": "\u6309 SKU \u5339\u914d\u5230\u8be5\u4ea7\u54c1\u7684 2026 YTD PO line \u8fdb\u8d27\u6210\u672c\u3002",
            "Total PO Cost ($ CAD)": "2024\u30012025\u30012026 \u4e09\u4e2a PO \u6587\u4ef6\u5339\u914d\u5230\u8be5 SKU \u7684\u8fdb\u8d27\u6210\u672c\u5408\u8ba1\u3002",
        },
        "SABC Summary": {
            "SABC Type": "\u6839\u636e\u7d2f\u8ba1\u9500\u552e\u8d21\u732e\u5f97\u5230\u7684\u9500\u552e\u4f18\u5148\u7ea7\u3002",
            "SKU_Count": "\u6bcf\u4e2a SABC \u5206\u7c7b\u4e0b\u7684 SKU \u6570\u91cf\u3002",
            "Qty": "\u8be5\u5206\u7c7b\u4e0b\u6240\u6709 SKU \u5728\u5f53\u524d\u5206\u6790\u533a\u95f4\u7684\u9500\u91cf\u5408\u8ba1\uff0c\u901a\u5e38\u662f\u6700\u65b0\u4e0a\u4f20\u5e74\u4efd\u3002",
            "Avg_Coverage": "\u8be5\u5206\u7c7b\u4e0b SKU \u7684\u5e73\u5747\u5e93\u5b58\u8986\u76d6\u6708\u6570\u3002",
            "Future_Inventory": "\u8be5\u5206\u7c7b\u4e0b\u8c03\u6574\u540e\u672a\u6765\u5e93\u5b58\u5408\u8ba1\u3002",
            "Qty Share": "\u8be5\u5206\u7c7b Qty / \u5168\u90e8 Qty\u3002",
        },
        "Inventory Status Summary": {
            "Inventory Status": "\u6839\u636e coverage \u548c\u9500\u552e\u9700\u6c42\u5f97\u5230\u7684\u5e93\u5b58\u5065\u5eb7\u72b6\u6001\u3002",
            "SKU_Count": "\u6bcf\u4e2a\u5e93\u5b58\u72b6\u6001\u4e0b\u7684 SKU \u6570\u91cf\u3002",
            "Qty": "\u8be5\u5e93\u5b58\u72b6\u6001\u4e0b SKU \u5728\u5f53\u524d\u5206\u6790\u533a\u95f4\u7684\u9500\u91cf\u5408\u8ba1\uff0c\u901a\u5e38\u662f\u6700\u65b0\u4e0a\u4f20\u5e74\u4efd\u3002",
            "Avg_Coverage": "\u8be5\u72b6\u6001\u4e0b SKU \u7684\u5e73\u5747\u5e93\u5b58\u8986\u76d6\u6708\u6570\u3002",
            "Future_Inventory": "\u8be5\u72b6\u6001\u4e0b SKU \u7684\u8c03\u6574\u540e\u672a\u6765\u5e93\u5b58\u5408\u8ba1\u3002",
            "Qty Share": "\u8be5\u72b6\u6001 Qty / \u5168\u90e8 Qty\u3002",
        },
        "Action Summary": {
            "Action": "\u5efa\u8bae\u52a8\u4f5c\uff0c\u4f8b\u5982 Replenish\u3001Monitor\u3001Reduce PO \u6216 Review SKU\u3002",
            "SKU_Count": "\u88ab\u5206\u914d\u5230\u8be5\u52a8\u4f5c\u7684 SKU \u6570\u91cf\u3002",
            "Qty": "\u8be5\u52a8\u4f5c\u4e0b SKU \u5728\u5f53\u524d\u5206\u6790\u533a\u95f4\u7684\u9500\u91cf\u5408\u8ba1\uff0c\u901a\u5e38\u662f\u6700\u65b0\u4e0a\u4f20\u5e74\u4efd\u3002",
            "Avg_Coverage": "\u8be5\u52a8\u4f5c\u4e0b SKU \u7684\u5e73\u5747\u5e93\u5b58\u8986\u76d6\u6708\u6570\u3002",
            "Future_Inventory": "\u8be5\u52a8\u4f5c\u4e0b SKU \u7684\u8c03\u6574\u540e\u672a\u6765\u5e93\u5b58\u5408\u8ba1\u3002",
            "Qty Share": "\u8be5\u52a8\u4f5c Qty / \u5168\u90e8 Qty\u3002",
        },
        "Matrix": {"Rows": "SABC \u9500\u552e\u4f18\u5148\u7ea7\u5206\u7c7b\u3002", "Columns": "\u5e93\u5b58\u72b6\u6001\u5206\u7c7b\u3002", "Values": "\u6bcf\u4e2a\u7ec4\u5408\u4e0b\u7684 SKU \u6570\u91cf\u3002"},
        "Priority": {
            "Qty": "\u5f53\u524d\u5206\u6790\u533a\u95f4\u7684\u9500\u552e\u6570\u91cf\uff0c\u901a\u5e38\u662f\u6700\u65b0\u4e0a\u4f20\u5e74\u4efd\u3002",
            "Sales Amount ($ CAD)": "\u5f53\u524d\u5206\u6790\u533a\u95f4\u7684\u9500\u552e\u989d\uff0c\u5982\u6709\u3002",
            "Profit ($ CAD)": "\u5f53\u524d\u5206\u6790\u533a\u95f4\u7684\u5229\u6da6\uff0c\u5982\u6709\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "Sales Margin %": "Profit / Sales Amount\uff0c\u5982\u6709\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "Adjusted Future Inventory": "\u5254\u9664\u8d1f\u5e93\u5b58\u540e\u7684\u53ef\u7528\u672a\u6765\u5e93\u5b58\u3002",
            "Coverage": "\u57fa\u4e8e\u5e73\u5747\u6708\u9500\u91cf\u8ba1\u7b97\u7684\u5e93\u5b58\u8986\u76d6\u6708\u6570\u3002",
            "Inventory Status": "\u5f53\u524d\u5e93\u5b58\u98ce\u9669\u6807\u7b7e\u3002",
            "Action": "\u5efa\u8bae\u4e0b\u4e00\u6b65\u4e1a\u52a1\u52a8\u4f5c\u3002",
        },
        "Detected Columns": {"Field": "\u7cfb\u7edf\u5185\u90e8\u9700\u8981\u8bc6\u522b\u7684\u5b57\u6bb5\u3002", "Detected Column": "\u4ece\u4e0a\u4f20 Excel \u4e2d\u81ea\u52a8\u5339\u914d\u5230\u7684\u5217\u540d\u3002"},
        "Sales Summary by Year": {
            "Year": "\u4e0a\u4f20\u65f6\u6807\u8bb0\u7684\u5e74\u4efd\uff0c\u6216\u5355\u4e2a sales \u6587\u4ef6\u4e2d\u4ece\u65e5\u671f\u8bc6\u522b\u7684\u5e74\u4efd\u3002",
            "SKU Count": "\u8be5\u5e74\u6709\u9500\u552e\u7684 SKU \u6570\u91cf\u3002",
            "Sales Qty": "\u8be5\u5e74\u4efd sales file \u4e2d\u7684\u9500\u552e\u6570\u91cf\u5408\u8ba1\u3002",
            "Sales Amount ($ CAD)": "\u5982 sales report \u6709\u91d1\u989d\u5217\uff0c\u5219\u4e3a\u8be5\u5e74\u9500\u552e\u989d\u3002",
            "Profit ($ CAD)": "\u5982 sales report \u6709 profit \u5217\uff0c\u5219\u4e3a\u8be5\u5e74\u5229\u6da6\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "Sales Margin %": "Profit / Sales Amount\uff0c\u8868\u793a\u8be5\u5e74\u7684\u9500\u552e\u5229\u6da6\u7387\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
        },
        "Sales vs PO by Year": {
            "Year": "\u7528\u4e8e\u8fde\u63a5 Sales file \u548c PO file \u7684\u5e74\u4efd\u3002",
            "Sales Qty": "\u8be5\u5e74\u4e0a\u4f20 sales file \u4e2d\u7684\u9500\u91cf\u3002",
            "Purchase Qty": "\u8be5\u5e74\u4e0a\u4f20 PO file \u4e2d\uff0c\u7b5b\u9009\u54c1\u724c\u548c location \u540e\u7684\u91c7\u8d2d\u6570\u91cf\u3002",
            "Sales Amount ($ CAD)": "\u5982 sales report \u6709\u91d1\u989d\u5217\uff0c\u5219\u4e3a\u8be5\u5e74\u9500\u552e\u989d\u3002",
            "Profit ($ CAD)": "\u5982 sales report \u6709 profit \u5217\uff0c\u5219\u4e3a\u8be5\u5e74\u5229\u6da6\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "Sales Margin %": "Profit / Sales Amount\uff0c\u8868\u793a\u6bcf 1 \u5757\u9500\u552e\u989d\u91cc\u6709\u591a\u5c11\u662f\u5229\u6da6\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "PO Cost ($ CAD)": "\u8be5\u5e74\u4e0a\u4f20 PO file \u4e2d\uff0c\u7b5b\u9009\u54c1\u724c\u548c location \u540e\u7684 PO line \u8fdb\u8d27\u6210\u672c\u3002",
            "PO Cost / Sales Qty ($ CAD)": "PO Cost / Sales Qty\uff0c\u8868\u793a\u6bcf\u5356\u51fa 1 \u4ef6\u5bf9\u5e94\u7684\u8fdb\u8d27\u6210\u672c\u5f3a\u5ea6\u3002\u901a\u5e38\u8d8a\u4f4e\u8d8a\u597d\uff0c\u4f46\u8981\u7ed3\u5408\u5f53\u524d\u5e93\u5b58\u548c\u662f\u5426\u9700\u8981\u8865\u8d27\u4e00\u8d77\u770b\u3002",
            "PO Cost / Sales Amount": "\u5982\u6709\u9500\u552e\u989d\uff0c\u5219\u4e3a PO Cost / Sales Amount\u3002\u901a\u5e38\u8d8a\u4f4e\u4ee3\u8868\u91c7\u8d2d\u6210\u672c\u76f8\u5bf9\u9500\u552e\u989d\u66f4\u6709\u6548\u7387\u3002",
        },
        "Purchase by SKU": {
            "SKU": "\u4ece PO line \u6587\u4ef6\u8bc6\u522b\u5230\u7684 SKU\u3002",
            "Product": "\u4ece PO line \u6587\u4ef6\u8bc6\u522b\u5230\u7684\u4ea7\u54c1\u540d\u79f0\u3002",
            "PO Cost ($ CAD)": "\u5957\u7528\u54c1\u724c\u548c location \u7b5b\u9009\u540e\u7684 PO line \u8fdb\u8d27\u6210\u672c\u5408\u8ba1\u3002",
            "Quantity": "\u5982\u6587\u4ef6\u4e2d\u6709\u6570\u91cf\u5217\uff0c\u5219\u4e3a\u91c7\u8d2d\u6570\u91cf\u5408\u8ba1\u3002",
        },
        "Latest Inventory View": {
            "Year": "\u6700\u65b0\u4e0a\u4f20\u7684 sales \u5e74\u4efd\u3002\u5f53\u524d\u5e93\u5b58\u53ea\u548c\u8fd9\u4e2a\u5e74\u4efd\u5bf9\u6bd4\u3002",
            "Location": "\u4ece Sales\u3001Stock \u6216 PO \u6587\u4ef6\u4e2d\u8bc6\u522b\u5230\u7684 location\u3002",
            "Sales Qty": "\u5982 Sales Report \u6709\u65e5\u671f\u548c location\uff0c\u5219\u4e3a\u8be5\u5e74 + location \u7684\u9500\u91cf\u3002",
            "Sales Amount ($ CAD)": "\u5982 Sales Report \u6709\u91d1\u989d\u5217\uff0c\u5219\u4e3a\u8be5\u5e74 + location \u7684\u9500\u552e\u989d\u3002",
            "Profit ($ CAD)": "\u5982 Sales Report \u6709 profit \u5217\uff0c\u5219\u4e3a\u8be5\u5e74 + location \u7684\u5229\u6da6\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "Sales Margin %": "Profit / Sales Amount\uff0c\u7528\u4e8e\u5224\u65ad\u8be5\u5e74 + location \u7684\u76c8\u5229\u6548\u7387\u3002\u901a\u5e38\u8d8a\u9ad8\u8d8a\u597d\u3002",
            "PO Cost ($ CAD)": "\u5957\u7528\u54c1\u724c\u5173\u952e\u8bcd\u7b5b\u9009\u540e\uff0c\u8be5\u5e74 + location \u7684 PO \u8fdb\u8d27\u6210\u672c\u3002",
            "Purchase Qty": "\u5957\u7528\u54c1\u724c\u5173\u952e\u8bcd\u7b5b\u9009\u540e\uff0c\u8be5\u5e74 + location \u7684\u91c7\u8d2d\u6570\u91cf\u3002",
            "Available / Incoming / On Hand / Future Inventory": "\u8fd9\u662f\u4eca\u5929 stock report \u7684\u5f53\u524d\u5e93\u5b58\u503c\uff0c\u4e0d\u662f 2024 \u6216 2025 \u7684\u5386\u53f2\u5e93\u5b58\u3002",
        },
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


def clean_filename_part(value: str) -> str:
    cleaned = "".join(ch for ch in value.strip() if ch.isalnum())
    return cleaned or "AllLocations"


def output_file_stem(brand_name: str, location_filter: str) -> str:
    return f"{clean_filename_part(clean_brand_name(brand_name))}_{clean_filename_part(location_filter)}"


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


def show_sales_po_insights(language: str, df: pd.DataFrame):
    with st.expander("Insight / \u5206\u6790\u6d1e\u5bdf", expanded=False):
        if df.empty:
            return
        work = df.sort_values("Year").copy()
        lines = []
        if len(work) >= 2:
            first = work.iloc[0]
            last = work.iloc[-1]
            sku_change = last.get("SKU Count", 0) - first.get("SKU Count", 0)
            qty_change = last.get("Sales Qty", 0) - first.get("Sales Qty", 0)
            if language == ZH:
                lines.append(f"从 {int(first['Year'])} 到 {int(last['Year'])}，SKU 数量变化 {sku_change:+,.0f}，Sales Qty 变化 {qty_change:+,.0f}。")
            else:
                lines.append(f"From {int(first['Year'])} to {int(last['Year'])}, SKU count changed by {sku_change:+,.0f} and Sales Qty changed by {qty_change:+,.0f}.")
        ratio_col = "PO Cost / Sales Qty ($ CAD)"
        if ratio_col in work.columns and work[ratio_col].notna().any():
            best = work.loc[work[ratio_col].idxmin()]
            worst = work.loc[work[ratio_col].idxmax()]
            if language == ZH:
                lines.append(f"{int(best['Year'])} 的 PO Cost / Sales Qty 最低，为 {best[ratio_col]:,.2f}，代表每卖出 1 件对应的进货成本强度最低。")
                lines.append(f"这个数值通常越低越好，但如果当前库存很低，也可能代表后续需要补货。")
            else:
                lines.append(f"{int(best['Year'])} has the lowest PO Cost / Sales Qty at {best[ratio_col]:,.2f}, meaning the lowest PO cost intensity per unit sold.")
                lines.append("Lower is usually better, but if current stock is low it may also indicate a need for replenishment.")
            if int(best["Year"]) != int(worst["Year"]) and language == ZH:
                lines.append(f"对比 {int(worst['Year'])}，{int(best['Year'])} 的采购效率更好。")
            elif int(best["Year"]) != int(worst["Year"]):
                lines.append(f"Compared with {int(worst['Year'])}, {int(best['Year'])} is more efficient on this metric.")
        if "Profit ($ CAD)" in work.columns and work["Profit ($ CAD)"].notna().any():
            top_profit = work.loc[work["Profit ($ CAD)"].idxmax()]
            if language == ZH:
                lines.append(f"{int(top_profit['Year'])} 的 Profit 最高，为 ${top_profit['Profit ($ CAD)']:,.2f} CAD。")
            else:
                lines.append(f"{int(top_profit['Year'])} has the highest Profit at ${top_profit['Profit ($ CAD)']:,.2f} CAD.")
        if "Sales Margin %" in work.columns and work["Sales Margin %"].notna().any():
            top_margin = work.loc[work["Sales Margin %"].idxmax()]
            if language == ZH:
                lines.append(f"{int(top_margin['Year'])} 的 Sales Margin % 最高，为 {top_margin['Sales Margin %']:.1%}。")
            else:
                lines.append(f"{int(top_margin['Year'])} has the strongest Sales Margin at {top_margin['Sales Margin %']:.1%}.")
        for line in lines:
            st.write(f"- {line}")


def chart_ready(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df[[col for col in columns if col in df.columns]].copy()
    for col in out.columns:
        if col != "Year" and col not in {"SABC Type", "Inventory Status", "Action"}:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    return out


def show_sales_po_charts(df: pd.DataFrame):
    trend = chart_ready(df, ["Year", "Sales Qty", "Purchase Qty", "Sales Amount ($ CAD)", "PO Cost ($ CAD)", "Profit ($ CAD)", "Sales Margin %"])
    if "Year" not in trend or trend.empty:
        return
    trend = trend.sort_values("Year").copy()
    trend["Year"] = trend["Year"].astype(int).astype(str)

    qty_cols = [col for col in ["Sales Qty", "Purchase Qty"] if col in trend.columns and trend[col].notna().any()]
    if qty_cols:
        qty_data = trend.melt("Year", value_vars=qty_cols, var_name="Metric", value_name="Qty").dropna(subset=["Qty"])
        qty_chart = (
            alt.Chart(qty_data)
            .mark_line(point=True, strokeWidth=3)
            .encode(
                x=alt.X("Year:O", title="Year", sort=trend["Year"].tolist()),
                y=alt.Y("Qty:Q", title="Qty"),
                color=alt.Color("Metric:N", title="Metric"),
                tooltip=["Year", "Metric", alt.Tooltip("Qty:Q", format=",.0f")],
            )
            .properties(title="Sales Qty vs Purchase Qty", height=300)
        )
        st.altair_chart(qty_chart, use_container_width=True)

    amount_cols = [col for col in ["Sales Amount ($ CAD)", "PO Cost ($ CAD)", "Profit ($ CAD)"] if col in trend.columns and trend[col].notna().any()]
    if "Sales Amount ($ CAD)" not in amount_cols:
        st.info("Sales Amount chart is hidden because the uploaded Sales reports do not include a sales amount column.")
    else:
        amount_data = trend.melt("Year", value_vars=amount_cols, var_name="Metric", value_name="Amount").dropna(subset=["Amount"])
        amount_chart = (
            alt.Chart(amount_data)
            .mark_line(point=True, strokeWidth=3)
            .encode(
                x=alt.X("Year:O", title="Year", sort=trend["Year"].tolist()),
                y=alt.Y("Amount:Q", title="Amount ($ CAD)"),
                color=alt.Color("Metric:N", title="Metric"),
                tooltip=["Year", "Metric", alt.Tooltip("Amount:Q", format="$,.2f")],
            )
            .properties(title="Sales Amount vs Purchase Cost and Profit", height=300)
        )
        st.altair_chart(amount_chart, use_container_width=True)

    if "Sales Margin %" in trend.columns and trend["Sales Margin %"].notna().any():
        margin_chart = (
            alt.Chart(trend.dropna(subset=["Sales Margin %"]))
            .mark_line(point=True, strokeWidth=3)
            .encode(
                x=alt.X("Year:O", title="Year", sort=trend["Year"].tolist()),
                y=alt.Y("Sales Margin %:Q", title="Sales Margin %", axis=alt.Axis(format="%")),
                color=alt.value("#D92D20"),
                tooltip=["Year", alt.Tooltip("Sales Margin %:Q", format=".1%")],
            )
            .properties(title="Sales Margin Trend", height=260)
        )
        st.altair_chart(margin_chart, use_container_width=True)


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
        st.caption(t["sales"])
        sales_2024 = st.file_uploader(t["sales_2024"], type=["xlsx", "xls"], key="sales_2024")
        sales_2025 = st.file_uploader(t["sales_2025"], type=["xlsx", "xls"], key="sales_2025")
        sales_2026 = st.file_uploader(t["sales_2026"], type=["xlsx", "xls"], key="sales_2026")
        stock_file = st.file_uploader(t["stock"], type=["xlsx", "xls"], key="stock")
        catalogue_file = st.file_uploader(t["catalogue"], type=["xlsx", "xls"], key="catalogue")
        st.caption(t["purchase"])
        purchase_2024 = st.file_uploader(t["purchase_2024"], type=["xlsx", "xls"], key="purchase_2024")
        purchase_2025 = st.file_uploader(t["purchase_2025"], type=["xlsx", "xls"], key="purchase_2025")
        purchase_2026 = st.file_uploader(t["purchase_2026"], type=["xlsx", "xls"], key="purchase_2026")
        purchase_keyword = st.text_input(t["purchase_keyword"], value="", placeholder="MILLEFEE / JUDYDOLL / JOOCYEE")
        location_filter = st.text_input(t["location_filter"], value="", placeholder="Toronto New / Vancouver", help=t["location_filter_help"])
        brand_name = st.text_input(t["brand"], value="", placeholder="MilleFee / Judydoll / Joocyee")
        generate = st.button(t["generate"], type="primary", use_container_width=True)

    sales_uploads = [(2024, sales_2024), (2025, sales_2025), (2026, sales_2026)]
    if not any(file is not None for _, file in sales_uploads) or not stock_file:
        st.info(t["missing_files"])
        return

    if generate:
        with st.spinner(t["spinner"]):
            try:
                sales_files = []
                sales_years = []
                for year, file in sales_uploads:
                    if file is not None:
                        sales_files.append(load_file(file))
                        sales_years.append(year)
                purchase_files = []
                purchase_years = []
                for year, file in [(2024, purchase_2024), (2025, purchase_2025), (2026, purchase_2026)]:
                    if file is not None:
                        purchase_files.append(load_file(file))
                        purchase_years.append(year)
                result, excel_bytes, word_bytes = generate_outputs(
                    sales_files,
                    load_file(stock_file),
                    load_file(catalogue_file),
                    purchase_files if purchase_files else None,
                    purchase_keyword.strip(),
                    purchase_years if purchase_files else None,
                    brand_name=clean_brand_name(brand_name),
                    sales_years=sales_years,
                    location_filter=location_filter.strip(),
                )
                enriched_catalogue_bytes = generate_enriched_catalogue(load_file(catalogue_file), result) if catalogue_file is not None else None
            except Exception as exc:
                st.error(f"{t['failed']}: {exc}")
                return

        st.session_state["result"] = result
        st.session_state["excel_bytes"] = excel_bytes
        st.session_state["word_bytes"] = word_bytes
        st.session_state["enriched_catalogue_bytes"] = enriched_catalogue_bytes

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
    if result.sales_purchase_year_summary is not None and not result.sales_purchase_year_summary.empty:
        st.subheader(t["sales_purchase_year_summary"])
        show_column_notes(language, "Sales vs PO by Year")
        show_sales_po_insights(language, result.sales_purchase_year_summary)
        show_sales_po_charts(result.sales_purchase_year_summary)
        compare_view = result.sales_purchase_year_summary.copy()
        if "Sales Margin %" in compare_view:
            compare_view["Sales Margin %"] = compare_view["Sales Margin %"].map(lambda x: "" if pd.isna(x) else f"{x:.1%}")
        for col in [c for c in compare_view.columns if "($ CAD)" in c or c == "PO Cost / Sales Amount"]:
            if "PO Cost / Sales Amount" == col:
                compare_view[col] = compare_view[col].map(lambda x: "" if pd.isna(x) else f"{x:.1%}")
            else:
                compare_view[col] = compare_view[col].map(lambda x: "" if pd.isna(x) else f"{x:,.2f}")
        st.dataframe(compare_view, use_container_width=True, hide_index=True)
    elif result.sales_year_summary is not None and not result.sales_year_summary.empty:
        st.subheader(t["sales_year_summary"])
        show_column_notes(language, "Sales Summary by Year")
        sales_year_view = result.sales_year_summary.copy()
        if "Sales Margin %" in sales_year_view:
            sales_year_view["Sales Margin %"] = sales_year_view["Sales Margin %"].map(lambda x: "" if pd.isna(x) else f"{x:.1%}")
        for col in [c for c in sales_year_view.columns if "($ CAD)" in c]:
            sales_year_view[col] = sales_year_view[col].map(lambda x: "" if pd.isna(x) else f"{x:,.2f}")
        st.dataframe(sales_year_view, use_container_width=True, hide_index=True)

    if result.location_year_business_view is not None and not result.location_year_business_view.empty:
        st.subheader(t["location_year_view"])
        show_column_notes(language, "Latest Inventory View")
        location_year_view = result.location_year_business_view.copy()
        if "Sales Margin %" in location_year_view:
            location_year_view["Sales Margin %"] = location_year_view["Sales Margin %"].map(lambda x: "" if pd.isna(x) else f"{x:.1%}")
        for col in [c for c in location_year_view.columns if "($ CAD)" in c]:
            location_year_view[col] = location_year_view[col].map(lambda x: "" if pd.isna(x) else f"{x:,.2f}")
        st.dataframe(location_year_view, use_container_width=True, hide_index=True)

    download_cols = st.columns(3)
    output_stem = output_file_stem(brand_name, location_filter)
    with download_cols[0]:
        st.download_button(t["download_excel"], data=st.session_state["excel_bytes"], file_name=f"{output_stem}_BP_Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with download_cols[1]:
        st.download_button(t["download_word"], data=st.session_state["word_bytes"], file_name=f"{output_stem}_Business_Analysis.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with download_cols[2]:
        if st.session_state.get("enriched_catalogue_bytes"):
            st.download_button(t["download_catalogue"], data=st.session_state["enriched_catalogue_bytes"], file_name=f"{output_stem}_Updated_Catalogue.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    tab1, tab2, tab3, tab4 = st.tabs([t["tab_final"], t["tab_summaries"], t["tab_priority"], t["tab_detected"]])
    with tab1:
        st.subheader(t["final_analysis"])
        show_column_notes(language, "Final Analysis")
        view = result.final.copy()
        for col in ["Contribution %", "Cumulative %", "Sales Margin %"]:
            if col in view:
                view[col] = view[col].map(lambda x: f"{x:.1%}" if pd.notna(x) else "")
        if "Coverage" in view:
            view["Coverage"] = view["Coverage"].map(lambda x: "" if pd.isna(x) else f"{x:,.2f}")
        for col in ["2024 PO Cost ($ CAD)", "2025 PO Cost ($ CAD)", "2026 PO Cost ($ CAD)", "Total PO Cost ($ CAD)"]:
            if col in view:
                view[col] = view[col].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
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
        priority_cols = ["Product SKU", "Product Name", "Qty", "Sales Amount ($ CAD)", "Profit ($ CAD)", "Sales Margin %", "SABC Type", "Adjusted Future Inventory", "Coverage", "Inventory Status", "Action"]
        priority_cols = [col for col in priority_cols if col in result.final.columns]
        st.subheader(t["replenishment"])
        show_column_notes(language, "Priority")
        st.dataframe(result.insights["urgent_table"][priority_cols], use_container_width=True)
        st.subheader(t["overstock_risk"])
        show_column_notes(language, "Priority")
        st.dataframe(result.insights["overstock_table"][priority_cols], use_container_width=True)
        if "top_profit_table" in result.insights and not result.insights["top_profit_table"].empty:
            st.subheader("Top Profit SKUs")
            st.dataframe(result.insights["top_profit_table"][priority_cols], use_container_width=True)
        if "low_sales_margin_table" in result.insights and not result.insights["low_sales_margin_table"].empty:
            st.subheader("Low Sales Margin Review")
            st.dataframe(result.insights["low_sales_margin_table"][priority_cols], use_container_width=True)

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
        for key in ["user", "result", "excel_bytes", "word_bytes", "enriched_catalogue_bytes"]:
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
