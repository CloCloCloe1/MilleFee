from __future__ import annotations

import math
import re
import tempfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from typing import BinaryIO, Iterable

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image, ImageDraw, ImageFont


ACCENT = "1D1D1F"
BLUE = "0071E3"
LIGHT_BLUE = "EAF4FF"
GREEN = "1F8A5B"
YELLOW = "F2A900"
RED = "D92D20"
GRAY = "F5F5F7"
MID_GRAY = "86868B"


@dataclass
class AnalysisResult:
    final: pd.DataFrame
    sabc_summary: pd.DataFrame
    inventory_status_summary: pd.DataFrame
    action_summary: pd.DataFrame
    matrix: pd.DataFrame
    insights: dict
    detected_columns: dict


def normalize_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def normalize_sku(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        if math.isnan(float(value)):
            return ""
        return str(int(round(float(value)))) if float(value).is_integer() else f"{float(value):.0f}"

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    text = text.replace("\u3000", " ").strip()
    text = re.sub(r"\.0$", "", text)
    if re.fullmatch(r"[0-9]+", text):
        return text
    if re.fullmatch(r"[0-9]+(\.[0-9]+)?[eE][+-]?[0-9]+", text) or re.fullmatch(r"[0-9]+\.[0-9]+", text):
        try:
            return str(int(Decimal(text).to_integral_value()))
        except (InvalidOperation, ValueError):
            return text
    return text


def first_existing_column(columns: Iterable[str], aliases: Iterable[str]) -> str | None:
    normalized = {normalize_header(c): c for c in columns}
    for alias in aliases:
        hit = normalized.get(normalize_header(alias))
        if hit is not None:
            return hit
    return None


def detect_column(df: pd.DataFrame, aliases: list[str], required: bool = True, label: str = "") -> str | None:
    hit = first_existing_column(df.columns, aliases)
    if hit is None and required:
        pretty = label or aliases[0]
        raise ValueError(f"Could not detect required column: {pretty}. Available columns: {list(df.columns)}")
    return hit


def read_excel_any(file: str | Path | BinaryIO | BytesIO) -> pd.DataFrame:
    xl = pd.ExcelFile(file)
    frames = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet)
        if not df.empty:
            frames.append(df)
    if not frames:
        raise ValueError("The uploaded Excel file does not contain readable rows.")
    return pd.concat(frames, ignore_index=True)


def latest_12_months(df: pd.DataFrame, date_col: str | None) -> pd.DataFrame:
    if date_col is None:
        return df.copy()
    work = df.copy()
    work[date_col] = pd.to_datetime(work[date_col], errors="coerce")
    valid_dates = work[date_col].dropna()
    if valid_dates.empty:
        return work
    end_month = valid_dates.max().to_period("M").to_timestamp()
    start_month = (end_month.to_period("M") - 11).to_timestamp()
    month_values = work[date_col].dt.to_period("M").dt.to_timestamp()
    return work[(month_values >= start_month) & (month_values <= end_month)].copy()


def mode_or_first(series: pd.Series) -> str:
    cleaned = series.dropna().astype(str).str.strip()
    cleaned = cleaned[cleaned != ""]
    if cleaned.empty:
        return ""
    mode = cleaned.mode()
    return mode.iloc[0] if not mode.empty else cleaned.iloc[0]


def aggregate_sales(sales_df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    sku_col = detect_column(sales_df, ["Product SKU", "SKU", "Product Code", "Item SKU", "Barcode"], label="Product SKU")
    name_col = detect_column(sales_df, ["Product / Service", "Product Name", "Name", "Item Name", "Description"], label="Product Name")
    qty_col = detect_column(sales_df, ["Quantity", "Qty", "Units", "Units Sold", "Sold Qty"], label="Quantity")
    date_col = detect_column(sales_df, ["Month, Year", "Month/Year", "Month Year", "Date", "Order Date"], required=False)

    work = latest_12_months(sales_df, date_col)
    work["_sku"] = work[sku_col].map(normalize_sku)
    work["_qty"] = pd.to_numeric(work[qty_col], errors="coerce").fillna(0)
    work = work[work["_sku"] != ""]
    grouped = (
        work.groupby("_sku", as_index=False)
        .agg(**{"Product SKU": ("_sku", "first"), "Product Name": (name_col, mode_or_first), "Qty": ("_qty", "sum")})
        .drop(columns=["_sku"], errors="ignore")
        .sort_values("Qty", ascending=False)
    )
    return grouped, {"sales_sku": sku_col, "sales_name": name_col, "sales_qty": qty_col, "sales_date": date_col or "Not detected"}


def aggregate_stock(stock_df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    sku_col = detect_column(stock_df, ["Product SKU", "SKU", "Product Code", "Item SKU", "Barcode"], label="SKU")
    name_col = detect_column(stock_df, ["Product / Service", "Product Name", "Name", "Item Name", "Description"], required=False)
    available_col = detect_column(stock_df, ["Available", "Available Qty", "Available Stock"], required=False)
    incoming_col = detect_column(stock_df, ["Incoming", "Incoming Qty", "On Order"], required=False)
    on_hand_col = detect_column(stock_df, ["On Hand", "On hand", "Onhand", "Stock On Hand"], required=False)

    work = stock_df.copy()
    work["_sku"] = work[sku_col].map(normalize_sku)
    work = work[work["_sku"] != ""]
    for col in [available_col, incoming_col, on_hand_col]:
        if col:
            work[col] = pd.to_numeric(work[col], errors="coerce").fillna(0)

    aggregations = {"Product SKU": ("_sku", "first")}
    if name_col:
        aggregations["Stock Product Name"] = (name_col, mode_or_first)
    aggregations["Available"] = (available_col, "sum") if available_col else ("_sku", lambda _: 0)
    aggregations["Incoming"] = (incoming_col, "sum") if incoming_col else ("_sku", lambda _: 0)
    aggregations["On Hand"] = (on_hand_col, "sum") if on_hand_col else ("_sku", lambda _: 0)
    grouped = work.groupby("_sku", as_index=False).agg(**aggregations).drop(columns=["_sku"], errors="ignore")

    return grouped, {
        "stock_sku": sku_col,
        "stock_name": name_col or "Not detected",
        "stock_available": available_col or "Not detected",
        "stock_incoming": incoming_col or "Not detected",
        "stock_on_hand": on_hand_col or "Not detected",
    }


def apply_catalogue(final: pd.DataFrame, catalogue_df: pd.DataFrame | None) -> tuple[pd.DataFrame, dict]:
    if catalogue_df is None:
        return final, {}
    sku_col = detect_column(catalogue_df, ["Product SKU", "SKU", "Product Code", "Code", "Barcode"], label="Catalogue SKU")
    cat = catalogue_df.copy()
    cat["_sku"] = cat[sku_col].map(normalize_sku)
    optional_map = {
        "Category": ["Category", "Product Category", "Collection"],
        "Recommendation / Lifecycle Status": ["Recommendation / Lifecycle Status", "Lifecycle Status", "Recommendation", "Status"],
        "RP USD": ["RP USD", "Retail Price", "Retail USD", "MSRP"],
        "Wholesale Price": ["Wholesale Price", "Wholesale", "WSP", "Cost"],
        "Inner Case": ["Inner Case", "Inner", "Inner Qty"],
        "Outer Case": ["Outer Case", "Outer", "Outer Qty"],
    }
    detected = {"catalogue_sku": sku_col}
    keep = ["_sku"]
    for out_col, aliases in optional_map.items():
        source = detect_column(cat, aliases, required=False)
        detected[f"catalogue_{out_col}"] = source or "Not detected"
        if source:
            cat[out_col] = cat[source]
            keep.append(out_col)
    cat = cat[keep].drop_duplicates("_sku")
    merged = final.merge(cat, left_on="Product SKU", right_on="_sku", how="left").drop(columns=["_sku"], errors="ignore")
    if {"RP USD", "Wholesale Price"}.issubset(merged.columns):
        rp = pd.to_numeric(merged["RP USD"], errors="coerce")
        wholesale = pd.to_numeric(merged["Wholesale Price"], errors="coerce")
        merged["Margin %"] = np.where(rp > 0, (rp - wholesale) / rp, np.nan)
    if "Outer Case" in merged.columns:
        outer = pd.to_numeric(merged["Outer Case"], errors="coerce")
        merged["MOQ risk"] = np.where((merged["Avg Monthly Sales"] > 0) & (outer > merged["Avg Monthly Sales"] * 3), "High MOQ Risk", "")
    return merged, detected


def build_analysis(
    sales_file: str | Path | BinaryIO | BytesIO,
    stock_file: str | Path | BinaryIO | BytesIO,
    catalogue_file: str | Path | BinaryIO | BytesIO | None = None,
) -> AnalysisResult:
    sales_df = read_excel_any(sales_file)
    stock_df = read_excel_any(stock_file)
    catalogue_df = read_excel_any(catalogue_file) if catalogue_file else None

    sales, sales_cols = aggregate_sales(sales_df)
    stock, stock_cols = aggregate_stock(stock_df)

    final = sales.merge(stock, on="Product SKU", how="outer")
    final["Qty"] = pd.to_numeric(final["Qty"], errors="coerce").fillna(0)
    final["Product Name"] = final["Product Name"].fillna(final.get("Stock Product Name", "")).fillna("")
    final = final.drop(columns=["Stock Product Name"], errors="ignore")
    for col in ["Available", "Incoming", "On Hand"]:
        final[col] = pd.to_numeric(final.get(col, 0), errors="coerce").fillna(0)

    final = final.sort_values("Qty", ascending=False).reset_index(drop=True)
    total_qty = final["Qty"].sum()
    final["Contribution %"] = np.where(total_qty > 0, final["Qty"] / total_qty, 0)
    final["Cumulative %"] = final["Contribution %"].cumsum()
    final["SABC Type"] = np.select(
        [
            final["Cumulative %"] <= 0.05,
            final["Cumulative %"] <= 0.80,
            final["Cumulative %"] <= 0.95,
        ],
        ["S", "A", "B"],
        default="C",
    )
    final.loc[final["Qty"] <= 0, "SABC Type"] = "C"

    final["Future Inventory"] = final["Available"] + final["Incoming"]
    final["Adjusted Future Inventory"] = final["Future Inventory"].clip(lower=0)
    final["Avg Monthly Sales"] = final["Qty"] / 12
    final["Coverage"] = np.where(final["Avg Monthly Sales"] > 0, final["Adjusted Future Inventory"] / final["Avg Monthly Sales"], np.nan)

    conditions = [
        final["Avg Monthly Sales"] <= 0,
        (final["Adjusted Future Inventory"] <= 0) & (final["Avg Monthly Sales"] > 0),
        final["Coverage"] < 1,
        final["Coverage"] < 3,
        final["Coverage"] < 6,
        final["Coverage"] >= 6,
    ]
    final["Inventory Status"] = np.select(
        conditions,
        ["No Sales / Review", "Stockout", "Urgent", "Healthy", "Monitor", "Overstock"],
        default="No Sales / Review",
    )
    action_map = {
        "Stockout": "Immediate Replenishment",
        "Urgent": "Replenish",
        "Healthy": "Monitor",
        "Monitor": "Review PO",
        "Overstock": "Reduce PO",
        "No Sales / Review": "Review SKU",
    }
    final["Action"] = final["Inventory Status"].map(action_map)
    final, catalogue_cols = apply_catalogue(final, catalogue_df)

    base_cols = [
        "Product SKU",
        "Product Name",
        "Qty",
        "Contribution %",
        "Cumulative %",
        "SABC Type",
        "Available",
        "Incoming",
        "Future Inventory",
        "Adjusted Future Inventory",
        "Avg Monthly Sales",
        "Coverage",
        "Inventory Status",
        "Action",
    ]
    extra_cols = [c for c in final.columns if c not in base_cols + ["On Hand"]]
    final = final[base_cols + extra_cols]

    sabc_summary = summarize(final, "SABC Type")
    inventory_status_summary = summarize(final, "Inventory Status")
    action_summary = summarize(final, "Action")
    matrix = pd.crosstab(final["SABC Type"], final["Inventory Status"], values=final["Product SKU"], aggfunc="count").fillna(0).astype(int)
    matrix = matrix.reindex(index=["S", "A", "B", "C"], fill_value=0)

    insights = build_insights(final, sabc_summary, inventory_status_summary, action_summary)
    return AnalysisResult(final, sabc_summary, inventory_status_summary, action_summary, matrix, insights, {**sales_cols, **stock_cols, **catalogue_cols})


def summarize(final: pd.DataFrame, group_col: str) -> pd.DataFrame:
    out = (
        final.groupby(group_col, dropna=False)
        .agg(
            SKU_Count=("Product SKU", "count"),
            Qty=("Qty", "sum"),
            Avg_Coverage=("Coverage", "mean"),
            Future_Inventory=("Adjusted Future Inventory", "sum"),
        )
        .reset_index()
    )
    total_qty = final["Qty"].sum()
    out["Qty Share"] = np.where(total_qty > 0, out["Qty"] / total_qty, 0)
    return out


def build_insights(final: pd.DataFrame, sabc: pd.DataFrame, status: pd.DataFrame, action: pd.DataFrame) -> dict:
    total_skus = len(final)
    total_qty = float(final["Qty"].sum())
    top_sku = final.iloc[0] if total_skus else None
    urgent = final[final["Inventory Status"].isin(["Stockout", "Urgent"])].sort_values(["SABC Type", "Qty"], ascending=[True, False])
    overstock = final[final["Inventory Status"] == "Overstock"].sort_values("Adjusted Future Inventory", ascending=False)
    core = final[final["SABC Type"].isin(["S", "A"])]
    no_sales = final[final["Inventory Status"] == "No Sales / Review"]
    return {
        "total_skus": total_skus,
        "total_qty": total_qty,
        "top_sku": "" if top_sku is None else top_sku["Product SKU"],
        "top_product": "" if top_sku is None else top_sku["Product Name"],
        "top_qty": 0 if top_sku is None else float(top_sku["Qty"]),
        "top_share": 0 if total_qty == 0 or top_sku is None else float(top_sku["Qty"] / total_qty),
        "urgent_count": len(urgent),
        "overstock_count": len(overstock),
        "no_sales_count": len(no_sales),
        "core_sku_count": len(core),
        "core_qty_share": 0 if total_qty == 0 else float(core["Qty"].sum() / total_qty),
        "urgent_table": urgent.head(10),
        "overstock_table": overstock.head(10),
    }


def dataframe_to_rows(df: pd.DataFrame) -> list[list[object]]:
    clean = df.copy()
    clean = clean.replace([np.inf, -np.inf], np.nan)
    clean = clean.where(pd.notna(clean), None)
    return [list(clean.columns)] + clean.values.tolist()


def write_df(ws, df: pd.DataFrame, table_name: str, percent_cols: Iterable[str] = ()):
    rows = dataframe_to_rows(df)
    for row in rows:
        ws.append(row)
    max_row, max_col = ws.max_row, ws.max_column
    ref = f"A1:{get_column_letter(max_col)}{max_row}"
    table = Table(displayName=table_name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)
    style_sheet(ws, percent_cols)


def style_sheet(ws, percent_cols: Iterable[str] = ()):
    header_fill = PatternFill("solid", fgColor=ACCENT)
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D2D2D7")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    percent_set = set(percent_cols)
    for col_idx, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
        header = col_cells[0].value
        width = min(max(len(str(header or "")) + 2, 12), 32)
        for cell in col_cells[1:]:
            if isinstance(cell.value, str):
                width = min(max(width, len(cell.value[:40]) + 2), 42)
            if header in percent_set:
                cell.number_format = "0.0%"
            elif isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00" if header in {"Coverage", "Avg Monthly Sales", "Avg_Coverage"} else "#,##0"
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def add_excel_dashboard_charts(wb: Workbook):
    ws = wb["Final Analysis"]
    status_ws = wb["Inventory Status Summary"]
    sabc_ws = wb["SABC Summary"]
    if status_ws.max_row > 1:
        chart = BarChart()
        chart.title = "Inventory Status by SKU Count"
        chart.y_axis.title = "SKU Count"
        chart.x_axis.title = "Status"
        data = Reference(status_ws, min_col=2, min_row=1, max_row=status_ws.max_row)
        cats = Reference(status_ws, min_col=1, min_row=2, max_row=status_ws.max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 7
        chart.width = 12
        ws.add_chart(chart, "P2")
    if sabc_ws.max_row > 1:
        pie = PieChart()
        pie.title = "Qty Share by SABC"
        data = Reference(sabc_ws, min_col=6, min_row=1, max_row=sabc_ws.max_row)
        cats = Reference(sabc_ws, min_col=1, min_row=2, max_row=sabc_ws.max_row)
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)
        pie.height = 7
        pie.width = 10
        ws.add_chart(pie, "P18")


def generate_excel(result: AnalysisResult) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    sheets = [
        ("Final Analysis", result.final, "FinalAnalysisTable", ["Contribution %", "Cumulative %", "Margin %"]),
        ("SABC Summary", result.sabc_summary, "SABCSummaryTable", ["Qty Share"]),
        ("Inventory Status Summary", result.inventory_status_summary, "InventoryStatusSummaryTable", ["Qty Share"]),
        ("Action Summary", result.action_summary, "ActionSummaryTable", ["Qty Share"]),
        ("SABC x Inventory Status Matrix", result.matrix.reset_index(), "SABCMatrixTable", []),
    ]
    for name, df, table_name, percent_cols in sheets:
        ws = wb.create_sheet(name)
        write_df(ws, df, table_name, percent_cols)

    ws = wb["Final Analysis"]
    headers = [cell.value for cell in ws[1]]
    if "Inventory Status" in headers:
        col = get_column_letter(headers.index("Inventory Status") + 1)
        rng = f"{col}2:{col}{ws.max_row}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Stockout"'], fill=PatternFill("solid", fgColor="FCE4E4")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Urgent"'], fill=PatternFill("solid", fgColor="FFF4CC")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Overstock"'], fill=PatternFill("solid", fgColor="EAF4FF")))
    add_excel_dashboard_charts(wb)
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def try_font(size: int, bold: bool = False):
    candidates = [
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def make_bar_chart_png(df: pd.DataFrame, label_col: str, value_col: str, title: str) -> str:
    data = df[[label_col, value_col]].copy().head(8)
    data[value_col] = pd.to_numeric(data[value_col], errors="coerce").fillna(0)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    width, height = 900, 420
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    title_font = try_font(24, True)
    label_font = try_font(15)
    small_font = try_font(13)
    draw.text((32, 24), title, fill=f"#{ACCENT}", font=title_font)
    max_value = max(float(data[value_col].max()), 1.0)
    top, left, bar_h, gap = 82, 240, 30, 16
    for i, (_, row) in enumerate(data.iterrows()):
        y = top + i * (bar_h + gap)
        label = str(row[label_col])[:26]
        value = float(row[value_col])
        draw.text((32, y + 5), label, fill=f"#{ACCENT}", font=label_font)
        draw.rounded_rectangle((left, y, width - 90, y + bar_h), radius=8, fill=f"#{GRAY}")
        bar_w = int((width - 90 - left) * value / max_value)
        draw.rounded_rectangle((left, y, left + bar_w, y + bar_h), radius=8, fill=f"#{BLUE}")
        draw.text((left + bar_w + 10, y + 6), f"{value:,.0f}", fill=f"#{ACCENT}", font=small_font)
    img.save(tmp.name, "PNG")
    return tmp.name


def make_sabc_chart_png(df: pd.DataFrame) -> str:
    return make_bar_chart_png(df.sort_values("Qty", ascending=False), "SABC Type", "Qty", "Sales Quantity by SABC Type")


def set_doc_styles(doc: Document):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(6)
    for style_name, size in [("Heading 1", 16), ("Heading 2", 12)]:
        style = styles[style_name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(29, 29, 31)
        style.paragraph_format.space_before = Pt(14)
        style.paragraph_format.space_after = Pt(6)


def add_table(doc: Document, df: pd.DataFrame, max_rows: int = 10):
    clean = df.head(max_rows).copy().replace([np.inf, -np.inf], np.nan)
    clean = clean.where(pd.notna(clean), "")
    table = doc.add_table(rows=1, cols=len(clean.columns))
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    for i, col in enumerate(clean.columns):
        hdr[i].text = str(col)
        hdr[i].paragraphs[0].runs[0].font.bold = True
        hdr[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        hdr[i]._tc.get_or_add_tcPr().append(_cell_shading(ACCENT))
    for _, row in clean.iterrows():
        cells = table.add_row().cells
        for i, value in enumerate(row):
            if isinstance(value, float):
                text = f"{value:.1%}" if "Share" in str(clean.columns[i]) or "%" in str(clean.columns[i]) else f"{value:,.2f}"
            else:
                text = str(value)
            cells[i].text = text
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    doc.add_paragraph()


def _cell_shading(fill: str):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    return shd


def add_bullet(doc: Document, text: str):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text)


def add_key_value_paragraph(doc: Document, label: str, value: str):
    p = doc.add_paragraph()
    run = p.add_run(f"{label}: ")
    run.bold = True
    p.add_run(value)


def generate_word_report(result: AnalysisResult, brand_name: str = "MilleFee") -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    set_doc_styles(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = title.add_run(f"{brand_name} Business Analysis & 6-Month Plan")
    run.font.name = "Arial"
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = RGBColor(29, 29, 31)
    subtitle = doc.add_paragraph("Auto-generated from Sales Report and Stock Levels Report")
    subtitle.runs[0].font.color.rgb = RGBColor(134, 134, 139)

    insights = result.insights
    doc.add_heading("Executive Summary", level=1)
    add_bullet(doc, f"Analyzed {insights['total_skus']:,} SKUs with {insights['total_qty']:,.0f} units sold in the latest available 12-month window.")
    add_bullet(doc, f"Top SKU: {insights['top_sku']} - {insights['top_product']} ({insights['top_qty']:,.0f} units, {insights['top_share']:.1%} of sales).")
    add_bullet(doc, f"S/A core SKUs represent {insights['core_sku_count']:,} SKUs and {insights['core_qty_share']:.1%} of quantity sold.")
    add_bullet(doc, f"{insights['urgent_count']:,} SKUs need immediate replenishment attention; {insights['overstock_count']:,} SKUs show overstock risk.")

    doc.add_heading("Data Sources", level=1)
    add_key_value_paragraph(doc, "Sales Report", "SKU-level quantity sold, product name, and month/year when available.")
    add_key_value_paragraph(doc, "Stock Levels Report", "SKU-level available stock, incoming stock, and on-hand inventory when available.")
    if any(k.startswith("catalogue_") for k in result.detected_columns):
        add_key_value_paragraph(doc, "Optional Catalogue", "Catalogue fields were linked by SKU and added to the final analysis.")

    doc.add_heading("Excel Column Explanation", level=1)
    explanation = pd.DataFrame(
        [
            ["Qty", "Total SKU quantity sold in latest 12 months"],
            ["Contribution %", "SKU Qty / Total Qty"],
            ["Cumulative %", "Running contribution after sorting by Qty"],
            ["SABC Type", "S <= 5%, A <= 80%, B <= 95%, C > 95%"],
            ["Future Inventory", "Available + Incoming"],
            ["Adjusted Future Inventory", "MAX(0, Future Inventory)"],
            ["Avg Monthly Sales", "Qty / 12"],
            ["Coverage", "Adjusted Future Inventory / Avg Monthly Sales"],
            ["Inventory Status", "Stockout, Urgent, Healthy, Monitor, Overstock, or No Sales / Review"],
            ["Action", "Operational recommendation linked to inventory status"],
        ],
        columns=["Column", "Meaning"],
    )
    add_table(doc, explanation, max_rows=20)

    doc.add_heading("SABC Sales Analysis", level=1)
    doc.add_paragraph("The SABC view separates high-impact SKUs from long-tail products, so replenishment can protect sales while keeping slow movers controlled.")
    add_table(doc, result.sabc_summary, max_rows=10)
    sabc_png = make_sabc_chart_png(result.sabc_summary)
    doc.add_picture(sabc_png, width=Inches(6.5))

    doc.add_heading("Inventory Coverage Analysis", level=1)
    doc.add_paragraph("Coverage translates stock into months of demand. Low coverage creates service risk; excessive coverage creates cash and warehouse pressure.")
    add_table(doc, result.inventory_status_summary, max_rows=10)
    status_png = make_bar_chart_png(result.inventory_status_summary, "Inventory Status", "SKU_Count", "Inventory Status by SKU Count")
    doc.add_picture(status_png, width=Inches(6.5))

    doc.add_heading("Replenishment Priority", level=1)
    priority_cols = ["Product SKU", "Product Name", "Qty", "SABC Type", "Adjusted Future Inventory", "Coverage", "Inventory Status", "Action"]
    add_table(doc, insights["urgent_table"][priority_cols], max_rows=10)

    doc.add_heading("Overstock Risk", level=1)
    add_table(doc, insights["overstock_table"][priority_cols], max_rows=10)

    doc.add_heading("Future 6-Month Business Plan", level=1)
    add_bullet(doc, "Month 1-2: protect S/A SKUs, solve Stockout and Urgent items first, and validate incoming PO timing.")
    add_bullet(doc, "Month 2-3: reduce or pause PO for Overstock SKUs, especially C-type long-tail products.")
    add_bullet(doc, "Month 3-4: build channel-specific bundles around best-selling product lines and display sets.")
    add_bullet(doc, "Month 4-6: review SKU lifecycle, discontinue low-sales SKUs where sell-through does not improve, and reallocate budget to hero SKUs.")

    doc.add_heading("Label / Packaging Optimization", level=1)
    add_bullet(doc, "Prioritize bilingual label clarity for S/A products first, because these SKUs carry the highest sales impact.")
    add_bullet(doc, "Standardize shade names, claims, and barcode visibility to reduce store and warehouse handling friction.")

    doc.add_heading("Channel Growth Plan", level=1)
    add_bullet(doc, "Use S/A SKUs as traffic drivers for marketplace, boutique retail, and social commerce campaigns.")
    add_bullet(doc, "Use B SKUs for curated sets and seasonal campaigns; keep C SKUs under tight inventory review.")

    doc.add_heading("New Product Plan", level=1)
    add_bullet(doc, "Expand adjacent shades or formats only where the current line has proven sell-through.")
    add_bullet(doc, "Require catalogue pricing, MOQ, inner/outer case data, and expected launch channel before committing to new SKUs.")

    doc.add_heading("Manufacturer Support Needed", level=1)
    add_bullet(doc, "Confirm replenishment lead time, MOQ flexibility, tester/display support, and launch assets for priority SKUs.")
    add_bullet(doc, "Request packaging files and ingredient/claim documentation for faster Canadian channel onboarding.")

    doc.add_heading("Key Business Conclusion", level=1)
    doc.add_paragraph(
        f"{brand_name} should manage the next 6 months with a focused SKU strategy: protect S/A winners, fix urgent replenishment gaps, "
        "and actively reduce long-tail overstock exposure. The business plan should stay data-led, with monthly updates as sales and stock reports refresh."
    )

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()


def generate_outputs(
    sales_file: str | Path | BinaryIO | BytesIO,
    stock_file: str | Path | BinaryIO | BytesIO,
    catalogue_file: str | Path | BinaryIO | BytesIO | None = None,
    brand_name: str = "MilleFee",
) -> tuple[AnalysisResult, bytes, bytes]:
    result = build_analysis(sales_file, stock_file, catalogue_file)
    return result, generate_excel(result), generate_word_report(result, brand_name=brand_name)
