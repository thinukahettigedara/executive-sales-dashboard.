#!/usr/bin/env python3
"""
Executive Sales Dashboard Generator
===================================
Reads sales data from an Excel file and builds an Excel dashboard
(KPI cards + native Excel charts) whose numbers are LIVE formulas.

Usage
-----
    pip install pandas openpyxl
    python sales_dashboard.py                       # uses sales_data.xlsx (creates a sample if missing)
    python sales_dashboard.py my_sales.xlsx         # your own file
    python sales_dashboard.py my_sales.xlsx out.xlsx --sheet Sales_Data

Required columns in the input sheet (Profit is optional, = Revenue - Cost):
    Date, Order_ID, Region, Category, Product, Sales_Rep, Segment,
    Units, Unit_Price, Revenue, Cost, Profit
"""
import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
COLUMNS = ["Date", "Order_ID", "Region", "Category", "Product", "Sales_Rep",
           "Segment", "Units", "Unit_Price", "Revenue", "Cost", "Profit"]
NUMERIC = ["Units", "Unit_Price", "Revenue", "Cost", "Profit"]
TEXT = ["Order_ID", "Region", "Category", "Product", "Sales_Rep", "Segment"]
LET = {c: get_column_letter(i + 1) for i, c in enumerate(COLUMNS)}

FONT = "Arial"
NAVY, BLUE, ORANGE, GREEN, GOLD, GRAY = "1F3864", "2E75B6", "ED7D31", "70AD47", "FFC000", "F2F2F2"
PALETTE = [NAVY, BLUE, ORANGE, GREEN, GOLD, "7F7F7F", "5B9BD5", "A5A5A5"]
TOP_REPS = 8
CUR = '$#,##0'


def rng(col):
    """Whole-column reference into the Data sheet, e.g. Data!$J:$J"""
    return f"Data!${LET[col]}:${LET[col]}"


# --------------------------------------------------------------------------- #
# Sample data (only used when no input file exists)
# --------------------------------------------------------------------------- #
def make_sample_data(path, seed=42, n=900):
    rng_ = np.random.default_rng(seed)
    regions = {"North America": ["Emily Carter", "James Wilson"],
               "Europe": ["Sophie Martin", "Lukas Weber"],
               "Asia Pacific": ["Aiko Tanaka", "Raj Patel"],
               "Latin America": ["Lucia Gomez", "Mateo Silva"],
               "Middle East": ["Omar Haddad", "Layla Nasser"]}
    region_w = [0.34, 0.26, 0.22, 0.10, 0.08]
    products = {"Laptops": [("Pro Laptop 15", 1450, .68), ("Air Laptop 13", 980, .70)],
                "Accessories": [("Wireless Mouse", 45, .45), ("Mech Keyboard", 120, .50)],
                "Monitors": [("UltraWide 34", 640, .66), ("4K Monitor 27", 420, .64)],
                "Software": [("Office Suite", 160, .20), ("Security Suite", 90, .18)],
                "Services": [("Support Plan", 300, .35), ("Install Service", 220, .40)]}
    cat_w = [0.30, 0.25, 0.18, 0.15, 0.12]
    segments, seg_w = ["Enterprise", "SMB", "Consumer"], [0.35, 0.40, 0.25]

    days = pd.date_range("2025-11-01", "2026-10-30")
    w = np.linspace(1, 1.9, len(days))            # growth through the year
    w = w * (1 + 0.25 * (days.month.isin([11, 12]).astype(float)))  # holiday bump
    dates = rng_.choice(days, size=n, p=w / w.sum())

    rows = []
    for i, d in enumerate(sorted(dates)):
        region = rng_.choice(list(regions), p=region_w)
        cat = rng_.choice(list(products), p=cat_w)
        prod, price, cost_ratio = products[cat][rng_.integers(0, 2)]
        units = int(max(1, rng_.gamma(2.0, 4.0 if price > 300 else 9.0)))
        disc = rng_.choice([0, 0.05, 0.10, 0.15], p=[.45, .25, .2, .1])
        revenue = round(units * price * (1 - disc), 2)
        cost = round(units * price * cost_ratio * rng_.uniform(0.95, 1.05), 2)
        rows.append([pd.Timestamp(d).to_pydatetime(), f"ORD-{10001 + i}", region, cat, prod,
                     rng_.choice(regions[region]), rng_.choice(segments, p=seg_w),
                     units, price, revenue, cost, round(revenue - cost, 2)])
    df = pd.DataFrame(rows, columns=COLUMNS)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales_Data"
    ws.append(COLUMNS)
    for r in df.itertuples(index=False):
        ws.append(list(r))
    for c in ws[1]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=FONT, size=10)
        row[0].number_format = "yyyy-mm-dd"
        for k in (8, 9, 10, 11):
            row[k].number_format = '#,##0.00'
    for i, wdt in enumerate([12, 12, 16, 14, 18, 16, 12, 8, 12, 13, 13, 13], 1):
        ws.column_dimensions[get_column_letter(i)].width = wdt
    ws.freeze_panes = "A2"
    wb.save(path)
    print(f"Sample data created: {path}")


# --------------------------------------------------------------------------- #
# Load + validate
# --------------------------------------------------------------------------- #
def load_data(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except ValueError as e:
        sys.exit(f"Could not read sheet '{sheet}' in {path}: {e}")
    df.columns = [str(c).strip() for c in df.columns]
    if "Profit" not in df.columns and {"Revenue", "Cost"} <= set(df.columns):
        df["Profit"] = df["Revenue"] - df["Cost"]
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        sys.exit(f"Missing required column(s): {', '.join(missing)}")
    df = df[COLUMNS].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    for c in NUMERIC:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    for c in TEXT:
        df[c] = df[c].fillna("").astype(str).str.strip()
    if df.empty:
        sys.exit("No valid rows found (check the Date column).")
    return df.sort_values("Date").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Styling helpers
# --------------------------------------------------------------------------- #
def f(bold=False, size=10, color="000000"):
    return Font(name=FONT, bold=bold, size=size, color=color)


def header(ws, row, col, labels):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font = f(True, 10, "FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put(ws, row, col, value, fmt=None, bold=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = f(bold)
    if fmt:
        c.number_format = fmt
    return c


def fix_axes(ch, y_fmt=None):
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    if y_fmt:
        ch.y_axis.number_format = y_fmt
    ch.legend.position = "b"


# --------------------------------------------------------------------------- #
# Build workbook
# --------------------------------------------------------------------------- #
def build_dashboard(df, out_path, source_name):
    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    summ = wb.create_sheet("Summary")
    data = wb.create_sheet("Data")

    # ---- Data sheet (copy of the source rows) ----
    data.append(COLUMNS)
    for r in df.itertuples(index=False):
        data.append([r.Date.to_pydatetime(), r.Order_ID, r.Region, r.Category, r.Product,
                     r.Sales_Rep, r.Segment, r.Units, r.Unit_Price, r.Revenue, r.Cost, r.Profit])
    header(data, 1, 1, COLUMNS)
    for row in data.iter_rows(min_row=2):
        for c in row:
            c.font = f()
        row[0].number_format = "yyyy-mm-dd"
        for k in (8, 9, 10, 11):
            row[k].number_format = "#,##0.00"
    for i, w in enumerate([12, 12, 16, 14, 18, 16, 12, 8, 12, 13, 13, 13], 1):
        data.column_dimensions[get_column_letter(i)].width = w
    data.freeze_panes = "A2"

    # ---- Summary sheet (formula-driven tables) ----
    months = pd.date_range(df["Date"].min().replace(day=1), df["Date"].max(), freq="MS")
    regions = df.groupby("Region")["Revenue"].sum().sort_values(ascending=False).index.tolist()
    cats = df.groupby("Category")["Revenue"].sum().sort_values(ascending=False).index.tolist()
    reps = df.groupby("Sales_Rep")["Revenue"].sum().sort_values(ascending=False).head(TOP_REPS).index.tolist()

    # Monthly (A:E)
    header(summ, 1, 1, ["Month", "Revenue", "Profit", "Orders", "Margin"])
    for i, m in enumerate(months):
        r = i + 2
        crit = f'{rng("Date")},">="&$A{r},{rng("Date")},"<"&EDATE($A{r},1)'
        put(summ, r, 1, m.to_pydatetime(), "mmm yyyy")
        put(summ, r, 2, f'=SUMIFS({rng("Revenue")},{crit})', CUR)
        put(summ, r, 3, f'=SUMIFS({rng("Profit")},{crit})', CUR)
        put(summ, r, 4, f'=COUNTIFS({crit})', "#,##0")
        put(summ, r, 5, f"=IFERROR(C{r}/B{r},0)", "0.0%")
    m_last = len(months) + 1

    # Region (G:K)
    header(summ, 1, 7, ["Region", "Revenue", "Profit", "Margin", "Share"])
    for i, name in enumerate(regions):
        r = i + 2
        put(summ, r, 7, name)
        put(summ, r, 8, f'=SUMIFS({rng("Revenue")},{rng("Region")},$G{r})', CUR)
        put(summ, r, 9, f'=SUMIFS({rng("Profit")},{rng("Region")},$G{r})', CUR)
        put(summ, r, 10, f"=IFERROR(I{r}/H{r},0)", "0.0%")
        put(summ, r, 11, f"=IFERROR(H{r}/SUM(H$2:H${len(regions)+1}),0)", "0.0%")
    r_last = len(regions) + 1

    # Category (M:P)
    header(summ, 1, 13, ["Category", "Revenue", "Profit", "Margin"])
    for i, name in enumerate(cats):
        r = i + 2
        put(summ, r, 13, name)
        put(summ, r, 14, f'=SUMIFS({rng("Revenue")},{rng("Category")},$M{r})', CUR)
        put(summ, r, 15, f'=SUMIFS({rng("Profit")},{rng("Category")},$M{r})', CUR)
        put(summ, r, 16, f"=IFERROR(O{r}/N{r},0)", "0.0%")
    c_last = len(cats) + 1

    # Top reps (R:T)
    header(summ, 1, 18, ["Sales Rep", "Revenue", "Orders"])
    for i, name in enumerate(reps):
        r = i + 2
        put(summ, r, 18, name)
        put(summ, r, 19, f'=SUMIFS({rng("Revenue")},{rng("Sales_Rep")},$R{r})', CUR)
        put(summ, r, 20, f'=COUNTIFS({rng("Sales_Rep")},$R{r})', "#,##0")
    p_last = len(reps) + 1

    for col, w in {"A": 12, "B": 14, "C": 14, "D": 10, "E": 10, "F": 3, "G": 18, "H": 14,
                   "I": 14, "J": 10, "K": 10, "L": 3, "M": 16, "N": 14, "O": 14, "P": 10,
                   "Q": 3, "R": 18, "S": 14, "T": 10}.items():
        summ.column_dimensions[col].width = w
    summ.freeze_panes = "A2"

    # ---- Dashboard sheet ----
    dash.sheet_view.showGridLines = False
    dash.sheet_view.zoomScale = 90
    grey = PatternFill("solid", fgColor=GRAY)
    for row in dash.iter_rows(min_row=1, max_row=42, min_col=1, max_col=20):
        for c in row:
            c.fill = grey
    dash.column_dimensions["A"].width = 2
    dash.column_dimensions["T"].width = 2
    for i in range(2, 20):
        dash.column_dimensions[get_column_letter(i)].width = 10.5

    dash.merge_cells("B1:S2")
    t = dash["B1"]
    t.value = "EXECUTIVE SALES DASHBOARD"
    t.font = f(True, 22, "FFFFFF")
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for row in dash["B1:S2"]:
        for c in row:
            c.fill = PatternFill("solid", fgColor=NAVY)
    dash["B3"].value = (f"Period: {df['Date'].min():%d %b %Y} – {df['Date'].max():%d %b %Y}   |   "
                        f"Source: {source_name}")
    dash["B3"].font = Font(name=FONT, size=9, italic=True, color="595959")

    # KPI cards
    prev_m = m_last - 1
    mom = (f"=IFERROR(Summary!B{m_last}/Summary!B{prev_m}-1,0)" if len(months) >= 2 else "=0")
    kpis = [("TOTAL REVENUE", f'=SUM({rng("Revenue")})', CUR, NAVY),
            ("TOTAL PROFIT", f'=SUM({rng("Profit")})', CUR, BLUE),
            ("PROFIT MARGIN", "=IFERROR(E5/B5,0)", "0.0%", GREEN),
            ("ORDERS", f'=COUNT({rng("Date")})', "#,##0", ORANGE),
            ("AVG ORDER VALUE", "=IFERROR(B5/K5,0)", CUR, "7F7F7F"),
            ("MoM REVENUE GROWTH", mom, '+0.0%;-0.0%;0.0%', "C00000")]
    white = Side(style="thick", color="FFFFFF")
    for i, (label, formula, fmt, color) in enumerate(kpis):
        c1, c2 = 2 + 3 * i, 4 + 3 * i
        L1, L2 = get_column_letter(c1), get_column_letter(c2)
        dash.merge_cells(f"{L1}4:{L2}4")
        dash.merge_cells(f"{L1}5:{L2}6")
        lab, val = dash[f"{L1}4"], dash[f"{L1}5"]
        lab.value, val.value = label, formula
        lab.font = f(True, 9, "FFFFFF")
        lab.alignment = Alignment(horizontal="center", vertical="center")
        val.font = f(True, 20, "FFFFFF")
        val.number_format = fmt
        val.alignment = Alignment(horizontal="center", vertical="center")
        for row in dash[f"{L1}4:{L2}6"]:
            for c in row:
                c.fill = PatternFill("solid", fgColor=color)
                c.border = Border(left=white, right=white)
    dash.row_dimensions[4].height = 20
    dash.row_dimensions[5].height = 20
    dash.row_dimensions[6].height = 20

    # Charts ----------------------------------------------------------------
    # 1) Monthly revenue (columns) + profit (line)
    bar = BarChart()
    bar.type = "col"
    bar.title = "Monthly Revenue & Profit"
    bar.add_data(Reference(summ, min_col=2, min_row=1, max_row=m_last), titles_from_data=True)
    bar.set_categories(Reference(summ, min_col=1, min_row=2, max_row=m_last))
    bar.series[0].graphicalProperties.solidFill = BLUE
    bar.gapWidth = 60
    line = LineChart()
    line.add_data(Reference(summ, min_col=3, min_row=1, max_row=m_last), titles_from_data=True)
    s = line.series[0]
    s.graphicalProperties.line.solidFill = ORANGE
    s.graphicalProperties.line.width = 28000
    s.marker.symbol = "circle"
    s.marker.size = 6
    s.marker.graphicalProperties.solidFill = ORANGE
    s.smooth = False
    bar.x_axis.number_format = "mmm yy"
    fix_axes(bar, CUR)
    line.x_axis.delete = False
    line.y_axis.delete = False
    bar += line
    bar.width, bar.height = 25.0, 7.6
    dash.add_chart(bar, "B8")

    # 2) Revenue by category (doughnut-style pie)
    pie = PieChart()
    pie.title = "Revenue by Category"
    pie.add_data(Reference(summ, min_col=14, min_row=1, max_row=c_last), titles_from_data=True)
    pie.set_categories(Reference(summ, min_col=13, min_row=2, max_row=c_last))
    from openpyxl.chart.series import DataPoint
    for idx in range(len(cats)):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = PALETTE[idx % len(PALETTE)]
        pie.series[0].dPt.append(pt)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = False
    pie.dataLabels.showSerName = False
    pie.dataLabels.showLeaderLines = False
    pie.legend.position = "b"
    pie.width, pie.height = 12.4, 7.6
    dash.add_chart(pie, "N8")

    # 3) Revenue by region (horizontal bars)
    rb = BarChart()
    rb.type = "bar"
    rb.title = "Revenue by Region"
    rb.add_data(Reference(summ, min_col=8, min_row=1, max_row=r_last), titles_from_data=True)
    rb.set_categories(Reference(summ, min_col=7, min_row=2, max_row=r_last))
    rb.series[0].graphicalProperties.solidFill = NAVY
    rb.x_axis.scaling.orientation = "maxMin"
    rb.gapWidth = 50
    rb.legend = None
    rb.x_axis.delete = False
    rb.y_axis.delete = False
    rb.y_axis.number_format = CUR
    rb.width, rb.height = 18.6, 7.6
    dash.add_chart(rb, "B24")

    # 4) Top sales reps
    pb = BarChart()
    pb.type = "bar"
    pb.title = f"Top {len(reps)} Sales Reps by Revenue"
    pb.add_data(Reference(summ, min_col=19, min_row=1, max_row=p_last), titles_from_data=True)
    pb.set_categories(Reference(summ, min_col=18, min_row=2, max_row=p_last))
    pb.series[0].graphicalProperties.solidFill = GREEN
    pb.x_axis.scaling.orientation = "maxMin"
    pb.gapWidth = 50
    pb.legend = None
    pb.x_axis.delete = False
    pb.y_axis.delete = False
    pb.y_axis.number_format = CUR
    pb.width, pb.height = 18.6, 7.6
    dash.add_chart(pb, "K24")

    dash["B40"].value = ("All figures are live Excel formulas over the 'Data' sheet — "
                         "edit or paste new rows there and the dashboard recalculates.")
    dash["B40"].font = Font(name=FONT, size=8, italic=True, color="7F7F7F")

    dash.page_setup.orientation = "landscape"
    dash.page_setup.fitToWidth = 1
    dash.page_setup.fitToHeight = 1
    dash.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    dash.print_area = "A1:T42"

    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(out_path)
    print(f"Dashboard created: {out_path}")


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Build an executive sales dashboard in Excel.")
    ap.add_argument("input", nargs="?", default="sales_data.xlsx", help="Excel file with sales data")
    ap.add_argument("output", nargs="?", default="Sales_Dashboard.xlsx", help="Dashboard file to create")
    ap.add_argument("--sheet", default="Sales_Data", help="Sheet name that holds the data")
    a = ap.parse_args()

    if not Path(a.input).exists():
        print(f"'{a.input}' not found - generating sample data.")
        make_sample_data(a.input)
    df = load_data(a.input, a.sheet)
    build_dashboard(df, a.output, Path(a.input).name)


if __name__ == "__main__":
    main()
