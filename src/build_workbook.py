#!/usr/bin/env python3
"""P23 — build the deliverable workbook for StyleHub.pk (SAMPLE):

  Dashboard              — KPIs (campaigns, spend, revenue, overall
                           ROAS, traffic, CTR / CPC / CPA / conversion
                           rate, revenue per click, gross profit at
                           45% margin, blended payback), best / worst
                           campaign and channel, Kill / Scale / Keep
                           counts — all live formulas over Campaigns
  Campaigns              — all 12 campaigns (values for the source
                           columns, live formulas for days, CTR, CPC,
                           CPA, conversion rate, ROAS, revenue per
                           click, gross profit, payback days and the
                           Kill / Scale / Keep call) plus the monthly
                           day-share split (spend / revenue /
                           conversions per month, live formulas)
  Channels               — 7 channels: campaigns, spend, traffic,
                           revenue, all rates, gross profit, campaign
                           days, payback — COUNTIF / SUMIF over
                           Campaigns
  Monthly Trend          — 6 months (day-share split): campaigns
                           active, spend, revenue, conversions, ROAS
                           + totals + chart
  Funnel                 — impressions -> clicks -> conversions with
                           conversion rates + per-channel funnel
  Method                 — how to read it, every formula in plain
                           words, the 45% margin assumption, the call
                           rule, their limits, how to re-verify

    python3 src/build_workbook.py
"""
import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"
OUT = PROJECT / "outputs" / "campaign_roi_report_SAMPLE.xlsx"

CHANNELS = ["Google Search", "Google Shopping", "Meta Ads",
            "TikTok Ads", "Influencer", "Email", "SMS/WhatsApp"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
# month: (first day, last day) — 2026 is not a leap year
MONTH_BOUNDS = {
    "2026-01": (date(2026, 1, 1), date(2026, 1, 31)),
    "2026-02": (date(2026, 2, 1), date(2026, 2, 28)),
    "2026-03": (date(2026, 3, 1), date(2026, 3, 31)),
    "2026-04": (date(2026, 4, 1), date(2026, 4, 30)),
    "2026-05": (date(2026, 5, 1), date(2026, 5, 31)),
    "2026-06": (date(2026, 6, 1), date(2026, 6, 30)),
}
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
COUNT_FMT = "#,##0"
PKR_FMT = "#,##0"
PCT_FMT = "0.00%"
AVG_FMT = "#,##0.00"
ROAS_FMT = "0.00"
DAYS_FMT = "0.0"
DATE_FMT = "yyyy-mm-dd"


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT


def dstr(d):
    return f"DATE({d.year},{d.month},{d.day})"


def main():
    raw = list(csv.DictReader(
        open(RAW / "campaigns_SAMPLE.csv", encoding="utf-8")))
    raw.sort(key=lambda r: (r["start_date"], r["campaign_id"]))
    n = len(raw)
    last = n + 1  # last data row on Campaigns (row 1 = header)

    wb = Workbook()

    # ── Campaigns ───────────────────────────────────────────────────
    cp = wb.active
    cp.title = "Campaigns"
    cp.append(["campaign_id", "campaign_name", "channel",
               "start_date", "end_date", "days", "spend_pkr",
               "impressions", "clicks", "conversions", "revenue_pkr",
               "ctr", "cpc_pkr", "cpa_pkr", "conversion_rate",
               "roas", "revenue_per_click_pkr", "gross_profit_pkr",
               "payback_days", "call",
               "spend_2026-01", "spend_2026-02", "spend_2026-03",
               "spend_2026-04", "spend_2026-05", "spend_2026-06",
               "revenue_2026-01", "revenue_2026-02",
               "revenue_2026-03", "revenue_2026-04",
               "revenue_2026-05", "revenue_2026-06",
               "conversions_2026-01", "conversions_2026-02",
               "conversions_2026-03", "conversions_2026-04",
               "conversions_2026-05", "conversions_2026-06"])
    for r in raw:
        cp.append([r["campaign_id"], r["campaign_name"], r["channel"],
                   date.fromisoformat(r["start_date"]),
                   date.fromisoformat(r["end_date"]),
                   None, int(r["spend_pkr"]), int(r["impressions"]),
                   int(r["clicks"]), int(r["conversions"]),
                   int(r["revenue_pkr"])] + [None] * 27)
    for i in range(2, last + 1):
        cp.cell(row=i, column=6).value = f"=E{i}-D{i}+1"
        cp.cell(row=i, column=12).value = f'=IF(H{i}=0,"—",I{i}/H{i})'
        cp.cell(row=i, column=13).value = f'=IF(I{i}=0,"—",G{i}/I{i})'
        cp.cell(row=i, column=14).value = f'=IF(J{i}=0,"—",G{i}/J{i})'
        cp.cell(row=i, column=15).value = f'=IF(I{i}=0,"—",J{i}/I{i})'
        cp.cell(row=i, column=16).value = f'=IF(G{i}=0,"—",K{i}/G{i})'
        cp.cell(row=i, column=17).value = f'=IF(I{i}=0,"—",K{i}/I{i})'
        cp.cell(row=i, column=18).value = f"=K{i}*0.45"
        cp.cell(row=i, column=19).value = (
            f'=IF(R{i}=0,"—",G{i}/(R{i}/F{i}))')
        cp.cell(row=i, column=20).value = (
            f'=IF(P{i}<1,"Kill",IF(AND(P{i}>=3,G{i}<900000),"Scale",'
            f'IF(P{i}>=3,"Keep — proven at scale","Keep / Watch")))')
        for k, mo in enumerate(MONTHS):
            first, lastd = MONTH_BOUNDS[mo]
            overlap = (f"MAX(0,MIN($E{i},{dstr(lastd)})"
                       f"-MAX($D{i},{dstr(first)})+1)")
            cp.cell(row=i, column=21 + k).value = (
                f"=ROUND($G{i}*{overlap}/$F{i},0)")
            cp.cell(row=i, column=27 + k).value = (
                f"=ROUND($K{i}*{overlap}/$F{i},0)")
            cp.cell(row=i, column=33 + k).value = (
                f"=ROUND($J{i}*{overlap}/$F{i},0)")
    style_header(cp, 38)
    cp.freeze_panes = "A2"
    widths = {"A": 11, "B": 26, "C": 15, "D": 12, "E": 12, "F": 7,
              "G": 12, "H": 13, "I": 10, "J": 12, "K": 13, "L": 9,
              "M": 10, "N": 11, "O": 13, "P": 8, "Q": 15, "R": 15,
              "S": 12, "T": 22}
    for col, w in widths.items():
        cp.column_dimensions[col].width = w
    for col in list("UVWXYZ") + ["AA", "AB", "AC", "AD", "AE", "AF",
                                 "AG", "AH", "AI", "AJ", "AK", "AL"]:
        cp.column_dimensions[col].width = 13
    for r in range(2, last + 1):
        for col in (4, 5):
            cp.cell(row=r, column=col).number_format = DATE_FMT
        for col in (7, 11, 18):
            cp.cell(row=r, column=col).number_format = PKR_FMT
        for col in (12, 15):
            cp.cell(row=r, column=col).number_format = PCT_FMT
        for col in (13, 14, 17):
            cp.cell(row=r, column=col).number_format = AVG_FMT
        cp.cell(row=r, column=16).number_format = ROAS_FMT
        cp.cell(row=r, column=19).number_format = DAYS_FMT
        for col in range(21, 39):
            cp.cell(row=r, column=col).number_format = PKR_FMT
        for col in range(33, 39):
            cp.cell(row=r, column=col).number_format = COUNT_FMT

    # ── Channels ────────────────────────────────────────────────────
    ch = wb.create_sheet("Channels")
    ch.append(["channel", "campaigns", "spend_pkr", "impressions",
               "clicks", "conversions", "revenue_pkr", "ctr",
               "cpc_pkr", "cpa_pkr", "conversion_rate", "roas",
               "revenue_per_click_pkr", "gross_profit_pkr",
               "campaign_days", "payback_days"])
    for j, label in enumerate(CHANNELS, start=2):
        ch.cell(row=j, column=1).value = label
        ch.cell(row=j, column=2).value = (
            f'=COUNTIF(Campaigns!$C$2:$C${last},$A{j})')
        ch.cell(row=j, column=3).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$G$2:$G${last})')
        ch.cell(row=j, column=4).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$H$2:$H${last})')
        ch.cell(row=j, column=5).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$I$2:$I${last})')
        ch.cell(row=j, column=6).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$J$2:$J${last})')
        ch.cell(row=j, column=7).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$K$2:$K${last})')
        ch.cell(row=j, column=8).value = (
            f'=IF(D{j}=0,"—",E{j}/D{j})')
        ch.cell(row=j, column=9).value = (
            f'=IF(E{j}=0,"—",C{j}/E{j})')
        ch.cell(row=j, column=10).value = (
            f'=IF(F{j}=0,"—",C{j}/F{j})')
        ch.cell(row=j, column=11).value = (
            f'=IF(E{j}=0,"—",F{j}/E{j})')
        ch.cell(row=j, column=12).value = (
            f'=IF(C{j}=0,"—",G{j}/C{j})')
        ch.cell(row=j, column=13).value = (
            f'=IF(E{j}=0,"—",G{j}/E{j})')
        ch.cell(row=j, column=14).value = f"=G{j}*0.45"
        ch.cell(row=j, column=15).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$F$2:$F${last})')
        ch.cell(row=j, column=16).value = (
            f'=IF(N{j}=0,"—",C{j}/(N{j}/O{j}))')
        for col in (3, 7, 14):
            ch.cell(row=j, column=col).number_format = PKR_FMT
        for col in (8, 11):
            ch.cell(row=j, column=col).number_format = PCT_FMT
        for col in (9, 10, 13):
            ch.cell(row=j, column=col).number_format = AVG_FMT
        ch.cell(row=j, column=12).number_format = ROAS_FMT
        ch.cell(row=j, column=16).number_format = DAYS_FMT
    # totals row (row 9)
    ch.cell(row=9, column=1).value = "TOTAL"
    ch.cell(row=9, column=1).font = Font(bold=True)
    for col, f in ((2, "=SUM(B2:B8)"), (3, "=SUM(C2:C8)"),
                   (4, "=SUM(D2:D8)"), (5, "=SUM(E2:E8)"),
                   (6, "=SUM(F2:F8)"), (7, "=SUM(G2:G8)"),
                   (14, "=SUM(N2:N8)"), (15, "=SUM(O2:O8)")):
        ch.cell(row=9, column=col).value = f
        ch.cell(row=9, column=col).font = Font(bold=True)
    ch.cell(row=9, column=8).value = "=IF(D9=0,\"—\",E9/D9)"
    ch.cell(row=9, column=9).value = "=IF(E9=0,\"—\",C9/E9)"
    ch.cell(row=9, column=10).value = "=IF(F9=0,\"—\",C9/F9)"
    ch.cell(row=9, column=11).value = "=IF(E9=0,\"—\",F9/E9)"
    ch.cell(row=9, column=12).value = "=IF(C9=0,\"—\",G9/C9)"
    ch.cell(row=9, column=13).value = "=IF(E9=0,\"—\",G9/E9)"
    ch.cell(row=9, column=16).value = "=IF(N9=0,\"—\",C9/(N9/O9))"
    for col in (3, 7, 14):
        ch.cell(row=9, column=col).number_format = PKR_FMT
    for col in (8, 11):
        ch.cell(row=9, column=col).number_format = PCT_FMT
    for col in (9, 10, 13):
        ch.cell(row=9, column=col).number_format = AVG_FMT
    ch.cell(row=9, column=12).number_format = ROAS_FMT
    ch.cell(row=9, column=16).number_format = DAYS_FMT
    style_header(ch, 16)
    ch.freeze_panes = "A2"
    ch.column_dimensions["A"].width = 15
    for col in "BCDEFGNO":
        ch.column_dimensions[col].width = 13
    ch.column_dimensions["H"].width = 9
    ch.column_dimensions["I"].width = 10
    ch.column_dimensions["J"].width = 11
    ch.column_dimensions["K"].width = 13
    ch.column_dimensions["L"].width = 8
    ch.column_dimensions["M"].width = 15
    ch.column_dimensions["P"].width = 12

    chart = BarChart()
    chart.type = "col"
    chart.title = "ROAS by channel — StyleHub.pk (SAMPLE)"
    chart.y_axis.title = "ROAS (revenue / spend)"
    chart.x_axis.title = "Channel"
    data = Reference(ch, min_col=12, min_row=1, max_row=8)
    cats = Reference(ch, min_col=1, min_row=2, max_row=8)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 7.5
    chart.width = 14
    ch.add_chart(chart, "A11")

    # ── Monthly Trend ───────────────────────────────────────────────
    mt = wb.create_sheet("Monthly Trend")
    mt.append(["month", "campaigns_active", "spend_pkr",
               "revenue_pkr", "conversions", "roas"])
    spend_cols = ["U", "V", "W", "X", "Y", "Z"]
    rev_cols = ["AA", "AB", "AC", "AD", "AE", "AF"]
    conv_cols = ["AG", "AH", "AI", "AJ", "AK", "AL"]
    for k, mo in enumerate(MONTHS):
        r = 2 + k
        first, lastd = MONTH_BOUNDS[mo]
        mt.cell(row=r, column=1).value = mo
        mt.cell(row=r, column=2).value = (
            f'=COUNTIFS(Campaigns!$D$2:$D${last},"<="&'
            f'{dstr(lastd)},Campaigns!$E$2:$E${last},">="&'
            f'{dstr(first)})')
        mt.cell(row=r, column=3).value = (
            f"=SUM(Campaigns!${spend_cols[k]}$2:"
            f"${spend_cols[k]}${last})")
        mt.cell(row=r, column=4).value = (
            f"=SUM(Campaigns!${rev_cols[k]}$2:"
            f"${rev_cols[k]}${last})")
        mt.cell(row=r, column=5).value = (
            f"=SUM(Campaigns!${conv_cols[k]}$2:"
            f"${conv_cols[k]}${last})")
        mt.cell(row=r, column=6).value = (
            f'=IF(C{r}=0,"—",D{r}/C{r})')
        for col in (3, 4):
            mt.cell(row=r, column=col).number_format = PKR_FMT
        mt.cell(row=r, column=6).number_format = ROAS_FMT
    mt.cell(row=8, column=1).value = "TOTAL"
    mt.cell(row=8, column=1).font = Font(bold=True)
    for col, f in ((3, "=SUM(C2:C7)"), (4, "=SUM(D2:D7)"),
                   (5, "=SUM(E2:E7)")):
        mt.cell(row=8, column=col).value = f
        mt.cell(row=8, column=col).font = Font(bold=True)
        mt.cell(row=8, column=col).number_format = PKR_FMT
    mt.cell(row=8, column=6).value = "=IF(C8=0,\"—\",D8/C8)"
    mt.cell(row=8, column=6).font = Font(bold=True)
    mt.cell(row=8, column=6).number_format = ROAS_FMT
    style_header(mt, 6)
    mt.freeze_panes = "A2"
    for col, w in zip("ABCDEF", (12, 17, 14, 14, 13, 9)):
        mt.column_dimensions[col].width = w

    mchart = BarChart()
    mchart.type = "col"
    mchart.title = "Spend vs revenue by month — StyleHub.pk (SAMPLE)"
    mchart.y_axis.title = "PKR"
    mchart.x_axis.title = "Month"
    mdata = Reference(mt, min_col=3, min_row=1, max_row=7)
    mcats = Reference(mt, min_col=1, min_row=2, max_row=7)
    mchart.add_data(mdata, titles_from_data=True)
    mchart.set_categories(mcats)
    mchart.height = 7.5
    mchart.width = 14
    mt.add_chart(mchart, "A10")

    # ── Funnel ──────────────────────────────────────────────────────
    fn = wb.create_sheet("Funnel")
    fn.cell(row=1, column=1).value = (
        "Funnel — impressions to orders (SAMPLE)")
    fn.cell(row=1, column=1).font = Font(bold=True, size=12,
                                         color="1F4E79")
    fn.append(["stage", "count", "conversion_from_previous",
               "share_of_impressions"])
    style_header(fn, 4, row=2)
    funnel_rows = [
        ("Impressions", f"=SUM(Campaigns!$H$2:$H${last})"),
        ("Clicks", f"=SUM(Campaigns!$I$2:$I${last})"),
        ("Conversions", f"=SUM(Campaigns!$J$2:$J${last})"),
    ]
    for k, (label, formula) in enumerate(funnel_rows):
        r = 3 + k
        fn.cell(row=r, column=1).value = label
        fn.cell(row=r, column=2).value = formula
        if k == 0:
            fn.cell(row=r, column=3).value = 1.0
        else:
            fn.cell(row=r, column=3).value = (
                f'=IF(B{r-1}=0,"—",B{r}/B{r-1})')
        fn.cell(row=r, column=4).value = (
            f'=IF($B$3=0,"—",B{r}/$B$3)')
        fn.cell(row=r, column=3).number_format = PCT_FMT
        fn.cell(row=r, column=4).number_format = PCT_FMT
    fn.cell(row=6, column=1).value = "Revenue from conversions (PKR)"
    fn.cell(row=6, column=1).font = Font(bold=True)
    fn.cell(row=6, column=2).value = f"=SUM(Campaigns!$K$2:$K${last})"
    fn.cell(row=6, column=2).number_format = PKR_FMT
    fn.cell(row=6, column=2).font = Font(bold=True)

    fn.cell(row=8, column=1).value = "Funnel by channel"
    fn.cell(row=8, column=1).font = Font(bold=True, size=12,
                                         color="1F4E79")
    fn.append(["channel", "impressions", "clicks", "conversions",
               "ctr", "conversion_rate", "revenue_pkr", "spend_pkr",
               "roas"])
    style_header(fn, 9, row=9)
    for j, label in enumerate(CHANNELS, start=10):
        fn.cell(row=j, column=1).value = label
        fn.cell(row=j, column=2).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$H$2:$H${last})')
        fn.cell(row=j, column=3).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$I$2:$I${last})')
        fn.cell(row=j, column=4).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$J$2:$J${last})')
        fn.cell(row=j, column=5).value = (
            f'=IF(B{j}=0,"—",C{j}/B{j})')
        fn.cell(row=j, column=6).value = (
            f'=IF(C{j}=0,"—",D{j}/C{j})')
        fn.cell(row=j, column=7).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$K$2:$K${last})')
        fn.cell(row=j, column=8).value = (
            f'=SUMIF(Campaigns!$C$2:$C${last},$A{j},'
            f'Campaigns!$G$2:$G${last})')
        fn.cell(row=j, column=9).value = (
            f'=IF(H{j}=0,"—",G{j}/H{j})')
        for col in (5, 6):
            fn.cell(row=j, column=col).number_format = PCT_FMT
        fn.cell(row=j, column=7).number_format = PKR_FMT
        fn.cell(row=j, column=8).number_format = PKR_FMT
        fn.cell(row=j, column=9).number_format = ROAS_FMT
    fn.cell(row=17, column=1).value = "TOTAL"
    fn.cell(row=17, column=1).font = Font(bold=True)
    for col, f in ((2, "=SUM(B10:B16)"), (3, "=SUM(C10:C16)"),
                   (4, "=SUM(D10:D16)"), (7, "=SUM(G10:G16)"),
                   (8, "=SUM(H10:H16)")):
        fn.cell(row=17, column=col).value = f
        fn.cell(row=17, column=col).font = Font(bold=True)
    fn.cell(row=17, column=5).value = "=IF(B17=0,\"—\",C17/B17)"
    fn.cell(row=17, column=6).value = "=IF(C17=0,\"—\",D17/C17)"
    fn.cell(row=17, column=9).value = "=IF(H17=0,\"—\",G17/H17)"
    fn.cell(row=17, column=5).number_format = PCT_FMT
    fn.cell(row=17, column=6).number_format = PCT_FMT
    fn.cell(row=17, column=7).number_format = PKR_FMT
    fn.cell(row=17, column=8).number_format = PKR_FMT
    fn.cell(row=17, column=9).number_format = ROAS_FMT
    fn.column_dimensions["A"].width = 24
    fn.column_dimensions["B"].width = 14
    fn.column_dimensions["C"].width = 12
    fn.column_dimensions["D"].width = 13
    fn.column_dimensions["E"].width = 10
    fn.column_dimensions["F"].width = 15
    fn.column_dimensions["G"].width = 15
    fn.column_dimensions["H"].width = 13
    fn.column_dimensions["I"].width = 9

    # ── Dashboard ───────────────────────────────────────────────────
    db = wb.create_sheet("Dashboard", 0)
    db.cell(row=1, column=1).value = (
        "StyleHub.pk — Marketing Campaign ROI Tracker")
    db.cell(row=1, column=1).font = TITLE_FONT
    db.cell(row=2, column=1).value = (
        "SAMPLE delivery on DEMO data · Jan–Jun 2026 · "
        "12 campaigns, 7 channels · All numbers are live formulas "
        "over Campaigns")
    db.cell(row=2, column=1).font = Font(italic=True, size=11)
    db.append(["KPI", "Value", "How it is calculated"])
    style_header(db, 3, row=3)
    L = last
    kpis = [
        ("Total campaigns", f"=COUNTA(Campaigns!$A$2:$A${L})",
         "Count of campaigns on Campaigns", COUNT_FMT),
        ("Total spend (PKR)", f"=SUM(Campaigns!$G$2:$G${L})",
         "Sum of campaign spend", PKR_FMT),
        ("Total revenue (PKR)", f"=SUM(Campaigns!$K$2:$K${L})",
         "Sum of campaign revenue", PKR_FMT),
        ("Overall ROAS", "=IF(B5=0,\"—\",B6/B5)",
         "Total revenue / total spend", ROAS_FMT),
        ("Total impressions", f"=SUM(Campaigns!$H$2:$H${L})",
         "Sum of impressions", COUNT_FMT),
        ("Total clicks", f"=SUM(Campaigns!$I$2:$I${L})",
         "Sum of clicks", COUNT_FMT),
        ("Total conversions", f"=SUM(Campaigns!$J$2:$J${L})",
         "Sum of orders", COUNT_FMT),
        ("Overall CTR", "=IF(B8=0,\"—\",B9/B8)",
         "Clicks / impressions", PCT_FMT),
        ("Overall CPC (PKR)", "=IF(B9=0,\"—\",B5/B9)",
         "Spend / clicks", AVG_FMT),
        ("Overall CPA (PKR)", "=IF(B10=0,\"—\",B5/B10)",
         "Spend / conversions", AVG_FMT),
        ("Overall conversion rate", "=IF(B9=0,\"—\",B10/B9)",
         "Conversions / clicks", PCT_FMT),
        ("Revenue per click (PKR)", "=IF(B9=0,\"—\",B6/B9)",
         "Revenue / clicks", AVG_FMT),
        ("Gross profit (PKR, 45% margin)", "=B6*0.45",
         "Revenue x 45% gross margin", PKR_FMT),
        ("Blended payback (days)", "=IF(B16=0,\"—\",B5/(B16/181))",
         "Spend / (gross profit / 181 days)", DAYS_FMT),
        ("Channels", "=COUNTA(Channels!$A$2:$A$8)",
         "Count of channels on Channels", COUNT_FMT),
    ]
    # KPI rows start at row 4; cross-references above use the fixed
    # layout: B4 campaigns, B5 spend, B6 revenue, B8 impressions,
    # B9 clicks, B10 conversions, B16 gross profit.
    for k, (label, formula, note, fmt) in enumerate(kpis):
        r = 4 + k
        db.cell(row=r, column=1).value = label
        db.cell(row=r, column=2).value = formula
        db.cell(row=r, column=3).value = note
        db.cell(row=r, column=2).number_format = fmt
        db.cell(row=r, column=1).font = Font(bold=True)
    assert kpis[0][0] == "Total campaigns" \
        and kpis[1][0] == "Total spend (PKR)"
    assert kpis[2][0] == "Total revenue (PKR)" \
        and kpis[12][0] == "Gross profit (PKR, 45% margin)"

    r0 = 4 + len(kpis) + 1  # first row of the leaders section
    db.cell(row=r0, column=1).value = "Winners, losers and calls"
    db.cell(row=r0, column=1).font = Font(bold=True, size=12,
                                          color="1F4E79")
    db.cell(row=r0 + 1, column=1).value = "Best campaign by ROAS"
    db.cell(row=r0 + 1, column=2).value = (
        f"=INDEX(Campaigns!$B$2:$B${L},MATCH("
        f"MAX(Campaigns!$P$2:$P${L}),Campaigns!$P$2:$P${L},0))")
    db.cell(row=r0 + 1, column=3).value = (
        f"=MAX(Campaigns!$P$2:$P${L})")
    db.cell(row=r0 + 1, column=3).number_format = ROAS_FMT
    db.cell(row=r0 + 2, column=1).value = "Worst campaign by ROAS"
    db.cell(row=r0 + 2, column=2).value = (
        f"=INDEX(Campaigns!$B$2:$B${L},MATCH("
        f"MIN(Campaigns!$P$2:$P${L}),Campaigns!$P$2:$P${L},0))")
    db.cell(row=r0 + 2, column=3).value = (
        f"=MIN(Campaigns!$P$2:$P${L})")
    db.cell(row=r0 + 2, column=3).number_format = ROAS_FMT
    db.cell(row=r0 + 3, column=1).value = "Best channel by ROAS"
    db.cell(row=r0 + 3, column=2).value = (
        "=INDEX(Channels!$A$2:$A$8,MATCH(MAX(Channels!$L$2:$L$8),"
        "Channels!$L$2:$L$8,0))")
    db.cell(row=r0 + 3, column=3).value = "=MAX(Channels!$L$2:$L$8)"
    db.cell(row=r0 + 3, column=3).number_format = ROAS_FMT
    db.cell(row=r0 + 4, column=1).value = "Worst channel by ROAS"
    db.cell(row=r0 + 4, column=2).value = (
        "=INDEX(Channels!$A$2:$A$8,MATCH(MIN(Channels!$L$2:$L$8),"
        "Channels!$L$2:$L$8,0))")
    db.cell(row=r0 + 4, column=3).value = "=MIN(Channels!$L$2:$L$8)"
    db.cell(row=r0 + 4, column=3).number_format = ROAS_FMT
    db.cell(row=r0 + 5, column=1).value = "Best channel by revenue"
    db.cell(row=r0 + 5, column=2).value = (
        "=INDEX(Channels!$A$2:$A$8,MATCH(MAX(Channels!$G$2:$G$8),"
        "Channels!$G$2:$G$8,0))")
    db.cell(row=r0 + 5, column=3).value = "=MAX(Channels!$G$2:$G$8)"
    db.cell(row=r0 + 5, column=3).number_format = PKR_FMT
    db.cell(row=r0 + 6, column=1).value = "Campaigns to Kill"
    db.cell(row=r0 + 6, column=2).value = (
        f'=COUNTIF(Campaigns!$T$2:$T${L},"Kill")')
    db.cell(row=r0 + 6, column=3).value = (
        f'=SUMIF(Campaigns!$T$2:$T${L},"Kill",'
        f'Campaigns!$G$2:$G${L})')
    db.cell(row=r0 + 6, column=3).number_format = PKR_FMT
    db.cell(row=r0 + 7, column=1).value = "Campaigns to Scale"
    db.cell(row=r0 + 7, column=2).value = (
        f'=COUNTIF(Campaigns!$T$2:$T${L},"Scale")')
    db.cell(row=r0 + 7, column=3).value = (
        f'=SUMIF(Campaigns!$T$2:$T${L},"Scale",'
        f'Campaigns!$G$2:$G${L})')
    db.cell(row=r0 + 7, column=3).number_format = PKR_FMT
    db.cell(row=r0 + 8, column=1).value = "Campaigns to Keep / Watch"
    db.cell(row=r0 + 8, column=2).value = (
        f'=COUNTIF(Campaigns!$T$2:$T${L},"Keep*")')
    db.cell(row=r0 + 8, column=3).value = (
        f'=SUMIF(Campaigns!$T$2:$T${L},"Keep*",'
        f'Campaigns!$G$2:$G${L})')
    db.cell(row=r0 + 8, column=3).number_format = PKR_FMT
    db.cell(row=r0 + 9, column=1).value = "Revenue lost to Kill campaigns"
    db.cell(row=r0 + 9, column=2).value = (
        f'=SUMIF(Campaigns!$T$2:$T${L},"Kill",'
        f'Campaigns!$K$2:$K${L})')
    db.cell(row=r0 + 9, column=2).number_format = PKR_FMT
    db.cell(row=r0 + 9, column=3).value = (
        "Revenue the Kill campaigns returned (spend is in the row "
        "above it)")
    db.column_dimensions["A"].width = 40
    db.column_dimensions["B"].width = 24
    db.column_dimensions["C"].width = 52
    for r in range(r0 + 1, r0 + 10):
        db.cell(row=r, column=1).font = Font(bold=True)

    # ── Method ──────────────────────────────────────────────────────
    md = wb.create_sheet("Method")
    md.cell(row=1, column=1).value = "Method — how to read this workbook"
    md.cell(row=1, column=1).font = Font(bold=True, size=12,
                                         color="1F4E79")
    lines = [
        "SAMPLE delivery on DEMO data. StyleHub.pk is a fictional "
        "Pakistani fashion e-commerce store; the 12 campaigns were "
        "generated to show the method. No real store, campaign or "
        "customer is involved.",
        "",
        "CTR = clicks / impressions. CPC = spend / clicks. "
        "CPA = spend / conversions. Conversion rate = conversions / "
        "clicks. ROAS = revenue / spend. Revenue per click = revenue "
        "/ clicks.",
        "",
        "Gross profit = revenue x 45%. The 45% is the stated "
        "assumption: StyleHub.pk sells fashion at roughly 45% gross "
        "margin after product cost. It is an assumption, not a "
        "measured margin — the workbook shows the unadjusted revenue "
        "next to every adjusted number so the reader can re-run it "
        "at their own margin.",
        "",
        "Payback days = spend / (gross profit / campaign days): how "
        "many days of gross profit it takes to earn the spend back. "
        "A campaign whose payback is longer than its flight never "
        "paid back while it ran — both Kill campaigns are in that "
        "position.",
        "",
        "The call rule: ROAS < 1.00 -> Kill (every rupee returns "
        "less than a rupee). ROAS >= 3.00 and spend < PKR 900,000 -> "
        "Scale (proven, with headroom to grow). ROAS >= 3.00 and "
        "spend >= PKR 900,000 -> Keep — proven at scale (already at "
        "the ceiling, hold and optimise). Otherwise -> Keep / Watch "
        "(profitable, but below the Scale bar). The rule is a "
        "formula in the Call column, not a typed label.",
        "",
        "Monthly trend: a campaign's spend, revenue and conversions "
        "are spread evenly across its flight days, so a campaign "
        "running Mar 1 – Apr 15 puts 31/46 of its spend in March "
        "and 15/46 in April. Monthly totals are the sum of those "
        "day-shares — the split columns on Campaigns are live "
        "formulas, not typed numbers.",
        "",
        "Funnel: impressions -> clicks -> conversions. Conversion "
        "at each step = this stage / the previous stage. The leak "
        "that matters is click -> order: only a small share of "
        "clicks become orders, and TikTok's share is the smallest "
        "of any paid channel.",
        "",
        "Small bases are noisy: Email's 15.98x ROAS sits on only "
        "PKR 83,000 of spend and 301 orders — the rate is real, the "
        "scale is not. Every rate in this workbook sits next to its "
        "spend and its order count, so the reader can see both.",
        "",
        "For Email and SMS/WhatsApp, 'impressions' means messages "
        "sent (emails delivered / SMS delivered), and CTR means "
        "clicks / messages sent — the same funnel, measured the way "
        "those channels are actually bought.",
        "",
        "Re-verify: run python3 src/checks.py — it recomputes every "
        "headline number from data/raw/campaigns_SAMPLE.csv, "
        "compares the workbook formulas after a LibreOffice recalc, "
        "and checks the PDF and charts exist. On a real client file, "
        "replace data/raw/campaigns_SAMPLE.csv with the ad-platform "
        "export (same columns) and skip src/generate_data.py — "
        "everything downstream runs unchanged.",
    ]
    for i, line in enumerate(lines, start=2):
        c = md.cell(row=i, column=1)
        c.value = line
        c.alignment = Alignment(wrap_text=True, vertical="top")
    md.column_dimensions["A"].width = 120

    wb.save(OUT)
    print(f"Wrote {OUT} — {wb.sheetnames}")


if __name__ == "__main__":
    main()
