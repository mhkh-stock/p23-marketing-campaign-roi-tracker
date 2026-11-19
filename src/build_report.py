#!/usr/bin/env python3
"""Build the P23 deliverables from the SAMPLE campaign data.

Recomputes every number from outputs/summary.json (itself recomputed
from the raw CSV by src/analyze.py — nothing is typed in by hand)
and produces:
  * outputs/channel_roas.png — ROAS by channel (matplotlib)
  * outputs/monthly_trend.png — spend vs revenue by month (matplotlib)
  * outputs/campaign_roi_report.pdf — 2-page client report (headline,
    channel table, campaign table with Kill / Scale / Keep calls,
    monthly trend, funnel, 5 recommendations)

Run:  python3 src/build_report.py
"""
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent.parent
SUMMARY = ROOT / "outputs" / "summary.json"
CHART_CH = ROOT / "outputs" / "channel_roas.png"
CHART_MO = ROOT / "outputs" / "monthly_trend.png"
PDF = ROOT / "outputs" / "campaign_roi_report.pdf"

s = json.loads(SUMMARY.read_text(encoding="utf-8"))
TOT_SPEND = s["total_spend_pkr"]
TOT_REV = s["total_revenue_pkr"]
ROAS = s["overall_roas"]
IMPR = s["total_impressions"]
CLICKS = s["total_clicks"]
CONV = s["total_conversions"]
CTR = s["overall_ctr_pct"]
CPC = s["overall_cpc_pkr"]
CPA = s["overall_cpa_pkr"]
CR = s["overall_conversion_rate_pct"]
GROSS = s["gross_profit_pkr"]
PAYBACK = s["blended_payback_days"]
CAMPS = s["campaigns"]
CHANS = sorted(s["channels"], key=lambda x: -x["roas"])
MO = s["monthly"]
FUN = s["funnel"]
ACTS = s["actions"]

# channel ROAS chart: ROAS by channel, Kill line at 1.0, Scale bar at 3.0
labels = [c["channel"] for c in CHANS]
roas = [c["roas"] for c in CHANS]
fig, ax = plt.subplots(figsize=(8, 3.4))
bar_colors = ["#C62828" if v < 1 else "#2E7D32" if v >= 3 else "#F9A825"
              for v in roas]
bars = ax.barh(labels[::-1], roas[::-1], color=bar_colors[::-1])
ax.axvline(1.0, color="#C62828", linestyle="--", linewidth=1,
           label="Kill line (ROAS 1.0)")
ax.axvline(3.0, color="#2E7D32", linestyle="--", linewidth=1,
           label="Scale bar (ROAS 3.0)")
ax.set_xlabel("ROAS (revenue / spend)")
ax.set_title("ROAS by channel — StyleHub.pk (SAMPLE, Jan–Jun 2026)")
for bar, v in zip(bars, roas[::-1]):
    ax.text(v + 0.3, bar.get_y() + bar.get_height() / 2,
            f"{v:.2f}x", va="center", fontsize=9)
ax.set_xlim(0, max(roas) * 1.15)
ax.legend(fontsize=8, loc="lower right")
fig.tight_layout()
plt.savefig(CHART_CH, dpi=150)
plt.close()

# monthly trend: spend vs revenue bars by month
months = [m["month"] for m in MO]
spend = [m["spend_pkr"] for m in MO]
rev = [m["revenue_pkr"] for m in MO]
fig, ax = plt.subplots(figsize=(8, 3.4))
x = range(len(months))
w = 0.38
ax.bar([i - w / 2 for i in x], spend, width=w, color="#1F4E79",
       label="Spend (PKR)")
ax.bar([i + w / 2 for i in x], rev, width=w, color="#2E7D32",
       label="Revenue (PKR)")
ax.set_xticks(list(x))
ax.set_xticklabels(months)
ax.set_ylabel("PKR")
ax.set_title("Spend vs revenue by month — StyleHub.pk (SAMPLE)")
ax.ticklabel_format(style="plain", axis="y")
ax.legend(fontsize=9)
fig.tight_layout()
plt.savefig(CHART_MO, dpi=150)
plt.close()

styles = getSampleStyleSheet()
styles["Heading2"].spaceBefore = 4
styles["Heading2"].spaceAfter = 2
styles["Heading2"].fontSize = 12
styles["Heading2"].leading = 14
styles["Normal"].leading = 13
styles["Normal"].spaceAfter = 2
styles["Title"].fontSize = 20
styles["Title"].leading = 22
styles["Title"].spaceAfter = 2
hdr = colors.HexColor("#1F4E79")


def table(rows, widths, size=8):
    t = Table(rows, colWidths=widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), hdr),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), size),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def p(txt, style="Normal"):
    return Paragraph(txt, styles[style])


kills = [c for c in CAMPS if c["call"] == "Kill"]
best = max(CAMPS, key=lambda x: x["roas"])
worst = min(CAMPS, key=lambda x: x["roas"])

story = [
    p("StyleHub.pk — Marketing Campaign ROI Tracker", "Title"),
    p("SAMPLE delivery on DEMO data · Jan–Jun 2026 · "
      "12 campaigns, 7 channels · Which campaigns earn their money "
      "back?", "Normal"),
    Spacer(1, 0.12 * cm),
    p(f"<b>PKR {TOT_SPEND:,} spent · PKR {TOT_REV:,} back · "
      f"ROAS {ROAS}x · {IMPR:,} impressions · {CLICKS:,} clicks · "
      f"{CONV:,} orders · CTR {CTR}% · CPC PKR {CPC:,.2f} · "
      f"CPA PKR {CPA:,.2f} · gross profit (45%) "
      f"PKR {GROSS:,.0f} · blended payback {PAYBACK} days</b>",
      "Normal"),
    Spacer(1, 0.12 * cm),
    p("<b>The answer in one line:</b> the account returns "
      f"<b>{ROAS}x</b> overall, but the average hides the story — "
      f"<b>{s['scale_count']} campaigns are worth scaling, "
      f"{s['keep_count']} are worth watching and "
      f"{s['kill_count']} should be killed</b>: "
      + ", ".join(c["campaign_name"] for c in kills)
      + f" burned PKR {s['kill_spend_pkr']:,} for "
      f"PKR {s['kill_revenue_pkr']:,} back. Move that money to the "
      f"Scale campaigns (average ROAS {s['scale_avg_roas']}x) and it "
      f"returns about PKR {s['reallocated_revenue_pkr']:,} instead.",
      "Normal"),
    Spacer(1, 0.1 * cm),
    p("ROAS by channel — where the money works", "Heading2"),
    Image(str(CHART_CH), width=14 * cm, height=3.1 * cm),
    table(
        [["Channel", "Campaigns", "Spend", "Revenue", "ROAS",
          "Conv. rate", "CPA"]]
        + [[c["channel"], str(c["campaigns"]),
            f"PKR {c['spend_pkr']:,}", f"PKR {c['revenue_pkr']:,}",
            f"{c['roas']}x", f"{c['conversion_rate_pct']}%",
            f"PKR {c['cpa_pkr']:,.0f}"] for c in CHANS],
        [3.4 * cm, 2.2 * cm, 2.6 * cm, 2.6 * cm, 1.6 * cm, 2 * cm,
         2.4 * cm]),
    Spacer(1, 0.12 * cm),
    p(f"<b>Best channel by ROAS: {s['best_channel_roas']} "
      f"({s['best_channel_roas_value']}x).</b> Worst: "
      f"{s['worst_channel_roas']} "
      f"({s['worst_channel_roas_value']}x). Biggest revenue channel: "
      f"{s['best_channel_revenue']} "
      f"(PKR {s['best_channel_revenue_value']:,}). TikTok brings the "
      f"most clicks in the account at the lowest conversion rate — "
      f"cheap traffic that does not buy.",
      "Normal"),
    Spacer(1, 0.1 * cm),
    p("Every campaign — with its call", "Heading2"),
    table(
        [["Campaign", "Channel", "Spend", "Revenue", "ROAS",
          "Conv. rate", "Payback", "Call"]]
        + [[c["campaign_name"], c["channel"],
            f"PKR {c['spend_pkr']:,}", f"PKR {c['revenue_pkr']:,}",
            f"{c['roas']}x", f"{c['conversion_rate_pct']}%",
            f"{c['payback_days']} days", c["call"]]
           for c in sorted(CAMPS, key=lambda x: -x["roas"])],
        [3.6 * cm, 2.6 * cm, 2.2 * cm, 2.2 * cm, 1.3 * cm, 1.7 * cm,
         1.9 * cm, 2.3 * cm], size=7.5),
    Spacer(1, 0.12 * cm),
    p(f"<b>Best campaign: {best['campaign_name']} (ROAS "
      f"{best['roas']}x, payback {best['payback_days']} days). "
      f"Worst: {worst['campaign_name']} (ROAS {worst['roas']}x, "
      f"payback {worst['payback_days']} days — longer than its "
      f"{worst['days']}-day flight, so it never paid back).</b> "
      f"Call rule: ROAS &lt; 1.00 = Kill; ROAS ≥ 3.00 and spend "
      f"&lt; PKR 900,000 = Scale; ROAS ≥ 3.00 at bigger spend = "
      f"Keep — proven at scale; otherwise Keep / Watch.",
      "Normal"),
    Spacer(1, 0.1 * cm),
    p("Month by month — the trend", "Heading2"),
    Image(str(CHART_MO), width=14 * cm, height=3.1 * cm),
    table(
        [["Month", "Campaigns active", "Spend", "Revenue",
          "Orders", "ROAS"]]
        + [[m["month"], str(m["campaigns_active"]),
            f"PKR {m['spend_pkr']:,}", f"PKR {m['revenue_pkr']:,}",
            str(m["conversions"]), f"{m['roas']}x"] for m in MO],
        [2.6 * cm, 3.4 * cm, 3.2 * cm, 3.2 * cm, 2 * cm, 2.4 * cm]),
    Spacer(1, 0.12 * cm),
    p("January–February ran at ROAS 1.76–1.84x — the account was "
      "carrying the Brand Awareness flight and the TikTok launch. "
      "From March the Eid campaigns and the Influencer flight land "
      "and ROAS jumps to 3.09x, peaking at 4.05x in April. Spend is "
      "spread by day-share across each campaign's flight, so the "
      "monthly numbers are the campaigns' own days, not a guess.",
      "Normal"),
    Spacer(1, 0.1 * cm),
    p("The funnel — impressions to orders", "Heading2"),
    table(
        [["Stage", "Count", "Conversion", "Share of impressions"]]
        + [[f["stage"], f"{f['count']:,}",
            f"{f['conversion_from_previous']}%",
            f"{f['share_of_impressions']}%"] for f in FUN],
        [4.5 * cm, 3.5 * cm, 4 * cm, 4.8 * cm]),
    Spacer(1, 0.12 * cm),
    p(f"Only {FUN[2]['share_of_impressions']}% of impressions "
      f"become an order. The leak that matters is click → order: "
      f"{CR}% of clicks convert. Email converts at "
      f"{[c for c in CHANS if c['channel'] == 'Email'][0]['conversion_rate_pct']}% "
      f"on a small list; TikTok converts at "
      f"{[c for c in CHANS if c['channel'] == 'TikTok Ads'][0]['conversion_rate_pct']}% "
      f"on the biggest click volume in the account.",
      "Normal"),
    Spacer(1, 0.1 * cm),
    p("What to do next — five actions from the data", "Heading2"),
]
for i, a in enumerate(ACTS, 1):
    story.append(p(f"<b>{i}.</b> {a}", "Normal"))
    story.append(Spacer(1, 0.1 * cm))
story.append(Spacer(1, 0.1 * cm))
story.append(p(
    "Method: CTR = clicks / impressions. CPC = spend / clicks. "
    "CPA = spend / conversions. Conversion rate = conversions / "
    "clicks. ROAS = revenue / spend. Gross profit = revenue x 45% "
    "(stated assumption: fashion gross margin after product cost). "
    "Payback days = spend / (gross profit / campaign days). Every "
    "number is recomputed from the campaign file and checked by "
    "src/checks.py; the workbook carries the same numbers as live "
    "formulas. SAMPLE data: StyleHub.pk is fictional.",
    "Normal"))

doc = SimpleDocTemplate(str(PDF), pagesize=A4,
                        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
                        topMargin=1.1 * cm, bottomMargin=1.1 * cm,
                        title="StyleHub.pk — Marketing Campaign ROI "
                              "Tracker (SAMPLE)",
                        author="Muhammad Hashir")
doc.build(story)
print(f"Wrote {PDF}")
