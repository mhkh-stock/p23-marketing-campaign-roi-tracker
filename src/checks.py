#!/usr/bin/env python3
"""Checks as code for P23. Recomputes every headline number from the
raw CSV and verifies the JSON, the processed CSVs, the markdown
report, the workbook formulas and the LibreOffice recalc against it
— three independent paths that must agree. CTR = clicks /
impressions; CPC = spend / clicks; CPA = spend / conversions;
conversion rate = conversions / clicks; ROAS = revenue / spend;
revenue per click = revenue / clicks; gross profit = revenue x 45%
(the stated margin assumption); payback days = spend / (gross profit
/ campaign days); call rule: ROAS < 1.00 -> Kill, ROAS >= 3.00 and
spend < PKR 900,000 -> Scale, ROAS >= 3.00 -> Keep — proven at
scale, else Keep / Watch; monthly split = day-share across each
campaign's flight.

Runs the LibreOffice recalc itself on every run, so the check
count is stable run to run.

    python3 src/checks.py

Exit code 0 when every check passes, 1 otherwise.
"""
import csv
import json
import subprocess
import sys
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"
PROC = PROJECT / "data" / "processed"
OUT = PROJECT / "outputs"
SUMMARY = OUT / "summary.json"
WORKBOOK = OUT / "campaign_roi_report_SAMPLE.xlsx"
RECALC_DIR = OUT / ".recalc"
RECALC = RECALC_DIR / "campaign_roi_report_SAMPLE.xlsx"
PDF = OUT / "campaign_roi_report.pdf"
MD = OUT / "campaign_roi_report.md"
README = PROJECT / "README.md"
PACKAGE = PROJECT / "PACKAGE-NOTE.md"
CHART_CH = OUT / "channel_roas.png"
CHART_MO = OUT / "monthly_trend.png"

CHANNELS = ["Google Search", "Google Shopping", "Meta Ads",
            "TikTok Ads", "Influencer", "Email", "SMS/WhatsApp"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
GROSS_MARGIN = 0.45
SCALE_CEILING = 900_000
fails = []
ran = []


def check(name, cond, detail=""):
    ran.append(name)
    if not cond:
        print("FAIL " + name + (f" — {detail}" if detail else ""))
        fails.append(name)


def close(a, b, tol=0.01):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def pct(n, d):
    return round(100.0 * n / d, 2) if d else 0.0


def r2(x):
    return round(x, 2)


def call_for(roas, spend):
    if roas < 1.0:
        return "Kill"
    if roas >= 3.0 and spend < SCALE_CEILING:
        return "Scale"
    if roas >= 3.0:
        return "Keep — proven at scale"
    return "Keep / Watch"


def metrics(spend, impressions, clicks, conversions, revenue, days):
    gross = revenue * GROSS_MARGIN
    return {
        "spend_pkr": spend,
        "impressions": impressions,
        "clicks": clicks,
        "conversions": conversions,
        "revenue_pkr": revenue,
        "ctr_pct": pct(clicks, impressions),
        "cpc_pkr": r2(spend / clicks) if clicks else 0.0,
        "cpa_pkr": r2(spend / conversions) if conversions else 0.0,
        "conversion_rate_pct": pct(conversions, clicks),
        "roas": r2(revenue / spend) if spend else 0.0,
        "revenue_per_click_pkr": r2(revenue / clicks)
        if clicks else 0.0,
        "gross_profit_pkr": r2(gross),
        "payback_days": r2(spend * days / gross) if gross else None,
        "days": days,
    }


def month_bounds(month):
    y, m = int(month[:4]), int(month[5:7])
    first = date(y, m, 1)
    last = date(y, m + 1, 1).fromordinal(
        date(y, m + 1, 1).toordinal() - 1) if m < 12 else date(y, 12, 31)
    return first, last


# ── raw file ────────────────────────────────────────────────────────
raw = list(csv.DictReader(open(RAW / "campaigns_SAMPLE.csv",
                               encoding="utf-8")))
check("raw: 12 campaigns", len(raw) == 12, f"got {len(raw)}")
check("raw: columns exact",
      list(raw[0].keys()) == [
          "campaign_id", "campaign_name", "channel", "start_date",
          "end_date", "days", "spend_pkr", "impressions", "clicks",
          "conversions", "revenue_pkr"],
      str(list(raw[0].keys())))
ids = [r["campaign_id"] for r in raw]
check("raw: campaign ids unique", len(set(ids)) == len(ids))
check("raw: campaign ids CMP-01..CMP-12",
      sorted(ids) == [f"CMP-{i:02d}" for i in range(1, 13)])
names = [r["campaign_name"] for r in raw]
check("raw: campaign names unique", len(set(names)) == len(names))
check("raw: 7 channels present",
      sorted(set(r["channel"] for r in raw)) == sorted(CHANNELS),
      str(sorted(set(r["channel"] for r in raw))))

rows = []
for r in raw:
    s = date.fromisoformat(r["start_date"])
    e = date.fromisoformat(r["end_date"])
    rows.append({
        "campaign_id": r["campaign_id"],
        "campaign_name": r["campaign_name"],
        "channel": r["channel"],
        "start_date": r["start_date"],
        "end_date": r["end_date"],
        "days": int(r["days"]),
        "spend_pkr": int(r["spend_pkr"]),
        "impressions": int(r["impressions"]),
        "clicks": int(r["clicks"]),
        "conversions": int(r["conversions"]),
        "revenue_pkr": int(r["revenue_pkr"]),
    })

for r in rows:
    d = r["campaign_id"]
    check(f"row {d}: channel known", r["channel"] in CHANNELS)
    check(f"row {d}: start in Jan-Jun 2026",
          "2026-01-01" <= r["start_date"] <= "2026-06-30")
    check(f"row {d}: end in Jan-Jun 2026",
          "2026-01-01" <= r["end_date"] <= "2026-06-30")
    check(f"row {d}: end on/after start",
          r["end_date"] >= r["start_date"])
    check(f"row {d}: days = end - start + 1",
          r["days"] == (date.fromisoformat(r["end_date"])
                        - date.fromisoformat(r["start_date"])).days + 1)
    check(f"row {d}: spend positive", r["spend_pkr"] > 0)
    check(f"row {d}: impressions positive", r["impressions"] > 0)
    check(f"row {d}: clicks positive", r["clicks"] > 0)
    check(f"row {d}: conversions positive", r["conversions"] > 0)
    check(f"row {d}: revenue positive", r["revenue_pkr"] > 0)
    check(f"row {d}: clicks <= impressions",
          r["clicks"] <= r["impressions"])
    check(f"row {d}: conversions <= clicks",
          r["conversions"] <= r["clicks"])
    check(f"row {d}: revenue is a multiple of 100",
          r["revenue_pkr"] % 100 == 0)

s = json.loads(SUMMARY.read_text(encoding="utf-8"))

# ── per-campaign recompute vs summary.json ──────────────────────────
campaigns = []
for r in rows:
    m = metrics(r["spend_pkr"], r["impressions"], r["clicks"],
                r["conversions"], r["revenue_pkr"], r["days"])
    campaigns.append({**r, **m, "call": call_for(m["roas"],
                                                 r["spend_pkr"])})
by_id = {c["campaign_id"]: c for c in s["campaigns"]}
check("json: 12 campaigns in summary", len(s["campaigns"]) == 12)
for c in campaigns:
    j = by_id[c["campaign_id"]]
    check(f"json {c['campaign_id']}: name", j["campaign_name"]
          == c["campaign_name"])
    check(f"json {c['campaign_id']}: channel", j["channel"]
          == c["channel"])
    for key in ("ctr_pct", "cpc_pkr", "cpa_pkr",
                "conversion_rate_pct", "roas",
                "revenue_per_click_pkr", "gross_profit_pkr",
                "payback_days"):
        check(f"json {c['campaign_id']}: {key}",
              close(j[key], c[key], 0.02),
              f"json={j[key]} recomputed={c[key]}")
    check(f"json {c['campaign_id']}: call", j["call"] == c["call"],
          f"json={j['call']} recomputed={c['call']}")

# ── story shape (the report must actually have a story) ─────────────
roases = [c["roas"] for c in campaigns]
check("story: a clear winner exists (ROAS >= 5)",
      max(roases) >= 5, f"max ROAS {max(roases)}")
check("story: a clear loser exists (ROAS < 1)",
      min(roases) < 1, f"min ROAS {min(roases)}")
check("story: at least 2 Kill calls",
      sum(1 for c in campaigns if c["call"] == "Kill") >= 2)
check("story: at least 1 Scale call",
      sum(1 for c in campaigns if c["call"] == "Scale") >= 1)
check("story: a tiny-scale high-ROAS campaign exists",
      any(c["roas"] >= 10 and c["spend_pkr"] <= 50_000
          for c in campaigns))
tiktok = [c for c in campaigns if c["channel"] == "TikTok Ads"]
check("story: TikTok has the most clicks of any channel",
      sum(c["clicks"] for c in tiktok)
      > max(sum(c["clicks"] for c in campaigns
                if c["channel"] == ch) for ch in CHANNELS
            if ch != "TikTok Ads"))
tiktok_ch = [c for c in s["channels"]
             if c["channel"] == "TikTok Ads"][0]
check("story: TikTok conversion rate is the lowest of any channel",
      tiktok_ch["conversion_rate_pct"]
      < min(c["conversion_rate_pct"] for c in s["channels"]
            if c["channel"] != "TikTok Ads"),
      f"TikTok {tiktok_ch['conversion_rate_pct']}%")
for c in campaigns:
    if c["call"] == "Kill":
        check(f"story {c['campaign_id']}: Kill payback exceeds flight",
              c["payback_days"] > c["days"],
              f"payback {c['payback_days']} vs {c['days']} days")

# ── totals ──────────────────────────────────────────────────────────
total = metrics(sum(r["spend_pkr"] for r in rows),
                sum(r["impressions"] for r in rows),
                sum(r["clicks"] for r in rows),
                sum(r["conversions"] for r in rows),
                sum(r["revenue_pkr"] for r in rows), 181)
check("total: spend", s["total_spend_pkr"] == total["spend_pkr"])
check("total: revenue", s["total_revenue_pkr"] == total["revenue_pkr"])
check("total: overall ROAS",
      close(s["overall_roas"], total["roas"], 0.01))
check("total: impressions",
      s["total_impressions"] == total["impressions"])
check("total: clicks", s["total_clicks"] == total["clicks"])
check("total: conversions",
      s["total_conversions"] == total["conversions"])
check("total: CTR", close(s["overall_ctr_pct"], total["ctr_pct"]))
check("total: CPC", close(s["overall_cpc_pkr"], total["cpc_pkr"]))
check("total: CPA", close(s["overall_cpa_pkr"], total["cpa_pkr"]))
check("total: conversion rate",
      close(s["overall_conversion_rate_pct"],
            total["conversion_rate_pct"]))
check("total: revenue per click",
      close(s["revenue_per_click_pkr"],
            total["revenue_per_click_pkr"]))
check("total: gross profit",
      close(s["gross_profit_pkr"], total["gross_profit_pkr"]))
check("total: blended payback",
      close(s["blended_payback_days"],
            r2(total["spend_pkr"] * 181 / total["gross_profit_pkr"])))
best = max(campaigns, key=lambda x: x["roas"])
worst = min(campaigns, key=lambda x: x["roas"])
check("total: best campaign", s["best_campaign"]
      == best["campaign_name"])
check("total: best campaign ROAS",
      close(s["best_campaign_roas"], best["roas"]))
check("total: worst campaign", s["worst_campaign"]
      == worst["campaign_name"])
check("total: worst campaign ROAS",
      close(s["worst_campaign_roas"], worst["roas"]))
check("total: kill count",
      s["kill_count"] == sum(1 for c in campaigns
                             if c["call"] == "Kill"))
check("total: scale count",
      s["scale_count"] == sum(1 for c in campaigns
                              if c["call"] == "Scale"))
check("total: keep count",
      s["keep_count"] == sum(1 for c in campaigns
                             if c["call"].startswith("Keep")))
check("total: kill spend",
      s["kill_spend_pkr"] == sum(c["spend_pkr"] for c in campaigns
                                 if c["call"] == "Kill"))
check("total: kill revenue",
      s["kill_revenue_pkr"] == sum(c["revenue_pkr"] for c in campaigns
                                   if c["call"] == "Kill"))
check("total: scale avg ROAS",
      close(s["scale_avg_roas"],
            r2(sum(c["revenue_pkr"] for c in campaigns
                   if c["call"] == "Scale")
               / sum(c["spend_pkr"] for c in campaigns
                     if c["call"] == "Scale"))))
check("total: reallocated revenue",
      s["reallocated_revenue_pkr"]
      == int(round(s["kill_spend_pkr"] * s["scale_avg_roas"])))

# ── channels ────────────────────────────────────────────────────────
by_ch = {c["channel"]: c for c in s["channels"]}
check("json: 7 channels in summary", len(s["channels"]) == 7)
for ch in CHANNELS:
    g = [c for c in campaigns if c["channel"] == ch]
    m = metrics(sum(c["spend_pkr"] for c in g),
                sum(c["impressions"] for c in g),
                sum(c["clicks"] for c in g),
                sum(c["conversions"] for c in g),
                sum(c["revenue_pkr"] for c in g),
                sum(c["days"] for c in g))
    j = by_ch[ch]
    check(f"channel {ch}: campaigns", j["campaigns"] == len(g))
    for key in ("spend_pkr", "impressions", "clicks", "conversions",
                "revenue_pkr"):
        check(f"channel {ch}: {key}", j[key] == m[key],
              f"json={j[key]} recomputed={m[key]}")
    for key in ("ctr_pct", "cpc_pkr", "cpa_pkr",
                "conversion_rate_pct", "roas",
                "revenue_per_click_pkr", "gross_profit_pkr",
                "payback_days"):
        check(f"channel {ch}: {key}", close(j[key], m[key], 0.02),
              f"json={j[key]} recomputed={m[key]}")
ch_by_roas = sorted(s["channels"], key=lambda x: -x["roas"])
check("channel: best by ROAS",
      s["best_channel_roas"] == ch_by_roas[0]["channel"])
check("channel: worst by ROAS",
      s["worst_channel_roas"] == ch_by_roas[-1]["channel"])
check("channel: best by revenue",
      s["best_channel_revenue"]
      == max(s["channels"], key=lambda x: x["revenue_pkr"])["channel"])
check("channel: spend sums to total",
      sum(c["spend_pkr"] for c in s["channels"])
      == s["total_spend_pkr"])
check("channel: revenue sums to total",
      sum(c["revenue_pkr"] for c in s["channels"])
      == s["total_revenue_pkr"])

# ── monthly (day-share) ─────────────────────────────────────────────
def monthly_split():
    out = []
    for mo in MONTHS:
        first, lastd = month_bounds(mo)
        spend = revenue = conv = 0.0
        active = 0
        for c in campaigns:
            st = date.fromisoformat(c["start_date"])
            en = date.fromisoformat(c["end_date"])
            overlap = max(0, (min(en, lastd) - max(st, first)).days + 1)
            if overlap:
                active += 1
                spend += c["spend_pkr"] * overlap / c["days"]
                revenue += c["revenue_pkr"] * overlap / c["days"]
                conv += c["conversions"] * overlap / c["days"]
        out.append({"month": mo, "campaigns_active": active,
                    "spend_pkr": int(round(spend)),
                    "revenue_pkr": int(round(revenue)),
                    "conversions": int(round(conv)),
                    "roas": r2(revenue / spend) if spend else 0.0})
    return out


monthly = monthly_split()
check("monthly: 6 months", len(s["monthly"]) == 6)
for m_exp, m_j in zip(monthly, s["monthly"]):
    mo = m_exp["month"]
    check(f"monthly {mo}: active", m_j["campaigns_active"]
          == m_exp["campaigns_active"])
    check(f"monthly {mo}: spend",
          abs(m_j["spend_pkr"] - m_exp["spend_pkr"]) <= 1,
          f"json={m_j['spend_pkr']} recomputed={m_exp['spend_pkr']}")
    check(f"monthly {mo}: revenue",
          abs(m_j["revenue_pkr"] - m_exp["revenue_pkr"]) <= 1)
    check(f"monthly {mo}: conversions",
          abs(m_j["conversions"] - m_exp["conversions"]) <= 1)
    check(f"monthly {mo}: ROAS",
          close(m_j["roas"], m_exp["roas"], 0.01))
check("monthly: spend sums to total",
      sum(m["spend_pkr"] for m in s["monthly"])
      == s["total_spend_pkr"])
check("monthly: revenue sums to total",
      sum(m["revenue_pkr"] for m in s["monthly"])
      == s["total_revenue_pkr"])

# ── funnel ──────────────────────────────────────────────────────────
check("funnel: 3 stages", len(s["funnel"]) == 3)
check("funnel: impressions", s["funnel"][0]["count"]
      == s["total_impressions"])
check("funnel: clicks", s["funnel"][1]["count"] == s["total_clicks"])
check("funnel: conversions", s["funnel"][2]["count"]
      == s["total_conversions"])
check("funnel: click conversion = CTR",
      close(s["funnel"][1]["conversion_from_previous"],
            s["overall_ctr_pct"]))
check("funnel: order conversion = conversion rate",
      close(s["funnel"][2]["conversion_from_previous"],
            s["overall_conversion_rate_pct"]))

# ── processed CSVs ──────────────────────────────────────────────────
def read_csv(path):
    return list(csv.DictReader(open(path, encoding="utf-8")))


enriched = read_csv(PROC / "campaigns_enriched.csv")
check("processed: enriched has 12 rows", len(enriched) == 12)
for row, c in zip(sorted(enriched, key=lambda r: r["campaign_id"]),
                  sorted(campaigns, key=lambda r: r["campaign_id"])):
    check(f"processed enriched {c['campaign_id']}: ROAS",
          close(row["roas"], c["roas"], 0.02))
    check(f"processed enriched {c['campaign_id']}: call",
          row["call"] == c["call"])
ch_csv = read_csv(PROC / "channel_summary.csv")
check("processed: channel summary has 7 rows", len(ch_csv) == 7)
for row in ch_csv:
    j = by_ch[row["channel"]]
    check(f"processed channel {row['channel']}: ROAS",
          close(row["roas"], j["roas"], 0.02))
mo_csv = read_csv(PROC / "monthly_summary.csv")
check("processed: monthly summary has 6 rows", len(mo_csv) == 6)
for row, m_exp in zip(mo_csv, monthly):
    check(f"processed monthly {m_exp['month']}: spend",
          abs(int(row["spend_pkr"]) - m_exp["spend_pkr"]) <= 1)
fn_csv = read_csv(PROC / "funnel_summary.csv")
check("processed: funnel summary has 3 rows", len(fn_csv) == 3)

# ── markdown report ─────────────────────────────────────────────────
check("md exists", MD.exists())
md = MD.read_text(encoding="utf-8") if MD.exists() else ""
for needle in (f"PKR {s['total_spend_pkr']:,}",
               f"PKR {s['total_revenue_pkr']:,}",
               f"{s['overall_roas']}",
               s["best_campaign"], s["worst_campaign"],
               "SAMPLE"):
    check(f"md contains {needle!r}", needle in md)

# ── workbook recalc ─────────────────────────────────────────────────
# always re-run the recalc, so the check count is stable run to run
RECALC_DIR.mkdir(parents=True, exist_ok=True)
res = subprocess.run(
    ["soffice", "--headless", "--calc", "--convert-to", "xlsx",
     "--outdir", str(RECALC_DIR), str(WORKBOOK)],
    capture_output=True, text=True, timeout=300)
check("recalc: soffice ran", res.returncode == 0, res.stderr[-300:])
check("recalc file exists", RECALC.exists())
wb = load_workbook(RECALC, data_only=True)

# no error values anywhere in the recalc
errs = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("#"):
                errs.append(f"{ws.title}!{cell.coordinate}={v}")
check("recalc: no error values", not errs, "; ".join(errs[:5]))

db = wb["Dashboard"]
check("recalc Dashboard total campaigns",
      db.cell(row=4, column=2).value == 12)
check("recalc Dashboard total spend",
      db.cell(row=5, column=2).value == s["total_spend_pkr"])
check("recalc Dashboard total revenue",
      db.cell(row=6, column=2).value == s["total_revenue_pkr"])
check("recalc Dashboard overall ROAS",
      close(db.cell(row=7, column=2).value, s["overall_roas"], 0.01))
check("recalc Dashboard impressions",
      db.cell(row=8, column=2).value == s["total_impressions"])
check("recalc Dashboard clicks",
      db.cell(row=9, column=2).value == s["total_clicks"])
check("recalc Dashboard conversions",
      db.cell(row=10, column=2).value == s["total_conversions"])
check("recalc Dashboard CTR",
      close(db.cell(row=11, column=2).value,
            s["overall_ctr_pct"] / 100, 0.0001))
check("recalc Dashboard CPC",
      close(db.cell(row=12, column=2).value, s["overall_cpc_pkr"]))
check("recalc Dashboard CPA",
      close(db.cell(row=13, column=2).value, s["overall_cpa_pkr"]))
check("recalc Dashboard conversion rate",
      close(db.cell(row=14, column=2).value,
            s["overall_conversion_rate_pct"] / 100, 0.0001))
check("recalc Dashboard revenue per click",
      close(db.cell(row=15, column=2).value,
            s["revenue_per_click_pkr"]))
check("recalc Dashboard gross profit",
      close(db.cell(row=16, column=2).value, s["gross_profit_pkr"]))
check("recalc Dashboard blended payback",
      close(db.cell(row=17, column=2).value,
            s["blended_payback_days"], 0.05))
check("recalc Dashboard channels",
      db.cell(row=18, column=2).value == 7)
r0 = 20
check("recalc Dashboard best campaign",
      db.cell(row=r0 + 1, column=2).value == s["best_campaign"],
      str(db.cell(row=r0 + 1, column=2).value))
check("recalc Dashboard best campaign ROAS",
      close(db.cell(row=r0 + 1, column=3).value,
            s["best_campaign_roas"], 0.01))
check("recalc Dashboard worst campaign",
      db.cell(row=r0 + 2, column=2).value == s["worst_campaign"],
      str(db.cell(row=r0 + 2, column=2).value))
check("recalc Dashboard worst campaign ROAS",
      close(db.cell(row=r0 + 2, column=3).value,
            s["worst_campaign_roas"], 0.01))
check("recalc Dashboard best channel ROAS",
      db.cell(row=r0 + 3, column=2).value == s["best_channel_roas"],
      str(db.cell(row=r0 + 3, column=2).value))
check("recalc Dashboard best channel ROAS value",
      close(db.cell(row=r0 + 3, column=3).value,
            s["best_channel_roas_value"], 0.01))
check("recalc Dashboard worst channel ROAS",
      db.cell(row=r0 + 4, column=2).value == s["worst_channel_roas"],
      str(db.cell(row=r0 + 4, column=2).value))
check("recalc Dashboard worst channel ROAS value",
      close(db.cell(row=r0 + 4, column=3).value,
            s["worst_channel_roas_value"], 0.01))
check("recalc Dashboard best channel revenue",
      db.cell(row=r0 + 5, column=2).value == s["best_channel_revenue"],
      str(db.cell(row=r0 + 5, column=2).value))
check("recalc Dashboard best channel revenue value",
      db.cell(row=r0 + 5, column=3).value
      == s["best_channel_revenue_value"])
check("recalc Dashboard Kill count",
      db.cell(row=r0 + 6, column=2).value == s["kill_count"])
check("recalc Dashboard Kill spend",
      db.cell(row=r0 + 6, column=3).value == s["kill_spend_pkr"])
check("recalc Dashboard Scale count",
      db.cell(row=r0 + 7, column=2).value == s["scale_count"])
check("recalc Dashboard Keep count",
      db.cell(row=r0 + 8, column=2).value == s["keep_count"])
check("recalc Dashboard Kill revenue",
      db.cell(row=r0 + 9, column=2).value == s["kill_revenue_pkr"])

# Campaigns sheet formulas recalculated — rows sorted by start date
cp = wb["Campaigns"]
sorted_camps = sorted(campaigns,
                      key=lambda c: (c["start_date"], c["campaign_id"]))
check("recalc Campaigns rows", cp.max_row == 13,
      f"max_row={cp.max_row}")
for k, c in enumerate(sorted_camps):
    r = 2 + k
    check(f"recalc Campaigns {c['campaign_id']} id",
          cp.cell(row=r, column=1).value == c["campaign_id"])
    check(f"recalc Campaigns {c['campaign_id']} days",
          cp.cell(row=r, column=6).value == c["days"])
    check(f"recalc Campaigns {c['campaign_id']} CTR",
          close(cp.cell(row=r, column=12).value,
                c["ctr_pct"] / 100, 0.0001))
    check(f"recalc Campaigns {c['campaign_id']} CPC",
          close(cp.cell(row=r, column=13).value, c["cpc_pkr"], 0.02))
    check(f"recalc Campaigns {c['campaign_id']} CPA",
          close(cp.cell(row=r, column=14).value, c["cpa_pkr"], 0.02))
    check(f"recalc Campaigns {c['campaign_id']} conversion rate",
          close(cp.cell(row=r, column=15).value,
                c["conversion_rate_pct"] / 100, 0.0001))
    check(f"recalc Campaigns {c['campaign_id']} ROAS",
          close(cp.cell(row=r, column=16).value, c["roas"], 0.01))
    check(f"recalc Campaigns {c['campaign_id']} revenue per click",
          close(cp.cell(row=r, column=17).value,
                c["revenue_per_click_pkr"], 0.02))
    check(f"recalc Campaigns {c['campaign_id']} gross profit",
          close(cp.cell(row=r, column=18).value,
                c["gross_profit_pkr"], 0.5))
    check(f"recalc Campaigns {c['campaign_id']} payback",
          close(cp.cell(row=r, column=19).value,
                c["payback_days"], 0.05))
    check(f"recalc Campaigns {c['campaign_id']} call",
          cp.cell(row=r, column=20).value == c["call"],
          str(cp.cell(row=r, column=20).value))
    # monthly day-share split
    for mi, mo in enumerate(MONTHS):
        first, lastd = month_bounds(mo)
        st = date.fromisoformat(c["start_date"])
        en = date.fromisoformat(c["end_date"])
        overlap = max(0, (min(en, lastd) - max(st, first)).days + 1)
        exp_spend = int(round(c["spend_pkr"] * overlap / c["days"]))
        exp_rev = int(round(c["revenue_pkr"] * overlap / c["days"]))
        exp_conv = int(round(c["conversions"] * overlap / c["days"]))
        check(f"recalc Campaigns {c['campaign_id']} spend {mo}",
              abs(cp.cell(row=r, column=21 + mi).value - exp_spend)
              <= 1,
              f"got {cp.cell(row=r, column=21 + mi).value} "
              f"expected {exp_spend}")
        check(f"recalc Campaigns {c['campaign_id']} revenue {mo}",
              abs(cp.cell(row=r, column=27 + mi).value - exp_rev)
              <= 1)
        check(f"recalc Campaigns {c['campaign_id']} conversions {mo}",
              abs(cp.cell(row=r, column=33 + mi).value - exp_conv)
              <= 1)

# Channels sheet
ch = wb["Channels"]
for j, label in enumerate(CHANNELS, start=2):
    jx = by_ch[label]
    check(f"recalc Channels {label} campaigns",
          ch.cell(row=j, column=2).value == jx["campaigns"])
    check(f"recalc Channels {label} spend",
          ch.cell(row=j, column=3).value == jx["spend_pkr"])
    check(f"recalc Channels {label} revenue",
          ch.cell(row=j, column=7).value == jx["revenue_pkr"])
    check(f"recalc Channels {label} CTR",
          close(ch.cell(row=j, column=8).value,
                jx["ctr_pct"] / 100, 0.0001))
    check(f"recalc Channels {label} CPC",
          close(ch.cell(row=j, column=9).value, jx["cpc_pkr"], 0.02))
    check(f"recalc Channels {label} CPA",
          close(ch.cell(row=j, column=10).value, jx["cpa_pkr"], 0.02))
    check(f"recalc Channels {label} conversion rate",
          close(ch.cell(row=j, column=11).value,
                jx["conversion_rate_pct"] / 100, 0.0001))
    check(f"recalc Channels {label} ROAS",
          close(ch.cell(row=j, column=12).value, jx["roas"], 0.01))
    check(f"recalc Channels {label} gross profit",
          close(ch.cell(row=j, column=14).value,
                jx["gross_profit_pkr"], 0.5))
    check(f"recalc Channels {label} payback",
          close(ch.cell(row=j, column=16).value,
                jx["payback_days"], 0.05))
check("recalc Channels total spend",
      ch.cell(row=9, column=3).value == s["total_spend_pkr"])
check("recalc Channels total revenue",
      ch.cell(row=9, column=7).value == s["total_revenue_pkr"])
check("recalc Channels total ROAS",
      close(ch.cell(row=9, column=12).value, s["overall_roas"], 0.01))

# Monthly Trend sheet
mt = wb["Monthly Trend"]
for k, m_exp in enumerate(monthly):
    r = 2 + k
    check(f"recalc Monthly {m_exp['month']} active",
          mt.cell(row=r, column=2).value == m_exp["campaigns_active"],
          f"got {mt.cell(row=r, column=2).value}")
    check(f"recalc Monthly {m_exp['month']} spend",
          abs(mt.cell(row=r, column=3).value - m_exp["spend_pkr"])
          <= 1)
    check(f"recalc Monthly {m_exp['month']} revenue",
          abs(mt.cell(row=r, column=4).value - m_exp["revenue_pkr"])
          <= 1)
    check(f"recalc Monthly {m_exp['month']} conversions",
          abs(mt.cell(row=r, column=5).value - m_exp["conversions"])
          <= 1)
    check(f"recalc Monthly {m_exp['month']} ROAS",
          close(mt.cell(row=r, column=6).value, m_exp["roas"], 0.01))
# monthly totals: each month rounds per campaign, so the workbook
# total can differ from the rounded month totals by a few rupees
check("recalc Monthly total spend",
      abs(mt.cell(row=8, column=3).value - s["total_spend_pkr"]) <= 6,
      f"got {mt.cell(row=8, column=3).value}")
check("recalc Monthly total revenue",
      abs(mt.cell(row=8, column=4).value - s["total_revenue_pkr"])
      <= 6,
      f"got {mt.cell(row=8, column=4).value}")
check("recalc Monthly total ROAS",
      close(mt.cell(row=8, column=6).value, s["overall_roas"], 0.01))

# Funnel sheet
fn = wb["Funnel"]
check("recalc Funnel impressions",
      fn.cell(row=3, column=2).value == s["total_impressions"])
check("recalc Funnel clicks",
      fn.cell(row=4, column=2).value == s["total_clicks"])
check("recalc Funnel conversions",
      fn.cell(row=5, column=2).value == s["total_conversions"])
check("recalc Funnel CTR",
      close(fn.cell(row=4, column=3).value,
            s["overall_ctr_pct"] / 100, 0.0001))
check("recalc Funnel conversion rate",
      close(fn.cell(row=5, column=3).value,
            s["overall_conversion_rate_pct"] / 100, 0.0001))
check("recalc Funnel revenue",
      fn.cell(row=6, column=2).value == s["total_revenue_pkr"])
for j, label in enumerate(CHANNELS, start=10):
    jx = by_ch[label]
    check(f"recalc Funnel channel {label} impressions",
          fn.cell(row=j, column=2).value == jx["impressions"])
    check(f"recalc Funnel channel {label} conversions",
          fn.cell(row=j, column=4).value == jx["conversions"])
    check(f"recalc Funnel channel {label} ROAS",
          close(fn.cell(row=j, column=9).value, jx["roas"], 0.01))
check("recalc Funnel total ROAS",
      close(fn.cell(row=17, column=9).value, s["overall_roas"], 0.01))

# ── PDF + charts ────────────────────────────────────────────────────
check("pdf exists", PDF.exists())
check("pdf size > 50KB", PDF.stat().st_size > 50_000,
      f"{PDF.stat().st_size} bytes")
res = subprocess.run(["pdfinfo", str(PDF)], capture_output=True,
                     text=True)
pages = None
for line in res.stdout.splitlines():
    if line.startswith("Pages:"):
        pages = int(line.split(":")[1].strip())
check("pdf is 2 pages", pages == 2, f"pages={pages}")
for chart_path, label in ((CHART_CH, "channel_roas"),
                          (CHART_MO, "monthly")):
    check(f"chart {label} exists", chart_path.exists())
    data = chart_path.read_bytes() if chart_path.exists() else b""
    check(f"chart {label} is a PNG",
          data[:8] == b"\x89PNG\r\n\x1a\n")
    check(f"chart {label} size > 10KB", len(data) > 10_000,
          f"{len(data)} bytes")

# ── README + PACKAGE-NOTE carry the same headline numbers ───────────
for path, label in ((README, "README"), (PACKAGE, "PACKAGE-NOTE")):
    check(f"{label} exists", path.exists())
    if path.exists():
        txt = path.read_text(encoding="utf-8")
        for needle in (f"{s['total_campaigns']}",
                       f"PKR {s['total_spend_pkr']:,}",
                       f"PKR {s['total_revenue_pkr']:,}",
                       f"{s['overall_roas']}",
                       "SAMPLE"):
            check(f"{label} contains {needle!r}", needle in txt)
        check(f"{label} says SAMPLE, not real client work",
              "SAMPLE" in txt and "fictional" in txt.lower())

print(f"\n{len(ran) - len(fails)}/{len(ran)} checks passed, "
      f"{len(fails)} failed")
if fails:
    print("Failed checks:")
    for f in fails:
        print("  - " + f)
    sys.exit(1)
print("ALL CHECKS PASSED")
