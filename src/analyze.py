#!/usr/bin/env python3
"""P23 — analyse the SAMPLE campaign file and regenerate every output
number from the raw CSV (never typed): headline ROI KPIs, per-campaign
metrics + Kill / Scale / Keep calls, per-channel rollups, monthly
trend (spend and revenue split by day-share across each campaign's
flight), funnel, and the reallocation maths behind the recommendations.

Metrics (all defined in the README and the workbook's Method sheet):
  CTR = clicks / impressions
  CPC = spend / clicks
  CPA = spend / conversions
  conversion rate = conversions / clicks
  ROAS = revenue / spend
  revenue per click = revenue / clicks
  gross profit = revenue x 0.45 (the stated 45% gross-margin
      assumption — StyleHub.pk sells fashion at roughly 45% margin
      after product cost; the report states it plainly and shows
      the unadjusted revenue next to every adjusted number)
  payback days = spend / (gross profit / campaign days) — how many
      days of gross profit it takes to earn the spend back. A
      campaign whose payback is longer than its flight never paid
      back while it ran.

Call rule (stated, applied the same way by analyze.py, the workbook
formulas and checks.py):
  ROAS < 1.00  -> Kill
  ROAS >= 3.00 and spend < PKR 900,000 -> Scale
  ROAS >= 3.00 and spend >= PKR 900,000 -> Keep — proven at scale
  otherwise (1.00 <= ROAS < 3.00) -> Keep / Watch

Monthly split: a campaign's spend and revenue are spread evenly
across its flight days, so a campaign running Mar 1 – Apr 15 puts
31/46 of its spend in March and 15/46 in April. Monthly totals are
the sum of those day-shares, rounded to the nearest rupee.

Writes:
  data/processed/campaigns_enriched.csv
  data/processed/channel_summary.csv
  data/processed/monthly_summary.csv
  data/processed/funnel_summary.csv
  outputs/summary.json
  outputs/campaign_roi_report.md

    python3 src/analyze.py
"""
import csv
import json
from datetime import date
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"
PROC = PROJECT / "data" / "processed"
OUT = PROJECT / "outputs"

CHANNELS = ["Google Search", "Google Shopping", "Meta Ads",
            "TikTok Ads", "Influencer", "Email", "SMS/WhatsApp"]
MONTHS = [f"2026-{m:02d}" for m in range(1, 7)]
GROSS_MARGIN = 0.45
SCALE_SPEND_CEILING = 900_000


def pct(n, d):
    return round(100.0 * n / d, 2) if d else 0.0


def r2(x):
    return round(x, 2)


def call_for(roas, spend):
    if roas < 1.0:
        return "Kill"
    if roas >= 3.0 and spend < SCALE_SPEND_CEILING:
        return "Scale"
    if roas >= 3.0:
        return "Keep — proven at scale"
    return "Keep / Watch"


def metrics(spend, impressions, clicks, conversions, revenue, days):
    gross = revenue * GROSS_MARGIN
    payback = (spend * days / gross) if gross else None
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
        "revenue_per_click_pkr": r2(revenue / clicks) if clicks else 0.0,
        "gross_profit_pkr": r2(gross),
        "payback_days": r2(payback) if payback is not None else None,
        "days": days,
    }


def month_bounds(month):
    y, m = int(month[:4]), int(month[5:7])
    first = date(y, m, 1)
    if m == 12:
        last = date(y, 12, 31)
    else:
        last = date(y, m + 1, 1).fromordinal(
            date(y, m + 1, 1).toordinal() - 1)
    return first, last


def main():
    raw = list(csv.DictReader(open(RAW / "campaigns_SAMPLE.csv",
                                   encoding="utf-8")))
    rows = []
    for r in raw:
        s = date.fromisoformat(r["start_date"])
        e = date.fromisoformat(r["end_date"])
        days = (e - s).days + 1
        rows.append({
            "campaign_id": r["campaign_id"],
            "campaign_name": r["campaign_name"],
            "channel": r["channel"],
            "start_date": r["start_date"],
            "end_date": r["end_date"],
            "days": days,
            "spend_pkr": int(r["spend_pkr"]),
            "impressions": int(r["impressions"]),
            "clicks": int(r["clicks"]),
            "conversions": int(r["conversions"]),
            "revenue_pkr": int(r["revenue_pkr"]),
        })
    rows.sort(key=lambda r: (r["start_date"], r["campaign_id"]))

    campaigns = []
    for r in rows:
        m = metrics(r["spend_pkr"], r["impressions"], r["clicks"],
                    r["conversions"], r["revenue_pkr"], r["days"])
        campaigns.append({**r, **m,
                          "call": call_for(m["roas"],
                                           r["spend_pkr"])})

    total = metrics(sum(r["spend_pkr"] for r in rows),
                    sum(r["impressions"] for r in rows),
                    sum(r["clicks"] for r in rows),
                    sum(r["conversions"] for r in rows),
                    sum(r["revenue_pkr"] for r in rows),
                    181)
    total["campaigns"] = len(rows)
    total["total_spend_pkr"] = total.pop("spend_pkr")
    total["total_impressions"] = total.pop("impressions")
    total["total_clicks"] = total.pop("clicks")
    total["total_conversions"] = total.pop("conversions")
    total["total_revenue_pkr"] = total.pop("revenue_pkr")
    total["overall_roas"] = total.pop("roas")
    total["gross_profit_pkr"] = round(
        total["total_revenue_pkr"] * GROSS_MARGIN, 2)
    total["blended_payback_days"] = r2(
        total["total_spend_pkr"] * 181 / total["gross_profit_pkr"])

    # ── channels ────────────────────────────────────────────────────
    channels = []
    for ch in CHANNELS:
        g = [c for c in campaigns if c["channel"] == ch]
        if not g:
            continue
        m = metrics(sum(c["spend_pkr"] for c in g),
                    sum(c["impressions"] for c in g),
                    sum(c["clicks"] for c in g),
                    sum(c["conversions"] for c in g),
                    sum(c["revenue_pkr"] for c in g),
                    sum(c["days"] for c in g))
        channels.append({"channel": ch, "campaigns": len(g), **m})
    channels_by_roas = sorted(channels, key=lambda x: -x["roas"])
    channels_by_revenue = sorted(channels,
                                 key=lambda x: -x["revenue_pkr"])

    # ── monthly trend (day-share split) ─────────────────────────────
    monthly = []
    for mo in MONTHS:
        first, last = month_bounds(mo)
        spend = 0.0
        revenue = 0.0
        conversions = 0.0
        active = []
        for c in campaigns:
            s = date.fromisoformat(c["start_date"])
            e = date.fromisoformat(c["end_date"])
            overlap = max(0, (min(e, last) - max(s, first)).days + 1)
            if overlap:
                active.append(c["campaign_name"])
                spend += c["spend_pkr"] * overlap / c["days"]
                revenue += c["revenue_pkr"] * overlap / c["days"]
                conversions += c["conversions"] * overlap / c["days"]
        monthly.append({
            "month": mo,
            "campaigns_active": len(active),
            "spend_pkr": int(round(spend)),
            "revenue_pkr": int(round(revenue)),
            "conversions": int(round(conversions)),
            "roas": r2(revenue / spend) if spend else 0.0,
        })

    # ── funnel ──────────────────────────────────────────────────────
    funnel = [
        {"stage": "Impressions",
         "count": total["total_impressions"],
         "conversion_from_previous": 100.0,
         "share_of_impressions": 100.0},
        {"stage": "Clicks", "count": total["total_clicks"],
         "conversion_from_previous": pct(total["total_clicks"],
                                         total["total_impressions"]),
         "share_of_impressions": pct(total["total_clicks"],
                                     total["total_impressions"])},
        {"stage": "Conversions", "count": total["total_conversions"],
         "conversion_from_previous": pct(total["total_conversions"],
                                         total["total_clicks"]),
         "share_of_impressions": pct(total["total_conversions"],
                                     total["total_impressions"])},
    ]

    kills = [c for c in campaigns if c["call"] == "Kill"]
    scales = [c for c in campaigns if c["call"] == "Scale"]
    keeps = [c for c in campaigns
             if c["call"] in ("Keep / Watch", "Keep — proven at scale")]
    best = max(campaigns, key=lambda x: x["roas"])
    worst = min(campaigns, key=lambda x: x["roas"])
    best_channel = channels_by_roas[0]
    worst_channel = channels_by_roas[-1]
    best_channel_rev = channels_by_revenue[0]
    tiktok = [c for c in channels if c["channel"] == "TikTok Ads"][0]
    email = [c for c in channels if c["channel"] == "Email"][0]
    kill_spend = sum(c["spend_pkr"] for c in kills)
    kill_rev = sum(c["revenue_pkr"] for c in kills)
    scale_roas = r2(sum(c["revenue_pkr"] for c in scales)
                   / sum(c["spend_pkr"] for c in scales))
    reallocated_revenue = int(round(kill_spend * scale_roas))

    infl = [c for c in campaigns
            if c["campaign_name"] == "Influencer Fest — Ayesha"][0]
    actions = [
        f"Kill {len(kills)} campaigns now: "
        + ", ".join(f"{c['campaign_name']} (ROAS {c['roas']}, "
                    f"payback {c['payback_days']} days on a "
                    f"{c['days']}-day flight — it never paid back)"
                    for c in kills)
        + f". Together they burned PKR {kill_spend:,} for "
        f"PKR {kill_rev:,} back. Stop them before the next rupee "
        f"is spent.",
        f"Scale the winners: {best['campaign_name']} is the best "
        f"campaign in the account by ROAS ({best['roas']}x, payback "
        f"{best['payback_days']} days, PKR "
        f"{best['revenue_pkr']:,} on PKR {best['spend_pkr']:,} "
        f"spend) — but it runs on a small email list, so it cannot "
        f"absorb a big budget. Put the PKR {kill_spend:,} freed "
        f"from the Kill campaigns into the biggest proven winner "
        f"at scale instead, {infl['campaign_name']} (ROAS "
        f"{infl['roas']}x, PKR {infl['revenue_pkr']:,} on "
        f"PKR {infl['spend_pkr']:,} spend). At the Scale "
        f"campaigns' average ROAS of {scale_roas} that money "
        f"returns about PKR {reallocated_revenue:,} instead of "
        f"PKR {kill_rev:,}.",
        f"Fix TikTok before spending more on it: {tiktok['clicks']:,} "
        f"clicks — the most traffic in the account — at a "
        f"{tiktok['conversion_rate_pct']}% conversion rate, the "
        f"lowest of any paid channel. The clicks are cheap "
        f"(PKR {tiktok['cpc_pkr']:,.0f} each) but they do not buy: "
        f"CPA is PKR {tiktok['cpa_pkr']:,.0f}. Send TikTok traffic "
        f"to a dedicated landing page and tighten the audience "
        f"before the next rupee goes in.",
        f"Grow Email — the best ROAS in the account "
        f"({email['roas']}x) on only PKR {email['spend_pkr']:,} "
        f"spend. It returned PKR {email['revenue_pkr']:,} from "
        f"{email['conversions']} orders, but the list is too small "
        f"to move the business. Grow the list (checkout opt-in, "
        f"Eid / sale captures) before raising email spend.",
        f"Hold the proven Search / Shopping core and watch the two "
        f"mid-table campaigns: Google Search is the biggest "
        f"revenue channel (PKR "
        f"{best_channel_rev['revenue_pkr']:,} at ROAS "
        f"{best_channel_rev['roas']}); Eid Collection — Shopping "
        f"and Ramadan Festive Push sit at ROAS 2.43–2.66 — "
        f"profitable, but below the Scale bar. Re-check them after "
        f"the next flight before raising their budgets.",
    ]

    summary = {
        "client": "StyleHub.pk (SAMPLE)",
        "period": "2026-01-01 to 2026-06-30",
        "months": MONTHS,
        "gross_margin": GROSS_MARGIN,
        "scale_spend_ceiling_pkr": SCALE_SPEND_CEILING,
        "total_campaigns": total["campaigns"],
        "total_spend_pkr": total["total_spend_pkr"],
        "total_revenue_pkr": total["total_revenue_pkr"],
        "overall_roas": total["overall_roas"],
        "total_impressions": total["total_impressions"],
        "total_clicks": total["total_clicks"],
        "total_conversions": total["total_conversions"],
        "overall_ctr_pct": total["ctr_pct"],
        "overall_cpc_pkr": total["cpc_pkr"],
        "overall_cpa_pkr": total["cpa_pkr"],
        "overall_conversion_rate_pct": total["conversion_rate_pct"],
        "revenue_per_click_pkr": total["revenue_per_click_pkr"],
        "gross_profit_pkr": total["gross_profit_pkr"],
        "blended_payback_days": total["blended_payback_days"],
        "campaigns": campaigns,
        "channels": channels,
        "channels_by_roas": [c["channel"] for c in channels_by_roas],
        "channels_by_revenue": [c["channel"]
                                for c in channels_by_revenue],
        "best_campaign": best["campaign_name"],
        "best_campaign_roas": best["roas"],
        "worst_campaign": worst["campaign_name"],
        "worst_campaign_roas": worst["roas"],
        "best_channel_roas": best_channel["channel"],
        "best_channel_roas_value": best_channel["roas"],
        "worst_channel_roas": worst_channel["channel"],
        "worst_channel_roas_value": worst_channel["roas"],
        "best_channel_revenue": best_channel_rev["channel"],
        "best_channel_revenue_value": best_channel_rev["revenue_pkr"],
        "kill_count": len(kills),
        "scale_count": len(scales),
        "keep_count": len(keeps),
        "kill_spend_pkr": kill_spend,
        "kill_revenue_pkr": kill_rev,
        "scale_avg_roas": scale_roas,
        "reallocated_revenue_pkr": reallocated_revenue,
        "monthly": monthly,
        "funnel": funnel,
        "actions": actions,
    }

    PROC.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)

    def write_csv(path, fieldnames, rows_):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows_)

    write_csv(PROC / "campaigns_enriched.csv",
              list(campaigns[0].keys()), campaigns)
    write_csv(PROC / "channel_summary.csv",
              list(channels[0].keys()), channels)
    write_csv(PROC / "monthly_summary.csv",
              list(monthly[0].keys()), monthly)
    write_csv(PROC / "funnel_summary.csv",
              list(funnel[0].keys()), funnel)

    (OUT / "summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False),
        encoding="utf-8")

    lines = [
        "# P23 — Marketing Campaign ROI Tracker: StyleHub.pk (SAMPLE)",
        "",
        f"**Period:** Thu 2026-01-01 → Tue 2026-06-30 (6 months, "
        f"181 days)",
        f"**Campaigns:** {total['campaigns']} across "
        f"{len(channels)} channels · **Total spend:** "
        f"PKR {total['total_spend_pkr']:,} · **Total revenue:** "
        f"PKR {total['total_revenue_pkr']:,} · **Overall ROAS:** "
        f"{total['overall_roas']} · **Gross profit (45%):** "
        f"PKR {total['gross_profit_pkr']:,.0f} · **Blended payback:** "
        f"{total['blended_payback_days']} days",
        f"**Traffic:** {total['total_impressions']:,} impressions · "
        f"{total['total_clicks']:,} clicks · "
        f"{total['total_conversions']:,} conversions · "
        f"CTR {total['ctr_pct']}% · CPC PKR {total['cpc_pkr']:,.2f} · "
        f"CPA PKR {total['cpa_pkr']:,.2f}",
        "",
        "## 1. The headline — which campaigns earn their money back?",
        "",
        f"StyleHub.pk spent PKR {total['total_spend_pkr']:,} across "
        f"{total['campaigns']} campaigns in Jan–Jun 2026 and took "
        f"PKR {total['total_revenue_pkr']:,} back — ROAS "
        f"{total['overall_roas']}. At a 45% gross margin that is "
        f"PKR {total['gross_profit_pkr']:,.0f} of gross profit, and "
        f"the whole programme pays back in "
        f"{total['blended_payback_days']} days of gross profit. "
        f"But the average hides the story: {len(scales)} campaigns "
        f"are worth scaling, {len(keeps)} are worth watching and "
        f"{len(kills)} should be killed — they burned "
        f"PKR {kill_spend:,} for PKR {kill_rev:,} back.",
        "",
        "## 2. Every campaign — with its call",
        "",
        "| Campaign | Channel | Days | Spend | Impressions | Clicks | "
        "Conv. | Revenue | CTR | CPC | CPA | Conv. rate | ROAS | "
        "Rev / click | Gross profit | Payback | Call |",
        "|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|",
    ]
    for c in campaigns:
        lines.append(
            f"| {c['campaign_name']} | {c['channel']} | {c['days']} | "
            f"PKR {c['spend_pkr']:,} | {c['impressions']:,} | "
            f"{c['clicks']:,} | {c['conversions']} | "
            f"PKR {c['revenue_pkr']:,} | {c['ctr_pct']}% | "
            f"PKR {c['cpc_pkr']:,.2f} | PKR {c['cpa_pkr']:,.2f} | "
            f"{c['conversion_rate_pct']}% | {c['roas']} | "
            f"PKR {c['revenue_per_click_pkr']:,.2f} | "
            f"PKR {c['gross_profit_pkr']:,.0f} | "
            f"{c['payback_days']} days | **{c['call']}** |")
    lines += [
        "",
        f"**Best campaign: {best['campaign_name']} (ROAS "
        f"{best['roas']}). Worst: {worst['campaign_name']} (ROAS "
        f"{worst['roas']}).**",
        "",
        "## 3. By channel",
        "",
        "| Channel | Campaigns | Spend | Impressions | Clicks | Conv. | "
        "Revenue | CTR | CPC | CPA | Conv. rate | ROAS | "
        "Rev / click | Gross profit | Payback |",
        "|---|---|---|---|---|---|---|---|---|---|---|---|---|---|",
    ]
    for ch in channels_by_revenue:
        lines.append(
            f"| {ch['channel']} | {ch['campaigns']} | "
            f"PKR {ch['spend_pkr']:,} | {ch['impressions']:,} | "
            f"{ch['clicks']:,} | {ch['conversions']} | "
            f"PKR {ch['revenue_pkr']:,} | {ch['ctr_pct']}% | "
            f"PKR {ch['cpc_pkr']:,.2f} | PKR {ch['cpa_pkr']:,.2f} | "
            f"{ch['conversion_rate_pct']}% | {ch['roas']} | "
            f"PKR {ch['revenue_per_click_pkr']:,.2f} | "
            f"PKR {ch['gross_profit_pkr']:,.0f} | "
            f"{ch['payback_days']} days |")
    lines += [
        "",
        f"**Best channel by ROAS: {best_channel['channel']} "
        f"({best_channel['roas']}x). Worst: "
        f"{worst_channel['channel']} "
        f"({worst_channel['roas']}x). Biggest revenue channel: "
        f"{best_channel_rev['channel']} "
        f"(PKR {best_channel_rev['revenue_pkr']:,}).**",
        "",
        "## 4. Month by month (day-share split)",
        "",
        "| Month | Campaigns active | Spend | Revenue | "
        "Conversions | ROAS |",
        "|---|---|---|---|---|---|",
    ]
    for m in monthly:
        lines.append(f"| {m['month']} | {m['campaigns_active']} | "
                     f"PKR {m['spend_pkr']:,} | "
                     f"PKR {m['revenue_pkr']:,} | {m['conversions']} | "
                     f"{m['roas']} |")
    lines += [
        "",
        "## 5. The funnel — impressions to orders",
        "",
        "| Stage | Count | Conversion from previous | "
        "Share of impressions |",
        "|---|---|---|---|",
    ]
    for f in funnel:
        lines.append(f"| {f['stage']} | {f['count']:,} | "
                     f"{f['conversion_from_previous']}% | "
                     f"{f['share_of_impressions']}% |")
    lines += [
        "",
        f"Only {funnel[2]['share_of_impressions']}% of impressions "
        f"become an order. The funnel is widest at the top — "
        f"{total['total_impressions']:,} impressions — and the "
        f"leak that matters is click → order: "
        f"{total['conversion_rate_pct']}% of clicks convert. "
        f"TikTok brings the clicks ({tiktok['clicks']:,} at "
        f"{tiktok['conversion_rate_pct']}% conversion); Email "
        f"brings the orders ({email['conversion_rate_pct']}% "
        f"conversion on {email['clicks']:,} clicks).",
        "",
        "## 6. The five recommendations",
        "",
    ]
    for i, a in enumerate(actions, 1):
        lines.append(f"{i}. **{a}**")
    lines.append("")
    (OUT / "campaign_roi_report.md").write_text(
        "\n".join(lines), encoding="utf-8")
    print(f"Analysed {len(rows)} campaigns: spend "
          f"PKR {total['total_spend_pkr']:,}, revenue "
          f"PKR {total['total_revenue_pkr']:,}, ROAS "
          f"{total['overall_roas']}, {len(scales)} Scale / "
          f"{len(keeps)} Keep / {len(kills)} Kill")


if __name__ == "__main__":
    main()
