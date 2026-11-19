#!/usr/bin/env python3
"""P23 — generate the SAMPLE campaign file for StyleHub.pk (SAMPLE).

Fictional Pakistani fashion e-commerce store, 12 campaigns across
7 channels, Thu 2026-01-01 -> Tue 2026-06-30 (181 days). Seeded, so
the file is reproducible. One file, one row per campaign:

  data/raw/campaigns_SAMPLE.csv — grain: one campaign
      (campaign_id, campaign_name, channel, start_date, end_date,
       days, spend_pkr, impressions, clicks, conversions,
       revenue_pkr)

How the demo data is built (and why the report can still find real
things in it): each campaign has fixed planning rates — cost per
click, click-through rate, conversion rate and average order value —
chosen per campaign so the account has a real story: a clear winner
(Influencer Fest, ROAS ~6.7), two clear losers to kill (Flash Sale
SMS and Brand Awareness — Meta, ROAS < 1, payback longer than the
flight), one traffic-heavy but poor-converting channel (TikTok Ads:
the most clicks in the account at the lowest conversion rates) and
two tiny-scale, high-ROAS email campaigns. A fixed seed jitters each
rate by at most +/-2%, so re-runs reproduce identical data and no
call flips. Derived: clicks = spend / CPC, impressions = clicks /
CTR, conversions = clicks x conversion rate, revenue = conversions
x AOV (rounded to the nearest PKR 100).

    python3 src/generate_data.py
"""
import csv
import random
from datetime import date
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"

SEED = 20261119
START = date(2026, 1, 1)
END = date(2026, 6, 30)

# (name, channel, start, end, spend PKR, CPC PKR, CTR, conv rate, AOV PKR)
CAMPAIGNS = [
    ("Eid Lawn Launch", "Google Search", "2026-03-01", "2026-04-15",
     850_000, 42, 0.065, 0.032, 4800),
    ("Summer Sale — Search", "Google Search", "2026-05-01",
     "2026-06-15", 620_000, 38, 0.060, 0.030, 4200),
    ("New Arrivals — Shopping", "Google Shopping", "2026-01-10",
     "2026-03-31", 540_000, 28, 0.045, 0.026, 3900),
    ("Eid Collection — Shopping", "Google Shopping", "2026-04-01",
     "2026-05-20", 480_000, 30, 0.048, 0.020, 4100),
    ("Ramadan Festive Push", "Meta Ads", "2026-02-20", "2026-04-10",
     720_000, 22, 0.028, 0.012, 4500),
    ("Influencer Fest — Ayesha", "Influencer", "2026-03-15",
     "2026-04-30", 300_000, 18, 0.035, 0.030, 4000),
    ("TikTok Viral Challenge", "TikTok Ads", "2026-02-01",
     "2026-03-31", 950_000, 8, 0.055, 0.0035, 3800),
    ("TikTok Eid Dance", "TikTok Ads", "2026-04-15", "2026-05-31",
     410_000, 9, 0.050, 0.009, 3600),
    ("Eid Email Blast", "Email", "2026-03-20", "2026-04-20",
     45_000, 12, 0.060, 0.040, 4200),
    ("Winback — Lapsed Buyers", "Email", "2026-05-05", "2026-06-10",
     38_000, 13, 0.070, 0.050, 4600),
    ("Flash Sale SMS", "SMS/WhatsApp", "2026-06-01", "2026-06-20",
     60_000, 26, 0.090, 0.007, 2800),
    ("Brand Awareness — Meta", "Meta Ads", "2026-01-15", "2026-02-28",
     680_000, 15, 0.018, 0.0032, 4000),
]

CHANNELS = ["Google Search", "Google Shopping", "Meta Ads",
            "TikTok Ads", "Influencer", "Email", "SMS/WhatsApp"]


def jitter(rng, value):
    return value * rng.uniform(0.98, 1.02)


def main():
    rng = random.Random(SEED)
    RAW.mkdir(parents=True, exist_ok=True)
    rows = []
    for i, (name, channel, start, end, spend, cpc, ctr, conv, aov) \
            in enumerate(CAMPAIGNS, start=1):
        s = date.fromisoformat(start)
        e = date.fromisoformat(end)
        days = (e - s).days + 1
        cpc_j = jitter(rng, cpc)
        ctr_j = jitter(rng, ctr)
        conv_j = jitter(rng, conv)
        aov_j = jitter(rng, aov)
        clicks = int(round(spend / cpc_j))
        impressions = int(round(clicks / ctr_j))
        conversions = int(round(clicks * conv_j))
        revenue = int(round(conversions * aov_j / 100.0)) * 100
        rows.append({
            "campaign_id": f"CMP-{i:02d}",
            "campaign_name": name,
            "channel": channel,
            "start_date": start,
            "end_date": end,
            "days": days,
            "spend_pkr": spend,
            "impressions": impressions,
            "clicks": clicks,
            "conversions": conversions,
            "revenue_pkr": revenue,
        })

    out = RAW / "campaigns_SAMPLE.csv"
    fields = list(rows[0].keys())
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    tot_spend = sum(r["spend_pkr"] for r in rows)
    tot_rev = sum(r["revenue_pkr"] for r in rows)
    print(f"Wrote {out} — {len(rows)} campaigns, "
          f"total spend PKR {tot_spend:,}, total revenue "
          f"PKR {tot_rev:,}, overall ROAS {tot_rev / tot_spend:.2f}")


if __name__ == "__main__":
    main()
