# P23 — Marketing Campaign ROI Tracker

**Client (SAMPLE):** StyleHub.pk — fictional Pakistani fashion
e-commerce store, 6 months (Thu 2026-01-01 → Tue 2026-06-30,
181 days) · **Status: SAMPLE delivery on DEMO data**
**Repo:** https://github.com/mhkh-stock/p23-marketing-campaign-roi-tracker
**Tools:** Python · Excel (live formulas) · reportlab
**Deliverable:** `outputs/campaign_roi_report.pdf` (2 pages) +
`outputs/campaign_roi_report_SAMPLE.xlsx`

> ⚠️ **What this is and is not.** The campaign file is a fictional
> store's demo data, generated to show the method. No real store,
> campaign or customer is involved. What is real is the analysis:
> the same CTR / CPC / CPA / ROAS / payback breakdowns and workbook
> run on a real client's ad-platform export the same way, and every
> number in the report is recomputed from the data, never typed.

## Business question it answers

> "Which of my campaigns actually earn their money back — and where
> should the next rupee go?"

A store that only watches total spend cannot tell whether its
revenue is carried by one winner, eaten by two losers, or bought
with cheap clicks that never become orders. This delivery answers
it from the campaigns: every campaign's ROAS, cost per order and
payback time, a Kill / Scale / Keep call on each one with the rule
stated, the channel rollup, the monthly trend, and the funnel from
impressions to orders — with the reallocation maths beside it so
the recommendation can be checked, not just believed.

## The method (plainly stated)

**CTR = clicks / impressions. CPC = spend / clicks. CPA = spend /
conversions. Conversion rate = conversions / clicks. ROAS =
revenue / spend. Revenue per click = revenue / clicks.**

**Gross profit = revenue × 45%.** The 45% is the stated
assumption: StyleHub.pk sells fashion at roughly 45% gross margin
after product cost. It is an assumption, not a measured margin —
the workbook shows the unadjusted revenue next to every adjusted
number so the reader can re-run it at their own margin.

**Payback days = spend / (gross profit / campaign days)** — how
many days of gross profit it takes to earn the spend back. A
campaign whose payback is longer than its flight never paid back
while it ran: Brand Awareness — Meta needs 115.27 days on a
45-day flight, Flash Sale SMS needs 56.14 days on a 20-day flight.

**The call rule, stated:** ROAS < 1.00 → Kill (every rupee
returns less than a rupee). ROAS ≥ 3.00 and spend < PKR 900,000
→ Scale (proven, with headroom to grow). ROAS ≥ 3.00 and spend
≥ PKR 900,000 → Keep — proven at scale (already at the ceiling,
hold and optimise). Otherwise → Keep / Watch (profitable, but
below the Scale bar). The rule is a formula in the workbook's
Call column, not a typed label.

**Monthly trend:** a campaign's spend, revenue and conversions
are spread evenly across its flight days, so a campaign running
Mar 1 – Apr 15 puts 31/46 of its spend in March and 15/46 in
April. Monthly totals are the sum of those day-shares — the split
columns on Campaigns are live formulas, not typed numbers.

**Limits, stated upfront:** Email's 15.98x ROAS sits on only
PKR 83,000 of spend and 301 orders — the rate is real, the scale
is not. Winback — Lapsed Buyers shows the best ROAS in the
account (18.17x) on PKR 38,000 of spend; it runs on a small list
and cannot absorb a big budget. For Email and SMS/WhatsApp,
"impressions" means messages sent (emails delivered / SMS
delivered), and CTR means clicks / messages sent — the same
funnel, measured the way those channels are actually bought.
January–February ran at ROAS 1.76–1.84x because the account was
carrying the Brand Awareness flight and the TikTok launch; the
March jump to 3.09x is the Eid campaigns landing, not a change
in the method.

## What is inside

| File | What it does |
|---|---|
| `outputs/campaign_roi_report.pdf` | The deliverable: 2-page client report — headline, ROAS by channel, every campaign with its Kill / Scale / Keep call, monthly trend, funnel, 5 recommendations |
| `outputs/campaign_roi_report_SAMPLE.xlsx` | The working report: Dashboard (KPIs, live formulas) + Campaigns + Channels + Monthly Trend + Funnel + Method sheets |
| `outputs/campaign_roi_report.md` | Findings report: headline numbers, all tables, the reasoning |
| `outputs/summary.json` | Machine-readable summary (headline KPIs + campaigns + channels + months + funnel + actions) |
| `outputs/channel_roas.png` | ROAS by channel — with the Kill line (1.0x) and Scale bar (3.0x) drawn on it |
| `outputs/monthly_trend.png` | Monthly trend — spend vs revenue by month (render of the same data) |
| `data/raw/campaigns_SAMPLE.csv` | The campaign file: 12 campaigns, channel + dates + spend + impressions + clicks + conversions + revenue |
| `data/processed/campaigns_enriched.csv` | All 12 campaigns as read by `src/analyze.py`, with every metric |
| `data/processed/channel_summary.csv` | 7-row channel summary produced by `src/analyze.py` |
| `data/processed/monthly_summary.csv` | 6-month summary produced by `src/analyze.py` |
| `data/processed/funnel_summary.csv` | 3-row funnel summary produced by `src/analyze.py` |
| `data/README.md` | Data dictionary + known quirks |
| `src/generate_data.py` | Makes the demo file (seeded; NOT run on a real client file) |
| `src/analyze.py` | Recomputes every headline number from the raw CSV |
| `src/build_workbook.py` | Builds the workbook: 6 sheets, live formulas |
| `src/build_report.py` | Builds the 2-page PDF + charts from `outputs/summary.json` |
| `src/checks.py` | Data-quality + reconciliation checks as code — 1,000+ checks, all passing |
| `requirements.txt` | Pinned: `openpyxl==3.1.2`, `reportlab==5.0.1`, `matplotlib==3.6.3` (Python 3.12) |

## How it fits together

```
Client ad-platform export (same columns) ──▶ src/analyze.py ──▶ summary.json
        (12 rows in the demo)                      + campaign_roi_report.md
                                                        ▼
                                              src/build_workbook.py ──▶ campaign_roi_report_SAMPLE.xlsx
                                                        │              (6 sheets, live formulas)
                                                        ▼
                                              src/build_report.py ──▶ campaign_roi_report.pdf (2 pages)
                                                        ▲ src/checks.py verifies all of it (1,000+ checks)
```

## Headline numbers (SAMPLE — StyleHub.pk, 6 months)

- **12 campaigns across 7 channels** · **total spend
  PKR 5,693,000** · **total revenue PKR 17,131,500** ·
  **overall ROAS 3.01x** · **gross profit (45%)
  PKR 7,709,175** · **blended payback 133.66 days**
- **8,703,740 impressions · 340,334 clicks · 4,139 orders** ·
  **CTR 3.91% · CPC PKR 16.73 · CPA PKR 1,375.45** ·
  **conversion rate 1.22% · revenue per click PKR 50.34**
- **Calls: 7 Scale · 3 Keep / Watch · 2 Kill** — the 2 Kill
  campaigns burned PKR 740,000 for PKR 637,400 back
- **Best campaign: Winback — Lapsed Buyers (ROAS 18.17x,
  payback 4.52 days)**; worst: **Flash Sale SMS (ROAS 0.79x,
  payback 56.14 days on a 20-day flight)**
- **Best channel by ROAS: Email (15.98x)**; worst:
  **SMS/WhatsApp (0.79x)**; biggest revenue channel:
  **Google Search (PKR 5,153,300 at ROAS 3.51x)**
- **Reallocation: the PKR 740,000 freed from the Kill campaigns,
  at the Scale campaigns' average ROAS of 4.24x, returns about
  PKR 3,137,600 instead of PKR 637,400**

## Every campaign — with its call

| Campaign | Channel | Days | Spend | Revenue | ROAS | Conv. rate | CPA | Payback | Call |
|---|---|---|---|---|---|---|---|---|---|
| Winback — Lapsed Buyers | Email | 37 | PKR 38,000 | PKR 690,600 | 18.17x | 5.02% | PKR 256.76 | 4.52 days | **Scale** |
| Eid Email Blast | Email | 32 | PKR 45,000 | PKR 635,500 | 14.12x | 4.07% | PKR 294.12 | 5.04 days | **Scale** |
| Influencer Fest — Ayesha | Influencer | 47 | PKR 300,000 | PKR 2,028,600 | 6.76x | 3.05% | PKR 583.66 | 15.45 days | **Scale** |
| Eid Lawn Launch | Google Search | 46 | PKR 850,000 | PKR 3,114,500 | 3.66x | 3.23% | PKR 1,315.79 | 27.9 days | **Scale** |
| TikTok Eid Dance | TikTok Ads | 47 | PKR 410,000 | PKR 1,465,600 | 3.57x | 0.9% | PKR 997.57 | 29.22 days | **Scale** |
| New Arrivals — Shopping | Google Shopping | 81 | PKR 540,000 | PKR 1,907,300 | 3.53x | 2.58% | PKR 1,084.34 | 50.96 days | **Scale** |
| Summer Sale — Search | Google Search | 46 | PKR 620,000 | PKR 2,038,800 | 3.29x | 2.97% | PKR 1,302.52 | 31.09 days | **Scale** |
| Eid Collection — Shopping | Google Shopping | 50 | PKR 480,000 | PKR 1,278,500 | 2.66x | 1.96% | PKR 1,538.46 | 41.72 days | **Keep / Watch** |
| Ramadan Festive Push | Meta Ads | 50 | PKR 720,000 | PKR 1,750,500 | 2.43x | 1.2% | PKR 1,818.18 | 45.7 days | **Keep / Watch** |
| TikTok Viral Challenge | TikTok Ads | 59 | PKR 950,000 | PKR 1,584,200 | 1.67x | 0.35% | PKR 2,245.86 | 78.62 days | **Keep / Watch** |
| Brand Awareness — Meta | Meta Ads | 45 | PKR 680,000 | PKR 589,900 | 0.87x | 0.32% | PKR 4,689.66 | 115.27 days | **Kill** |
| Flash Sale SMS | SMS/WhatsApp | 20 | PKR 60,000 | PKR 47,500 | 0.79x | 0.73% | PKR 3,529.41 | 56.14 days | **Kill** |

Winback — Lapsed Buyers has the best ROAS in the account
(18.17x) but only PKR 38,000 of spend behind it — it is a small
list, not a big lever. The biggest proven winner at scale is
Influencer Fest — Ayesha (6.76x on PKR 300,000). The two Kill
campaigns both fail the same test twice: ROAS below 1.00 and a
payback longer than the flight that ran them.

## By channel

| Channel | Campaigns | Spend | Impressions | Clicks | Orders | Revenue | CTR | CPC | CPA | Conv. rate | ROAS | Payback |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Email | 2 | PKR 83,000 | 103,543 | 6,707 | 301 | PKR 1,326,100 | 6.48% | PKR 12.38 | PKR 275.75 | 4.49% | 15.98x | 9.6 days |
| Influencer | 1 | PKR 300,000 | 481,764 | 16,847 | 514 | PKR 2,028,600 | 3.5% | PKR 17.81 | PKR 583.66 | 3.05% | 6.76x | 15.45 days |
| Google Search | 2 | PKR 1,470,000 | 568,452 | 36,014 | 1,122 | PKR 5,153,300 | 6.34% | PKR 40.82 | PKR 1,310.16 | 3.12% | 3.51x | 58.32 days |
| Google Shopping | 2 | PKR 1,020,000 | 751,583 | 35,223 | 810 | PKR 3,185,800 | 4.69% | PKR 28.96 | PKR 1,259.26 | 2.3% | 3.12x | 93.21 days |
| TikTok Ads | 2 | PKR 1,360,000 | 3,072,524 | 164,779 | 834 | PKR 3,049,800 | 5.36% | PKR 8.25 | PKR 1,630.7 | 0.51% | 2.24x | 105.04 days |
| Meta Ads | 2 | PKR 1,400,000 | 3,699,584 | 78,443 | 541 | PKR 2,340,400 | 2.12% | PKR 17.85 | PKR 2,587.8 | 0.69% | 1.67x | 126.28 days |
| SMS/WhatsApp | 1 | PKR 60,000 | 26,290 | 2,321 | 17 | PKR 47,500 | 8.83% | PKR 25.85 | PKR 3,529.41 | 0.73% | 0.79x | 56.14 days |

TikTok Ads brings the most clicks in the account (164,779) at
the lowest conversion rate (0.51%) — cheap traffic that does not
buy. Email brings the orders (4.49% conversion) on the smallest
spend in the account. Google Search is the revenue engine:
PKR 5,153,300, more than any other channel, at ROAS 3.51x.

## Month by month (day-share split)

| Month | Campaigns active | Spend | Revenue | Orders | ROAS |
|---|---|---|---|---|---|
| 2026-01 | 2 | PKR 403,556 | PKR 740,883 | 190 | 1.84x |
| 2026-02 | 4 | PKR 1,190,225 | PKR 2,093,276 | 534 | 1.76x |
| 2026-03 | 6 | PKR 1,850,431 | PKR 5,718,604 | 1,337 | 3.09x |
| 2026-04 | 6 | PKR 1,068,363 | PKR 4,323,764 | 1,041 | 4.05x |
| 2026-05 | 4 | PKR 907,981 | PKR 3,355,998 | 825 | 3.7x |
| 2026-06 | 3 | PKR 272,444 | PKR 898,975 | 212 | 3.3x |

January–February ran at ROAS 1.76–1.84x — the account was
carrying the Brand Awareness flight and the TikTok launch. From
March the Eid campaigns and the Influencer flight land and ROAS
jumps to 3.09x, peaking at 4.05x in April.

## The funnel — impressions to orders

| Stage | Count | Conversion from previous | Share of impressions |
|---|---|---|---|
| Impressions | 8,703,740 | 100.0% | 100.0% |
| Clicks | 340,334 | 3.91% | 3.91% |
| Conversions | 4,139 | 1.22% | 0.05% |

Only 0.05% of impressions become an order. The leak that
matters is click → order: 1.22% of clicks convert. Email
converts at 4.49% on a small list; TikTok converts at 0.51% on
the biggest click volume in the account.

## The five recommendations (from the data, not from a template)

1. **Kill 2 campaigns now: Brand Awareness — Meta (ROAS 0.87,
   payback 115.27 days on a 45-day flight — it never paid back),
   Flash Sale SMS (ROAS 0.79, payback 56.14 days on a 20-day
   flight — it never paid back). Together they burned
   PKR 740,000 for PKR 637,400 back. Stop them before the next
   rupee is spent.**
2. **Scale the winners: Winback — Lapsed Buyers is the best
   campaign in the account by ROAS (18.17x, payback 4.52 days,
   PKR 690,600 on PKR 38,000 spend) — but it runs on a small
   email list, so it cannot absorb a big budget. Put the
   PKR 740,000 freed from the Kill campaigns into the biggest
   proven winner at scale instead, Influencer Fest — Ayesha
   (ROAS 6.76x, PKR 2,028,600 on PKR 300,000 spend). At the Scale
   campaigns' average ROAS of 4.24 that money returns about
   PKR 3,137,600 instead of PKR 637,400.**
3. **Fix TikTok before spending more on it: 164,779 clicks — the
   most traffic in the account — at a 0.51% conversion rate, the
   lowest of any paid channel. The clicks are cheap (PKR 8 each)
   but they do not buy: CPA is PKR 1,631. Send TikTok traffic to
   a dedicated landing page and tighten the audience before the
   next rupee goes in.**
4. **Grow Email — the best ROAS in the account (15.98x) on only
   PKR 83,000 spend. It returned PKR 1,326,100 from 301 orders,
   but the list is too small to move the business. Grow the list
   (checkout opt-in, Eid / sale captures) before raising email
   spend.**
5. **Hold the proven Search / Shopping core and watch the two
   mid-table campaigns: Google Search is the biggest revenue
   channel (PKR 5,153,300 at ROAS 3.51); Eid Collection —
   Shopping and Ramadan Festive Push sit at ROAS 2.43–2.66 —
   profitable, but below the Scale bar. Re-check them after the
   next flight before raising their budgets.**

## Same logic, two ways — and they must agree

The headline KPIs, the campaign metrics, the channel rollups,
the monthly split and the calls exist in two forms: the Python
in `src/analyze.py` and this workbook's formulas. `src/checks.py`
runs both against the raw file and compares them — a mistake in
either one shows up as a failure. The workbook's formulas were
also recalculated in LibreOffice: every total matches the
recompute exactly, with zero error values.

## Reproduce it (2 minutes)

```bash
python3 src/generate_data.py   # rebuild the SAMPLE file (seed 20261119)
python3 src/analyze.py         # summary.json + campaign_roi_report.md
python3 src/build_workbook.py  # outputs/campaign_roi_report_SAMPLE.xlsx
python3 src/build_report.py    # outputs/campaign_roi_report.pdf (2 pages)
python3 src/checks.py          # 1,000+ checks — expect all passing
```

On a real client file, replace `data/raw/campaigns_SAMPLE.csv`
with the ad-platform export (same columns) and skip
`generate_data.py` — everything downstream runs unchanged.
