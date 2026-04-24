# quarter-corro

Quarterly Sales Report pipeline for **Equestrian Labs / Corro**.

Pulls raw data from the **Shopify API** → aggregates → writes to **Google Sheets** → renders a standalone **HTML dashboard**.

---

## What this generates

| Sheet | Description |
|-------|-------------|
| `SUMMARY` | KPI snapshot: Total/Net Sales, GP, Margin, Units, Orders |
| `top100_all` | Top 100 products sorted by Total Sales |
| `top100_gross_profit` | Top 100 sorted by **Gross Profit** (Ceci v2 requirement) |
| `top100_dropship` | Top 100 Dropship products only |
| `top100_best_margin` | Top 100 by GP% (min 5 units sold) |
| `monthly_breakdown` | Per-product split: Jan / Feb / Mar |
| `cost_zero_items` | COGS = $0 but customer paid — needs Shopify fix |
| `zero_sales` | SKUs with $0 revenue this quarter |
| `product_types` | All raw Shopify Product Types → 7 main categories |
| `by_category` | Revenue aggregated by main category |
| `by_channel` | Revenue by channel (Online, Concierge, Wellington) |

---

## Product Type Mapping

Shopify's `product_type` field uses a hierarchical format (`Parent > Sub > Sub2`).
This pipeline takes the **first segment** and maps to one of 7 parent categories:

| Shopify Product Type (first segment) | Maps To |
|--------------------------------------|---------|
| Horse Wear | **Horse Wear** |
| Rider | **Rider** |
| Horse Care | **Horse Care** |
| Stable | **Stable** |
| Tack & Equipment | **Tack & Equipment** |
| Pharmacy | **Pharmacy** |
| Dog | **Dog** |
| Everything else | **Other** |

> **⚠️ Note (Shopify AI):** equestrian-labs has ~20 unique product types but with duplication issues. For example, "English Saddle Pads > Dressage Saddle Pads" appears both as `Tack & Equipment > English Saddle Pads > Dressage Saddle Pads` and as just `English Saddle Pads > Dressage Saddle Pads` (missing parent). Also 156 products have **no type assigned** — these fall into `Other`. Recommend cleaning up in Shopify Admin.

---

## Usage

```bash
# Install dependencies
pip install requests gspread google-auth pytz

# Set env vars
export SHOPIFY_TOKEN_CORRO="shpat_xxxx"
export SHOPIFY_STORE="equestrian-labs.myshopify.com"
export GOOGLE_CREDENTIALS='{"type":"service_account",...}'
export SHEET_ID_QUARTER="1rDF18OHvX9X-M5kdMR51MKrQGoG9UgB17YzXtp7L7qQ"

# Run for Q1 2026
python scripts/quarter_pipeline.py --quarter 1 --year 2026

# Run for Q2 2026
python scripts/quarter_pipeline.py --quarter 2 --year 2026
```

---

## GitHub Actions

Add these secrets to the repo:
- `SHOPIFY_TOKEN_CORRO`
- `SHOPIFY_STORE`
- `GOOGLE_CREDENTIALS` (service account JSON)
- `SHEET_ID_QUARTER`

The workflow `.github/workflows/quarterly.yml` runs automatically or on manual trigger.

---

## Dashboard

`dashboard/q1_2026_dashboard.html` — Standalone HTML dashboard with:
- All 6 product views (All, Gross Profit, Dropship, Best Margin, Cost Zero, Zero Sales)
- Category breakdown using **Product Type** (not collections)
- Channel breakdown (Online / Concierge / Wellington)
- Product Type audit table (raw Shopify types → mapped category)
- Sidebar navigation

When the pipeline writes to Google Sheets, the dashboard can be connected to read live data via the Sheets API (see `SHEET_ID` constant in the HTML).

---

## Key Design Decisions

1. **Sorted by Gross Profit** — Per Ceci's v2 requirement, products are ranked by GP (not Net Sales). Pareto zones (A/B/C) are assigned by cumulative GP (Zone A = top products generating 80% of GP).
2. **Deduplication fix** — Orders are fetched in two batches (paid + refunded) and deduplicated by order ID. This fixes the v1.0 bug that returned 18K instead of ~8.8K orders.
3. **Product Type categories** — Based on Shopify's native `product_type` field (not collections). First segment of the hierarchy is used for grouping.
4. **COGS = 0 flagging** — Products where customers paid but COGS is $0 are isolated in `cost_zero_items` sheet and flagged in the dashboard. This means GP is overstated for those products.
5. **Concierge detection** — Orders tagged `concierge` (case-insensitive) in Shopify tags OR source_name are grouped as Concierge channel.
