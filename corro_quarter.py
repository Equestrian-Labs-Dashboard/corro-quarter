"""
CORRO QUARTERLY REPORT — Q1 2026
=================================
Genera el reporte trimestral de Corro desde Shopify API.
Estructura igual al Excel Q1_2026: SKU-level con breakdown mensual.

Hojas que genera en Google Sheets:
  q_summary        — KPIs totales + monthly breakdown (como SUMMARY del Excel)
  q_top_products   — Top 100 products por Gross Profit (con Jan/Feb/Mar)
  q_zero_cost      — Productos con costo $0 para el cliente (descuento 100%)
                     categorizados por tipo: Team Rider, Influencer, Marketing,
                     Advent Calendar, Sponsorship, Other
  q_discounts      — Códigos de descuento usados en el período

Run:
  python corro_quarter.py --quarter q1 --year 2026
  python corro_quarter.py --quarter q2 --year 2026

Secrets requeridos (GitHub Actions o .env):
  SHOPIFY_TOKEN_CORRO   GOOGLE_CREDENTIALS   SHEET_ID_QUARTER
"""

import os, json, requests, gspread, argparse, calendar
from google.oauth2.service_account import Credentials
from datetime import datetime, date
from collections import defaultdict
import pytz

TIMEZONE    = pytz.timezone("America/Bogota")
API_VERSION = "2024-10"
STORE_URL   = os.environ.get("SHOPIFY_STORE", "equestrian-labs.myshopify.com")
TOKEN       = os.environ.get("SHOPIFY_TOKEN_CORRO", "")
SHEET_ID    = os.environ.get("SHEET_ID_QUARTER",
              "1NnH7Ln3HP9AuJ5ohxgvVk6A5BnG9_iz9WPC9SxaaidI")
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets",
               "https://www.googleapis.com/auth/drive"]

import time

# ── QUARTER PERIODS ───────────────────────────────────────────────
QUARTER_MONTHS = {
    "q1": [1, 2, 3],
    "q2": [4, 5, 6],
    "q3": [7, 8, 9],
    "q4": [10, 11, 12],
}
MONTH_NAMES = {
    1:"January",2:"February",3:"March",4:"April",
    5:"May",6:"June",7:"July",8:"August",
    9:"September",10:"October",11:"November",12:"December"
}

def quarter_range(q, y):
    months = QUARTER_MONTHS[q]
    start  = date(y, months[0], 1)
    end    = date(y, months[-1], calendar.monthrange(y, months[-1])[1])
    today  = datetime.now(TIMEZONE).date()
    if end > today:
        end = today
    return start, end, months

# ── ZERO COST TAG CATEGORIES ──────────────────────────────────────
# Ceci quiere agrupar los productos con costo $0 para el cliente
# (ya sea precio=0 o descuento 100%) por tipo de iniciativa.
# Se detecta por order tags.
ZERO_COST_CATEGORIES = [
    ("Advent Calendar",  ["advent calendar", "advent_calendar"]),
    ("Team Rider",       ["team rider", "team_rider"]),
    ("Influencer",       ["influencer"]),
    ("Marketing",        ["marketing", "marketing - sponsorship"]),
    ("Sponsorship",      ["sponsorship", "sponsor"]),
    ("Internal / Staff", ["staff", "internal", "employee"]),
]

def classify_zero_cost(order_tags_str):
    tags_lower = (order_tags_str or "").lower()
    for category, keywords in ZERO_COST_CATEGORIES:
        if any(k in tags_lower for k in keywords):
            return category
    return "Other"

# ── SHOPIFY REST ──────────────────────────────────────────────────
def shopify_get(endpoint, params):
    url     = f"https://{STORE_URL}/admin/api/{API_VERSION}/{endpoint}"
    headers = {"X-Shopify-Access-Token": TOKEN}
    results = []
    while url:
        for attempt in range(8):
            try:
                r = requests.get(url, headers=headers, params=params, timeout=60)
            except requests.exceptions.ConnectionError as e:
                wait = min(2**attempt, 60)
                print(f"    connection error — retrying in {wait}s: {e}")
                time.sleep(wait); continue
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", 2**attempt))
                print(f"    rate-limited — waiting {wait}s...")
                time.sleep(wait); continue
            if r.status_code in (502, 503, 504):
                time.sleep(min(2**attempt, 60)); continue
            r.raise_for_status()
            break
        data = r.json()
        key  = [k for k in data if k != "errors"][0]
        results.extend(data[key])
        link = r.headers.get("Link", ""); url = None; params = {}
        if 'rel="next"' in link:
            for part in link.split(","):
                if 'rel="next"' in part:
                    url = part.split(";")[0].strip().strip("<>")
        time.sleep(0.3)
    return results

def shopify_graphql(query, variables=None):
    url     = f"https://{STORE_URL}/admin/api/2026-01/graphql.json"
    headers = {"X-Shopify-Access-Token": TOKEN, "Content-Type": "application/json"}
    payload = {"query": query}
    if variables:
        payload["variables"] = variables
    for attempt in range(8):
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=60)
        except requests.exceptions.ConnectionError as e:
            time.sleep(min(2**attempt, 60)); continue
        if r.status_code == 429:
            time.sleep(int(r.headers.get("Retry-After", 2**attempt))); continue
        if r.status_code in (502, 503, 504):
            time.sleep(min(2**attempt, 60)); continue
        r.raise_for_status()
        d = r.json()
        if any((e.get("extensions") or {}).get("code") == "THROTTLED"
               for e in (d.get("errors") or [])):
            time.sleep(min(2**attempt, 60)); continue
        return d
    raise RuntimeError("GraphQL failed after 8 attempts")

# ── FETCH ORDERS ──────────────────────────────────────────────────
def fetch_orders(start, end):
    print(f"  Fetching orders {start} → {end}...")
    seen_ids, all_orders = set(), []
    for status in ["paid,partially_paid", "partially_refunded,refunded"]:
        batch = shopify_get("orders.json", {
            "status": "any",
            "financial_status": status,
            "created_at_min": f"{start}T00:00:00-05:00",
            "created_at_max": f"{end}T23:59:59-05:00",
            "limit": 250,
            "fields": "id,name,created_at,subtotal_price,total_price,"
                      "total_discounts,discount_codes,source_name,tags,"
                      "line_items,customer",
        })
        for o in batch:
            oid = o.get("id")
            if oid and oid not in seen_ids:
                seen_ids.add(oid)
                all_orders.append(o)
    print(f"  → {len(all_orders)} unique orders")
    return all_orders

# ── FETCH PRODUCT MAP (GraphQL — product_type, vendor, COGS) ──────
def fetch_product_map():
    print("  Fetching product catalogue + COGS via GraphQL...")
    QUERY = """
    query($cursor: String) {
      productVariants(first: 100, after: $cursor) {
        pageInfo { hasNextPage endCursor }
        edges { node {
          id sku
          product { id title vendor productType }
          inventoryItem { id unitCost { amount } }
        }}
      }
    }
    """
    vmap = {}; title_map = {}; cursor = None
    while True:
        data  = shopify_graphql(QUERY, {"cursor": cursor})
        pv    = (data.get("data") or {}).get("productVariants", {})
        for edge in pv.get("edges", []):
            node = edge["node"]
            vid  = node["id"].split("/")[-1]
            prod = node.get("product") or {}
            inv  = node.get("inventoryItem") or {}
            uc   = inv.get("unitCost") or {}
            sku  = node.get("sku") or ""
            title  = prod.get("title", "")
            ptype  = prod.get("productType") or "Uncategorized"
            vendor = prod.get("vendor", "")
            cost   = float(uc.get("amount") or 0)
            pid    = prod.get("id","").split("/")[-1]
            vmap[vid] = {
                "product_id":   pid,
                "product_title":title,
                "sku":          sku,
                "vendor":       vendor,
                "product_type": ptype,
                "unit_cost":    cost,
            }
            if title:
                title_map[title.lower()] = {"product_id":pid,"product_type":ptype,"vendor":vendor}
        pi = pv.get("pageInfo", {})
        if not pi.get("hasNextPage"):
            break
        cursor = pi["endCursor"]
        time.sleep(0.3)
    filled = sum(1 for v in vmap.values() if v["unit_cost"] > 0)
    print(f"  → {len(vmap)} variants | {filled} with COGS > 0")
    return vmap, title_map

# ── SHOPIFYQL SALES ───────────────────────────────────────────────
def fetch_shopifyql_sales(start, end):
    """ShopifyQL: gross_profit, gross_sales, net_sales, orders, units per product+vendor."""
    print(f"  Fetching ShopifyQL sales {start} → {end}...")

    GQL = """
    query shopifyqlSales($q: String!) {
      shopifyqlQuery(query: $q) {
        tableData { columns { name dataType } rows }
        parseErrors
      }
    }
    """

    def run_q(q):
        d  = shopify_graphql(GQL, {"q": q})
        sq = (d.get("data") or {}).get("shopifyqlQuery")
        if not sq:
            errs = d.get("errors") or []
            print(f"  ✗ shopifyqlQuery null: {errs[:1]}")
            return [], [], True
        raw_err = sq.get("parseErrors")
        if raw_err:
            print(f"  ⚠ parseErrors: {raw_err}")
            return [], [], True
        table = sq.get("tableData") or {}
        cols  = [c["name"] for c in (table.get("columns") or [])]
        rows  = table.get("rows") or []
        print(f"    → cols={cols} rows={len(rows)}")
        return cols, rows, False

    q_main = (
        f"FROM sales "
        f"SHOW gross_profit, gross_sales, net_sales, orders, net_items_sold "
        f"GROUP BY product_title, product_vendor "
        f"SINCE {start} UNTIL {end} "
        f"ORDER BY gross_profit DESC"
    )
    cols, rows, err = run_q(q_main)
    if err:
        q_fallback = (
            f"FROM sales "
            f"SHOW gross_profit, gross_sales, net_sales, orders, net_items_sold "
            f"GROUP BY product_title "
            f"SINCE {start} UNTIL {end} "
            f"ORDER BY gross_profit DESC"
        )
        cols, rows, err2 = run_q(q_fallback)
        if err2:
            print("  ✗ ShopifyQL unavailable — GP will be estimated from COGS")
            return []

    results = []; seen = set()
    for row in rows:
        if not isinstance(row, dict):
            row = dict(zip(cols, row)) if isinstance(row, list) else {}
        title = (row.get("product_title") or "").strip()
        if not title:
            continue
        if title in seen:
            for r in results:
                if r["product_title"] == title:
                    r["gross_profit"] += float(row.get("gross_profit") or 0)
                    r["gross_sales"]  += float(row.get("gross_sales") or 0)
                    r["net_sales"]    += float(row.get("net_sales") or 0)
                    r["orders"]       += int(row.get("orders") or 0)
                    r["units"]        += int(row.get("net_items_sold") or 0)
                    break
            continue
        seen.add(title)
        ns = float(row.get("net_sales") or 0)
        gp = float(row.get("gross_profit") or 0)
        gs = float(row.get("gross_sales") or 0) or ns
        results.append({
            "product_title": title,
            "vendor":        (row.get("product_vendor") or "").strip(),
            "gross_profit":  gp,
            "gross_sales":   gs,
            "discounts":     max(gs - ns, 0),
            "net_sales":     ns,
            "orders":        int(row.get("orders") or 0),
            "units":         int(row.get("net_items_sold") or 0),
        })
    print(f"  → {len(results)} products from ShopifyQL")
    return results

# ── AGGREGATE BY SKU (monthly breakdown) ─────────────────────────
def aggregate_by_sku(orders, vmap, months):
    """
    Aggregates orders at SKU (variant) level with monthly breakdown.
    Returns dict: variant_id → {product_title, sku, vendor, product_type,
                                 unit_cost, gross_sales, discounts, returns,
                                 net_sales, cogs, gross_profit, units, orders,
                                 monthly: {month_int: {gross_sales, net_sales, units}}}
    """
    agg = {}
    for order in orders:
        order_month = int(order.get("created_at", "")[:7].split("-")[1] or 0)
        disc_total  = float(order.get("total_discounts", 0) or 0)
        li_count    = len(order.get("line_items", [])) or 1
        disc_per_li = disc_total / li_count

        for li in order.get("line_items", []):
            vid  = str(li.get("variant_id") or "")
            info = vmap.get(vid, {
                "product_id": str(li.get("product_id", "")),
                "product_title": li.get("title", "Unknown"),
                "sku": li.get("sku", ""),
                "vendor": "",
                "product_type": "Uncategorized",
                "unit_cost": 0.0,
            })
            qty   = int(li.get("quantity", 0) or 0)
            price = float(li.get("price", 0) or 0) * qty
            disc  = disc_per_li
            net   = max(price - disc, 0)
            cost  = info["unit_cost"] * qty

            # Track refunds via financial_status — approximate
            is_refund = (order.get("financial_status") or "").startswith("refund")
            ret  = -net if is_refund else 0.0

            key = vid or f"nv_{li.get('product_id','x')}_{li.get('title','')}"
            if key not in agg:
                agg[key] = {
                    "variant_id":    vid,
                    "product_id":    info["product_id"],
                    "product_title": info["product_title"],
                    "sku":           info["sku"],
                    "vendor":        info["vendor"],
                    "product_type":  info["product_type"],
                    "unit_cost":     info["unit_cost"],
                    "gross_sales":   0.0, "discounts": 0.0,
                    "returns":       0.0, "net_sales":  0.0,
                    "cogs":          0.0, "gross_profit": 0.0,
                    "units":         0,   "orders":       set(),
                    "monthly":       {m: {"gross_sales":0.0,"net_sales":0.0,"units":0}
                                      for m in months},
                    "is_dropship":   "YES" if info["unit_cost"] > 0 else "NO",
                }
            r = agg[key]
            r["gross_sales"]  += price
            r["discounts"]    += disc
            r["returns"]      += ret
            r["net_sales"]    += net
            r["cogs"]         += cost
            r["gross_profit"] += max(net - cost, 0)
            r["units"]        += qty
            r["orders"].add(order["id"])
            if order_month in months:
                r["monthly"][order_month]["gross_sales"] += price
                r["monthly"][order_month]["net_sales"]   += net
                r["monthly"][order_month]["units"]       += qty
    return agg

# ── AGGREGATE ZERO COST ORDERS ────────────────────────────────────
def aggregate_zero_cost(orders, vmap, months):
    """
    Finds line items where the customer paid $0 (price=0 OR net after discounts ~0).
    Groups by order tag category (Team Rider, Influencer, Marketing, etc.)
    and by month with units and COGS.
    """
    # agg[category][month] = {units, cogs, product_titles: set}
    agg = defaultdict(lambda: {
        m: {"units": 0, "cogs": 0.0, "products": set()} for m in months
    })
    agg_detail = []  # list of individual zero-cost line items for detail sheet

    for order in orders:
        order_month = int(order.get("created_at", "")[:7].split("-")[1] or 0)
        if order_month not in months:
            continue
        tags_str   = order.get("tags") or ""
        category   = classify_zero_cost(tags_str)
        disc_total = float(order.get("total_discounts", 0) or 0)
        li_count   = len(order.get("line_items", [])) or 1
        disc_per_li = disc_total / li_count

        for li in order.get("line_items", []):
            price = float(li.get("price", 0) or 0)
            qty   = int(li.get("quantity", 0) or 0)
            net   = max(price * qty - disc_per_li, 0)
            # Zero cost = customer paid $0 (price=0 OR net≈0 after 100% discount)
            if price == 0 or net < 0.01:
                vid  = str(li.get("variant_id") or "")
                info = vmap.get(vid, {})
                unit_cost = info.get("unit_cost", 0.0)
                cogs_line  = unit_cost * qty
                title = info.get("product_title") or li.get("title", "Unknown")

                agg[category][order_month]["units"]   += qty
                agg[category][order_month]["cogs"]    += cogs_line
                agg[category][order_month]["products"].add(title)

                agg_detail.append({
                    "order_name":    order.get("name", ""),
                    "created_at":    order.get("created_at", "")[:10],
                    "month":         MONTH_NAMES.get(order_month, ""),
                    "category":      category,
                    "tags":          tags_str,
                    "product_title": title,
                    "sku":           info.get("sku") or li.get("sku", ""),
                    "vendor":        info.get("vendor", ""),
                    "product_type":  info.get("product_type", ""),
                    "units":         qty,
                    "original_price": price * qty,
                    "discount":      disc_per_li,
                    "net_paid":      net,
                    "unit_cost":     unit_cost,
                    "cogs":          cogs_line,
                })

    return agg, agg_detail

# ── AGGREGATE DISCOUNTS ───────────────────────────────────────────
def aggregate_discounts(orders):
    agg = defaultdict(lambda: {
        "discount_code": "", "discount_type": "",
        "total_discounts": 0.0, "net_sales": 0.0, "orders": 0,
    })
    for order in orders:
        disc_total = float(order.get("total_discounts", 0) or 0)
        if disc_total <= 0:
            continue
        sub   = float(order.get("subtotal_price", 0) or 0)
        codes = order.get("discount_codes") or []
        if codes:
            for dc in codes:
                code = (dc.get("code") or "").strip() or "(no code)"
                agg[code]["discount_code"]    = code
                agg[code]["discount_type"]    = dc.get("type", "")
                agg[code]["total_discounts"] += float(dc.get("amount") or 0)
                agg[code]["net_sales"]       += sub
                agg[code]["orders"]          += 1
        else:
            agg["(automatic)"]["discount_code"]    = "(automatic)"
            agg["(automatic)"]["discount_type"]    = "Automatic"
            agg["(automatic)"]["total_discounts"] += disc_total
            agg["(automatic)"]["net_sales"]       += sub
            agg["(automatic)"]["orders"]          += 1
    return agg

# ── SHEETS ────────────────────────────────────────────────────────
def get_gc():
    creds = Credentials.from_service_account_info(
        json.loads(os.environ["GOOGLE_CREDENTIALS"]), scopes=SCOPES)
    return gspread.authorize(creds)

def sheets_call(fn, *args, **kwargs):
    for attempt in range(8):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = e.response.status_code if hasattr(e, "response") else 0
            if status == 429 or status >= 500:
                wait = min(15*(attempt+1), 90)
                print(f"    Sheets {status} — waiting {wait}s...")
                time.sleep(wait); continue
            raise
        except Exception as e:
            if attempt < 3:
                print(f"    Sheets error: {e}"); time.sleep(10); continue
            raise

def write_tab(sh, tab_name, headers, rows):
    """Write rows to a tab, replacing all content."""
    try:    ws = sh.worksheet(tab_name)
    except: ws = sheets_call(sh.add_worksheet, tab_name,
                             rows=max(5000, len(rows)+100), cols=len(headers)+2)
    time.sleep(2)

    all_data = [headers]
    for r in rows:
        clean = []
        for v in r:
            if v is None: clean.append("")
            elif isinstance(v, float) and v != v: clean.append("")
            elif isinstance(v, set): clean.append(", ".join(sorted(v)[:5]))
            else: clean.append(str(v) if not isinstance(v,(int,float,bool)) else v)
        all_data.append(clean)

    total = len(all_data)
    if total > ws.row_count or len(headers) > ws.col_count:
        sheets_call(ws.resize, rows=total+50, cols=len(headers)+2)
        time.sleep(1)

    sheets_call(ws.clear); time.sleep(1)
    BATCH = 500
    for i in range(0, total, BATCH):
        sheets_call(ws.append_rows, all_data[i:i+BATCH],
                    value_input_option="RAW", insert_data_option="INSERT_ROWS")
        print(f"    {tab_name}: {min(i+BATCH,total)}/{total} rows...")
        if i+BATCH < total: time.sleep(2)
    time.sleep(3)
    print(f"    ✓ {tab_name}: {len(rows)} data rows written")

# ── MAIN ──────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--quarter", default="q1",
                        choices=["q1","q2","q3","q4"])
    parser.add_argument("--year", type=int, default=2026)
    args = parser.parse_args()

    q, y      = args.quarter, args.year
    start, end, months = quarter_range(q, y)
    label     = f"{q.upper()} {y}"
    month_labels = [MONTH_NAMES[m] for m in months]
    now_str   = datetime.now(TIMEZONE).strftime("%Y-%m-%d %H:%M")

    print(f"\n{'='*60}")
    print(f"  CORRO QUARTERLY REPORT — {label}")
    print(f"  {start} → {end}")
    print(f"  Months: {', '.join(month_labels)}")
    print(f"  Sheet : {SHEET_ID}")
    print(f"{'='*60}\n")

    # ── Fetch data ────────────────────────────────────────────────
    vmap, title_map = fetch_product_map()
    orders          = fetch_orders(start, end)
    sq_rows         = fetch_shopifyql_sales(start, end)
    sq_map          = {r["product_title"].lower(): r for r in sq_rows}

    # ── Aggregate ─────────────────────────────────────────────────
    sku_agg   = aggregate_by_sku(orders, vmap, months)
    zc_agg, zc_detail = aggregate_zero_cost(orders, vmap, months)
    disc_agg  = aggregate_discounts(orders)

    sh = get_gc().open_by_key(SHEET_ID)

    # ── 1. SUMMARY TAB ───────────────────────────────────────────
    print("\n  Building summary...")
    total_gross_sales = sum(v["gross_sales"] for v in sku_agg.values())
    total_discounts   = sum(v["discounts"]   for v in sku_agg.values())
    total_returns     = sum(v["returns"]     for v in sku_agg.values())
    total_net_sales   = sum(v["net_sales"]   for v in sku_agg.values())
    total_cogs        = sum(v["cogs"]        for v in sku_agg.values())
    total_gp          = sum(v["gross_profit"]for v in sku_agg.values())
    total_units       = sum(v["units"]       for v in sku_agg.values())
    total_orders      = len(set(o["id"] for o in orders))
    gp_margin         = round(total_gp/total_net_sales*100,1) if total_net_sales else 0

    # Monthly totals
    monthly_totals = {m: {"gross_sales":0.0,"discounts":0.0,"net_sales":0.0,
                           "cogs":0.0,"gross_profit":0.0,"units":0}
                      for m in months}
    for row in sku_agg.values():
        for m in months:
            monthly_totals[m]["gross_sales"]  += row["monthly"][m]["gross_sales"]
            monthly_totals[m]["net_sales"]    += row["monthly"][m]["net_sales"]
            monthly_totals[m]["units"]        += row["monthly"][m]["units"]

    h_sum = ["updated_at","quarter","period_start","period_end","metric",
             month_labels[0], month_labels[1], month_labels[2], f"{label} Total"]
    sum_rows = [
        [now_str,label,str(start),str(end),"Gross Sales",
         round(monthly_totals[months[0]]["gross_sales"],2),
         round(monthly_totals[months[1]]["gross_sales"],2),
         round(monthly_totals[months[2]]["gross_sales"],2),
         round(total_gross_sales,2)],
        [now_str,label,str(start),str(end),"Discounts",
         "","","",round(total_discounts,2)],
        [now_str,label,str(start),str(end),"Returns",
         "","","",round(total_returns,2)],
        [now_str,label,str(start),str(end),"Net Sales",
         round(monthly_totals[months[0]]["net_sales"],2),
         round(monthly_totals[months[1]]["net_sales"],2),
         round(monthly_totals[months[2]]["net_sales"],2),
         round(total_net_sales,2)],
        [now_str,label,str(start),str(end),"COGS",
         "","","",round(total_cogs,2)],
        [now_str,label,str(start),str(end),"Gross Profit",
         "","","",round(total_gp,2)],
        [now_str,label,str(start),str(end),"Gross Margin %",
         "","","",f"{gp_margin}%"],
        [now_str,label,str(start),str(end),"Units Sold",
         monthly_totals[months[0]]["units"],
         monthly_totals[months[1]]["units"],
         monthly_totals[months[2]]["units"],
         total_units],
        [now_str,label,str(start),str(end),"Total Orders",
         "","","",total_orders],
    ]
    write_tab(sh, f"q_summary_{q}_{y}", h_sum, sum_rows)

    # ── 2. TOP PRODUCTS TAB ──────────────────────────────────────
    print("\n  Building top products...")
    # Merge SKU agg with ShopifyQL GP (more accurate)
    prod_agg = defaultdict(lambda: {
        "product_title":"","vendor":"","product_type":"","is_dropship":"NO",
        "units":0,"gross_sales":0.0,"discounts":0.0,"returns":0.0,
        "net_sales":0.0,"cogs":0.0,"gross_profit":0.0,
        "orders":set(),
        "monthly":{m:{"gross_sales":0.0,"net_sales":0.0,"units":0} for m in months},
    })
    for row in sku_agg.values():
        t = row["product_title"]
        p = prod_agg[t]
        p["product_title"] = t
        p["vendor"]        = p["vendor"] or row["vendor"]
        p["product_type"]  = p["product_type"] or row["product_type"]
        p["is_dropship"]   = row["is_dropship"]
        p["units"]        += row["units"]
        p["gross_sales"]  += row["gross_sales"]
        p["discounts"]    += row["discounts"]
        p["returns"]      += row["returns"]
        p["net_sales"]    += row["net_sales"]
        p["cogs"]         += row["cogs"]
        p["gross_profit"] += row["gross_profit"]
        p["orders"]       |= row["orders"]
        for m in months:
            p["monthly"][m]["gross_sales"] += row["monthly"][m]["gross_sales"]
            p["monthly"][m]["net_sales"]   += row["monthly"][m]["net_sales"]
            p["monthly"][m]["units"]       += row["monthly"][m]["units"]

    # Override GP with ShopifyQL if available (more accurate)
    for title, sq in sq_map.items():
        key = next((k for k in prod_agg if k.lower() == title), None)
        if key:
            prod_agg[key]["gross_profit"] = sq["gross_profit"]
            prod_agg[key]["gross_sales"]  = sq["gross_sales"]
            prod_agg[key]["net_sales"]    = sq["net_sales"]

    sorted_prods = sorted(prod_agg.values(),
                          key=lambda x: x["gross_profit"], reverse=True)
    total_gp_prods = sum(p["gross_profit"] for p in sorted_prods) or 1
    total_ns_prods = sum(p["net_sales"]    for p in sorted_prods) or 1

    h_prod = ["updated_at","quarter","rank","product_title","vendor","product_type",
              "is_dropship","units","gross_sales","discounts","returns","net_sales",
              "cogs","gross_profit","gross_margin",
              f"{month_labels[0]} sales",f"{month_labels[1]} sales",f"{month_labels[2]} sales",
              f"{month_labels[0]} units",f"{month_labels[1]} units",f"{month_labels[2]} units",
              "pct_gross_profit","pct_net_sales","orders"]
    prod_rows = []
    for i, p in enumerate(sorted_prods):
        ns  = p["net_sales"] or 0
        gp  = p["gross_profit"] or 0
        margin = round(gp/ns*100,1) if ns else 0
        prod_rows.append([
            now_str, label, i+1,
            p["product_title"], p["vendor"], p["product_type"], p["is_dropship"],
            p["units"],
            round(p["gross_sales"],2), round(p["discounts"],2),
            round(p["returns"],2),     round(ns,2),
            round(p["cogs"],2),        round(gp,2),
            f"{margin}%",
            round(p["monthly"][months[0]]["net_sales"],2),
            round(p["monthly"][months[1]]["net_sales"],2),
            round(p["monthly"][months[2]]["net_sales"],2),
            p["monthly"][months[0]]["units"],
            p["monthly"][months[1]]["units"],
            p["monthly"][months[2]]["units"],
            f"{round(gp/total_gp_prods*100,2)}%",
            f"{round(ns/total_ns_prods*100,2)}%",
            len(p["orders"]),
        ])
    write_tab(sh, f"q_products_{q}_{y}", h_prod, prod_rows)

    # ── 3. ZERO COST TAB ─────────────────────────────────────────
    print("\n  Building zero cost sheet...")
    # Summary: category × month
    all_categories = [c for c,_ in ZERO_COST_CATEGORIES] + ["Other"]
    h_zc = ["updated_at","quarter","category",
             f"{month_labels[0]} units",f"{month_labels[0]} cogs",
             f"{month_labels[1]} units",f"{month_labels[1]} cogs",
             f"{month_labels[2]} units",f"{month_labels[2]} cogs",
             "total units","total cogs","products sample"]
    zc_rows = []
    for cat in all_categories:
        if cat not in zc_agg:
            continue
        m_data = zc_agg[cat]
        all_prods = set()
        for m in months:
            all_prods |= m_data[m]["products"]
        total_units_cat = sum(m_data[m]["units"] for m in months)
        total_cogs_cat  = sum(m_data[m]["cogs"]  for m in months)
        zc_rows.append([
            now_str, label, cat,
            m_data[months[0]]["units"], round(m_data[months[0]]["cogs"],2),
            m_data[months[1]]["units"], round(m_data[months[1]]["cogs"],2),
            m_data[months[2]]["units"], round(m_data[months[2]]["cogs"],2),
            total_units_cat, round(total_cogs_cat,2),
            ", ".join(sorted(all_prods)[:5]),
        ])
    write_tab(sh, f"q_zero_cost_{q}_{y}", h_zc, zc_rows)

    # Detail rows
    h_zc_det = ["updated_at","quarter","order_name","created_at","month",
                "category","tags","product_title","sku","vendor","product_type",
                "units","original_price","discount","net_paid","unit_cost","cogs"]
    zc_det_rows = [[now_str, label] + [d[k] for k in [
        "order_name","created_at","month","category","tags",
        "product_title","sku","vendor","product_type",
        "units","original_price","discount","net_paid","unit_cost","cogs"
    ]] for d in sorted(zc_detail, key=lambda x: x["created_at"])]
    write_tab(sh, f"q_zero_cost_detail_{q}_{y}", h_zc_det, zc_det_rows)

    # ── 4. DISCOUNTS TAB ─────────────────────────────────────────
    print("\n  Building discounts sheet...")
    sorted_disc = sorted(disc_agg.values(),
                         key=lambda x: -x["total_discounts"])
    total_disc  = sum(d["total_discounts"] for d in sorted_disc) or 1
    h_disc = ["updated_at","quarter","rank","discount_code","discount_type",
              "total_discounts","pct_of_total","cumulative_pct",
              "orders","avg_discount_per_order","net_sales"]
    disc_rows = []
    cum_d = 0.0
    for i, d in enumerate(sorted_disc):
        cum_d += d["total_discounts"]
        orders_n = d["orders"] or 1
        disc_rows.append([
            now_str, label, i+1,
            d["discount_code"], d["discount_type"],
            round(d["total_discounts"],2),
            f"{round(d['total_discounts']/total_disc*100,2)}%",
            f"{round(cum_d/total_disc*100,2)}%",
            d["orders"],
            round(d["total_discounts"]/orders_n,2),
            round(d["net_sales"],2),
        ])
    write_tab(sh, f"q_discounts_{q}_{y}", h_disc, disc_rows)

    print(f"\n{'='*60}")
    print(f"  ✓ DONE — {label}")
    print(f"  Orders: {total_orders} | Net Sales: ${total_net_sales:,.2f}")
    print(f"  Gross Profit: ${total_gp:,.2f} ({gp_margin}% margin)")
    print(f"  Sheets written:")
    print(f"    q_summary_{q}_{y}")
    print(f"    q_products_{q}_{y}  ({len(prod_rows)} products)")
    print(f"    q_zero_cost_{q}_{y}  ({len(zc_rows)} categories)")
    print(f"    q_zero_cost_detail_{q}_{y}  ({len(zc_det_rows)} line items)")
    print(f"    q_discounts_{q}_{y}  ({len(disc_rows)} codes)")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()
