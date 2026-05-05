"""
CORRO QUARTERLY REPORT — v2.1
==============================
Generates Corro's quarterly report from the Shopify API.
This script is ONLY for the Quarter Report / Quarterly Dashboard.

Sheets generated (each tab receives suffix _{q}_{y}):
  q_summary              — total KPIs + monthly breakdown
  q_top_sku              — Top 100 by individual SKU/variant
  q_top_products         — Top 100 grouped by product
  q_top_dropship         — Top 100 dropship products
  q_top_margin           — Top 100 by Gross Margin % (min 5 units)
  q_top_gp               — Top 100 by Gross Profit $
  q_cost_zero            — products sold with COGS = 0 in Shopify
  q_discount_zero        — 100% discounted / free items by category
  q_discount_zero_detail — order-by-order 100% discount detail
  q_staff_orders         — staff/internal orders flagged separately
  q_unknown_vendors      — sold products with missing/Unknown vendor
  q_monthly_breakdown    — monthly product breakdown
  q_zero_sales           — products with no activity in the period
  q_vendors              — vendor Pareto by Gross Profit

Run:
  python corro_quarter.py --quarter q1 --year 2026
  python corro_quarter.py --quarter q2 --year 2026

Secrets:
  SHOPIFY_TOKEN_CORRO   GOOGLE_CREDENTIALS   SHEET_ID_QUARTER
"""

import os, json, requests, gspread, argparse, calendar
from google.oauth2.service_account import Credentials
from datetime import datetime, date
from collections import defaultdict
import pytz, time

TIMEZONE    = pytz.timezone("America/Bogota")
API_VERSION = "2024-10"
STORE_URL   = os.environ.get("SHOPIFY_STORE", "equestrian-labs.myshopify.com")
TOKEN       = os.environ.get("SHOPIFY_TOKEN_CORRO", "")
SHEET_ID    = os.environ.get("SHEET_ID_QUARTER", "")
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets",
               "https://www.googleapis.com/auth/drive"]

QUARTER_MONTHS = {
    "q1":[1,2,3], "q2":[4,5,6], "q3":[7,8,9], "q4":[10,11,12]
}
MONTH_NAMES = {
    1:"January",2:"February",3:"March",4:"April",
    5:"May",6:"June",7:"July",8:"August",
    9:"September",10:"October",11:"November",12:"December"
}

# ── 100% discount / free item categories ─────────────────────────
DISC_ZERO_CATS = [
    ("Advent Calendar",  ["advent calendar","advent_calendar"]),
    ("Team Rider",       ["team rider","team_rider"]),
    ("Influencer",       ["influencer"]),
    ("Marketing",        ["marketing","marketing - sponsorship"]),
    ("Sponsorship",      ["sponsorship","sponsor"]),
    ("Internal / Staff", ["staff","internal","employee"]),
    ("Elite Cart Gift",  ["elite cart","elite_cart","elitecart"]),
]
def classify_disc_zero(tags_str):
    t = (tags_str or "").lower()
    for cat, kws in DISC_ZERO_CATS:
        if any(k in t for k in kws):
            return cat
    return "Other"


def money(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0

def order_shipping_paid(order):
    """Order-level shipping paid, used for staff/internal audit rows."""
    total = 0.0
    for sl in (order.get("shipping_lines") or []):
        total += money(sl.get("price"))
    if total:
        return total
    tsp = order.get("total_shipping_price_set") or {}
    shop_money = tsp.get("shop_money") or {}
    return money(shop_money.get("amount"))

def customer_identity(order):
    c = order.get("customer") or {}
    first = (c.get("first_name") or "").strip()
    last  = (c.get("last_name") or "").strip()
    name  = (f"{first} {last}".strip()) or ""
    email = (c.get("email") or "").strip()
    return name, email

def is_staff_tag(tags_str):
    t = (tags_str or "").lower()
    return any(k in t for k in ["staff", "internal", "employee"])

# ── Elite Cart manual COGS overrides (product_title → unit_cost) ──
ELITE_CART_COGS = {
    "mystery cavali": 44.00,
    "equine moscow":   8.94,
}

def quarter_range(q, y):
    months = QUARTER_MONTHS[q]
    start  = date(y, months[0], 1)
    end    = date(y, months[-1], calendar.monthrange(y, months[-1])[1])
    today  = datetime.now(TIMEZONE).date()
    return start, min(end, today), months

# ── SHOPIFY REST ──────────────────────────────────────────────────
def shopify_get(endpoint, params):
    url = f"https://{STORE_URL}/admin/api/{API_VERSION}/{endpoint}"
    headers = {"X-Shopify-Access-Token": TOKEN}
    results = []
    while url:
        for attempt in range(8):
            try:
                r = requests.get(url, headers=headers, params=params, timeout=60)
            except requests.exceptions.ConnectionError as e:
                time.sleep(min(2**attempt,60)); continue
            if r.status_code == 429:
                time.sleep(int(r.headers.get("Retry-After", 2**attempt))); continue
            if r.status_code in (502,503,504):
                time.sleep(min(2**attempt,60)); continue
            r.raise_for_status(); break
        data = r.json()
        key  = [k for k in data if k != "errors"][0]
        results.extend(data[key])
        link = r.headers.get("Link",""); url = None; params = {}
        if 'rel="next"' in link:
            for part in link.split(","):
                if 'rel="next"' in part:
                    url = part.split(";")[0].strip().strip("<>")
        time.sleep(0.3)
    return results

def shopify_graphql(query, variables=None):
    url = f"https://{STORE_URL}/admin/api/2026-01/graphql.json"
    headers = {"X-Shopify-Access-Token": TOKEN, "Content-Type": "application/json"}
    payload = {"query": query}
    if variables: payload["variables"] = variables
    for attempt in range(8):
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=60)
        except: time.sleep(min(2**attempt,60)); continue
        if r.status_code == 429: time.sleep(int(r.headers.get("Retry-After",2**attempt))); continue
        if r.status_code in (502,503,504): time.sleep(min(2**attempt,60)); continue
        r.raise_for_status()
        d = r.json()
        if any((e.get("extensions") or {}).get("code")=="THROTTLED" for e in (d.get("errors") or [])):
            time.sleep(min(2**attempt,60)); continue
        return d
    raise RuntimeError("GraphQL failed")

# ── FETCH ─────────────────────────────────────────────────────────
def fetch_orders(start, end):
    print(f"  Fetching orders {start} → {end}...")
    seen, all_orders = set(), []
    for status in ["paid,partially_paid","partially_refunded,refunded"]:
        batch = shopify_get("orders.json", {
            "status":"any","financial_status":status,
            "created_at_min":f"{start}T00:00:00-05:00",
            "created_at_max":f"{end}T23:59:59-05:00",
            "limit":250,
            "fields":"id,name,created_at,financial_status,subtotal_price,total_price,"
                     "total_discounts,discount_codes,source_name,tags,line_items,customer,"
                     "shipping_lines,total_shipping_price_set",
        })
        for o in batch:
            oid = o.get("id")
            if oid and oid not in seen:
                seen.add(oid); all_orders.append(o)
    print(f"  → {len(all_orders)} unique orders")
    return all_orders

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
    }"""
    vmap = {}; all_product_ids = set(); cursor = None
    while True:
        data = shopify_graphql(QUERY, {"cursor": cursor})
        pv   = (data.get("data") or {}).get("productVariants", {})
        for edge in pv.get("edges", []):
            node = edge["node"]
            vid  = node["id"].split("/")[-1]
            prod = node.get("product") or {}
            inv  = node.get("inventoryItem") or {}
            uc   = inv.get("unitCost") or {}
            sku  = node.get("sku") or ""
            pid  = prod.get("id","").split("/")[-1]
            title = prod.get("title","")
            # Apply Elite Cart manual COGS override if unit cost is missing
            unit_cost = float(uc.get("amount") or 0)
            if unit_cost == 0:
                for key, override_cost in ELITE_CART_COGS.items():
                    if key in title.lower():
                        unit_cost = override_cost
                        break
            vmap[vid] = {
                "product_id":   pid,
                "product_title":title,
                "sku":          sku,
                "vendor":       prod.get("vendor",""),
                "product_type": prod.get("productType") or "Uncategorized",
                "unit_cost":    unit_cost,
            }
            all_product_ids.add(pid)
        pi = pv.get("pageInfo",{})
        if not pi.get("hasNextPage"): break
        cursor = pi["endCursor"]; time.sleep(0.3)
    filled = sum(1 for v in vmap.values() if v["unit_cost"] > 0)
    print(f"  → {len(vmap)} variants | {filled} with COGS > 0")
    return vmap, all_product_ids

def fetch_shopifyql(start, end):
    """ShopifyQL: gross_profit per product from Shopify Analytics directly."""
    print(f"  Fetching ShopifyQL gross_profit {start} → {end}...")
    GQL = """
    query($q: String!) {
      shopifyqlQuery(query: $q) {
        tableData { columns { name dataType } rows }
        parseErrors
      }
    }"""
    q = (f"FROM sales SHOW gross_profit, gross_sales, net_sales, orders, net_items_sold "
         f"GROUP BY product_title, product_vendor "
         f"SINCE {start} UNTIL {end} ORDER BY gross_profit DESC")
    d  = shopify_graphql(GQL, {"q": q})
    sq = (d.get("data") or {}).get("shopifyqlQuery")
    if not sq or sq.get("parseErrors"):
        print(f"  ⚠ ShopifyQL unavailable — GP estimated from COGS"); return {}
    table = sq.get("tableData") or {}
    cols  = [c["name"] for c in (table.get("columns") or [])]
    rows  = table.get("rows") or []
    result = {}
    for row in rows:
        if not isinstance(row, dict): row = dict(zip(cols, row))
        t = (row.get("product_title") or "").strip()
        if not t: continue
        result[t.lower()] = {
            "gross_profit": float(row.get("gross_profit") or 0),
            "gross_sales":  float(row.get("gross_sales")  or 0),
            "net_sales":    float(row.get("net_sales")    or 0),
            "vendor":       (row.get("product_vendor") or "").strip(),
        }
    print(f"  → {len(result)} products from ShopifyQL"); return result

# ── AGGREGATE ─────────────────────────────────────────────────────
def mk_sku_row(info, months):
    return {
        "product_id":   info.get("product_id",""),
        "product_title":info.get("product_title",""),
        "sku":          info.get("sku",""),
        "vendor":       info.get("vendor",""),
        "product_type": info.get("product_type","Uncategorized"),
        "unit_cost":    info.get("unit_cost",0.0),
        "gross_sales":  0.0,"discounts":0.0,"returns":0.0,
        "net_sales":    0.0,"cogs":0.0,"gross_profit":0.0,
        "units":        0,"orders":set(),
        "monthly":      {m:{"gross_sales":0.0,"net_sales":0.0,"units":0} for m in months},
    }

def aggregate(orders, vmap, months):
    """Returns SKU aggregates, 100% discount rows, and staff/internal rows."""
    sku_agg = {}
    disc_zero_rows = []
    staff_rows = []

    for order in orders:
        om = int((order.get("created_at", "")[:7]).split("-")[-1] or 0)
        tags_str = order.get("tags") or ""
        category = classify_disc_zero(tags_str)
        staff_flag = category == "Internal / Staff" or is_staff_tag(tags_str)
        shipping_paid = order_shipping_paid(order)
        customer_name, customer_email = customer_identity(order)
        is_refund = (order.get("financial_status") or "").startswith("refund")

        line_items = order.get("line_items", []) or []
        order_discount_total = money(order.get("total_discounts"))
        gross_line_total = sum(money(li.get("price")) * int(li.get("quantity", 0) or 0) for li in line_items) or 0.0

        for li in line_items:
            vid = str(li.get("variant_id") or "")
            info = vmap.get(vid, {
                "product_id": str(li.get("product_id", "")),
                "product_title": li.get("title", "Unknown"),
                "sku": li.get("sku", ""),
                "vendor": "",
                "product_type": "Uncategorized",
                "unit_cost": 0.0,
            })
            qty = int(li.get("quantity", 0) or 0)
            gross = money(li.get("price")) * qty

            # Prefer Shopify's line-level discount. If it is missing, allocate the
            # order discount proportionally by gross line value. Never let a line's
            # discount exceed its own gross value; this avoids impossible rows such
            # as original price $5 / discount $16.
            raw_line_disc = money(li.get("total_discount"))
            if raw_line_disc <= 0 and order_discount_total > 0 and gross_line_total > 0:
                raw_line_disc = order_discount_total * (gross / gross_line_total)
            disc = min(raw_line_disc, gross) if gross > 0 else max(raw_line_disc, 0)
            net = max(gross - disc, 0)
            discount_pct = round((disc / gross), 4) if gross > 0 else (1.0 if net < 0.01 else 0.0)

            cost = info["unit_cost"] * qty
            ret = -net if is_refund else 0.0
            gross_profit = net - cost

            key = vid or f"nv_{info['product_id']}_{info['product_title']}"
            if key not in sku_agg:
                sku_agg[key] = mk_sku_row(info, months)
            r = sku_agg[key]
            r["gross_sales"] += gross
            r["discounts"] += disc
            r["returns"] += ret
            r["net_sales"] += net
            r["cogs"] += cost
            r["gross_profit"] += gross_profit
            r["units"] += qty
            r["orders"].add(order["id"])
            if om in months:
                r["monthly"][om]["gross_sales"] += gross
                r["monthly"][om]["net_sales"] += net
                r["monthly"][om]["units"] += qty

            # Discount Zero = customer paid $0 for the line item.
            # It is only considered a 100% discount when net paid is zero and the
            # discount covers the full gross line amount, or when the item price is
            # genuinely $0 in Shopify.
            is_free_price = gross <= 0.01 and net <= 0.01
            is_full_discount = gross > 0.01 and net <= 0.01 and discount_pct >= 0.999
            if is_free_price or is_full_discount:
                row = {
                    "order_name": order.get("name", ""),
                    "created_at": order.get("created_at", "")[:10],
                    "month": MONTH_NAMES.get(om, ""),
                    "category": category,
                    "tags": tags_str,
                    "customer_name": customer_name,
                    "customer_email": customer_email,
                    "product_title": info["product_title"],
                    "sku": info["sku"],
                    "vendor": info["vendor"] or "Unknown",
                    "product_type": info["product_type"],
                    "units": qty,
                    "gross_sales": gross,
                    "original_price": gross,
                    "discount": disc,
                    "discount_pct": discount_pct,
                    "net_paid": net,
                    "unit_cost": info["unit_cost"],
                    "cogs": cost,
                    "gross_profit": gross_profit,
                    "shipping_paid": shipping_paid,
                }
                disc_zero_rows.append(row)
                if staff_flag:
                    expected_item_payment = round(cost * 1.10, 2)
                    row_staff = dict(row)
                    row_staff.update({
                        "expected_item_payment": expected_item_payment,
                        "payment_gap": round(net - expected_item_payment, 2),
                        "is_dropship": "YES" if info["unit_cost"] > 0 else "NO",
                    })
                    staff_rows.append(row_staff)

    return sku_agg, disc_zero_rows, staff_rows

def build_product_agg(sku_agg, sq_map, months):
    """Group SKUs by product_title."""
    prod = defaultdict(lambda: {
        "product_title":"","vendor":"","product_type":"",
        "gross_sales":0.0,"discounts":0.0,"returns":0.0,
        "net_sales":0.0,"cogs":0.0,"gross_profit":0.0,
        "units":0,"orders":set(),
        "monthly":{m:{"gross_sales":0.0,"net_sales":0.0,"units":0} for m in months},
        "is_dropship":"NO",
    })
    for row in sku_agg.values():
        t = row["product_title"]
        p = prod[t]
        p["product_title"] = t
        p["vendor"]        = p["vendor"] or row["vendor"]
        p["product_type"]  = p["product_type"] or row["product_type"]
        p["is_dropship"]   = "YES" if row["unit_cost"] > 0 else p["is_dropship"]
        p["gross_sales"]  += row["gross_sales"]
        p["discounts"]    += row["discounts"]
        p["returns"]      += row["returns"]
        p["net_sales"]    += row["net_sales"]
        p["cogs"]         += row["cogs"]
        p["gross_profit"] += row["gross_profit"]
        p["units"]        += row["units"]
        p["orders"]       |= row["orders"]
        for m in months:
            p["monthly"][m]["gross_sales"] += row["monthly"][m]["gross_sales"]
            p["monthly"][m]["net_sales"]   += row["monthly"][m]["net_sales"]
            p["monthly"][m]["units"]       += row["monthly"][m]["units"]
    # Use ShopifyQL GP when available, but do not overwrite a valid calculated
    # gross profit with blank/zero ShopifyQL values. This keeps the dashboard from
    # showing GP as zero when ShopifyQL is unavailable or returns incomplete data.
    for t, p in prod.items():
        sq = sq_map.get(t.lower())
        if sq:
            if abs(sq.get("gross_profit", 0)) > 0.01:
                p["gross_profit"] = sq["gross_profit"]
            if abs(sq.get("gross_sales", 0)) > 0.01:
                p["gross_sales"] = sq["gross_sales"]
            if abs(sq.get("net_sales", 0)) > 0.01:
                p["net_sales"] = sq["net_sales"]
            p["vendor"] = p["vendor"] or sq.get("vendor", "")
    return prod

def build_vendor_agg(prod_agg):
    vend = defaultdict(lambda: {
        "vendor":"","skus":set(),"units":0,
        "gross_sales":0.0,"discounts":0.0,"returns":0.0,
        "net_sales":0.0,"cogs":0.0,"gross_profit":0.0,
    })
    for p in prod_agg.values():
        v = p["vendor"] or "Unknown"
        vend[v]["vendor"]      = v
        vend[v]["skus"].add(p["product_title"])
        vend[v]["units"]      += p["units"]
        vend[v]["gross_sales"]+= p["gross_sales"]
        vend[v]["discounts"]  += p["discounts"]
        vend[v]["returns"]    += p["returns"]
        vend[v]["net_sales"]  += p["net_sales"]
        vend[v]["cogs"]       += p["cogs"]
        vend[v]["gross_profit"]+= p["gross_profit"]
    return vend

def build_disc_agg(orders):
    agg = defaultdict(lambda:{"discount_code":"","discount_type":"","total_discounts":0.0,"net_sales":0.0,"orders":0})
    for order in orders:
        disc = float(order.get("total_discounts",0) or 0)
        if disc <= 0: continue
        sub  = float(order.get("subtotal_price",0) or 0)
        codes = order.get("discount_codes") or []
        if codes:
            for dc in codes:
                code = (dc.get("code") or "").strip() or "(no code)"
                agg[code]["discount_code"]    = code
                agg[code]["discount_type"]    = dc.get("type","")
                agg[code]["total_discounts"] += float(dc.get("amount") or 0)
                agg[code]["net_sales"]       += sub
                agg[code]["orders"]          += 1
        else:
            agg["(automatic)"]["discount_code"]    = "(automatic)"
            agg["(automatic)"]["discount_type"]    = "Automatic"
            agg["(automatic)"]["total_discounts"] += disc
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
        try: return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = e.response.status_code if hasattr(e,"response") else 0
            if status==429 or status>=500:
                wait=min(15*(attempt+1),90); print(f"    Sheets {status}—wait {wait}s"); time.sleep(wait); continue
            raise
        except Exception as e:
            if attempt<3: print(f"    Sheets err: {e}"); time.sleep(10); continue
            raise

def write_tab(sh, name, headers, rows):
    try:    ws = sh.worksheet(name)
    except: ws = sheets_call(sh.add_worksheet, name, rows=max(5000,len(rows)+100), cols=len(headers)+2)
    time.sleep(2)
    all_data = [headers]
    for r in rows:
        clean = []
        for v in r:
            if v is None: clean.append("")
            elif isinstance(v,float) and v!=v: clean.append("")
            elif isinstance(v,set): clean.append(str(len(v)))
            else: clean.append(str(v) if not isinstance(v,(int,float,bool)) else v)
        all_data.append(clean)
    total = len(all_data)
    if total>ws.row_count or len(headers)>ws.col_count:
        sheets_call(ws.resize,rows=total+50,cols=len(headers)+2); time.sleep(1)
    sheets_call(ws.clear); time.sleep(1)
    BATCH=500
    for i in range(0,total,BATCH):
        sheets_call(ws.append_rows,all_data[i:i+BATCH],value_input_option="RAW",insert_data_option="INSERT_ROWS")
        print(f"    {name}: {min(i+BATCH,total)}/{total}")
        if i+BATCH<total: time.sleep(2)
    time.sleep(3)
    print(f"    ✓ {name}: {len(rows)} rows")

# ── MAIN ──────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--quarter",default="q1",choices=["q1","q2","q3","q4"])
    parser.add_argument("--year",type=int,default=2026)
    args = parser.parse_args()
    q, y = args.quarter, args.year
    start, end, months = quarter_range(q, y)
    label = f"{q.upper()} {y}"
    mnames = [MONTH_NAMES[m] for m in months]
    sfx   = f"{q}_{y}"
    now   = datetime.now(TIMEZONE).strftime("%Y-%m-%d %H:%M")

    print(f"\n{'='*60}")
    print(f"  CORRO QUARTERLY REPORT v2.0 — {label}")
    print(f"  {start} → {end} | Months: {', '.join(mnames)}")
    print(f"{'='*60}\n")

    vmap, all_pids = fetch_product_map()
    orders         = fetch_orders(start, end)
    sq_map         = fetch_shopifyql(start, end)
    sku_agg, disc_zero_rows, staff_rows = aggregate(orders, vmap, months)
    prod_agg       = build_product_agg(sku_agg, sq_map, months)
    vendor_agg     = build_vendor_agg(prod_agg)
    disc_agg       = build_disc_agg(orders)

    sh = get_gc().open_by_key(SHEET_ID)

    # ── Totals ────────────────────────────────────────────────────
    total_gs   = sum(p["gross_sales"]   for p in prod_agg.values())
    total_disc = sum(p["discounts"]     for p in prod_agg.values())
    total_ret  = sum(p["returns"]       for p in prod_agg.values())
    total_ns   = sum(p["net_sales"]     for p in prod_agg.values())
    total_cogs = sum(p["cogs"]          for p in prod_agg.values())
    total_gp   = sum(p["gross_profit"]  for p in prod_agg.values())
    total_units= sum(p["units"]         for p in prod_agg.values())
    total_ords = len(set(o["id"] for o in orders))
    gp_margin  = round(total_gp/total_ns*100,1) if total_ns else 0

    monthly_gs = {m:sum(p["monthly"][m]["gross_sales"] for p in prod_agg.values()) for m in months}
    monthly_ns = {m:sum(p["monthly"][m]["net_sales"]   for p in prod_agg.values()) for m in months}
    monthly_u  = {m:sum(p["monthly"][m]["units"]       for p in prod_agg.values()) for m in months}

    # ── 1. SUMMARY ────────────────────────────────────────────────
    h = ["updated_at","quarter","metric",mnames[0],mnames[1],mnames[2],"Q Total"]
    monthly_gp_approx = {}
    for m in months:
        m_ns = monthly_ns[m] or 0
        monthly_gp_approx[m] = round(total_gp * (m_ns / total_ns), 2) if total_ns else 0

    rows_s = [
        [now,label,"Gross Sales",  round(monthly_gs[months[0]],2),round(monthly_gs[months[1]],2),round(monthly_gs[months[2]],2),round(total_gs,2)],
        [now,label,"Discounts",    "","","",round(total_disc,2)],
        [now,label,"Returns",      "","","",round(total_ret,2)],
        [now,label,"Net Sales",    round(monthly_ns[months[0]],2),round(monthly_ns[months[1]],2),round(monthly_ns[months[2]],2),round(total_ns,2)],
        [now,label,"COGS",         "","","",round(total_cogs,2)],
        [now,label,"Gross Profit", round(monthly_gp_approx[months[0]],2),round(monthly_gp_approx[months[1]],2),round(monthly_gp_approx[months[2]],2),round(total_gp,2)],
        [now,label,"Gross Margin", "","","",f"{gp_margin}%"],
        [now,label,"Units Sold",   monthly_u[months[0]],monthly_u[months[1]],monthly_u[months[2]],total_units],
        [now,label,"Orders",       "","","",total_ords],
    ]
    write_tab(sh, f"q_summary_{sfx}", h, rows_s)

    # Helper: format rows for top-N sheets
    def prod_row(i, p, rank_by="gross_profit"):
        ns  = p["net_sales"] or 0
        gp  = p["gross_profit"] or 0
        margin = round(gp/ns,4) if ns else 0
        total_sales = round(p["gross_sales"],2)
        return [
            now, label, i+1,
            p["product_title"], p.get("vendor",""), p.get("product_type",""),
            p.get("is_dropship","NO"), p["units"],
            round(p["gross_sales"],2), round(p["discounts"],2),
            round(p["returns"],2),     round(ns,2),
            round(p["cogs"],2),        round(gp,2),
            round(margin,4), total_sales, len(p["orders"]),
            round(p["monthly"][months[0]]["net_sales"],2),
            round(p["monthly"][months[1]]["net_sales"],2),
            round(p["monthly"][months[2]]["net_sales"],2),
            p["monthly"][months[0]]["units"],
            p["monthly"][months[1]]["units"],
            p["monthly"][months[2]]["units"],
        ]

    h_prod = ["updated_at","quarter","rank","product_title","vendor","product_type",
              "is_dropship","units","gross_sales","discounts","returns","net_sales",
              "cogs","gross_profit","gross_margin","total_sales","orders",
              f"{mnames[0]} net_sales",f"{mnames[1]} net_sales",f"{mnames[2]} net_sales",
              f"{mnames[0]} units",f"{mnames[1]} units",f"{mnames[2]} units"]

    # Helper: format rows for SKU-level
    def sku_row(i, r):
        ns  = r["net_sales"] or 0
        gp  = r["gross_profit"] or 0
        margin = round(gp/ns,4) if ns else 0
        return [
            now, label, i+1,
            r["product_title"], r["sku"], r.get("vendor",""), r.get("product_type",""),
            "YES" if r["unit_cost"]>0 else "NO", r["units"],
            round(r["gross_sales"],2), round(r["discounts"],2),
            round(r["returns"],2),     round(ns,2),
            round(r["cogs"],2),        round(gp,2),
            round(margin,3), round(r["gross_sales"],2), len(r["orders"]),
            round(r["monthly"][months[0]]["net_sales"],2),
            round(r["monthly"][months[1]]["net_sales"],2),
            round(r["monthly"][months[2]]["net_sales"],2),
        ]

    h_sku = ["updated_at","quarter","rank","product_title","sku","vendor","product_type",
             "is_dropship","units","gross_sales","discounts","returns","net_sales",
             "cogs","gross_profit","gross_margin","total_sales","orders",
             f"{mnames[0]} net_sales",f"{mnames[1]} net_sales",f"{mnames[2]} net_sales"]

    # ── 2. TOP 100 BY SKU ─────────────────────────────────────────
    sorted_skus = sorted(sku_agg.values(), key=lambda x: x["gross_profit"], reverse=True)
    write_tab(sh, f"q_top_sku_{sfx}", h_sku,
              [sku_row(i,r) for i,r in enumerate(sorted_skus[:100])])

    # ── 3. TOP 100 ALL PRODUCTS ───────────────────────────────────
    sorted_prods = sorted(prod_agg.values(), key=lambda x: x["gross_profit"], reverse=True)
    write_tab(sh, f"q_top_products_{sfx}", h_prod,
              [prod_row(i,p) for i,p in enumerate(sorted_prods[:100])])

    # ── 4. TOP 100 DROPSHIP ───────────────────────────────────────
    drop_prods = sorted([p for p in prod_agg.values() if p.get("is_dropship")=="YES"],
                        key=lambda x: x["gross_sales"], reverse=True)
    write_tab(sh, f"q_top_dropship_{sfx}", h_prod,
              [prod_row(i,p) for i,p in enumerate(drop_prods[:100])])

    # ── 5. TOP 100 BEST MARGIN (min 5 units) ─────────────────────
    margin_prods = sorted(
        [p for p in prod_agg.values() if p["units"]>=5 and p["net_sales"]>0],
        key=lambda x: (x["gross_profit"]/x["net_sales"] if x["net_sales"] else 0), reverse=True)
    write_tab(sh, f"q_top_margin_{sfx}", h_prod,
              [prod_row(i,p) for i,p in enumerate(margin_prods[:100])])

    # ── 6. TOP 100 GROSS PROFIT ───────────────────────────────────
    gp_prods = sorted(prod_agg.values(), key=lambda x: x["gross_profit"], reverse=True)
    write_tab(sh, f"q_top_gp_{sfx}", h_prod,
              [prod_row(i,p) for i,p in enumerate(gp_prods[:100])])

    # ── 7. COST ZERO (COGS = 0, but product sold) ─────────────────
    cost_zero = sorted(
        [p for p in prod_agg.values() if p["cogs"] == 0 and p["gross_sales"] > 0],
        key=lambda x: x["gross_sales"], reverse=True)
    h_cz = ["updated_at","quarter","rank","product_title","vendor","product_type",
            "units","gross_sales","discounts","net_sales","cogs","gross_profit",
            "gross_margin","orders"]
    write_tab(sh, f"q_cost_zero_{sfx}", h_cz,
              [[now,label,i+1,p["product_title"],p.get("vendor","") or "Unknown",p.get("product_type",""),
                p["units"],round(p["gross_sales"],2),round(p["discounts"],2),round(p["net_sales"],2),
                round(p["cogs"],2),round(p["gross_profit"],2),
                round((p["gross_profit"]/p["net_sales"]),4) if p["net_sales"] else 0,
                len(p["orders"])]
               for i,p in enumerate(cost_zero)])

    # ── 8. DISCOUNT ZERO (100% discount / customer paid $0) ───────
    cat_month = defaultdict(lambda: {m:{"units":0,"cogs":0.0,"discount":0.0,"products":set()} for m in months})
    for d in disc_zero_rows:
        om = next((m for m in months if MONTH_NAMES[m] == d["month"]), None)
        if om:
            cat_month[d["category"]][om]["units"] += d["units"]
            cat_month[d["category"]][om]["cogs"] += d["cogs"]
            cat_month[d["category"]][om]["discount"] += d["discount"]
            cat_month[d["category"]][om]["products"].add(d["product_title"])

    all_cats = [c for c,_ in DISC_ZERO_CATS] + ["Other"]
    h_dz = ["updated_at","quarter","category",
            f"{mnames[0]} units",f"{mnames[0]} discount",f"{mnames[0]} cogs",
            f"{mnames[1]} units",f"{mnames[1]} discount",f"{mnames[1]} cogs",
            f"{mnames[2]} units",f"{mnames[2]} discount",f"{mnames[2]} cogs",
            "total units","total discount","total cogs","sample products"]
    dz_rows = []
    for cat in all_cats:
        if cat not in cat_month: continue
        cm = cat_month[cat]
        all_prods = set()
        for m in months: all_prods |= cm[m]["products"]
        dz_rows.append([
            now, label, cat,
            cm[months[0]]["units"], round(cm[months[0]]["discount"],2), round(cm[months[0]]["cogs"],2),
            cm[months[1]]["units"], round(cm[months[1]]["discount"],2), round(cm[months[1]]["cogs"],2),
            cm[months[2]]["units"], round(cm[months[2]]["discount"],2), round(cm[months[2]]["cogs"],2),
            sum(cm[m]["units"] for m in months),
            round(sum(cm[m]["discount"] for m in months),2),
            round(sum(cm[m]["cogs"] for m in months),2),
            ", ".join(sorted(all_prods)[:5]),
        ])
    write_tab(sh, f"q_discount_zero_{sfx}", h_dz, dz_rows)

    # Order-by-order detail for 100% discounts
    h_dz_det = ["updated_at","quarter","order_name","created_at","month","category",
                "customer_name","customer_email","tags","product_title","sku","vendor","product_type",
                "units","gross_sales","discount","discount_pct","net_paid","unit_cost","cogs",
                "gross_profit","shipping_paid"]
    write_tab(sh, f"q_discount_zero_detail_{sfx}", h_dz_det,
              [[now,label]+[d.get(k,"") for k in ["order_name","created_at","month","category",
               "customer_name","customer_email","tags","product_title","sku","vendor","product_type",
               "units","gross_sales","discount","discount_pct","net_paid","unit_cost","cogs",
               "gross_profit","shipping_paid"]]
               for d in sorted(disc_zero_rows, key=lambda x: x["created_at"])])

    # ── 8b. STAFF / INTERNAL AUDIT ────────────────────────────────
    h_staff = ["updated_at","quarter","order_name","created_at","month","customer_name","customer_email",
               "category","tags","product_title","sku","vendor","product_type","is_dropship",
               "units","gross_sales","discount","discount_pct","net_paid","unit_cost","cogs",
               "expected_item_payment","payment_gap","gross_profit","shipping_paid"]
    write_tab(sh, f"q_staff_orders_{sfx}", h_staff,
              [[now,label]+[d.get(k,"") for k in ["order_name","created_at","month","customer_name","customer_email",
               "category","tags","product_title","sku","vendor","product_type","is_dropship",
               "units","gross_sales","discount","discount_pct","net_paid","unit_cost","cogs",
               "expected_item_payment","payment_gap","gross_profit","shipping_paid"]]
               for d in sorted(staff_rows, key=lambda x: x["created_at"])])

    # ── 8c. UNKNOWN / MISSING VENDOR LIST ─────────────────────────
    unknown_vendor = sorted(
        [p for p in prod_agg.values() if (not (p.get("vendor") or "").strip()) or (p.get("vendor","").strip().lower() == "unknown")],
        key=lambda x: x["gross_sales"], reverse=True)
    h_uv = ["updated_at","quarter","rank","product_title","vendor","product_type",
            "units","gross_sales","discounts","net_sales","cogs","gross_profit",
            "gross_margin","orders"]
    write_tab(sh, f"q_unknown_vendors_{sfx}", h_uv,
              [[now,label,i+1,p["product_title"],p.get("vendor","") or "Unknown",p.get("product_type",""),
                p["units"],round(p["gross_sales"],2),round(p["discounts"],2),round(p["net_sales"],2),
                round(p["cogs"],2),round(p["gross_profit"],2),
                round((p["gross_profit"]/p["net_sales"]),4) if p["net_sales"] else 0,
                len(p["orders"])]
               for i,p in enumerate(unknown_vendor)])

    # ── 9. MONTHLY BREAKDOWN ──────────────────────────────────────
    h_mb = ["updated_at","quarter","rank","product_title","sku","vendor",
            f"{mnames[0]} units",f"{mnames[0]} sales",f"{mnames[0]} margin",
            f"{mnames[1]} units",f"{mnames[1]} sales",f"{mnames[1]} margin",
            f"{mnames[2]} units",f"{mnames[2]} sales",f"{mnames[2]} margin",
            "Q total units","Q total sales","Q avg margin"]
    mb_rows = []
    for i, r in enumerate(sorted(sku_agg.values(), key=lambda x: x["gross_profit"], reverse=True)[:100]):
        ns = r["net_sales"] or 0
        gp = r["gross_profit"] or 0
        avg_m = round(gp/ns,3) if ns else 0
        row_data = [now, label, i+1, r["product_title"], r["sku"], r.get("vendor","")]
        for m in months:
            ms = r["monthly"][m]["net_sales"] or 0
            mg = r["gross_profit"] * (ms/ns) if ns else 0
            row_data += [r["monthly"][m]["units"], round(ms,2), round(mg/ms if ms else 0,3)]
        row_data += [r["units"], round(r["gross_sales"],2), avg_m]
        mb_rows.append(row_data)
    write_tab(sh, f"q_monthly_breakdown_{sfx}", h_mb, mb_rows)

    # ── 10. ZERO SALES ────────────────────────────────────────────
    active_pids = {r["product_id"] for r in sku_agg.values()}
    zero_sales  = [v for pid,v in vmap.items()
                   if v["product_id"] not in active_pids and v["product_title"]]
    seen_titles = set(); zs_dedup = []
    for v in sorted(zero_sales, key=lambda x: x["product_title"]):
        if v["product_title"] not in seen_titles:
            seen_titles.add(v["product_title"]); zs_dedup.append(v)
    h_zs = ["updated_at","quarter","rank","product_title","sku","vendor","product_type","is_dropship"]
    write_tab(sh, f"q_zero_sales_{sfx}", h_zs,
              [[now,label,i+1,v["product_title"],v["sku"],v["vendor"],v["product_type"],
                "YES" if v["unit_cost"]>0 else "NO"]
               for i,v in enumerate(zs_dedup)])

    # ── 11. VENDORS PARETO — sorted by gross_profit ───────────────
    sorted_vend = sorted(vendor_agg.values(), key=lambda x: x["gross_profit"], reverse=True)
    total_sales_v = sum(v["gross_sales"] for v in sorted_vend) or 1
    h_v = ["updated_at","quarter","rank","vendor","skus","units","gross_sales","discounts",
           "returns","net_sales","cogs","gross_profit","gross_margin",
           "pct_of_sales","cumulative_pct"]
    vend_rows = []; cum = 0.0
    for i, v in enumerate(sorted_vend):
        cum += v["gross_sales"]
        ns = v["net_sales"] or 0
        gp = v["gross_profit"] or 0
        vend_rows.append([
            now, label, i+1, v["vendor"], len(v["skus"]), v["units"],
            round(v["gross_sales"],2), round(v["discounts"],2),
            round(v["returns"],2),     round(ns,2),
            round(v["cogs"],2),        round(gp,2),
            round(gp/ns,3) if ns else 0,
            round(v["gross_sales"]/total_sales_v,6),
            round(cum/total_sales_v,6),
        ])
    write_tab(sh, f"q_vendors_{sfx}", h_v, vend_rows)

    print(f"\n{'='*60}")
    print(f"  DONE — {label}")
    print(f"  Net Sales: ${total_ns:,.2f} | GP: ${total_gp:,.2f} ({gp_margin}%)")
    print(f"  Sheets written (all suffixed _{sfx}):")
    for t in ["q_summary","q_top_sku","q_top_products","q_top_dropship",
              "q_top_margin","q_top_gp","q_cost_zero","q_discount_zero",
              "q_discount_zero_detail","q_staff_orders","q_unknown_vendors",
              "q_monthly_breakdown","q_zero_sales","q_vendors"]:
        print(f"    {t}_{sfx}")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()
