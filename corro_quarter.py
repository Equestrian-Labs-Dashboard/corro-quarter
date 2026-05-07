"""
CORRO QUARTERLY REPORT — v2.5
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
  q_cost_zero_detail     — order-by-order detail for products with COGS = 0
  q_discount_zero        — 100% discounted / free items by category
  q_discount_zero_detail — order-by-order 100% discount detail
  q_staff_orders         — staff/internal orders flagged separately
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

# Staff audit source of truth. Shopify recommended querying orders by tag and
# reading staff_member from ShopifyQL instead of grouping by customer.
STAFF_ORDER_TAG = os.environ.get("STAFF_ORDER_TAG", "employee order")
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


def chunks(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i+size]

def fetch_order_staff_map(order_ids):
    """Return Shopify staff member / creator info by numeric order id.

    Shopify exposes Order.staffMember in GraphQL for orders that have a
    staff/POS creator and when the access token has permission. If the token
    cannot read this field, the report still runs and the creator columns stay
    blank instead of breaking the quarter export.
    """
    ids = [str(x) for x in order_ids if x]
    if not ids:
        return {}

    print("  Fetching order staff creators via GraphQL...")
    out = {}
    QUERY = """
    query($ids: [ID!]!) {
      nodes(ids: $ids) {
        ... on Order {
          id
          name
          staffMember {
            id
            name
            email
          }
        }
      }
    }"""

    for batch in chunks(ids, 100):
        gql_ids = [f"gid://shopify/Order/{oid}" for oid in batch]
        try:
            d = shopify_graphql(QUERY, {"ids": gql_ids})
        except Exception as e:
            print(f"  ⚠ staffMember lookup failed: {e}")
            return out

        if d.get("errors"):
            print("  ⚠ staffMember unavailable — leaving created_by fields blank")
            # Common when the app/token lacks read_users or store plan access.
            return out

        for node in ((d.get("data") or {}).get("nodes") or []):
            if not node:
                continue
            oid = str(node.get("id", "")).split("/")[-1]
            sm = node.get("staffMember") or {}
            out[oid] = {
                "order_staff_member_id": str(sm.get("id", "")).split("/")[-1] if sm.get("id") else "",
                "order_staff_name": sm.get("name") or "",
                "order_staff_email": sm.get("email") or "",
            }
        time.sleep(0.2)

    filled = sum(1 for v in out.values() if v.get("order_staff_name") or v.get("order_staff_email"))
    print(f"  → staff creators found for {filled} orders")
    return out

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

def norm_key(k):
    return str(k or "").strip().lower().replace(" ", "_").replace("-", "_")

def row_value(row, *names, default=""):
    wanted = {norm_key(n) for n in names}
    for k, v in row.items():
        if norm_key(k) in wanted:
            return v
    return default

def row_float(row, *names):
    return money(row_value(row, *names, default=0))

def row_int(row, *names):
    try:
        return int(round(float(row_value(row, *names, default=0) or 0)))
    except (TypeError, ValueError):
        return 0

def order_lookup_keys(value):
    """Build safe matching keys for Shopify REST/GraphQL/ShopifyQL order ids.

    Shopify surfaces orders differently depending on API/reporting context:
    REST uses numeric id + name (#152362), while ShopifyQL order_id can be
    returned as a number, a name, or a gid-like value depending on the report.
    """
    raw = str(value or "").strip()
    keys = set()
    if not raw:
        return keys
    keys.add(raw)
    keys.add(raw.lstrip("#"))
    if raw.startswith("gid://") or "/" in raw:
        last = raw.rstrip("/").split("/")[-1]
        if last:
            keys.add(last)
            keys.add(f"#{last}")
    if raw.isdigit():
        keys.add(f"#{raw}")
    return {k for k in keys if k}

def find_order_meta(order, meta_map):
    if not meta_map:
        return {}
    candidates = set()
    for field in ("id", "name", "order_number", "number"):
        candidates |= order_lookup_keys(order.get(field))
    for key in candidates:
        if key in meta_map:
            return meta_map[key]
    return {}

def shopifyql_table(query, label):
    """Run a ShopifyQL query and return normalized row dictionaries."""
    print(f"  Fetching ShopifyQL {label}...")
    GQL = """
    query($q: String!) {
      shopifyqlQuery(query: $q) {
        tableData { columns { name dataType } rows }
        parseErrors
      }
    }"""
    d = shopify_graphql(GQL, {"q": query})
    sq = (d.get("data") or {}).get("shopifyqlQuery")
    if not sq or sq.get("parseErrors"):
        print(f"  ⚠ ShopifyQL {label} unavailable — using REST fallback")
        if sq and sq.get("parseErrors"):
            print(f"    parseErrors: {sq.get('parseErrors')}")
        return []
    table = sq.get("tableData") or {}
    cols = [c.get("name", "") for c in (table.get("columns") or [])]
    rows = table.get("rows") or []
    out = []
    for row in rows:
        if not isinstance(row, dict):
            row = dict(zip(cols, row))
        out.append(row)
    print(f"  → {len(out)} rows from ShopifyQL {label}")
    return out

def fetch_shopifyql(start, end):
    """Exact product-level metrics from ShopifyQL / Shopify Analytics."""
    q = (
        f"FROM sales "
        f"SHOW gross_sales, discounts, net_sales, cost_of_goods_sold, "
        f"net_items_sold, orders, gross_profit "
        f"GROUP BY product_title, product_vendor "
        f"SINCE {start} UNTIL {end} "
        f"ORDER BY gross_profit DESC"
    )
    rows = shopifyql_table(q, "product metrics")
    result = {}
    for row in rows:
        title = str(row_value(row, "product_title", "Product title", default="") or "").strip()
        if not title or title.lower() == "summary":
            continue
        result[title.lower()] = {
            "gross_sales": row_float(row, "gross_sales", "Gross sales"),
            "discounts": row_float(row, "discounts", "Discounts"),
            "net_sales": row_float(row, "net_sales", "Net sales"),
            "cogs": row_float(row, "cost_of_goods_sold", "Cost of goods sold"),
            "units": row_int(row, "net_items_sold", "Net items sold"),
            "orders": row_int(row, "orders", "Orders"),
            "gross_profit": row_float(row, "gross_profit", "Gross profit"),
            "vendor": str(row_value(row, "product_vendor", "Product vendor", default="") or "").strip(),
        }
    return result

def fetch_shopifyql_sku(start, end):
    """Exact SKU-level metrics from ShopifyQL / Shopify Analytics.

    This matches the Shopify Analytics query used for Top 100 by SKU:
    FROM sales SHOW net_sales, discounts, cost_of_goods_sold, net_items_sold,
    orders, gross_profit WHERE product_variant_sku IS NOT NULL
    GROUP BY product_variant_sku ORDER BY gross_profit DESC.
    """
    q = (
        f"FROM sales "
        f"SHOW gross_sales, discounts, net_sales, cost_of_goods_sold, "
        f"net_items_sold, orders, gross_profit "
        f"WHERE product_variant_sku IS NOT NULL "
        f"GROUP BY product_variant_sku "
        f"SINCE {start} UNTIL {end} "
        f"ORDER BY gross_profit DESC "
        f"LIMIT 100"
    )
    rows = shopifyql_table(q, "SKU metrics")
    result = {}
    for row in rows:
        sku = str(row_value(row, "product_variant_sku", "Product variant SKU", default="") or "").strip()
        if not sku or sku.lower() == "summary":
            continue
        result[sku] = {
            "sku": sku,
            "gross_sales": row_float(row, "gross_sales", "Gross sales"),
            "discounts": row_float(row, "discounts", "Discounts"),
            "net_sales": row_float(row, "net_sales", "Net sales"),
            "cogs": row_float(row, "cost_of_goods_sold", "Cost of goods sold"),
            "units": row_int(row, "net_items_sold", "Net items sold"),
            "orders": row_int(row, "orders", "Orders"),
            "gross_profit": row_float(row, "gross_profit", "Gross profit"),
        }
    return result

def fetch_shopifyql_staff_orders(start, end, tag=STAFF_ORDER_TAG):
    """Fetch staff/internal orders from ShopifyQL using order_tags + staff_member.

    This is the source of truth for the Staff tab. It intentionally does not use
    the customer as the staff/person key. Customer remains available only as
    audit context in the detail rows.
    """
    safe_tag = str(tag or "employee order").replace("'", "\'")
    q = (
        f"FROM orders "
        f"SHOW order_id, total_price, discount_amount, staff_member "
        f"WHERE order_tags CONTAINS '{safe_tag}' "
        f"SINCE {start} UNTIL {end} "
        f"ORDER BY total_price DESC"
    )
    rows = shopifyql_table(q, f"staff orders by tag '{tag}'")

    result = {}
    for row in rows:
        order_id = str(row_value(row, "order_id", "Order ID", "order", "Order", default="") or "").strip()
        staff_member = str(row_value(row, "staff_member", "Staff member", "staff", default="") or "").strip()
        total_price = row_float(row, "total_price", "Total price")
        discount_amount = row_float(row, "discount_amount", "Discount amount")
        meta = {
            "shopifyql_order_id": order_id,
            "staff_member": staff_member,
            "staff_person_name": staff_member or "Unknown",
            "shopifyql_total_price": total_price,
            "shopifyql_discount_amount": discount_amount,
            "shopifyql_staff_source": "ShopifyQL orders.staff_member",
        }
        for key in order_lookup_keys(order_id):
            result[key] = meta

    filled = sum(1 for v in result.values() if v.get("staff_member"))
    print(f"  → ShopifyQL staff orders mapped: {len(rows)} rows | {filled} staff keys")
    return result


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
        "monthly":      {m:{
            "gross_sales":0.0,"discounts":0.0,"returns":0.0,
            "net_sales":0.0,"cogs":0.0,"gross_profit":0.0,"units":0
        } for m in months},
    }

def aggregate(orders, vmap, months, order_staff_map=None, shopifyql_staff_order_map=None):
    """Returns SKU aggregates, all order line rows, 100% discount rows, and staff/internal audit rows.

    Important:
    - Discount Zero only captures line items paid at $0 / 100% discounted.
    - Staff captures every line from staff/internal/employee tagged orders, even when
      the staff member paid something. This is required to audit whether they paid
      at least COGS + 10% and whether shipping was paid.
    """
    order_staff_map = order_staff_map or {}
    shopifyql_staff_order_map = shopifyql_staff_order_map or {}
    sku_agg = {}
    all_line_rows = []
    disc_zero_rows = []
    staff_rows = []

    for order in orders:
        om = int((order.get("created_at", "")[:7]).split("-")[-1] or 0)
        tags_str = order.get("tags") or ""
        category = classify_disc_zero(tags_str)
        shipping_paid = order_shipping_paid(order)
        customer_name, customer_email = customer_identity(order)
        oid_str = str(order.get("id") or "")

        # GraphQL staffMember is used only as fallback/enrichment.
        # ShopifyQL orders.staff_member is the Staff tab source of truth.
        staff_meta = order_staff_map.get(oid_str, {})
        ql_staff_meta = find_order_meta(order, shopifyql_staff_order_map)

        order_staff_name = (staff_meta.get("order_staff_name") or "").strip()
        order_staff_email = (staff_meta.get("order_staff_email") or "").strip()
        order_staff_member_id = staff_meta.get("order_staff_member_id") or ""

        shopifyql_staff_member = (ql_staff_meta.get("staff_member") or "").strip()
        shopifyql_order_id = ql_staff_meta.get("shopifyql_order_id", "")
        shopifyql_total_price = ql_staff_meta.get("shopifyql_total_price", "")
        shopifyql_discount_amount = ql_staff_meta.get("shopifyql_discount_amount", "")

        # IMPORTANT: Staff report must not group by customer.
        # Customer is only context; the accountable person is staff_member.
        staff_person_name = shopifyql_staff_member or order_staff_name or "Unknown"
        staff_person_email = order_staff_email or ""

        # If ShopifyQL returned staff orders, use that list. If ShopifyQL is
        # unavailable/empty, fall back to the original employee/internal tags so
        # the export does not go blank.
        staff_flag = bool(ql_staff_meta) or (not shopifyql_staff_order_map and (category == "Internal / Staff" or is_staff_tag(tags_str)))
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
            raw_line_disc = sum(money(a.get("amount")) for a in (li.get("discount_allocations") or []))
            if raw_line_disc <= 0:
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
                r["monthly"][om]["discounts"] += disc
                r["monthly"][om]["returns"] += ret
                r["monthly"][om]["net_sales"] += net
                r["monthly"][om]["cogs"] += cost
                r["monthly"][om]["gross_profit"] += gross_profit
                r["monthly"][om]["units"] += qty

            base_audit_row = {
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
                "financial_status": order.get("financial_status", ""),
                "source_name": order.get("source_name", ""),
                "order_id": order.get("id", ""),
                "order_staff_member_id": order_staff_member_id,
                "order_staff_name": order_staff_name,
                "order_staff_email": order_staff_email,
                "staff_member": shopifyql_staff_member,
                "shopifyql_order_id": shopifyql_order_id,
                "shopifyql_total_price": shopifyql_total_price,
                "shopifyql_discount_amount": shopifyql_discount_amount,
                "staff_person_name": staff_person_name,
                "staff_person_email": staff_person_email,
            }
            all_line_rows.append(base_audit_row)

            # Discount Zero = customer paid $0 for the line item.
            # It is only considered a 100% discount when net paid is zero and the
            # discount covers the full gross line amount, or when the item price is
            # genuinely $0 in Shopify.
            is_free_price = gross <= 0.01 and net <= 0.01
            is_full_discount = gross > 0.01 and net <= 0.01 and discount_pct >= 0.999
            if is_free_price or is_full_discount:
                disc_zero_rows.append(base_audit_row)

            # Staff audit must NOT depend on Discount Zero. Staff members may have
            # paid something, so we need every staff/internal/employee tagged line.
            if staff_flag:
                expected_item_payment = round(cost * 1.10, 2)
                row_staff = dict(base_audit_row)
                row_staff.update({
                    "expected_item_payment": expected_item_payment,
                    "payment_gap": round(net - expected_item_payment, 2),
                    "is_dropship": "YES" if info["unit_cost"] > 0 else "NO",
                })
                staff_rows.append(row_staff)

    return sku_agg, disc_zero_rows, staff_rows, all_line_rows

def build_product_agg(sku_agg, sq_map, months):
    """Group SKUs by product_title."""
    prod = defaultdict(lambda: {
        "product_title":"","vendor":"","product_type":"",
        "gross_sales":0.0,"discounts":0.0,"returns":0.0,
        "net_sales":0.0,"cogs":0.0,"gross_profit":0.0,
        "units":0,"orders":set(),"orders_count":0,
        "monthly":{m:{
            "gross_sales":0.0,"discounts":0.0,"returns":0.0,
            "net_sales":0.0,"cogs":0.0,"gross_profit":0.0,"units":0
        } for m in months},
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
            p["monthly"][m]["gross_sales"]  += row["monthly"][m]["gross_sales"]
            p["monthly"][m]["discounts"]    += row["monthly"][m]["discounts"]
            p["monthly"][m]["returns"]      += row["monthly"][m]["returns"]
            p["monthly"][m]["net_sales"]    += row["monthly"][m]["net_sales"]
            p["monthly"][m]["cogs"]         += row["monthly"][m]["cogs"]
            p["monthly"][m]["gross_profit"] += row["monthly"][m]["gross_profit"]
            p["monthly"][m]["units"]        += row["monthly"][m]["units"]
    # ShopifyQL / Shopify Analytics is the source of truth for product-level
    # sales metrics. REST aggregation is kept as fallback and for monthly splits.
    for t, p in prod.items():
        sq = sq_map.get(t.lower())
        if sq:
            p["gross_sales"] = sq.get("gross_sales", p["gross_sales"])
            p["discounts"] = sq.get("discounts", p["discounts"])
            if "returns" in sq:
                p["returns"] = sq["returns"]
            p["net_sales"] = sq.get("net_sales", p["net_sales"])
            p["cogs"] = sq.get("cogs", p["cogs"])
            p["gross_profit"] = sq.get("gross_profit", p["gross_profit"])
            p["units"] = sq.get("units", p["units"])
            p["orders_count"] = sq.get("orders", len(p["orders"]))
            p["vendor"] = p["vendor"] or sq.get("vendor", "")
    return prod

def build_exact_sku_rows(sku_agg, sq_sku_map, vmap, months):
    """Return SKU rows with ShopifyQL metrics as the source of truth.

    The REST order aggregation is still used for descriptive fields and monthly
    columns, but the main SKU metrics (net sales, discounts, COGS, units,
    orders, gross profit) come from ShopifyQL so they match Shopify Analytics.
    """
    grouped = {}

    def merge_into(dst, src):
        dst["gross_sales"] += src.get("gross_sales", 0)
        dst["discounts"] += src.get("discounts", 0)
        dst["returns"] += src.get("returns", 0)
        dst["net_sales"] += src.get("net_sales", 0)
        dst["cogs"] += src.get("cogs", 0)
        dst["gross_profit"] += src.get("gross_profit", 0)
        dst["units"] += src.get("units", 0)
        dst["orders"] |= src.get("orders", set())
        for m in months:
            dst["monthly"][m]["gross_sales"] += src["monthly"][m]["gross_sales"]
            dst["monthly"][m]["discounts"] += src["monthly"][m]["discounts"]
            dst["monthly"][m]["returns"] += src["monthly"][m]["returns"]
            dst["monthly"][m]["net_sales"] += src["monthly"][m]["net_sales"]
            dst["monthly"][m]["cogs"] += src["monthly"][m]["cogs"]
            dst["monthly"][m]["gross_profit"] += src["monthly"][m]["gross_profit"]
            dst["monthly"][m]["units"] += src["monthly"][m]["units"]

    for src in sku_agg.values():
        sku = (src.get("sku") or "").strip()
        if not sku:
            continue
        if sku not in grouped:
            grouped[sku] = mk_sku_row(src, months)
            grouped[sku]["orders"] = set()
        merge_into(grouped[sku], src)

    if not sq_sku_map:
        return list(grouped.values())

    info_by_sku = {}
    for info in vmap.values():
        sku = (info.get("sku") or "").strip()
        if sku and sku not in info_by_sku:
            info_by_sku[sku] = info

    exact_rows = []
    for sku, exact in sq_sku_map.items():
        base = grouped.get(sku)
        if not base:
            base = mk_sku_row(info_by_sku.get(sku, {
                "sku": sku,
                "product_title": "",
                "vendor": "",
                "product_type": "Uncategorized",
                "unit_cost": 0.0,
                "product_id": "",
            }), months)
        row = dict(base)
        row["monthly"] = {m: dict(base["monthly"][m]) for m in months}
        row["sku"] = sku
        # ShopifyQL is the source of truth for the main metrics.
        row["gross_sales"] = exact.get("gross_sales", 0.0)
        row["discounts"] = exact.get("discounts", 0.0)
        row["returns"] = exact.get("returns", base.get("returns", 0.0))
        row["net_sales"] = exact.get("net_sales", 0.0)
        row["cogs"] = exact.get("cogs", 0.0)
        row["gross_profit"] = exact.get("gross_profit", 0.0)
        row["units"] = exact.get("units", 0)
        if row["cogs"] > 0 and row["units"] > 0:
            row["unit_cost"] = row["cogs"] / row["units"]
        row["orders_count"] = exact.get("orders", 0)
        row["shopifyql_exact"] = "YES"
        exact_rows.append(row)
    return exact_rows


def build_cost_zero_from_lines(all_line_rows):
    """Build Cost Zero strictly from line items whose current unit COGS is 0.

    This avoids pulling products whose product-level ShopifyQL COGS looks like 0
    but whose actual variant Cost per item is already loaded. The Cost Zero tab
    should only show rows that need COGS cleanup.
    """
    grouped = defaultdict(lambda: {
        "product_title":"", "vendor":"", "product_type":"",
        "gross_sales":0.0, "discounts":0.0, "net_sales":0.0,
        "cogs":0.0, "gross_profit":0.0, "units":0, "orders":set(),
    })
    detail = []
    for d in all_line_rows:
        unit_cost = money(d.get("unit_cost"))
        gross = money(d.get("gross_sales"))
        units = row_int(d, "units")
        if unit_cost > 0.0001 or gross <= 0.0001 or units <= 0:
            continue
        row = dict(d)
        # Force the detail to remain audit-clean: Cost Zero means COGS really 0.
        row["unit_cost"] = 0.0
        row["cogs"] = 0.0
        row["gross_profit"] = money(row.get("net_paid"))
        detail.append(row)

        key = (row.get("product_title") or "Unknown").strip().lower()
        p = grouped[key]
        p["product_title"] = row.get("product_title") or "Unknown"
        p["vendor"] = p["vendor"] or row.get("vendor") or "Unknown"
        p["product_type"] = p["product_type"] or row.get("product_type") or "Uncategorized"
        p["gross_sales"] += gross
        p["discounts"] += money(row.get("discount"))
        p["net_sales"] += money(row.get("net_paid"))
        p["cogs"] += 0.0
        p["gross_profit"] += money(row.get("net_paid"))
        p["units"] += units
        if row.get("order_name"):
            p["orders"].add(row.get("order_name"))

    return list(grouped.values()), detail

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
    print(f"  CORRO QUARTERLY REPORT v2.5 — {label}")
    print(f"  {start} → {end} | Months: {', '.join(mnames)}")
    print(f"{'='*60}\n")

    vmap, all_pids = fetch_product_map()
    orders         = fetch_orders(start, end)
    order_staff_map= fetch_order_staff_map([o.get("id") for o in orders])
    staff_ql_map   = fetch_shopifyql_staff_orders(start, end)
    sq_map         = fetch_shopifyql(start, end)
    sq_sku_map     = fetch_shopifyql_sku(start, end)
    sku_agg, disc_zero_rows, staff_rows, all_line_rows = aggregate(orders, vmap, months, order_staff_map, staff_ql_map)
    sku_exact_rows = build_exact_sku_rows(sku_agg, sq_sku_map, vmap, months)
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

    monthly_gs   = {m:sum(p["monthly"][m]["gross_sales"]  for p in prod_agg.values()) for m in months}
    monthly_disc = {m:sum(p["monthly"][m]["discounts"]    for p in prod_agg.values()) for m in months}
    monthly_ret  = {m:sum(p["monthly"][m]["returns"]      for p in prod_agg.values()) for m in months}
    monthly_ns   = {m:sum(p["monthly"][m]["net_sales"]    for p in prod_agg.values()) for m in months}
    monthly_cogs = {m:sum(p["monthly"][m]["cogs"]         for p in prod_agg.values()) for m in months}
    monthly_gp   = {m:sum(p["monthly"][m]["gross_profit"] for p in prod_agg.values()) for m in months}
    monthly_u    = {m:sum(p["monthly"][m]["units"]        for p in prod_agg.values()) for m in months}
    monthly_margin = {m:(round(monthly_gp[m]/monthly_ns[m]*100,1) if monthly_ns[m] else 0) for m in months}

    # ── 1. SUMMARY ────────────────────────────────────────────────
    h = ["updated_at","quarter","metric",mnames[0],mnames[1],mnames[2],"Q Total"]

    rows_s = [
        [now,label,"Gross Sales",  round(monthly_gs[months[0]],2),round(monthly_gs[months[1]],2),round(monthly_gs[months[2]],2),round(total_gs,2)],
        [now,label,"Discounts",    round(monthly_disc[months[0]],2),round(monthly_disc[months[1]],2),round(monthly_disc[months[2]],2),round(total_disc,2)],
        [now,label,"Returns",      round(monthly_ret[months[0]],2),round(monthly_ret[months[1]],2),round(monthly_ret[months[2]],2),round(total_ret,2)],
        [now,label,"Net Sales",    round(monthly_ns[months[0]],2),round(monthly_ns[months[1]],2),round(monthly_ns[months[2]],2),round(total_ns,2)],
        [now,label,"COGS",         round(monthly_cogs[months[0]],2),round(monthly_cogs[months[1]],2),round(monthly_cogs[months[2]],2),round(total_cogs,2)],
        [now,label,"Gross Profit", round(monthly_gp[months[0]],2),round(monthly_gp[months[1]],2),round(monthly_gp[months[2]],2),round(total_gp,2)],
        [now,label,"Gross Margin", f"{monthly_margin[months[0]]}%",f"{monthly_margin[months[1]]}%",f"{monthly_margin[months[2]]}%",f"{gp_margin}%"],
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
            round(margin,4), total_sales, p.get("orders_count") or len(p["orders"]),
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
            round(margin,3), round(r["gross_sales"],2), r.get("orders_count") or len(r["orders"]),
            round(r["monthly"][months[0]]["net_sales"],2),
            round(r["monthly"][months[1]]["net_sales"],2),
            round(r["monthly"][months[2]]["net_sales"],2),
        ]

    h_sku = ["updated_at","quarter","rank","product_title","sku","vendor","product_type",
             "is_dropship","units","gross_sales","discounts","returns","net_sales",
             "cogs","gross_profit","gross_margin","total_sales","orders",
             f"{mnames[0]} net_sales",f"{mnames[1]} net_sales",f"{mnames[2]} net_sales"]

    # ── 2. TOP 100 BY SKU ─────────────────────────────────────────
    sorted_skus = sorted(sku_exact_rows, key=lambda x: x["gross_profit"], reverse=True)
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

    # ── 7. COST ZERO (strict: actual line unit_cost / COGS = 0) ──
    cost_zero, cost_zero_detail_rows = build_cost_zero_from_lines(all_line_rows)
    cost_zero = sorted(cost_zero, key=lambda x: x["gross_sales"], reverse=True)
    h_cz = ["updated_at","quarter","rank","product_title","vendor","product_type",
            "units","gross_sales","discounts","net_sales","cogs","gross_profit",
            "gross_margin","orders"]
    write_tab(sh, f"q_cost_zero_{sfx}", h_cz,
              [[now,label,i+1,p["product_title"],p.get("vendor","") or "Unknown",p.get("product_type",""),
                p["units"],round(p["gross_sales"],2),round(p["discounts"],2),round(p["net_sales"],2),
                0,round(p["gross_profit"],2),
                round((p["gross_profit"]/p["net_sales"]),4) if p["net_sales"] else 0,
                len(p["orders"])]
               for i,p in enumerate(cost_zero)])

    # Order-by-order detail for Cost Zero products.
    # Only lines with unit_cost = 0 are written here. If a product has mixed
    # variants, loaded-cost variants are excluded from Cost Zero.
    h_cz_det = ["updated_at","quarter","order_name","created_at","month","category",
                "customer_name","customer_email","tags","product_title","sku","vendor","product_type",
                "units","gross_sales","discount","discount_pct","net_paid","unit_cost","cogs",
                "gross_profit","shipping_paid","financial_status","source_name","order_id",
                "order_staff_name","order_staff_email"]
    write_tab(sh, f"q_cost_zero_detail_{sfx}", h_cz_det,
              [[now,label]+[d.get(k,"") for k in ["order_name","created_at","month","category",
               "customer_name","customer_email","tags","product_title","sku","vendor","product_type",
               "units","gross_sales","discount","discount_pct","net_paid","unit_cost","cogs",
               "gross_profit","shipping_paid","financial_status","source_name","order_id",
               "order_staff_name","order_staff_email"]]
               for d in sorted(cost_zero_detail_rows, key=lambda x: (x.get("product_title", ""), x.get("created_at", "")))])

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
    h_staff = ["updated_at","quarter","order_name","created_at","month",
               "staff_member","staff_person_name","staff_person_email",
               "customer_name","customer_email",
               "order_staff_name","order_staff_email","source_name",
               "shopifyql_order_id","shopifyql_total_price","shopifyql_discount_amount",
               "category","tags","product_title","sku","vendor","product_type","is_dropship",
               "units","gross_sales","discount","discount_pct","net_paid","unit_cost","cogs",
               "expected_item_payment","payment_gap","gross_profit","shipping_paid"]
    write_tab(sh, f"q_staff_orders_{sfx}", h_staff,
              [[now,label]+[d.get(k,"") for k in ["order_name","created_at","month",
               "staff_member","staff_person_name","staff_person_email",
               "customer_name","customer_email",
               "order_staff_name","order_staff_email","source_name",
               "shopifyql_order_id","shopifyql_total_price","shopifyql_discount_amount",
               "category","tags","product_title","sku","vendor","product_type","is_dropship",
               "units","gross_sales","discount","discount_pct","net_paid","unit_cost","cogs",
               "expected_item_payment","payment_gap","gross_profit","shipping_paid"]]
               for d in sorted(staff_rows, key=lambda x: x["created_at"])])

    # ── 9. MONTHLY BREAKDOWN ──────────────────────────────────────
    h_mb = ["updated_at","quarter","rank","product_title","sku","vendor",
            f"{mnames[0]} units",f"{mnames[0]} sales",f"{mnames[0]} margin",
            f"{mnames[1]} units",f"{mnames[1]} sales",f"{mnames[1]} margin",
            f"{mnames[2]} units",f"{mnames[2]} sales",f"{mnames[2]} margin",
            "Q total units","Q total sales","Q avg margin"]
    mb_rows = []
    for i, r in enumerate(sorted_skus[:100]):
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
              "q_top_margin","q_top_gp","q_cost_zero","q_cost_zero_detail",
              "q_discount_zero","q_discount_zero_detail","q_staff_orders",
              "q_monthly_breakdown","q_zero_sales","q_vendors"]:
        print(f"    {t}_{sfx}")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()
