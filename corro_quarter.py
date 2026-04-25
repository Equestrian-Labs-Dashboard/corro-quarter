"""
CORRO QUARTERLY REPORT — v2.0
==============================
Genera el reporte trimestral de Corro desde Shopify API.
Replica exactamente las hojas del Excel Q1_2026 + hoja nueva de Discount Zero.

Hojas que genera (sufijo _{q}_{y} en cada tab):
  q_summary           — KPIs totales + monthly breakdown
  q_top_sku           — Top 100 por SKU individual (con Jan/Feb/Mar)
  q_top_products      — Top 100 agrupado por producto
  q_top_dropship      — Top 100 solo dropship
  q_top_margin        — Top 100 por Gross Margin % (min 5 unidades)
  q_top_gp            — Top 100 por Gross Profit $
  q_cost_zero         — Productos con COGS = 0 (sin costo cargado en Shopify)
  q_discount_zero     — Productos donde el cliente pagó $0 (descuento 100%)
                        por categoría: Team Rider, Influencer, Marketing, etc.
  q_monthly_breakdown — Desglose mensual por producto (Jan/Feb/Mar)
  q_zero_sales        — Productos sin actividad en el período
  q_vendors           — Pareto por vendor

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

# ── Discount-zero tag categories (Ceci request) ───────────────────
DISC_ZERO_CATS = [
    ("Advent Calendar",  ["advent calendar","advent_calendar"]),
    ("Team Rider",       ["team rider","team_rider"]),
    ("Influencer",       ["influencer"]),
    ("Marketing",        ["marketing","marketing - sponsorship"]),
    ("Sponsorship",      ["sponsorship","sponsor"]),
    ("Internal / Staff", ["staff","internal","employee"]),
]
def classify_disc_zero(tags_str):
    t = (tags_str or "").lower()
    for cat, kws in DISC_ZERO_CATS:
        if any(k in t for k in kws):
            return cat
    return "Other"

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
            "fields":"id,name,created_at,financial_status,subtotal_price,"
                     "total_discounts,discount_codes,source_name,tags,line_items,customer",
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
            vmap[vid] = {
                "product_id":   pid,
                "product_title":prod.get("title",""),
                "sku":          sku,
                "vendor":       prod.get("vendor",""),
                "product_type": prod.get("productType") or "Uncategorized",
                "unit_cost":    float(uc.get("amount") or 0),
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
    """Returns sku_agg (by variant_id) and disc_zero_rows (customer paid $0)."""
    sku_agg       = {}
    disc_zero_rows = []  # customer paid $0 — Ceci's new sheet

    for order in orders:
        om   = int((order.get("created_at","")[:7]).split("-")[-1] or 0)
        disc_total  = float(order.get("total_discounts",0) or 0)
        li_count    = len(order.get("line_items",[])) or 1
        disc_per_li = disc_total / li_count
        is_refund   = (order.get("financial_status") or "").startswith("refund")
        tags_str    = order.get("tags") or ""

        for li in order.get("line_items",[]):
            vid  = str(li.get("variant_id") or "")
            info = vmap.get(vid, {
                "product_id":str(li.get("product_id","")),
                "product_title":li.get("title","Unknown"),
                "sku":li.get("sku",""),"vendor":"",
                "product_type":"Uncategorized","unit_cost":0.0,
            })
            qty   = int(li.get("quantity",0) or 0)
            price = float(li.get("price",0) or 0) * qty
            disc  = disc_per_li
            net   = max(price - disc, 0)
            cost  = info["unit_cost"] * qty
            ret   = -net if is_refund else 0.0

            key = vid or f"nv_{info['product_id']}_{info['product_title']}"
            if key not in sku_agg:
                sku_agg[key] = mk_sku_row(info, months)
            r = sku_agg[key]
            r["gross_sales"]  += price
            r["discounts"]    += disc
            r["returns"]      += ret
            r["net_sales"]    += net
            r["cogs"]         += cost
            r["gross_profit"] += max(net - cost, 0)
            r["units"]        += qty
            r["orders"].add(order["id"])
            if om in months:
                r["monthly"][om]["gross_sales"] += price
                r["monthly"][om]["net_sales"]   += net
                r["monthly"][om]["units"]       += qty

            # Discount zero: customer paid $0 (price=0 OR net≈0)
            if price == 0 or net < 0.01:
                disc_zero_rows.append({
                    "order_name":   order.get("name",""),
                    "created_at":   order.get("created_at","")[:10],
                    "month":        MONTH_NAMES.get(om,""),
                    "category":     classify_disc_zero(tags_str),
                    "tags":         tags_str,
                    "product_title":info["product_title"],
                    "sku":          info["sku"],
                    "vendor":       info["vendor"],
                    "product_type": info["product_type"],
                    "units":        qty,
                    "original_price":price,
                    "discount":     disc,
                    "net_paid":     net,
                    "unit_cost":    info["unit_cost"],
                    "cogs":         cost,
                })

    return sku_agg, disc_zero_rows

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
    # Override GP with ShopifyQL if available
    for t, p in prod.items():
        sq = sq_map.get(t.lower())
        if sq:
            p["gross_profit"] = sq["gross_profit"]
            p["gross_sales"]  = sq["gross_sales"]
            p["net_sales"]    = sq["net_sales"]
            p["vendor"]       = p["vendor"] or sq["vendor"]
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
    sku_agg, disc_zero_rows = aggregate(orders, vmap, months)
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
    # Monthly GP approximated proportionally from ShopifyQL total
    # (ShopifyQL doesn't break GP by month directly)
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
        margin = round(gp/ns,4) if ns else 0  # decimal e.g. 0.19 = 19%
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
        margin = round(gp/ns,4) if ns else 0  # decimal
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
        [p for p in prod_agg.values() if p["cogs"]==0 and p["gross_sales"]>0],
        key=lambda x: x["gross_sales"], reverse=True)
    h_cz = ["updated_at","quarter","rank","product_title","vendor","product_type",
             "units","gross_sales","cogs","gross_profit","orders"]
    write_tab(sh, f"q_cost_zero_{sfx}", h_cz,
              [[now,label,i+1,p["product_title"],p.get("vendor",""),p.get("product_type",""),
                p["units"],round(p["gross_sales"],2),0,round(p["gross_profit"],2),len(p["orders"])]
               for i,p in enumerate(cost_zero)])

    # ── 8. DISCOUNT ZERO (customer paid $0) — Ceci request ───────
    # Summary by category × month
    cat_month = defaultdict(lambda: {m:{"units":0,"cogs":0.0,"products":set()} for m in months})
    for d in disc_zero_rows:
        om = next((m for m in months if MONTH_NAMES[m]==d["month"]), None)
        if om:
            cat_month[d["category"]][om]["units"]   += d["units"]
            cat_month[d["category"]][om]["cogs"]    += d["cogs"]
            cat_month[d["category"]][om]["products"].add(d["product_title"])

    all_cats = [c for c,_ in DISC_ZERO_CATS] + ["Other"]
    h_dz = ["updated_at","quarter","category",
            f"{mnames[0]} units",f"{mnames[0]} cogs",
            f"{mnames[1]} units",f"{mnames[1]} cogs",
            f"{mnames[2]} units",f"{mnames[2]} cogs",
            "total units","total cogs","sample products"]
    dz_rows = []
    for cat in all_cats:
        if cat not in cat_month: continue
        cm = cat_month[cat]
        all_prods = set()
        for m in months: all_prods |= cm[m]["products"]
        dz_rows.append([
            now, label, cat,
            cm[months[0]]["units"], round(cm[months[0]]["cogs"],2),
            cm[months[1]]["units"], round(cm[months[1]]["cogs"],2),
            cm[months[2]]["units"], round(cm[months[2]]["cogs"],2),
            sum(cm[m]["units"] for m in months),
            round(sum(cm[m]["cogs"] for m in months),2),
            ", ".join(sorted(all_prods)[:5]),
        ])
    write_tab(sh, f"q_discount_zero_{sfx}", h_dz, dz_rows)

    # Detail
    h_dz_det = ["updated_at","quarter","order_name","created_at","month","category",
                "tags","product_title","sku","vendor","product_type",
                "units","original_price","discount","net_paid","unit_cost","cogs"]
    write_tab(sh, f"q_discount_zero_detail_{sfx}", h_dz_det,
              [[now,label]+[d[k] for k in ["order_name","created_at","month","category",
               "tags","product_title","sku","vendor","product_type",
               "units","original_price","discount","net_paid","unit_cost","cogs"]]
               for d in sorted(disc_zero_rows, key=lambda x: x["created_at"])])

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
            mg = r["gross_profit"] * (ms/ns) if ns else 0  # approx monthly GP
            row_data += [r["monthly"][m]["units"], round(ms,2), round(mg/ms if ms else 0,3)]
        row_data += [r["units"], round(r["gross_sales"],2), avg_m]
        mb_rows.append(row_data)
    write_tab(sh, f"q_monthly_breakdown_{sfx}", h_mb, mb_rows)

    # ── 10. ZERO SALES ────────────────────────────────────────────
    active_pids = {r["product_id"] for r in sku_agg.values()}
    zero_sales  = [v for pid,v in vmap.items()
                   if v["product_id"] not in active_pids and v["product_title"]]
    # Deduplicate by product_title
    seen_titles = set(); zs_dedup = []
    for v in sorted(zero_sales, key=lambda x: x["product_title"]):
        if v["product_title"] not in seen_titles:
            seen_titles.add(v["product_title"]); zs_dedup.append(v)
    h_zs = ["updated_at","quarter","rank","product_title","sku","vendor","product_type","is_dropship"]
    write_tab(sh, f"q_zero_sales_{sfx}", h_zs,
              [[now,label,i+1,v["product_title"],v["sku"],v["vendor"],v["product_type"],
                "YES" if v["unit_cost"]>0 else "NO"]
               for i,v in enumerate(zs_dedup)])

    # ── 11. VENDORS PARETO ────────────────────────────────────────
    sorted_vend = sorted(vendor_agg.values(), key=lambda x: x["gross_sales"], reverse=True)
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
    print(f"  ✓ DONE — {label}")
    print(f"  Net Sales: ${total_ns:,.2f} | GP: ${total_gp:,.2f} ({gp_margin}%)")
    print(f"  Sheets written (all suffixed _{sfx}):")
    for t in ["q_summary","q_top_sku","q_top_products","q_top_dropship",
              "q_top_margin","q_top_gp","q_cost_zero","q_discount_zero",
              "q_discount_zero_detail","q_monthly_breakdown","q_zero_sales","q_vendors"]:
        print(f"    {t}_{sfx}")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()
