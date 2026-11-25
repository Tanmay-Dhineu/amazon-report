"""
Amazon Fees Report -> .xlsx

Uses the same SP-API .env credentials as amazon_orders_report.py.

1) Fees actually CHARGED (Finances API, by posted date, max 180 days):
   python amazon_fees_report.py 2026-11-01 2026-11-25
   Sheets:
     By Fee Type - total of each fee (commission, closing fee, shipping/FBA fees, storage, ...)
     By SKU      - per SKU: units shipped, sales (principal), each fee, total fees, fee % of sales
     Fee Lines   - every individual fee line with order id, SKU and date

2) Fee ESTIMATES for your current listings at their current price (Product Fees API):
   python amazon_fees_report.py estimate
   Sheet:
     Fee Estimates - per SKU: price, FBA/MFN, each estimated fee, total, fee % of price
"""
import sys
from collections import defaultdict
from datetime import date, timedelta
from urllib.parse import urlencode

from openpyxl import Workbook

from amazon_inventory_report import fetch_report
from amazon_orders_report import MARKETPLACE, get_token, sp
from amazon_payments_report import TXN_COLS, paged, to_ist, walk
from amazon_settlement_report import parse

NO_SKU = "(no SKU)"
fmt = lambda d: f"{d[8:10]}-{d[5:7]}-{d[2:4]}"


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def is_fee(r):
    return "Fee" in (r.get("component") or "") or "Fee" in (r.get("amount-type") or "")


def fee_name(r):
    t = r.get("amount-type") or ""
    return r.get("description") or t if t in ("FeeAmount", "TotalAmount") else t


def charged(date_from, date_to):
    days = (date.fromisoformat(date_to) - date.fromisoformat(date_from)).days + 1
    if not 1 <= days <= 180:
        sys.exit("Date range must be 1-180 days")

    token = get_token()
    params = {"MaxResultsPerPage": 100, "PostedAfter": f"{date_from}T00:00:00+05:30"}
    if date_to < date.today().isoformat():  # "before" must be at least 2 min in the past
        params["PostedBefore"] = f"{date_to}T23:59:59+05:30"

    rows, pages = [], 0
    for page in paged(token, "/finances/v0/financialEvents", params):
        for list_name, events in (page.get("FinancialEvents") or {}).items():
            et = list_name.removesuffix("EventList")
            walk(events, {"event-type": et}, et, rows)
        pages += 1
        print(f"page {pages}: {len(rows)} money lines so far")

    fees = [r for r in rows if is_fee(r)]
    print(f"{len(fees)} fee lines out of {len(rows)} money lines")
    if not fees:
        sys.exit("No fees found in this range.")

    # Per-SKU sales & units from shipped items (Principal charge lines)
    sales, units = defaultdict(float), defaultdict(int)
    for r in rows:
        if r.get("event-type") == "Shipment" and r.get("amount-type") == "Principal":
            sku = r.get("sku") or NO_SKU
            sales[sku] += num(r["amount"])
            units[sku] += int(num(r.get("quantity")))

    by_type = defaultdict(lambda: [0, 0.0])
    by_sku = defaultdict(lambda: defaultdict(float))
    for r in fees:
        name = fee_name(r)
        by_type[name][0] += 1
        by_type[name][1] += num(r["amount"])
        by_sku[r.get("sku") or NO_SKU][name] += num(r["amount"])

    wb = Workbook()
    ws = wb.active
    ws.title = "By Fee Type"
    ws.append(["fee-type", "lines", "total"])
    for name, (n, total) in sorted(by_type.items(), key=lambda kv: kv[1][1]):
        ws.append([name, n, round(total, 2)])
    ws.append(["TOTAL", sum(v[0] for v in by_type.values()),
               round(sum(v[1] for v in by_type.values()), 2)])

    ws = wb.create_sheet("By SKU")
    types = [n for n, _ in sorted(by_type.items(), key=lambda kv: kv[1][1])]
    ws.append(["sku", "units-shipped", "sales (principal)", *types, "total-fees", "fee % of sales"])
    for sku, f in sorted(by_sku.items(), key=lambda kv: sum(kv[1].values())):
        total = sum(f.values())
        pct = round(-total / sales[sku] * 100, 2) if sales[sku] else None
        ws.append([sku, units[sku] or None, round(sales[sku], 2) or None,
                   *[round(f[t], 2) if t in f else None for t in types], round(total, 2), pct])
    ws.freeze_panes = "B2"

    ws = wb.create_sheet("Fee Lines")
    cols = [c for c in TXN_COLS if c != "payment-id"]
    ws.append(cols)
    for r in sorted(fees, key=lambda r: r.get("posted-date") or ""):
        r["posted-date"] = to_ist(r.get("posted-date"))
        ws.append([r.get(c) for c in cols])
    ws.freeze_panes = "A2"

    out = f"Amazon_Fees-{fmt(date_from)}_to_{fmt(date_to)}.xlsx"
    wb.save(out)
    print(f"Saved -> {out}")


def estimate():
    token = get_token()
    headers, listings = parse(fetch_report(token, "GET_MERCHANT_LISTINGS_ALL_DATA"))
    listings = [l for l in listings if l.get("seller-sku") and num(l.get("price")) > 0]
    print(f"{len(listings)} listings with a price")
    if not listings:
        sys.exit("No priced listings found.")

    results = {}
    for i in range(0, len(listings), 20):  # API takes max 20 per call
        batch = listings[i:i + 20]
        body = [{
            "IdType": "SellerSKU",
            "IdValue": l["seller-sku"],
            "FeesEstimateRequest": {
                "MarketplaceId": MARKETPLACE,
                "IsAmazonFulfilled": l.get("fulfillment-channel", "DEFAULT") != "DEFAULT",
                "PriceToEstimateFees": {
                    "ListingPrice": {"CurrencyCode": "INR", "Amount": num(l["price"])}},
                "Identifier": l["seller-sku"],
            },
        } for l in batch]
        for res in sp(token, "POST", "/products/fees/v0/feesEstimate", body):
            results[res["FeesEstimateIdentifier"]["IdValue"]] = res
        print(f"estimated {min(i + 20, len(listings))}/{len(listings)}")

    def fee_map(res):
        details = ((res.get("FeesEstimate") or {}).get("FeeDetailList")) or []
        return {d["FeeType"]: (d.get("FinalFee") or {}).get("Amount") for d in details}

    types = sorted({t for res in results.values() for t in fee_map(res)})
    wb = Workbook()
    ws = wb.active
    ws.title = "Fee Estimates"
    ws.append(["sku", "asin", "item-name", "channel", "status", "price", *types,
               "total-fees", "fee % of price", "error"])
    for l in listings:
        res = results.get(l["seller-sku"], {})
        fm = fee_map(res)
        total = ((res.get("FeesEstimate") or {}).get("TotalFeesEstimate") or {}).get("Amount")
        price = num(l["price"])
        ws.append([
            l["seller-sku"], l.get("asin1"), l.get("item-name"),
            "MFN" if l.get("fulfillment-channel", "DEFAULT") == "DEFAULT" else "FBA",
            l.get("status"), price, *[fm.get(t) for t in types], total,
            round(total / price * 100, 2) if total is not None else None,
            (res.get("Error") or {}).get("Message"),
        ])
    ws.freeze_panes = "B2"

    out = f"Amazon_Fee_Estimates-{fmt(date.today().isoformat())}.xlsx"
    wb.save(out)
    print(f"Saved -> {out}")


def main():
    if len(sys.argv) == 2 and sys.argv[1].lower() == "estimate":
        estimate()
    elif len(sys.argv) == 3:
        charged(sys.argv[1], sys.argv[2])
    else:
        sys.exit("Usage: python amazon_fees_report.py YYYY-MM-DD YYYY-MM-DD\n"
                 "       python amazon_fees_report.py estimate")


if __name__ == "__main__":
    main()
