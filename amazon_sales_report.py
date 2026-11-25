"""
Amazon Sales & Traffic Report -> .xlsx  (same data as Seller Central > Reports > Business Reports)

Uses the same SP-API .env credentials as amazon_orders_report.py.
Run:  python amazon_sales_report.py 2026-11-01 2026-11-25            (daily, default)
      python amazon_sales_report.py 2026-01-01 2026-11-25 week       (weekly)
      python amazon_sales_report.py 2026-01-01 2026-11-25 month      (monthly)

Sheets:
  By Date - per day/week/month: ordered sales, units, order items, B2B, refunds,
            shipped sales, sessions, page views, conversion %, Buy Box %, + TOTAL row
  By SKU  - per SKU/ASIN for the whole range: sales, units, sessions, page views,
            conversion %, Buy Box %, sorted by sales
"""
import json
import sys

from openpyxl import Workbook

from amazon_inventory_report import fetch_report
from amazon_orders_report import get_token

REPORT_TYPE = "GET_SALES_AND_TRAFFIC_REPORT"
GRANULARITY = {"day": "DAY", "week": "WEEK", "month": "MONTH"}
NOT_ADDITIVE = ("average", "Percentage", "Rate")  # don't sum these in the TOTAL row


def flatten(d, out=None):
    """{'salesByDate': {'orderedProductSales': {'amount': 5, 'currencyCode': 'INR'}}}
    -> {'orderedProductSales': 5}"""
    out = {} if out is None else out
    for k, v in d.items():
        if isinstance(v, dict) and "amount" in v:
            out[k] = v["amount"]
        elif isinstance(v, dict):
            flatten(v, out)
        else:
            out[k] = v
    return out


def columns(rows, first):
    cols = list(first)
    for r in rows:
        cols += [k for k in r if k not in cols]
    return cols


def main():
    if len(sys.argv) not in (3, 4):
        sys.exit("Usage: python amazon_sales_report.py YYYY-MM-DD YYYY-MM-DD [day|week|month]")
    date_from, date_to = sys.argv[1], sys.argv[2]
    gran = sys.argv[3].lower() if len(sys.argv) == 4 else "day"
    if gran not in GRANULARITY:
        sys.exit("Granularity must be day, week or month")

    token = get_token()
    text = fetch_report(token, REPORT_TYPE, date_from, date_to,
                        {"dateGranularity": GRANULARITY[gran], "asinGranularity": "SKU"})
    data = json.loads(text)

    by_date = [flatten(x) for x in data.get("salesAndTrafficByDate", [])]
    by_sku = [flatten(x) for x in data.get("salesAndTrafficByAsin", [])]
    print(f"{len(by_date)} {gran} rows, {len(by_sku)} SKU rows")

    wb = Workbook()

    ws = wb.active
    ws.title = "By Date"
    cols = columns(by_date, ["date"])
    ws.append(cols)
    by_date.sort(key=lambda r: r.get("date") or "")
    for r in by_date:
        ws.append([r.get(c) for c in cols])
    if by_date:
        total = ["TOTAL"]
        for c in cols[1:]:
            vals = [r.get(c) for r in by_date if isinstance(r.get(c), (int, float))]
            additive = not any(w in c for w in NOT_ADDITIVE)
            total.append(round(sum(vals), 2) if additive and vals else None)
        ws.append(total)
    ws.freeze_panes = "B2"

    ws = wb.create_sheet("By SKU")
    cols = columns(by_sku, ["sku", "childAsin", "parentAsin"])
    ws.append(cols)
    by_sku.sort(key=lambda r: r.get("orderedProductSales") or 0, reverse=True)
    for r in by_sku:
        ws.append([r.get(c) for c in cols])
    ws.freeze_panes = "B2"

    fmt = lambda d: f"{d[8:10]}-{d[5:7]}-{d[2:4]}"
    out = f"Amazon_Sales-{fmt(date_from)}_to_{fmt(date_to)}.xlsx"
    wb.save(out)
    print(f"Saved -> {out}")


if __name__ == "__main__":
    main()
