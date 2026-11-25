"""
Amazon Payments Report -> .xlsx  (SP-API Finances API - same data as Seller Central > Payments)

Unlike the settlement report, this also includes the current OPEN period (money earned
but not yet paid out), and it is live - no waiting for Amazon to generate a file.

Uses the same SP-API .env credentials as amazon_orders_report.py.
Run:  python amazon_payments_report.py                          (payment periods started in last 90 days)
      python amazon_payments_report.py 2026-10-01 2026-11-25    (payment periods started in range)

Sheets:
  Payments     - one row per payment period (payout): status, period, transfer date,
                 Amazon's total, and the sum of its transactions
  Transactions - every money line in each period: order charges, fees, refunds,
                 promotions, TCS/TDS withheld, ads charges, adjustments, ...
"""
import sys
from datetime import date, datetime, timedelta, timezone
from urllib.parse import quote, urlencode

from openpyxl import Workbook

from amazon_orders_report import get_token, sp

IST = timezone(timedelta(hours=5, minutes=30))

# Scalar fields copied onto every money line found beneath them
CONTEXT = {
    "PostedDate": "posted-date", "postedDate": "posted-date",
    "AmazonOrderId": "order-id", "SellerOrderId": "seller-order-id",
    "MarketplaceName": "marketplace", "SellerSKU": "sku", "ASIN": "asin",
    "QuantityShipped": "quantity", "Quantity": "quantity", "quantity": "quantity",
    "FeeDescription": "description", "productDescription": "description",
    "AdjustmentType": "category", "FeeReason": "category", "transactionType": "category",
    "TransactionType": "category",
}
# Field that names the money amount next to it (e.g. ChargeType=Principal + ChargeAmount)
LABEL_KEYS = ("ChargeType", "FeeType", "PromotionType")
# Money fields that are not cash movements or duplicate other lines
SKIP_FIELDS = {"PerUnitAmount", "transactionValue", "BaseAmount", "OriginalAmount",
               "TotalExpense", "ChargeInstrumentList"}
# Event totals to skip when an item breakdown is present (avoid counting twice)
PARENT_TOTALS = {
    "AdjustmentAmount": "AdjustmentItemList",
    "ReimbursedAmount": "SAFETReimbursementItemList",
    "RecoveryAmount": "DebtRecoveryItemList",
    "TotalAmount": "FeeComponent",
}
TXN_COLS = ["payment-id", "event-type", "posted-date", "order-id", "seller-order-id", "sku",
            "asin", "quantity", "category", "component", "amount-type", "description",
            "currency", "amount", "marketplace"]


def to_ist(s):
    """'2026-11-01T10:00:00Z' -> naive IST datetime so Excel shows it as a date."""
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(IST).replace(tzinfo=None)
    except ValueError:
        return s


def money(m):
    return isinstance(m, dict) and "CurrencyAmount" in m


def paged(token, path, params):
    """Follow NextToken pages of a Finances v0 endpoint; yields each payload."""
    while True:
        res = sp(token, "GET", f"{path}?{urlencode(params)}")["payload"]
        yield res
        if not res.get("NextToken"):
            return
        params = {**params, "NextToken": res["NextToken"]}


def walk(node, ctx, component, out):
    """Flatten any financial event into one row per money amount."""
    if isinstance(node, list):
        for x in node:
            walk(x, ctx, component, out)
        return
    if not isinstance(node, dict):
        return
    ctx = dict(ctx)
    for k, v in node.items():
        if k in CONTEXT and isinstance(v, (str, int, float)):
            ctx[CONTEXT[k]] = v
    label = next((node[k] for k in LABEL_KEYS if isinstance(node.get(k), str)), None)
    for k, v in node.items():
        if k in SKIP_FIELDS or (k in PARENT_TOTALS and node.get(PARENT_TOTALS[k])):
            continue
        if money(v):
            out.append({**ctx, "component": component, "amount-type": label or k,
                        "currency": v.get("CurrencyCode"), "amount": v.get("CurrencyAmount")})
        elif isinstance(v, (dict, list)):
            walk(v, ctx, k[:-4] if k.endswith("List") else k, out)


def main():
    if len(sys.argv) == 3:
        date_from, date_to = sys.argv[1], sys.argv[2]
    elif len(sys.argv) == 1:
        date_to = date.today().isoformat()
        date_from = (date.today() - timedelta(days=89)).isoformat()
    else:
        sys.exit("Usage: python amazon_payments_report.py [YYYY-MM-DD YYYY-MM-DD]")

    token = get_token()

    # 1. Payment periods (financial event groups)
    params = {"MaxResultsPerPage": 100,
              "FinancialEventGroupStartedAfter": f"{date_from}T00:00:00+05:30"}
    if date_to < date.today().isoformat():  # "before" must not be in the future
        params["FinancialEventGroupStartedBefore"] = f"{date_to}T23:59:59+05:30"
    groups = []
    for page in paged(token, "/finances/v0/financialEventGroups", params):
        groups += page.get("FinancialEventGroupList", [])
    groups.sort(key=lambda g: g.get("FinancialEventGroupStart", ""))
    print(f"Found {len(groups)} payment period(s) started {date_from} to {date_to}")
    if not groups:
        sys.exit("No payment periods in this range.")

    # 2. Transactions in each period
    txns, sums = [], {}
    for g in groups:
        gid = g["FinancialEventGroupId"]
        rows = []
        path = f"/finances/v0/financialEventGroups/{quote(gid, safe='')}/financialEvents"
        for page in paged(token, path, {"MaxResultsPerPage": 100}):
            for list_name, events in (page.get("FinancialEvents") or {}).items():
                et = list_name.removesuffix("EventList")
                walk(events, {"payment-id": gid, "event-type": et}, et, rows)
        sums[gid] = round(sum(float(r["amount"] or 0) for r in rows), 2)
        total = (g.get("OriginalTotal") or {}).get("CurrencyAmount")
        print(f"{g.get('ProcessingStatus', ''):6} {gid}: {len(rows)} lines, "
              f"Amazon total {total}, lines total {sums[gid]}")
        txns += rows

    # 3. Write workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Payments"
    ws.append(["payment-id", "status", "transfer-status", "period-start", "period-end",
               "transfer-date", "currency", "amazon-total", "beginning-balance",
               "transactions-total", "account-tail", "trace-id"])
    for g in groups:
        total = g.get("OriginalTotal") or {}
        ws.append([
            g["FinancialEventGroupId"], g.get("ProcessingStatus"), g.get("FundTransferStatus"),
            to_ist(g.get("FinancialEventGroupStart")), to_ist(g.get("FinancialEventGroupEnd")),
            to_ist(g.get("FundTransferDate")), total.get("CurrencyCode"),
            total.get("CurrencyAmount"), (g.get("BeginningBalance") or {}).get("CurrencyAmount"),
            sums[g["FinancialEventGroupId"]], g.get("AccountTail"), g.get("TraceId"),
        ])
    ws.freeze_panes = "B2"

    ws = wb.create_sheet("Transactions")
    ws.append(TXN_COLS)
    txns.sort(key=lambda r: (r["payment-id"], r.get("posted-date") or ""))
    for r in txns:
        r["posted-date"] = to_ist(r.get("posted-date"))
        ws.append([r.get(c) for c in TXN_COLS])
    ws.freeze_panes = "A2"

    fmt = lambda d: f"{d[8:10]}-{d[5:7]}-{d[2:4]}"
    out = f"Amazon_Payments-{fmt(date_from)}_to_{fmt(date_to)}.xlsx"
    wb.save(out)
    print(f"Saved {len(groups)} payment periods, {len(txns)} lines -> {out}")


if __name__ == "__main__":
    main()
