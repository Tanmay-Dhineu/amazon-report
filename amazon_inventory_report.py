"""
Amazon Current Inventory Report -> .xlsx

Uses the same .env credentials as amazon_orders_report.py.
Run:  python amazon_inventory_report.py          (both FBA + listings)
      python amazon_inventory_report.py fba      (FBA inventory only)
      python amazon_inventory_report.py mfn      (all listings / merchant-fulfilled qty only)

Sheets:
  FBA      - GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA (Manage FBA Inventory: fulfillable,
             reserved, unsellable, inbound quantities per SKU)
  Listings - GET_MERCHANT_LISTINGS_ALL_DATA (every listing with price, quantity, status)
"""
import gzip
import sys
import time
from datetime import date

import requests
from openpyxl import Workbook

from amazon_orders_report import MARKETPLACE, get_token, sp

REPORTS = {
    "fba": ("FBA", "GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA"),
    "mfn": ("Listings", "GET_MERCHANT_LISTINGS_ALL_DATA"),
}
NUMERIC_HINTS = ("quantity", "price", "volume")


def fetch_report(token, report_type):
    res = sp(token, "POST", "/reports/2021-06-30/reports", {
        "reportType": report_type,
        "marketplaceIds": [MARKETPLACE],
    })
    report_id = res["reportId"]
    print(f"[{report_type}] reportId: {report_id}")

    while True:
        report = sp(token, "GET", f"/reports/2021-06-30/reports/{report_id}")
        status = report["processingStatus"]
        print(f"[{report_type}] status: {status}")
        if status == "DONE":
            break
        if status in ("CANCELLED", "FATAL"):
            raise RuntimeError(f"{report_type} report {status}")
        time.sleep(15)

    doc = sp(token, "GET", f"/reports/2021-06-30/documents/{report['reportDocumentId']}")
    raw = requests.get(doc["url"], timeout=120).content
    if doc.get("compressionAlgorithm") == "GZIP":
        raw = gzip.decompress(raw)
    return raw.decode("utf-8", errors="replace")


def to_value(header, v):
    if v == "":
        return None
    if any(k in header.strip().lower() for k in NUMERIC_HINTS):
        try:
            n = float(v)
            return int(n) if n.is_integer() else n
        except ValueError:
            return v
    return v


def write_sheet(wb, title, text):
    lines = [l for l in text.splitlines() if l]
    headers = lines[0].split("\t")
    ws = wb.create_sheet(title)
    ws.append(headers)
    for line in lines[1:]:
        cells = line.split("\t")
        cells += [""] * (len(headers) - len(cells))
        ws.append([to_value(h, c) for h, c in zip(headers, cells)])
    ws.freeze_panes = "A2"
    return len(lines) - 1


def main():
    which = sys.argv[1].lower() if len(sys.argv) > 1 else "all"
    if which not in ("all", *REPORTS):
        sys.exit("Usage: python amazon_inventory_report.py [all|fba|mfn]")
    keys = list(REPORTS) if which == "all" else [which]

    token = get_token()
    wb = Workbook()
    wb.remove(wb.active)

    for key in keys:
        title, report_type = REPORTS[key]
        try:
            text = fetch_report(token, report_type)
        except RuntimeError as e:
            # FATAL is normal for FBA if the account has no FBA inventory
            print(f"Skipping {title}: {e}")
            continue
        rows = write_sheet(wb, title, text)
        print(f"{title}: {rows} rows")

    if not wb.sheetnames:
        sys.exit("No inventory data returned.")

    out = f"Amazon_Inventory-{date.today():%d-%m-%y}.xlsx"
    wb.save(out)
    print(f"Saved -> {out}")


if __name__ == "__main__":
    main()
