"""
Amazon All Orders Report -> .xlsx (same columns as Seller Central download)

pip install requests openpyxl python-dotenv
.env: SP_CLIENT_ID, SP_CLIENT_SECRET, SP_REFRESH_TOKEN
Run:  python amazon_orders_report.py 2026-03-01 2026-03-31   (max 30 days)
"""
import gzip
import os
import sys
import time

import requests
from dotenv import load_dotenv
from openpyxl import Workbook

load_dotenv()

BASE = "https://sellingpartnerapi-eu.amazon.com"
MARKETPLACE = "A21TJRUUN4KGV"  # amazon.in
REPORT_TYPE = "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL"

NUMERIC = {
    "quantity", "item-price", "item-tax", "shipping-price", "shipping-tax",
    "gift-wrap-price", "gift-wrap-tax", "item-promotion-discount", "ship-promotion-discount",
}
BOOLEAN = {"is-business-order", "is-iba"}


def get_token():
    r = requests.post(
        "https://api.amazon.com/auth/o2/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": os.environ["SP_REFRESH_TOKEN"],
            "client_id": os.environ["SP_CLIENT_ID"],
            "client_secret": os.environ["SP_CLIENT_SECRET"],
        },
        timeout=30,
    )
    j = r.json()
    if "access_token" not in j:
        raise RuntimeError(f"Token error: {j}")
    return j["access_token"]


def sp(token, method, path, body=None):
    r = requests.request(
        method,
        BASE + path,
        headers={"x-amz-access-token": token, "Content-Type": "application/json"},
        json=body,
        timeout=60,
    )
    if not r.ok:
        raise RuntimeError(f"{r.status_code} {path}: {r.text}")
    return r.json()


def to_value(header, v):
    if v == "":
        return None
    h = header.strip()
    if h in NUMERIC:
        try:
            n = float(v)
            return int(n) if n.is_integer() else n
        except ValueError:
            return v
    if h in BOOLEAN:
        return v.lower() == "true"
    return v


def main():
    if len(sys.argv) != 3:
        sys.exit("Usage: python amazon_orders_report.py YYYY-MM-DD YYYY-MM-DD (max 30 days)")
    date_from, date_to = sys.argv[1], sys.argv[2]

    token = get_token()

    # 1. Create report
    res = sp(token, "POST", "/reports/2021-06-30/reports", {
        "reportType": REPORT_TYPE,
        "marketplaceIds": [MARKETPLACE],
        "dataStartTime": f"{date_from}T00:00:00+05:30",
        "dataEndTime": f"{date_to}T23:59:59+05:30",
    })
    report_id = res["reportId"]
    print("reportId:", report_id)

    # 2. Poll until DONE
    while True:
        report = sp(token, "GET", f"/reports/2021-06-30/reports/{report_id}")
        status = report["processingStatus"]
        print("status:", status)
        if status == "DONE":
            break
        if status in ("CANCELLED", "FATAL"):
            raise RuntimeError(f"Report {status} (CANCELLED usually means no orders in range)")
        time.sleep(15)

    # 3. Get document URL & download (no auth header needed for this URL)
    doc = sp(token, "GET", f"/reports/2021-06-30/documents/{report['reportDocumentId']}")
    raw = requests.get(doc["url"], timeout=120).content
    if doc.get("compressionAlgorithm") == "GZIP":
        raw = gzip.decompress(raw)
    text = raw.decode("utf-8", errors="replace")

    # 4. Parse TSV
    lines = [l for l in text.splitlines() if l]
    headers = lines[0].split("\t")

    # 5. Write xlsx
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(headers)
    for line in lines[1:]:
        cells = line.split("\t")
        cells += [""] * (len(headers) - len(cells))
        ws.append([to_value(h, c) for h, c in zip(headers, cells)])

    fmt = lambda d: f"{d[8:10]}-{d[5:7]}-{d[2:4]}"
    out = f"Amazon_Order-{fmt(date_from)}_to_{fmt(date_to)}.xlsx"
    wb.save(out)
    print(f"Saved {len(lines) - 1} rows -> {out}")


if __name__ == "__main__":
    main()
