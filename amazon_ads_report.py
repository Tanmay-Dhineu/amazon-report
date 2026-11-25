"""
Amazon Sponsored Products Ads Report -> .xlsx  (Amazon Ads API, reporting v3)

.env: ADS_CLIENT_ID, ADS_CLIENT_SECRET, ADS_REFRESH_TOKEN, ADS_PROFILE_ID (optional)
      These are Amazon Ads API credentials, NOT the SP-API ones.
      If ADS_PROFILE_ID is empty, the amazon.in profile is detected automatically.

Run:  python amazon_ads_report.py 2026-11-01 2026-11-25                (campaign, default)
      python amazon_ads_report.py 2026-11-01 2026-11-25 product        (per ASIN/SKU)
      python amazon_ads_report.py 2026-11-01 2026-11-25 searchterm     (customer search terms)
Max 31 days per run; Amazon keeps only the last ~95 days of ads data.
"""
import gzip
import json
import os
import re
import sys
import time
from datetime import date

import requests
from dotenv import load_dotenv
from openpyxl import Workbook

load_dotenv()

BASE = "https://advertising-api-eu.amazon.com"  # amazon.in is served by the EU endpoint
COUNTRY = "IN"

METRICS = [
    "impressions", "clicks", "cost", "purchases7d", "sales7d",
    "unitsSoldClicks7d", "clickThroughRate", "costPerClick",
]
REPORTS = {
    "campaign": {
        "reportTypeId": "spCampaigns",
        "groupBy": ["campaign"],
        "columns": ["date", "campaignName", "campaignId", "campaignStatus",
                    "campaignBudgetAmount", *METRICS],
    },
    "product": {
        "reportTypeId": "spAdvertisedProduct",
        "groupBy": ["advertiser"],
        "columns": ["date", "campaignName", "adGroupName", "advertisedAsin",
                    "advertisedSku", *METRICS],
    },
    "searchterm": {
        "reportTypeId": "spSearchTerm",
        "groupBy": ["searchTerm"],
        "columns": ["date", "campaignName", "adGroupName", "keyword", "matchType",
                    "searchTerm", *METRICS],
    },
}


def get_token():
    r = requests.post(
        "https://api.amazon.com/auth/o2/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": os.environ["ADS_REFRESH_TOKEN"],
            "client_id": os.environ["ADS_CLIENT_ID"],
            "client_secret": os.environ["ADS_CLIENT_SECRET"],
        },
        timeout=30,
    )
    j = r.json()
    if "access_token" not in j:
        raise RuntimeError(f"Token error: {j}")
    return j["access_token"]


def headers(token, profile_id=None, content_type="application/json"):
    h = {
        "Amazon-Advertising-API-ClientId": os.environ["ADS_CLIENT_ID"],
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type,
    }
    if profile_id:
        h["Amazon-Advertising-API-Scope"] = str(profile_id)
    return h


def get_profile_id(token):
    pid = os.environ.get("ADS_PROFILE_ID", "").strip()
    if pid:
        return pid
    r = requests.get(BASE + "/v2/profiles", headers=headers(token), timeout=30)
    if not r.ok:
        raise RuntimeError(f"{r.status_code} /v2/profiles: {r.text}")
    profiles = [p for p in r.json() if p.get("countryCode") == COUNTRY]
    if not profiles:
        raise RuntimeError(f"No {COUNTRY} advertising profile found: {r.text}")
    pid = profiles[0]["profileId"]
    print(f"profileId: {pid} (set ADS_PROFILE_ID={pid} in .env to skip this lookup)")
    return pid


def create_report(token, profile_id, kind, date_from, date_to):
    cfg = REPORTS[kind]
    body = {
        "name": f"SP {kind} {date_from} to {date_to}",
        "startDate": date_from,
        "endDate": date_to,
        "configuration": {
            "adProduct": "SPONSORED_PRODUCTS",
            "reportTypeId": cfg["reportTypeId"],
            "groupBy": cfg["groupBy"],
            "columns": cfg["columns"],
            "timeUnit": "DAILY",
            "format": "GZIP_JSON",
        },
    }
    r = requests.post(
        BASE + "/reporting/reports",
        headers=headers(token, profile_id, "application/vnd.createasyncreportrequest.v3+json"),
        json=body,
        timeout=60,
    )
    if r.status_code == 425:  # identical report already requested -> reuse it
        m = re.search(r"duplicate of\s*:\s*([\w-]+)", r.text)
        if m:
            return m.group(1)
    if not r.ok:
        raise RuntimeError(f"{r.status_code} create report: {r.text}")
    return r.json()["reportId"]


def wait_for_report(token, profile_id, report_id):
    while True:
        r = requests.get(BASE + f"/reporting/reports/{report_id}",
                         headers=headers(token, profile_id), timeout=60)
        if not r.ok:
            raise RuntimeError(f"{r.status_code} get report: {r.text}")
        rep = r.json()
        status = rep["status"]
        print("status:", status)
        if status == "COMPLETED":
            return rep["url"]
        if status == "FAILED":
            raise RuntimeError(f"Report FAILED: {rep.get('failureReason')}")
        time.sleep(30)


def main():
    if len(sys.argv) not in (3, 4):
        sys.exit("Usage: python amazon_ads_report.py YYYY-MM-DD YYYY-MM-DD [campaign|product|searchterm]")
    date_from, date_to = sys.argv[1], sys.argv[2]
    kind = sys.argv[3].lower() if len(sys.argv) == 4 else "campaign"
    if kind not in REPORTS:
        sys.exit(f"Report type must be one of: {', '.join(REPORTS)}")
    days = (date.fromisoformat(date_to) - date.fromisoformat(date_from)).days + 1
    if not 1 <= days <= 31:
        sys.exit("Date range must be 1-31 days")

    token = get_token()
    profile_id = get_profile_id(token)

    report_id = create_report(token, profile_id, kind, date_from, date_to)
    print("reportId:", report_id)
    url = wait_for_report(token, profile_id, report_id)

    raw = requests.get(url, timeout=120).content
    rows = json.loads(gzip.decompress(raw))

    cols = REPORTS[kind]["columns"]
    wb = Workbook()
    ws = wb.active
    ws.title = kind
    ws.append(cols + ["ACOS %", "ROAS"])
    rows.sort(key=lambda x: (x.get("date") or "", x.get("campaignName") or ""))
    for row in rows:
        cost, sales = row.get("cost") or 0, row.get("sales7d") or 0
        acos = round(cost / sales * 100, 2) if sales else None
        roas = round(sales / cost, 2) if cost else None
        ws.append([row.get(c) for c in cols] + [acos, roas])
    ws.freeze_panes = "A2"

    fmt = lambda d: f"{d[8:10]}-{d[5:7]}-{d[2:4]}"
    out = f"Amazon_Ads_{kind}-{fmt(date_from)}_to_{fmt(date_to)}.xlsx"
    wb.save(out)
    print(f"Saved {len(rows)} rows -> {out}")


if __name__ == "__main__":
    main()
