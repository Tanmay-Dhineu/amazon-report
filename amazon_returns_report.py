"""
Amazon Returns Report -> .xlsx

Uses the same SP-API .env credentials as amazon_orders_report.py.
Run:  python amazon_returns_report.py 2026-11-01 2026-11-25          (both FBA + MFN)
      python amazon_returns_report.py 2026-11-01 2026-11-25 fba      (FBA returns only)
      python amazon_returns_report.py 2026-11-01 2026-11-25 mfn      (self-ship / Easy Ship only)
Max 60 days per run. Dates are return dates (not order dates).

Sheets:
  FBA Returns - GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA (returns received at Amazon
                warehouses: reason, disposition, fulfillment center, customer comments)
  MFN Returns - GET_FLAT_FILE_RETURNS_DATA_BY_RETURN_DATE (returns of orders you shipped:
                return status, RMA, tracking, reason, refund/order amount)
"""
import sys
from datetime import date

from openpyxl import Workbook

from amazon_inventory_report import fetch_report, write_sheet
from amazon_orders_report import get_token

REPORTS = {
    "fba": ("FBA Returns", "GET_FBA_FULFILLMENT_CUSTOMER_RETURNS_DATA"),
    "mfn": ("MFN Returns", "GET_FLAT_FILE_RETURNS_DATA_BY_RETURN_DATE"),
}


def main():
    if len(sys.argv) not in (3, 4):
        sys.exit("Usage: python amazon_returns_report.py YYYY-MM-DD YYYY-MM-DD [all|fba|mfn]")
    date_from, date_to = sys.argv[1], sys.argv[2]
    which = sys.argv[3].lower() if len(sys.argv) == 4 else "all"
    if which not in ("all", *REPORTS):
        sys.exit("Report type must be one of: all, fba, mfn")
    days = (date.fromisoformat(date_to) - date.fromisoformat(date_from)).days + 1
    if not 1 <= days <= 60:
        sys.exit("Date range must be 1-60 days")
    keys = list(REPORTS) if which == "all" else [which]

    token = get_token()
    wb = Workbook()
    wb.remove(wb.active)

    for key in keys:
        title, report_type = REPORTS[key]
        try:
            text = fetch_report(token, report_type, date_from, date_to)
        except RuntimeError as e:
            # CANCELLED = no returns in range; FATAL = e.g. no FBA on the account
            print(f"Skipping {title}: {e}")
            continue
        if not text.strip():
            print(f"Skipping {title}: empty report")
            continue
        rows = write_sheet(wb, title, text)
        print(f"{title}: {rows} rows")

    if not wb.sheetnames:
        sys.exit("No returns data returned for this range.")

    fmt = lambda d: f"{d[8:10]}-{d[5:7]}-{d[2:4]}"
    out = f"Amazon_Returns-{fmt(date_from)}_to_{fmt(date_to)}.xlsx"
    wb.save(out)
    print(f"Saved -> {out}")


if __name__ == "__main__":
    main()
