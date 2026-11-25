# Amazon All Orders Report -> Excel

1. Copy `.env.example` to `.env` and fill in your credentials
2. Install: `pip install -r requirements.txt`
3. Run: `python amazon_orders_report.py 2026-03-01 2026-03-31`  (max 30 days per run)

Output: `Amazon_Order-01-03-26_to_31-03-26.xlsx` in this folder.
## Sales & traffic (Business Reports)

Run: `python amazon_sales_report.py 2026-11-01 2026-11-25 [day|week|month]`

Output: `Amazon_Sales-01-11-26_to_25-11-26.xlsx` with sheets:
- **By Date** – sales, units, orders, refunds, sessions, page views, conversion %, Buy Box % + TOTAL row
- **By SKU** – the same per SKU for the whole range, sorted by sales

## Current inventory

Run: `python amazon_inventory_report.py` (or `fba` / `mfn` for just one sheet)

Output: `Amazon_Inventory-25-11-26.xlsx` (today's date) with sheets:
- **FBA** – fulfillable / reserved / unsellable / inbound qty per SKU
- **Listings** – all listings with price, quantity and status

## Returns

Run: `python amazon_returns_report.py 2026-11-01 2026-11-25 [all|fba|mfn]` (max 60 days, by return date)

Output: `Amazon_Returns-01-11-26_to_25-11-26.xlsx` with sheets:
- **FBA Returns** – returns received at Amazon warehouses: reason, disposition, customer comments
- **MFN Returns** – returns of self-ship / Easy Ship orders: status, RMA, tracking, reason, amounts

## Settlement (payments)

Run: `python amazon_settlement_report.py` (last 90 days) or `python amazon_settlement_report.py 2026-10-01 2026-11-25`

Amazon creates settlement reports by itself every settlement period; this downloads the ones
already created (last 90 days only).

Output: `Amazon_Settlement-27-08-26_to_25-11-26.xlsx` with sheets:
- **Summary** – one row per settlement: period, deposit date, total paid, totals per transaction type
- **Transactions** – every order, refund, fee and adjustment line

## Payments (live, includes unpaid open period)

Run: `python amazon_payments_report.py` (last 90 days) or `python amazon_payments_report.py 2026-10-01 2026-11-25`

Output: `Amazon_Payments-27-08-26_to_25-11-26.xlsx` with sheets:
- **Payments** – one row per payout period (Open / Closed), transfer date, Amazon total
- **Transactions** – every charge, fee, refund, promotion, TCS/TDS and adjustment line

## Fees

- Charged: `python amazon_fees_report.py 2026-11-01 2026-11-25` (max 180 days)
  → `Amazon_Fees-01-11-26_to_25-11-26.xlsx`: totals by fee type, per-SKU fees & fee % of sales, every fee line
- Estimate: `python amazon_fees_report.py estimate`
  → `Amazon_Fee_Estimates-25-11-26.xlsx`: Amazon's estimated fees for every listing at its current price

## Ads report (Sponsored Products)

Needs **Amazon Ads API** credentials in `.env` (separate from SP-API):
`ADS_CLIENT_ID`, `ADS_CLIENT_SECRET`, `ADS_REFRESH_TOKEN`, and optionally `ADS_PROFILE_ID`.

Run: `python amazon_ads_report.py 2026-11-01 2026-11-25 [campaign|product|searchterm]`
(max 31 days per run, last ~95 days only)

Output: `Amazon_Ads_campaign-01-11-26_to_25-11-26.xlsx` – daily impressions, clicks,
spend, 7-day sales/orders/units, CTR, CPC, ACOS %, ROAS.

Do not share or commit the `.env` file.
