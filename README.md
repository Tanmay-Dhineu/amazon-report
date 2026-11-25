# Amazon All Orders Report -> Excel

1. Copy `.env.example` to `.env` and fill in your credentials
2. Install: `pip install -r requirements.txt`
3. Run: `python amazon_orders_report.py 2026-03-01 2026-03-31`  (max 30 days per run)

Output: `Amazon_Order-01-03-26_to_31-03-26.xlsx` in this folder.
## Current inventory

Run: `python amazon_inventory_report.py` (or `fba` / `mfn` for just one sheet)

Output: `Amazon_Inventory-25-11-26.xlsx` (today's date) with sheets:
- **FBA** – fulfillable / reserved / unsellable / inbound qty per SKU
- **Listings** – all listings with price, quantity and status

## Ads report (Sponsored Products)

Needs **Amazon Ads API** credentials in `.env` (separate from SP-API):
`ADS_CLIENT_ID`, `ADS_CLIENT_SECRET`, `ADS_REFRESH_TOKEN`, and optionally `ADS_PROFILE_ID`.

Run: `python amazon_ads_report.py 2026-11-01 2026-11-25 [campaign|product|searchterm]`
(max 31 days per run, last ~95 days only)

Output: `Amazon_Ads_campaign-01-11-26_to_25-11-26.xlsx` – daily impressions, clicks,
spend, 7-day sales/orders/units, CTR, CPC, ACOS %, ROAS.

Do not share or commit the `.env` file.
