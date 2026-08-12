# How to run — complete SQLite migration

**Every business table in this ERP PoC is now SQLite-backed.** This
checkpoint completes the migration that started with BOM_Items and
ended with Payments/Payment_Applications — S2C, MFG, and O2C are all
fully migrated.

## Requirements
```
pip install streamlit openpyxl pandas reportlab
```
(Python 3.10+ recommended; sqlite3 is stdlib. reportlab is used to
generate the real, downloadable E-Way Bill PDF from the Position &
Transfers screen — added alongside that feature; nothing else here
depends on it.)

## Setup
Put every file in this folder in the **same directory**.

`erp_pilot.db` is already migrated and seeded (one demo Purchase
Bundle). If you ever want to reset it from scratch, run every
migration script in this exact order (later ones depend on earlier
ones already being loaded):
```
rm erp_pilot.db
python migrate_bom_to_sqlite.py
python migrate_reference_data.py
python migrate_s2c_remainder.py
python migrate_mfg_remainder.py
python migrate_o2c_customer.py
python migrate_o2c_quotation.py
python migrate_o2c_sales_order.py
python migrate_o2c_fulfillment.py
python migrate_o2c_billing.py
python migrate_o2c_accounting.py
python migrate_o2c_cash_application.py
python seed_bundle_from_bom.py
```

## Run (each in its own terminal)
```
streamlit run erp_ui.py     # Source-to-Pay: Bundles, PR -> PO, RFx, Vendors, Contracts
streamlit run mfg_ui.py     # Manufacturing: Goods Receipt, QC, BOM, Production, Inventory
streamlit run o2c_ui.py     # Order-to-Cash: Quotes, Sales Orders, Fulfillment, Billing, Accounting, Cash
```

## What's SQLite-backed (all of it)

**S2C:** `BOM_Items`, `PR_Header`/`PR_Items`, `PO_Header`/`PO_Items`, `RFP`,
`Item Master`, `Vendor_Master`/`Vendor_Documents`, `Delivery_Locations`,
`Contracts`/`Contract_Items`, `VRQ_Requests`/`VRQ_Responses`,
`RFx_Quotes`/`RFx_Invitations`, `Purchase Bundles` (new capability)

**MFG:** `GR_Header`/`GR_Items`, `Quality_Inspections`,
`Production_Confirmations`, `Inventory_Transactions`

**O2C:** `Customer_Master`/`Customer_Documents`, `Org_Profile`, `Quotes`/
`Quote_Items`, `Sales_Orders`/`Sales_Order_Items`, `Fulfillments`/
`Fulfillment_Items`, `Invoices`/`Invoice_Items`, `Chart_of_Accounts`/
`Journal_Entries`/`Journal_Entry_Lines`, `Payments`/`Payment_Applications`

`data.xlsx` remains present only for its own historical record and
because `openpyxl` is still used to generate standalone downloadable
documents (POs, invoices, delivery notes, etc.) — no business table is
read from or written to it anymore.

## A real bug found and fixed in this final pass
While doing a full sweep for any remaining direct Excel access,
`bom.py`'s `_batch_open_po_exposure()` (used by `net_requirements()`
for open-PO netting) was found still reading `PO_Items`/`GR_Header`/
`GR_Items` via `openpyxl` — both had already moved to SQLite in
earlier slices of this migration and were never being written to in
Excel anymore. This meant `net_requirements()`'s open-PO netting had
been silently computing against empty/stale data since the MFG slice
was completed — not caught earlier because nothing since had
re-exercised that specific code path against a real open PO. Fixed
and verified directly: created a real 20-unit PO, partially received
8 units, confirmed `_batch_open_po_exposure()` correctly reports 12
units outstanding, and confirmed `net_requirements()` nets against
that 12 correctly in both a fully-covered and a partially-covered
scenario.

## Active-flag policy (unchanged, still in effect everywhere)
Master data (`Item Master`, `Vendor_Master`, `Delivery_Locations`) has
`Active` filtering on every listing/selection dropdown — verified
directly by deactivating records and confirming exclusion. Transaction
tables (GR, QC, Production, Inventory) have no `Active` column of
their own; their write paths check referenced master data and warn
(never block) if it's inactive, since a transaction is often
completing a commitment made while the reference was still active.

This checkpoint has been tested end-to-end at every stage — not just
"does it load" — including full real-money-math scenarios (GST
CGST/SGST/IGST splits, COGS postings, trial balance reconciliation,
credit holds, partial shipments/receipts/payments) and a full 21-page
regression across all three apps.
