# How to run — complete SQLite migration

**Every business table in this ERP PoC is now SQLite-backed.** This
checkpoint completes the migration that started with BOM_Items and
ended with Payments/Payment_Applications — S2C, MFG, and O2C are all
fully migrated.

## Requirements
```
pip install -r requirements.txt
```
(Python 3.12 recommended; sqlite3 is stdlib.)

## Deploy to Streamlit Community Cloud

1. Push this folder to a GitHub repository (keep `erp_pilot.db` in the repository;
   it contains the demo data).
2. In Streamlit Community Cloud, select **Create app**, choose that repository and
   branch, and set the main file path to `erp_ui.py`.
3. Deploy. `requirements.txt`, `runtime.txt`, and `.streamlit/config.toml` are
   already configured for the cloud runtime.

The bundled SQLite database is suitable for a demonstration deployment. Changes
made in the hosted app are stored on the instance's ephemeral filesystem and can
be lost when Streamlit restarts or redeploys the app. Use an external database
before treating this as a production or multi-user system.

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
streamlit run streamlit_app.py  # Complete unified ERP suite
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
