"""
db.py — SQLite foundation for the BOM -> PR -> Consolidation -> PO pilot.

Scope (per agreed pilot design): only these 6 tables move to SQLite.
Everything else (Item Master, Vendor_Master, Delivery_Locations, Sales
Orders, Inventory, GR's own tables, Contracts, etc.) stays in
data.xlsx — deliberately hybrid, to prove "SQLite table + Excel
reference data read together correctly" before deciding on a full
migration.

One live database file, `erp_pilot.db`, mirrors the "one live
data.xlsx" pattern: a single shared file, not per-session state.

Indexes are placed specifically on the columns that caused every real
performance bug found earlier in this project with openpyxl (no
indexing at all, .cell(r,c) in a loop measured at 277x slower than
iter_rows()):
  - status columns (PR_Items.status, RFP.status) — repeatedly filtered
    on "Open" / "PO Proposed" / "RFP" across Consolidate and resets.
  - material code columns — BOM explosion, netting, and coverage
    checks all look up by material code repeatedly.
  - PO_Number / PR_Number — every join between header and line tables,
    and every numbering-sequence scan, keys off these.

get_connection() returns row objects that behave like dicts via
sqlite3.Row + row_factory, so dict(row) produces the exact same shape
every function in this project already expects from its Excel reads —
no downstream code needs to know or care which storage it's talking to.
"""

import os
import sqlite3
from datetime import datetime, date

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "erp_pilot.db")

SCHEMA = """
CREATE TABLE IF NOT EXISTS bom_items (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    parent_code      TEXT NOT NULL,
    parent_desc      TEXT,
    component_code   TEXT NOT NULL,
    component_desc   TEXT,
    qty_per          REAL NOT NULL,
    uom              TEXT,
    notes            TEXT,
    UNIQUE(parent_code, component_code)
);
CREATE INDEX IF NOT EXISTS idx_bom_parent    ON bom_items(parent_code);
CREATE INDEX IF NOT EXISTS idx_bom_component ON bom_items(component_code);

CREATE TABLE IF NOT EXISTS pr_header (
    pr_number        TEXT PRIMARY KEY,
    requester_id     TEXT,
    requester_name   TEXT,
    requester_dept   TEXT,
    project_id       TEXT,
    po_type          TEXT,
    legal_entity     TEXT,
    purchase_entity  TEXT,
    purchasing_group TEXT,
    currency         TEXT,
    plant_code       TEXT,
    pr_date          TEXT
);

CREATE TABLE IF NOT EXISTS pr_items (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    pr_number           TEXT NOT NULL,
    pr_line_item        INTEGER NOT NULL,
    preferred_vendor    TEXT,
    material_code       TEXT,
    material_desc       TEXT,
    uom                 TEXT,
    quantity            REAL,
    required_date       TEXT,
    delivery_location   TEXT,
    delivery_geolocation TEXT,
    status              TEXT,
    quantity_accepted   REAL,
    last_accepted_date  TEXT,
    po_number           TEXT,
    po_item             INTEGER,
    UNIQUE(pr_number, pr_line_item),
    FOREIGN KEY (pr_number) REFERENCES pr_header(pr_number)
);
CREATE INDEX IF NOT EXISTS idx_pri_pr_number    ON pr_items(pr_number);
CREATE INDEX IF NOT EXISTS idx_pri_status       ON pr_items(status);
CREATE INDEX IF NOT EXISTS idx_pri_material     ON pr_items(material_code);
CREATE INDEX IF NOT EXISTS idx_pri_po_number     ON pr_items(po_number);

CREATE TABLE IF NOT EXISTS po_header (
    po_number             TEXT PRIMARY KEY,
    po_type               TEXT,
    legal_entity          TEXT,
    purchase_entity       TEXT,
    purchasing_group      TEXT,
    currency               TEXT,
    plant_code            TEXT,
    supplier_id           TEXT,
    supplier_name         TEXT,
    supplier_geolocation  TEXT,
    status                TEXT DEFAULT 'Proposed',
    po_date               TEXT
);

CREATE TABLE IF NOT EXISTS po_items (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    po_number             TEXT NOT NULL,
    po_item               INTEGER NOT NULL,
    material_code         TEXT,
    material_desc         TEXT,
    uom                   TEXT,
    quantity              REAL,
    unit_price            REAL,
    delivery_date         TEXT,
    delivery_location     TEXT,
    delivery_geolocation  TEXT,
    source_pr_number      TEXT,
    source_pr_line_item   INTEGER,
    requester_id          TEXT,
    requester_dept        TEXT,
    project_id            TEXT,
    UNIQUE(po_number, po_item),
    FOREIGN KEY (po_number) REFERENCES po_header(po_number)
);
CREATE INDEX IF NOT EXISTS idx_poi_po_number  ON po_items(po_number);
CREATE INDEX IF NOT EXISTS idx_poi_material   ON po_items(material_code);
CREATE INDEX IF NOT EXISTS idx_poi_source_pr  ON po_items(source_pr_number);

CREATE TABLE IF NOT EXISTS rfp (
    rfp_number           TEXT PRIMARY KEY,
    rfp_date             TEXT,
    material_code        TEXT,
    material_desc        TEXT,
    uom                  TEXT,
    total_qty            REAL,
    required_by_date     TEXT,
    delivery_location    TEXT,
    delivery_geolocation TEXT,
    source_pr_numbers    TEXT,
    source_pr_lines      TEXT,
    requester_depts      TEXT,
    project_ids          TEXT,
    specifications       TEXT,
    closing_date         TEXT,
    status               TEXT
);
CREATE INDEX IF NOT EXISTS idx_rfp_status    ON rfp(status);
CREATE INDEX IF NOT EXISTS idx_rfp_material  ON rfp(material_code);

CREATE TABLE IF NOT EXISTS item_master (
    item_code       TEXT PRIMARY KEY,
    item_desc       TEXT,
    category        TEXT,
    subcategory     TEXT,
    uom             TEXT,
    unit_price      REAL,
    lead_time_days  INTEGER,
    in_stock        REAL,
    tags            TEXT,
    active          TEXT,
    hsn_code        TEXT,
    gst_rate        REAL,
    weight_kg       REAL
);
CREATE INDEX IF NOT EXISTS idx_item_active   ON item_master(active);
CREATE INDEX IF NOT EXISTS idx_item_category ON item_master(category);

CREATE TABLE IF NOT EXISTS vendor_master (
    vendor_id         TEXT PRIMARY KEY,
    vendor_name       TEXT,
    geolocation       TEXT,
    city              TEXT,
    country           TEXT,
    address           TEXT,
    contact_name      TEXT,
    contact_email     TEXT,
    active            TEXT,
    gstin             TEXT,
    pan               TEXT,
    bank_account_no   TEXT,
    ifsc              TEXT,
    bank_name         TEXT,
    bank_branch       TEXT,
    onboarding_status TEXT,
    kyc_flag          TEXT,
    onboarded_date    TEXT,
    vendor_type       TEXT
);
CREATE INDEX IF NOT EXISTS idx_vendor_active ON vendor_master(active);

CREATE TABLE IF NOT EXISTS delivery_locations (
    location_id   TEXT PRIMARY KEY,
    location_name TEXT,
    geolocation   TEXT,
    city          TEXT,
    state         TEXT,
    country       TEXT,
    address       TEXT,
    active        TEXT
);
CREATE INDEX IF NOT EXISTS idx_delivloc_active ON delivery_locations(active);

CREATE TABLE IF NOT EXISTS vendor_documents (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id     TEXT,
    vendor_id       TEXT,
    doc_type        TEXT,
    filename        TEXT,
    uploaded_date   TEXT,
    status          TEXT,
    notes           TEXT
);
CREATE INDEX IF NOT EXISTS idx_vdoc_vendor ON vendor_documents(vendor_id);

CREATE TABLE IF NOT EXISTS contracts (
    contract_id       TEXT PRIMARY KEY,
    vendor_id         TEXT,
    vendor_name       TEXT,
    status            TEXT,
    start_date        TEXT,
    end_date          TEXT,
    payment_terms     TEXT,
    delivery_sla_days INTEGER,
    currency          TEXT,
    created_date      TEXT,
    source_po         TEXT,
    auto_renew        TEXT,
    notes             TEXT
);
CREATE INDEX IF NOT EXISTS idx_contracts_status ON contracts(status);
CREATE INDEX IF NOT EXISTS idx_contracts_vendor ON contracts(vendor_id);

CREATE TABLE IF NOT EXISTS contract_items (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    contract_id           TEXT NOT NULL,
    line_item             INTEGER NOT NULL,
    material_code         TEXT,
    material_desc         TEXT,
    uom                   TEXT,
    contracted_unit_price REAL,
    min_order_qty         REAL,
    lead_time_days        INTEGER,
    UNIQUE(contract_id, line_item)
);
CREATE INDEX IF NOT EXISTS idx_citems_contract ON contract_items(contract_id);
CREATE INDEX IF NOT EXISTS idx_citems_material ON contract_items(material_code);

CREATE TABLE IF NOT EXISTS vrq_requests (
    vrq_id        TEXT PRIMARY KEY,
    vendor_name   TEXT,
    contact_email TEXT,
    sent_date     TEXT,
    status        TEXT,
    filename      TEXT,
    vendor_id     TEXT
);
CREATE INDEX IF NOT EXISTS idx_vrq_status ON vrq_requests(status);

CREATE TABLE IF NOT EXISTS vrq_responses (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    vrq_id       TEXT NOT NULL,
    question_key TEXT NOT NULL,
    section      TEXT,
    question     TEXT,
    answer       TEXT,
    UNIQUE(vrq_id, question_key)
);
CREATE INDEX IF NOT EXISTS idx_vrqresp_vrq ON vrq_responses(vrq_id);

CREATE TABLE IF NOT EXISTS rfx_quotes (
    quote_id       TEXT PRIMARY KEY,
    rfp_number     TEXT,
    vendor_id      TEXT,
    vendor_name    TEXT,
    quoted_price   REAL,
    lead_time_days INTEGER,
    moq            REAL,
    quote_date     TEXT,
    status         TEXT,
    notes          TEXT
);
CREATE INDEX IF NOT EXISTS idx_rfxq_rfp    ON rfx_quotes(rfp_number);
CREATE INDEX IF NOT EXISTS idx_rfxq_vendor ON rfx_quotes(vendor_id);
CREATE INDEX IF NOT EXISTS idx_rfxq_status ON rfx_quotes(status);

CREATE TABLE IF NOT EXISTS rfx_invitations (
    invitation_id TEXT PRIMARY KEY,
    rfp_number    TEXT,
    vendor_id     TEXT,
    vendor_name   TEXT,
    invited_date  TEXT,
    filename      TEXT
);
CREATE INDEX IF NOT EXISTS idx_rfxi_rfp    ON rfx_invitations(rfp_number);
CREATE INDEX IF NOT EXISTS idx_rfxi_vendor ON rfx_invitations(vendor_id);

CREATE TABLE IF NOT EXISTS purchase_bundles (
    bundle_id    TEXT PRIMARY KEY,
    bundle_name  TEXT NOT NULL,
    description  TEXT,
    department   TEXT,
    created_by   TEXT,
    created_date TEXT,
    active       TEXT DEFAULT 'Yes'
);
CREATE INDEX IF NOT EXISTS idx_bundles_active ON purchase_bundles(active);
CREATE INDEX IF NOT EXISTS idx_bundles_dept   ON purchase_bundles(department);

CREATE TABLE IF NOT EXISTS purchase_bundle_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    bundle_id     TEXT NOT NULL,
    material_code TEXT NOT NULL,
    material_desc TEXT,
    uom           TEXT,
    default_qty   REAL NOT NULL,
    notes         TEXT,
    UNIQUE(bundle_id, material_code)
);
CREATE INDEX IF NOT EXISTS idx_bitems_bundle   ON purchase_bundle_items(bundle_id);
CREATE INDEX IF NOT EXISTS idx_bitems_material ON purchase_bundle_items(material_code);

CREATE TABLE IF NOT EXISTS gr_header (
    gr_id             TEXT PRIMARY KEY,
    po_number         TEXT,
    vendor_id         TEXT,
    vendor_name       TEXT,
    gr_date           TEXT,
    status            TEXT,
    delivery_location TEXT,
    received_by       TEXT,
    notes             TEXT
);
CREATE INDEX IF NOT EXISTS idx_grh_po     ON gr_header(po_number);
CREATE INDEX IF NOT EXISTS idx_grh_status ON gr_header(status);

CREATE TABLE IF NOT EXISTS gr_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    gr_id         TEXT NOT NULL,
    line_item     INTEGER,
    po_item       INTEGER,
    material_code TEXT,
    material_desc TEXT,
    uom           TEXT,
    po_qty        REAL,
    qty_received  REAL,
    unit_price    REAL,
    notes         TEXT
);
CREATE INDEX IF NOT EXISTS idx_gri_gr       ON gr_items(gr_id);
CREATE INDEX IF NOT EXISTS idx_gri_material ON gr_items(material_code);

CREATE TABLE IF NOT EXISTS quality_inspections (
    qi_id           TEXT PRIMARY KEY,
    gr_id           TEXT,
    po_number       TEXT,
    line_item       INTEGER,
    po_item         INTEGER,
    material_code   TEXT,
    material_desc   TEXT,
    qty_received    REAL,
    qty_passed      REAL,
    qty_failed      REAL,
    inspected_by    TEXT,
    inspection_date TEXT,
    notes           TEXT,
    UNIQUE(gr_id, po_item)
);
CREATE INDEX IF NOT EXISTS idx_qi_gr ON quality_inspections(gr_id);
CREATE INDEX IF NOT EXISTS idx_qi_po ON quality_inspections(po_number);

CREATE TABLE IF NOT EXISTS production_confirmations (
    confirmation_id   TEXT PRIMARY KEY,
    parent_item_code  TEXT,
    parent_item_desc  TEXT,
    quantity_built    REAL,
    location_id       TEXT,
    confirmation_date TEXT,
    confirmed_by      TEXT,
    notes             TEXT
);
CREATE INDEX IF NOT EXISTS idx_pconf_parent ON production_confirmations(parent_item_code);

CREATE TABLE IF NOT EXISTS inventory_transactions (
    txn_id         TEXT PRIMARY KEY,
    txn_date       TEXT,
    material_code  TEXT,
    material_desc  TEXT,
    location_id    TEXT,
    location_name  TEXT,
    quantity       REAL,
    txn_type       TEXT,
    reference_type TEXT,
    reference_id   TEXT,
    notes          TEXT
);
CREATE INDEX IF NOT EXISTS idx_txn_material          ON inventory_transactions(material_code);
CREATE INDEX IF NOT EXISTS idx_txn_location          ON inventory_transactions(location_id);
CREATE INDEX IF NOT EXISTS idx_txn_material_location ON inventory_transactions(material_code, location_id);

CREATE TABLE IF NOT EXISTS stock_transfers (
    transfer_id    TEXT PRIMARY KEY,
    material_code  TEXT,
    material_desc  TEXT,
    uom            TEXT,
    quantity       REAL,
    from_location  TEXT,
    to_location    TEXT,
    status         TEXT DEFAULT 'In Transit',
    shipped_date   TEXT,
    shipped_by     TEXT,
    received_date  TEXT,
    received_by    TEXT,
    received_qty   REAL,
    discrepancy_type  TEXT,
    discrepancy_notes TEXT,
    gl_goods_value REAL,
    gl_igst_amount REAL,
    cancelled_date TEXT,
    cancelled_by   TEXT,
    source_type    TEXT DEFAULT 'Ad Hoc',
    source_doc     TEXT,
    eway_bill_number TEXT,
    eway_bill_valid_until TEXT,
    carrier        TEXT,
    tracking_ref   TEXT,
    notes          TEXT
);
CREATE INDEX IF NOT EXISTS idx_txn_reference         ON inventory_transactions(reference_id);

CREATE TABLE IF NOT EXISTS sto_header (
    sto_id         TEXT PRIMARY KEY,
    material_code  TEXT,
    material_desc  TEXT,
    hub_location   TEXT,
    total_qty      REAL,
    allocation_rule TEXT,
    created_date   TEXT,
    created_by     TEXT,
    notes          TEXT
);

CREATE TABLE IF NOT EXISTS sto_lines (
    sto_id         TEXT,
    line_no        INTEGER,
    to_location    TEXT,
    requested_qty  REAL,
    allocated_qty  REAL,
    transfer_id    TEXT,
    PRIMARY KEY (sto_id, line_no)
);
CREATE INDEX IF NOT EXISTS idx_sto_lines_sto ON sto_lines(sto_id);

CREATE TABLE IF NOT EXISTS reservations (
    reservation_id    TEXT PRIMARY KEY,
    material_code     TEXT,
    material_desc     TEXT,
    location_id       TEXT,
    quantity          REAL,
    so_id             TEXT,
    so_line_item      INTEGER,
    status            TEXT DEFAULT 'Open',
    created_date      TEXT,
    resolved_date     TEXT,
    resolution_type   TEXT,
    resolution_notes  TEXT
);
CREATE INDEX IF NOT EXISTS idx_reservations_matloc ON reservations(material_code, location_id);
CREATE INDEX IF NOT EXISTS idx_reservations_so ON reservations(so_id);

CREATE TABLE IF NOT EXISTS backorders (
    backorder_id      TEXT PRIMARY KEY,
    material_code     TEXT,
    material_desc     TEXT,
    location_id       TEXT,
    original_qty      REAL,
    open_qty          REAL,
    so_id             TEXT,
    so_line_item      INTEGER,
    status            TEXT DEFAULT 'Open',
    created_date      TEXT,
    resolved_date     TEXT,
    resolution_notes  TEXT
);
CREATE INDEX IF NOT EXISTS idx_backorders_matloc ON backorders(material_code, location_id);
CREATE INDEX IF NOT EXISTS idx_backorders_so ON backorders(so_id);

CREATE TABLE IF NOT EXISTS plant_pair_lead_time_estimates (
    from_location  TEXT NOT NULL,
    to_location    TEXT NOT NULL,
    estimated_days INTEGER,
    PRIMARY KEY (from_location, to_location)
);

CREATE TABLE IF NOT EXISTS material_location_planning_params (
    material_code       TEXT NOT NULL,
    location            TEXT NOT NULL,
    min_qty             REAL,
    max_qty             REAL,
    reorder_cadence_days INTEGER,
    PRIMARY KEY (material_code, location)
);

CREATE TABLE IF NOT EXISTS customer_master (
    customer_id       TEXT PRIMARY KEY,
    customer_name     TEXT,
    customer_type     TEXT,
    geolocation       TEXT,
    city              TEXT,
    country           TEXT,
    address           TEXT,
    contact_name      TEXT,
    contact_email     TEXT,
    contact_phone     TEXT,
    gstin             TEXT,
    pan               TEXT,
    credit_limit      REAL,
    credit_status     TEXT,
    payment_terms     TEXT,
    onboarding_status TEXT,
    kyc_flag          TEXT,
    default_delivery_location TEXT,
    onboarded_date    TEXT,
    active            TEXT
);
CREATE INDEX IF NOT EXISTS idx_cm_active ON customer_master(active);

CREATE TABLE IF NOT EXISTS customer_documents (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id   TEXT,
    customer_id   TEXT,
    doc_type      TEXT,
    filename      TEXT,
    uploaded_date TEXT,
    status        TEXT,
    notes         TEXT
);
CREATE INDEX IF NOT EXISTS idx_cdoc_customer ON customer_documents(customer_id);

CREATE TABLE IF NOT EXISTS org_profile (
    org_id          TEXT PRIMARY KEY,
    legal_name      TEXT,
    gstin           TEXT,
    pan             TEXT,
    address         TEXT,
    city            TEXT,
    state           TEXT,
    country         TEXT,
    bank_account_no TEXT,
    ifsc            TEXT,
    bank_name       TEXT,
    contact_email   TEXT,
    contact_phone   TEXT
);

CREATE TABLE IF NOT EXISTS org_defaults (
    org_element   TEXT PRIMARY KEY,
    default_value TEXT
);

CREATE TABLE IF NOT EXISTS legal_entities (
    le_id           TEXT PRIMARY KEY,
    le_name         TEXT,
    gstin           TEXT,
    pan             TEXT,
    address         TEXT,
    city            TEXT,
    state           TEXT,
    country         TEXT,
    bank_account_no TEXT,
    ifsc            TEXT,
    bank_name       TEXT,
    contact_email   TEXT,
    contact_phone   TEXT
);

CREATE TABLE IF NOT EXISTS vendor_types (
    vendor_type TEXT PRIMARY KEY
);

CREATE TABLE IF NOT EXISTS customer_types (
    customer_type TEXT PRIMARY KEY
);

CREATE TABLE IF NOT EXISTS quotes (
    quote_id      TEXT PRIMARY KEY,
    customer_id   TEXT,
    customer_name TEXT,
    quote_date    TEXT,
    valid_until   TEXT,
    payment_terms TEXT,
    status        TEXT,
    currency      TEXT,
    total_value   REAL,
    notes         TEXT,
    filename      TEXT
);
CREATE INDEX IF NOT EXISTS idx_quotes_customer ON quotes(customer_id);
CREATE INDEX IF NOT EXISTS idx_quotes_status   ON quotes(status);

CREATE TABLE IF NOT EXISTS quote_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    quote_id      TEXT NOT NULL,
    line_item     INTEGER,
    material_code TEXT,
    material_desc TEXT,
    uom           TEXT,
    qty           REAL,
    unit_price    REAL,
    line_total    REAL
);
CREATE INDEX IF NOT EXISTS idx_qitems_quote ON quote_items(quote_id);

CREATE TABLE IF NOT EXISTS sales_orders (
    so_id                   TEXT PRIMARY KEY,
    customer_id             TEXT,
    customer_name           TEXT,
    order_date              TEXT,
    status                  TEXT,
    source_quote            TEXT,
    payment_terms           TEXT,
    currency                TEXT,
    total_value             REAL,
    delivery_location       TEXT,
    delivery_geolocation    TEXT,
    requested_delivery_date TEXT,
    notes                   TEXT
);
CREATE INDEX IF NOT EXISTS idx_so_customer     ON sales_orders(customer_id);
CREATE INDEX IF NOT EXISTS idx_so_status       ON sales_orders(status);
CREATE INDEX IF NOT EXISTS idx_so_source_quote ON sales_orders(source_quote);

CREATE TABLE IF NOT EXISTS sales_order_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    so_id         TEXT NOT NULL,
    line_item     INTEGER,
    material_code TEXT,
    material_desc TEXT,
    uom           TEXT,
    qty           REAL,
    unit_price    REAL,
    line_total    REAL,
    atp_outcome      TEXT,
    promised_qty     REAL,
    backordered_qty  REAL,
    reservation_id   TEXT,
    backorder_id     TEXT
);
CREATE INDEX IF NOT EXISTS idx_soitems_so ON sales_order_items(so_id);

CREATE TABLE IF NOT EXISTS fulfillments (
    fulfillment_id    TEXT PRIMARY KEY,
    so_id             TEXT,
    customer_id       TEXT,
    customer_name     TEXT,
    status            TEXT,
    created_date      TEXT,
    shipped_date      TEXT,
    delivered_date    TEXT,
    carrier           TEXT,
    tracking_ref      TEXT,
    delivery_location TEXT,
    pod_reference     TEXT,
    notes             TEXT
);
CREATE INDEX IF NOT EXISTS idx_ful_so     ON fulfillments(so_id);
CREATE INDEX IF NOT EXISTS idx_ful_status ON fulfillments(status);

CREATE TABLE IF NOT EXISTS fulfillment_items (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    fulfillment_id TEXT NOT NULL,
    line_item      INTEGER,
    material_code  TEXT,
    material_desc  TEXT,
    uom            TEXT,
    qty_ordered    REAL,
    qty_shipped    REAL
);
CREATE INDEX IF NOT EXISTS idx_fulitems_ful ON fulfillment_items(fulfillment_id);

CREATE TABLE IF NOT EXISTS invoices (
    invoice_id      TEXT PRIMARY KEY,
    fulfillment_id  TEXT,
    so_id           TEXT,
    customer_id     TEXT,
    customer_name   TEXT,
    customer_gstin  TEXT,
    invoice_date    TEXT,
    due_date        TEXT,
    status          TEXT,
    payment_terms   TEXT,
    currency        TEXT,
    place_of_supply TEXT,
    subtotal        REAL,
    cgst_total      REAL,
    sgst_total      REAL,
    igst_total      REAL,
    grand_total     REAL,
    notes           TEXT
);
CREATE INDEX IF NOT EXISTS idx_inv_fulfillment ON invoices(fulfillment_id);
CREATE INDEX IF NOT EXISTS idx_inv_status       ON invoices(status);
CREATE INDEX IF NOT EXISTS idx_inv_customer     ON invoices(customer_id);

CREATE TABLE IF NOT EXISTS invoice_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id    TEXT NOT NULL,
    line_item     INTEGER,
    material_code TEXT,
    material_desc TEXT,
    hsn_code      TEXT,
    uom           TEXT,
    qty           REAL,
    unit_price    REAL,
    taxable_value REAL,
    gst_rate      REAL,
    cgst_amount   REAL,
    sgst_amount   REAL,
    igst_amount   REAL,
    line_total    REAL
);
CREATE INDEX IF NOT EXISTS idx_invitems_invoice ON invoice_items(invoice_id);

CREATE TABLE IF NOT EXISTS chart_of_accounts (
    account_code TEXT PRIMARY KEY,
    account_name TEXT,
    account_type TEXT,
    description  TEXT
);

CREATE TABLE IF NOT EXISTS journal_entries (
    je_id        TEXT PRIMARY KEY,
    entry_date   TEXT,
    source_type  TEXT,
    source_id    TEXT,
    description  TEXT,
    total_debit  REAL,
    total_credit REAL
);
CREATE INDEX IF NOT EXISTS idx_je_source ON journal_entries(source_type, source_id);

CREATE TABLE IF NOT EXISTS journal_entry_lines (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    je_id        TEXT NOT NULL,
    line_item    INTEGER,
    account_code TEXT,
    account_name TEXT,
    debit        REAL,
    credit       REAL,
    description  TEXT
);
CREATE INDEX IF NOT EXISTS idx_jelines_je      ON journal_entry_lines(je_id);
CREATE INDEX IF NOT EXISTS idx_jelines_account ON journal_entry_lines(account_code);

CREATE TABLE IF NOT EXISTS payments (
    payment_id       TEXT PRIMARY KEY,
    customer_id      TEXT,
    customer_name    TEXT,
    payment_date     TEXT,
    amount           REAL,
    payment_method   TEXT,
    reference_no     TEXT,
    unapplied_amount REAL,
    notes            TEXT
);
CREATE INDEX IF NOT EXISTS idx_pay_customer ON payments(customer_id);

CREATE TABLE IF NOT EXISTS payment_applications (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    payment_id           TEXT NOT NULL,
    line_item            INTEGER,
    invoice_id           TEXT,
    applied_amount       REAL,
    short_payment_reason TEXT,
    application_date     TEXT
);
CREATE INDEX IF NOT EXISTS idx_payapp_payment ON payment_applications(payment_id);
CREATE INDEX IF NOT EXISTS idx_payapp_invoice ON payment_applications(invoice_id);

CREATE TABLE IF NOT EXISTS vendor_invoices (
    invoice_id     TEXT PRIMARY KEY,
    gr_id          TEXT,
    po_number      TEXT,
    vendor_id      TEXT,
    vendor_name    TEXT,
    invoice_number TEXT,
    invoice_date   TEXT,
    amount         REAL,
    paid_amount    REAL,
    status         TEXT,
    notes          TEXT
);
CREATE INDEX IF NOT EXISTS idx_vi_gr     ON vendor_invoices(gr_id);
CREATE INDEX IF NOT EXISTS idx_vi_status ON vendor_invoices(status);

CREATE TABLE IF NOT EXISTS vendor_invoice_payments (
    payment_id     TEXT PRIMARY KEY,
    invoice_id     TEXT,
    vendor_id      TEXT,
    vendor_name    TEXT,
    payment_date   TEXT,
    amount         REAL,
    payment_method TEXT,
    reference_no   TEXT,
    notes          TEXT
);
CREATE INDEX IF NOT EXISTS idx_vip_invoice ON vendor_invoice_payments(invoice_id);
"""


def get_connection(db_file=None):
    """One connection per call, row_factory set so dict(row) works
    exactly like the dict shapes every existing function expects from
    Excel reads. Foreign keys on, since real referential integrity
    (unlike Excel) is one of the actual points of this migration."""
    path = db_file or DB_FILE
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _add_column_if_missing(conn, table, column, coltype):
    """CREATE TABLE IF NOT EXISTS alone never retroactively adds a
    column to a table that already exists — a schema change here only
    takes effect for a brand-new database unless there's also an
    explicit ALTER TABLE for anyone whose database predates the change.
    Returns True if the column was actually added (so callers can run
    a one-time backfill only when it's genuinely needed)."""
    cols = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coltype}")
        return True
    return False


def _backfill_pr_dates(conn):
    """One-time courtesy for any pr_header row that existed before
    pr_date was added (2026-07-29) — parses each PR's creation date out
    of its own number (PR-YYYYMMDD-NNN) one last time, so existing data
    doesn't go blank the moment the column appears. New rows never need
    this — pr_consolidation.create_pr() sets pr_date directly from here
    on, not by parsing anything."""
    rows = conn.execute("SELECT pr_number FROM pr_header WHERE pr_date IS NULL").fetchall()
    for r in rows:
        try:
            d = datetime.strptime(r["pr_number"].split("-")[1], "%Y%m%d").date().isoformat()
            conn.execute("UPDATE pr_header SET pr_date=? WHERE pr_number=?", (d, r["pr_number"]))
        except (IndexError, ValueError):
            pass  # non-standard PR number, nothing to backfill from


def _backfill_po_dates(conn):
    """One-time courtesy for any po_header row that existed before
    po_date was added (2026-07-29) — but unlike pr_date, there's no
    ID-embedded date to recover here at all: PO numbers are
    'PO-YYYY-NNNN', year-level only, never day-level. Backfilling to
    today is a deliberate, honest approximation, not a real recovered
    date — chosen because leaving it NULL would silently and
    permanently exclude every pre-existing PO from any date-aware
    projection (get_open_po_lines_with_dates() and everything that
    reads it), which is worse than a slightly-wrong-but-present one.
    In practice this only affects POs still open long enough to still
    matter after this migration runs — anything already fully received
    doesn't show up in open-PO exposure regardless of its date."""
    conn.execute("UPDATE po_header SET po_date=? WHERE po_date IS NULL", (date.today().isoformat(),))


def _backfill_customer_default_locations(conn):
    """One-time courtesy for any customer_master row with no
    default_delivery_location set (2026-07-30) — matches on city
    against delivery_locations, case-insensitive. Deliberately does
    NOT guess when there's no exact city match: left NULL rather than
    picking a plausible-looking nearest location, since a wrong default
    silently applied to every future order for that customer is worse
    than requiring an explicit location on each one. This is a courtesy
    backfill for data that predates the field, not a substitute for
    customer_onboarding.py actually capturing a real default going
    forward."""
    conn.execute(
        "UPDATE customer_master SET default_delivery_location = ("
        "  SELECT dl.location_id FROM delivery_locations dl"
        "  WHERE LOWER(dl.city) = LOWER(customer_master.city) AND dl.active = 'Yes'"
        "  LIMIT 1"
        ") WHERE default_delivery_location IS NULL AND city IS NOT NULL"
    )


def init_schema(db_file=None):
    """Idempotent — safe to call every time the app starts, same
    spirit as the Excel side's ensure_sheets()/ensure_schema()."""
    conn = get_connection(db_file)
    try:
        conn.executescript(SCHEMA)
        _add_column_if_missing(conn, "pr_header", "pr_date", "TEXT")
        _add_column_if_missing(conn, "po_header", "po_date", "TEXT")
        _add_column_if_missing(conn, "customer_master", "default_delivery_location", "TEXT")
        _add_column_if_missing(conn, "item_master", "weight_kg", "REAL")
        _add_column_if_missing(conn, "stock_transfers", "cancelled_date", "TEXT")
        _add_column_if_missing(conn, "stock_transfers", "cancelled_by", "TEXT")
        _add_column_if_missing(conn, "stock_transfers", "received_qty", "REAL")
        _add_column_if_missing(conn, "stock_transfers", "discrepancy_type", "TEXT")
        _add_column_if_missing(conn, "stock_transfers", "discrepancy_notes", "TEXT")
        # DEFAULT is specified here, not just "TEXT" -- a column added via
        # ALTER TABLE only gets a real SQLite-level default if the ALTER
        # statement itself says so; the schema string's own "DEFAULT
        # 'Ad Hoc'" above only applies when a table is created fresh via
        # that exact CREATE TABLE statement, not to a column added later
        # via migration. Found this directly: an ad hoc ship on an
        # already-migrated database was inserting NULL for source_type
        # instead of 'Ad Hoc', since the earlier migration call here only
        # said "TEXT" with no default at all.
        added_source_type = _add_column_if_missing(conn, "stock_transfers", "source_type",
                                                    "TEXT DEFAULT 'Ad Hoc'")
        if added_source_type:
            # A migrated database's existing rows would otherwise be NULL here,
            # not 'Ad Hoc' — every transfer that existed before this column did
            # was genuinely an ad hoc one (STOs didn't exist yet either), so
            # this is a real, correct backfill, not a guess.
            conn.execute("UPDATE stock_transfers SET source_type='Ad Hoc' WHERE source_type IS NULL")
        _add_column_if_missing(conn, "stock_transfers", "source_doc", "TEXT")
        _add_column_if_missing(conn, "stock_transfers", "eway_bill_number", "TEXT")
        _add_column_if_missing(conn, "stock_transfers", "eway_bill_valid_until", "TEXT")
        _add_column_if_missing(conn, "stock_transfers", "gl_goods_value", "REAL")
        _add_column_if_missing(conn, "stock_transfers", "gl_igst_amount", "REAL")
        _add_column_if_missing(conn, "sales_order_items", "atp_outcome", "TEXT")
        _add_column_if_missing(conn, "sales_order_items", "promised_qty", "REAL")
        _add_column_if_missing(conn, "sales_order_items", "backordered_qty", "REAL")
        _add_column_if_missing(conn, "sales_order_items", "reservation_id", "TEXT")
        _add_column_if_missing(conn, "sales_order_items", "backorder_id", "TEXT")
        # None of these backfills are gated on their column having just
        # been added — a row with a NULL value can in principle turn up
        # later too (anything that ever inserts without going through
        # the normal create_pr()/mark_po_created()/customer_onboarding
        # path), and all three are cheap no-ops once nothing's missing,
        # safe to run on every startup.
        _backfill_pr_dates(conn)
        _backfill_po_dates(conn)
        _backfill_customer_default_locations(conn)
        conn.commit()
    finally:
        conn.close()


if __name__ == "__main__":
    init_schema()
    print(f"Schema initialized at {DB_FILE}")
