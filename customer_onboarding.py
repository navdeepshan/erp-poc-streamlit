"""
customer_onboarding.py — Customer intake & credit management.

This is the O2C mirror of vendor_onboarding.py, and deliberately reuses it
rather than re-implementing anything: GSTIN checksum and PAN format
validation are generic Indian tax-ID logic, not vendor-specific, so this
module imports vendor_onboarding.py and calls its validate_gstin() /
validate_pan_format() directly. One checksum implementation, two intake
forms — if it's ever upgraded (e.g. a real GSTN active-registration
lookup), both sides benefit without touching this file.

Where it genuinely differs from the vendor side: credit risk runs the
opposite direction. Vendor onboarding vets who WE pay; this vets who owes
US. So alongside the same Format-Verified -> Approved gate, a customer
also carries a Credit_Limit and a Credit_Status that can be put on hold
independently of KYC approval — an onboarded, KYC-clean customer can still
be credit-held (e.g. for being delinquent) without re-running onboarding.

Honest scope limit: check_credit_available() can currently only compare a
proposed order against the stated Credit_Limit — it has no outstanding-AR
figure to net against, because Billing/Cash Application don't exist yet.
That's noted in the function itself rather than faked. Once Cash
Application exists, the natural integration is the same pattern
contracts.py used for pr_consolidation.py: Sales Order Management takes an
optional credit_check_fn callable, so this module stays independent and
gets wired in by whatever page creates sales orders — no import needed in
either direction.

New sheets:
  Customer_Master     Customer_ID | Customer_Name | Customer_Type |
                       Geolocation | City | Country | Address |
                       Contact_Name | Contact_Email | Contact_Phone |
                       GSTIN | PAN | Credit_Limit | Credit_Status |
                       Payment_Terms | Onboarding_Status | KYC_Flag |
                       Onboarded_Date | Active
  Customer_Documents   Document_ID | Customer_ID | Doc_Type | Filename |
                       Uploaded_Date | Status | Notes

SQLite pilot: Customer_Master and Customer_Documents now live in
erp_pilot.db (tables `customer_master`, `customer_documents`) — this
module is the exclusive owner of both, same "single owner" pattern as
vendor_onboarding.py for Vendor_Master.
"""

import os
from datetime import date

import db
import vendor_onboarding as vo  # reuse GSTIN/PAN validators — see module docstring

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

CUSTOMER_MASTER_SHEET = "Customer_Master"
CUSTOMER_DOCS_SHEET = "Customer_Documents"

CUSTOMER_MASTER_COLS = ["Customer_ID", "Customer_Name", "Customer_Type", "Geolocation",
                        "City", "Country", "Address", "Contact_Name", "Contact_Email",
                        "Contact_Phone", "GSTIN", "PAN", "Credit_Limit", "Credit_Status",
                        "Payment_Terms", "Onboarding_Status", "KYC_Flag", "Onboarded_Date", "Active"]
CUSTOMER_DOCS_COLS = ["Document_ID", "Customer_ID", "Doc_Type", "Filename",
                      "Uploaded_Date", "Status", "Notes"]

# Fallback if Customer_Types has nothing seeded yet (fresh install before
# any seed data is loaded) — same values that used to be the only option,
# hardcoded directly here. Now the real source of truth is the
# Customer_Types table; list_customer_types() below is what the UI
# actually reads from.
_CUSTOMER_TYPES_FALLBACK = ["Municipal Corporation", "Water & Sewerage Board",
                            "Government / PSU", "Industrial Plant",
                            "Sanitation Contractor", "Distributor", "Other"]


def list_customer_types(data_file=None):
    """Customer_Types reference list — what Customer Onboarding's
    dropdown reads from, instead of a hardcoded taxonomy baked into the
    UI code. Falls back to the pre-this-feature hardcoded list if the
    table is empty, so behavior doesn't change for a fresh install that
    hasn't seeded reference data yet."""
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT customer_type FROM customer_types ORDER BY customer_type").fetchall()
    finally:
        conn.close()
    types = [r["customer_type"] for r in rows]
    return types or list(_CUSTOMER_TYPES_FALLBACK)


# ── Sheet bootstrap ────────────────────────────────────────────────────────────
def ensure_sheets(wb=None):
    """Kept for signature compatibility — Customer_Master/Customer_Documents
    no longer live in the Excel workbook. `wb` is accepted and ignored."""
    db.init_schema()


# ── Read ──────────────────────────────────────────────────────────────────────
_COL_MAP = {  # SQLite column -> the Capitalized key shape callers already expect
    "customer_id": "Customer_ID", "customer_name": "Customer_Name",
    "customer_type": "Customer_Type", "geolocation": "Geolocation", "city": "City",
    "country": "Country", "address": "Address", "contact_name": "Contact_Name",
    "contact_email": "Contact_Email", "contact_phone": "Contact_Phone", "gstin": "GSTIN",
    "pan": "PAN", "credit_limit": "Credit_Limit", "credit_status": "Credit_Status",
    "payment_terms": "Payment_Terms", "onboarding_status": "Onboarding_Status",
    "kyc_flag": "KYC_Flag", "onboarded_date": "Onboarded_Date", "active": "Active",
    "default_delivery_location": "Default_Delivery_Location",
}
_REV_COL_MAP = {v: k for k, v in _COL_MAP.items()}


def list_customers(data_file=None, include_inactive=True):
    conn = db.get_connection()
    try:
        if include_inactive:
            rows = conn.execute("SELECT * FROM customer_master ORDER BY customer_id").fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM customer_master WHERE active = 'Yes' ORDER BY customer_id"
            ).fetchall()
    finally:
        conn.close()
    return [{_COL_MAP[k]: v for k, v in dict(r).items() if k in _COL_MAP} for r in rows]


def get_customer(customer_id, data_file=None):
    conn = db.get_connection()
    try:
        r = conn.execute(
            "SELECT * FROM customer_master WHERE customer_id = ?", (customer_id,)
        ).fetchone()
    finally:
        conn.close()
    if not r:
        return None
    return {_COL_MAP[k]: v for k, v in dict(r).items() if k in _COL_MAP}


# ── Write ─────────────────────────────────────────────────────────────────────
def upsert_customer(customer_id, fields, data_file=None):
    """
    fields: dict with any of Customer_Name, Customer_Type, Geolocation, City,
    Country, Address, Contact_Name, Contact_Email, Contact_Phone, GSTIN, PAN,
    Credit_Limit, Payment_Terms, Default_Delivery_Location. Runs the same GSTIN/PAN checks vendor intake
    uses and sets Onboarding_Status/KYC_Flag accordingly. New customers save
    Active='No' and Credit_Status='Not Set' until approved — see
    approve_customer(). Returns {checks, onboarding_status, kyc_flag, is_new}.
    (No longer returns an Excel row number — nothing downstream used it.)

    Onboarding_Status/KYC_Flag/Onboarded_Date are only (re)computed for a
    brand new customer, or when the caller is actually submitting GSTIN
    or PAN for (re)validation — never as a side effect of an unrelated
    partial edit. Same bug, same fix as vendor_onboarding.upsert_vendor()
    — confirmed there first: calling this with e.g. just
    {"Credit_Limit": 500000} on an already-Approved customer used to
    silently reset it to "Needs Review" / "Failed Validation", because
    those fields were recomputed and overwritten on every single call
    regardless of what was actually being changed.
    """
    db.init_schema()
    checks = {}
    if "GSTIN" in fields:
        ok, msg, _ = vo.validate_gstin(fields["GSTIN"])
        checks["GSTIN"] = {"ok": ok, "message": msg}
    if "PAN" in fields:
        ok, msg = vo.validate_pan_format(fields["PAN"])
        checks["PAN"] = {"ok": ok, "message": msg}

    all_ok = all(c["ok"] for c in checks.values()) if checks else False
    onboarding_status = "Format Verified" if all_ok else "Needs Review"
    kyc_flag = "Pending Approval" if all_ok else "Failed Validation"

    conn = db.get_connection()
    try:
        existing = conn.execute(
            "SELECT customer_id FROM customer_master WHERE customer_id = ?", (customer_id,)
        ).fetchone()
        is_new = existing is None
        revalidating = bool(checks)

        write_vals = dict(fields)
        if is_new or revalidating:
            write_vals["Onboarding_Status"] = onboarding_status
            write_vals["KYC_Flag"] = kyc_flag
        if is_new:
            write_vals["Onboarded_Date"] = write_vals.get("Onboarded_Date") or date.today().strftime("%Y-%m-%d")
            write_vals.setdefault("Active", "No")
            write_vals.setdefault("Credit_Status", "Not Set")

        db_vals = {_REV_COL_MAP[k]: v for k, v in write_vals.items() if k in _REV_COL_MAP}

        if is_new:
            cols = ["customer_id"] + list(db_vals.keys())
            placeholders = ",".join("?" for _ in cols)
            conn.execute(
                f"INSERT INTO customer_master ({', '.join(cols)}) VALUES ({placeholders})",
                [customer_id] + list(db_vals.values()),
            )
        elif db_vals:
            set_clause = ", ".join(f"{col} = ?" for col in db_vals)
            conn.execute(
                f"UPDATE customer_master SET {set_clause} WHERE customer_id = ?",
                list(db_vals.values()) + [customer_id],
            )
        conn.commit()
    finally:
        conn.close()

    return {"checks": checks, "onboarding_status": onboarding_status,
            "kyc_flag": kyc_flag, "is_new": is_new}


def approve_customer(customer_id, data_file=None):
    """
    Flips a format-verified customer to Active + Approved — the KYC gate,
    same as vendor approval. Credit_Status is set here too: 'Active' if a
    Credit_Limit is already on file, otherwise 'On Hold' — KYC approval and
    credit approval are related but not the same decision, so a customer
    with no stated credit limit stays held even once onboarded.
    """
    conn = db.get_connection()
    try:
        row = conn.execute(
            "SELECT credit_limit FROM customer_master WHERE customer_id = ?", (customer_id,)
        ).fetchone()
        found = row is not None
        if found:
            limit = row["credit_limit"]
            credit_status = "Active" if limit and float(limit) > 0 else "On Hold"
            conn.execute(
                "UPDATE customer_master SET active='Yes', onboarding_status='Approved', "
                "kyc_flag='Approved', credit_status=? WHERE customer_id=?",
                (credit_status, customer_id),
            )
            conn.commit()
    finally:
        conn.close()
    return found


# ── Credit management ────────────────────────────────────────────────────────
def set_credit_limit(customer_id, limit, data_file=None):
    """Updates the credit limit. If the customer was On Hold purely for lack
    of a limit, setting a positive one releases the hold automatically."""
    conn = db.get_connection()
    try:
        row = conn.execute(
            "SELECT credit_status, onboarding_status FROM customer_master WHERE customer_id = ?",
            (customer_id,),
        ).fetchone()
        found = row is not None
        if found:
            new_status = row["credit_status"]
            is_approved = row["onboarding_status"] == "Approved"
            if is_approved and new_status in ("On Hold", "Not Set") and limit and float(limit) > 0:
                new_status = "Active"
            conn.execute(
                "UPDATE customer_master SET credit_limit = ?, credit_status = ? WHERE customer_id = ?",
                (limit, new_status, customer_id),
            )
            conn.commit()
    finally:
        conn.close()
    return found


def hold_customer(customer_id, reason="", data_file=None):
    """Manual credit hold — e.g. delinquent payment. Independent of KYC/
    Onboarding_Status: a held customer is still a legitimate onboarded
    customer, just currently blocked from new orders."""
    conn = db.get_connection()
    try:
        conn.execute(
            "UPDATE customer_master SET credit_status = 'On Hold' WHERE customer_id = ?",
            (customer_id,),
        )
        conn.commit()
    finally:
        conn.close()


def release_hold(customer_id, data_file=None):
    """Lifts a manual hold back to Active. Refuses if there's no credit
    limit on file — releasing a hold shouldn't silently grant unlimited
    credit to a customer who was never actually given a limit."""
    conn = db.get_connection()
    try:
        row = conn.execute(
            "SELECT credit_limit FROM customer_master WHERE customer_id = ?", (customer_id,)
        ).fetchone()
        if row is not None:
            limit = row["credit_limit"]
            if not limit or float(limit) <= 0:
                raise ValueError(f"{customer_id} has no credit limit on file — "
                                  "set one before releasing the hold.")
            conn.execute(
                "UPDATE customer_master SET credit_status = 'Active' WHERE customer_id = ?",
                (customer_id,),
            )
            conn.commit()
    finally:
        conn.close()


def check_credit_available(customer_id, order_value, data_file=None):
    """
    Whether a proposed order fits the customer's credit limit.
    HONEST LIMITATION: this can only check order_value against the stated
    Credit_Limit — it has no outstanding-AR figure to net against, because
    Billing/Cash Application don't exist yet. Once they do, the natural
    extension is (existing_open_AR + order_value) <= Credit_Limit, and the
    natural integration point is the same dependency-injection pattern
    contracts.py uses on pr_consolidation.py: Sales Order Management takes
    an optional credit_check_fn, this module never needs to import it.
    """
    c = get_customer(customer_id, data_file)
    if c is None:
        return {"ok": False, "reason": "Customer not found"}
    if str(c.get("Active") or "").lower() != "yes":
        return {"ok": False, "reason": "Customer not approved/active"}
    if c.get("Credit_Status") != "Active":
        return {"ok": False, "reason": f"Credit status is '{c.get('Credit_Status')}', not Active"}
    limit = c.get("Credit_Limit")
    if not limit or float(limit) <= 0:
        return {"ok": False, "reason": "No credit limit on file"}
    if order_value > float(limit):
        return {"ok": False, "reason": f"Order \u20b9{order_value:,.2f} exceeds credit limit "
                 f"\u20b9{float(limit):,.2f} (not yet netted against open AR — see docstring)"}
    return {"ok": True, "reason": "Within credit limit", "limit": float(limit)}


# ── Documents (reference only, mirrors vendor_onboarding.py) ────────────────────
def _next_doc_id(conn):
    rows = conn.execute(
        "SELECT document_id FROM customer_documents WHERE document_id LIKE 'CDOC-%'"
    ).fetchall()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["document_id"].split("-")[1]))
        except Exception: pass
    return f"CDOC-{mx+1:05d}"


def record_document(customer_id, doc_type, filename, notes="", data_file=None):
    db.init_schema()
    conn = db.get_connection()
    try:
        did = _next_doc_id(conn)
        conn.execute(
            "INSERT INTO customer_documents (document_id, customer_id, doc_type, filename, "
            "uploaded_date, status, notes) VALUES (?,?,?,?,?,?,?)",
            (did, customer_id, doc_type, filename, date.today().strftime("%Y-%m-%d"),
             "Logged (not yet reviewed)", notes),
        )
        conn.commit()
    finally:
        conn.close()
    return did


def get_documents(customer_id=None, data_file=None):
    conn = db.get_connection()
    try:
        if customer_id:
            rows = conn.execute(
                "SELECT * FROM customer_documents WHERE customer_id = ? ORDER BY document_id",
                (customer_id,),
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM customer_documents ORDER BY document_id").fetchall()
    finally:
        conn.close()
    return [{"doc_id": r["document_id"], "customer_id": r["customer_id"], "doc_type": r["doc_type"],
             "filename": r["filename"], "uploaded_date": r["uploaded_date"],
             "status": r["status"], "notes": r["notes"]} for r in rows]


def stats(data_file=None):
    customers = list_customers(data_file)
    by_status = {}
    by_credit = {}
    for c in customers:
        s = c.get("Onboarding_Status") or "Draft"
        by_status[s] = by_status.get(s, 0) + 1
        cs = c.get("Credit_Status") or "Not Set"
        by_credit[cs] = by_credit.get(cs, 0) + 1
    approved = sum(1 for c in customers if str(c.get("Active") or "").lower() == "yes")
    return {"total": len(customers), "approved": approved,
            "by_status": by_status, "by_credit": by_credit}


if __name__ == "__main__":
    ok, msg, _ = vo.validate_gstin("27AABCU9603R1ZN")
    print(f"Reused GSTIN validator self-test: {'PASS' if ok else 'FAIL'} — {msg}")
    print("Customer stats:", stats())
