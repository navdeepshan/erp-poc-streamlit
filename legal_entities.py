"""
legal_entities.py — Legal entity master data.

Previously "Legal Entity" was just a free-text field (typically "LE-001")
scattered through PR/PO creation with no master table behind it at all —
any string was accepted, nothing validated it meant anything real. This
gives it a real reference table, even though — per current scope — a
single-legal-entity business like this pilot's Genrobotics setup only
ever needs exactly one row, with the same identity details as the org
profile itself (org_profile.py's Org_ID and this module's one LE_ID
happen to describe the same company from two different angles: "who we
are" vs. "which legal entity are we transacting as").

Deliberately minimal — a read/list/upsert module, no dedicated
management UI yet. If a business ever needs more than one legal entity
(a holding company with multiple registered subsidiaries, for
instance), this table is already shaped for that; nothing here assumes
exactly one row, it's just what's populated today.

New table:
  Legal_Entities   LE_ID | LE_Name | GSTIN | PAN | Address | City |
                   State | Country | Bank_Account_No | IFSC |
                   Bank_Name | Contact_Email | Contact_Phone
"""

import os

import db
import vendor_onboarding as vo  # reuse GSTIN/PAN validators

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

_COL_MAP = {
    "le_id": "LE_ID", "le_name": "LE_Name", "gstin": "GSTIN", "pan": "PAN",
    "address": "Address", "city": "City", "state": "State", "country": "Country",
    "bank_account_no": "Bank_Account_No", "ifsc": "IFSC", "bank_name": "Bank_Name",
    "contact_email": "Contact_Email", "contact_phone": "Contact_Phone",
}
_REV_COL_MAP = {v: k for k, v in _COL_MAP.items()}


def ensure_sheets(wb=None):
    db.init_schema()


def list_legal_entities(data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT * FROM legal_entities ORDER BY le_id").fetchall()
    finally:
        conn.close()
    return [{_COL_MAP[k]: v for k, v in dict(r).items() if k in _COL_MAP} for r in rows]


def get_legal_entity(le_id, data_file=None):
    for le in list_legal_entities(data_file):
        if le["LE_ID"] == le_id:
            return le
    return None


def upsert_legal_entity(le_id, fields, data_file=None):
    """fields: any of LE_Name, GSTIN, PAN, Address, City, State, Country,
    Bank_Account_No, IFSC, Bank_Name, Contact_Email, Contact_Phone."""
    checks = {}
    if fields.get("GSTIN"):
        ok, msg, _ = vo.validate_gstin(fields["GSTIN"])
        checks["GSTIN"] = {"ok": ok, "message": msg}
    if fields.get("PAN"):
        ok, msg = vo.validate_pan_format(fields["PAN"])
        checks["PAN"] = {"ok": ok, "message": msg}

    db.init_schema()
    conn = db.get_connection()
    try:
        existing = conn.execute(
            "SELECT le_id FROM legal_entities WHERE le_id = ?", (le_id,)
        ).fetchone()
        db_vals = {_REV_COL_MAP[k]: v for k, v in fields.items() if k in _REV_COL_MAP}
        if existing is None:
            cols = ["le_id"] + list(db_vals.keys())
            placeholders = ",".join("?" for _ in cols)
            conn.execute(
                f"INSERT INTO legal_entities ({', '.join(cols)}) VALUES ({placeholders})",
                [le_id] + list(db_vals.values()),
            )
        else:
            set_clause = ", ".join(f"{col} = ?" for col in db_vals)
            conn.execute(
                f"UPDATE legal_entities SET {set_clause} WHERE le_id = ?",
                list(db_vals.values()) + [le_id],
            )
        conn.commit()
    finally:
        conn.close()
    return {"checks": checks}


if __name__ == "__main__":
    print("Legal entities:", list_legal_entities())
