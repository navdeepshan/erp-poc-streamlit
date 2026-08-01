"""
org_profile.py — Your own company's identity (O2C, prerequisite for Billing).

Every invoice legally needs to say who's billing, not just who's being
billed. Nothing in this app has ever stored that — pr_consolidation.py has
always used hardcoded placeholder org defaults ("LE-001", "PLANT-01")
because nothing downstream needed them to be real. Billing does, so this
exists now, checked as missing before being built rather than assumed.

Deliberately a single-record settings module, not a full master-data list
— reuses vendor_onboarding.py's GSTIN/PAN validators, same as every other
identity in this app. If a business needs multiple GST registrations
(India requires a separate GSTIN per state a business operates in), that's
a natural extension of ORG_COLS into a real list later; not built now
because nothing yet needs more than one.

billing.py checks is_configured() before allowing an invoice to be
created — an invoice literally cannot be generated without a real seller
GSTIN on file. That's the enforcement point for "don't generate something
that looks compliant and isn't."

Org_ID is a real, user-editable business code (e.g. "GRB"), not just an
internal row identifier — this used to be hardcoded to a literal
"DEFAULT" with no way to see or change it in the UI. It's still a
technical singleton underneath (there's only ever one row, exactly one
org profile), but every read/write now operates on "whichever single
row currently exists" rather than a lookup hardcoded to that old
constant, so the ID itself is free to be whatever the user sets — and
changing it doesn't create a second orphaned row or break the lookup.

New sheet:
  Org_Profile   Org_ID | Legal_Name | GSTIN | PAN | Address | City | State |
                Country | Bank_Account_No | IFSC | Bank_Name |
                Contact_Email | Contact_Phone

SQLite pilot: Org_Profile now lives in erp_pilot.db (table `org_profile`).
"""

import os

import db
import vendor_onboarding as vo  # reuse GSTIN/PAN validators

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

DEFAULT_ORG_ID = "DEFAULT"  # only used as the fallback ID for a brand-new
                            # profile that hasn't been given a real one yet

_COL_MAP = {
    "org_id": "Org_ID", "legal_name": "Legal_Name", "gstin": "GSTIN", "pan": "PAN",
    "address": "Address", "city": "City", "state": "State", "country": "Country",
    "bank_account_no": "Bank_Account_No", "ifsc": "IFSC", "bank_name": "Bank_Name",
    "contact_email": "Contact_Email", "contact_phone": "Contact_Phone",
}
_REV_COL_MAP = {v: k for k, v in _COL_MAP.items()}


def ensure_sheets(wb=None):
    """Kept for signature compatibility — Org_Profile no longer lives in
    the Excel workbook. `wb` is accepted and ignored."""
    db.init_schema()


def get_org_profile(data_file=None):
    """Whichever single row currently exists — not a lookup keyed to a
    hardcoded ID, since Org_ID is now a real, user-editable value (see
    module docstring). There's still only ever one row; LIMIT 1 is a
    safety net, not evidence more than one is expected."""
    conn = db.get_connection()
    try:
        r = conn.execute("SELECT * FROM org_profile LIMIT 1").fetchone()
    finally:
        conn.close()
    if not r:
        return None
    return {_COL_MAP[k]: v for k, v in dict(r).items() if k in _COL_MAP}


def is_configured(data_file=None):
    """False until a real GSTIN is on file — the gate billing.py checks."""
    profile = get_org_profile(data_file)
    if profile is None or not profile.get("GSTIN"):
        return False
    ok, _, _ = vo.validate_gstin(profile["GSTIN"])
    return ok


def set_org_profile(fields, data_file=None):
    """
    fields: any of Org_ID, Legal_Name, GSTIN, PAN, Address, City, State,
    Country, Bank_Account_No, IFSC, Bank_Name, Contact_Email, Contact_Phone.
    Validates GSTIN/PAN with the same checksum logic used everywhere else.
    Returns {checks, is_configured}.

    Org_ID is free-text here deliberately — it's a business code (e.g.
    "GRB"), not validated against any format, the same way a company's
    own short name isn't a thing that has a "correct" shape to check.
    Changing it on an existing profile updates that row's ID in place
    (a real SQL UPDATE of the primary key column, which SQLite handles
    fine) rather than inserting a second row and orphaning the first.
    """
    db.init_schema()
    checks = {}
    if "GSTIN" in fields and fields["GSTIN"]:
        ok, msg, _ = vo.validate_gstin(fields["GSTIN"])
        checks["GSTIN"] = {"ok": ok, "message": msg}
    if "PAN" in fields and fields["PAN"]:
        ok, msg = vo.validate_pan_format(fields["PAN"])
        checks["PAN"] = {"ok": ok, "message": msg}

    conn = db.get_connection()
    try:
        existing = conn.execute("SELECT org_id FROM org_profile LIMIT 1").fetchone()
        new_org_id = fields.get("Org_ID") or None
        db_vals = {_REV_COL_MAP[k]: v for k, v in fields.items()
                   if k in _REV_COL_MAP and k != "Org_ID"}

        if existing is None:
            org_id = new_org_id or DEFAULT_ORG_ID
            cols = ["org_id"] + list(db_vals.keys())
            placeholders = ",".join("?" for _ in cols)
            conn.execute(
                f"INSERT INTO org_profile ({', '.join(cols)}) VALUES ({placeholders})",
                [org_id] + list(db_vals.values()),
            )
        else:
            current_id = existing["org_id"]
            if new_org_id and new_org_id != current_id:
                conn.execute("UPDATE org_profile SET org_id = ? WHERE org_id = ?",
                            (new_org_id, current_id))
                current_id = new_org_id
            if db_vals:
                set_clause = ", ".join(f"{col} = ?" for col in db_vals)
                conn.execute(
                    f"UPDATE org_profile SET {set_clause} WHERE org_id = ?",
                    list(db_vals.values()) + [current_id],
                )
        conn.commit()
    finally:
        conn.close()

    return {"checks": checks, "is_configured": is_configured(data_file)}


if __name__ == "__main__":
    print("Org profile:", get_org_profile())
    print("Is configured:", is_configured())
