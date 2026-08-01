"""
vendor_onboarding.py — Vendor intake & lightweight KYC scaffolding.

v1 scope, deliberately kept lean:
  - Local, zero-cost, zero-dependency format + checksum validation for
    GSTIN and PAN (same "no external API" philosophy as categorization.py).
    This catches typos instantly and for free, before anyone pays for a
    real GSTIN-active lookup.
  - A live IFSC lookup against Razorpay's free, keyless public API
    (https://ifsc.razorpay.com) to auto-fill bank name/branch as the
    vendor types an IFSC code. Requires outbound internet access from
    wherever this app runs — it could not be tested from Claude's sandboxed
    dev environment (that network is allow-listed and doesn't include this
    domain), but the endpoint itself needs no API key and no signup, so it
    should work as-is once you run the app locally. If it doesn't, the form
    still works with manual bank-detail entry.
  - Document *references* only (type, filename, notes) are logged to a new
    Vendor_Documents sheet — actual upload + OCR extraction is intentionally
    not built yet; that's the natural next slice.

Gating: a newly-added vendor is saved with Active = "No" until someone
explicitly approves them. RFx's vendor suggestion logic already filters on
Active == "Yes", so unapproved vendors are automatically excluded from
RFQ suggestions without any change needed there — the gate is "free".

Vendor_Master and Vendor_Documents now live in SQLite (erp_pilot.db,
tables `vendor_master` and `vendor_documents`) — this module is the
canonical owner both pr_consolidation.py and rfx.py delegate to for
vendor geo/lookup data, closing a real inconsistency that predated this
migration: pr_consolidation.py's own vendor reader didn't filter on
Active at all, while rfx.py's separate copy did — two independently
written "read Vendor_Master" implementations that could (and did)
disagree. Now there's exactly one.
"""

import os, re, json
from datetime import date
from urllib import request as urlrequest, error as urlerror

import db

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

VENDOR_MASTER_SHEET = "Vendor_Master"
VENDOR_DOCS_SHEET = "Vendor_Documents"

VENDOR_MASTER_ALL_COLS = ["Vendor_ID", "Vendor_Name", "Geolocation", "City", "Country",
                          "Address", "Contact_Name", "Contact_Email", "Active",
                          "GSTIN", "PAN", "Bank_Account_No", "IFSC", "Bank_Name",
                          "Bank_Branch", "Onboarding_Status", "KYC_Flag", "Onboarded_Date"]


def ensure_schema(wb=None):
    """Kept for signature compatibility with any external caller — no
    longer touches an Excel workbook (Vendor_Master/Vendor_Documents
    live in SQLite now). `wb` is accepted and ignored."""
    db.init_schema()


# ── Local validators (no network, no cost) ──────────────────────────────────────
_B36 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
_PAN_RE = re.compile(r'^[A-Z]{5}[0-9]{4}[A-Z]$')
_GSTIN_RE = re.compile(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$')
_IFSC_RE = re.compile(r'^[A-Z]{4}0[A-Z0-9]{6}$')

# Census 2011 GST state codes, the common ones — used only to make the
# "unknown state code" warning friendlier, not to hard-block anything.
_GST_STATE_CODES = {
    "01":"Jammu & Kashmir","02":"Himachal Pradesh","03":"Punjab","04":"Chandigarh",
    "05":"Uttarakhand","06":"Haryana","07":"Delhi","08":"Rajasthan","09":"Uttar Pradesh",
    "10":"Bihar","11":"Sikkim","12":"Arunachal Pradesh","13":"Nagaland","14":"Manipur",
    "15":"Mizoram","16":"Tripura","17":"Meghalaya","18":"Assam","19":"West Bengal",
    "20":"Jharkhand","21":"Odisha","22":"Chhattisgarh","23":"Madhya Pradesh","24":"Gujarat",
    "27":"Maharashtra","29":"Karnataka","30":"Goa","31":"Lakshadweep","32":"Kerala",
    "33":"Tamil Nadu","34":"Puducherry","35":"Andaman & Nicobar","36":"Telangana",
    "37":"Andhra Pradesh","38":"Ladakh",
}


def compute_gstin_check_digit(gstin14):
    """Public: compute the 15th (checksum) character for a 14-char GSTIN prefix.
    Exposed so other modules (e.g. vrq.py's demo data) can build genuinely
    checksum-valid dummy GSTINs instead of arbitrary-looking fake ones."""
    factor = 1
    total = 0
    for ch in gstin14.upper():
        idx = _B36.find(ch)
        d = idx * factor
        d = (d // 36) + (d % 36)
        total += d
        factor = 2 if factor == 1 else 1
    return _B36[(36 - (total % 36)) % 36]


def demo_gstin(state_code="32", pan="AABCU9603R", entity="1"):
    """A structurally + checksum valid dummy GSTIN for demo/simulation use."""
    prefix = f"{state_code}{pan.upper()}{entity}Z"
    return prefix + compute_gstin_check_digit(prefix)


def validate_pan_format(pan):
    """Format-only check (5 letters, 4 digits, 1 letter). PAN's own check
    letter uses an unpublished algorithm, so this can't verify checksum —
    only that the shape is right."""
    pan = (pan or "").strip().upper()
    if not pan:
        return False, "PAN is required"
    if not _PAN_RE.match(pan):
        return False, "Invalid PAN format — expected AAAAA9999A"
    holder_types = {"P":"Individual","C":"Company","H":"HUF","F":"Firm/LLP","A":"AOP",
                     "T":"Trust","B":"BOI","L":"Local Authority","J":"Artificial Judicial Person",
                     "G":"Government"}
    holder = holder_types.get(pan[3], "Unknown")
    return True, f"Format valid ({holder})"


def _gstin_checksum_ok(gstin15):
    return compute_gstin_check_digit(gstin15[:-1]) == gstin15[-1]


def validate_gstin(gstin):
    """Format + Luhn-mod-36 checksum validation, plus a structural check that
    the embedded PAN segment (chars 3-12) is itself PAN-shaped. This is a
    strong typo-catcher; it does NOT confirm the registration is active —
    that needs a real GSTN lookup (e.g. via a Sandbox.co.in-style API)."""
    gstin = (gstin or "").strip().upper()
    if not gstin:
        return False, "GSTIN is required", {}
    if len(gstin) != 15:
        return False, f"GSTIN must be 15 characters (got {len(gstin)})", {}
    if not _GSTIN_RE.match(gstin):
        return False, "Invalid GSTIN format", {}
    if not _gstin_checksum_ok(gstin):
        return False, "GSTIN checksum failed — likely a typo", {}
    embedded_pan = gstin[2:12]
    pan_ok, pan_msg = validate_pan_format(embedded_pan)
    if not pan_ok:
        return False, "Embedded PAN segment is malformed", {}
    state_code = gstin[:2]
    state_name = _GST_STATE_CODES.get(state_code, "Unrecognized state code")
    return True, f"Format + checksum valid ({state_name})", {
        "state_code": state_code, "state_name": state_name, "embedded_pan": embedded_pan,
    }


def lookup_ifsc(ifsc_code, timeout=5):
    """Live lookup via Razorpay's free, keyless public IFSC API. Returns a
    dict; always has 'ok'. Never raises — network/format problems come back
    as ok=False with an 'error' message so the UI can fall back to manual entry."""
    ifsc_code = (ifsc_code or "").strip().upper()
    if not _IFSC_RE.match(ifsc_code):
        return {"ok": False, "error": "Invalid IFSC format — expected e.g. HDFC0001234"}
    url = f"https://ifsc.razorpay.com/{ifsc_code}"
    try:
        req = urlrequest.Request(url, headers={"User-Agent": "vendor-onboarding-poc"})
        with urlrequest.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode())
        return {"ok": True, "bank": data.get("BANK"), "branch": data.get("BRANCH"),
                "address": data.get("ADDRESS"), "city": data.get("CITY"),
                "state": data.get("STATE")}
    except urlerror.HTTPError as e:
        if e.code == 404:
            return {"ok": False, "error": "IFSC code not found"}
        return {"ok": False, "error": f"Lookup failed (HTTP {e.code})"}
    except Exception as e:
        return {"ok": False, "error": f"Network error — {e}"}


# ── Vendor_Master read/write ─────────────────────────────────────────────────────
_COL_MAP = {  # SQLite column -> the Capitalized key shape callers already expect
    "vendor_id": "Vendor_ID", "vendor_name": "Vendor_Name", "geolocation": "Geolocation",
    "city": "City", "country": "Country", "address": "Address",
    "contact_name": "Contact_Name", "contact_email": "Contact_Email", "active": "Active",
    "gstin": "GSTIN", "pan": "PAN", "bank_account_no": "Bank_Account_No", "ifsc": "IFSC",
    "bank_name": "Bank_Name", "bank_branch": "Bank_Branch",
    "onboarding_status": "Onboarding_Status", "kyc_flag": "KYC_Flag",
    "onboarded_date": "Onboarded_Date", "vendor_type": "Vendor_Type",
}
_REV_COL_MAP = {v: k for k, v in _COL_MAP.items()}


def list_vendor_types(data_file=None):
    """Vendor_Types reference list — what the Vendor Onboarding UI's
    dropdown reads from, instead of a hardcoded taxonomy baked into the
    UI code. Empty until seeded (see migrate_org_reference_data.py)."""
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT vendor_type FROM vendor_types ORDER BY vendor_type").fetchall()
    finally:
        conn.close()
    return [r["vendor_type"] for r in rows]


def list_vendors(data_file=None, include_inactive=True):
    """Vendor_Master now lives in SQLite — the canonical reader every
    other module (pr_consolidation.py, rfx.py) should delegate to for
    vendor data, rather than each maintaining its own read. Returns the
    same Capitalized-key dict shape as before (Vendor_ID, Vendor_Name,
    ...) so existing callers didn't need to change."""
    conn = db.get_connection()
    try:
        if include_inactive:
            rows = conn.execute("SELECT * FROM vendor_master ORDER BY vendor_id").fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM vendor_master WHERE active = 'Yes' ORDER BY vendor_id"
            ).fetchall()
    finally:
        conn.close()
    return [{_COL_MAP[k]: v for k, v in dict(r).items() if k in _COL_MAP} for r in rows]


def get_vendor(vendor_id, data_file=None):
    conn = db.get_connection()
    try:
        r = conn.execute(
            "SELECT * FROM vendor_master WHERE vendor_id = ?", (vendor_id,)
        ).fetchone()
    finally:
        conn.close()
    if not r:
        return None
    return {_COL_MAP[k]: v for k, v in dict(r).items() if k in _COL_MAP}


def upsert_vendor(vendor_id, fields, data_file=None):
    """
    fields: dict with any of Vendor_Name, Vendor_Type, Geolocation, City,
    Country, Address, Contact_Name, Contact_Email, GSTIN, PAN,
    Bank_Account_No, IFSC, Bank_Name, Bank_Branch. Runs local validation
    on GSTIN/PAN and sets Onboarding_Status + KYC_Flag accordingly. New
    vendors are saved Active='No' until approved.
    Returns {checks: {...}, onboarding_status, kyc_flag, is_new}. (No longer
    returns a Excel row number — `row` dropped since SQLite rows have no
    stable position; nothing downstream used it besides logging.)

    Onboarding_Status/KYC_Flag are only (re)computed for a brand new
    vendor, or when the caller is actually submitting GSTIN or PAN for
    (re)validation — never as a side effect of an unrelated partial
    edit. This was a real, confirmed bug: calling this with e.g. just
    {"Vendor_Type": "..."} on an already-Approved vendor silently reset
    it to "Needs Review" / "Failed Validation", because those two
    fields used to be recomputed and overwritten on every single call
    regardless of what was actually being changed. Same reasoning for
    Onboarded_Date — it should only be set once, at genuine first
    onboarding, not re-stamped to today() on every later edit.
    """
    db.init_schema()
    checks = {}
    if "GSTIN" in fields:
        ok, msg, _ = validate_gstin(fields["GSTIN"])
        checks["GSTIN"] = {"ok": ok, "message": msg}
    if "PAN" in fields:
        ok, msg = validate_pan_format(fields["PAN"])
        checks["PAN"] = {"ok": ok, "message": msg}

    all_ok = all(c["ok"] for c in checks.values()) if checks else False
    onboarding_status = "Format Verified" if all_ok else "Needs Review"
    kyc_flag = "Pending Approval" if all_ok else "Failed Validation"

    conn = db.get_connection()
    try:
        existing = conn.execute(
            "SELECT vendor_id FROM vendor_master WHERE vendor_id = ?", (vendor_id,)
        ).fetchone()
        is_new = existing is None
        revalidating = bool(checks)  # caller included GSTIN and/or PAN this call

        write_vals = dict(fields)
        if is_new or revalidating:
            write_vals["Onboarding_Status"] = onboarding_status
            write_vals["KYC_Flag"] = kyc_flag
        if is_new:
            write_vals["Onboarded_Date"] = write_vals.get("Onboarded_Date") or date.today().strftime("%Y-%m-%d")
            write_vals.setdefault("Active", "No")  # gate: not suggested until approved

        db_vals = {_REV_COL_MAP[k]: v for k, v in write_vals.items() if k in _REV_COL_MAP}

        if is_new:
            cols = ["vendor_id"] + list(db_vals.keys())
            placeholders = ",".join("?" for _ in cols)
            conn.execute(
                f"INSERT INTO vendor_master ({', '.join(cols)}) VALUES ({placeholders})",
                [vendor_id] + list(db_vals.values()),
            )
        elif db_vals:
            set_clause = ", ".join(f"{col} = ?" for col in db_vals)
            conn.execute(
                f"UPDATE vendor_master SET {set_clause} WHERE vendor_id = ?",
                list(db_vals.values()) + [vendor_id],
            )
        conn.commit()
    finally:
        conn.close()

    return {"checks": checks, "onboarding_status": onboarding_status,
            "kyc_flag": kyc_flag, "is_new": is_new}


def approve_vendor(vendor_id, data_file=None):
    """Flip a format-verified vendor to Active + Approved, making it eligible
    for RFx vendor suggestions (which filter on Active == 'Yes')."""
    conn = db.get_connection()
    try:
        cur = conn.execute(
            "UPDATE vendor_master SET active='Yes', onboarding_status='Approved', "
            "kyc_flag='Approved' WHERE vendor_id = ?",
            (vendor_id,),
        )
        found = cur.rowcount > 0
        conn.commit()
    finally:
        conn.close()
    return found


# ── Vendor_Documents (references only — no upload/OCR yet) ──────────────────────
def _next_doc_id(conn):
    rows = conn.execute(
        "SELECT document_id FROM vendor_documents WHERE document_id LIKE 'DOC-%'"
    ).fetchall()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["document_id"].split("-")[1]))
        except Exception: pass
    return f"DOC-{mx+1:05d}"


def record_document(vendor_id, doc_type, filename, notes="", data_file=None):
    db.init_schema()
    conn = db.get_connection()
    try:
        did = _next_doc_id(conn)
        conn.execute(
            "INSERT INTO vendor_documents (document_id, vendor_id, doc_type, filename, "
            "uploaded_date, status, notes) VALUES (?,?,?,?,?,?,?)",
            (did, vendor_id, doc_type, filename, date.today().strftime("%Y-%m-%d"),
             "Logged (not yet reviewed)", notes),
        )
        conn.commit()
    finally:
        conn.close()
    return did


def get_documents(vendor_id=None, data_file=None):
    conn = db.get_connection()
    try:
        if vendor_id:
            rows = conn.execute(
                "SELECT * FROM vendor_documents WHERE vendor_id = ? ORDER BY document_id",
                (vendor_id,),
            ).fetchall()
        else:
            rows = conn.execute("SELECT * FROM vendor_documents ORDER BY document_id").fetchall()
    finally:
        conn.close()
    return [{"doc_id": r["document_id"], "vendor_id": r["vendor_id"], "doc_type": r["doc_type"],
             "filename": r["filename"], "uploaded_date": r["uploaded_date"],
             "status": r["status"], "notes": r["notes"]} for r in rows]


def stats(data_file=None):
    vendors = list_vendors(data_file)
    by_status = {}
    for v in vendors:
        s = v.get("Onboarding_Status") or "Draft"
        by_status[s] = by_status.get(s, 0) + 1
    approved = sum(1 for v in vendors if str(v.get("Active") or "").lower() == "yes")
    return {"total": len(vendors), "approved": approved, "by_status": by_status}


if __name__ == "__main__":
    # Self-test the checksum against a known-valid published GSTIN.
    ok, msg, details = validate_gstin("27AABCU9603R1ZN")
    print(f"GSTIN self-test: {'PASS' if ok else 'FAIL'} — {msg}")
    ok2, msg2 = validate_pan_format("AABCU9603R")
    print(f"PAN self-test:   {'PASS' if ok2 else 'FAIL'} — {msg2}")
    bad_ok, bad_msg, _ = validate_gstin("27AABCU9603R1ZX")  # wrong check digit
    print(f"Negative test:   {'PASS (correctly rejected)' if not bad_ok else 'FAIL (should have been rejected)'} — {bad_msg}")
    print()
    print("Current vendors:", stats())
