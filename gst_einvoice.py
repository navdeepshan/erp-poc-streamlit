"""
gst_einvoice.py — GST e-Invoicing: real IRN generation via the GST
Invoice Registration Portal (IRP), scoped 2026-08-10 as O2C-E09 (see
CONTEXT_HANDOFF_v2.md for the full design writeup).

WHY THIS EXISTS: since Rule 48(4) of the CGST Rules (turnover threshold
now Rs. 5 crore, from Aug 2023), a B2B tax invoice from a taxpayer above
threshold is not legally valid without a government-issued Invoice
Reference Number (IRN) and signed QR code, obtained by registering the
invoice with the IRP *before* handing it to the customer. Nothing in
this codebase attempted that before this module — billing.py computes
correct GST, but an "Issued" invoice here was purely an internal
document.

REAL, NOT SIMULATED — by explicit choice, unlike eway_bill.py's own
deterministic simulation. This module makes a genuine outbound HTTPS
call to NIC's public e-invoice API sandbox (einv-apisandbox.nic.in,
confirmed reachable from this environment). That is a first for this
codebase: every other integration (courier tracking, e-way bills) is a
local, offline simulation. Two real consequences of that:

  1. Credentials are never hardcoded here. They're read from environment
     variables (GST_EINV_CLIENT_ID, GST_EINV_CLIENT_SECRET,
     GST_EINV_USERNAME, GST_EINV_PASSWORD) that must be set outside this
     repo. is_configured() reports whether they are, the same shape as
     org_profile.is_configured() gates billing.py.

  2. Getting those credentials is NOT fully self-serve. Registering on
     the sandbox under the "ERP" category needs only a real PAN, not a
     real client GSTIN (GSPs/ERPs can generate their own dummy GSTINs
     for testing) — but the registration flow requires OTP verification
     to a real phone number, a one-time human action this module cannot
     perform on its own. Until that's done and the four env vars above
     are set, every function below that calls the API fails cleanly with
     a clear "not configured" error — never a silent no-op, never a
     fabricated IRN.

SCOPING DECISION, stated rather than hidden: this module talks to the
IRP using a plain JSON request/response contract (the shape NIC's own
sandbox documentation publishes, and what the large majority of real
integrations use). NIC's raw production API additionally wraps every
request/response body in an AES session-key exchange keyed by an RSA
handshake — real cryptographic complexity that most real-world
integrations avoid entirely by going through a GST Suvidha Provider
(GSP), which terminates that encryption on their own side and exposes
the simpler contract this module implements. If a future production
deployment goes direct-to-NIC instead of through a GSP, that encryption
layer is a real, separate piece of work this module does not attempt.

SCOPE, v1: IRN generation and cancellation only (the per-invoice, real-
time compliance event). GSTR-1 (the periodic, invoice-level return
filing every GST-registered business owes regardless of e-invoicing
threshold) is a genuinely different shape of work — a batch export
formatter, not a live API integration — deliberately deferred to a
follow-on, not built here.

No GL impact anywhere in this module: an IRN is a compliance/document
event, not a financial transaction. The invoice's own GL entry already
posted at Issue time (accounting.post_invoice_entry()); this only
attaches a government registration to an invoice that already exists.

New table:
  e_invoices   Invoice_ID | IRN | Ack_No | Ack_Date | Signed_Invoice |
               Signed_QR_Code | Status | Generated_Date | Generated_By |
               Cancel_Reason | Cancel_Date | Error_Message

SQLite: e_invoices lives in erp_pilot.db, this module is its exclusive
owner — same one-table-one-owner shape as every other module here.
"""

import os
from datetime import date, datetime, timedelta

import requests

import db
import billing as bl
import org_profile as op
import customer_onboarding as co
import vendor_onboarding as vo

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

# Configurable, not hardcoded as gospel — see module docstring's scoping
# decision. These are NIC's own documented sandbox paths; a real
# deployment through a GSP will have entirely different paths and should
# override every one of these via env var, not edit this file.
BASE_URL = os.environ.get("GST_EINV_BASE_URL", "https://einv-apisandbox.nic.in")
AUTH_PATH = os.environ.get("GST_EINV_AUTH_PATH", "/eivital/v1.04/auth")
GENERATE_PATH = os.environ.get("GST_EINV_GENERATE_PATH", "/eicore/v1.03/Invoice")
CANCEL_PATH = os.environ.get("GST_EINV_CANCEL_PATH", "/eicore/v1.03/Invoice/Cancel")

CANCEL_WINDOW_HOURS = 24  # Real IRP rule — an IRN can only be cancelled
                          # within 24 hours of generation; past that, the
                          # only real remedy is a credit note (rma.py's
                          # own issue_credit_memo(), already built).

# In-memory only — a fresh auth token every process start is a fine
# trade for a PoC; a real deployment would want this cached alongside
# the credentials themselves, not in a module-level dict that resets
# on every Streamlit rerun of a long-lived process.
_token_cache = {"token": None, "expires_at": None}


def is_configured():
    """False until all four sandbox credentials are set as real
    environment variables — the gate every API-calling function below
    checks first, same shape as org_profile.is_configured()."""
    return all(os.environ.get(k) for k in
               ("GST_EINV_CLIENT_ID", "GST_EINV_CLIENT_SECRET",
                "GST_EINV_USERNAME", "GST_EINV_PASSWORD"))


def _credentials():
    return {
        "client_id": os.environ.get("GST_EINV_CLIENT_ID"),
        "client_secret": os.environ.get("GST_EINV_CLIENT_SECRET"),
        "username": os.environ.get("GST_EINV_USERNAME"),
        "password": os.environ.get("GST_EINV_PASSWORD"),
    }


# ── Read ──────────────────────────────────────────────────────────────────────
def get_einvoice(invoice_id, data_file=None):
    conn = db.get_connection()
    try:
        r = conn.execute("SELECT * FROM e_invoices WHERE invoice_id = ?", (invoice_id,)).fetchone()
    finally:
        conn.close()
    if not r:
        return None
    return {"invoice_id": r["invoice_id"], "irn": r["irn"], "ack_no": r["ack_no"],
            "ack_date": r["ack_date"], "signed_invoice": r["signed_invoice"],
            "signed_qr_code": r["signed_qr_code"], "status": r["status"],
            "generated_date": r["generated_date"], "generated_by": r["generated_by"],
            "cancel_reason": r["cancel_reason"], "cancel_date": r["cancel_date"],
            "error_message": r["error_message"]}


def is_cancellable(invoice_id, data_file=None):
    """Real 24-hour IRP rule, not a UI-only guess — computed off the
    real generated_date/time on file."""
    einv = get_einvoice(invoice_id, data_file)
    if not einv or einv["status"] != "Generated":
        return False, "No active IRN on file."
    try:
        generated = datetime.strptime(einv["generated_date"], "%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return False, "Generated timestamp on file is unreadable — can't verify the window."
    deadline = generated + timedelta(hours=CANCEL_WINDOW_HOURS)
    if datetime.now() > deadline:
        return False, (f"The {CANCEL_WINDOW_HOURS}-hour cancellation window closed at "
                       f"{deadline.strftime('%Y-%m-%d %H:%M:%S')} — a credit memo "
                       "(rma.py) is the real remedy past this point, not cancellation.")
    return True, ""


# ── Payload construction ─────────────────────────────────────────────────────
def _fmt_date_ddmmyyyy(iso_date):
    return datetime.strptime(iso_date, "%Y-%m-%d").strftime("%d/%m/%Y")


def check_prerequisites(invoice_id, data_file=None):
    """
    Everything that must be real and present before an IRN can even be
    attempted, checked up front and named individually — the same
    gate-and-name-what's-missing discipline billing.create_invoice()
    already uses for HSN/GST rate. Returns (ok, list_of_problems).
    """
    fpath = data_file or DATA_FILE
    problems = []

    if not is_configured():
        problems.append("Sandbox credentials are not configured (GST_EINV_CLIENT_ID / "
                        "GST_EINV_CLIENT_SECRET / GST_EINV_USERNAME / GST_EINV_PASSWORD "
                        "environment variables) — see module docstring for how to obtain them.")

    inv = bl.get_invoice(invoice_id, fpath)
    if inv is None:
        problems.append(f"{invoice_id} not found.")
        return False, problems
    if inv["status"] != "Issued":
        problems.append(f"{invoice_id} is '{inv['status']}' — only an Issued invoice can "
                        "get a real IRN.")

    existing = get_einvoice(invoice_id, fpath)
    if existing and existing["status"] == "Generated":
        problems.append(f"{invoice_id} already has an active IRN on file ({existing['irn']}).")

    org = op.get_org_profile(fpath) or {}
    if not org.get("GSTIN"):
        problems.append("Your organization has no GSTIN on file (Org Profile).")
    elif not vo.validate_gstin(org["GSTIN"])[0]:
        problems.append("Your organization's GSTIN on file is invalid.")
    if not org.get("Pincode"):
        problems.append("Your organization has no Pincode on file (Org Profile) — required "
                        "for the IRP's SellerDtls.Pin.")

    customer = co.get_customer(inv["customer_id"], fpath) or {}
    if not customer.get("Pincode"):
        problems.append(f"{inv['customer_id']} has no Pincode on file (Customer Onboarding) "
                        "— required for the IRP's BuyerDtls.Pin.")

    items = bl.get_invoice_items(invoice_id, fpath)
    if not items:
        problems.append(f"{invoice_id} has no line items — nothing to register.")
    for it in items:
        if not it["hsn_code"]:
            problems.append(f"Line {it['line_item']} ({it['mat_code']}) has no HSN code on file.")

    return (len(problems) == 0), problems


def _build_payload(invoice_id, data_file=None):
    fpath = data_file or DATA_FILE
    inv = bl.get_invoice(invoice_id, fpath)
    items = bl.get_invoice_items(invoice_id, fpath)
    org = op.get_org_profile(fpath)
    customer = co.get_customer(inv["customer_id"], fpath)

    _, _, seller_details = vo.validate_gstin(org["GSTIN"])
    _, _, buyer_details = vo.validate_gstin(customer["GSTIN"])
    seller_state = seller_details["state_code"]
    buyer_state = buyer_details["state_code"]

    item_list = []
    for it in items:
        taxable = float(it["taxable_value"] or 0)
        item_list.append({
            "SlNo": str(it["line_item"]),
            "PrdDesc": it["mat_desc"],
            "IsServc": "N",
            "HsnCd": it["hsn_code"],
            "Qty": float(it["qty"]),
            "Unit": it["uom"] or "NOS",
            "UnitPrice": float(it["unit_price"]),
            "TotAmt": taxable,
            "AssAmt": taxable,
            "GstRt": float(it["gst_rate"] or 0),
            "CgstAmt": float(it["cgst_amount"] or 0),
            "SgstAmt": float(it["sgst_amount"] or 0),
            "IgstAmt": float(it["igst_amount"] or 0),
            "TotItemVal": float(it["line_total"]),
        })

    payload = {
        "Version": "1.1",
        "TranDtls": {"TaxSch": "GST", "SupTyp": "B2B", "RegRev": "N", "IgstOnIntra": "N"},
        "DocDtls": {"Typ": "INV", "No": invoice_id, "Dt": _fmt_date_ddmmyyyy(inv["invoice_date"])},
        "SellerDtls": {
            "Gstin": org["GSTIN"], "LglNm": org["Legal_Name"], "Addr1": org["Address"],
            "Loc": org["City"], "Pin": int(org["Pincode"]), "Stcd": seller_state,
            "Ph": org.get("Contact_Phone") or "", "Em": org.get("Contact_Email") or "",
        },
        "BuyerDtls": {
            "Gstin": customer["GSTIN"], "LglNm": customer["Customer_Name"],
            "Pos": buyer_state, "Addr1": customer["Address"], "Loc": customer["City"],
            "Pin": int(customer["Pincode"]), "Stcd": buyer_state,
            "Ph": customer.get("Contact_Phone") or "", "Em": customer.get("Contact_Email") or "",
        },
        "ItemList": item_list,
        "ValDtls": {
            "AssVal": float(inv["subtotal"]), "CgstVal": float(inv["cgst_total"] or 0),
            "SgstVal": float(inv["sgst_total"] or 0), "IgstVal": float(inv["igst_total"] or 0),
            "TotInvVal": float(inv["grand_total"]),
        },
    }
    return payload


# ── Auth ──────────────────────────────────────────────────────────────────────
def _get_auth_token(data_file=None):
    """
    Cached in-process, real ~6-hour NIC token lifetime honored by
    re-authenticating once the cached expiry passes. Raises ValueError
    with the IRP's own error text on failure — never swallowed, matching
    accounting.py's own refuse-rather-than-guess discipline.
    """
    fpath = data_file or DATA_FILE
    now = datetime.now()
    if _token_cache["token"] and _token_cache["expires_at"] and now < _token_cache["expires_at"]:
        return _token_cache["token"]

    org = op.get_org_profile(fpath) or {}
    creds = _credentials()
    headers = {
        "Content-Type": "application/json",
        "client-id": creds["client_id"],
        "client-secret": creds["client_secret"],
        "gstin": org.get("GSTIN", ""),
    }
    body = {"UserName": creds["username"], "Password": creds["password"]}
    try:
        resp = requests.post(BASE_URL + AUTH_PATH, json=body, headers=headers, timeout=20)
    except requests.RequestException as e:
        raise ValueError(f"Could not reach the e-invoice sandbox ({BASE_URL}): {e}")

    data = _safe_json(resp)
    token = (data.get("AuthToken") or data.get("Data", {}).get("AuthToken")
            if isinstance(data, dict) else None)
    if resp.status_code != 200 or not token:
        raise ValueError(f"Authentication failed ({resp.status_code}): {_error_text(data)}")

    _token_cache["token"] = token
    _token_cache["expires_at"] = now + timedelta(hours=6)
    return token


def _safe_json(resp):
    try:
        return resp.json()
    except ValueError:
        return {"raw": resp.text}


def _error_text(data):
    if isinstance(data, dict):
        return data.get("ErrorDetails") or data.get("Message") or data.get("error") or str(data)
    return str(data)


# ── Generate / Cancel ─────────────────────────────────────────────────────────
def generate_irn(invoice_id, generated_by="", data_file=None):
    """
    The real, live call. Raises ValueError naming every unmet
    prerequisite (check_prerequisites()) before ever attempting the
    network call — a strict government schema rejecting a request for a
    reason this module could have caught locally isn't a useful failure
    mode. On a genuine IRP-side rejection, the raw error is stored on
    the e_invoices row (status='Error') and re-raised, not swallowed —
    the UI shows exactly what the IRP said, and the invoice stays
    retriable.
    """
    fpath = data_file or DATA_FILE
    ok, problems = check_prerequisites(invoice_id, fpath)
    if not ok:
        raise ValueError("Can't generate an IRN for " + invoice_id + ":\n- " + "\n- ".join(problems))

    token = _get_auth_token(fpath)
    payload = _build_payload(invoice_id, fpath)
    org = op.get_org_profile(fpath)
    creds = _credentials()
    headers = {
        "Content-Type": "application/json",
        "client-id": creds["client_id"],
        "client-secret": creds["client_secret"],
        "gstin": org.get("GSTIN", ""),
        "user_name": creds["username"],
        "authorization": token,
    }

    conn = db.get_connection()
    try:
        try:
            resp = requests.post(BASE_URL + GENERATE_PATH, json=payload, headers=headers, timeout=30)
        except requests.RequestException as e:
            msg = f"Could not reach the e-invoice sandbox ({BASE_URL}): {e}"
            _upsert_error(conn, invoice_id, msg)
            raise ValueError(msg)

        data = _safe_json(resp)
        result = data.get("Data", data) if isinstance(data, dict) else {}
        irn = result.get("Irn")
        if resp.status_code != 200 or not irn:
            msg = f"IRP rejected the invoice ({resp.status_code}): {_error_text(data)}"
            _upsert_error(conn, invoice_id, msg)
            raise ValueError(msg)

        ack_date = result.get("AckDt") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "INSERT INTO e_invoices (invoice_id, irn, ack_no, ack_date, signed_invoice, "
            "signed_qr_code, status, generated_date, generated_by, error_message) "
            "VALUES (?,?,?,?,?,?,?,?,?,?) "
            "ON CONFLICT(invoice_id) DO UPDATE SET irn=excluded.irn, ack_no=excluded.ack_no, "
            "ack_date=excluded.ack_date, signed_invoice=excluded.signed_invoice, "
            "signed_qr_code=excluded.signed_qr_code, status=excluded.status, "
            "generated_date=excluded.generated_date, generated_by=excluded.generated_by, "
            "error_message=NULL, cancel_reason=NULL, cancel_date=NULL",
            (invoice_id, irn, result.get("AckNo"), ack_date, result.get("SignedInvoice"),
             result.get("SignedQRCode"), "Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
             generated_by, None))
        conn.commit()
    finally:
        conn.close()

    return {"invoice_id": invoice_id, "irn": irn, "ack_no": result.get("AckNo"),
            "ack_date": ack_date, "signed_qr_code": result.get("SignedQRCode")}


def _upsert_error(conn, invoice_id, message):
    conn.execute(
        "INSERT INTO e_invoices (invoice_id, status, error_message) VALUES (?,?,?) "
        "ON CONFLICT(invoice_id) DO UPDATE SET status='Error', error_message=excluded.error_message",
        (invoice_id, "Error", message))
    conn.commit()


def cancel_irn(invoice_id, reason, cancelled_by="", data_file=None):
    """
    Real 24-hour-window enforcement (is_cancellable()) before ever
    calling the IRP — same up-front-gate discipline as generate_irn().
    Past the window, the real remedy is a credit memo (rma.py), not a
    cancellation this module would otherwise let through incorrectly.
    """
    fpath = data_file or DATA_FILE
    if not reason or not reason.strip():
        raise ValueError("A cancellation reason is required.")
    ok, msg = is_cancellable(invoice_id, fpath)
    if not ok:
        raise ValueError(msg)

    einv = get_einvoice(invoice_id, fpath)
    token = _get_auth_token(fpath)
    org = op.get_org_profile(fpath)
    creds = _credentials()
    headers = {
        "Content-Type": "application/json",
        "client-id": creds["client_id"],
        "client-secret": creds["client_secret"],
        "gstin": org.get("GSTIN", ""),
        "user_name": creds["username"],
        "authorization": token,
    }
    body = {"Irn": einv["irn"], "CnlRsn": "1", "CnlRem": reason.strip()[:100]}

    try:
        resp = requests.post(BASE_URL + CANCEL_PATH, json=body, headers=headers, timeout=30)
    except requests.RequestException as e:
        raise ValueError(f"Could not reach the e-invoice sandbox ({BASE_URL}): {e}")

    data = _safe_json(resp)
    result = data.get("Data", data) if isinstance(data, dict) else {}
    if resp.status_code != 200 or (isinstance(result, dict) and result.get("Status") not in (1, "1", "CNL", None)):
        raise ValueError(f"IRP rejected the cancellation ({resp.status_code}): {_error_text(data)}")

    conn = db.get_connection()
    try:
        conn.execute(
            "UPDATE e_invoices SET status='Cancelled', cancel_reason=?, cancel_date=? "
            "WHERE invoice_id=?",
            (reason.strip(), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), invoice_id))
        conn.commit()
    finally:
        conn.close()
    return {"invoice_id": invoice_id, "status": "Cancelled"}


def stats(data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT status, COUNT(*) c FROM e_invoices GROUP BY status").fetchall()
    finally:
        conn.close()
    return {r["status"]: r["c"] for r in rows}


if __name__ == "__main__":
    print("Sandbox configured:", is_configured())
    print("Stats:", stats())
