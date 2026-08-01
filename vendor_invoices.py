"""
vendor_invoices.py — Vendor Invoices, stage 2 of the proper 3-way match
(Goods Receipt / Invoice Receipt / Payment).

goods_receipt.create_gr() -> accounting.post_gr_entry() posts stage 1:
Dr Inventory / Cr GR/IR Clearing, ex-GST — a temporary clearing
liability, not the final vendor liability yet. This module is stage 2:
Dr GR/IR Clearing (clearing what stage 1 posted) + Dr GST Input (GST is
determined HERE, not at GR time — see compute_invoice_amounts()) /
Cr Accounts Payable, the real, final liability, credited for the first
time. Stage 3 is record_payment() below: Dr Accounts Payable / Cr Cash.

This mirrors real practice (SAP's GR/IR process, for instance) more
than an earlier version of this module did — that version had GR post
straight to Accounts Payable with GST already claimed, which collapsed
stages 1 and 2 into one, left this module with nothing to post at
invoice time, and meant GR/IR Clearing's balance (now a genuine "goods
received, not yet invoiced" signal) didn't exist as a concept at all.

This module is the mirror image of billing.py + cash_application.py,
built for the purchase side instead of the sales side. Terminology
deliberately follows the same word on both sides — this module says
"invoice," not "bill," for the same document customer-side calls an
invoice.

  create_invoice()  -> Dr GR/IR Clearing + Dr GST Input CGST/SGST/IGST
                       (if determinable) / Cr Accounts Payable. No
                       amount override — always posts exactly what
                       compute_invoice_amounts() computes; see that
                       function's docstring for why a caller-supplied
                       override isn't offered.
  record_payment()   -> Dr Accounts Payable / Cr Cash and Bank.
                       Unchanged by this rework — this was already the
                       correct stage-3 posting. Supports partial
                       payment, same as cash_application.py's AR-side
                       equivalent, for the same reason: a short payment
                       is just as real on the AP side.

Deliberately simpler than cash_application.py's Payments/
Payment_Applications shape: an invoice here belongs to exactly one GR
(no many-invoices-one-payment matching flexibility), since that's the
actual shape of what a GR-driven AP flow needs for this PoC. A payment
can still be partial and an invoice can still take more than one
payment over time — just always against the one invoice it was
created for.

New tables:
  Vendor_Invoices          Invoice_ID | GR_ID | PO_Number | Vendor_ID |
                           Vendor_Name | Invoice_Number | Invoice_Date |
                           Amount | Paid_Amount | Status | Notes
  Vendor_Invoice_Payments  Payment_ID | Invoice_ID | Vendor_ID |
                           Vendor_Name | Payment_Date | Amount |
                           Payment_Method | Reference_No | Notes
"""

import os
from datetime import date

import db
import goods_receipt as gr
import accounting as acct
import vendor_onboarding as vo
import item_tax as it
import org_profile as op
import billing as bl  # reuse _gst_split — same math as the sales side, roles swapped

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")


def ensure_sheets(wb=None):
    """Kept for signature compatibility with the rest of this app's
    modules — Vendor_Invoices/Vendor_Invoice_Payments live only in
    SQLite, there was never an Excel version. `wb` is accepted and
    ignored."""
    db.init_schema()


# ── Read ──────────────────────────────────────────────────────────────────────
def get_invoices(gr_id=None, status=None, data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT * FROM vendor_invoices ORDER BY invoice_id").fetchall()
    finally:
        conn.close()
    out = []
    for r in rows:
        row = {"invoice_id": r["invoice_id"], "gr_id": r["gr_id"], "po_number": r["po_number"],
               "vendor_id": r["vendor_id"], "vendor_name": r["vendor_name"],
               "invoice_number": r["invoice_number"], "invoice_date": r["invoice_date"],
               "amount": r["amount"], "paid_amount": r["paid_amount"] or 0,
               "status": r["status"], "notes": r["notes"]}
        if gr_id and row["gr_id"] != gr_id:
            continue
        if status and row["status"] != status:
            continue
        out.append(row)
    return out


def get_invoice(invoice_id, data_file=None):
    for inv in get_invoices(data_file=data_file):
        if inv["invoice_id"] == invoice_id:
            return inv
    return None


def invoice_for_gr(gr_id, data_file=None):
    """The existing invoice against this GR, if any — create_invoice()
    refuses a second one, so this is always 0 or 1 result."""
    invoices = get_invoices(gr_id=gr_id, data_file=data_file)
    return invoices[0] if invoices else None


def get_invoice_payment_info(invoice_id, data_file=None):
    inv = get_invoice(invoice_id, data_file)
    if inv is None:
        raise ValueError(f"{invoice_id} not found.")
    balance_due = round(inv["amount"] - inv["paid_amount"], 2)
    return {"amount": inv["amount"], "paid_amount": inv["paid_amount"],
            "balance_due": balance_due, "status": inv["status"]}


def get_invoice_payments(invoice_id, data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute(
            "SELECT * FROM vendor_invoice_payments WHERE invoice_id = ? ORDER BY payment_id",
            (invoice_id,),
        ).fetchall()
    finally:
        conn.close()
    return [{"payment_id": r["payment_id"], "invoice_id": r["invoice_id"], "vendor_id": r["vendor_id"],
             "vendor_name": r["vendor_name"], "payment_date": r["payment_date"],
             "amount": r["amount"], "payment_method": r["payment_method"],
             "reference_no": r["reference_no"], "notes": r["notes"]} for r in rows]


def gr_clearing_amount(gr_id, data_file=None):
    """The GR's own posted GR/IR Clearing amount (ex-GST) — read straight
    off its journal entry (the account 2100 credit line), which is the
    authoritative record of what GR actually posted. Falls back to a
    plain qty*price recompute only if the GR was never posted to
    accounting — a best-effort default rather than a hard block, so
    create_invoice() still works even for a GR whose accounting posting
    failed or was skipped."""
    for je in acct.get_journal_entries(source_type="GR", data_file=data_file):
        if je["source_id"] == gr_id:
            for line in acct.get_journal_entry_lines(je["je_id"], data_file):
                if line["account_code"] == "2100":
                    return round(line["credit"], 2)
    items = gr.get_gr_items(gr_id, data_file)
    return round(sum((i["qty_received"] or 0) * (i["unit_price"] or 0) for i in items), 2)


def compute_invoice_amounts(gr_id, data_file=None):
    """
    Everything create_invoice() needs to post, computed once here and
    shared with the UI — so a preview shown before confirming is
    guaranteed to match exactly what gets posted, not a second,
    separately-maintained calculation that could quietly drift from it.

    GST determination moved here from accounting.post_gr_entry() — see
    that function's docstring for why (real practice determines tax at
    invoice verification, not goods receipt). Same _gst_split() math as
    the sales side, just with roles swapped: our org is the "buyer"
    whose state gets compared against the vendor's, not the customer's.

    Same graceful-degradation reasoning as before too: a line with no
    GST rate on file, or a vendor with no GSTIN on file (unregistered/
    composition-scheme vendors are a legitimate real case), simply
    doesn't get a GST Input claim for that portion — it doesn't block
    the invoice from posting at all, it's not an error condition.

    Returns {clearing_amount, cgst, sgst, igst, total}.
    """
    fpath = data_file or DATA_FILE
    g = gr.get_gr(gr_id, fpath)
    if g is None:
        raise ValueError(f"{gr_id} not found.")

    clearing_amount = gr_clearing_amount(gr_id, fpath)

    org_state = None
    if op.is_configured(fpath):
        org = op.get_org_profile(fpath)
        ok, _, details = vo.validate_gstin(org["GSTIN"])
        if ok:
            org_state = details["state_code"]
    vendor = vo.get_vendor(g["vendor_id"], fpath) if g.get("vendor_id") else None
    vendor_state = None
    if vendor and vendor.get("GSTIN"):
        ok, _, details = vo.validate_gstin(vendor["GSTIN"])
        if ok:
            vendor_state = details["state_code"]

    cgst_t = sgst_t = igst_t = 0.0
    items = [i for i in gr.get_gr_items(gr_id, fpath) if (i["qty_received"] or 0) > 0]
    for item in items:
        taxable = round(float(item["qty_received"]) * float(item["unit_price"] or 0), 2)
        tax = it.get_item_tax_info(item["mat_code"], fpath)
        gst_rate = tax["gst_rate"] if tax else None
        if gst_rate is not None and org_state and vendor_state:
            cgst, sgst, igst = bl._gst_split(taxable, float(gst_rate), org_state, vendor_state)
            cgst_t += cgst; sgst_t += sgst; igst_t += igst

    cgst_t, sgst_t, igst_t = round(cgst_t, 2), round(sgst_t, 2), round(igst_t, 2)
    total = round(clearing_amount + cgst_t + sgst_t + igst_t, 2)
    return {"clearing_amount": clearing_amount, "cgst": cgst_t, "sgst": sgst_t,
            "igst": igst_t, "total": total}


# ── Create an invoice (stage 2 of the 3-way match — the real posting) ───────────
def _next_invoice_id(conn):
    rows = conn.execute("SELECT invoice_id FROM vendor_invoices WHERE invoice_id LIKE 'VI-%'").fetchall()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["invoice_id"].split("-")[1]))
        except Exception: pass
    return f"VI-{mx+1:05d}"


def create_invoice(gr_id, invoice_number="", invoice_date=None, notes="", data_file=None):
    """
    Dr GR/IR Clearing (clearing what GR posted) + Dr GST Input CGST/SGST/
    IGST (determined now, not at GR time — see compute_invoice_amounts())
    / Cr Accounts Payable (the total, first time AP is credited for this
    GR — the real, final vendor liability).

    No amount override parameter anymore — this always posts exactly
    what compute_invoice_amounts() computes, and that's also what's
    stored as the invoice's tracked amount. Deliberately: letting a
    caller substitute an arbitrary number here would either produce an
    unbalanced entry or need a price-variance account to absorb the
    difference, and this PoC doesn't model price variance — better to
    not pretend to than to plug the gap with something that looks
    like it balances but doesn't mean anything.
    """
    fpath = data_file or DATA_FILE
    g = gr.get_gr(gr_id, fpath)
    if g is None:
        raise ValueError(f"{gr_id} not found.")
    if invoice_for_gr(gr_id, fpath):
        raise ValueError(f"{gr_id} already has an invoice recorded against it.")

    amounts = compute_invoice_amounts(gr_id, fpath)
    total = amounts["total"]
    if not total or total <= 0:
        raise ValueError(f"Couldn't determine a positive amount for {gr_id}.")

    db.init_schema()
    conn = db.get_connection()
    try:
        invoice_id = _next_invoice_id(conn)
        conn.execute(
            "INSERT INTO vendor_invoices (invoice_id, gr_id, po_number, vendor_id, vendor_name, "
            "invoice_number, invoice_date, amount, paid_amount, status, notes) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (invoice_id, gr_id, g["po_number"], g["vendor_id"], g["vendor_name"],
             invoice_number, (invoice_date or date.today().strftime("%Y-%m-%d")),
             total, 0, "Open", notes),
        )
        conn.commit()
    finally:
        conn.close()

    lines = [{"account_code": "2100", "debit": amounts["clearing_amount"], "credit": 0,
             "description": f"GR/IR cleared for {gr_id}"}]
    if amounts["cgst"]:
        lines.append({"account_code": "1150", "debit": amounts["cgst"], "credit": 0,
                      "description": f"GST Input CGST on {invoice_id}"})
    if amounts["sgst"]:
        lines.append({"account_code": "1160", "debit": amounts["sgst"], "credit": 0,
                      "description": f"GST Input SGST on {invoice_id}"})
    if amounts["igst"]:
        lines.append({"account_code": "1170", "debit": amounts["igst"], "credit": 0,
                      "description": f"GST Input IGST on {invoice_id}"})
    lines.append({"account_code": "2000", "debit": 0, "credit": total,
                  "description": f"AP for {invoice_id} — {g['vendor_name']}"})
    acct.post_journal_entry("VendorInvoice", invoice_id,
        f"Vendor Invoice {invoice_id} — {g['vendor_name']}", lines, fpath)

    return invoice_id


# ── Payment (the one new real posting) ───────────────────────────────────────
def _next_payment_id(conn):
    rows = conn.execute(
        "SELECT payment_id FROM vendor_invoice_payments WHERE payment_id LIKE 'VIP-%'"
    ).fetchall()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["payment_id"].split("-")[1]))
        except Exception: pass
    return f"VIP-{mx+1:05d}"


def record_payment(invoice_id, amount, payment_date=None, payment_method="Bank Transfer",
                   reference_no="", notes="", data_file=None):
    fpath = data_file or DATA_FILE
    inv = get_invoice(invoice_id, fpath)
    if inv is None:
        raise ValueError(f"{invoice_id} not found.")
    if amount <= 0:
        raise ValueError("Payment amount must be positive.")
    info = get_invoice_payment_info(invoice_id, fpath)
    if amount > info["balance_due"] + 0.005:
        raise ValueError(f"{invoice_id} only has \u20b9{info['balance_due']:,.2f} outstanding — "
                          "can't pay more than that.")

    db.init_schema()
    conn = db.get_connection()
    try:
        payment_id = _next_payment_id(conn)
        conn.execute(
            "INSERT INTO vendor_invoice_payments (payment_id, invoice_id, vendor_id, vendor_name, "
            "payment_date, amount, payment_method, reference_no, notes) VALUES (?,?,?,?,?,?,?,?,?)",
            (payment_id, invoice_id, inv["vendor_id"], inv["vendor_name"],
             (payment_date or date.today().strftime("%Y-%m-%d")),
             round(amount, 2), payment_method, reference_no, notes),
        )
        new_paid = round(inv["paid_amount"] + amount, 2)
        new_status = "Paid" if new_paid >= inv["amount"] - 0.005 else "Partially Paid"
        conn.execute(
            "UPDATE vendor_invoices SET paid_amount = ?, status = ? WHERE invoice_id = ?",
            (new_paid, new_status, invoice_id),
        )
        conn.commit()
    finally:
        conn.close()

    je_id = acct.post_journal_entry(
        "VendorPayment", payment_id,
        f"Payment {payment_id} — {invoice_id} ({inv['vendor_name']})",
        [{"account_code": "2000", "debit": amount, "credit": 0,
          "description": f"AP settled, {invoice_id}"},
         {"account_code": "1000", "debit": 0, "credit": amount,
          "description": f"Cash paid, {payment_id}"}],
        fpath)
    return {"payment_id": payment_id, "je_id": je_id}


def stats(data_file=None):
    invoices = get_invoices(data_file=data_file)
    return {"total_invoices": len(invoices),
            "open": sum(1 for i in invoices if i["status"] == "Open"),
            "partially_paid": sum(1 for i in invoices if i["status"] == "Partially Paid"),
            "paid": sum(1 for i in invoices if i["status"] == "Paid"),
            "total_invoiced": round(sum(i["amount"] for i in invoices), 2),
            "total_paid": round(sum(i["paid_amount"] for i in invoices), 2)}


if __name__ == "__main__":
    print("Vendor Invoices stats:", stats())
