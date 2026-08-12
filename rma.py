"""
rma.py — O2C-E07: Returns & Credit Management (RMA-US-01 through 04).

Design scoped and recorded in CONTEXT_HANDOFF_v2.md ss7k before any code
here was written; read that section for the full reasoning. Two real
adaptations from the source document (O2C_User_Stories_V1_10.docx) to
this actual codebase, both found by reading real code, not assumed:

- Identity key is so_id + material_code, never the raw so_line_item
  number -- fulfillment_items, invoice_items, Reservations and
  Backorders all already discard line-item distinctness by the time
  anything downstream touches them, so RMA follows the same convention
  rather than inventing a second one nothing else honors.
- Pending Disposition needs no new "excluded transaction type".
  inventory.py has no type-filtering anywhere -- on-hand is a flat sum
  of signed transactions, and Quality Hold's own exclusion works by
  posting a real *negative* transaction at fail-time. So RMA-US-02's
  receipt posts NO inventory transaction at all; received qty and
  condition live only in this module's own `rmas` row. Stock only
  actually enters inventory.py's ledger at RMA-US-03's Sellable
  disposition -- which matches the document's own wording ("moves into
  sellable on-hand... at disposition") more literally than a
  receipt-time posting would have.

RMA-US-03's Return-to-Vendor outcome is deliberately NOT built here.
rtv.py's own vendor lookup is hard-wired to a Quality Hold's gr_id; an
RMA-received unit has no GR to trace a vendor from without real
lot/batch genealogy (Traceability, TRC-US-01 -- not built). Asked
directly, resolved: defer. Only Sellable and Scrap are real
dispositions in this build; DISPOSITION_ACCOUNTS below is the actual,
current, complete set -- not a partial list awaiting more entries.

Status on a `rmas` row represents physical/authorization state only:
Authorized -> Received/Partially Received (set once receipt happens),
or Cancelled (only possible pre-receipt). Disposition and credit-memo
issuance are each their own independently-queryable fact layered on
top (via the `disposition` column and the `credit_memos` table
respectively) -- never folded into `status` itself, since the source
document allows a credit memo before disposition (RMA-US-04's own
precondition is "Received", not "Disposed"), and overloading `status`
to also mean "credited" would make a real disposition silently
disappear from the Pending Disposition worklist the moment a credit
memo posted first.
"""

from datetime import date, timedelta

import db
import sales_order as so
import fulfillment as ful
import billing as bl
import inventory as inv
import accounting as acct
import po_export

RETURN_WINDOW_DAYS = 30

DISPOSITION_ACCOUNTS = {
    "Sellable": {"debit": "1200", "credit": "4100"},
    "Scrap": {"debit": "5200", "credit": "1200"},
}


def _next_rma_id(conn):
    rows = conn.execute("SELECT rma_id FROM rmas WHERE rma_id LIKE 'RMA-%'").fetchall()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["rma_id"].split("-")[1]))
        except Exception: pass
    return f"RMA-{mx+1:05d}"


def _next_credit_memo_id(conn):
    rows = conn.execute(
        "SELECT credit_memo_id FROM credit_memos WHERE credit_memo_id LIKE 'CM-%'"
    ).fetchall()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["credit_memo_id"].split("-")[1]))
        except Exception: pass
    return f"CM-{mx+1:05d}"


def _row_to_rma(r):
    d = dict(r)
    d["mat_code"] = d.pop("material_code")
    d["mat_desc"] = d.pop("material_desc")
    return d


def _row_to_credit_memo(r):
    d = dict(r)
    d["mat_code"] = d.pop("material_code")
    d["mat_desc"] = d.pop("material_desc")
    return d


# ── Read ──────────────────────────────────────────────────────────────────────
def get_rmas(status=None, so_id=None, data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT * FROM rmas ORDER BY rma_id").fetchall()
    finally:
        conn.close()
    out = [_row_to_rma(r) for r in rows]
    if status:
        out = [r for r in out if r["status"] == status]
    if so_id:
        out = [r for r in out if r["so_id"] == so_id]
    return out


def get_rma(rma_id, data_file=None):
    for r in get_rmas(data_file=data_file):
        if r["rma_id"] == rma_id:
            return r
    return None


def get_credit_memos(rma_id=None, invoice_id=None, data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT * FROM credit_memos ORDER BY credit_memo_id").fetchall()
    finally:
        conn.close()
    out = [_row_to_credit_memo(r) for r in rows]
    if rma_id:
        out = [r for r in out if r["rma_id"] == rma_id]
    if invoice_id:
        out = [r for r in out if r["invoice_id"] == invoice_id]
    return out


def get_pending_receipt(data_file=None):
    """Authorized RMAs with nothing physically received yet — RMA-US-02's worklist."""
    return get_rmas(status="Authorized", data_file=data_file)


def get_pending_disposition(data_file=None):
    """Received (or Partially Received) stock with no disposition recorded yet — RMA-US-03's worklist."""
    return [r for r in get_rmas(data_file=data_file)
            if r["status"] in ("Received", "Partially Received") and not r["disposition"]]


def get_ready_for_credit(data_file=None):
    """Received (or Partially Received) RMAs with no credit memo issued yet — RMA-US-04's worklist.
    Deliberately independent of disposition state — the source document's own precondition for a
    credit memo is Received, not Disposed."""
    credited_ids = {cm["rma_id"] for cm in get_credit_memos(data_file=data_file)}
    return [r for r in get_rmas(data_file=data_file)
            if r["status"] in ("Received", "Partially Received") and r["rma_id"] not in credited_ids]


# ── RMA-US-01: Create and authorize ─────────────────────────────────────────────
def _goods_issued_fulfillment(so_id, data_file=None):
    """The one active (non-Cancelled) fulfillment for this SO, if any —
    fulfillment.so_already_fulfilled() already enforces at most one exists."""
    fs = [f for f in ful.get_fulfillments(so_id=so_id, data_file=data_file) if f["status"] != "Cancelled"]
    return fs[0] if fs else None


def get_goods_issued_qty(so_id, mat_code, data_file=None):
    f = _goods_issued_fulfillment(so_id, data_file)
    if f is None:
        return 0.0
    items = {i["mat_code"]: i for i in ful.get_fulfillment_items(f["fulfillment_id"], data_file)}
    item = items.get(mat_code)
    return float(item["qty_shipped"] or 0) if item else 0.0


def get_remaining_returnable_qty(so_id, mat_code, data_file=None):
    """Goods-issued quantity net of every non-Cancelled RMA already authorized against
    this line — computed live, never stored, so it can never drift from the real ledger."""
    issued = get_goods_issued_qty(so_id, mat_code, data_file)
    committed = sum(r["requested_qty"] for r in get_rmas(so_id=so_id, data_file=data_file)
                     if r["mat_code"] == mat_code and r["status"] != "Cancelled")
    return round(issued - committed, 3)


def create_rma(so_id, mat_code, requested_qty, reason, authorized_by="",
               window_override_reason="", data_file=None):
    if requested_qty <= 0:
        raise ValueError("Requested return quantity must be positive.")
    if not (reason or "").strip():
        raise ValueError("A return reason is required.")

    f = _goods_issued_fulfillment(so_id, data_file)
    if f is None:
        raise ValueError(f"{so_id} has no Goods Issue on record — nothing to return.")
    items = {i["mat_code"]: i for i in ful.get_fulfillment_items(f["fulfillment_id"], data_file)}
    item = items.get(mat_code)
    if item is None or not (item["qty_shipped"] or 0):
        raise ValueError(f"{mat_code} was never goods-issued on {so_id} — nothing to return.")

    remaining = get_remaining_returnable_qty(so_id, mat_code, data_file)
    if requested_qty > remaining + 0.005:
        raise ValueError(f"Requested quantity {requested_qty:g} exceeds the remaining returnable "
                          f"quantity of {remaining:g} for {mat_code} on {so_id}.")

    outside_window = False
    if f["shipped_date"]:
        window_end = date.fromisoformat(f["shipped_date"][:10]) + timedelta(days=RETURN_WINDOW_DAYS)
        outside_window = date.today() > window_end
    if outside_window and not (window_override_reason or "").strip():
        raise ValueError(f"This return is outside the {RETURN_WINDOW_DAYS}-day window (shipped "
                          f"{f['shipped_date']}) — an override reason is required to authorize it anyway.")

    db.init_schema()
    conn = db.get_connection()
    try:
        rma_id = _next_rma_id(conn)
        conn.execute(
            "INSERT INTO rmas (rma_id, so_id, fulfillment_id, material_code, material_desc, "
            "requested_qty, reason, status, authorized_date, authorized_by, window_override_reason) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (rma_id, so_id, f["fulfillment_id"], mat_code, item["mat_desc"], requested_qty,
             reason, "Authorized", date.today().strftime("%Y-%m-%d"), authorized_by,
             window_override_reason if outside_window else ""),
        )
        conn.commit()
    finally:
        conn.close()
    return {"rma_id": rma_id, "status": "Authorized", "outside_window": outside_window}


def cancel_rma(rma_id, data_file=None):
    r = get_rma(rma_id, data_file)
    if r is None:
        raise ValueError(f"{rma_id} not found.")
    if r["status"] != "Authorized":
        raise ValueError(f"{rma_id} is '{r['status']}' — only an Authorized RMA with nothing "
                          "physically received yet can be cancelled.")
    conn = db.get_connection()
    try:
        conn.execute("UPDATE rmas SET status='Cancelled' WHERE rma_id=?", (rma_id,))
        conn.commit()
    finally:
        conn.close()


# ── RMA-US-02: Receive ──────────────────────────────────────────────────────────
def receive_rma(rma_id, received_qty, condition_note, data_file=None):
    r = get_rma(rma_id, data_file)
    if r is None:
        raise ValueError(f"{rma_id} not found.")
    if r["status"] != "Authorized":
        raise ValueError(f"{rma_id} is '{r['status']}' — only an Authorized RMA can be received.")
    if received_qty <= 0:
        raise ValueError("Received quantity must be positive.")
    if received_qty > r["requested_qty"] + 0.005:
        raise ValueError(f"Received quantity {received_qty:g} exceeds the authorized quantity "
                          f"of {r['requested_qty']:g}.")
    if not (condition_note or "").strip():
        raise ValueError("A condition note is required to confirm this receipt.")

    f = ful.get_fulfillment(r["fulfillment_id"], data_file)
    receiving_location = f["delivery_location"] if f else ""
    new_status = "Received" if received_qty >= r["requested_qty"] - 0.005 else "Partially Received"

    conn = db.get_connection()
    try:
        conn.execute(
            "UPDATE rmas SET status=?, received_qty=?, received_date=?, condition_note=?, "
            "receiving_location=? WHERE rma_id=?",
            (new_status, received_qty, date.today().strftime("%Y-%m-%d"), condition_note,
             receiving_location, rma_id),
        )
        conn.commit()
    finally:
        conn.close()
    return {"rma_id": rma_id, "status": new_status}


# ── RMA-US-03: Disposition ───────────────────────────────────────────────────────
def _item_cost(mat_code, data_file=None):
    cost_by_code = {i["code"]: i["price"] for i in po_export.load_item_master(data_file)}
    cost = cost_by_code.get(mat_code)
    if cost is None:
        raise ValueError(f"No cost found for {mat_code} in Item Master — can't post a "
                          "disposition GL entry.")
    return float(cost)


def dispose_rma(rma_id, disposition, disposed_by="", notes="", data_file=None):
    if disposition not in DISPOSITION_ACCOUNTS:
        raise ValueError(f"'{disposition}' is not a supported disposition in this build — only "
                          f"{', '.join(DISPOSITION_ACCOUNTS)} (Return-to-Vendor is deferred until "
                          "real lot/batch traceability exists to identify a vendor of origin).")
    r = get_rma(rma_id, data_file)
    if r is None:
        raise ValueError(f"{rma_id} not found.")
    if r["status"] not in ("Received", "Partially Received"):
        raise ValueError(f"{rma_id} is '{r['status']}' — only received stock can be dispositioned.")
    if r["disposition"]:
        raise ValueError(f"{rma_id} was already dispositioned '{r['disposition']}' — not silently "
                          "re-dispositionable; a correction is a separate, explicit reversal.")

    qty = r["received_qty"]
    unit_cost = _item_cost(r["mat_code"], data_file)
    value = round(unit_cost * qty, 2)
    accts = DISPOSITION_ACCOUNTS[disposition]

    je_id = None
    if value > 0:
        je_id = acct.post_journal_entry(
            "RMA", rma_id,
            f"RMA {disposition} disposition — {r['mat_desc']} ({qty:g} units)",
            [{"account_code": accts["debit"], "debit": value, "credit": 0,
              "description": f"{disposition} disposition — {rma_id}"},
             {"account_code": accts["credit"], "debit": 0, "credit": value,
              "description": f"{disposition} disposition — {rma_id}"}],
            data_file=data_file)

    if disposition == "Sellable":
        inv.record_transaction(r["mat_code"], r["mat_desc"], r["receiving_location"], qty,
                               "RMA Sellable Return", reference_type="RMA", reference_id=rma_id,
                               notes=f"Dispositioned Sellable from {rma_id}", data_file=data_file)

    conn = db.get_connection()
    try:
        conn.execute(
            "UPDATE rmas SET disposition=?, disposed_qty=?, disposed_date=?, disposed_by=?, "
            "disposition_notes=?, disposition_je_id=? WHERE rma_id=?",
            (disposition, qty, date.today().strftime("%Y-%m-%d"), disposed_by, notes, je_id, rma_id),
        )
        conn.commit()
    finally:
        conn.close()
    return {"rma_id": rma_id, "disposition": disposition, "je_id": je_id}


# ── RMA-US-04: Credit memo ───────────────────────────────────────────────────────
def _find_original_invoice_line(so_id, mat_code, data_file=None):
    """The most recent invoice covering this so_id + mat_code, and its own line —
    credited price and GST always mirror this line exactly, never re-priced."""
    candidates = []
    for invoice in bl.get_invoices(data_file=data_file):
        if invoice["so_id"] != so_id:
            continue
        for item in bl.get_invoice_items(invoice["invoice_id"], data_file):
            if item["mat_code"] == mat_code:
                candidates.append((invoice, item))
    if not candidates:
        return None, None
    candidates.sort(key=lambda pair: pair[0]["invoice_date"] or "", reverse=True)
    return candidates[0]


def issue_credit_memo(rma_id, issued_by="", data_file=None):
    r = get_rma(rma_id, data_file)
    if r is None:
        raise ValueError(f"{rma_id} not found.")
    if r["status"] not in ("Received", "Partially Received"):
        raise ValueError(f"{rma_id} is '{r['status']}' — a credit memo can only be issued "
                          "against a received return.")
    if get_credit_memos(rma_id=rma_id, data_file=data_file):
        raise ValueError(f"{rma_id} already has a credit memo issued — not re-issuable; a "
                          "correction is a new, separate adjustment.")

    invoice, line = _find_original_invoice_line(r["so_id"], r["mat_code"], data_file)
    if invoice is None:
        raise ValueError(f"No invoice found for {r['mat_code']} on {r['so_id']} — can't issue "
                          "a credit memo without a resolvable original invoice.")

    qty = r["received_qty"]
    line_qty = line["qty"] or 1
    ratio = qty / line_qty
    taxable_value = round(line["taxable_value"] * ratio, 2)
    cgst = round((line["cgst_amount"] or 0) * ratio, 2)
    sgst = round((line["sgst_amount"] or 0) * ratio, 2)
    igst = round((line["igst_amount"] or 0) * ratio, 2)
    credit_total = round(taxable_value + cgst + sgst + igst, 2)

    gl_lines = [{"account_code": "4000", "debit": taxable_value, "credit": 0,
                 "description": f"Sales Revenue reversal — {rma_id}"}]
    if cgst:
        gl_lines.append({"account_code": "2200", "debit": cgst, "credit": 0,
                          "description": f"CGST reversal — {rma_id}"})
    if sgst:
        gl_lines.append({"account_code": "2210", "debit": sgst, "credit": 0,
                          "description": f"SGST reversal — {rma_id}"})
    if igst:
        gl_lines.append({"account_code": "2220", "debit": igst, "credit": 0,
                          "description": f"IGST reversal — {rma_id}"})
    gl_lines.append({"account_code": "1100", "debit": 0, "credit": credit_total,
                      "description": f"AR reduced — {rma_id}"})

    je_id = acct.post_journal_entry("CreditMemo", rma_id,
                                    f"Credit memo for {rma_id} against {invoice['invoice_id']}",
                                    gl_lines, data_file=data_file)

    db.init_schema()
    conn = db.get_connection()
    try:
        cm_id = _next_credit_memo_id(conn)
        conn.execute(
            "INSERT INTO credit_memos (credit_memo_id, rma_id, invoice_id, material_code, "
            "material_desc, qty, unit_price, taxable_value, cgst_amount, sgst_amount, "
            "igst_amount, credit_total, issued_date, issued_by, je_id) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (cm_id, rma_id, invoice["invoice_id"], r["mat_code"], r["mat_desc"], qty,
             line["unit_price"], taxable_value, cgst, sgst, igst, credit_total,
             date.today().strftime("%Y-%m-%d"), issued_by, je_id),
        )
        conn.commit()
    finally:
        conn.close()
    return {"credit_memo_id": cm_id, "invoice_id": invoice["invoice_id"],
            "credit_total": credit_total, "je_id": je_id}


def stats(data_file=None):
    return {"pending_receipt": len(get_pending_receipt(data_file)),
            "pending_disposition": len(get_pending_disposition(data_file)),
            "ready_for_credit": len(get_ready_for_credit(data_file))}


if __name__ == "__main__":
    print("RMA stats:", stats())
