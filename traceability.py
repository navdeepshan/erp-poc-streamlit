"""
traceability.py — S2S-E06: TRC-US-01, batch/lot/serial capture with
expiry, FEFO suggestion, forward/backward trace, Near-Expiry.

Design scoped and recorded in CONTEXT_HANDOFF_v2.md ss7l before any
code here was written; read that section for the full reasoning. The
one thing worth repeating here: this module owns no separate ledger.
inventory.py's own on-hand is a flat SUM(quantity) over
inventory_transactions with no type-filtering anywhere -- so a lot's
own remaining quantity is never stored on the `lots` row, it's always
computed live as that same sum, additionally filtered by lot_id. A
lot's aggregate balance and the sum of its own lots can never drift
apart, because they're literally the same rows, just grouped
differently -- not two numbers kept in sync by convention.

Serial is a batch of size 1, not a second table or a second identity
column: a Serial-tracked receipt creates N `lots` rows (qty=1 each,
lot_number = the serial number itself) instead of one row with qty=N.
Tracking Type is fixed per material, so a `lots` row is never
ambiguous about which meaning its own lot_number carries.

Scope, deliberately: GR-US-01 (this module's own create_lots_for_
receipt(), called from goods_receipt.create_gr()), INV-US-01 (the
drill-down/trace/near-expiry reads below), and FEFO-aware ad hoc
transfers (inventory.ship_transfer()'s own optional lot_id param). O2C's
FUL-US-02/03 pick/pack lot capture and sto.py's own FEFO integration
are both explicitly out of scope for this build -- the first because
the source document itself calls it a follow-on decision for those
stories, the second by direct instruction (asked, resolved: ad hoc
transfer only for v1).
"""

from datetime import date, timedelta

import db

NEAR_EXPIRY_WINDOW_DAYS = 90


def get_tracking_info(material_code, data_file=None):
    """{'tracking_type': 'None'/'Batch'/'Serial', 'shelf_life_tracked': bool}
    for a material — defaults to untracked for a material not found at
    all, the same honest-default spirit as every other lookup here."""
    conn = db.get_connection()
    try:
        row = conn.execute(
            "SELECT tracking_type, shelf_life_tracked FROM item_master WHERE item_code=?",
            (material_code,)).fetchone()
    finally:
        conn.close()
    if row is None:
        return {"tracking_type": "None", "shelf_life_tracked": False}
    return {"tracking_type": row["tracking_type"] or "None",
            "shelf_life_tracked": (row["shelf_life_tracked"] or "No") == "Yes"}


def validate_lot_capture(material_code, qty, lot_number=None, serials=None,
                          expiry_date=None, data_file=None):
    """
    Raises with the document's own specific error language if a
    Batch/Serial-tracked material's receipt is missing what its own
    Tracking Type requires — a hard block, never a warning, since an
    untracked receipt of a tracked material would break every
    downstream trace query. A no-op (returns info, raises nothing) for
    a None-tracked material regardless of what's passed in.
    """
    info = get_tracking_info(material_code, data_file)
    tracking_type = info["tracking_type"]
    if tracking_type == "None":
        return info
    if tracking_type == "Batch":
        if not (lot_number or "").strip():
            raise ValueError(f"A lot/batch number is required for {material_code}.")
        if info["shelf_life_tracked"] and not expiry_date:
            raise ValueError(f"An expiry date is required for {material_code}.")
    elif tracking_type == "Serial":
        if not serials or len(serials) != qty:
            raise ValueError(f"Enter exactly {qty:g} serial numbers to match the received "
                              f"quantity for {material_code}.")
        if len(set(serials)) != len(serials):
            raise ValueError(f"Serial numbers for {material_code} must be unique within "
                              "this receipt.")
    return info


def create_lots_for_receipt(material_code, material_desc, qty, location_id, po_number,
                             vendor_id, vendor_name, gr_id, lot_number=None, serials=None,
                             expiry_date=None, data_file=None):
    """
    Called by goods_receipt.create_gr() for a Batch- or Serial-tracked
    line. Returns the list of newly-created lot rows (each a dict) —
    the caller threads each one's own "id" back into its own
    inv.record_transaction() call as lot_id, one call per lot/serial,
    so the physical identity survives on the ledger itself. Returns []
    for a None-tracked material — nothing to create.
    """
    info = validate_lot_capture(material_code, qty, lot_number, serials, expiry_date, data_file)
    tracking_type = info["tracking_type"]
    if tracking_type == "None":
        return []

    db.init_schema()
    conn = db.get_connection()
    try:
        if tracking_type == "Batch":
            rows_to_insert = [(lot_number, qty)]
        else:  # Serial — each serial is its own lots row, qty always 1
            for serial in serials:
                if conn.execute("SELECT 1 FROM lots WHERE material_code=? AND lot_number=?",
                                 (material_code, serial)).fetchone():
                    raise ValueError(f"Serial number '{serial}' already exists for "
                                      f"{material_code} — serials must be unique tenant-wide.")
            rows_to_insert = [(serial, 1) for serial in serials]

        new_lots = []
        for lot_number_val, lot_qty in rows_to_insert:
            cur = conn.execute(
                "INSERT INTO lots (lot_number, material_code, material_desc, qty_received, "
                "location_id, expiry_date, po_number, vendor_id, vendor_name, gr_id, "
                "received_date) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (lot_number_val, material_code, material_desc, lot_qty, location_id,
                 expiry_date, po_number, vendor_id, vendor_name, gr_id,
                 date.today().strftime("%Y-%m-%d")))
            new_lots.append({"id": cur.lastrowid, "lot_number": lot_number_val,
                             "material_code": material_code, "material_desc": material_desc,
                             "qty_received": lot_qty, "location_id": location_id,
                             "expiry_date": expiry_date, "po_number": po_number,
                             "vendor_id": vendor_id, "vendor_name": vendor_name, "gr_id": gr_id})
        conn.commit()
    finally:
        conn.close()
    return new_lots


def get_lot(lot_id, data_file=None):
    conn = db.get_connection()
    try:
        row = conn.execute("SELECT * FROM lots WHERE id=?", (lot_id,)).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def get_lot_remaining_qty(lot_id, location_id=None, data_file=None):
    """Live sum over inventory_transactions filtered by lot_id — the
    exact same computed-not-stored discipline inventory.get_balance()
    already uses at the aggregate level, now filtered one level deeper.

    location_id is optional and filters to just that location's own
    share — a real, found bug fixed here: a lot's stock can genuinely
    split across more than one location once any of it transfers
    (ship_transfer()/receive_transfer() both carry lot_id through), so
    a global, location-agnostic sum is the wrong question to ask when
    the caller actually means "how much of this lot is sitting HERE."
    Every existing caller that only ever cared about the lot's total
    everywhere keeps working unchanged (location_id defaults to None)."""
    conn = db.get_connection()
    try:
        query = "SELECT SUM(quantity) AS total FROM inventory_transactions WHERE lot_id=?"
        params = [lot_id]
        if location_id:
            query += " AND location_id=?"
            params.append(location_id)
        row = conn.execute(query, params).fetchone()
    finally:
        conn.close()
    return round(row["total"] or 0.0, 3)


def get_lots_for_material(material_code, location_id=None, data_file=None):
    """
    One row per real (lot, location) pairing with actual remaining
    quantity there, soonest-expiry first — the drill-down INV-US-01's
    Stock by Location/By Material views need beneath the existing
    aggregate figure, which stays unchanged.

    Deliberately does NOT filter on (or trust) the lots table's own
    static location_id column, and does NOT return one row per lot —
    a real bug, found and fixed here: that column only ever records
    where a lot was originally received. The moment any of it transfers
    (ship_transfer()/receive_transfer() both carry lot_id through), the
    lot's real stock can genuinely span more than one location, and a
    query keyed on the original receiving location alone would make a
    transferred-in lot invisible at the location it actually sits now
    — exactly the case this function exists to show correctly. Instead,
    every real (lot_id, location_id) combination that ever had a
    transaction is checked directly against the ledger, and only pairs
    with real remaining quantity are returned; the row's own
    "location_id" reflects that specific pairing, not the lot's origin.
    """
    conn = db.get_connection()
    try:
        query = ("SELECT DISTINCT lot_id, location_id FROM inventory_transactions "
                 "WHERE lot_id IN (SELECT id FROM lots WHERE material_code=?)")
        params = [material_code]
        if location_id:
            query += " AND location_id=?"
            params.append(location_id)
        pairs = conn.execute(query, params).fetchall()
    finally:
        conn.close()
    out = []
    for pair in pairs:
        remaining = get_lot_remaining_qty(pair["lot_id"], pair["location_id"], data_file)
        if abs(remaining) < 0.0005:
            continue
        lot = get_lot(pair["lot_id"], data_file)
        if lot is None:
            continue
        lot["location_id"] = pair["location_id"]
        lot["remaining_qty"] = remaining
        out.append(lot)
    out.sort(key=lambda l: l["expiry_date"] or "9999-99-99")
    return out


def trace_lot(lot_number, material_code=None, data_file=None):
    """
    Forward/backward trace for a real lot/serial number — every real
    transaction referencing it, forward and backward, never a sample.
    Returns None (an explicit not-found) if no lot genuinely matches,
    never a false empty history that could be mistaken for "this lot
    had no transactions." material_code narrows the search but is
    optional — a real recall query is sometimes "have we seen this
    serial anywhere," not always tied to a remembered material.
    """
    conn = db.get_connection()
    try:
        query = "SELECT * FROM lots WHERE lot_number=?"
        params = [lot_number]
        if material_code:
            query += " AND material_code=?"
            params.append(material_code)
        lot_rows = conn.execute(query, params).fetchall()
        if not lot_rows:
            return None
        results = []
        for lot_row in lot_rows:
            txns = conn.execute(
                "SELECT * FROM inventory_transactions WHERE lot_id=? ORDER BY txn_date, txn_id",
                (lot_row["id"],)).fetchall()
            results.append({"lot": dict(lot_row), "transactions": [dict(t) for t in txns]})
    finally:
        conn.close()
    return results


def get_near_expiry_lots(window_days=None, location_id=None, data_file=None):
    """Every lot with real remaining quantity expiring within the
    window, across every location unless one is given, soonest first —
    visible before it becomes urgent, independent of whether any pick
    or transfer is currently in progress against it."""
    window = NEAR_EXPIRY_WINDOW_DAYS if window_days is None else window_days
    cutoff = (date.today() + timedelta(days=window)).strftime("%Y-%m-%d")
    conn = db.get_connection()
    try:
        query = "SELECT * FROM lots WHERE expiry_date IS NOT NULL AND expiry_date <= ?"
        params = [cutoff]
        if location_id:
            query += " AND location_id=?"
            params.append(location_id)
        rows = conn.execute(query, params).fetchall()
    finally:
        conn.close()
    out = []
    for r in rows:
        remaining = get_lot_remaining_qty(r["id"], data_file)
        if remaining <= 0.0005:
            continue
        d = dict(r)
        d["remaining_qty"] = remaining
        d["days_to_expiry"] = (date.fromisoformat(r["expiry_date"]) - date.today()).days
        out.append(d)
    out.sort(key=lambda l: l["expiry_date"])
    return out


def suggest_fefo_lot(material_code, location_id, qty_needed, data_file=None):
    """
    Earliest-expiring lot(s) with available quantity at this location,
    covering qty_needed — a suggestion, never a silent default; the
    confirming user can pick a different lot instead, with the reason
    recorded by the caller (inventory.ship_transfer()'s own optional
    fefo_override_reason), not here. Only ever applies to Shelf-Life
    Batch lots — a Serial-tracked material has no expiry-driven
    ordering, only individual unit selection, so this returns an empty
    plan for one (the caller picks a specific serial directly instead).
    """
    lots_here = [l for l in get_lots_for_material(material_code, location_id, data_file)
                 if l["expiry_date"]]
    lots_here.sort(key=lambda l: l["expiry_date"])
    plan = []
    remaining = qty_needed
    for l in lots_here:
        if remaining <= 0.0005:
            break
        take = min(remaining, l["remaining_qty"])
        plan.append({"lot_id": l["id"], "lot_number": l["lot_number"],
                     "expiry_date": l["expiry_date"], "qty": round(take, 3)})
        remaining = round(remaining - take, 3)
    return {"plan": plan, "fully_covered": remaining <= 0.0005, "shortfall": round(max(remaining, 0), 3)}


def stats(data_file=None):
    conn = db.get_connection()
    try:
        total_lots = conn.execute("SELECT COUNT(*) c FROM lots").fetchone()["c"]
    finally:
        conn.close()
    return {"total_lots": total_lots, "near_expiry": len(get_near_expiry_lots(data_file=data_file))}


if __name__ == "__main__":
    print("Traceability stats:", stats())
