"""
rtv.py — RTV-US-01: Return already-received, quality-failed stock to the vendor.

The physical second half of a Quality Hold disposed "Return to Vendor"
(quality_inspection.dispose_quality_hold()) — nothing in this module
creates or disposes a hold itself, it only acts on one already sitting
in "Pending RTV Shipment" state. Explicitly distinct from goods_
receipt.py's own existing at-dock over-receipt "return to vendor" rule
(a same-day rejection before the goods are ever put away), which is
unchanged — this is for stock that was already received, shelved, and
only found to have failed inspection afterward.

Deliberately doesn't post any GL entry at disposition time — quality_
inspection.dispose_quality_hold() already doesn't, for the identical
reason: the real financial consequence belongs to the moment the
physical shipment actually happens, not the moment someone decides it
should. Mirrors GR/IR Clearing's own "don't assume resolution
prematurely" pattern exactly:

  Dr RTV Clearing (1310) / Cr Inventory Clearing (1200)

at the GR's own originally-received unit cost — a clearing/suspense
entry, left open pending the vendor's own eventual credit note or
replacement, which is genuinely out of scope here (the same way Vendor
Invoice verification is referenced but not built by goods_receipt.py).

No inventory quantity movement posts here — the held quantity already
left available on-hand the moment quality_inspection.record_inspection()
posted the original Quality Hold transaction; this module's only
ledger consequence is the GL entry above.

E-way bill: reuses eway_bill.py's own real threshold/distance-validity
constants (EWAY_BILL_VALUE_THRESHOLD, KM_PER_VALIDITY_DAY, its haversine
helper) rather than duplicating them, but can't call eway_bill.py's own
is_eway_bill_required()/generate_eway_bill() directly — both are
hardcoded to resolve *both* ends of a movement from this org's own
Delivery_Locations, and a vendor is never one of those. The vendor's
own state is resolved the same way shipping.py's build_customer_
shipment_details() already resolves a customer's — from vendor_
onboarding.validate_gstin()'s own state-code lookup, not a second,
free-text-parsed one.

New table: rtv_shipments (see db.py) — one row per physical return
shipment, one hold per shipment. Consolidating several holds bound for
the same vendor into one shipment, the way shipping.py's own route
consolidation does for couriers, is a real future enhancement, not
built here.
"""

import math
import random
from datetime import date, timedelta

import db
import quality_inspection as qi
import pr_consolidation as pc
import vendor_onboarding as vo
import goods_receipt as gr
import eway_bill as ewb


def _next_rtv_id(conn):
    rows = conn.execute("SELECT rtv_id FROM rtv_shipments WHERE rtv_id LIKE 'RTV-%'").fetchall()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["rtv_id"].split("-")[1]))
        except Exception: pass
    return f"RTV-{mx+1:05d}"


def get_pending_rtv_holds(data_file=None):
    """Quality Holds disposed 'Return to Vendor' but not yet physically shipped."""
    return qi.get_quality_holds(status="Pending RTV Shipment", data_file=data_file)


def get_rtv_shipments(data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT * FROM rtv_shipments ORDER BY rtv_id").fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]


def get_rtv_shipment(rtv_id, data_file=None):
    for r in get_rtv_shipments(data_file=data_file):
        if r["rtv_id"] == rtv_id:
            return r
    return None


def _vendor_for_hold(hold, data_file=None):
    """The real vendor this hold's GR was received against — a hold
    doesn't carry vendor_id directly, so this reads it off the GR
    header rather than duplicating it onto every hold row."""
    g = gr.get_gr(hold["gr_id"], data_file)
    if g is None:
        raise ValueError(f"{hold['gr_id']} not found for {hold['hold_id']}.")
    vendor = vo.get_vendor(g["vendor_id"], data_file)
    if vendor is None:
        raise ValueError(f"Vendor {g['vendor_id']} not found for {hold['gr_id']}.")
    return vendor


def _rtv_eway_bill(from_location, vendor, declared_value, data_file=None):
    """
    Same real rule shape as eway_bill.is_eway_bill_required()/
    generate_eway_bill() — inter-state + value threshold, real
    haversine distance — adapted for a vendor destination. Returns
    None if no e-way bill is required for this return.
    """
    locs = {l["id"]: l for l in pc.get_delivery_locations(active_only=False)}
    from_loc = locs.get(from_location, {})
    from_state = from_loc.get("state")

    to_state = None
    if vendor.get("GSTIN"):
        ok, _, details = vo.validate_gstin(vendor["GSTIN"])
        if ok:
            to_state = details["state_name"]

    inter_state = bool(from_state and to_state and from_state != to_state)
    if not inter_state or declared_value < ewb.EWAY_BILL_VALUE_THRESHOLD:
        return None

    distance_km = round(ewb._haversine_km(from_loc.get("geo", "0,0"), vendor.get("Geolocation", "0,0")))
    validity_days = max(1, math.ceil(distance_km / ewb.KM_PER_VALIDITY_DAY))
    generated_date = date.today()
    valid_until = generated_date + timedelta(days=validity_days)
    ewb_number = f"{random.randint(10**11, 10**12 - 1)}"
    return {"ewb_number": ewb_number, "generated_date": str(generated_date),
            "valid_until": str(valid_until), "distance_km": distance_km}


def ship_return_to_vendor(hold_id, shipped_by="", notes="", data_file=None):
    """
    Executes the real physical return: generates an e-way bill where
    the real rule requires one, posts the real GL entry (Dr RTV
    Clearing / Cr Inventory Clearing at the GR's own received cost),
    and marks the hold Returned to Vendor. Refuses anything not
    currently 'Pending RTV Shipment' — there's no honest shipment to
    execute otherwise.
    """
    hold = qi.get_quality_hold(hold_id, data_file)
    if hold is None:
        raise ValueError(f"{hold_id} not found.")
    if hold["status"] != "Pending RTV Shipment":
        raise ValueError(f"{hold_id} is '{hold['status']}' — only a hold disposed "
                          f"'Return to Vendor' and awaiting shipment can ship.")

    vendor = _vendor_for_hold(hold, data_file)
    gr_items = {i["po_item"]: i for i in gr.get_gr_items(hold["gr_id"], data_file)}
    unit_price = (gr_items.get(hold["po_item"]) or {}).get("unit_price") or 0
    declared_value = round(unit_price * hold["qty"], 2)

    ewb_result = _rtv_eway_bill(hold["location_id"], vendor, declared_value, data_file)

    je_id = None
    if declared_value > 0:
        import accounting as acct
        je_id = acct.post_journal_entry(
            "RTV", hold_id,
            f"Return to vendor — {hold['mat_desc']} ({hold['qty']:g} units) to "
            f"{vendor['Vendor_Name']}",
            [{"account_code": "1310", "debit": declared_value, "credit": 0,
              "description": f"RTV Clearing — {hold['mat_desc']} to {vendor['Vendor_Name']}"},
             {"account_code": "1200", "debit": 0, "credit": declared_value,
              "description": f"Inventory Clearing relieved — {hold_id}"}],
            data_file=data_file)

    conn = db.get_connection()
    try:
        rtv_id = _next_rtv_id(conn)
        conn.execute(
            "INSERT INTO rtv_shipments (rtv_id, hold_id, material_code, material_desc, qty, "
            "location_id, vendor_id, vendor_name, shipped_date, shipped_by, eway_bill_number, "
            "eway_bill_valid_until, je_id, notes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (rtv_id, hold_id, hold["mat_code"], hold["mat_desc"], hold["qty"],
             hold["location_id"], vendor["Vendor_ID"], vendor["Vendor_Name"],
             date.today().strftime("%Y-%m-%d"), shipped_by,
             (ewb_result or {}).get("ewb_number"), (ewb_result or {}).get("valid_until"),
             je_id, notes),
        )
        conn.execute("UPDATE quality_holds SET status='Returned to Vendor' WHERE hold_id=?",
                     (hold_id,))
        conn.commit()
    finally:
        conn.close()

    return {"rtv_id": rtv_id, "je_id": je_id, "eway_bill_number": (ewb_result or {}).get("ewb_number"),
            "eway_bill_valid_until": (ewb_result or {}).get("valid_until")}


def stats(data_file=None):
    pending = get_pending_rtv_holds(data_file)
    shipped = get_rtv_shipments(data_file)
    return {"pending": len(pending), "shipped": len(shipped)}


if __name__ == "__main__":
    print("RTV stats:", stats())
