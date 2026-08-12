"""
atp.py — ATP-US-01: Check real availability and promise a confirmed
Sales Order line.

Scoped to this PoC's own configured defaults, matching the story's own
Default Configuration Values table exactly:
  - Concurrency Guarantee Level: Strict (reservation.
    create_reservation_up_to()'s own real BEGIN IMMEDIATE atomic
    guarantee, proven under real concurrent load, unchanged here).
  - ATP Sourcing Scope: Single-Plant only. Network ATP (consulting
    INV-US-03's own transfer network for a sourcing Plant when the
    assigned Plant alone can't cover a line) is a real, separate
    future extension, deliberately not built in this pass.
  - Reservation/Backorder Priority: First-Confirmed-First-Served is
    implicit in this check's own synchronous, atomic nature — whichever
    order's own confirmation call reaches create_reservation_up_to()
    first genuinely gets the stock, proven directly by the same
    concurrency stress testing reservation.py's own primitive already
    passed. A Customer Priority Tier override is a real, separate
    future extension, not built here.
  - Reservation Granularity: quantity-only (no Batch/Shelf-Life lot
    binding exists in this PoC to defer in the first place).
  - Reservation Visibility to Planning: handled separately, in
    bom.py's own transfer-opportunity calculation, not in this module.
"""

import reservation as res
import backorder as bo


def check_and_promise_line(so_id, line_item, mat_code, mat_desc, location_id, qty,
                           data_file=None):
    """
    The real, single atomic check-and-promise operation for one Sales
    Order line — one call into reservation.create_reservation_up_to(),
    never a separate check-then-reserve sequence that could race
    against a concurrent order for the same stock.

    Returns a real ATP outcome:
      - Promised: the full requested quantity was genuinely available
        and is now reserved.
      - Partially Promised: only part of it was; that part is reserved,
        the remainder becomes a real ATP-US-03 Backorder (created here,
        not left as a bare number this function reports and forgets).
      - Backordered: none of it was available; the full quantity
        becomes a real Backorder, with no reservation created at all.
    """
    if qty <= 0:
        raise ValueError("ATP check requires a positive quantity.")

    import org_defaults as od
    od.validate_atp_policy("ATP Sourcing Scope", data_file=data_file)

    result = res.create_reservation_up_to(mat_code, mat_desc, location_id, qty,
                                          so_id, so_line_item=line_item,
                                          notes=f"ATP-US-01 check at {so_id} confirmation",
                                          data_file=data_file)

    if result["shortfall_qty"] <= 0.005:
        outcome = "Promised"
    elif result["reserved_qty"] <= 0.005:
        outcome = "Backordered"
    else:
        outcome = "Partially Promised"

    backorder_id = None
    if result["shortfall_qty"] > 0.005:
        bo_result = bo.create_backorder(mat_code, mat_desc, location_id,
                                        result["shortfall_qty"], so_id,
                                        so_line_item=line_item, data_file=data_file)
        backorder_id = bo_result["backorder_id"]

    return {
        "outcome": outcome,
        "promised_qty": result["reserved_qty"],
        "backordered_qty": result["shortfall_qty"],
        "reservation_id": result["reservation_id"],
        "backorder_id": backorder_id,
    }


def get_live_line_status(so_item, data_file=None):
    """
    The real, current ATP state for one Sales Order line — not the
    static snapshot check_and_promise_line() wrote once, at
    confirmation. A real gap found and fixed here directly: nothing
    ever revisited that snapshot afterward, so a line whose own real
    Backorder had since been fully or partially fulfilled (by
    reevaluate_backorders() on new supply arrival, or by a real
    shipment resolving it directly via fulfillment.py's own record_
    shipment()) still displayed its own original, stale outcome
    indefinitely. Mirrors exactly how bom.py's own planning fix
    already reads live Backorder state rather than the same stale
    columns, for the same real reason.

    so_item: one dict as returned by sales_order.get_order_items() --
    needs its own qty, atp_outcome, and backorder_id.

    Falls back to the line's own original, static values for a pre-ATP
    order (atp_outcome is None), the same real fallback check_and_
    promise_line()'s own caller already applies elsewhere.
    """
    if so_item.get("atp_outcome") is None:
        return {"outcome": None, "promised_qty": None, "backordered_qty": None}

    if not so_item.get("backorder_id"):
        # Fully Promised at confirmation, nothing ever backordered --
        # still accurate; there's no Backorder record to have since
        # moved.
        return {"outcome": "Promised", "promised_qty": so_item["qty"], "backordered_qty": 0.0}

    live_bo = bo.get_backorder(so_item["backorder_id"], data_file=data_file)
    live_backordered = (live_bo["open_qty"]
                        if live_bo and live_bo["status"] in ("Open", "Partially Fulfilled")
                        else 0.0)
    live_promised = round(so_item["qty"] - live_backordered, 3)

    if live_backordered <= 0.005:
        outcome = "Promised"
    elif live_promised <= 0.005:
        outcome = "Backordered"
    else:
        outcome = "Partially Promised"

    return {"outcome": outcome, "promised_qty": live_promised, "backordered_qty": live_backordered}
