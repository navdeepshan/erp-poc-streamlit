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
