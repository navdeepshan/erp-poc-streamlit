"""
backorder.py — ATP-US-03: Manage a backorder when a reservation cannot
be fully satisfied.

Replaces the old "flagged for the Warehouse Planner" placeholder with
a real, first-class Backorder record: its own status lifecycle (Open
-> Partially Fulfilled -> Fulfilled, or Cancelled at any point before
Fulfilled), automatic re-evaluation on every real supply-arrival event,
and fulfillment strictly in the configured priority order — First-
Confirmed-First-Served by default (this PoC's own configured default;
Customer Priority Tier is a real, separate future extension, not built
here, matching ATP-US-01's own scope decision).

Every function below accepts a data_file parameter for API consistency
with the rest of this codebase, but never actually uses it for its own
database connection — matching reservation.py's own established
pattern and rationale (see that module's own docstring for the full
reasoning): this codebase has fully migrated to one single real
SQLite database, so ignoring data_file entirely here closes off the
same real bug class found and fixed multiple times this session,
rather than leaving it open for a future caller to trip over again.
"""

import db
from datetime import date


def _next_backorder_id(conn):
    rows = conn.execute("SELECT backorder_id FROM backorders").fetchall()
    mx = 0
    for r in rows:
        try:
            mx = max(mx, int(r["backorder_id"].split("-")[1]))
        except Exception:
            pass
    return f"BO-{mx+1:05d}"


def create_backorder(mat_code, mat_desc, location_id, qty, so_id, so_line_item=None,
                     data_file=None):
    """
    ATP-US-01's own un-promised quantity becomes a real Backorder —
    timestamped to the moment of the original order confirmation (this
    function is called synchronously from within that same
    confirmation flow), never to whenever the backorder happens to be
    looked at later. That real creation timestamp, plus this table's
    own auto-incrementing id as a stable tiebreak, is what
    reevaluate_backorders() below sorts on for the default First-
    Confirmed-First-Served priority.
    """
    if qty <= 0:
        raise ValueError("Backorder quantity must be positive.")

    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        backorder_id = _next_backorder_id(conn)
        conn.execute(
            "INSERT INTO backorders (backorder_id, material_code, material_desc, "
            "location_id, original_qty, open_qty, so_id, so_line_item, status, "
            "created_date) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (backorder_id, mat_code, mat_desc, location_id, qty, qty, so_id,
             so_line_item, "Open", str(date.today())))
        conn.commit()
    finally:
        conn.close()

    return {"backorder_id": backorder_id, "open_qty": qty, "status": "Open"}


def reevaluate_backorders(mat_code, location_id, data_file=None):
    """
    The real, automatic re-evaluation this story's own Business Rules
    require on every relevant supply-arrival event — called from
    goods_receipt.py's own GR posting and inventory.receive_transfer(),
    the two real events that increase real on-hand for a material at a
    Plant. Fetches every currently-active (Open or Partially Fulfilled)
    backorder for this exact (material, location), sorted strictly by
    the default First-Confirmed-First-Served priority (created_date,
    then backorder_id as a stable tiebreak for same-day confirmations),
    and attempts each in that order against reservation.
    create_reservation_up_to()'s own real, atomic reserve-up-to-
    available primitive — the same one ATP-US-01 already uses and has
    already been proven correct under real concurrent load.

    A real, honest, documented limitation: this function's own loop
    processes one backorder at a time, each against its own atomic
    reservation call, but the loop itself isn't wrapped in one single
    lock spanning every backorder. Two supply-arrival events for the
    same (material, location) arriving at genuinely the same instant
    could have their own two re-evaluation loops interleave, which
    could mean a later-arriving wave's own backorder claims real stock
    slightly ahead of an earlier-priority backorder from the other
    wave's own loop. The underlying reservation atomicity still
    guarantees the total ever reserved can never exceed real available
    supply — no double-booking is possible either way — but strict
    priority ordering *across two concurrent re-evaluation waves*
    specifically isn't fully guaranteed the way it is within one
    single wave's own loop. A real, narrower gap than ATP-US-01's own
    already-proven guarantee, not a correctness break, and left
    honestly documented rather than silently assumed away.

    Returns the list of backorders touched, each with its own real
    before/after state.
    """
    import org_defaults as od
    od.validate_atp_policy("Reservation/Backorder Priority", data_file=data_file)

    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        rows = conn.execute(
            "SELECT * FROM backorders WHERE material_code=? AND location_id=? "
            "AND status IN ('Open', 'Partially Fulfilled') "
            "ORDER BY created_date, backorder_id",
            (mat_code, location_id)).fetchall()
    finally:
        conn.close()

    import reservation as res
    touched = []
    for row in rows:
        bo = dict(row)
        if bo["open_qty"] <= 0:
            continue
        result = res.create_reservation_up_to(
            bo["material_code"], bo["material_desc"], bo["location_id"],
            bo["open_qty"], bo["so_id"], so_line_item=bo["so_line_item"],
            notes=f"ATP-US-03 backorder fulfillment ({bo['backorder_id']})",
            data_file=data_file)

        if result["reserved_qty"] <= 0:
            continue  # nothing available for this one; leave Open, try the next wave

        new_open_qty = round(bo["open_qty"] - result["reserved_qty"], 3)
        new_status = "Fulfilled" if new_open_qty <= 0.005 else "Partially Fulfilled"

        conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
        try:
            conn.execute(
                "UPDATE backorders SET open_qty=?, status=?, resolved_date=? "
                "WHERE backorder_id=?",
                (max(0.0, new_open_qty), new_status,
                 str(date.today()) if new_status == "Fulfilled" else None,
                 bo["backorder_id"]))
            conn.commit()
        finally:
            conn.close()

        touched.append({"backorder_id": bo["backorder_id"], "reserved_qty": result["reserved_qty"],
                        "new_open_qty": max(0.0, new_open_qty), "new_status": new_status,
                        "reservation_id": result["reservation_id"]})

    return touched


def get_backorder(backorder_id, data_file=None):
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        row = conn.execute("SELECT * FROM backorders WHERE backorder_id=?",
                           (backorder_id,)).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def resolve_backorder_via_shipment(backorder_id, shipped_qty, data_file=None):
    """
    A real, second way a Backorder resolves — genuinely distinct from
    reevaluate_backorders()'s own new-supply-arrival path, and a real
    gap found and fixed here directly: fulfillment.py's own
    record_shipment() ships real stock and posts a real Goods Issue,
    but until now never touched this backorder's own record at all,
    so a fully, physically delivered order's own original shortfall
    just sat here forever, Open, long after the customer actually had
    the goods.

    Deliberately does NOT create a Reservation the way reevaluate_
    backorders() does — there's nothing left to reserve; the stock
    already left the building via this same shipment's own Goods
    Issue. This function only ever reduces this backorder's own real
    open_qty by the real quantity that just shipped (capped at what
    was genuinely still open — a shipment can't resolve more backorder
    than actually exists), moving it to Fulfilled once nothing remains
    open, or leaving it Partially Fulfilled with a real, smaller
    open_qty otherwise. No-op, returning None, if the backorder is
    already Fulfilled or Cancelled — a caller doesn't need to check
    status first.
    """
    if shipped_qty <= 0:
        return None
    bo = get_backorder(backorder_id, data_file)
    if bo is None:
        raise ValueError(f"{backorder_id} not found.")
    if bo["status"] not in ("Open", "Partially Fulfilled"):
        return None

    resolved_qty = min(shipped_qty, bo["open_qty"])
    new_open_qty = round(bo["open_qty"] - resolved_qty, 3)
    new_status = "Fulfilled" if new_open_qty <= 0.005 else "Partially Fulfilled"

    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        conn.execute(
            "UPDATE backorders SET open_qty=?, status=?, resolved_date=? WHERE backorder_id=?",
            (max(0.0, new_open_qty), new_status,
             str(date.today()) if new_status == "Fulfilled" else None, backorder_id))
        conn.commit()
    finally:
        conn.close()

    return {"backorder_id": backorder_id, "resolved_qty": resolved_qty,
           "new_open_qty": max(0.0, new_open_qty), "new_status": new_status}


def cancel_backorder(backorder_id, reason="", data_file=None):
    """
    Explicit cancellation — removed from future fulfillment
    consideration entirely. A Backorder already Fulfilled cannot be
    cancelled (there's no open quantity left to cancel); a Partially
    Fulfilled one is cancelled only for its own real remaining open_qty
    — the already-fulfilled portion's own real Reservation is
    completely untouched by this, since this function never looks at
    reservations at all, only this backorder's own status.
    """
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        row = conn.execute("SELECT * FROM backorders WHERE backorder_id=?",
                           (backorder_id,)).fetchone()
    finally:
        conn.close()
    if row is None:
        raise ValueError(f"{backorder_id} not found.")
    bo = dict(row)
    if bo["status"] == "Fulfilled":
        raise ValueError(f"{backorder_id} is already Fulfilled — nothing open to cancel.")
    if bo["status"] == "Cancelled":
        raise ValueError(f"{backorder_id} is already Cancelled.")

    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        conn.execute(
            "UPDATE backorders SET status='Cancelled', resolved_date=?, "
            "resolution_notes=? WHERE backorder_id=?",
            (str(date.today()), reason, backorder_id))
        conn.commit()
    finally:
        conn.close()
    return {"backorder_id": backorder_id, "status": "Cancelled"}


def get_open_backorders(mat_code=None, location_id=None, data_file=None):
    """Every currently-active (Open or Partially Fulfilled) backorder —
    the real worklist this story's own screen shows."""
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        query = "SELECT * FROM backorders WHERE status IN ('Open', 'Partially Fulfilled')"
        params = []
        if mat_code:
            query += " AND material_code=?"
            params.append(mat_code)
        if location_id:
            query += " AND location_id=?"
            params.append(location_id)
        query += " ORDER BY created_date, backorder_id"
        rows = conn.execute(query, params).fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]


def get_all_backorders(data_file=None):
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        rows = conn.execute(
            "SELECT * FROM backorders ORDER BY created_date DESC, backorder_id DESC"
        ).fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]
