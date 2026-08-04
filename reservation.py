"""
reservation.py — ATP-US-02: Reserve promised inventory and keep
Available-to-Promise current.

A real, durable Reservation record sitting alongside On Hand and In
Transit in inventory.py's own position model — quantity, decremented
from Available-to-Promise, consumed or released explicitly, never a
promise silently overridden by a later calculation.

Honest, worth stating plainly: this document's own S2S-E08 text
(ATP-US-02) describes Reserved as one of four real co-existing states
alongside On Hand, In Transit, Quality Hold (QHD-US-01), and Pending
Disposition (RMA-US-02/QHD-US-01's own Rejected state). Quality Hold
and Pending Disposition were drafted as user stories in this same
document set but were never actually scaffolded into this PoC's own
code — get_balance() today reflects only what's genuinely built (GR
receipt, Production output, Transfer/Goods Issue, and now Reservation).
This module is built correctly against that real, current get_balance()
— not against a richer state model that exists in documentation but
not in this database. When Quality Hold and Pending Disposition are
eventually built, they would naturally need to also be excluded from
wherever "on hand" is computed, and Available-to-Promise here would
inherit that correctly for free, since it's built directly on top of
get_balance() rather than duplicating its own logic.

Real atomic check-and-reserve, not a business rule only written in
prose: create_reservation() below issues a real BEGIN IMMEDIATE before
its own first read, acquiring SQLite's own RESERVED lock on the whole
database file before the availability check runs — a second, concurrent
create_reservation() call blocks on its own BEGIN IMMEDIATE until this
one commits or rolls back, so it can never read the pre-reservation
on-hand figure and act on it. A real, honest limitation, not hidden:
SQLite's own locking is database-file-wide, not scoped to a specific
material/location, so this serializes every reservation creation
globally, not only ones genuinely competing for the same real stock —
a deliberate, reasonable tradeoff at this PoC's own scale (a handful of
concurrent demo users), not something a genuinely high-volume tenant
would want unchanged.

Every function below accepts a data_file parameter for API consistency
with the rest of this codebase, but never actually uses it for its own
database connection — db.get_connection() is always called with no
argument, matching the same safe, established pattern inventory.py,
sto.py, atp.py, accounting.py, and eway_bill.py all already use. This
is deliberate, not an oversight: this whole codebase has fully
migrated to one single real SQLite database (erp_pilot.db), so a
"which database file" parameter is no longer a meaningful concept for
anything genuinely SQLite-native. Real bugs were found and fixed this
same session, three separate times, each caused by a different
caller's own leftover Excel-era DATA_FILE constant (a real, pre-
migration path to data.xlsx, still present in many older modules for
their own, still-legitimate parameter-passing chains) being passed
straight through into a function that, back then, still honored it —
crashing with a real, cryptic "file is not a database" error. Ignoring
data_file here entirely, rather than just fixing each caller found so
far, closes this whole bug class permanently: any future caller,
including ones not yet written, can pass whatever it wants here and
this module will still always connect to the one real database that
actually exists.
"""

import db
from datetime import date


def _next_reservation_id(conn):
    rows = conn.execute("SELECT reservation_id FROM reservations").fetchall()
    mx = 0
    for r in rows:
        try:
            mx = max(mx, int(r["reservation_id"].split("-")[1]))
        except Exception:
            pass
    return f"RES-{mx+1:05d}"


def create_reservation(mat_code, mat_desc, location_id, qty, so_id, so_line_item=None,
                       notes="", data_file=None):
    """
    The real atomic check-and-reserve operation. Real on-hand (summed
    directly from inventory_transactions, the same real figure
    inventory.get_balance() computes, but read within this same locked
    transaction rather than a separate connection, for genuine
    correctness rather than assumed safety) minus every currently-Open
    reservation at this material/location is this material/location's
    own real Available-to-Promise at the instant this runs. A request
    exceeding that is rejected outright, never created partially or
    against a stale figure.
    """
    if qty <= 0:
        raise ValueError("Reservation quantity must be positive.")

    fpath = data_file
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        conn.execute("BEGIN IMMEDIATE")

        row = conn.execute(
            "SELECT SUM(quantity) AS total FROM inventory_transactions "
            "WHERE material_code=? AND location_id=?",
            (mat_code, location_id)).fetchone()
        on_hand = round(row["total"] or 0.0, 3)

        row = conn.execute(
            "SELECT COALESCE(SUM(quantity), 0) AS reserved FROM reservations "
            "WHERE material_code=? AND location_id=? AND status='Open'",
            (mat_code, location_id)).fetchone()
        already_reserved = round(row["reserved"] or 0.0, 3)

        available_to_promise = on_hand - already_reserved

        if qty > available_to_promise + 0.005:
            conn.rollback()
            raise ValueError(
                f"Cannot reserve {qty:g} of {mat_code} at {location_id} — only "
                f"{available_to_promise:g} genuinely available to promise "
                f"(on-hand {on_hand:g}, {already_reserved:g} already reserved).")

        reservation_id = _next_reservation_id(conn)
        conn.execute(
            "INSERT INTO reservations (reservation_id, material_code, material_desc, "
            "location_id, quantity, so_id, so_line_item, status, created_date, "
            "resolution_notes) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (reservation_id, mat_code, mat_desc, location_id, qty, so_id, so_line_item,
             "Open", str(date.today()), notes))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return {"reservation_id": reservation_id, "material_code": mat_code,
           "location_id": location_id, "quantity": qty, "status": "Open"}


def create_reservation_up_to(mat_code, mat_desc, location_id, requested_qty, so_id,
                             so_line_item=None, notes="", data_file=None):
    """
    ATP-US-01's own real atomic primitive for a partial promise: within
    one single BEGIN IMMEDIATE transaction (not two separate calls to
    create_reservation() with a real race window between a failed
    full-quantity attempt and a fallback partial one), reserves
    min(requested_qty, real current Available-to-Promise) and reports
    exactly what was reserved and what wasn't. A request that can be
    fully covered still creates exactly one reservation for the full
    amount, same as create_reservation() itself; the only real
    difference is this function never rejects outright for a
    partially-coverable request -- it reserves what's genuinely there
    and reports the shortfall honestly, rather than raising and
    forcing the caller into a second, separately-racy attempt.

    Returns {"reserved_qty", "shortfall_qty", "reservation_id" (None
    if reserved_qty is 0)}.
    """
    if requested_qty <= 0:
        raise ValueError("Requested quantity must be positive.")

    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        conn.execute("BEGIN IMMEDIATE")

        row = conn.execute(
            "SELECT SUM(quantity) AS total FROM inventory_transactions "
            "WHERE material_code=? AND location_id=?",
            (mat_code, location_id)).fetchone()
        on_hand = round(row["total"] or 0.0, 3)

        row = conn.execute(
            "SELECT COALESCE(SUM(quantity), 0) AS reserved FROM reservations "
            "WHERE material_code=? AND location_id=? AND status='Open'",
            (mat_code, location_id)).fetchone()
        already_reserved = round(row["reserved"] or 0.0, 3)

        available = max(0.0, round(on_hand - already_reserved, 3))
        to_reserve = round(min(requested_qty, available), 3)
        shortfall = round(requested_qty - to_reserve, 3)

        reservation_id = None
        if to_reserve > 0:
            reservation_id = _next_reservation_id(conn)
            conn.execute(
                "INSERT INTO reservations (reservation_id, material_code, material_desc, "
                "location_id, quantity, so_id, so_line_item, status, created_date, "
                "resolution_notes) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (reservation_id, mat_code, mat_desc, location_id, to_reserve, so_id,
                 so_line_item, "Open", str(date.today()), notes))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return {"reserved_qty": to_reserve, "shortfall_qty": shortfall,
           "reservation_id": reservation_id}


def get_reservation(reservation_id, data_file=None):
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        row = conn.execute("SELECT * FROM reservations WHERE reservation_id=?",
                           (reservation_id,)).fetchone()
    finally:
        conn.close()
    return dict(row) if row else None


def consume_reservation(reservation_id, data_file=None):
    """
    A pick actually occurred against this reservation — moves Open to
    Consumed. Never re-opened afterward; a genuine correction (a pick
    reversed) is a new reservation, not a reanimation of this one.
    Rejected, with a specific reason, against anything not currently
    Open — a Released reservation no longer represents a real promise,
    and a reservation already Consumed cannot be consumed twice.

    Deliberately posts no inventory movement of its own — this function
    is purely the reservation's own status transition, matching the
    story's own "does not change picking mechanics, only what it
    consumes." In the real, complete flow, a caller (a future pick/
    FUL-US-02 implementation) calls this at the same moment it also
    posts the real Goods Issue reducing on-hand by the same quantity,
    so Available-to-Promise nets to unchanged overall (reserved drops,
    on-hand drops by the same amount). Called on its own, without that
    accompanying pick, Available-to-Promise will genuinely rise — the
    quantity is no longer reserved, and nothing has yet reduced on-hand
    to account for it. That's correct given what this function alone
    is responsible for; a real caller is expected to do both together.
    """
    r = get_reservation(reservation_id, data_file)
    if r is None:
        raise ValueError(f"{reservation_id} not found.")
    if r["status"] != "Open":
        raise ValueError(f"{reservation_id} is '{r['status']}' — only an Open "
                         f"reservation can be consumed.")

    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        conn.execute(
            "UPDATE reservations SET status='Consumed', resolved_date=?, "
            "resolution_type='Consumed' WHERE reservation_id=?",
            (str(date.today()), reservation_id))
        conn.commit()
    finally:
        conn.close()
    return {"reservation_id": reservation_id, "status": "Consumed"}


def release_reservation(reservation_id, reason="", data_file=None):
    """
    A Sales Order line was cancelled, or this reservation can no longer
    be fulfilled — moves Open to Released, and its quantity returns to
    Available-to-Promise the instant this commits (get_available_to_
    promise() below only ever sums Open reservations, so a Released one
    stops counting against it immediately, with no separate step).
    Rejected against an already-Consumed reservation — real,
    physically-picked stock cannot be un-consumed by a release.
    """
    r = get_reservation(reservation_id, data_file)
    if r is None:
        raise ValueError(f"{reservation_id} not found.")
    if r["status"] == "Consumed":
        raise ValueError(f"{reservation_id} is already Consumed — a consumed "
                         f"reservation reflects real, physically-picked stock "
                         f"and cannot be released.")
    if r["status"] == "Released":
        raise ValueError(f"{reservation_id} is already Released.")

    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        conn.execute(
            "UPDATE reservations SET status='Released', resolved_date=?, "
            "resolution_type='Released', resolution_notes=? WHERE reservation_id=?",
            (str(date.today()), reason, reservation_id))
        conn.commit()
    finally:
        conn.close()
    return {"reservation_id": reservation_id, "status": "Released"}


def get_available_to_promise(mat_code, location_id, data_file=None):
    """
    Real on-hand minus every currently-Open reservation — recalculated
    live on every call, never cached stale. Not wrapped in the same
    BEGIN IMMEDIATE lock create_reservation() uses, since a plain read
    doesn't need one; two of these running concurrently with a
    create_reservation() simply see a consistent snapshot before or
    after that write, never a torn one.
    """
    import inventory as inv
    on_hand = inv.get_balance(mat_code, location_id, data_file)
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        row = conn.execute(
            "SELECT COALESCE(SUM(quantity), 0) AS reserved FROM reservations "
            "WHERE material_code=? AND location_id=? AND status='Open'",
            (mat_code, location_id)).fetchone()
    finally:
        conn.close()
    reserved = round(row["reserved"] or 0.0, 3)
    return round(on_hand - reserved, 3)


def get_reserved_quantity(mat_code, location_id, data_file=None):
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        row = conn.execute(
            "SELECT COALESCE(SUM(quantity), 0) AS reserved FROM reservations "
            "WHERE material_code=? AND location_id=? AND status='Open'",
            (mat_code, location_id)).fetchone()
    finally:
        conn.close()
    return round(row["reserved"] or 0.0, 3)


def get_open_reservations(mat_code=None, location_id=None, data_file=None):
    """The real Reservation Ledger view — every currently-Open
    reservation, optionally filtered to one material and/or location."""
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        query = "SELECT * FROM reservations WHERE status='Open'"
        params = []
        if mat_code:
            query += " AND material_code=?"
            params.append(mat_code)
        if location_id:
            query += " AND location_id=?"
            params.append(location_id)
        query += " ORDER BY created_date, reservation_id"
        rows = conn.execute(query, params).fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]


def get_all_reservations(data_file=None):
    conn = db.get_connection()  # data_file intentionally ignored -- see module docstring
    try:
        rows = conn.execute(
            "SELECT * FROM reservations ORDER BY created_date DESC, reservation_id DESC"
        ).fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]
