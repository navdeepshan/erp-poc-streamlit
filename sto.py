"""
sto.py — Stock Transfer Order creation from an approved Hub allocation.

Real STO-US-02 scaffolding: auto-creates STOs from an approved
allocation, and tags every resulting real transfer as STO-sourced
rather than ad hoc. The "approved allocation" this consumes is
deliberately a stand-in for the future full STO-US-01 (AI-driven
recommendation + human review/override + Operations Approver sign-off,
including its own OCR/scan path) — not a replacement for it, and not
pretending to be one. `simple_allocate_and_create_sto()` wraps the
SAME real transfer-opportunity engine INV-US-03 already uses (largest-
shortage-first, when a Hub can't fully satisfy every requesting Plant)
and treats its output as already-approved, skipping the review/
override/sign-off workflow entirely. Swapping in the real STO-US-01
later means replacing only this one function's own source of
allocation lines — `create_sto_from_allocation()` itself already takes
real, already-decided lines and doesn't care where they came from.
"""

import db
import bom
import inventory as inv
from datetime import date

# Simulated internal logistics rate for Hub-to-Plant STO freight, not
# verified against any real carrier tariff (no real transport vendor
# contract to check this against) -- a flat per-shipment base plus a
# per-unit component, the same "real shape, not verified accurate"
# honesty standard as shipping.py's own courier rate cards.
STO_FREIGHT_BASE_CHARGE = 500    # INR, per STO
STO_FREIGHT_PER_UNIT = 20        # INR, per allocated unit


def _next_sto_id(conn):
    rows = conn.execute("SELECT sto_id FROM sto_header").fetchall()
    mx = 0
    for r in rows:
        try:
            mx = max(mx, int(r["sto_id"].split("-")[1]))
        except Exception:
            pass
    return f"STO-{mx+1:05d}"


def create_sto_from_allocation(mat_code, mat_desc, hub_location, allocation_lines,
                               allocation_rule="Simple", created_by="", notes="",
                               data_file=None):
    """
    allocation_lines: [{"to_location": ..., "requested_qty": ...,
    "allocated_qty": ...}, ...] — already-decided, already-approved
    lines, from wherever they came from (this module's own
    simple_allocate_and_create_sto() below, or a real future
    STO-US-01). This function's own job is purely mechanical: create
    the STO header/lines, and for every line with a real allocated_qty
    > 0, ship a real stock_transfers record via inventory.
    ship_transfer() — the same real Ship mechanics INV-US-05 already
    owns, not a separate STO-only shipping path — tagged source_type=
    'STO' / source_doc=this STO's own id, distinguishing it from an ad
    hoc transfer in every list that already carries that column.

    A line with allocated_qty of 0 (a real requester whose own
    shortage wasn't covered at all, e.g. the Hub genuinely didn't have
    enough) still gets a real sto_lines row — visible on the STO
    itself as an honest zero, not silently dropped.
    """
    conn = db.get_connection()
    try:
        sto_id = _next_sto_id(conn)
    finally:
        conn.close()

    total_qty = sum(l["allocated_qty"] for l in allocation_lines)

    conn = db.get_connection()
    try:
        conn.execute(
            "INSERT INTO sto_header (sto_id, material_code, material_desc, hub_location, "
            "total_qty, allocation_rule, created_date, created_by, notes) "
            "VALUES (?,?,?,?,?,?,?,?,?)",
            (sto_id, mat_code, mat_desc, hub_location, total_qty, allocation_rule,
             str(date.today()), created_by, notes))
        conn.commit()
    finally:
        conn.close()

    results = []
    for i, line in enumerate(allocation_lines, 1):
        transfer_id = None
        if line["allocated_qty"] > 0:
            # ship_transfer() itself now generates a real e-way bill automatically,
            # whenever the real rule requires one -- this used to be duplicated
            # here, a real gap found directly: an ad hoc shipment never got the
            # same treatment. One real implementation now, reached by every
            # caller, not two that could drift.
            result = inv.ship_transfer(mat_code, mat_desc, hub_location,
                line["to_location"], line["allocated_qty"], shipped_by=created_by,
                notes=f"STO {sto_id}", data_file=data_file)
            transfer_id = result["transfer_id"]
            conn = db.get_connection()
            try:
                conn.execute(
                    "UPDATE stock_transfers SET source_type='STO', source_doc=? "
                    "WHERE transfer_id=?", (sto_id, transfer_id))
                conn.commit()
            finally:
                conn.close()

        conn = db.get_connection()
        try:
            conn.execute(
                "INSERT INTO sto_lines (sto_id, line_no, to_location, requested_qty, "
                "allocated_qty, transfer_id) VALUES (?,?,?,?,?,?)",
                (sto_id, i, line["to_location"], line["requested_qty"],
                 line["allocated_qty"], transfer_id))
            conn.commit()
        finally:
            conn.close()
        results.append({"to_location": line["to_location"], "allocated_qty": line["allocated_qty"],
                        "transfer_id": transfer_id})

    # Real freight GL posting, scoped deliberately narrow: a self-contained,
    # balanced freight entry (Dr 5100 Freight Expense / Cr 2110 Freight
    # Accrual), not the fuller 1300 Inter-Plant Stock-in-Transit / 1200
    # Inventory posting the reviewed story also describes. Deliberately not
    # built here — this system's own inventory transfers have never posted
    # to the GL at all (INV-US-05's own "no accounting entry" principle for
    # a movement between the org's own locations), so introducing a general
    # inventory-asset GL account for the first time, only for STO freight,
    # with nothing else in this system feeding or reconciling against it,
    # would be a disconnected entry rather than a real one. The freight
    # entry below doesn't have that problem — it's genuinely self-contained.
    shipped_lines = [r for r in results if r["transfer_id"]]
    if shipped_lines:
        total_freight = round(STO_FREIGHT_BASE_CHARGE + STO_FREIGHT_PER_UNIT * total_qty, 2)
        je_lines = [{"account_code": "5100", "debit": total_freight, "credit": 0,
                    "description": f"{sto_id} freight — {hub_location} to "
                                   f"{len(shipped_lines)} Plant(s)"}]
        allocated = 0
        for idx, r in enumerate(shipped_lines):
            is_last = idx == len(shipped_lines) - 1
            # Remainder to the last line, same rounding discipline used
            # elsewhere in this codebase for proportional allocation —
            # never let independent per-line rounding leave the entry
            # unbalanced by a paisa or two.
            share = round(total_freight - allocated, 2) if is_last else \
                round(total_freight * (r["allocated_qty"] / total_qty), 2)
            allocated += share
            je_lines.append({"account_code": "2110", "debit": 0, "credit": share,
                            "description": f"{sto_id} line to {r['to_location']} "
                                           f"({r['allocated_qty']:g} units)"})
        import accounting
        accounting.post_journal_entry("STO", sto_id,
            f"{sto_id} — freight accrual for {mat_desc}", je_lines, data_file=data_file)

    return {"sto_id": sto_id, "lines": results}


def simple_allocate_and_create_sto(mat_code, hub_location, created_by="", data_file=None):
    """
    The real, honest stand-in for the future STO-US-01: reuses
    INV-US-03's own get_transfer_opportunities() (the same largest-
    shortage-first engine already proven there, real not-enough-supply
    handling included) for every requesting Plant currently short of
    this material with real, available stock at this Hub, and treats
    that output as already-approved — no Draft/Review/Override/
    Approve step, no confidence indicator, no persona split. A real
    future STO-US-01 replaces only this function's own source of
    allocation lines; create_sto_from_allocation() above doesn't
    change at all when that happens.
    """
    opportunities = bom.get_transfer_opportunities(data_file=data_file)
    matching = [o for o in opportunities if o["mat_code"] == mat_code
               and o["from_location"] == hub_location]
    if not matching:
        raise ValueError(f"No real transfer opportunity found for {mat_code} from "
                         f"{hub_location} — nothing to allocate.")

    mat_desc = matching[0]["mat_desc"]
    allocation_lines = [{"to_location": o["to_location"], "requested_qty": o["shortage_at_destination"],
                        "allocated_qty": o["suggested_qty"]} for o in matching]

    return create_sto_from_allocation(mat_code, mat_desc, hub_location, allocation_lines,
                                      allocation_rule="Simple (stand-in for STO-US-01)",
                                      created_by=created_by, data_file=data_file)


def get_sto(sto_id, data_file=None):
    conn = db.get_connection()
    try:
        header = conn.execute("SELECT * FROM sto_header WHERE sto_id=?", (sto_id,)).fetchone()
        if header is None:
            return None
        lines = conn.execute("SELECT * FROM sto_lines WHERE sto_id=? ORDER BY line_no",
                             (sto_id,)).fetchall()
    finally:
        conn.close()
    return {"header": dict(header), "lines": [dict(l) for l in lines]}


def get_all_stos(data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT * FROM sto_header ORDER BY sto_id DESC").fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]
