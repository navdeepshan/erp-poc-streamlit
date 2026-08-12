"""
control_tower.py — cross-cutting exception aggregation, v1.

Not a new business capability — a read-only aggregation layer over
exception-generating queries that already exist, scattered today across
three separate apps (S2C/MFG/O2C) that this project's own CLAUDE.md
deliberately keeps un-merged. This module never mutates anything; every
adapter below calls an existing module's existing read function and maps
its own natural shape into one common exception record. Fixing an
exception still happens on its owning screen — this is a worklist, not
a second place business logic lives.

Same "kind registry" pattern already proven in shipping.py (_KINDS) —
one shared aggregator, N small adapters, each owning its own domain's
severity judgment. The tower itself doesn't know backorder aging rules
or AR overdue thresholds; each adapter does, since that's exactly the
kind of domain knowledge that belongs with the module that already owns
the underlying data, not duplicated here.

Adding a new source is one new adapter function plus one line in
_SOURCES — never a change to the adapters already here, and never a
change to the aggregation/sort logic below. RMA's own source
(_rma_disposition_source) was added exactly this way (2026-08-09),
with zero changes to the nine already here — proving the pattern
scales the way it was designed to.

Exception record shape (every adapter returns a list of these):
    id            the real business-object ID (backorder_id, hold_id, po_number, ...)
    category      human-readable grouping, e.g. "Backorder", "Credit Hold"
    severity      "Critical" | "Attention" | "Info" — computed by the adapter
    title         one-line human summary, already including the real numbers
    age_days      real days since the exception's own creation/due event
    owner_app     "S2C" | "MFG" | "O2C" — which app owns the fix
    owner_screen  where in that app to go act on it
    source_module which module this came from, for traceability

v1 is deliberately read-only text navigation (owner_app/owner_screen
shown as instructions) rather than clickable cross-app deep links —
real clickability needs each of the three existing apps to read a
launch query param, a small change to already-working code, correctly
sequenced *after* this registry pattern is proven, not before.
"""

from datetime import date

SEVERITY_ORDER = {"Critical": 0, "Attention": 1, "Info": 2}


def _age_days(date_str):
    """Days since a YYYY-MM-DD date string, the one date format used
    consistently everywhere in this codebase. Returns 0 for a missing
    or unparseable date rather than crashing the whole tower over one
    bad row — an exception that can't be dated still deserves to show
    up, just without an age-based severity escalation."""
    if not date_str:
        return 0
    try:
        y, m, d = (int(x) for x in str(date_str)[:10].split("-"))
        return max(0, (date.today() - date(y, m, d)).days)
    except Exception:
        return 0


# ── Adapters — one per existing exception source ────────────────────────────
def _backorder_source(data_file=None):
    import backorder as bo
    out = []
    for b in bo.get_open_backorders(data_file=data_file):
        age = _age_days(b["created_date"])
        out.append({
            "id": b["backorder_id"], "category": "Backorder",
            "severity": "Critical" if age > 14 else "Attention",
            "title": f"{b['open_qty']:g} of {b['material_desc']} "
                     f"backordered for {b['so_id']} — {age}d open",
            "age_days": age, "owner_app": "O2C", "owner_screen": "Sales Orders → Backorders",
            "source_module": "backorder",
        })
    return out


def _at_risk_source(data_file=None):
    import bom
    out = []
    for r in bom.get_procurement_recommendations(data_file=data_file):
        if r["outcome"] != "At Risk":
            continue
        out.append({
            "id": f"{r['mat_code']}@{r['location']}", "category": "At-Risk Shortage",
            "severity": "Critical",
            "title": f"{r['mat_desc']} short {r['remaining_gap']:g} at {r['location']} — "
                     f"stockout {r['stockout_date']}, normal lead time can't beat it",
            "age_days": max(0, -r["days_until_stockout"]) if r["days_until_stockout"] < 0
                        else 0,
            "owner_app": "MFG", "owner_screen": "Inventory → Time-Phased Planning",
            "source_module": "bom",
        })
    return out


def _quality_hold_source(data_file=None):
    import quality_inspection as qi
    out = []
    for h in qi.get_quality_holds(status="Held", data_file=data_file):
        age = _age_days(h["created_date"])
        out.append({
            "id": h["hold_id"], "category": "Quality Hold — Undisposed",
            "severity": "Critical" if age > 7 else "Attention",
            "title": f"{h['qty']:g} of {h['mat_desc']} held at {h['location_id']} — "
                     f"{age}d awaiting Scrap/Return-to-Vendor decision",
            "age_days": age, "owner_app": "MFG",
            "owner_screen": "Quality Inspection → Quality Holds & RTV",
            "source_module": "quality_inspection",
        })
    return out


def _pending_rtv_source(data_file=None):
    import rtv
    out = []
    for h in rtv.get_pending_rtv_holds(data_file=data_file):
        age = _age_days(h["disposed_date"] or h["created_date"])
        out.append({
            "id": h["hold_id"], "category": "RTV Awaiting Shipment",
            "severity": "Critical" if age > 21 else "Attention",
            "title": f"{h['qty']:g} of {h['mat_desc']} disposed Return-to-Vendor — "
                     f"{age}d not yet physically shipped",
            "age_days": age, "owner_app": "MFG",
            "owner_screen": "Quality Inspection → Quality Holds & RTV",
            "source_module": "rtv",
        })
    return out


def _gr_pending_inspection_source(data_file=None):
    import quality_inspection as qi
    import goods_receipt as gr
    out = []
    for g in qi.get_grs_needing_inspection(data_file=data_file):
        header = gr.get_gr(g["gr_id"], data_file)
        age = _age_days(header["gr_date"]) if header else 0
        out.append({
            "id": g["gr_id"], "category": "GR Pending Inspection",
            "severity": "Critical" if age > 5 else "Attention",
            "title": f"{g['gr_id']} ({g['vendor_name']}) — {g['pending_lines']} of {g['lines']} "
                     f"line(s) not yet inspected, {age}d since receipt",
            "age_days": age, "owner_app": "MFG",
            "owner_screen": "Quality Inspection → Record Inspection",
            "source_module": "quality_inspection",
        })
    return out


def _po_awaiting_receipt_source(data_file=None):
    import goods_receipt as gr
    import pr_consolidation as pc
    out = []
    for p in gr.get_receivable_pos(data_file=data_file):
        header = pc.get_po_header(p["po_number"], data_file)
        age = _age_days(header["po_date"]) if header else 0
        out.append({
            "id": p["po_number"], "category": "PO Awaiting Receipt",
            # Deliberately coarse in v1 — a real "is this actually late" signal
            # needs each line's own delivery_date compared to today, not just
            # PO age; noted here rather than faked as more precise than it is.
            "severity": "Attention" if age > 30 else "Info",
            "title": f"{p['po_number']} ({p['vendor_name']}) — {p['outstanding_lines']} of "
                     f"{p['lines']} line(s) still outstanding, {age}d since PO",
            "age_days": age, "owner_app": "MFG", "owner_screen": "Goods Receipt → Create GR",
            "source_module": "goods_receipt",
        })
    return out


def _credit_hold_source(data_file=None):
    import sales_order as so
    out = []
    for o in so.get_orders(status="Credit Hold", data_file=data_file):
        age = _age_days(o["order_date"])
        out.append({
            "id": o["so_id"], "category": "Credit Hold",
            "severity": "Critical",
            "title": f"{o['so_id']} ({o['customer_name']}) ₹{o['total_value']:,.0f} "
                     f"blocked on Credit Hold — {age}d",
            "age_days": age, "owner_app": "O2C", "owner_screen": "Sales Orders → Manage Orders",
            "source_module": "sales_order",
        })
    return out


def _overdue_ar_source(data_file=None):
    import cash_application as ca
    out = []
    for inv in ca.get_overdue_invoices(data_file=data_file):
        age = _age_days(inv["due_date"])
        out.append({
            "id": inv["invoice_id"], "category": "Overdue AR",
            "severity": "Critical" if age > 30 else "Attention",
            "title": f"{inv['invoice_id']} ({inv['customer_name']}) ₹{inv['balance_due']:,.0f} "
                     f"overdue {age}d ({inv['payment_status']})",
            "age_days": age, "owner_app": "O2C",
            "owner_screen": "Cash Application → Manage & Reports",
            "source_module": "cash_application",
        })
    return out


def _rma_disposition_source(data_file=None):
    import rma
    out = []
    for r in rma.get_pending_disposition(data_file=data_file):
        age = _age_days(r["received_date"])
        out.append({
            "id": r["rma_id"], "category": "RMA Awaiting Disposition",
            "severity": "Critical" if age > 7 else "Attention",
            "title": f"{r['received_qty']:g} of {r['mat_desc']} received against {r['rma_id']} — "
                     f"{age}d awaiting Sellable/Scrap decision",
            "age_days": age, "owner_app": "MFG",
            "owner_screen": "Quality Inspection → RMA Receipt & Disposition",
            "source_module": "rma",
        })
    return out


def _rfx_clarification_source(data_file=None):
    import rfx
    out = []
    for c in rfx.get_clarifications(data_file=data_file):
        if c["status"] != "Sent":
            continue
        age = _age_days(c["requested_date"])
        out.append({
            "id": c["clarification_id"], "category": "RFx Clarification Outstanding",
            "severity": "Attention" if age > 7 else "Info",
            "title": f"{c['vendor_name']} hasn't answered on {c['rfp_number']} — {age}d "
                     f"(“{c['question'][:60]}{'...' if len(c['question']) > 60 else ''}”)",
            "age_days": age, "owner_app": "S2C", "owner_screen": "RFx Management → Select Winners",
            "source_module": "rfx",
        })
    return out


_SOURCES = [
    _backorder_source, _at_risk_source, _quality_hold_source, _pending_rtv_source,
    _gr_pending_inspection_source, _po_awaiting_receipt_source, _credit_hold_source,
    _overdue_ar_source, _rfx_clarification_source, _rma_disposition_source,
]


def get_all_exceptions(data_file=None, category=None, severity=None, owner_app=None):
    """
    The one real entry point. Every adapter is called fresh on every
    call — no caching, no staleness, matching this whole codebase's own
    "recompute from the real ledger, never a stored snapshot" discipline.
    A single adapter raising doesn't take down the whole tower — its
    own exceptions are just missing for this call, surfaced as a real
    partial-failure note rather than a blank page or a hidden crash.
    """
    all_ex = []
    failed_sources = []
    for src in _SOURCES:
        try:
            all_ex.extend(src(data_file))
        except Exception as e:
            failed_sources.append((src.__name__, str(e)))

    if category:
        all_ex = [e for e in all_ex if e["category"] == category]
    if severity:
        all_ex = [e for e in all_ex if e["severity"] == severity]
    if owner_app:
        all_ex = [e for e in all_ex if e["owner_app"] == owner_app]

    all_ex.sort(key=lambda e: (SEVERITY_ORDER.get(e["severity"], 9), -e["age_days"]))
    return all_ex, failed_sources


def get_categories(data_file=None):
    """Every real category currently registered — for a UI filter list
    that never goes stale relative to _SOURCES above."""
    all_ex, _ = get_all_exceptions(data_file)
    seen = []
    for e in all_ex:
        if e["category"] not in seen:
            seen.append(e["category"])
    return seen


def stats(data_file=None):
    all_ex, failed = get_all_exceptions(data_file)
    by_severity = {}
    for e in all_ex:
        by_severity[e["severity"]] = by_severity.get(e["severity"], 0) + 1
    return {"total": len(all_ex), "by_severity": by_severity,
            "sources_registered": len(_SOURCES), "sources_failed": len(failed)}


if __name__ == "__main__":
    print("Control Tower stats:", stats())
