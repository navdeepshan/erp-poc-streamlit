"""
vendor_scorecard.py — VSC-US-01 (rate a vendor) + VSC-US-02 (computed
vendor scorecard), S2S-E09.

VSC-US-01: a 1-5 buyer rating + review, tied to a real Closed PO — a
PO where every line is Fully Received, derived live from
goods_receipt.get_po_receipt_status() rather than a stored po_header
status this codebase doesn't have (no new status column, no second
place "closed" could drift from what receiving actually shows). One
rating per (PO, rater), immutable once submitted — a correction is a
new, explicit action this module doesn't provide, matching the same
"no silent edit" discipline as quality_inspection.py's own disposed
holds. A single vendor response per rating, also immutable once set.

VSC-US-02: combines four real components into one tenant-weighted
score (org_defaults' own "Vendor Scorecard Weight - *" keys, default
30/30/20/20) — no component is invented or estimated:
  - On-time delivery % — a Closed PO line's own real Delivery_Date vs.
    the real date its receipt actually completed (the latest GR
    touching that line, since a line isn't delivered until all of it
    arrived).
  - Quality reject rate — quality_inspection.py's own real Fail/
    Partial Pass line count over total inspected line count for POs
    from this vendor. Not qty-weighted, matching the story's own
    "Fail count over total inspected count" wording.
  - Price consistency — coefficient of variation across this vendor's
    own real po_items.unit_price history, per material, averaged
    across materials with more than one data point (a single
    data point has nothing to vary against, so it's excluded rather
    than scored as perfectly consistent by default).
  - Aggregated rating — VSC-US-01's own real average rating, scaled
    from its native 1-5 to 0-100 for the same weighted blend.

A vendor below the configured minimum transaction count is flagged
Low Volume rather than hidden — a thin score built on one or two data
points is real, informative context for a Buyer, not something to
suppress. RTV-US-01's own confirmed returns are surfaced alongside
the quality component as supporting context, per the story's own
explicit instruction not to double-count them into the reject rate
itself (not every Fail results in an RTV).

New table:
  vendor_ratings   Rating_ID | PO_Number | Vendor_ID | Rated_By |
                   Rating | Review | Rating_Date | Vendor_Response |
                   Vendor_Response_Date
"""

import statistics
from datetime import date

import db
import goods_receipt as gr
import pr_consolidation as pc
import quality_inspection as qi
import rtv
import vendor_onboarding as vo
import org_defaults as od


# ── Closed PO eligibility ────────────────────────────────────────────────────
def is_po_closed(po_number, data_file=None):
    """A PO is Closed when every real line is Fully Received — derived
    live from goods_receipt's own receipt status, never a second,
    separately-maintained 'closed' flag that could disagree with it."""
    lines = gr.get_po_receipt_status(po_number, data_file)
    return bool(lines) and all(l["receipt_status"] == "Fully Received" for l in lines)


def get_closed_pos(vendor_id=None, data_file=None):
    """Every real Closed PO, optionally scoped to one vendor — the
    real pool of Rating-eligible transactions."""
    out = []
    for h in pc.get_all_po_headers(data_file):
        if vendor_id and h["supplier_id"] != vendor_id:
            continue
        if h["status"] != "Created":
            continue  # a Proposed PO was never actually sent; nothing to close
        if is_po_closed(h["po_number"], data_file):
            out.append(h)
    return out


def get_unrated_closed_pos(vendor_id, rated_by, data_file=None):
    """Closed POs for this vendor this rater hasn't already rated —
    the real candidate list for the 'rate a vendor' UI."""
    closed = get_closed_pos(vendor_id, data_file)
    conn = db.get_connection()
    try:
        rated = {r["po_number"] for r in conn.execute(
            "SELECT po_number FROM vendor_ratings WHERE vendor_id=? AND rated_by=?",
            (vendor_id, rated_by)).fetchall()}
    finally:
        conn.close()
    return [h for h in closed if h["po_number"] not in rated]


# ── VSC-US-01: rate a vendor ─────────────────────────────────────────────────
def _next_rating_id(conn):
    rows = conn.execute("SELECT rating_id FROM vendor_ratings WHERE rating_id LIKE 'VR-%'").fetchall()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["rating_id"].split("-")[1]))
        except Exception: pass
    return f"VR-{mx+1:05d}"


def rate_vendor(po_number, vendor_id, rated_by, rating, review="", data_file=None):
    """
    One real rating per (PO, rater) — a second call for the same pair
    is refused, never silently overwritten; a genuine change of mind
    isn't built here, the same "no silent edit" stance disposed
    Quality Holds already take.
    """
    if rating not in (1, 2, 3, 4, 5):
        raise ValueError("Rating must be an integer from 1 to 5.")
    header = pc.get_po_header(po_number, data_file)
    if header is None:
        raise ValueError(f"{po_number} not found.")
    if header["supplier_id"] != vendor_id:
        raise ValueError(f"{po_number} wasn't issued to {vendor_id}.")
    if not is_po_closed(po_number, data_file):
        raise ValueError(f"{po_number} isn't Closed yet — every line must be Fully "
                          f"Received before it's eligible for a rating.")

    conn = db.get_connection()
    try:
        existing = conn.execute(
            "SELECT rating_id FROM vendor_ratings WHERE po_number=? AND rated_by=?",
            (po_number, rated_by)).fetchone()
        if existing:
            raise ValueError(f"{rated_by} already rated {po_number} ({existing['rating_id']}) "
                              f"— a rating can't be silently resubmitted.")
        rating_id = _next_rating_id(conn)
        conn.execute(
            "INSERT INTO vendor_ratings (rating_id, po_number, vendor_id, rated_by, rating, "
            "review, rating_date, vendor_response, vendor_response_date) "
            "VALUES (?,?,?,?,?,?,?,?,?)",
            (rating_id, po_number, vendor_id, rated_by, rating, review,
             date.today().strftime("%Y-%m-%d"), "", ""))
        conn.commit()
    finally:
        conn.close()
    return {"rating_id": rating_id}


def respond_to_rating(rating_id, response, data_file=None):
    """The vendor's own single response — immutable once set, same as
    the rating itself."""
    conn = db.get_connection()
    try:
        row = conn.execute("SELECT vendor_response FROM vendor_ratings WHERE rating_id=?",
                           (rating_id,)).fetchone()
        if row is None:
            raise ValueError(f"{rating_id} not found.")
        if row["vendor_response"]:
            raise ValueError(f"{rating_id} already has a vendor response on file — "
                              f"it can't be replaced.")
        conn.execute(
            "UPDATE vendor_ratings SET vendor_response=?, vendor_response_date=? WHERE rating_id=?",
            (response, date.today().strftime("%Y-%m-%d"), rating_id))
        conn.commit()
    finally:
        conn.close()
    return {"rating_id": rating_id}


def get_ratings(vendor_id=None, po_number=None, data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT * FROM vendor_ratings ORDER BY rating_id").fetchall()
    finally:
        conn.close()
    out = []
    for r in rows:
        row = dict(r)
        if vendor_id and row["vendor_id"] != vendor_id:
            continue
        if po_number and row["po_number"] != po_number:
            continue
        out.append(row)
    return out


# ── VSC-US-02: computed scorecard ────────────────────────────────────────────
def _line_completion_date(po_number, po_item, data_file=None):
    """The real date this specific PO line's receipt actually
    completed — the latest GR that posted against it, since a line
    received across more than one GR isn't genuinely delivered until
    the last of them arrives. None if never received at all."""
    grs = {g["gr_id"]: g for g in gr.get_grs(po_number=po_number, data_file=data_file)
          if g["status"] != "Cancelled"}
    latest = None
    for gr_id, g in grs.items():
        for item in gr.get_gr_items(gr_id, data_file):
            if item["po_item"] == po_item and (item["qty_received"] or 0) > 0:
                if latest is None or g["gr_date"] > latest:
                    latest = g["gr_date"]
    return latest


def compute_on_time_pct(vendor_id, data_file=None):
    """% of real Closed-PO lines whose actual completion date was on
    or before the PO line's own real Delivery_Date. Returns
    (pct_or_None, sample_size) — None when there's nothing to measure
    yet, never a fabricated 100%."""
    closed = get_closed_pos(vendor_id, data_file)
    on_time = 0
    total = 0
    for h in closed:
        for item in pc.get_po_items(h["po_number"], data_file):
            completed = _line_completion_date(h["po_number"], item["po_item"], data_file)
            if not completed or not item["delivery_date"]:
                continue
            total += 1
            if completed <= item["delivery_date"]:
                on_time += 1
    if total == 0:
        return None, 0
    return round(100 * on_time / total, 1), total


def compute_quality_reject_rate(vendor_id, data_file=None):
    """Real Fail/Partial Pass line count over total inspected line
    count, across every GR received from this vendor — count-based,
    matching QHD-US-01's own documented definition, not qty-weighted."""
    grs = [g for g in gr.get_grs(data_file=data_file) if g["vendor_id"] == vendor_id]
    rejected = 0
    total = 0
    for g in grs:
        for line in qi.get_gr_quality_status(g["gr_id"], data_file):
            if line["status"] == qi.STATUS_NOT_DONE:
                continue
            total += 1
            if line["status"] in (qi.STATUS_FAILED, qi.STATUS_PARTIAL):
                rejected += 1
    if total == 0:
        return None, 0
    return round(100 * rejected / total, 1), total


def compute_price_consistency(vendor_id, data_file=None):
    """
    Coefficient of variation (stdev / mean) across this vendor's own
    real po_items.unit_price history, per material, averaged across
    every material with more than one real data point — a single
    quote has nothing to vary against, so it's excluded rather than
    scored as perfectly consistent by default (that would flatter a
    vendor rated on one order the same as one proven consistent across
    ten). Returns (score_0_to_100_or_None, sample_size) where the
    score is 100 at zero variation, falling to 0 by 100% variation
    and clamped there beyond it — a simple, named mapping, not a
    claim of statistical rigor.
    """
    by_material = {}
    for h in pc.get_all_po_headers(data_file):
        if h["supplier_id"] != vendor_id:
            continue
        for item in pc.get_po_items(h["po_number"], data_file):
            if item["unit_price"] is None:
                continue
            by_material.setdefault(item["material_code"], []).append(item["unit_price"])

    cvs = []
    for prices in by_material.values():
        if len(prices) < 2:
            continue
        mean = statistics.mean(prices)
        if mean <= 0:
            continue
        cv = statistics.stdev(prices) / mean
        cvs.append(cv)
    if not cvs:
        return None, 0
    avg_cv = statistics.mean(cvs)
    score = max(0.0, round(100 - avg_cv * 100, 1))
    return score, len(cvs)


def compute_aggregated_rating(vendor_id, data_file=None):
    ratings = get_ratings(vendor_id=vendor_id, data_file=data_file)
    if not ratings:
        return None, 0
    avg = statistics.mean(r["rating"] for r in ratings)
    return round(avg, 2), len(ratings)


def get_vendor_scorecard(vendor_id, data_file=None):
    """
    The real combined score. Every component that has no real data
    yet is None, not a fabricated neutral value, and is excluded from
    the weighted blend entirely (its weight redistributes across
    whatever components ARE real, proportionally) rather than silently
    counting an unmeasured component as zero or average.
    """
    on_time_pct, on_time_n = compute_on_time_pct(vendor_id, data_file)
    reject_pct, reject_n = compute_quality_reject_rate(vendor_id, data_file)
    price_score, price_n = compute_price_consistency(vendor_id, data_file)
    rating_avg, rating_n = compute_aggregated_rating(vendor_id, data_file)

    components = {
        "on_time": (on_time_pct, float(od.get_default("Vendor Scorecard Weight - On-Time", data_file))),
        "quality": (100 - reject_pct if reject_pct is not None else None,
                   float(od.get_default("Vendor Scorecard Weight - Quality", data_file))),
        "price_consistency": (price_score, float(od.get_default("Vendor Scorecard Weight - Price Consistency", data_file))),
        "rating": (rating_avg * 20 if rating_avg is not None else None,
                  float(od.get_default("Vendor Scorecard Weight - Rating", data_file))),
    }
    real = {k: v for k, (v, w) in components.items() if v is not None}
    weights = {k: w for k, (v, w) in components.items() if v is not None}
    total_weight = sum(weights.values())
    overall = (round(sum(real[k] * weights[k] for k in real) / total_weight, 1)
              if total_weight > 0 else None)

    min_txn = int(od.get_default("Vendor Scorecard Min Transactions", data_file))
    closed_count = len(get_closed_pos(vendor_id, data_file))
    rtv_count = len([r for r in rtv.get_rtv_shipments(data_file) if r["vendor_id"] == vendor_id])

    return {
        "vendor_id": vendor_id,
        "overall_score": overall,
        "low_volume": closed_count < min_txn,
        "closed_po_count": closed_count,
        "on_time_pct": on_time_pct, "on_time_sample": on_time_n,
        "quality_reject_pct": reject_pct, "quality_sample": reject_n,
        "price_consistency_score": price_score, "price_sample": price_n,
        "aggregated_rating": rating_avg, "rating_sample": rating_n,
        "rtv_count": rtv_count,
    }


def get_all_scorecards(data_file=None):
    return [get_vendor_scorecard(v["Vendor_ID"], data_file)
            for v in vo.list_vendors(data_file=data_file, include_inactive=False)]


def stats(data_file=None):
    ratings = get_ratings(data_file=data_file)
    return {"total_ratings": len(ratings),
            "responded": len([r for r in ratings if r["vendor_response"]])}


if __name__ == "__main__":
    print("Vendor Scorecard stats:", stats())
