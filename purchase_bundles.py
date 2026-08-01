"""
purchase_bundles.py — Purchase Bundles (S2C, sits before PR creation).

A Purchase Bundle is a named, reusable set of materials + default
quantities that a requester picks as a group instead of hunting down
each item individually — e.g. "New Hire Desk Setup," "Site HSE Kit,"
"Monthly Housekeeping Consumables." Conceptually a sibling of BOM_Items
(bom.py): same "parent -> line items with quantities" shape — but
deliberately simpler and scoped differently:

  - FLAT, not hierarchical. A bundle contains materials directly; it
    cannot contain another bundle. No recursion, no cycle detection —
    there's nothing to recurse into.
  - Procurement-only. A bundle exists purely to speed up requisitioning
    (Create PR); it has no role in production, has no relationship to
    BOM_Items, and bom.py's explode_bom()/production.py's consumption
    logic never touch it. Picking a bundle just pre-populates several
    PR lines at once with sensible default quantities — the requester
    can still edit quantities, drop lines, or add more before saving,
    exactly like the individual item picker already allows.
  - "Exploded" only at pick time. Unlike BOM_Items (which has a live
    multi-level explosion computed at PR-creation time via
    bom.propose_pr_lines()), a bundle's explosion is a single flat
    join: bundle_id -> its line items, each already carrying its own
    quantity — there's no aggregation-across-paths to compute, so
    explode_bundle() is a straight read, not a graph walk.

Once a bundle's items are picked into a PR, they are ordinary PR_Items
rows — pr_consolidation.run() and everything downstream (RFx, PO,
GR, Contracts) has no idea a bundle was ever involved. That's the
entire point: this module's job ends the moment the PR is created.

Line items carry material_desc/uom captured at add-time (same
replication convention BOM_Items already uses), NOT a live join against
Item Master — so a bundle's contents stay stable even if Item Master's
description changes later. Vendor tag and current price ARE looked up
fresh from Item Master at explode time (via po_export.get_item_by_code),
same as BOM's explode_bom_detailed() does, since those genuinely should
reflect current data, not what was true when the bundle was created.
"""

import os
from datetime import date

import db
import po_export

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")


def _vendor_from_tags(tags):
    """Same rule Create PR's item picker and bom.py use — an ALL-CAPS
    first word of Tags/Keywords, not just any first word. See bom.py's
    own copy of this function for the real bug this fixes."""
    t = (tags or "").strip()
    first = t.split()[0] if t else ""
    return first if first.isupper() else ""


# ── Bundle header CRUD ────────────────────────────────────────────────────────
def next_bundle_id(data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute(
            "SELECT bundle_id FROM purchase_bundles WHERE bundle_id LIKE 'BDL-%'"
        ).fetchall()
    finally:
        conn.close()
    mx = 0
    for r in rows:
        try: mx = max(mx, int(r["bundle_id"].split("-")[1]))
        except (ValueError, IndexError): pass
    return f"BDL-{mx+1:05d}"


def create_bundle(bundle_name, description="", department="", created_by="",
                   items=None, data_file=None):
    """
    items: list of dicts {mat_code, qty, notes(optional)}. material_desc
    and uom are looked up from Item Master at creation time and stored
    with the line (see module docstring on why). Unknown material codes
    raise ValueError up front — better to fail loudly at creation than
    silently produce a bundle with a dead reference.
    """
    if not bundle_name or not bundle_name.strip():
        raise ValueError("Bundle name is required.")
    items = items or []
    if not items:
        raise ValueError("A bundle needs at least one line item.")

    enriched = []
    for it in items:
        code = it["mat_code"]
        qty = it.get("qty")
        if qty is None or qty <= 0:
            raise ValueError(f"{code}: quantity must be greater than zero.")
        master = po_export.get_item_by_code(code, active_only=False)
        if not master:
            raise ValueError(f"{code}: not found in Item Master — check the code.")
        enriched.append({"mat_code": code, "mat_desc": master["desc"], "uom": master["uom"],
                          "qty": qty, "notes": it.get("notes", "")})

    db.init_schema()
    conn = db.get_connection()
    try:
        bundle_id = next_bundle_id(data_file)
        conn.execute(
            "INSERT INTO purchase_bundles (bundle_id, bundle_name, description, department, "
            "created_by, created_date, active) VALUES (?,?,?,?,?,?,?)",
            (bundle_id, bundle_name.strip(), description, department, created_by,
             date.today().strftime("%Y-%m-%d"), "Yes"),
        )
        for it in enriched:
            conn.execute(
                "INSERT INTO purchase_bundle_items (bundle_id, material_code, material_desc, "
                "uom, default_qty, notes) VALUES (?,?,?,?,?,?)",
                (bundle_id, it["mat_code"], it["mat_desc"], it["uom"], it["qty"], it["notes"]),
            )
        conn.commit()
    finally:
        conn.close()
    return bundle_id


def list_bundles(active_only=True, department=None, data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute(
            "SELECT * FROM purchase_bundles ORDER BY bundle_name"
        ).fetchall()
    finally:
        conn.close()
    out = []
    for r in rows:
        if active_only and (r["active"] or "Yes") != "Yes":
            continue
        if department and r["department"] != department:
            continue
        out.append({"bundle_id": r["bundle_id"], "bundle_name": r["bundle_name"],
                    "description": r["description"] or "", "department": r["department"] or "",
                    "created_by": r["created_by"] or "", "created_date": r["created_date"],
                    "active": r["active"] or "Yes"})
    return out


def get_bundle(bundle_id, data_file=None):
    conn = db.get_connection()
    try:
        r = conn.execute(
            "SELECT * FROM purchase_bundles WHERE bundle_id = ?", (bundle_id,)
        ).fetchone()
    finally:
        conn.close()
    if not r:
        return None
    return {"bundle_id": r["bundle_id"], "bundle_name": r["bundle_name"],
            "description": r["description"] or "", "department": r["department"] or "",
            "created_by": r["created_by"] or "", "created_date": r["created_date"],
            "active": r["active"] or "Yes"}


def get_bundle_items(bundle_id, data_file=None):
    conn = db.get_connection()
    try:
        rows = conn.execute(
            "SELECT * FROM purchase_bundle_items WHERE bundle_id = ? ORDER BY id",
            (bundle_id,),
        ).fetchall()
    finally:
        conn.close()
    return [{"mat_code": r["material_code"], "mat_desc": r["material_desc"],
             "uom": r["uom"], "default_qty": r["default_qty"], "notes": r["notes"] or ""}
            for r in rows]


def set_bundle_active(bundle_id, active, data_file=None):
    """active: True/False. Deactivating just removes the bundle from the
    Create PR picker (list_bundles(active_only=True)) — it's not
    deleted, and any PR already created from it is completely
    unaffected (its items are already ordinary PR_Items rows)."""
    conn = db.get_connection()
    try:
        cur = conn.execute(
            "UPDATE purchase_bundles SET active = ? WHERE bundle_id = ?",
            ("Yes" if active else "No", bundle_id),
        )
        found = cur.rowcount > 0
        conn.commit()
    finally:
        conn.close()
    return found


# ── Bundle item CRUD (edit an existing bundle) ───────────────────────────────
def add_or_update_bundle_item(bundle_id, mat_code, qty, notes="", data_file=None):
    """Upsert: same bundle+material pair updates qty/notes instead of
    duplicating (same UNIQUE-constraint pattern as bom.add_bom_line)."""
    if qty is None or qty <= 0:
        raise ValueError("Quantity must be greater than zero.")
    master = po_export.get_item_by_code(mat_code, active_only=False)
    if not master:
        raise ValueError(f"{mat_code}: not found in Item Master — check the code.")
    if get_bundle(bundle_id) is None:
        raise ValueError(f"Bundle {bundle_id} not found.")

    db.init_schema()
    conn = db.get_connection()
    try:
        conn.execute(
            "INSERT INTO purchase_bundle_items (bundle_id, material_code, material_desc, "
            "uom, default_qty, notes) VALUES (?,?,?,?,?,?) "
            "ON CONFLICT(bundle_id, material_code) DO UPDATE SET "
            "default_qty=excluded.default_qty, notes=excluded.notes",
            (bundle_id, mat_code, master["desc"], master["uom"], qty, notes),
        )
        conn.commit()
    finally:
        conn.close()


def remove_bundle_item(bundle_id, mat_code, data_file=None):
    conn = db.get_connection()
    try:
        cur = conn.execute(
            "DELETE FROM purchase_bundle_items WHERE bundle_id = ? AND material_code = ?",
            (bundle_id, mat_code),
        )
        removed = cur.rowcount > 0
        conn.commit()
    finally:
        conn.close()
    return removed


# ── Explosion (pick time) ────────────────────────────────────────────────────
def explode_bundle(bundle_id, multiplier=1, data_file=None):
    """
    Returns the bundle's line items ready to stage into a PR — same
    dict shape Create PR's individual item picker already produces
    (code/desc/vendor/uom/price/qty), so the picker UI can treat a
    bundle pick and an individual item pick identically once exploded.
    Vendor tag and unit price come from a FRESH Item Master lookup
    (current data), while desc/uom come from what was stored on the
    bundle line (stable snapshot) — see module docstring.

    multiplier: scales every line's default_qty (e.g. picking a
    "Site HSE Kit" bundle for 3 sites at once). Defaults to 1 (use the
    bundle's stored default quantities as-is).
    """
    bundle = get_bundle(bundle_id, data_file)
    if bundle is None:
        raise ValueError(f"Bundle {bundle_id} not found.")
    lines = get_bundle_items(bundle_id, data_file)
    out = []
    for ln in lines:
        master = po_export.get_item_by_code(ln["mat_code"], active_only=False)
        vendor = _vendor_from_tags(master["tags"]) if master else ""
        price = master["price"] if master else None
        out.append({"code": ln["mat_code"], "desc": ln["mat_desc"], "vendor": vendor,
                    "uom": ln["uom"], "price": price, "qty": ln["default_qty"] * multiplier,
                    "notes": ln["notes"]})
    return out


def discover_bundle_candidates(min_po_count=2, data_file=None):
    """
    PR-US-04's agent-assisted discovery step, made real: analyzes real
    PO line-item co-occurrence — materials that repeatedly appear
    together on the same Purchase Order — and proposes candidate
    bundles a Buyer or Category Manager hasn't yet defined. Never
    creates anything on its own; every candidate still requires an
    explicit Buyer action through create_bundle()/add_or_update_
    bundle_item() (steps 1-3 of this story's own Business Flow) before
    it becomes a real, selectable bundle.

    Algorithm, deliberately simple and explainable rather than a black
    box: for every PO with 2+ distinct materials, every pairwise
    combination of materials on it is one co-occurrence observation.
    A pair observed together on at least min_po_count distinct POs is
    a real, repeated pattern, not a coincidence — proposed as a
    candidate with the specific PO numbers behind it as evidence. A
    pair already fully covered by an existing bundle (both materials
    already members of the same bundle) is excluded — no point
    proposing what's already been captured.

    Returns a list of {mat_codes: [a, b], mat_descs: [...], po_count,
    po_refs: [...]}, sorted by po_count descending (the most-repeated,
    most-confident pattern first).
    """
    import itertools
    conn = db.get_connection()
    try:
        rows = conn.execute(
            "SELECT po_number, material_code, material_desc FROM po_items "
            "WHERE material_code IS NOT NULL AND material_code != ''"
        ).fetchall()
    finally:
        conn.close()

    by_po = {}
    desc_by_code = {}
    for r in rows:
        by_po.setdefault(r["po_number"], set()).add(r["material_code"])
        desc_by_code[r["material_code"]] = r["material_desc"]

    pair_pos = {}  # frozenset({matA, matB}) -> set of po_numbers
    for po_number, materials in by_po.items():
        if len(materials) < 2:
            continue
        for a, b in itertools.combinations(sorted(materials), 2):
            key = frozenset((a, b))
            pair_pos.setdefault(key, set()).add(po_number)

    # Materials already bundled together — excluded from candidacy
    already_bundled = set()
    for b in list_bundles(active_only=False, data_file=data_file):
        items = get_bundle_items(b["bundle_id"], data_file=data_file)
        codes = {i["mat_code"] for i in items}
        for a, c in itertools.combinations(sorted(codes), 2):
            already_bundled.add(frozenset((a, c)))

    candidates = []
    for pair, pos in pair_pos.items():
        if len(pos) < min_po_count or pair in already_bundled:
            continue
        mat_codes = sorted(pair)
        candidates.append({
            "mat_codes": mat_codes,
            "mat_descs": [desc_by_code.get(c, c) for c in mat_codes],
            "po_count": len(pos),
            "po_refs": sorted(pos),
        })
    return sorted(candidates, key=lambda c: -c["po_count"])


def stats(data_file=None):
    bundles = list_bundles(active_only=False)
    conn = db.get_connection()
    try:
        total_lines = conn.execute("SELECT COUNT(*) FROM purchase_bundle_items").fetchone()[0]
    finally:
        conn.close()
    return {"total_bundles": len(bundles),
            "active_bundles": sum(1 for b in bundles if b["active"] == "Yes"),
            "total_lines": total_lines}


if __name__ == "__main__":
    print("Purchase Bundles stats:", stats())
