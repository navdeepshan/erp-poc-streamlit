"""
demo_scenario_ids.py — "One Container, Four Cities": an inventory-
optimization demo scenario for IDS Denmed, the trading/distribution
counterpart to demo_scenario.py's Genrobotics manufacturing story.

Built entirely through real app functions (not raw SQL), same
philosophy as demo_scenario.py. Requires Demand Detection Mode =
"All Items" (already set in SeedData_IDS.xlsx's Org_Defaults) — IDS
Denmed doesn't manufacture anything, so the BOM-only default would
show zero demand no matter what data exists here.

The story: a bulk import consignment lands entirely at Chennai (the
port). Four real customers in four different cities then place orders
— Bangalore, Mumbai, Hyderabad, and Delhi — while all the stock still
sits at the port. Executing the suggested transfers resolves three of
the four items completely; the fourth (sterilization pouches, a fast-
moving consumable) genuinely runs short even after using up everything
at Chennai — triggering a direct reorder to MELAG (not competitive RFx:
a distributor reordering an exclusive branded line doesn't shop it
around the way Genrobotics' generic servo motor legitimately could).
Delhi is IDS Denmed's own HQ state, so that one order lands same-state
(CGST+SGST) while the other three are inter-state (IGST) — deliberate,
to show both tax paths in one scenario.

Same phase-split structure as demo_scenario.py, applied from the start
this time (that was a retrofit there — lesson learned):
  run_setup()      — phases 0-2, the 'before' state, live transfer
                      opportunities left for a presenter to execute
  run_resolution()  — phases 4-8, assumes the transfers already
                      happened live
  run_all()         — everything, one shot, no live interaction

Known limitation, found while building this (not fixed here — see
CONTEXT_HANDOFF_v2.md): bom.get_transfer_opportunities() doesn't
decrement source-location stock across multiple destination
suggestions computed in the same pass, so two cities both drawing on
the same scarce material at the same source could show suggestions
that together exceed what's actually there. This scenario's quantities
are deliberately chosen to stay within that limitation (every material
two destinations compete for has enough combined supply at Chennai;
the one genuine shortage has only one destination drawing on it) —
not a workaround so much as staying inside a boundary that's worth
fixing properly later.
"""

from datetime import date

import db
import customer_onboarding as co
import pr_consolidation as pc
import goods_receipt as gr
import accounting as acct
import inventory as inv
import sales_order as so

SCALER = "IDS-USC-0001"
AIRPOLISHER = "IDS-USC-0002"
AUTOCLAVE = "IDS-STR-0001"
POUCHES = "IDS-STR-0003"
CHENNAI = "IDS_DL_CHN"
CHENNAI_GEO = "13.024762,80.163549"


def _backdate_gr(gr_id, new_date):
    """Same technique as demo_scenario.py's version — create_gr() always
    dates itself today, this patches the fields a person actually sees
    afterward so the timeline looks organic."""
    conn = db.get_connection()
    conn.execute("UPDATE gr_header SET gr_date=? WHERE gr_id=?", (new_date, gr_id))
    conn.execute("UPDATE journal_entries SET entry_date=? WHERE source_type='GR' AND source_id=?",
                (new_date, gr_id))
    conn.execute("UPDATE inventory_transactions SET txn_date=? WHERE reference_type='GR' AND reference_id=?",
                (new_date, gr_id))
    conn.commit()
    conn.close()


def phase0_setup():
    """Nothing to configure beyond what the seed file already sets
    (Demand Detection Mode = All Items, realistic credit limits already
    on the customer records) — kept as its own phase for symmetry with
    demo_scenario.py and as a natural place to add setup steps later."""
    return {"demand_mode": "All Items (already set in seed data)"}


def phase1_bulk_import():
    """
    The container lands at Chennai — but not out of nowhere. A real PR
    gets raised for the consolidated import (Operations planning ahead
    for the regional offices' known demand pattern), consolidated
    through the normal S2C screens into two POs — one per brand
    (Woodpecker, MELAG) — same as any other procurement in this app.
    Full paper trail: PR -> Consolidate -> PO -> GR, all traceable in
    the S2C app. (An earlier version of this function called
    pr_consolidation.insert_po() directly, skipping the PR entirely —
    fine for the GR/accounting/inventory mechanics, but left nothing to
    find when someone went looking for the PR that led to it. Fixed
    2026-07-28, per direct request, to have the real chain instead.)

    Each PR line also carries deliv_geo — Chennai's real coordinates —
    which turned out to matter for a second, separate reason: the S2C
    app's "All POs Map" (pr_consolidation.build_all_pos_geo_data())
    silently skips any route missing either the vendor's or the
    delivery location's geolocation. The original version left
    deliv_geo blank on every line, so even though the PO genuinely
    existed, it was invisible on that map — not a bug in the app, a
    bug in this script.
    """
    pr_id = "PR-20260610-101"
    pc.create_pr(pr_id, requester_id="REQ-OPS", requester_name="Chennai Import Desk",
                requester_dept="Operations", project_id="Q3-BULK-IMPORT",
                pr_date="2026-06-10",
                lines=[
                    {"vendor": "WOODPECKER", "mat_code": SCALER,
                     "mat_desc": "Woodpecker UDS-J Ultrasonic Scaler", "uom": "pcs", "qty": 10,
                     "req_date": "2026-06-15", "deliv_loc": CHENNAI, "deliv_geo": CHENNAI_GEO},
                    {"vendor": "WOODPECKER", "mat_code": AIRPOLISHER,
                     "mat_desc": "Woodpecker Air Polisher AP-P", "uom": "pcs", "qty": 6,
                     "req_date": "2026-06-15", "deliv_loc": CHENNAI, "deliv_geo": CHENNAI_GEO},
                    {"vendor": "MELAG", "mat_code": AUTOCLAVE,
                     "mat_desc": "MELAG Vacuclave 23B Autoclave", "uom": "pcs", "qty": 4,
                     "req_date": "2026-06-15", "deliv_loc": CHENNAI, "deliv_geo": CHENNAI_GEO},
                    {"vendor": "MELAG", "mat_code": POUCHES,
                     "mat_desc": "MELAG Sterilization Pouches Box of 200", "uom": "box", "qty": 30,
                     "req_date": "2026-06-15", "deliv_loc": CHENNAI, "deliv_geo": CHENNAI_GEO},
                ])
    result = pc.run()  # vendor-tagged lines route straight to PO — one per vendor, no RFP needed

    gr_ids = {}
    for summary in result["po_summary"]:
        po_number = summary["po_number"]
        pc.mark_po_created(po_number)
        raw_items = pc.get_po_items(po_number)
        line_receipts = {it["po_item"]: it["quantity"] for it in raw_items}
        gr_id = gr.create_gr(po_number, line_receipts, delivery_location=CHENNAI,
                             received_by="Chennai Import Desk",
                             notes=f"Bulk import consignment — {summary['vendor']}")
        acct.post_gr_entry(gr_id)
        _backdate_gr(gr_id, "2026-06-18")
        gr_ids[summary["vendor"]] = gr_id

    return {"pr_id": pr_id, "po_summary": result["po_summary"], "gr_ids": gr_ids}


def phase2_demand():
    """Four real customers, four different cities. Delhi is IDS
    Denmed's own HQ state (same-state, CGST+SGST once invoiced);
    Bangalore, Mumbai, and Hyderabad are all inter-state (IGST)."""
    orders = {}

    orders["so_blr"] = so.create_direct_order("SMILECARE-BLR",
        [{"mat_code": SCALER, "mat_desc": "Woodpecker UDS-J Ultrasonic Scaler", "uom": "pcs",
          "qty": 3, "unit_price": 9500},
         {"mat_code": AIRPOLISHER, "mat_desc": "Woodpecker Air Polisher AP-P", "uom": "pcs",
          "qty": 2, "unit_price": 22000}],
        delivery_location="IDS_DL_BLR", delivery_geo="12.948019,77.574706",
        requested_delivery_date="2026-08-05",
        notes="Clinic hygiene equipment upgrade — Bangalore")

    orders["so_del"] = so.create_direct_order("CAREPLUS-DEL",
        [{"mat_code": AUTOCLAVE, "mat_desc": "MELAG Vacuclave 23B Autoclave", "uom": "pcs",
          "qty": 1, "unit_price": 395000}],
        delivery_location="IDS_DL_DEL", delivery_geo="28.630487,77.127103",
        requested_delivery_date="2026-08-01",
        notes="Sterilization compliance upgrade — same-state order (Delhi)")

    orders["so_bom"] = so.create_direct_order("METRODENTAL-BOM",
        [{"mat_code": SCALER, "mat_desc": "Woodpecker UDS-J Ultrasonic Scaler", "uom": "pcs",
          "qty": 2, "unit_price": 9500},
         {"mat_code": AUTOCLAVE, "mat_desc": "MELAG Vacuclave 23B Autoclave", "uom": "pcs",
          "qty": 1, "unit_price": 395000}],
        delivery_location="IDS_DL_BOM", delivery_geo="19.135825,72.811831",
        requested_delivery_date="2026-08-08",
        notes="Lab equipment refresh — Mumbai")

    orders["so_hyd"] = so.create_direct_order("GOVTDENTAL-HYD",
        [{"mat_code": POUCHES, "mat_desc": "MELAG Sterilization Pouches Box of 200", "uom": "box",
          "qty": 35, "unit_price": 1450}],
        delivery_location="IDS_DL_HYD", delivery_geo="17.392754,78.469955",
        requested_delivery_date="2026-07-30",
        notes="Institutional bulk consumables order — Hyderabad")

    conn = db.get_connection()
    dates = {"so_blr": "2026-07-14", "so_del": "2026-07-15",
             "so_bom": "2026-07-16", "so_hyd": "2026-07-17"}
    for key, so_date in dates.items():
        conn.execute("UPDATE sales_orders SET order_date=? WHERE so_id=?",
                     (so_date, orders[key]["so_id"]))
    conn.commit()
    conn.close()

    return orders


def run_setup():
    """Phases 0-2 only — the 'before' state. Leaves real transfer
    opportunities (Chennai -> 4 cities) visible on the Inventory page
    for live execution, and one genuine shortfall (sterilization
    pouches at Hyderabad) that no transfer fully resolves."""
    print("Phase 0 — setup:", phase0_setup())
    print("Phase 1 — bulk import:", phase1_bulk_import())
    print("Phase 2 — demand:", phase2_demand())
    print("\nSetup done — transfer opportunities should now be visible on "
          "the Inventory page (Chennai fanning out to Bangalore, Delhi, "
          "Mumbai, Hyderabad). Execute them there, then call run_resolution().")


def phase6_reorder_pouches():
    """The honest remainder, addressed *after* fulfillment/billing below
    have already captured Hyderabad's genuine partial shipment (30 of
    35 — what was actually available at the time) — this closes the
    gap for the future, it doesn't retroactively complete an order
    that's already been recorded as partially shipped. That's the
    living-pipeline beat for this story, same spirit as Genrobotics'
    Mysore order staying at 1-of-3.

    Goes straight to a PO with the vendor already known — MELAG makes
    this exact branded consumable, there's no one else to competitively
    quote it against, unlike Genrobotics' generic servo motor. The PR
    line already carries the vendor tag, so pr_consolidation.run()
    routes it straight to a PO instead of RFP."""
    pr_id = "PR-20260722-102"
    pc.create_pr(pr_id, requester_id="REQ-OPS", requester_name="Chennai Import Desk",
                requester_dept="Operations", project_id="Q3-RESTOCK",
                pr_date="2026-07-22",
                lines=[{"vendor": "MELAG", "mat_code": POUCHES,
                        "mat_desc": "MELAG Sterilization Pouches Box of 200",
                        "uom": "box", "qty": 5, "req_date": "2026-07-22",
                        "deliv_loc": CHENNAI, "deliv_geo": CHENNAI_GEO}])
    result = pc.run()
    po_number = result["po_summary"][0]["po_number"]
    pc.mark_po_created(po_number)

    import po_export
    hdr = pc.get_po_header(po_number)
    raw_items = pc.get_po_items(po_number)
    po_lns = [{"mat_code": it["material_code"], "mat_desc": it["material_desc"], "uom": it["uom"],
              "qty": it["quantity"], "deliv_date": it["delivery_date"], "deliv_loc": it["delivery_location"]}
             for it in raw_items]
    po_export.make_av_bytes(po_number, "MELAG", hdr["po_type"], hdr["legal_entity"],
                 hdr["purchase_entity"], hdr["purchasing_group"], hdr["currency"],
                 hdr["plant_code"], po_lns)

    gr_id = gr.create_gr(po_number, {1: 5}, delivery_location=CHENNAI,
                         received_by="Chennai Import Desk", notes="Quick reorder — pouches shortfall")
    acct.post_gr_entry(gr_id)
    _backdate_gr(gr_id, "2026-07-24")

    import vendor_invoices as vi
    invoice_id = vi.create_invoice(gr_id, invoice_number="MELAG-INV-3391")
    payment = vi.record_payment(invoice_id, vi.get_invoice_payment_info(invoice_id)["amount"])

    # The 5 fresh boxes sit at Chennai, ready for Hyderabad's second
    # shipment whenever that gets scheduled — deliberately NOT
    # transferred or shipped here, so the order stays genuinely open.
    return {"pr_id": pr_id, "po_number": po_number, "gr_id": gr_id,
            "invoice_id": invoice_id, "payment": payment}


def phase5_fulfillment_and_billing():
    """Bangalore, Delhi, and Mumbai ship complete — everything they
    ordered was available after the transfers. Hyderabad ships 30 of
    35 boxes — a real partial shipment, what was actually on hand at
    the time, not staged for effect. Delhi is same-state (CGST+SGST);
    the other three are inter-state (IGST) — both tax paths shown in
    one scenario."""
    import fulfillment as ful
    import billing as bl

    results = {}

    fid_blr = ful.create_fulfillment("SO-00001")
    ful.record_shipment(fid_blr, {SCALER: 3, AIRPOLISHER: 2}, carrier="Safexpress",
                        tracking_ref="SFX-BLR-91021")
    ful.record_delivery(fid_blr, pod_reference="POD-BLR-001")
    acct.post_fulfillment_entry(fid_blr)
    inv_blr = bl.create_invoice(fid_blr, due_days=30, notes="Complete — Bangalore")
    bl.mark_issued(inv_blr["invoice_id"])
    acct.post_invoice_entry(inv_blr["invoice_id"])
    results["blr"] = {"fulfillment": fid_blr, "invoice": inv_blr}

    fid_del = ful.create_fulfillment("SO-00002")
    ful.record_shipment(fid_del, {AUTOCLAVE: 1}, carrier="VRL Logistics",
                        tracking_ref="VRL-DEL-44510")
    ful.record_delivery(fid_del, pod_reference="POD-DEL-001")
    acct.post_fulfillment_entry(fid_del)
    inv_del = bl.create_invoice(fid_del, due_days=30, notes="Complete — Delhi (same-state)")
    bl.mark_issued(inv_del["invoice_id"])
    acct.post_invoice_entry(inv_del["invoice_id"])
    results["del"] = {"fulfillment": fid_del, "invoice": inv_del}

    fid_bom = ful.create_fulfillment("SO-00003")
    ful.record_shipment(fid_bom, {SCALER: 2, AUTOCLAVE: 1}, carrier="Safexpress",
                        tracking_ref="SFX-BOM-77234")
    ful.record_delivery(fid_bom, pod_reference="POD-BOM-001")
    acct.post_fulfillment_entry(fid_bom)
    inv_bom = bl.create_invoice(fid_bom, due_days=45, notes="Complete — Mumbai")
    bl.mark_issued(inv_bom["invoice_id"])
    acct.post_invoice_entry(inv_bom["invoice_id"])
    results["bom"] = {"fulfillment": fid_bom, "invoice": inv_bom}

    fid_hyd = ful.create_fulfillment("SO-00004")
    ful.record_shipment(fid_hyd, {POUCHES: 30}, carrier="VRL Logistics",
                        tracking_ref="VRL-HYD-33087")
    ful.record_delivery(fid_hyd, pod_reference="POD-HYD-001")
    acct.post_fulfillment_entry(fid_hyd)
    inv_hyd = bl.create_invoice(fid_hyd, due_days=60,
                                notes="Partial — 30 of 35 boxes; balance pending next batch")
    bl.mark_issued(inv_hyd["invoice_id"])
    acct.post_invoice_entry(inv_hyd["invoice_id"])
    results["hyd"] = {"fulfillment": fid_hyd, "invoice": inv_hyd}

    conn = db.get_connection()
    for key, ship_date in [("blr", "2026-07-25"), ("del", "2026-07-25"),
                            ("bom", "2026-07-26"), ("hyd", "2026-07-27")]:
        fid = results[key]["fulfillment"]
        conn.execute("UPDATE fulfillments SET shipped_date=?, delivered_date=? WHERE fulfillment_id=?",
                     (ship_date, ship_date, fid))
    conn.commit()
    conn.close()

    return results


def phase7_cash_application():
    """Bangalore, Delhi, and Hyderabad pay in full. Mumbai pays 70% —
    a short payment with a reason on file, the same AR-aging beat as
    Genrobotics' Indore order, just distributed differently across
    four customers instead of two."""
    import cash_application as ca
    import billing as bl

    results = {}
    pay_dates = {"blr": date(2026, 7, 28), "del": date(2026, 7, 28),
                 "bom": date(2026, 7, 29), "hyd": date(2026, 7, 30)}

    for key, cust_id, ref_prefix in [("blr", "SMILECARE-BLR", "BLR-NEFT"),
                                      ("del", "CAREPLUS-DEL", "DEL-NEFT"),
                                      ("hyd", "GOVTDENTAL-HYD", "HYD-RTGS")]:
        inv_row = [i for i in bl.get_invoices() if i["customer_id"] == cust_id][-1]
        pay_id = ca.record_payment(cust_id, inv_row["grand_total"], payment_date=pay_dates[key],
                                   payment_method="NEFT", reference_no=f"{ref_prefix}-{cust_id[:4]}",
                                   notes="Full settlement")
        ca.apply_payment(pay_id, inv_row["invoice_id"], inv_row["grand_total"])
        results[key] = {"payment": pay_id, "invoice": inv_row["invoice_id"]}

    inv_bom = [i for i in bl.get_invoices() if i["customer_id"] == "METRODENTAL-BOM"][-1]
    partial_amt = round(inv_bom["grand_total"] * 0.7, 2)
    pay_bom = ca.record_payment("METRODENTAL-BOM", partial_amt, payment_date=pay_dates["bom"],
                                payment_method="RTGS", reference_no="BOM-RTGS-METR",
                                notes="Partial — balance pending internal approval")
    ca.apply_payment(pay_bom, inv_bom["invoice_id"], partial_amt,
                     short_payment_reason="Internal approval delay — balance to follow")
    results["bom"] = {"payment": pay_bom, "invoice": inv_bom["invoice_id"], "partial_amount": partial_amt}

    return results


def phase8_reorder_qty_planning_demo():
    """
    Configures Reorder Qty Based demand for a couple of (material,
    location) pairs — the mode this scenario's own business case (a
    multi-location distributor) most naturally reaches for once it's
    using this platform for S2S alone, with no O2C Sales Orders to
    drive the Sales Order Based mode at all. Real min/max + a typical
    reorder cadence, not tied to any Sales Order or PR — this
    configuration exists independently of whether O2C is even in use,
    and works identically whether it is or not. Unlike Genrobotics'
    phase9 (which has to run last, after every other phase's own PR
    consolidation sweeps), this only writes to material_location_
    planning_params — no PR/PO involved, so no ordering constraint;
    kept at the end purely for narrative consistency with that phase.
    """
    import bom
    bom.set_planning_params(POUCHES, "IDS_DL_BLR", min_qty=10, max_qty=40, reorder_cadence_days=21)
    bom.set_planning_params(SCALER, "IDS_DL_BOM", min_qty=3, max_qty=12, reorder_cadence_days=30)
    return {"pouches_blr": {"min": 10, "max": 40, "cadence": 21},
            "scaler_bom": {"min": 3, "max": 12, "cadence": 30}}


STARTER_KIT_ITEMS = [
    {"mat_code": "IDS-CHR-0001", "qty": 1},   # Woodpecker Dental Chair Unit W-1
    {"mat_code": "IDS-HPC-0001", "qty": 2},   # High Speed Dental Handpiece Standard (no vendor tag -> RFP)
    {"mat_code": "IDS-USC-0001", "qty": 1},   # Woodpecker UDS-J Ultrasonic Scaler
    {"mat_code": "IDS-EMT-0002", "qty": 1},   # Woodpecker Endo Motor DTC
    {"mat_code": "IDS-LCU-0001", "qty": 1},   # Woodpecker LED-D Curing Light
    {"mat_code": "IDS-RAD-0001", "qty": 1},   # Woodpecker Portable Dental X-Ray Unit
    {"mat_code": "IDS-STR-0001", "qty": 1},   # MELAG Vacuclave 23B Autoclave
    {"mat_code": "IDS-STR-0003", "qty": 5},   # MELAG Sterilization Pouches Box of 200
]


def phase9_new_customer_pipeline():
    """
    Three new clinic customers, three real quotation -> Sales Order ->
    bundle-based PR pipelines, staggered roughly a month apart across
    the next three months — the IDS Denmed counterpart to
    demo_scenario.py's phase13, built the same day for the same reason
    (a real, staged, live PR-consolidation demo scaling from small to
    large) but genuinely adapted, not copy-pasted, to fit a distributor
    rather than a manufacturer.

    The one design decision that actually matters here: unlike
    Genrobotics (where every customer's PR delivers to the Factory,
    since components get built into finished units there before
    shipping out), IDS Denmed doesn't manufacture anything — there's no
    "build" step, and the equivalent procurement PRs need to replenish
    stock at Chennai (the import hub), the exact same pattern phase1's
    own bulk-import PR already establishes. Routing each new customer's
    PR to their own city instead would have meant nothing would ever
    merge in consolidation — three different destinations never share
    a (material, location) pair no matter how much co-occurs. Sales
    Orders still deliver to each clinic's own city, same as the
    existing four customers; only the PROCUREMENT side routes to
    Chennai.

    New "Dental Clinic Starter Kit" bundle (STARTER_KIT_ITEMS above) —
    a distributor-appropriate bundle (discrete equipment a new clinic
    would order together), not a manufacturing BOM, since IDS Denmed
    has no BOM to reuse the way Genrobotics did. Built and verified
    directly before writing this: one line (IDS-HPC-0001) has no
    vendor tag, so it correctly routes to RFP with everything else
    correctly resolving to WOODPECKER or MELAG — checked this deliberately,
    not by accident, and found and fixed a real, general bug in
    _vendor_from_tags() along the way (see CONTEXT_HANDOFF_v2.md) that
    would otherwise have silently mis-tagged this and 8 other real
    items across this item master.

    Quantities vary slightly (1/2/1 kits) rather than being identical,
    same reasoning as phase13's 2/1/3 — realistic, and the merge still
    shows clearly since quantities don't need to match to combine.
    GSTINs are real, checksum-valid dummy values for each state
    (vendor_onboarding.demo_gstin()), not arbitrary-looking strings.

    Deliberately NOT called by run_setup()/run_resolution()/run_all()
    — this is its own, separately-triggered stage (the Settings page's
    "Add New Customer Wave" button), meant to run live after the
    audience has already seen a smaller example.
    """
    import purchase_bundles as pb
    import quotation as qt

    bundle_id = pb.create_bundle(
        "Dental Clinic Starter Kit",
        description="Core equipment for a new single-chair to multi-chair dental "
                    "clinic opening — chair, handpiece, scaler, endo motor, curing "
                    "light, X-ray, and sterilization gear.",
        department="Sales", created_by="", items=STARTER_KIT_ITEMS)
    # One lookup, not one per item per customer — get_bundle_items() is a
    # real DB call, and indexing into a fresh call of it inside a loop
    # both wastes calls and risks a real bug if row order ever isn't
    # guaranteed to match STARTER_KIT_ITEMS' own order.
    bundle_items_by_code = {i["mat_code"]: i for i in pb.get_bundle_items(bundle_id)}

    customers = [
        {"id": "SMILEZONE-AMD", "name": "SMILEZONE DENTAL CLINIC", "city": "Ahmedabad",
         "loc": "IDS_DL_AMD", "geo": "23.022500,72.571400",
         "address": "CG Road, Navrangpura, Ahmedabad, Gujarat 380009",
         "contact": "Dr. Kavita Shah", "email": "info@smilezonedental.in",
         "phone": "079-26301122", "gstin": "24AAALS9871K1ZD", "pan": "AAALS9871K",
         "credit_limit": 1200000, "kits": 1, "quote_date": "2026-07-25",
         "delivery_date": "2026-08-30", "pr_date": "2026-08-01",
         "pr_number": "PR-20260901-501", "project_id": "SMILEZONE-CLINIC"},
        {"id": "KOLKATADENTAL-CCU", "name": "KOLKATA DENTAL CARE CENTRE", "city": "Kolkata",
         "loc": "IDS_DL_CCU", "geo": "22.565860,88.356975",
         "address": "Park Street, Kolkata, West Bengal 700016",
         "contact": "Dr. Anirban Sen", "email": "contact@kolkatadentalcare.in",
         "phone": "033-22293344", "gstin": "19AAALK4523M1ZT", "pan": "AAALK4523M",
         "credit_limit": 2000000, "kits": 2, "quote_date": "2026-08-25",
         "delivery_date": "2026-09-30", "pr_date": "2026-09-01",
         "pr_number": "PR-20260901-502", "project_id": "KOLKATADENTAL-CLINIC"},
        {"id": "PINKCITY-JAI", "name": "PINK CITY DENTAL CLINIC", "city": "Jaipur",
         "loc": "IDS_DL_JAI", "geo": "26.819311,75.777112",
         "address": "MI Road, Jaipur, Rajasthan 302001",
         "contact": "Dr. Vikram Singh Rathore", "email": "care@pinkcitydental.in",
         "phone": "0141-2365577", "gstin": "08AAALJ6789N1Z7", "pan": "AAALJ6789N",
         "credit_limit": 1200000, "kits": 1, "quote_date": "2026-09-25",
         "delivery_date": "2026-10-30", "pr_date": "2026-10-01",
         "pr_number": "PR-20260901-503", "project_id": "PINKCITY-CLINIC"},
    ]

    results = []
    for c in customers:
        co.upsert_customer(c["id"], {
            "Customer_Name": c["name"], "Customer_Type": "Dental Clinic",
            "Geolocation": c["geo"], "City": c["city"], "Country": "India",
            "Address": c["address"], "Contact_Name": c["contact"],
            "Contact_Email": c["email"], "Contact_Phone": c["phone"],
            "GSTIN": c["gstin"], "PAN": c["pan"],
            "Credit_Limit": c["credit_limit"], "Payment_Terms": "Net 30",
        })
        co.approve_customer(c["id"])

        quote_lines = [{"mat_code": it["mat_code"],
                        "mat_desc": bundle_items_by_code[it["mat_code"]]["mat_desc"],
                        "uom": bundle_items_by_code[it["mat_code"]]["uom"],
                        "qty": it["qty"] * c["kits"],
                        "unit_price": _po_price(it["mat_code"])}
                       for it in STARTER_KIT_ITEMS]
        quote_id = qt.create_quote(c["id"], quote_lines,
            notes=f"New clinic equipment setup — {c['name'].title()}")
        qt.mark_sent(quote_id)
        qt.record_response(quote_id, accepted=True)
        conn = db.get_connection()
        conn.execute("UPDATE quotes SET quote_date=? WHERE quote_id=?", (c["quote_date"], quote_id))
        conn.commit()
        conn.close()

        order = so.create_order_from_quote(quote_id, delivery_location=c["loc"],
            delivery_geo=c["geo"], requested_delivery_date=c["delivery_date"],
            notes=f"{c['name'].title()} — new clinic setup")

        lines = pb.explode_bundle(bundle_id, multiplier=c["kits"])
        pr_lines = [{"vendor": l["vendor"], "mat_code": l["code"], "mat_desc": l["desc"],
                    "uom": l["uom"], "qty": l["qty"], "req_date": c["pr_date"],
                    "deliv_loc": CHENNAI, "deliv_geo": CHENNAI_GEO} for l in lines]
        pc.create_pr(c["pr_number"], requester_id="REQ-IMPORT", requester_name="Import Desk",
                    requester_dept="Operations", project_id=c["project_id"], lines=pr_lines,
                    pr_date=c["pr_date"])

        results.append({"customer": c["id"], "quote": quote_id, "so": order["so_id"],
                        "so_status": order["status"], "pr": c["pr_number"],
                        "kits": c["kits"], "delivery_date": c["delivery_date"]})
    return results


def _po_price(mat_code):
    import po_export
    item = po_export.get_item_by_code(mat_code)
    return item["price"] if item else 0


# Pilot-agnostic alias — see demo_scenario.py's own copy of this same
# alias for the full reasoning. This module's real phase number is 9,
# not 13; external callers use this name instead of either number.
new_customer_wave = phase9_new_customer_pipeline


def run_resolution():
    """Phases 5-7 — deliberately skips the transfer execution (phase 3
    equivalent, done live via the Inventory page's real Execute
    buttons — that's the whole point of splitting this out) but also
    the reorder is folded into fulfillment/billing's sequence here
    since Hyderabad's genuine partial shipment needs to be captured
    BEFORE the reorder closes the gap, not after."""
    print("Phase 5 — fulfillment & billing:", run_fulfillment_and_reorder())
    print("Phase 7 — cash application:", phase7_cash_application())
    print("Phase 8 — reorder-qty planning demo data:", phase8_reorder_qty_planning_demo())
    print("\nDone.")


def run_fulfillment_and_reorder():
    """Internal helper for run_resolution() — fulfillment/billing must
    run before the reorder, so Hyderabad's partial shipment reflects
    what was genuinely available at the time."""
    fulfillment_result = phase5_fulfillment_and_billing()
    reorder_result = phase6_reorder_pouches()
    return {"fulfillment": fulfillment_result, "reorder": reorder_result}


def run_all():
    """The whole story, start to finish, on top of whatever master data
    is currently seeded. Prints a short summary after each phase."""
    print("Phase 0 — setup:", phase0_setup())
    print("Phase 1 — bulk import:", phase1_bulk_import())
    print("Phase 2 — demand:", phase2_demand())
    import bom, inventory as inv
    opps = bom.get_transfer_opportunities()
    for o in opps:
        inv.execute_transfer(o["mat_code"], o["mat_desc"], o["from_location"],
                             o["to_location"], o["suggested_qty"])
    print(f"Phase 3 — executed {len(opps)} transfer(s)")
    print("Phase 5 — fulfillment & billing:", phase5_fulfillment_and_billing())
    print("Phase 6 — reorder:", phase6_reorder_pouches())
    print("Phase 7 — cash application:", phase7_cash_application())
    print("Phase 8 — reorder-qty planning demo data:", phase8_reorder_qty_planning_demo())
    print("\nDone.")


if __name__ == "__main__":
    run_all()
