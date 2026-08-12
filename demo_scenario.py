"""
demo_scenario.py — "Two City Contracts, One Smart Transfer": an
inventory-optimization demo scenario for Genrobotics, built entirely
through the same functions the UI calls (not raw SQL), so every number
that shows up anywhere in the app is exactly as consistent as if a
person had clicked through it by hand.

Deliberately a separate, optional layer on top of seed_manager.py's
master-data reset, not folded into the seed template itself — a reset
gives a clean slate; this then tells one specific story on top of it.
Designed to be re-run after any reset.

The story: R&D (Trivandrum) has surplus stock of two components left
over from a wrapped-up pilot batch. Factory (Palakkad) is about to get
squeezed by two real customer orders and doesn't have enough of either.
Executing the suggested transfer resolves the battery shortage
completely and most of the servo motor shortage — leaving a real,
honest remainder that goes through full competitive procurement
(RFx -> contract award -> PO -> GR -> Vendor Invoice -> Payment). The
rest of the BOM is sourced via the existing Purchase Bundle. Two of
the four ordered units get carried all the way to cash — the other two
are left visibly in progress, not a fully "finished" story.

Run phases individually (each is independent enough to inspect before
moving on) or call run_all() for the whole thing in order.
"""

from datetime import date

import db
import customer_onboarding as co
import pr_consolidation as pc
import goods_receipt as gr
import accounting as acct
import inventory as inv
import sales_order as so

BATTERY = "GRB-PWR-0001"
SERVO = "GRB-MEC-0001"
FACTORY = "GRB_DL_PKD_Factory"
RND = "GRB_DL_TVM_RnD"
BANDICOOT = "GRB-FG-0001"
UNIT_PRICE = 1200000  # a believable market price for a fielded municipal robotic unit


def _geo(location_id):
    """Live lookup, not a hardcoded copy of Delivery_Locations' own
    real coordinates — found and fixed 2026-08-11 after exactly this
    class of bug: FACTORY's own literal was correct everywhere it was
    copied, but RND's ("8.5241,76.9366") had quietly drifted from the
    real seeded value (8.5578,76.8807), and phase1_historical_stock()'s
    four lines had never carried any deliv_geo at all — silently
    missing every route on the S2C app's 'All POs Map'
    (pr_consolidation.build_all_pos_geo_data() skips a route missing
    either endpoint's geolocation) despite the PO genuinely existing,
    same bug class demo_scenario_ids.py's own phase1_bulk_import()
    docstring already names and fixed once for Chennai. Every call site
    below reads live from Delivery_Locations now, so a location's own
    geo can never again drift out of sync with a copy pasted into this
    file."""
    import pr_consolidation as _pc
    for loc in _pc.get_delivery_locations(active_only=False):
        if loc["id"] == location_id:
            return loc.get("geo", "")
    return ""


def _backdate_gr(gr_id, new_date):
    """create_gr() always dates itself today — this patches the date
    fields a person actually sees (GR date, its journal entry, its
    inventory transactions) after the fact, so the demo's timeline
    looks organic instead of everything landing on today's date. PR/PO
    numbers keep their real creation-date pattern regardless — those
    embed the date in the ID itself, not worth reconstructing for a
    detail nobody scrutinizes mid-demo."""
    conn = db.get_connection()
    conn.execute("UPDATE gr_header SET gr_date=? WHERE gr_id=?", (new_date, gr_id))
    conn.execute("UPDATE journal_entries SET entry_date=? WHERE source_type='GR' AND source_id=?",
                (new_date, gr_id))
    conn.execute("UPDATE inventory_transactions SET txn_date=? WHERE reference_type='GR' AND reference_id=?",
                (new_date, gr_id))
    conn.commit()
    conn.close()


def phase0_setup():
    """Realistic credit limits for actual municipal capital-equipment
    contracts — the seed file's defaults (500K / 1.2M) were sized for
    an empty demo, not a real multi-unit Bandicoot order; left as-is,
    Mysore's order below would trip a Credit Hold and derail the story
    before it starts."""
    co.set_credit_limit("MCC", 4000000)
    co.set_credit_limit("IMC", 1500000)
    return {"mcc_limit": 4000000, "imc_limit": 1500000}


def phase1_historical_stock():
    """The imbalance this whole story hinges on: R&D received a pilot
    batch six weeks ago and never drew it down; Factory's had normal
    routine restocking since, at a smaller scale. Both are real GRs,
    posted through accounting normally, just backdated afterward."""
    pc.insert_po("PO-DEMO-RND1",
        {"po_type": "NB", "legal_entity": "LE-001", "purch_entity": "PE-001", "purch_group": "PG-001",
         "currency": "INR", "plant_code": "PLANT-01", "supplier_id": "COSTAPWR",
         "supplier_name": "Costa Power Industries Pvt Ltd", "supplier_geo": ""},
        [{"mat_code": BATTERY, "mat_desc": "Lithium-Ion Battery Pack 24V 20Ah", "uom": "pcs",
          "qty": 6, "unit_price": 24000, "deliv_date": "2026-06-15", "deliv_loc": RND, "deliv_geo": _geo(RND),
          "source_pr": "", "source_pr_line": "", "req_id": "", "req_dept": "", "project_id": "PILOT-BATCH-1"},
         {"mat_code": SERVO, "mat_desc": "Servo Motor 12V High Torque Arm Joint Drive", "uom": "pcs",
          "qty": 6, "unit_price": 8500, "deliv_date": "2026-06-15", "deliv_loc": RND, "deliv_geo": _geo(RND),
          "source_pr": "", "source_pr_line": "", "req_id": "", "req_dept": "", "project_id": "PILOT-BATCH-1"}])
    pc.mark_po_created("PO-DEMO-RND1")
    gr_rnd = gr.create_gr("PO-DEMO-RND1", {1: 6, 2: 6}, delivery_location=RND,
                          received_by="R&D Team", notes="Pilot batch build stock",
                          line_lots={1: {"lot_number": "LOT-BATT-2026-001", "expiry_date": "2028-06-15"},
                                     2: {"serials": [f"SN-SERVO-{i:04d}" for i in range(1, 7)]}})
    acct.post_gr_entry(gr_rnd)
    _backdate_gr(gr_rnd, "2026-06-15")

    pc.insert_po("PO-DEMO-FAC1",
        {"po_type": "NB", "legal_entity": "LE-001", "purch_entity": "PE-001", "purch_group": "PG-001",
         "currency": "INR", "plant_code": "PLANT-01", "supplier_id": "COSTAPWR",
         "supplier_name": "Costa Power Industries Pvt Ltd", "supplier_geo": ""},
        [{"mat_code": BATTERY, "mat_desc": "Lithium-Ion Battery Pack 24V 20Ah", "uom": "pcs",
          "qty": 2, "unit_price": 24000, "deliv_date": "2026-07-06", "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY),
          "source_pr": "", "source_pr_line": "", "req_id": "", "req_dept": "", "project_id": ""},
         {"mat_code": SERVO, "mat_desc": "Servo Motor 12V High Torque Arm Joint Drive", "uom": "pcs",
          "qty": 2, "unit_price": 8500, "deliv_date": "2026-07-06", "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY),
          "source_pr": "", "source_pr_line": "", "req_id": "", "req_dept": "", "project_id": ""}])
    pc.mark_po_created("PO-DEMO-FAC1")
    gr_fac = gr.create_gr("PO-DEMO-FAC1", {1: 2, 2: 2}, delivery_location=FACTORY,
                          received_by="Factory Stores", notes="Routine restock",
                          line_lots={1: {"lot_number": "LOT-BATT-2026-002", "expiry_date": "2028-07-06"},
                                     2: {"serials": ["SN-SERVO-0007", "SN-SERVO-0008"]}})
    acct.post_gr_entry(gr_fac)
    _backdate_gr(gr_fac, "2026-07-06")

    return {"gr_rnd": gr_rnd, "gr_factory": gr_fac}


def phase2_demand():
    """Two real customers, two real inter-state orders — both correctly
    land on IGST once invoiced later, not CGST/SGST, since neither
    Karnataka nor Madhya Pradesh is Kerala. delivery_location is
    Factory on both — that's genrobotics' own fulfilling location
    (what get_inventory_position() aggregates demand against), not the
    customer's city; the customer's city lives on the customer record
    itself, not the order."""
    so_mcc = so.create_direct_order("MCC",
        [{"mat_code": BANDICOOT, "mat_desc": "Bandicoot Robotic Scavenger - Standard Unit",
          "uom": "pcs", "qty": 3, "unit_price": UNIT_PRICE}],
        delivery_location=FACTORY, delivery_geo=_geo(FACTORY),
        requested_delivery_date="2026-08-15",
        notes="Multi-ward sewer-cleaning fleet rollout — Mysore City Corporation")
    so_imc = so.create_direct_order("IMC",
        [{"mat_code": BANDICOOT, "mat_desc": "Bandicoot Robotic Scavenger - Standard Unit",
          "uom": "pcs", "qty": 1, "unit_price": UNIT_PRICE}],
        delivery_location=FACTORY, delivery_geo=_geo(FACTORY),
        requested_delivery_date="2026-08-10",
        notes="URGENT — Indore Municipal Corporation priority unit")

    conn = db.get_connection()
    conn.execute("UPDATE sales_orders SET order_date=? WHERE so_id=?", ("2026-07-17", so_mcc["so_id"]))
    conn.execute("UPDATE sales_orders SET order_date=? WHERE so_id=?", ("2026-07-19", so_imc["so_id"]))
    conn.commit()
    conn.close()
    return {"so_mcc": so_mcc, "so_imc": so_imc}


def phase3_execute_transfer():
    """The punchline moment: both suggested transfers, executed. Battery
    fully resolves Factory's shortage (2 on hand + 6 transferred = 8,
    exactly the 8 needed). Servo motor only gets Factory to 8 of the
    12 needed — a real, remaining gap phase4 sources properly."""
    battery = inv.execute_transfer(BATTERY, "Lithium-Ion Battery Pack 24V 20Ah", RND, FACTORY, 6,
                                   notes="Suggested transfer — R&D pilot-batch surplus covers Factory shortage")
    servo = inv.execute_transfer(SERVO, "Servo Motor 12V High Torque Arm Joint Drive", RND, FACTORY, 6,
                                 notes="Suggested transfer — partial coverage, remainder to be procured")
    return {"battery_transfer": battery, "servo_transfer": servo}


def phase4_procure_remaining_servo_motors():
    """The honest remainder after the transfer: 4 more Servo Motor 12V,
    genuinely not available anywhere internally. Goes through full
    competitive sourcing — three electronics vendors quote, the best
    one wins — rather than just reordering from whoever supplied it
    last time."""
    import rfx

    pr_id = "PR-20260720-101"
    pc.create_pr(pr_id, requester_id="REQ-PROD", requester_name="Production Planning",
                requester_dept="Manufacturing", project_id="MCC-IMC-FLEET",
                pr_date="2026-07-20",
                lines=[{"vendor": "", "mat_code": SERVO,
                        "mat_desc": "Servo Motor 12V High Torque Arm Joint Drive",
                        "uom": "pcs", "qty": 4, "req_date": "2026-07-28",
                        "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY)}])
    result = pc.run()

    rfp_number = None
    for rfp in rfx.get_open_rfps():
        if rfp["mat_code"] == SERVO:
            rfp_number = rfp["rfp_number"]
    if rfp_number is None:
        raise RuntimeError("Servo motor PR didn't land in RFP as expected.")

    q1 = rfx.record_quote(rfp_number, "ARKPWR", "ARK Power Controls Private Limited", 8500, 7)
    q2 = rfx.record_quote(rfp_number, "IPCS", "Ingenious Power & Control Systems", 8900, 10)
    q3 = rfx.record_quote(rfp_number, "KELTRON", "Kerala State Electronics Development Corporation", 8750, 14)
    rfx.select_winner(rfp_number, q1)  # ARKPWR wins on price and lead time

    po_results = rfx.generate_pos()
    po_number = po_results[0]["po_number"]

    import po_export
    hdr = pc.get_po_header(po_number)
    raw_items = pc.get_po_items(po_number)
    po_lns = [{"mat_code": it["material_code"], "mat_desc": it["material_desc"], "uom": it["uom"],
              "qty": it["quantity"], "deliv_date": it["delivery_date"], "deliv_loc": it["delivery_location"]}
             for it in raw_items]
    po_export.make_av_bytes(po_number, "ARKPWR", hdr["po_type"], hdr["legal_entity"], hdr["purchase_entity"],
                 hdr["purchasing_group"], hdr["currency"], hdr["plant_code"], po_lns)

    gr_id = gr.create_gr(po_number, {1: 4}, delivery_location=FACTORY,
                         received_by="Factory Stores", notes="Servo motor top-up — RFx award",
                         line_lots={1: {"serials": ["SN-SERVO-0009", "SN-SERVO-0010",
                                                     "SN-SERVO-0011", "SN-SERVO-0012"]}})
    acct.post_gr_entry(gr_id)
    _backdate_gr(gr_id, "2026-07-24")

    import vendor_invoices as vi
    invoice_id = vi.create_invoice(gr_id, invoice_number="ARK-INV-8842")
    payment = vi.record_payment(invoice_id, vi.get_invoice_payment_info(invoice_id)["amount"])

    return {"pr_id": pr_id, "rfp_number": rfp_number, "po_number": po_number,
            "gr_id": gr_id, "invoice_id": invoice_id, "payment": payment}


def phase5_bundle_procurement():
    """The rest of the BOM — everything except the two hero components,
    which are already resolved by Phase 3's transfer and Phase 4's RFx.
    Uses the existing 'Bandicoot Standard Unit — Full Build Kit' bundle
    at 2x (enough for the 2 units Phase 6 will actually build), showing
    off Purchase Bundles' real value: one shot instead of 48 separate
    sourcing decisions for parts whose vendor and price are already
    known and don't need re-litigating each time."""
    import purchase_bundles as pb

    lines = pb.explode_bundle("BDL-00001", multiplier=2)
    lines = [l for l in lines if l["code"] not in (BATTERY, SERVO)]

    pr_id = "PR-20260721-102"
    pr_lines = [{"vendor": l["vendor"], "mat_code": l["code"], "mat_desc": l["desc"],
                "uom": l["uom"], "qty": l["qty"], "req_date": "2026-07-25",
                "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY)} for l in lines]
    pc.create_pr(pr_id, requester_id="REQ-PROD", requester_name="Production Planning",
                requester_dept="Manufacturing", project_id="MCC-IMC-FLEET", lines=pr_lines,
                pr_date="2026-07-21")
    result = pc.run()

    # Bundle lines whose Item Master entry has no vendor tag fall to RFP
    # instead of a PO — real behavior, not a bug, just needs the same
    # resolution any RFP needs: a quote, a winner, an award. Bulk-resolved
    # here with one quote each (list price, a category-appropriate
    # vendor) rather than walked through individually — this layer is
    # meant to show breadth, not repeat Phase 4's competitive-bidding
    # beat 33 more times.
    import rfx
    import po_export
    CATEGORY_VENDOR = {
        "MEC": ("PROBOTS", "Probots Engineering Pvt Ltd"),
        "ELE": ("IPCS", "Ingenious Power & Control Systems"),
        "COM": ("IPCS", "Ingenious Power & Control Systems"),
        "SEN": ("KELTRON", "Kerala State Electronics Development Corporation"),
        "PNU": ("RAGHAV", "Raghavendra Chemicals & Pneumatics"),
        "SAF": ("PLINTRO", "PlinTroNics Pvt Ltd"),
        "MAT": ("SINGHANIA", "Singhania International"),
    }
    open_rfps = rfx.get_open_rfps()
    for r in open_rfps:
        prefix = r["mat_code"].split("-")[1]
        vendor_id, vendor_name = CATEGORY_VENDOR.get(prefix, ("PROBOTS", "Probots Engineering Pvt Ltd"))
        price = po_export.get_item_by_code(r["mat_code"])["price"] or 1000
        quote_id = rfx.record_quote(r["rfp_number"], vendor_id, vendor_name, price, 10)
        rfx.select_winner(r["rfp_number"], quote_id)
    fallout_pos = rfx.generate_pos() if open_rfps else []

    import vendor_invoices as vi
    outcomes = []
    for summary in result["po_summary"]:
        po_number = summary["po_number"]
        pc.mark_po_created(po_number)
        hdr = pc.get_po_header(po_number)
        raw_items = pc.get_po_items(po_number)
        po_lns = [{"mat_code": it["material_code"], "mat_desc": it["material_desc"], "uom": it["uom"],
                  "qty": it["quantity"], "deliv_date": it["delivery_date"], "deliv_loc": it["delivery_location"]}
                 for it in raw_items]
        po_export.make_av_bytes(po_number, hdr["supplier_id"], hdr["po_type"], hdr["legal_entity"],
                     hdr["purchase_entity"], hdr["purchasing_group"], hdr["currency"],
                     hdr["plant_code"], po_lns)

        line_receipts = {it["po_item"]: it["quantity"] for it in raw_items}
        gr_id = gr.create_gr(po_number, line_receipts, delivery_location=FACTORY,
                             received_by="Factory Stores", notes="Bundle-driven build materials")
        acct.post_gr_entry(gr_id)
        _backdate_gr(gr_id, "2026-07-25")

        invoice_id = vi.create_invoice(gr_id, invoice_number=f"BULK-{po_number}")
        payment = vi.record_payment(invoice_id, vi.get_invoice_payment_info(invoice_id)["amount"])
        outcomes.append({"po_number": po_number, "gr_id": gr_id,
                         "invoice_id": invoice_id, "payment": payment})

    # Same GR/Invoice/Payment treatment for the fallout POs — already
    # 'Created' immediately by generate_pos(), no separate send step
    for summary in fallout_pos:
        po_number = summary["po_number"]
        hdr = pc.get_po_header(po_number)
        raw_items = pc.get_po_items(po_number)
        line_receipts = {it["po_item"]: it["quantity"] for it in raw_items}
        gr_id = gr.create_gr(po_number, line_receipts, delivery_location=FACTORY,
                             received_by="Factory Stores", notes="Bundle-driven build materials")
        acct.post_gr_entry(gr_id)
        _backdate_gr(gr_id, "2026-07-25")
        invoice_id = vi.create_invoice(gr_id, invoice_number=f"BULK-{po_number}")
        payment = vi.record_payment(invoice_id, vi.get_invoice_payment_info(invoice_id)["amount"])
        outcomes.append({"po_number": po_number, "gr_id": gr_id,
                         "invoice_id": invoice_id, "payment": payment})

    return {"pr_id": pr_id, "pos": outcomes}


def phase6_production():
    """2 of the 4 ordered units get built — the other 2 are left
    genuinely unbuilt (not enough of most components procured for
    them), a real, visible 'still in progress' state rather than a
    manufactured one. Quality Inspection in this app is GR-side only
    (incoming materials from a vendor) — there's no equivalent QC gate
    on finished-goods output, so this goes straight to Fulfillment."""
    import production as prod
    unit1 = prod.confirm_production(BANDICOOT, 1, FACTORY, confirmed_by="Production Team",
                                    notes="Unit 1 of 3 — Mysore City Corporation fleet order")
    unit2 = prod.confirm_production(BANDICOOT, 1, FACTORY, confirmed_by="Production Team",
                                    notes="Unit 1 of 1 — Indore Municipal Corporation priority order")
    conn = db.get_connection()
    conn.execute("UPDATE production_confirmations SET confirmation_date=? WHERE confirmation_id=?",
                ("2026-07-26", unit1["confirmation_id"]))
    conn.execute("UPDATE production_confirmations SET confirmation_date=? WHERE confirmation_id=?",
                ("2026-07-26", unit2["confirmation_id"]))
    conn.commit()
    conn.close()
    return {"unit_for_mcc": unit1, "unit_for_imc": unit2}


def phase7_fulfillment_and_billing():
    """MCC gets 1 of their 3 units shipped — a real partial fulfillment,
    2 still open, not a manufactured gap. IMC's single-unit order ships
    complete. Both invoices bill off what actually shipped, not what
    was ordered — correctly IGST on both, Karnataka and Madhya Pradesh
    each being inter-state from Kerala."""
    import fulfillment as ful
    import billing as bl

    fid_mcc = ful.create_fulfillment("SO-00001")
    ful.record_shipment(fid_mcc, {BANDICOOT: 1}, carrier="VRL Logistics",
                        tracking_ref="VRL-MYS-88214")
    ful.record_delivery(fid_mcc, pod_reference="POD-MYS-001")
    acct.post_fulfillment_entry(fid_mcc)
    inv_mcc = bl.create_invoice(fid_mcc, due_days=30,
                                notes="Unit 1 of 3 — remaining 2 units to follow")
    bl.mark_issued(inv_mcc["invoice_id"])
    acct.post_invoice_entry(inv_mcc["invoice_id"])

    fid_imc = ful.create_fulfillment("SO-00002")
    ful.record_shipment(fid_imc, {BANDICOOT: 1}, carrier="Safexpress",
                        tracking_ref="SFX-IND-44190")
    ful.record_delivery(fid_imc, pod_reference="POD-IND-001")
    acct.post_fulfillment_entry(fid_imc)
    inv_imc = bl.create_invoice(fid_imc, due_days=30, notes="Priority order — complete")
    bl.mark_issued(inv_imc["invoice_id"])
    acct.post_invoice_entry(inv_imc["invoice_id"])

    conn = db.get_connection()
    for fid, ship_date in [(fid_mcc, "2026-07-27"), (fid_imc, "2026-07-27")]:
        conn.execute("UPDATE fulfillments SET shipped_date=?, delivered_date=? WHERE fulfillment_id=?",
                     (ship_date, ship_date, fid))
    conn.commit()
    conn.close()

    return {"fulfillment_mcc": fid_mcc, "invoice_mcc": inv_mcc,
            "fulfillment_imc": fid_imc, "invoice_imc": inv_imc}


def phase8_cash_application():
    """Mysore pays in full — clean close. Indore pays 60% now, citing a
    budget-release delay on their end — a real, common municipal
    payment pattern, and exactly the AR-aging/collections story worth
    having live in the demo data rather than every invoice closing out
    perfectly."""
    import cash_application as ca
    import billing as bl

    inv_mcc = [i for i in bl.get_invoices() if i["customer_id"] == "MCC"][-1]
    pay_mcc = ca.record_payment("MCC", inv_mcc["grand_total"], payment_date=date(2026, 7, 27),
                                payment_method="NEFT", reference_no="MCC-NEFT-77213",
                                notes="Full settlement — Unit 1 of 3")
    ca.apply_payment(pay_mcc, inv_mcc["invoice_id"], inv_mcc["grand_total"])

    inv_imc = [i for i in bl.get_invoices() if i["customer_id"] == "IMC"][-1]
    partial_amt = round(inv_imc["grand_total"] * 0.6, 2)
    pay_imc = ca.record_payment("IMC", partial_amt, payment_date=date(2026, 7, 27),
                                payment_method="RTGS", reference_no="IMC-RTGS-30587",
                                notes="Partial — balance pending municipal budget release")
    ca.apply_payment(pay_imc, inv_imc["invoice_id"], partial_amt,
                     short_payment_reason="Budget release delay — balance to follow")

    return {"payment_mcc": pay_mcc, "invoice_mcc": inv_mcc["invoice_id"],
            "payment_imc": pay_imc, "invoice_imc": inv_imc["invoice_id"],
            "imc_partial_amount": partial_amt}


def run_setup():
    """Phases 0-2 only — the 'before' state, for a live walkthrough. Leaves
    real transfer opportunities visible on the Inventory page for the
    presenter to execute themselves via the actual Execute buttons
    (Act 1 -> Act 2 of DEMO_GUIDE.md), rather than pre-resolving them.
    Pair with run_resolution() to fast-forward everything after that
    live moment, or run_all() instead of this for a single-click
    end-to-end preview with no live interaction at all."""
    print("Phase 0 — setup:", phase0_setup())
    print("Phase 1 — historical stock:", phase1_historical_stock())
    print("Phase 2 — demand:", phase2_demand())
    print("\nSetup done — transfer opportunities should now be visible on "
          "the Inventory page. Execute them there, then call run_resolution().")


def phase9_time_phased_planning_demo():
    """
    Deliberately runs LAST, after every other phase's own PR
    consolidation calls — pr_consolidation.run() sweeps every currently
    Open PR line tenant-wide with no scoping of its own, so adding
    these earlier would risk them being consolidated away by an
    earlier phase's own run() before anyone gets to see them sitting
    Open on the Time-Phased Planning tab.

    Leaves real, genuinely unconsolidated PRs on the books specifically
    so each of the three demand-signal modes has something real to
    show a first-time viewer, not an empty panel:

    1. 'Already Covered by Existing PR/PO' — GRB-MEC-0009 (Universal Joint
       Coupling) has a real, persistent 8-unit gap at the Factory even
       after every other phase completes (most of the 50 BOM
       components are never explicitly restocked by phases 4-5's
       narrower scope — genuinely still short, not staged). Covering
       it with a real Open PR, dated to its own natural stock-out,
       gives Sales Order Based mode's third outcome something to show
       — otherwise never demonstrable, since everything else in this
       script gets consolidated before anyone sees it as Open.
    2. 'Optimize Existing PRs' — Insufficient Lead Time: a Servo Motor
       PR at the Factory (Phase 4's own Factory-bound servo PR is
       already closed/consolidated by this point, so a fresh Open one
       here doesn't collide with it) with a required date tighter than
       the item's real 14-day lead time. Deliberately kept off R&D
       (the duplicate-pair PRs' own location below) so the two
       vignettes stay narratively distinct — otherwise this PR would
       also show up inside the duplicate group purely because it
       shares that group's location, muddying which effect is which.
    3. 'Optimize Existing PRs' — cross-location duplicate: two
       competing PRs for the same material at the same location
       (comfortable lead time on both, so the duplicate warning is the
       only thing firing, not also an insufficiency flag).
    """
    pc.create_pr("PR-20260901-201", requester_id="REQ-PLAN", requester_name="Planning",
                requester_dept="Manufacturing", project_id="TIME-PHASED-DEMO",
                pr_date="2026-08-25",
                lines=[{"vendor": "", "mat_code": "GRB-MEC-0009",
                        "mat_desc": "Universal Joint Coupling Precision Ground",
                        "uom": "pcs", "qty": 8, "req_date": "2026-08-10",
                        "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY)}])

    pc.create_pr("PR-20260901-202", requester_id="REQ-PROD", requester_name="Production Planning",
                requester_dept="Manufacturing", project_id="TIME-PHASED-DEMO",
                pr_date="2026-08-25",
                lines=[{"vendor": "", "mat_code": SERVO,
                        "mat_desc": "Servo Motor 12V High Torque Arm Joint Drive",
                        "uom": "pcs", "qty": 6, "req_date": "2026-08-05",
                        "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY)}])

    pc.create_pr("PR-20260901-203", requester_id="REQ-RND", requester_name="R&D Stores",
                requester_dept="R&D", project_id="TIME-PHASED-DEMO",
                pr_date="2026-08-26",
                lines=[{"vendor": "", "mat_code": SERVO,
                        "mat_desc": "Servo Motor 12V High Torque Arm Joint Drive",
                        "uom": "pcs", "qty": 8, "req_date": "2026-10-15",
                        "deliv_loc": RND, "deliv_geo": _geo(RND)}])
    pc.create_pr("PR-20260901-204", requester_id="REQ-PROD", requester_name="Production Planning",
                requester_dept="Manufacturing", project_id="TIME-PHASED-DEMO",
                pr_date="2026-08-27",
                lines=[{"vendor": "", "mat_code": SERVO,
                        "mat_desc": "Servo Motor 12V High Torque Arm Joint Drive",
                        "uom": "pcs", "qty": 5, "req_date": "2026-10-20",
                        "deliv_loc": RND, "deliv_geo": _geo(RND)}])

    return {"already_covered_pr": "PR-20260901-201", "insufficient_lead_time_pr": "PR-20260901-202",
            "duplicate_prs": ["PR-20260901-203", "PR-20260901-204"]}


def phase10_bundle_discovery_demo():
    """
    Real evidence for PR-US-04's bundle-discovery step
    (purchase_bundles.discover_bundle_candidates()) to find. The
    existing 'Bandicoot Standard Unit — Full Build Kit' bundle already
    covers ~50 of Genrobotics' 67 materials, so almost any pair drawn
    from within it is already-bundled and correctly excluded from
    candidacy — checked this directly before picking materials, rather
    than assuming any pair would work. These two (Camera & Sensor
    Module, Control & Power Module Assembly) are real, higher-level
    sub-assemblies genuinely outside that bundle, ordered together on
    two separate POs here — a real repeated pattern discovery can
    actually find, not a contrived one.

    No ordering constraint the way phase9 has (nothing here touches PR
    consolidation's own sweep) — kept at the end purely for narrative
    consistency with phases 9 and this file's own ordering.
    """
    header = {"po_type": "NB", "legal_entity": "LE-001", "purch_entity": "PE-001",
              "purch_group": "PG-001", "currency": "INR", "plant_code": FACTORY,
              "supplier_id": "VEN-GRB-ELEC", "supplier_name": "Precision Electronics Assemblies Pvt Ltd",
              "supplier_geo": ""}
    lines_common = [
        {"mat_code": "GRB-ASM-0002", "mat_desc": "Camera & Sensor Module - Complete Unit",
         "uom": "pcs", "unit_price": 68000.0, "deliv_date": "2026-08-20",
         "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY),
         "source_pr": "", "source_pr_line": "", "req_id": "", "req_dept": "", "project_id": ""},
        {"mat_code": "GRB-ASM-0004", "mat_desc": "Control & Power Module Assembly",
         "uom": "pcs", "unit_price": 72000.0, "deliv_date": "2026-08-20",
         "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY),
         "source_pr": "", "source_pr_line": "", "req_id": "", "req_dept": "", "project_id": ""},
    ]
    for po_number, qty in [("PO-DEMO-ASM1", 3), ("PO-DEMO-ASM2", 2)]:
        lines = [dict(l, qty=qty) for l in lines_common]
        pc.insert_po(po_number, header, lines)
        pc.mark_po_created(po_number, po_date="2026-08-01")
    return {"pos_created": ["PO-DEMO-ASM1", "PO-DEMO-ASM2"],
            "pattern": ["GRB-ASM-0002", "GRB-ASM-0004"]}


def phase11_multi_pr_consolidation_demo():
    """
    Three real, deliberately-open PRs — same material, same location,
    three different requesters/departments who each independently
    noticed they needed the same waterproof connector, with no
    coordination between them. Left genuinely Open (never consolidated
    by this script) specifically so a presenter can run PR
    Consolidation LIVE with the existing "1:1 PR -> PO Lines" toggle
    switched OFF, and watch three separate 20/15/10-unit requests
    merge into one real 45-unit PO line in front of the audience —
    the actual value proposition of consolidation, not just three
    lines with the same vendor sitting side by side.

    Positioned last, same reasoning as phase9 and phase10: run() (the
    consolidation this phase is deliberately NOT triggering) sweeps
    every currently-Open PR line tenant-wide with no scoping of its
    own, so anything created earlier in this script's own flow would
    already be gone by the time a presenter gets here if it ran before
    every other phase's own run() calls finished.
    """
    material = {"mat_code": "GRB-CAB-0002", "mat_desc": "Waterproof Connector 8-Pin",
                "uom": "pcs"}
    requesters = [
        ("PR-20260901-301", "REQ-ASM", "Assembly Team Lead", "Assembly", 20),
        ("PR-20260901-302", "REQ-QC", "QC Supervisor", "Quality Control", 15),
        ("PR-20260901-303", "REQ-LOG", "Logistics Coordinator", "Field Service", 10),
    ]
    for pr_number, req_id, req_name, dept, qty in requesters:
        pc.create_pr(pr_number, requester_id=req_id, requester_name=req_name,
                     requester_dept=dept, project_id="MULTI-PR-DEMO", pr_date="2026-08-25",
                     lines=[{"vendor": "", "mat_code": material["mat_code"],
                             "mat_desc": material["mat_desc"], "uom": material["uom"],
                             "qty": qty, "req_date": "2026-09-15", "deliv_loc": FACTORY,
                             "deliv_geo": _geo(FACTORY)}])
    return {"prs_created": [r[0] for r in requesters], "material": material["mat_desc"],
            "combined_qty": sum(r[4] for r in requesters)}


def phase12_reorder_qty_planning_demo():
    """
    Configures Reorder Qty Based demand for a routine consumable at
    R&D — matching IDS Denmed's own phase8, and deliberately a
    different material and location from phase11's Waterproof
    Connector (Factory), so each Time-Phased Planning mode's demo data
    stays narratively distinct rather than three different features all
    pointing at the same material. Hex bolts are a classic "keep some
    on hand, reorder roughly every N days" item — exactly the profile
    this mode is for, not tied to any Sales Order or PR.
    """
    import bom
    bom.set_planning_params("GRB-HW-0001", RND, min_qty=100, max_qty=400,
                            reorder_cadence_days=25)
    return {"material": "GRB-HW-0001", "location": RND, "min": 100, "max": 400, "cadence": 25}


def phase13_new_customer_pipeline():
    """
    Three new customers, three real quotation -> Sales Order -> bundle-
    based PR pipelines, staggered roughly a month apart across the next
    three months. Built specifically to make PR/PO consolidation and
    time-phased planning demonstrable at real scale: all three route
    through the SAME 'Bandicoot Standard Unit — Full Build Kit' bundle
    (BDL-00001, the same one phase5 already uses), so running PR
    Consolidation live doesn't merge one connector across three PRs —
    it can merge potentially all ~50 real components simultaneously
    across three genuinely independent customer orders. Every PR is
    left deliberately Open (no pc.run() call here) for that live demo.

    Quantities vary slightly (2/1/3 units) rather than being identical
    — realistic, and still shows the consolidation merge clearly since
    quantities don't need to match to combine. GSTINs are real,
    checksum-valid dummy values (vendor_onboarding.demo_gstin()) for
    each state, not arbitrary-looking strings, so nothing here would
    fail this platform's own GSTIN validation if re-entered by hand.

    Deliberately NOT called by run_resolution() or run_all() — this is
    its own, separately-triggered stage (the Settings page's own
    "Add New Customer Wave" button), meant to run live, in front of an
    audience, after they've already seen a smaller PR consolidation
    example (the Waterproof Connector PRs from phase11). The contrast
    between that small merge and this one — a handful of lines vs.
    potentially the whole BOM, three genuine customer orders deep — is
    the actual demo payoff; folding this into an already-automatic
    phase would collapse that staging into one flat reveal instead.
    Has no real dependency on any earlier phase (BDL-00001 is seed
    data, not phase-created) — it's safe to call any time there's at
    least one Active customer type on file, though the demo guide's
    own sequencing is what makes it land well, not any technical
    requirement.
    """
    import customer_onboarding as co
    import quotation as qt
    import sales_order as so
    import purchase_bundles as pb

    customers = [
        {"id": "SMC", "name": "SURAT MUNICIPAL CORPORATION", "type": "Municipal Corporation",
         "geo": "21.1702, 72.8311", "city": "Surat",
         "address": "Muglisara, Surat, Gujarat 395003",
         "contact": "Municipal Commissioner", "email": "commissioner@suratmunicipal.gov.in",
         "phone": "0261-2422244", "gstin": "24AAALS1234R1ZQ", "pan": "AAALS1234R",
         "credit_limit": 3000000, "qty": 2, "quote_date": "2026-07-25",
         "delivery_date": "2026-08-30", "pr_date": "2026-08-01",
         "pr_number": "PR-20260901-401", "project_id": "SMC-FLEET"},
        {"id": "GMC", "name": "GUWAHATI MUNICIPAL CORPORATION", "type": "Municipal Corporation",
         "geo": "26.1445, 91.7362", "city": "Guwahati",
         "address": "Panbazar, Guwahati, Assam 781001",
         "contact": "Municipal Commissioner", "email": "commissioner@gmc.assam.gov.in",
         "phone": "0361-2540111", "gstin": "18AAALG5678R1Z7", "pan": "AAALG5678R",
         "credit_limit": 2000000, "qty": 1, "quote_date": "2026-08-25",
         "delivery_date": "2026-09-30", "pr_date": "2026-09-01",
         "pr_number": "PR-20260901-402", "project_id": "GMC-FLEET"},
        {"id": "DJB", "name": "DELHI JAL BOARD", "type": "Government Utility",
         "geo": "28.6139, 77.2090", "city": "New Delhi",
         "address": "Varunalaya Phase II, Karol Bagh, New Delhi 110005",
         "contact": "Chief Engineer", "email": "ce@delhijalboard.nic.in",
         "phone": "011-25742420", "gstin": "07AAAGD9012R1Z4", "pan": "AAAGD9012R",
         "credit_limit": 4500000, "qty": 3, "quote_date": "2026-09-25",
         "delivery_date": "2026-10-30", "pr_date": "2026-10-01",
         "pr_number": "PR-20260901-403", "project_id": "DJB-FLEET"},
    ]

    results = []
    for c in customers:
        co.upsert_customer(c["id"], {
            "Customer_Name": c["name"], "Customer_Type": c["type"], "Geolocation": c["geo"],
            "City": c["city"], "Country": "India", "Address": c["address"],
            "Contact_Name": c["contact"], "Contact_Email": c["email"],
            "Contact_Phone": c["phone"], "GSTIN": c["gstin"], "PAN": c["pan"],
            "Credit_Limit": c["credit_limit"], "Payment_Terms": "Net 30",
        })
        co.approve_customer(c["id"])

        quote_id = qt.create_quote(c["id"],
            [{"mat_code": BANDICOOT, "mat_desc": "Bandicoot Robotic Scavenger - Standard Unit",
              "uom": "pcs", "qty": c["qty"], "unit_price": UNIT_PRICE}],
            notes=f"Sewer-cleaning fleet expansion — {c['name'].title()}")
        qt.mark_sent(quote_id)
        qt.record_response(quote_id, accepted=True)
        conn = db.get_connection()
        conn.execute("UPDATE quotes SET quote_date=? WHERE quote_id=?", (c["quote_date"], quote_id))
        conn.commit()
        conn.close()

        order = so.create_order_from_quote(quote_id, delivery_location=FACTORY,
            delivery_geo=_geo(FACTORY), requested_delivery_date=c["delivery_date"],
            notes=f"{c['name'].title()} — fleet expansion")

        lines = pb.explode_bundle("BDL-00001", multiplier=c["qty"])
        pr_lines = [{"vendor": l["vendor"], "mat_code": l["code"], "mat_desc": l["desc"],
                    "uom": l["uom"], "qty": l["qty"], "req_date": c["pr_date"],
                    "deliv_loc": FACTORY, "deliv_geo": _geo(FACTORY)} for l in lines]
        pc.create_pr(c["pr_number"], requester_id="REQ-PROD", requester_name="Production Planning",
                    requester_dept="Manufacturing", project_id=c["project_id"], lines=pr_lines,
                    pr_date=c["pr_date"])

        results.append({"customer": c["id"], "quote": quote_id, "so": order["so_id"],
                        "so_status": order["status"], "pr": c["pr_number"],
                        "qty": c["qty"], "delivery_date": c["delivery_date"]})
    return results


# Pilot-agnostic alias — the Agent Console (and anything else driving
# multiple demo_scenario modules interchangeably) calls this name
# regardless of which pilot is active, rather than hardcoding either
# this module's own phase13_new_customer_pipeline or demo_scenario_ids
# .py's differently-numbered phase9_new_customer_pipeline. Each
# module's own phase function keeps its own real number, matching its
# own phase sequence — this is purely an external-calling convenience,
# not a rename.
new_customer_wave = phase13_new_customer_pipeline


def run_resolution():
    """Phases 4-8 — deliberately skips phase3 (the transfer), on the
    assumption the presenter already executed it live via the real
    Execute buttons on the Inventory page (that's the whole point of
    splitting this out — Act 2 needs to actually be interactive, not
    replayed). Safe to call regardless of whether the transfer moved
    stock through the UI or through phase3_execute_transfer() directly
    — phase4's servo-motor shortfall is a fixed number either way,
    since the transfer quantities are deterministic given phase1/2's
    fixed historical stock and demand, not read live from whichever
    path produced them."""
    print("Phase 4 — procure remaining servo motors:", phase4_procure_remaining_servo_motors())
    print("Phase 5 — bundle procurement:", phase5_bundle_procurement())
    print("Phase 6 — production:", phase6_production())
    print("Phase 7 — fulfillment & billing:", phase7_fulfillment_and_billing())
    print("Phase 8 — cash application:", phase8_cash_application())
    print("Phase 9 — time-phased planning demo data:", phase9_time_phased_planning_demo())
    print("Phase 10 — bundle discovery demo data:", phase10_bundle_discovery_demo())
    print("Phase 11 — multi-PR consolidation demo data:", phase11_multi_pr_consolidation_demo())
    print("Phase 12 — reorder-qty planning demo data:", phase12_reorder_qty_planning_demo())
    print("(Phase 13 — new customer wave — deliberately NOT run here; triggered "
          "separately, live, via the Settings page own button, for a staged "
          "small-then-large PR consolidation demo. See phase13_new_customer_pipeline().)")
    print("\nDone.")


def run_all():
    """The whole story, start to finish, on top of whatever master data
    is currently seeded. Prints a short summary after each phase so a
    failure anywhere is immediately traceable to its phase."""
    print("Phase 0 — setup:", phase0_setup())
    print("Phase 1 — historical stock:", phase1_historical_stock())
    print("Phase 2 — demand:", phase2_demand())
    print("Phase 3 — execute transfer:", phase3_execute_transfer())
    print("Phase 4 — procure remaining servo motors:", phase4_procure_remaining_servo_motors())
    print("Phase 5 — bundle procurement:", phase5_bundle_procurement())
    print("Phase 6 — production:", phase6_production())
    print("Phase 7 — fulfillment & billing:", phase7_fulfillment_and_billing())
    print("Phase 8 — cash application:", phase8_cash_application())
    print("Phase 9 — time-phased planning demo data:", phase9_time_phased_planning_demo())
    print("Phase 10 — bundle discovery demo data:", phase10_bundle_discovery_demo())
    print("Phase 11 — multi-PR consolidation demo data:", phase11_multi_pr_consolidation_demo())
    print("Phase 12 — reorder-qty planning demo data:", phase12_reorder_qty_planning_demo())
    print("(Phase 13 — new customer wave — deliberately NOT run here; triggered "
          "separately, live, via the Settings page own button, for a staged "
          "small-then-large PR consolidation demo. See phase13_new_customer_pipeline().)")
    print("\nDone.")


if __name__ == "__main__":
    run_all()