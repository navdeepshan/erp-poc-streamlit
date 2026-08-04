"""
mfg_ui.py — Manufacturing, as its own Streamlit app.

Same reasoning as the O2C split: this is a genuinely distinct process area
(production support on the S2P side — Goods Receipt, Three-Way Match,
Quality Inspection, BOM/MRP), not a natural extension of erp_ui.py's
7-page S2P app or o2c_ui.py's O2C app. Keeping it separate means neither
of those has to grow to accommodate it, and this app can be iterated on
without any risk to either.

Reuses backend modules where they already fit exactly like O2C did:
pr_consolidation.py for shared Excel-styling helpers and the PO schema
this module reads, po_export.py for the item catalog. No UI code is
shared — same reasoning as o2c_ui.py's docstring: importing another
app's UI module would execute that module's own page config and router,
which is exactly the coupling this split exists to avoid.

Run with: streamlit run mfg_ui.py
"""

import streamlit as st
import openpyxl
import pandas as pd
import os, sys, json, traceback
from datetime import date

import db
import po_export
import shipping
import sto
import reservation as res
import pr_consolidation
import goods_receipt as gr
import quality_inspection as qi
import bom
import org_defaults as od
import inventory as inv
import production as prod
import accounting as acct
import vendor_invoices as vi

def load_delivery_locs():
    """Delivery_Locations now lives in SQLite — delegates to
    pr_consolidation.py's canonical reader (also used by erp_ui.py and
    o2c_ui.py) instead of this file's own separate Excel read."""
    try:
        return pr_consolidation.get_delivery_locations(active_only=True)
    except Exception:
        return []

_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(_DIR, "data.xlsx")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Manufacturing", page_icon="\U0001f527", layout="wide")

from ui_theme import apply_theme
apply_theme()


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### \U0001f527 ERP Suite")
    st.caption("Manufacturing")
    st.divider()
    page = st.radio("", ["\U0001f4e5  Goods Receipt",
                         "\U0001f50e  Quality Inspection",
                         "\U0001f9e9  BOM & Explosion",
                         "\U0001f3ed  Production",
                         "\U0001f4e6  Inventory"],
                    label_visibility="collapsed")
    st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Goods Receipt
# ══════════════════════════════════════════════════════════════════════════════
def page_goods_receipt():
    st.markdown("## \U0001f4e5 Goods Receipt")
    st.caption("Captures received quantities separately from ordered quantities, "
               "supporting partial receipt across one or more Goods Receipts against "
               "the same PO. This is the event Three-Way Match and Quality Inspection "
               "both rely on.")
    st.divider()

    s = gr.stats()
    receivable = gr.get_receivable_pos()
    m1, m2, m3 = st.columns(3)
    m1.metric("Total GRs", s["total"])
    m2.metric("Posted", s["by_status"].get("Posted", 0))
    m3.metric("POs Awaiting Receipt", len(receivable))
    st.divider()

    tab1, tab2 = st.tabs(["\u2795 Create GR", "\U0001f4cb Manage GRs"])

    # ── TAB 1 — receive against a PO ──────────────────────────────────────────────
    with tab1:
        if not receivable:
            st.info("No POs with outstanding lines to receive against.")
        else:
            labels = {p["po_number"]: f"{p['po_number']} — {p['vendor_name']} "
                      f"({p['outstanding_lines']} of {p['lines']} line(s) outstanding)"
                      for p in receivable}
            sel_po = st.selectbox("Purchase Order", list(labels.keys()),
                                  format_func=lambda k: labels[k], key="gr_po_sel")

            lines = gr.get_po_receipt_status(sel_po)
            # Batch: one call covers every line of this PO, instead of one
            # call per line — the per-line version was the confirmed cause
            # of a 2.8s render for a 7-line PO with 6-way PR consolidation.
            source_prs_by_line = gr.get_po_lines_source_prs(sel_po)
            loc_names = {d["id"]: d["name"] for d in load_delivery_locs()}
            st.markdown("##### Enter quantity received per line")
            line_qtys = {}
            pr_overrides = {}
            for i, ln in enumerate(lines):
                po_item = ln["po_item"]
                c1, c2 = st.columns([3, 1.3])
                loc_label = loc_names.get(ln.get("deliv_loc"), ln.get("deliv_loc") or "")
                c1.write(f"**{ln['mat_desc']}**  \n{ln['mat_code']} \u00b7 {ln['uom']} \u00b7 "
                        f"ordered {ln['po_qty']:g} \u00b7 already received {ln['received_qty']:g} "
                        f"\u00b7 outstanding {ln['outstanding_qty']:g}  \n"
                        f"\U0001f4cd *{loc_label}*")
                default_qty = ln["outstanding_qty"] if ln["receipt_status"] != "Fully Received" else 0
                qty_received = c2.number_input("Qty received", min_value=0.0,
                    value=float(default_qty), key=f"gr_qty_{sel_po}_{i}", label_visibility="collapsed")
                line_qtys[po_item] = qty_received

                source_prs = source_prs_by_line.get(po_item, [])
                if len(source_prs) > 1 and qty_received > 0:
                    with st.expander(f"\U0001f9ee This line consolidates {len(source_prs)} PR lines — "
                                     "review the split before confirming"):
                        default_alloc = gr.compute_default_allocation(source_prs, qty_received)
                        st.caption("Default is proportional by requested quantity — adjust any line "
                                  "below (e.g. for a more urgent PR) before creating the GR. "
                                  "**The total allocated below is what actually gets recorded as "
                                  "received for this line** — not the field above, which only sets "
                                  "the starting point for the default split.")
                        line_overrides = []
                        for sp in source_prs:
                            key = (sp["pr_number"], sp["pr_line"])
                            default_val = default_alloc.get(key, 0)
                            oc1, oc2 = st.columns([3, 1.3])
                            oc1.write(f"{sp['pr_number']} line {sp['pr_line']} "
                                     f"(requested {sp['requested_qty']:g}, already accepted {sp['accepted_qty']:g})")
                            v = oc2.number_input("Allocate", min_value=0.0, value=float(default_val),
                                key=f"gr_alloc_{sel_po}_{i}_{sp['pr_number']}_{sp['pr_line']}",
                                label_visibility="collapsed")
                            line_overrides.append((sp["pr_number"], sp["pr_line"], v))
                        allocated_total = sum(v for _, _, v in line_overrides)
                        if abs(allocated_total - qty_received) > 0.01:
                            st.info(f"\u2139\ufe0f This line will record **{allocated_total:g} received** "
                                   f"(the sum below), not the {qty_received:g} shown above.")
                        else:
                            st.success(f"\u2705 This line will record {allocated_total:g} received.")
                        pr_overrides[po_item] = line_overrides

            c1, c2 = st.columns(2)
            with c1:
                deliv_locs = load_delivery_locs()
                if deliv_locs:
                    loc_labels = {d["id"]: d["name"] for d in deliv_locs}
                    po_default_loc = gr.get_po_delivery_location(sel_po)
                    default_idx = 0
                    if po_default_loc in loc_labels:
                        default_idx = list(loc_labels.keys()).index(po_default_loc)
                    sel_loc_id = st.selectbox("Delivery Location", list(loc_labels.keys()),
                        format_func=lambda k: loc_labels[k], index=default_idx, key="gr_deliv")
                    deliv_loc = sel_loc_id
                    st.caption("Defaults to the PO's own delivery location — change this only if "
                              "the goods actually arrived somewhere else.")
                else:
                    st.warning("No delivery locations set up — add one on the S2P app first.")
                    deliv_loc = ""
            with c2:
                received_by = st.text_input("Received By", key="gr_recvby")
            notes = st.text_input("Notes", key="gr_notes")

            create_clicked = st.button("\U0001f4e5 Create Goods Receipt", type="primary", key="gr_create")
            if create_clicked:
                try:
                    gr_id = gr.create_gr(sel_po, line_qtys, deliv_loc, received_by, notes,
                                         pr_allocations=pr_overrides or None)
                    st.cache_data.clear()
                    st.success(f"\u2705 {gr_id} created for {sel_po}. Head to **Manage GRs** "
                              "to generate the receipt note.")
                    try:
                        je_id = acct.post_gr_entry(gr_id)
                        st.success(f"\U0001f4d2 Posted to ledger as {je_id} "
                                  "(Inventory / GST Input / Accounts Payable).")
                    except Exception:
                        st.warning("\u26a0\ufe0f Goods receipt recorded, but the accounting "
                                  "entry failed to post:")
                        st.code(traceback.format_exc())
                except Exception:
                    st.error("\u274c Error creating goods receipt:")
                    st.code(traceback.format_exc())

    # ── TAB 2 — view and manage existing GRs ──────────────────────────────────────
    with tab2:
        grs = gr.get_grs()
        if not grs:
            st.info("No goods receipts yet.")
        else:
            gdf = pd.DataFrame(grs)[["gr_id","po_number","vendor_name","status","gr_date"]]
            gdf.columns = ["GR","PO","Vendor","Status","Date"]
            st.dataframe(gdf, use_container_width=True, hide_index=True)

            labels = {g["gr_id"]: f"{g['gr_id']} — {g['po_number']} ({g['status']})" for g in grs}
            sel = st.selectbox("Work on", list(labels.keys()), format_func=lambda k: labels[k], key="gr_work_sel")
            g = next(x for x in grs if x["gr_id"] == sel)
            items = gr.get_gr_items(sel)

            left, right = st.columns([3, 2], gap="large")
            with left:
                st.markdown("##### \U0001f4c4 Line Items")
                idf = pd.DataFrame(items)[["mat_code","mat_desc","uom","po_qty","qty_received"]]
                idf.columns = ["Code","Description","UOM","PO Qty","Received"]
                st.dataframe(idf, use_container_width=True, hide_index=True)
                st.caption(f"Vendor: {g['vendor_name']}  \u00b7  Received by: {g['received_by'] or 'n/a'}")

            with right:
                st.markdown("##### \u2696\ufe0f Actions")
                gen_clicked = st.button("\U0001f4c4 Generate Receipt Note", key="gr_gen_doc")
                if gen_clicked:
                    fname, fbytes = gr.generate_gr_note(sel)
                    st.session_state.gr_generated_doc = {"gr_id": sel, "filename": fname, "bytes": fbytes}
                gd = st.session_state.get("gr_generated_doc")
                if gd and gd["gr_id"] == sel:
                    st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                        file_name=gd["filename"], mime=XLSX_MIME, key=f"gr_dl_{sel}")

                if g["status"] == "Posted":
                    with st.popover("\u274c Cancel GR"):
                        reason = st.text_input("Reason", key=f"gr_cancel_reason_{sel}")
                        if st.button("Confirm Cancellation", key=f"gr_cancel_btn_{sel}"):
                            gr.cancel_gr(sel, reason)
                            st.cache_data.clear()
                            st.success(f"{sel} cancelled.")
                            st.rerun()
                else:
                    st.warning("Cancelled.")

                st.divider()
                st.markdown("##### \U0001f4b0 Vendor Invoice")
                st.caption("Stage 2 of the proper 3-way match: GR posted Dr Inventory / "
                          "Cr GR/IR Clearing (ex-GST) — this clears that and, for the "
                          "first time, credits Accounts Payable. GST is determined here, "
                          "not at GR time, matching real invoice-verification practice.")
                vinv = vi.invoice_for_gr(sel)
                if vinv is None:
                    if g["status"] != "Posted":
                        st.caption("Only a Posted GR can be invoiced.")
                    else:
                        amounts = vi.compute_invoice_amounts(sel)
                        inv_no = st.text_input("Vendor's invoice number", key=f"vi_invno_{sel}")
                        c1, c2 = st.columns(2)
                        c1.metric("GR/IR Clearing", f"\u20b9{amounts['clearing_amount']:,.2f}")
                        gst_total = amounts['cgst'] + amounts['sgst'] + amounts['igst']
                        c2.metric("GST Input", f"\u20b9{gst_total:,.2f}")
                        st.caption(f"**Total: \u20b9{amounts['total']:,.2f}** — computed from "
                                  "the GR's own items, not editable here. A vendor invoice "
                                  "differing from this would represent a price variance, "
                                  "which isn't tracked here.")
                        if st.button("\U0001f9fe Simulate Invoice", key=f"vi_create_{sel}"):
                            try:
                                invoice_id = vi.create_invoice(sel, invoice_number=inv_no)
                                st.cache_data.clear()
                                st.success(f"\u2705 {invoice_id} recorded — \u20b9{amounts['total']:,.2f} "
                                          f"owed to {g['vendor_name']}.")
                            except Exception as e:
                                st.error(f"\u274c {e}")
                else:
                    info = vi.get_invoice_payment_info(vinv["invoice_id"])
                    st.write(f"**{vinv['invoice_id']}**" +
                            (f"  \u00b7  vendor ref: {vinv['invoice_number']}" if vinv['invoice_number'] else ""))
                    st.caption(f"Invoiced \u20b9{info['amount']:,.2f}  \u00b7  Paid \u20b9{info['paid_amount']:,.2f}  "
                              f"\u00b7  Due \u20b9{info['balance_due']:,.2f}  \u00b7  **{info['status']}**")
                    if info["balance_due"] > 0.005:
                        pay_amt = st.number_input("Payment amount", min_value=0.0,
                            max_value=float(info["balance_due"]), value=float(info["balance_due"]),
                            key=f"vi_payamt_{vinv['invoice_id']}",
                            help="Defaults to the full balance due — reduce for a partial payment.")
                        if st.button("\U0001f4b8 Simulate Payment", key=f"vi_pay_{vinv['invoice_id']}"):
                            try:
                                result = vi.record_payment(vinv["invoice_id"], pay_amt)
                                st.cache_data.clear()
                                st.success(f"\u2705 {result['payment_id']} recorded, posted as "
                                          f"{result['je_id']}.")
                            except Exception as e:
                                st.error(f"\u274c {e}")
                    else:
                        st.success("\u2705 Fully paid.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Quality Inspection
# ══════════════════════════════════════════════════════════════════════════════
def page_quality_inspection():
    st.markdown("## \U0001f50e Quality Inspection")
    st.caption("Record-only — inspection results here never block Goods Receipt or "
               "anything else. They provide a warning signal for a payment-proposal "
               "process to check before release, not an enforcement gate. 'Not yet "
               "inspected' and 'Failed' are tracked as distinct statuses, since they "
               "carry very different weight for downstream review.")
    st.divider()

    s = qi.stats()
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Passed", s["by_status"].get(qi.STATUS_PASSED, 0))
    m2.metric("Partial Pass", s["by_status"].get(qi.STATUS_PARTIAL, 0))
    m3.metric("Failed", s["by_status"].get(qi.STATUS_FAILED, 0))
    m4.metric("In Progress", s["by_status"].get(qi.STATUS_IN_PROGRESS, 0))
    m5.metric("GRs Pending", s["grs_pending"])
    st.divider()

    tab1, tab2 = st.tabs(["\u2795 Record Inspection", "\U0001f4cb Manage Inspections"])

    # ── TAB 1 — record results per GR line ────────────────────────────────────────
    with tab1:
        pending = qi.get_grs_needing_inspection()
        if not pending:
            st.info("No GRs with lines still awaiting inspection.")
        else:
            labels = {p["gr_id"]: f"{p['gr_id']} — {p['vendor_name']} "
                      f"({p['pending_lines']} of {p['lines']} line(s) pending)" for p in pending}
            sel_gr = st.selectbox("Goods Receipt", list(labels.keys()),
                                  format_func=lambda k: labels[k], key="qi_gr_sel")

            lines = qi.get_gr_quality_status(sel_gr)
            st.markdown("##### Enter results per line")
            if st.button("\u2705 Mark all lines fully passed", key=f"qi_markall_{sel_gr}",
                        help="Sets Passed = Received and Failed = 0 for every line below — "
                             "then just adjust any line that actually had a failure."):
                for i, ln in enumerate(lines):
                    st.session_state[f"qi_pass_{sel_gr}_{i}"] = float(ln["qty_received"] or 0)
                    st.session_state[f"qi_fail_{sel_gr}_{i}"] = 0.0
                st.rerun()
            for i, ln in enumerate(lines):
                c1, c2, c3 = st.columns([3, 1, 1])
                c1.write(f"**{ln['mat_desc']}**  \n{ln['mat_code']} \u00b7 received "
                        f"{ln['qty_received']:g} \u00b7 current status: **{ln['status']}**")
                passed = c2.number_input("Passed", min_value=0.0,
                    value=float(ln['qty_passed'] or 0), key=f"qi_pass_{sel_gr}_{i}")
                failed = c3.number_input("Failed", min_value=0.0,
                    value=float(ln['qty_failed'] or 0), key=f"qi_fail_{sel_gr}_{i}")
                st.session_state[f"qi_line_{sel_gr}_{i}"] = (ln['po_item'], ln['mat_code'], passed, failed)

            inspected_by = st.text_input("Inspected By", key="qi_inspector")
            notes = st.text_input("Notes", key="qi_notes")

            save_clicked = st.button("\U0001f50e Save Inspection Results", type="primary", key="qi_save")
            if save_clicked:
                errors = []
                saved = 0
                for i in range(len(lines)):
                    po_item, mat_code, passed, failed = st.session_state[f"qi_line_{sel_gr}_{i}"]
                    if passed == 0 and failed == 0:
                        continue  # untouched line, skip
                    try:
                        qi.record_inspection(sel_gr, po_item, passed, failed, inspected_by, notes)
                        saved += 1
                    except Exception as e:
                        errors.append(f"{mat_code}: {e}")
                st.cache_data.clear()
                if saved:
                    st.success(f"\u2705 Recorded results for {saved} line(s).")
                for e in errors:
                    st.error(f"\u274c {e}")
                if saved:
                    st.rerun()

    # ── TAB 2 — view results, generate report ─────────────────────────────────────
    with tab2:
        all_grs = gr.get_grs(status="Posted")
        if not all_grs:
            st.info("No goods receipts yet.")
        else:
            labels = {g["gr_id"]: f"{g['gr_id']} — {g['po_number']} ({g['vendor_name']})" for g in all_grs}
            sel = st.selectbox("Goods Receipt", list(labels.keys()), format_func=lambda k: labels[k], key="qi_view_sel")
            lines = qi.get_gr_quality_status(sel)
            ldf = pd.DataFrame(lines)[["mat_code","mat_desc","qty_received","qty_passed","qty_failed","status"]]
            ldf.columns = ["Code","Description","Received","Passed","Failed","Status"]
            st.dataframe(ldf, use_container_width=True, hide_index=True)

            gen_clicked = st.button("\U0001f4c4 Generate Inspection Report", key="qi_gen_doc")
            if gen_clicked:
                fname, fbytes = qi.generate_inspection_report(sel)
                st.session_state.qi_generated_doc = {"gr_id": sel, "filename": fname, "bytes": fbytes}
            gd = st.session_state.get("qi_generated_doc")
            if gd and gd["gr_id"] == sel:
                st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                    file_name=gd["filename"], mime=XLSX_MIME, key=f"qi_dl_{sel}")

            po_status = qi.get_po_quality_status(gr.get_gr(sel)["po_number"])
            st.divider()
            st.markdown("##### \U0001f4e6 Full PO quality picture (what a payment-proposal system would see)")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Not Yet Inspected", po_status["not_yet_inspected"])
            c2.metric("Passed", po_status["passed_lines"])
            c3.metric("Partial", po_status["partial_lines"])
            c4.metric("Failed", po_status["failed_lines"])
            if po_status["has_failures"]:
                st.warning("\u26a0\ufe0f This PO has failed or partial lines — a payment "
                          "proposal system would flag this, not block it.")
            elif po_status["not_yet_inspected"] > 0:
                st.info("\u2139\ufe0f Some lines on this PO haven't been inspected yet.")
            elif po_status["clean"]:
                st.success("\u2705 Every line on this PO passed inspection cleanly.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — BOM & Explosion
# ══════════════════════════════════════════════════════════════════════════════
def page_bom():
    st.markdown("## \U0001f9e9 BOM & Explosion")
    st.caption("Manual trigger, not continuous MRP. Nets against BOTH on-hand inventory "
               "at the delivery location AND open PO exposure — genuine gross-to-net now "
               "that a real inventory ledger exists. Generated PR lines flow through the "
               "exact same PO-vs-RFP split as any other PR.")
    st.divider()

    s = bom.stats()
    m1, m2 = st.columns(2)
    m1.metric("Finished Goods / Assemblies", s["finished_goods"])
    m2.metric("BOM Lines", s["bom_lines"])
    st.divider()

    tab1, tab2 = st.tabs(["\U0001f4a5 Explode & Propose PR", "\U0001f333 View BOM"])

    # ── TAB 1 — explode and propose ───────────────────────────────────────────────
    with tab1:
        fgs = bom.get_finished_goods()
        if not fgs:
            st.info("No BOMs defined yet.")
        else:
            matched_sos = bom.get_bom_matched_open_sales_orders()
            source = st.radio("Explode from",
                ["\U0001f9e9 Pick a finished good manually"] +
                (["\U0001f4e6 A confirmed Sales Order"] if matched_sos else []),
                key="bom_source", horizontal=True)

            so_link = None
            if source.startswith("\U0001f4e6") and matched_sos:
                so_labels = {f"{m['so_id']}|{m['mat_code']}": f"{m['so_id']} — {m['customer_name']} "
                            f"— {m['qty']:g} x {m['mat_desc']}" for m in matched_sos}
                sel_so_key = st.selectbox("Confirmed Sales Order", list(so_labels.keys()),
                    format_func=lambda k: so_labels[k], key="bom_so_sel")
                so_link = next(m for m in matched_sos
                               if f"{m['so_id']}|{m['mat_code']}" == sel_so_key)
                sel_fg = so_link["mat_code"]
                qty = so_link["qty"]
                st.caption(f"Building {qty:g} x {sel_fg} for {so_link['so_id']} "
                          f"({so_link['customer_name']}) — quantity and finished good are "
                          "fixed by the order; delivery location below defaults to the "
                          "order's, but can be changed.")
            else:
                labels = {f["code"]: f"{f['code']} — {f['desc']}" for f in fgs}
                sel_fg = st.selectbox("Finished Good / Assembly", list(labels.keys()),
                                      format_func=lambda k: labels[k], key="bom_fg_sel")
                qty = st.number_input("Quantity to build", min_value=1, value=1, key="bom_qty")

            net_pos = st.checkbox("Net against on-hand inventory + open PO exposure",
                                  value=True, key="bom_net")

            deliv_locs = load_delivery_locs()
            if deliv_locs:
                loc_labels = {d["id"]: d["name"] for d in deliv_locs}
                default_idx = 0
                if so_link:
                    # Sales Orders now store location_id here too (see o2c_ui.py) —
                    # matching on id, not name, is what actually makes this default
                    # line up with the order that triggered this explosion.
                    match_idx = next((i for i, d in enumerate(deliv_locs)
                                      if d["id"] == so_link["delivery_location"]), None)
                    if match_idx is not None:
                        default_idx = match_idx
                sel_loc_id = st.selectbox("Delivery Location", list(loc_labels.keys()),
                    format_func=lambda k: loc_labels[k], index=default_idx, key="bom_deliv")
                # Store the ID, not the display name — inventory balances (on-hand,
                # open-PO netting) are keyed by location_id everywhere else in the
                # app (GR posts against the PO line's own Delivery_Location, which
                # is an ID). Passing the name here instead used to make on-hand
                # stock invisible to this exact netting call: get_balance() would
                # look up "Genrobotics Manufacturing Plant - Kanjikode Palakkad"
                # while every GR had posted stock under "GRB_DL_PKD_Factory" — same
                # real place, two different keys, so on-hand always read as 0.
                deliv_loc = sel_loc_id
                deliv_geo = next(d["geo"] for d in deliv_locs if d["id"] == sel_loc_id)
            else:
                st.warning("No delivery locations set up — add one on the S2P app first.")
                deliv_loc, deliv_geo = "", ""

            detailed = bom.explode_bom_detailed(sel_fg, qty)
            lines = bom.net_requirements(detailed, deliv_loc if net_pos else None) if net_pos else \
                    [{**d, "on_hand_qty": 0, "open_po_qty": 0, "net_qty": d["gross_qty"]} for d in detailed]

            st.markdown(f"##### Exploded requirement — {len(lines)} line(s) to requisition")
            if lines:
                ldf = pd.DataFrame(lines)[["mat_code","mat_desc","gross_qty","on_hand_qty","open_po_qty","net_qty"]]
                ldf.columns = ["Code","Description","Gross Qty","On Hand","Open PO Qty","Net to Requisition"]
                st.dataframe(ldf, use_container_width=True, hide_index=True)
            else:
                st.success("\u2705 Nothing to requisition — fully covered by on-hand stock and open POs.")

            st.markdown("##### PR details")
            c1, c2 = st.columns(2)
            with c1:
                requester_id = st.text_input("Requester ID", key="bom_req_id")
                requester_name = st.text_input("Requester Name", key="bom_req_name")
                requester_dept = st.text_input("Department", value="Manufacturing", key="bom_req_dept")
            with c2:
                default_proj = so_link["so_id"] if so_link else ""
                project_id = st.text_input("Project ID", value=default_proj, key="bom_proj",
                    help="Set to the Sales Order ID when exploding from one, for traceability.")
                required_date = st.date_input("Required By", key="bom_req_date")

            propose_clicked = st.button("\U0001f9e9 Propose Purchase Requisition", type="primary",
                                        key="bom_propose", disabled=not lines)
            if propose_clicked:
                try:
                    result = bom.propose_pr_lines(sel_fg, qty, requester_id, requester_name,
                        requester_dept, project_id, deliv_loc, deliv_geo, required_date, net_pos)
                    st.cache_data.clear()
                    st.success(f"\u2705 {result['pr_number']} created — {result['lines']} line(s), "
                              f"{result['with_vendor']} with a preferred vendor already. "
                              "Head to Consolidate on the S2P app to route them.")
                except Exception:
                    st.error("\u274c Error proposing PR:")
                    st.code(traceback.format_exc())

    # ── TAB 2 — view a BOM tree ───────────────────────────────────────────────────
    with tab2:
        fgs = bom.get_finished_goods()
        if not fgs:
            st.info("No BOMs defined yet.")
        else:
            labels = {f["code"]: f"{f['code']} — {f['desc']}" for f in fgs}
            sel = st.selectbox("Finished Good / Assembly", list(labels.keys()),
                               format_func=lambda k: labels[k], key="bom_view_sel")
            children = bom.get_bom(sel)
            cdf = pd.DataFrame(children)[["component_code","component_desc","qty_per","uom"]]
            cdf.columns = ["Component","Description","Qty Per Unit","UOM"]
            st.dataframe(cdf, use_container_width=True, hide_index=True)
            st.caption("Direct children only — components with their own BOM (sub-assemblies) "
                      "can be selected above too, to drill into their own components.")

            gen_clicked = st.button("\U0001f4c4 Generate BOM Document", key="bom_gen_doc")
            if gen_clicked:
                fname, fbytes = bom.generate_bom_document(sel)
                st.session_state.bom_generated_doc = {"code": sel, "filename": fname, "bytes": fbytes}
            gd = st.session_state.get("bom_generated_doc")
            if gd and gd["code"] == sel:
                st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                    file_name=gd["filename"], mime=XLSX_MIME, key=f"bom_dl_{sel}")


def _parse_geo(s):
    try:
        p = str(s or "").split(",")
        if len(p) == 2:
            return float(p[0].strip()), float(p[1].strip())
    except Exception:
        pass
    return None


def _build_transfer_map_data(transfers):
    """Aggregates get_transfer_opportunities()'s per-material rows into
    one route per (from, to) location pair, and resolves each location's
    real coordinates — same {locations, routes} shape a Leaflet map
    needs, adapted from erp_ui.py's vendor->delivery PO map. Unlike that
    map, both ends here are the same kind of node (a delivery location),
    so this uses one marker style throughout rather than two."""
    if not transfers:
        return [], []
    loc_lookup = {d["id"]: d for d in pr_consolidation.get_delivery_locations(active_only=False)}

    grouped = {}
    for t in transfers:
        key = (t["from_location"], t["to_location"])
        grouped.setdefault(key, []).append(t)

    used_ids = {lid for pair in grouped for lid in pair}
    locations = []
    for lid in used_ids:
        d = loc_lookup.get(lid)
        if not d or not d.get("geo"):
            continue
        ll = _parse_geo(d["geo"])
        if not ll:
            continue
        locations.append({"id": lid, "name": d["name"], "lat": ll[0], "lng": ll[1]})

    routes = []
    for (f, t), items in grouped.items():
        total_qty = sum(i["suggested_qty"] for i in items)
        materials = "<br>".join(f"{i['suggested_qty']:g}\u00d7 {i['mat_desc']}" for i in items[:6])
        if len(items) > 6:
            materials += f"<br>+{len(items)-6} more"
        routes.append({"from": f, "to": t, "total_qty": total_qty,
                       "materials": materials, "count": len(items)})
    return locations, routes


def _transfer_map(locations, routes):
    """Pin-and-route map for suggested inventory transfers between
    delivery locations — same Leaflet pattern as erp_ui.py's PO
    consolidation map (arced polylines, a Fit button, a legend), adapted
    for location->location flow instead of vendor->delivery flow. One
    marker style for every pin (every node here is the same kind of
    thing, unlike vendor-vs-delivery), one accent color for every route
    since "suggested transfer" is the one thing being shown, not a
    category needing its own color per instance."""
    lj = json.dumps(locations); rj = json.dumps(routes)
    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"/>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"></script>
<style>body{{margin:0}}#map{{width:100%;height:420px}}
.leg{{background:rgba(255,255,255,.95);border-radius:8px;padding:10px 14px;
border:1px solid #E2E8F0;font-size:12px;color:#374151;line-height:2}}</style></head><body>
<div id="map"></div><script>
const L_=({lj}),R=({rj});
const map=L.map("map",{{preferCanvas:true,attributionControl:false}});
L.tileLayer("https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png",{{attribution:"\u00a9 OpenStreetMap",maxZoom:18}}).addTo(map);
L_.forEach(d=>{{if(!d.lat||!d.lng)return;
L.marker([d.lat,d.lng],{{icon:L.divIcon({{html:`<div style="width:16px;height:16px;background:#0F766E;border:2px solid #fff;border-radius:50%;box-shadow:0 1px 4px rgba(0,0,0,.25)"></div>`,className:"",iconAnchor:[8,8],popupAnchor:[0,-10]}})}}).addTo(map)
.bindPopup(`<b style="color:#0F766E">${{d.name}}</b><br><small>${{d.id}}</small>`);
}});
function arc(a,b,c,d,n=50){{const p=[];for(let i=0;i<=n;i++){{const t=i/n;p.push([a+(c-a)*t+Math.sin(Math.PI*t)*Math.sqrt((c-a)**2+(d-b)**2)*.18,b+(d-b)*t]);}}return p;}}
R.forEach(r=>{{const f=L_.find(x=>x.id===r.from),t=L_.find(x=>x.id===r.to);
if(!f||!t)return;
const w=Math.min(2+Math.sqrt(r.total_qty)*0.8,8);
L.polyline(arc(f.lat,f.lng,t.lat,t.lng),{{color:"#D97706",weight:w,opacity:.75,dashArray:"6,4"}}).addTo(map)
.bindPopup(`<b>${{f.name}}</b> \u2192 <b>${{t.name}}</b><br><small>${{r.materials}}</small>`);
}});
const pts=L_.filter(d=>d.lat).map(d=>[d.lat,d.lng]);
function doFit(){{
  map.invalidateSize(true);
  if(pts.length){{map.fitBounds(pts,{{padding:[50,50],maxZoom:9}});}}
}}
setTimeout(doFit,200);
setTimeout(doFit,600);
setTimeout(doFit,1200);
let autoFitDone=false;
const ro=new ResizeObserver(entries=>{{
  const r=entries[0].contentRect;
  if(!autoFitDone && r.width>10 && r.height>10){{autoFitDone=true; doFit();}}
}});
ro.observe(document.getElementById("map"));
const fitBtn=L.control({{position:"topright"}});
fitBtn.onAdd=()=>{{const d=L.DomUtil.create("div","leaflet-bar");
d.innerHTML='<a href="#" title="Fit to all points" style="width:34px;height:30px;line-height:30px;text-align:center;display:block;font-size:11px;font-weight:600;text-decoration:none;color:#374151;background:#fff">Fit</a>';
d.onclick=(e)=>{{e.preventDefault(); doFit();}};
L.DomEvent.disableClickPropagation(d);
return d;}};
fitBtn.addTo(map);
const leg=L.control({{position:"bottomright"}});
leg.onAdd=()=>{{const d=L.DomUtil.create("div","leg");
d.innerHTML="<b style='font-size:11px;color:#6B7280;text-transform:uppercase'>Legend</b>"+
"<div><span style=\\"width:12px;height:12px;background:#0F766E;display:inline-block;border-radius:50%;margin-right:6px;vertical-align:middle\\"></span>Location</div>"+
"<div><span style=\\"display:inline-block;width:16px;height:2.5px;background:#D97706;margin-right:6px;vertical-align:middle\\"></span>Suggested transfer</div>";
return d;}};
leg.addTo(map);
</script></body></html>"""
    st.components.v1.html(html, height=460)
    st.caption("Line thickness scales with suggested transfer quantity \u00b7 "
              "click a route for the material breakdown \u00b7 **Fit** re-frames the map.")


def _material_flow_svg(vendor_label, vendor_sub, from_name, stays_qty, destinations, uom):
    """
    Live version of the Sankey-style concept sketch: Supplier(s) ->
    from_location -> {stays here, one branch per destination}.

    destinations: list of {"name": str, "qty": float} — every
    destination this material+source pair is suggested to move to,
    shown as separate outgoing branches in ONE diagram rather than one
    diagram per destination. This isn't just tidier — it's a
    correctness fix: when the same source feeds two destinations
    simultaneously, showing them independently meant each diagram
    computed "stays here" as on_hand minus only THAT destination's
    quantity, silently ignoring the other destination's claim on the
    same stock. Combined here, stays_qty is computed once by the
    caller as on_hand minus the sum of every destination, so it's
    correct regardless of how many destinations exist.

    Height is computed from the number of branches, not fixed — a
    fixed-height version previously came up 6px short of its own
    content (300 viewBox + 16px outer div padding vs. a 310px iframe),
    clipping the bottom branch. Both the SVG's own height attribute and
    the iframe height passed to st.components.v1.html() are now
    derived from the same layout math instead of guessed independently.
    """
    branches = [{"name": "Stays here", "qty": stays_qty, "is_stay": True}] + \
               [{"name": d["name"], "qty": d["qty"], "is_stay": False} for d in destinations]
    on_hand = stays_qty + sum(d["qty"] for d in destinations)
    ref = on_hand or 1
    def w(q):
        return max(2.0, min(8.0, 2.0 + (q / ref) * 6.0))

    def esc(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    BOX_H, GAP, TOP_MARGIN, BOTTOM_MARGIN = 56, 16, 20, 46
    n = len(branches)
    branches_h = n * BOX_H + (n - 1) * GAP
    viewbox_h = TOP_MARGIN + branches_h + BOTTOM_MARGIN
    src_cy = TOP_MARGIN + branches_h / 2  # vertically center the source box on the branch block

    paths, boxes = [], []
    for i, b in enumerate(branches):
        y = TOP_MARGIN + i * (BOX_H + GAP)
        cy = y + BOX_H / 2
        color = "#888780" if b["is_stay"] else "#D97706"
        opacity = "0.55" if b["is_stay"] else "0.9"
        paths.append(f'<path d="M420,{src_cy:g} C440,{src_cy:g} 440,{cy:g} 440,{cy:g}" '
                     f'fill="none" stroke="{color}" stroke-width="{w(b["qty"])}" opacity="{opacity}"/>')
        cls = "mf-gray" if b["is_stay"] else "mf-coral"
        label = "Stays here" if b["is_stay"] else esc(b["name"])
        sub = f'{b["qty"]:g} {esc(uom)}' if b["is_stay"] else f'Suggested +{b["qty"]:g} {esc(uom)}'
        boxes.append(f'<g><rect class="{cls}" x="440" y="{y:g}" width="190" height="{BOX_H}" rx="8" stroke-width="0.5"/>'
                     f'<text class="mf-t" x="535" y="{y+18:g}" text-anchor="middle" dominant-baseline="central">{label}</text>'
                     f'<text class="mf-ts" x="535" y="{y+36:g}" text-anchor="middle" dominant-baseline="central">{sub}</text></g>')

    svg = f"""<div style="padding:8px 0">
<svg width="100%" height="{viewbox_h}" viewBox="0 0 680 {viewbox_h}" role="img">
<title>Material flow from {esc(vendor_label)} through {esc(from_name)} to {len(destinations)} destination(s)</title>
<desc>{on_hand:g} {esc(uom)} on hand at {esc(from_name)}: {stays_qty:g} stays there, the rest suggested for transfer across {len(destinations)} destination(s).</desc>
<style>
.mf-t{{font:14px -apple-system,sans-serif;font-weight:500;fill:var(--text-primary,#1A1A2E)}}
.mf-ts{{font:12px -apple-system,sans-serif;fill:var(--text-secondary,#64748B)}}
.mf-gray{{fill:#F1EFE8;stroke:#B4B2A9}}
.mf-teal{{fill:#E1F5EE;stroke:#5DCAA5}}
.mf-coral{{fill:#FAECE7;stroke:#F0997B}}
</style>
{''.join(paths)}
<line x1="170" y1="{src_cy:g}" x2="230" y2="{src_cy:g}" stroke="#888780" stroke-width="{w(on_hand)}" opacity="0.55"/>

<text class="mf-ts" x="200" y="{src_cy-18:g}" text-anchor="middle">{on_hand:g} {esc(uom)}</text>

<g><rect class="mf-gray" x="20" y="{src_cy-28:g}" width="150" height="56" rx="8" stroke-width="0.5"/>
<text class="mf-t" x="95" y="{src_cy-8:g}" text-anchor="middle" dominant-baseline="central">{esc(vendor_label)}</text>
<text class="mf-ts" x="95" y="{src_cy+10:g}" text-anchor="middle" dominant-baseline="central">{esc(vendor_sub)}</text></g>

<g><rect class="mf-teal" x="230" y="{src_cy-28:g}" width="190" height="56" rx="8" stroke-width="0.5"/>
<text class="mf-t" x="325" y="{src_cy-8:g}" text-anchor="middle" dominant-baseline="central">{esc(from_name)}</text>
<text class="mf-ts" x="325" y="{src_cy+10:g}" text-anchor="middle" dominant-baseline="central">{on_hand:g} {esc(uom)} on hand today</text></g>

{''.join(boxes)}

<text class="mf-ts" x="340" y="{viewbox_h-16:g}" text-anchor="middle">Gray = received or held \u00b7 amber = suggested transfer</text>
</svg></div>"""
    st.components.v1.html(svg, height=viewbox_h + 26)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Inventory
# ══════════════════════════════════════════════════════════════════════════════
def page_inventory():
    st.markdown("## \U0001f4e6 Inventory")
    st.caption("Transaction-based — every balance below is computed live from "
               "transaction history, not stored separately. Goods Receipt, Production "
               "Confirmation, and O2C Fulfillment shipments all post here.")
    st.divider()

    s = inv.stats()
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Materials Tracked", s["materials_tracked"])
    m2.metric("Locations Tracked", s["locations_tracked"])
    m3.metric("Total Units On Hand", f"{s['total_units_on_hand']:g}")
    m4.metric("In Transit", f"{s['total_units_in_transit']:g}",
              help="Shipped but not yet received — real stock, moving between "
                   "locations, not usable at either end until Confirm Receipt.")
    m5.metric("Negative Balances", s["negative_balances"],
              delta="review these" if s["negative_balances"] else None,
              delta_color="inverse")
    st.divider()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["\U0001f4ca Stock by Location", "\U0001f50d By Material",
                                "\U0001f4dc Transaction History", "\U0001f3af Position & Transfers",
                                "\U0001f4c5 Time-Phased Planning"])

    # ── TAB 1 — full stock position ────────────────────────────────────────────
    with tab1:
        balances = inv.get_all_balances()
        in_transit = inv.get_stock_transfers(status="In Transit")
        if not balances and not in_transit:
            st.info("No inventory transactions yet — post a Goods Receipt to see stock appear here.")
        else:
            bdf = pd.DataFrame(balances)[["mat_code","mat_desc","location_name","balance"]]
            bdf.columns = ["Code","Description","Location","Balance"]
            bdf["Status"] = "On Hand"
            if in_transit:
                loc_name = {l["id"]: l["name"] for l in pr_consolidation.get_delivery_locations(active_only=False)}
                tdf = pd.DataFrame([{
                    "Code": t["material_code"], "Description": t["material_desc"],
                    "Location": loc_name.get(t["to_location"], t["to_location"]),
                    "Balance": t["quantity"], "Status": "In Transit",
                } for t in in_transit])
                bdf = pd.concat([bdf, tdf], ignore_index=True)
            open_reservations = res.get_all_reservations()
            open_reservations = [r for r in open_reservations if r["status"] == "Open"]
            if open_reservations:
                loc_name = {l["id"]: l["name"] for l in pr_consolidation.get_delivery_locations(active_only=False)}
                rdf = pd.DataFrame([{
                    "Code": r["material_code"], "Description": r["material_desc"],
                    "Location": loc_name.get(r["location_id"], r["location_id"]),
                    "Balance": r["quantity"], "Status": "Reserved",
                } for r in open_reservations])
                bdf = pd.concat([bdf, rdf], ignore_index=True)
            st.dataframe(bdf[["Code","Description","Location","Balance","Status"]],
                        use_container_width=True, hide_index=True)

            gen_clicked = st.button("\U0001f4c4 Generate Stock Report", key="inv_gen_doc")
            if gen_clicked:
                fname, fbytes = inv.generate_stock_report()
                st.session_state.inv_generated_doc = {"filename": fname, "bytes": fbytes}
            gd = st.session_state.get("inv_generated_doc")
            if gd:
                st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                    file_name=gd["filename"], mime=XLSX_MIME, key="inv_dl")

            st.divider()
            with st.expander("\U0001f512 Reservation Ledger (ATP-US-02)"):
                st.caption("Real, currently-Open reservations \u2014 quantity spoken for "
                          "against a Sales Order, decremented from Available-to-Promise "
                          "at the material/location shown, until picked (Consumed) or "
                          "released.")
                if not open_reservations:
                    st.caption("No open reservations right now.")
                else:
                    loc_name = {l["id"]: l["name"] for l in pr_consolidation.get_delivery_locations(active_only=False)}
                    for r in open_reservations:
                        c1, c2 = st.columns([4, 1.3])
                        c1.write(f"**{r['reservation_id']}** \u2014 {r['material_desc']} \u00b7 "
                                f"{r['quantity']:g} \u00b7 {loc_name.get(r['location_id'], r['location_id'])} "
                                f"\u00b7 {r['so_id']} line {r['so_line_item']} \u00b7 "
                                f"reserved {r['created_date']}")
                        if c2.button("\U0001f513 Release", key=f"release_res_{r['reservation_id']}"):
                            try:
                                res.release_reservation(r["reservation_id"],
                                    reason="Released manually from Reservation Ledger")
                                st.cache_data.clear()
                                st.success(f"\u2705 {r['reservation_id']} released \u2014 "
                                          f"{r['quantity']:g} units returned to "
                                          f"Available-to-Promise.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"\u274c {e}")

    # ── TAB 2 — drill into one material across locations ──────────────────────────
    with tab2:
        balances = inv.get_all_balances()
        in_transit_all = inv.get_stock_transfers(status="In Transit")
        materials = sorted({b["mat_code"] for b in balances} |
                           {t["material_code"] for t in in_transit_all})
        if not materials:
            st.info("No inventory transactions yet.")
        else:
            mat_desc = {b["mat_code"]: b["mat_desc"] for b in balances}
            mat_desc.update({t["material_code"]: t["material_desc"] for t in in_transit_all})
            sel_mat = st.selectbox("Material", materials,
                format_func=lambda k: f"{k} — {mat_desc[k]}", key="inv_mat_sel")
            by_loc = inv.get_material_balance_by_location(sel_mat)
            ldf = pd.DataFrame(by_loc)[["location_name","balance"]]
            ldf.columns = ["Location", "Balance"]
            sel_in_transit = [t for t in in_transit_all if t["material_code"] == sel_mat]
            if sel_in_transit:
                loc_name = {l["id"]: l["name"] for l in pr_consolidation.get_delivery_locations(active_only=False)}
                for t in sel_in_transit:
                    dest = loc_name.get(t["to_location"], t["to_location"])
                    ldf = pd.concat([ldf, pd.DataFrame([{
                        "Location": f"{dest} (In Transit)", "Balance": t["quantity"]}])],
                        ignore_index=True)
            st.dataframe(ldf, use_container_width=True, hide_index=True)
            c1, c2 = st.columns(2)
            c1.metric("Total on hand across all locations", f"{sum(b['balance'] for b in by_loc):g}")
            transit_total = sum(t["quantity"] for t in sel_in_transit)
            if transit_total:
                c2.metric("In Transit", f"{transit_total:g}",
                         help="Shipped but not yet received — not usable at either end yet.")

    # ── TAB 3 — raw transaction history ────────────────────────────────────────────
    with tab3:
        txns = inv.get_transactions()
        if not txns:
            st.info("No transactions yet.")
        else:
            tdf = pd.DataFrame(txns)[["txn_id","txn_date","mat_code","location_name",
                                       "quantity","txn_type","reference_id"]]
            tdf.columns = ["Txn","Date","Material","Location","Qty","Type","Reference"]
            st.dataframe(tdf, use_container_width=True, hide_index=True)

    # ── TAB 4 — persistent demand vs. supply position + transfer suggestions ──────
    with tab4:
        # A message set right before st.rerun() (Ship/Confirm Receipt below) would
        # otherwise never actually be seen — st.rerun() halts the current script
        # immediately and starts completely fresh, discarding anything rendered
        # (including a just-called st.success()) before the browser ever draws that
        # frame. Real Streamlit behavior, not an artifact of any testing tool — a
        # second, related bug found right after fixing the first one, from watching
        # a real success message never actually appear despite the action itself
        # having genuinely succeeded. Storing it here and showing it on the very
        # next render is the standard fix.
        flash = st.session_state.pop("tp_flash_message", None)
        if flash:
            flash_fn = {"success": st.success, "warning": st.warning}.get(flash[0], st.error)
            flash_fn(flash[1])

        st.caption("Demand here means confirmed, BOM-matched Sales Orders only — the "
                  "one real demand signal in this system right now. No forecasting, "
                  "no reorder points.")
        position = bom.get_inventory_position()
        if not position:
            st.info("No confirmed Sales Orders for BOM-matched items right now — "
                   "nothing to show a position against.")
        else:
            loc_names = {d["id"]: d["name"] for d in load_delivery_locs()}
            loc_short = {d["id"]: (d["city"] or d["name"][:14]) for d in load_delivery_locs()}
            def _lname(loc_id): return loc_names.get(loc_id, loc_id)
            def _lshort(loc_id): return loc_short.get(loc_id, loc_id)

            pdf = pd.DataFrame(position)[["mat_code","mat_desc","location_id","gross_demand",
                                          "on_hand","open_po","net_position"]]
            pdf["location_id"] = pdf["location_id"].map(_lname)
            pdf.columns = ["Code","Description","Location","Demand","On Hand","Open PO","Net Position"]
            st.dataframe(pdf, use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("##### \U0001f504 Cross-location transfer opportunities")
            transfers = bom.get_transfer_opportunities()
            if not transfers:
                st.success("\u2705 No transfer opportunities right now — either nothing's "
                          "short, or nothing short has stock sitting elsewhere.")
            else:
                map_locs, map_routes = _build_transfer_map_data(transfers)
                if map_locs:
                    _transfer_map(map_locs, map_routes)
                tdf = pd.DataFrame(transfers)[["mat_code","mat_desc","from_location","to_location",
                                               "suggested_qty","shortage_at_destination"]]
                tdf["from_location"] = tdf["from_location"].map(_lname)
                tdf["to_location"] = tdf["to_location"].map(_lname)
                tdf.columns = ["Code","Description","From","To","Suggested Qty","Shortage at Destination"]
                st.dataframe(tdf, use_container_width=True, hide_index=True)

                use_sto = od.get_default("Use Stock Transfer Orders") == "Yes"
                if use_sto:
                    st.markdown("###### \U0001f3ed Hub Allocation (Stock Transfer Order)")
                    st.caption("Creates a Stock Transfer Order covering every Plant shown "
                              "above for this material and Hub, generates an e-way bill "
                              "where required (inter-state movement above the statutory "
                              "value threshold), and posts the freight accrual. Allocation "
                              "uses the same largest-shortage-first split shown above; a "
                              "configurable approval step will be available in a future "
                              "release.")
                    sto_groups = {}
                    for t in transfers:
                        key = (t["mat_code"], t["from_location"])
                        sto_groups.setdefault(key, []).append(t)
                    for (mat_code, from_loc), group in sto_groups.items():
                        mat_desc = group[0]["mat_desc"]
                        destinations = ", ".join(_lname(g["to_location"]) for g in group)
                        gc1, gc2 = st.columns([4, 1.3])
                        gc1.write(f"**{mat_desc}** from {_lname(from_loc)} \u2192 {destinations}")
                        if gc2.button("\U0001f3ed Create STO", key=f"create_sto_{mat_code}_{from_loc}"):
                            try:
                                result = sto.simple_allocate_and_create_sto(
                                    mat_code, from_loc, created_by="ERP UI")
                                st.cache_data.clear()
                                legs = ", ".join(f"{_lname(l['to_location'])} "
                                                f"({l['allocated_qty']:g})" for l in result["lines"])
                                st.session_state["tp_flash_message"] = ("success",
                                    f"\u2705 {result['sto_id']} created — {legs}. Real freight "
                                    f"accrual posted; e-way bill generated for any inter-state "
                                    f"leg above the real value threshold.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"\u274c {e}")

                else:
                    st.markdown("###### Ship")
                    st.caption("Ships real stock — the source location's balance drops "
                              "immediately, but the destination doesn't receive it until "
                              "someone confirms receipt below. Genuinely in transit in "
                              "between, not silently teleported. The quantity is pre-filled "
                              "with the suggestion (largest shortage served first when several "
                              "destinations compete for the same source) but freely editable — "
                              "shipping itself always re-checks real, current availability, "
                              "so an edited quantity can't over-allocate the same stock twice. "
                              "Select the lines to ship together and pick one courier for all "
                              "of them — a later version can let an agent or an optimizer "
                              "choose the best courier per line instead of one shared choice "
                              "here; this is the manual baseline that unlocks.")

                    row_keys = [f"{t['mat_code']}_{t['from_location']}_{t['to_location']}"
                               for t in transfers]

                    def _select_all_ship(row_keys=row_keys):
                        val = st.session_state["select_all_ship"]
                        for rk in row_keys:
                            st.session_state[f"chk_ship_{rk}"] = val

                    sa1, sa2 = st.columns([1, 5])
                    sa1.checkbox("Select All", key="select_all_ship",
                                on_change=_select_all_ship)

                    for i, t in enumerate(transfers):
                        item_info = po_export.get_item_by_code(t["mat_code"], active_only=False)
                        uom = item_info["uom"] if item_info else "pcs"
                        # Stable, content-based key — not the loop index. A Ship action can
                        # change which opportunities exist at all (e.g. draining a source
                        # down to zero removes it from transfers entirely), which shifts
                        # every index after it. A key built from the loop position alone
                        # means the NEXT click can silently land on nothing: real bug found
                        # 2026-07-30 — shipping two rows in sequence correctly moved stock
                        # for the first, but the second click (whose index-based key no
                        # longer matched anything once the list had changed shape) never
                        # even reached ship_transfer() at all, not a validation failure,
                        # just a lost click. row_key ties every widget for this specific
                        # opportunity to what it actually is, immune to reordering.
                        row_key = row_keys[i]
                        c0, c1, c2 = st.columns([0.4, 3.9, 1.3])
                        c0.checkbox("Select", key=f"chk_ship_{row_key}",
                                   label_visibility="collapsed")
                        c1.write(f"**{t['mat_desc']}** · {_lname(t['from_location'])} \u2192 "
                                f"{_lname(t['to_location'])}")
                        c2.number_input(f"Qty ({uom})", min_value=0.0,
                            max_value=float(t["available_at_source"]), value=float(t["suggested_qty"]),
                            step=1.0, key=f"exec_qty_{row_key}", label_visibility="collapsed")

                    st.divider()
                    st.caption("Courier for selected lines — 'Let system choose' leaves each "
                              "leg open for a least-cost, route-consolidating recommendation "
                              "at Submit to Courier time, rather than committing to one "
                              "courier for the whole selection now.")
                    bc1, bc2 = st.columns([1.5, 1])
                    LEAST_COST_OPTION = "Let system choose (least cost)"
                    courier_choice = bc1.selectbox("Courier for selected lines",
                        [LEAST_COST_OPTION] + shipping.COURIERS,
                        key="batch_courier_ship", label_visibility="collapsed")
                    batch_courier = "" if courier_choice == LEAST_COST_OPTION else courier_choice
                    if bc2.button("\U0001f69a Ship Selected", type="primary", key="ship_selected_btn"):
                        selected = [(i, t) for i, t in enumerate(transfers)
                                   if st.session_state.get(f"chk_ship_{row_keys[i]}")]
                        if not selected:
                            st.warning("No lines selected — check at least one row first.")
                        else:
                            shipped, failed = [], []
                            for i, t in selected:
                                row_key = row_keys[i]
                                item_info = po_export.get_item_by_code(t["mat_code"], active_only=False)
                                uom = item_info["uom"] if item_info else "pcs"
                                qty = st.session_state.get(f"exec_qty_{row_key}", t["suggested_qty"])
                                try:
                                    result = inv.ship_transfer(t["mat_code"], t["mat_desc"],
                                        t["from_location"], t["to_location"], qty, carrier=batch_courier)
                                    shipped.append(f"{result['transfer_id']} ({qty:g} {uom})")
                                except Exception as e:
                                    failed.append(f"{t['mat_desc']}: {e}")
                                # Real bug found and fixed here: assigning False to an
                                # already-instantiated checkbox's own session_state key
                                # (st.session_state[key] = False) raises — Streamlit
                                # disallows modifying a widget's value in the same run
                                # it was already rendered in. Deleting the key instead
                                # is allowed and has the same practical effect: on the
                                # next render (this whole block already ends in
                                # st.rerun()) the checkbox reverts to its unchecked
                                # default rather than the previously-set True.
                                if f"chk_ship_{row_key}" in st.session_state:
                                    del st.session_state[f"chk_ship_{row_key}"]
                            st.cache_data.clear()
                            msg_parts = []
                            if shipped:
                                courier_label = batch_courier if batch_courier else \
                                    "no pinned courier — least-cost proposal at Submit to Courier"
                                msg_parts.append(f"\u2705 Shipped, {courier_label}: " + ", ".join(shipped))
                            if failed:
                                msg_parts.append(f"\u274c Failed: " + "; ".join(failed))
                            st.session_state["tp_flash_message"] = (
                                "success" if not failed else "warning", "\n\n".join(msg_parts))
                            st.rerun()


                    st.divider()
                    st.markdown("##### \U0001f9ee Drill into one material's flow")
                    st.caption("Every destination competing for the same material from the "
                              "same source shows together in one diagram, not one diagram "
                              "each — that's also what makes 'stays here' correct when more "
                              "than one destination is drawing on the same stock.")

                    groups = {}
                    for t in transfers:
                        groups.setdefault((t["mat_code"], t["from_location"]), []).append(t)
                    group_keys = list(groups.keys())
                    drill_labels = {}
                    for i, key in enumerate(group_keys):
                        group = groups[key]
                        dest_summary = ", ".join(f"{_lname(g['to_location'])} ({g['suggested_qty']:g})"
                                                 for g in group)
                        drill_labels[i] = f"{group[0]['mat_desc']} — {_lname(key[1])} \u2192 {dest_summary}"
                    sel_i = st.selectbox("Material + source to trace", list(drill_labels.keys()),
                        format_func=lambda k: drill_labels[k], key="inv_drill_sel")
                    mat_code, from_loc = group_keys[sel_i]
                    group = groups[(mat_code, from_loc)]

                    detail = bom.get_material_flow_detail(mat_code, from_loc)
                    total_suggested = sum(g["suggested_qty"] for g in group)
                    stays_qty = max(0.0, detail["on_hand"] - total_suggested)
                    if detail["vendors"]:
                        vendor_label = detail["vendors"][0][0]
                        vendor_sub = ("Supplier" if len(detail["vendors"]) == 1
                                      else f"+ {len(detail['vendors'])-1} more supplier(s)")
                    else:
                        vendor_label, vendor_sub = "Unknown source", "No GR on file"
                    item_info = po_export.get_item_by_code(mat_code, active_only=False)
                    vendor_short = vendor_label if len(vendor_label) <= 15 else vendor_label[:14] + "\u2026"
                    from_full = _lname(from_loc)
                    from_short = _lshort(from_loc)
                    from_short = from_short if len(from_short) <= 20 else from_short[:19] + "\u2026"

                    destinations, dest_captions = [], []
                    for g in group:
                        to_full = _lname(g["to_location"])
                        to_short = _lshort(g["to_location"])
                        to_short = to_short if len(to_short) <= 20 else to_short[:19] + "\u2026"
                        destinations.append({"name": to_short, "qty": g["suggested_qty"]})
                        if to_short != to_full:
                            dest_captions.append(f"{to_short} = {to_full}")

                    _material_flow_svg(vendor_short, vendor_sub, from_short, stays_qty, destinations,
                                       item_info["uom"] if item_info else "pcs")
                    captions = []
                    if vendor_short != vendor_label:
                        captions.append(f"Supplier: {vendor_label}")
                    if from_short != from_full:
                        captions.append(f"{from_short} = {from_full}")
                    captions.extend(dest_captions)
                    if captions:
                        st.caption("  \u00b7  ".join(captions))

            in_transit = inv.get_stock_transfers(status="In Transit")
            if in_transit:
                st.divider()
                st.markdown("##### \U0001f69a In Transit — confirm receipt")
                st.caption("Shipped but not yet arrived — this stock isn't counted at "
                          "either location until confirmed here. Shown regardless of "
                          "whether any new transfer opportunities exist right now — a "
                          "shipment already on its way doesn't disappear just because "
                          "everything else has been resolved. Select the lines to act "
                          "on, then either submit them to the courier or confirm "
                          "they've genuinely arrived.")

                @st.dialog("\U0001f4e6 Tracking Status (Simulated)")
                def _show_tracking(transfer_id):
                    status = shipping.get_tracking_status(transfer_id)
                    if not status:
                        st.write("No tracking information on file yet.")
                        return
                    st.caption("No real courier tracking API is called here — this is a "
                              "deterministic simulation from the real ship date, not a "
                              "live status.")
                    st.markdown(f"**AWB {status['awb_number']}** \u2014 "
                               f"*{status['current_status']}*")
                    st.divider()
                    for cp in status["checkpoints"]:
                        icon = "\u2705" if cp["completed"] else "\u26aa"
                        weight = "**" if cp["completed"] else ""
                        st.markdown(f"{icon} {weight}{cp['label']}{weight} \u2014 "
                                   f"{cp['location']} \u00b7 {cp['date']}")
                    st.divider()
                    if status["current_status"] != "Delivered":
                        st.caption("Demo control \u2014 advances this simulation to its next "
                                  "real checkpoint rather than waiting for real elapsed time.")
                        if st.button("\u23e9 Skip Ahead", key=f"skip_ahead_{transfer_id}"):
                            shipping.skip_ahead_tracking(transfer_id)
                            st.rerun()

                def _select_all_receive(transfer_ids=[t["transfer_id"] for t in in_transit]):
                    val = st.session_state["select_all_receive"]
                    for tid in transfer_ids:
                        st.session_state[f"chk_receive_{tid}"] = val

                ra1, ra2 = st.columns([1, 5])
                ra1.checkbox("Select All", key="select_all_receive",
                            on_change=_select_all_receive)

                for t in in_transit:
                    c0, c1, c3, c4 = st.columns([0.4, 3.3, 1.1, 0.9])
                    c0.checkbox("Select", key=f"chk_receive_{t['transfer_id']}",
                               label_visibility="collapsed")
                    c1.write(f"**{t['transfer_id']}** — {t['material_desc']} · "
                            f"{t['quantity']:g} {t['uom']} · "
                            f"{_lname(t['from_location'])} \u2192 {_lname(t['to_location'])} "
                            f"· shipped {t['shipped_date']}"
                            + (f" via {t['carrier']}" if t['carrier'] else "")
                            + (f" · {t['source_doc']}" if t['source_type'] == 'STO' and t['source_doc']
                               else " · Ad Hoc")
                            + (f" · E-Way Bill {t['eway_bill_number']} (valid to "
                               f"{t['eway_bill_valid_until']})" if t['eway_bill_number'] else ""))
                    if t["carrier"]:
                        shipment = shipping.build_shipment_details(t)
                        fname, xlsx_bytes = shipping.generate_shipping_excel(shipment)
                        c3.download_button(f"\u2b07 {t['carrier']} Details", data=xlsx_bytes,
                            file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"shipdoc_{t['transfer_id']}")
                        if t["tracking_ref"]:
                            # AWB as a clickable link, not static text -- clicking opens
                            # the simulated tracking dialog for this specific shipment.
                            if st.button(f"\U0001f517 AWB {t['tracking_ref']}",
                                        key=f"awb_link_{t['transfer_id']}"):
                                _show_tracking(t["transfer_id"])
                    if c4.button("\u274c Cancel", key=f"cancel_{t['transfer_id']}"):
                        try:
                            inv.cancel_transfer(t["transfer_id"], cancelled_by="ERP UI")
                            st.session_state.pop(f"courier_issue_{t['transfer_id']}", None)
                            st.cache_data.clear()
                            st.session_state["tp_flash_message"] = ("success",
                                f"\u2705 {t['transfer_id']} cancelled \u2014 stock restored at "
                                f"{_lname(t['from_location'])}, visible again as a transfer "
                                f"opportunity.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"\u274c {e}")

                    # Real gap found from direct testing, not anticipated up front: a
                    # transfer too heavy for every available courier used to fail only
                    # at Submit Selected to Courier time, and the resulting warning
                    # rendered at the very top of this whole tab -- easy to miss after
                    # scrolling down to click a button in this section, and looked like
                    # a silent failure even though a real reason was always generated.
                    # This shows the same real reason inline, next to the row it's
                    # actually about, computed proactively (own weight vs. every
                    # available courier's own max) even before a submit is ever
                    # attempted, and overwritten with the more specific real reason
                    # (which can depend on what else it gets grouped with) after an
                    # actual attempt.
                    if not t["tracking_ref"]:
                        issue = st.session_state.get(f"courier_issue_{t['transfer_id']}")
                        if issue is None:
                            item_info = po_export.get_item_by_code(t["material_code"], active_only=False)
                            weight_per_unit = item_info.get("weight_kg") if item_info else None
                            if weight_per_unit is not None:
                                own_weight = weight_per_unit * t["quantity"]
                                max_limit = max(shipping.RATE_CARDS[c]["max_weight_kg"]
                                                for c in shipping.COURIERS)
                                if own_weight > max_limit:
                                    issue = (f"{own_weight:g}kg exceeds every available "
                                            f"courier's maximum ({max_limit:g}kg) on its own.")
                        if issue:
                            st.warning(f"\u26a0\ufe0f {issue}", icon="\u26a0\ufe0f")
                st.divider()
                bcol1, bcol2 = st.columns(2)
                if bcol1.button("\U0001f50c Submit Selected to Courier", key="submit_selected_btn"):
                    selected_ids = [t["transfer_id"] for t in in_transit
                                    if st.session_state.get(f"chk_receive_{t['transfer_id']}")]
                    if not selected_ids:
                        st.warning("No lines selected — check at least one row first.")
                    else:
                        by_id = {t["transfer_id"]: t for t in in_transit}
                        # A transfer with no carrier on file is no longer skipped as an
                        # error case — it's the real "let system choose" state (see the
                        # Ship section's own courier dropdown), and submit_batch_to_courier()
                        # resolves it via real least-cost, route-consolidating logic. Only
                        # an already-booked transfer is genuinely skippable here.
                        bookable_ids = [tid for tid in selected_ids if not by_id[tid]["tracking_ref"]]
                        already_booked = [tid for tid in selected_ids if by_id[tid]["tracking_ref"]]

                        batch_result = shipping.submit_batch_to_courier(bookable_ids) if bookable_ids \
                            else {"bookings": [], "unbookable": []}

                        # Real reasons from an actual attempt are more specific than the
                        # proactive per-row check (they can reflect what a leg got grouped
                        # with), so they overwrite it here. A leg that just booked
                        # successfully has its old issue cleared -- nothing left to warn about.
                        for u in batch_result["unbookable"]:
                            st.session_state[f"courier_issue_{u['transfer_id']}"] = u["reason"]
                        for b in batch_result["bookings"]:
                            for l in b["consolidated_legs"]:
                                st.session_state.pop(f"courier_issue_{l['transfer_id']}", None)

                        for tid in selected_ids:
                            if f"chk_receive_{tid}" in st.session_state:
                                del st.session_state[f"chk_receive_{tid}"]
                        st.cache_data.clear()

                        msg_parts = []
                        bookings = batch_result["bookings"]
                        if bookings:
                            n_legs = sum(len(b["consolidated_legs"]) for b in bookings)
                            consolidation_note = (f" ({n_legs} shipments \u2192 {len(bookings)} "
                                                  f"real courier booking{'s' if len(bookings) != 1 else ''}"
                                                  f"{', consolidated' if n_legs > len(bookings) else ''})")
                            lines = []
                            for b in bookings:
                                leg_ids = ", ".join(l["transfer_id"] for l in b["consolidated_legs"])
                                lines.append(f"{b['courier']} AWB {b['awb_number']} "
                                            f"(\u20b9{b['total_cost']:,.2f}, {b['total_weight_kg']:g}kg): "
                                            f"{leg_ids}")
                            msg_parts.append(f"\u2705 Booked (simulated){consolidation_note}:\n" +
                                            "\n".join(f"  \u2022 {l}" for l in lines))
                        failed = batch_result["unbookable"]
                        if failed:
                            reasons = "; ".join(f"{u['transfer_id']}: {u['reason']}" for u in failed)
                            msg_parts.append(f"\u274c Could not book: {reasons}")
                        if already_booked:
                            msg_parts.append("\u2139\ufe0f Already booked, skipped: " +
                                            ", ".join(already_booked))
                        st.session_state["tp_flash_message"] = (
                            "success" if not failed else "warning", "\n\n".join(msg_parts))
                        st.rerun()
                if bcol2.button("\u2705 Confirm Receipt Selected", type="primary",
                            key="receive_selected_btn"):
                    selected_ids = [t["transfer_id"] for t in in_transit
                                    if st.session_state.get(f"chk_receive_{t['transfer_id']}")]
                    if not selected_ids:
                        st.warning("No lines selected — check at least one row first.")
                    else:
                        received, failed = [], []
                        by_id = {t["transfer_id"]: t for t in in_transit}
                        for tid in selected_ids:
                            try:
                                inv.receive_transfer(tid)
                                received.append(f"{tid} at {_lname(by_id[tid]['to_location'])}")
                            except Exception as e:
                                failed.append(f"{tid}: {e}")
                            if f"chk_receive_{tid}" in st.session_state:
                                del st.session_state[f"chk_receive_{tid}"]
                        st.cache_data.clear()
                        msg_parts = []
                        if received:
                            msg_parts.append("\u2705 Received: " + ", ".join(received))
                        if failed:
                            msg_parts.append("\u274c Failed: " + "; ".join(failed))
                        st.session_state["tp_flash_message"] = (
                            "success" if not failed else "warning", "\n\n".join(msg_parts))
                        st.rerun()

        st.divider()
        with st.expander("\U0001f4cb Stock Transfer Orders"):
            # Real bug fixed here: this section runs unconditionally (regardless of
            # whether the transfer-opportunities block above ever ran), but that
            # block's own _lname() is only defined inside its own "if position:"
            # branch -- using it here crashed with a real scoping error whenever
            # position was empty (e.g. right after a fresh reset, before any
            # confirmed demand exists). A fresh, independent lookup here, the
            # same pattern tab5 already uses for its own independent _lname().
            _loc_names_sto = {d["id"]: d["name"] for d in load_delivery_locs()}
            def _lname_sto(loc_id): return _loc_names_sto.get(loc_id, loc_id)
            stos = sto.get_all_stos()
            if not stos:
                st.caption("No Stock Transfer Orders created yet — use \"Create STO\" above, "
                          "on any material/Hub group with more than one requesting Plant.")
            else:
                for s in stos:
                    detail = sto.get_sto(s["sto_id"])
                    st.markdown(f"**{s['sto_id']}** — {s['material_desc']} from "
                               f"{_lname_sto(s['hub_location'])} · {s['total_qty']:g} total · "
                               f"{s['allocation_rule']} · created {s['created_date']} "
                               f"by {s['created_by'] or 'unspecified'}")
                    for l in detail["lines"]:
                        gap_note = "" if l["allocated_qty"] >= l["requested_qty"] else \
                            f" \u26a0\ufe0f {l['requested_qty'] - l['allocated_qty']:g} short of requested"
                        st.caption(f"\u2003\u2192 {_lname_sto(l['to_location'])}: "
                                  f"{l['allocated_qty']:g} allocated (requested "
                                  f"{l['requested_qty']:g}){gap_note} · "
                                  f"{l['transfer_id'] or 'not shipped'}")

    # ── TAB 5 — time-phased planning ─────────────────────────────────────────
    with tab5:
        _loc_names5 = {d["id"]: d["name"] for d in load_delivery_locs()}
        def _lname(loc_id): return _loc_names5.get(loc_id, loc_id)

        st.caption("A forward projection, not a snapshot — nets real on-hand against every "
                  "known future event on its own actual date, over a 90-day horizon.")

        mode_options = ["Sales Order Based", "Optimize Existing PRs", "Reorder Qty Based"]
        current_mode = od.get_default("Time-Phased Planning Mode")
        sel_mode = st.selectbox("Demand signal", mode_options,
            index=mode_options.index(current_mode) if current_mode in mode_options else 0,
            key="tp_mode_select",
            help="Sales Order Based (default) — real confirmed orders, the most trustworthy "
                 "signal where O2C is in use. Optimize Existing PRs — no new demand injected "
                 "(would cancel against that PR's own later PO arrival); instead checks "
                 "whether each Open PR gives itself enough lead time, and flags cross-"
                 "location duplicates. Reorder Qty Based — manual min/max + assumed "
                 "reorder cadence per material/location, for a site with no Sales Order "
                 "history yet; every figure it produces is an assumption, not a confirmed "
                 "order, and is labelled as such below.")
        if sel_mode != current_mode:
            od.set_org_default("Time-Phased Planning Mode", sel_mode)
            st.cache_data.clear()
            st.rerun()

        if sel_mode == "Optimize Existing PRs":
            st.info("\u2139\ufe0f This mode doesn't add new projected demand — an Open PR's "
                   "own need would roughly cancel against that same PR's later arrival, "
                   "telling a trajectory nothing real. Instead it checks whether every "
                   "currently Open PR gives itself enough lead time, and flags material+"
                   "location combinations with more than one competing PR.")
        elif sel_mode == "Reorder Qty Based":
            st.warning("\u26a0\ufe0f Every figure below this point is an **assumption** — a "
                      "manually declared min/max and reorder cadence, not a confirmed order. "
                      "Treat accordingly.")
        st.divider()

        if sel_mode == "Optimize Existing PRs":
            analysis = bom.get_pr_optimization_analysis()
            insufficient = [r for r in analysis["sufficiency"] if r["status"] == "Insufficient Lead Time"]
            sufficient = [r for r in analysis["sufficiency"] if r["status"] == "Sufficient"]
            m1, m2, m3 = st.columns(3)
            m1.metric("\U0001f6a8 Insufficient Lead Time", len(insufficient))
            m2.metric("\u2705 Sufficient", len(sufficient))
            m3.metric("\u26a0\ufe0f Possible Duplicates", len(analysis["duplicates"]))
            st.divider()

            if insufficient:
                st.markdown("##### \U0001f6a8 Insufficient Lead Time")
                st.caption("If consolidated into a PO today, these would still arrive after "
                          "the PR's own stated required date.")
                idf = pd.DataFrame(insufficient)[["pr_number", "mat_desc", "location",
                                                   "required_date", "implied_arrival_date",
                                                   "shortfall_days"]]
                idf["location"] = idf["location"].map(_lname)
                idf.columns = ["PR", "Material", "Location", "Required By", "Implied Arrival",
                              "Shortfall (days)"]
                st.dataframe(idf, use_container_width=True, hide_index=True)
                st.divider()

            if analysis["duplicates"]:
                st.markdown("##### \u26a0\ufe0f Possible Duplicates (same material + location)")
                st.caption("Warning only, never a block — two independently-raised PRs for "
                          "the same material and location can both be legitimate.")
                for d in analysis["duplicates"]:
                    lines_desc = ", ".join(f"{c['pr_number']} ({c['qty']:g}, due {c['required_date']})"
                                           for c in d["competing_lines"])
                    st.write(f"**{d['mat_code']}** @ {_lname(d['location'])}: {lines_desc}")
                st.divider()

            if not insufficient and not analysis["duplicates"]:
                st.success("\u2705 Every Open PR gives itself enough lead time, and no "
                          "material+location combination has competing PRs.")

        elif sel_mode == "Reorder Qty Based":
            st.markdown("##### \u2699\ufe0f Configure min/max + reorder cadence")
            with st.form("tp_rate_config"):
                c1, c2, c3, c4 = st.columns(4)
                mat_input = c1.text_input("Material Code")
                loc_input = c2.text_input("Location ID")
                min_input = c3.number_input("Min Qty", min_value=0.0, step=1.0)
                max_input = c4.number_input("Max Qty", min_value=0.0, step=1.0)
                cadence_input = st.number_input("Typical reorder cadence (days)", min_value=1, step=1, value=30)
                if st.form_submit_button("Save"):
                    if mat_input and loc_input and max_input > min_input:
                        bom.set_planning_params(mat_input, loc_input, min_input, max_input, cadence_input)
                        st.cache_data.clear()
                        st.success(f"\u2705 Saved — implied rate: {(max_input-min_input)/cadence_input:.1f} "
                                  f"units/day.")
                    else:
                        st.error("\u274c Material, Location required, and Max must exceed Min.")
            existing = bom.get_planning_params()
            if existing:
                edf = pd.DataFrame(existing)
                edf.columns = ["Material", "Location", "Min", "Max", "Cadence (days)"]
                st.dataframe(edf, use_container_width=True, hide_index=True)
            st.divider()

        recs = bom.get_procurement_recommendations()
        at_risk = [r for r in recs if r["outcome"] == "At Risk"]
        action_needed = [r for r in recs if r["outcome"] == "Action Needed"]
        covered = [r for r in recs if r["outcome"] == "Already Covered by Existing PR/PO"]

        m1, m2, m3 = st.columns(3)
        m1.metric("\U0001f6a8 At Risk", len(at_risk))
        m2.metric("\U0001f4cb Action Needed", len(action_needed))
        m3.metric("\u2705 Already Covered", len(covered))
        st.divider()

        if at_risk:
            st.markdown("##### \U0001f6a8 At Risk — needs a human decision, not a PR")
            st.caption("Normal procurement genuinely can't beat these deadlines. A PR would "
                      "arrive too late regardless — this needs expediting, splitting an "
                      "order, an alternate vendor, or a customer conversation, not a "
                      "system-generated document.")
            ardf = pd.DataFrame(at_risk)[["mat_desc", "location", "remaining_gap",
                                          "stockout_date", "days_until_stockout",
                                          "pipeline_lead_time_days"]]
            ardf["location"] = ardf["location"].map(_lname)
            ardf.columns = ["Material", "Location", "Gap", "Stock-Out Date",
                            "Days Left", "Normal Lead Time (days)"]
            st.dataframe(ardf, use_container_width=True, hide_index=True)
            st.divider()

        if action_needed:
            st.markdown("##### \U0001f4cb Action Needed — timed requisitions")
            st.caption("Enough lead-time headroom for normal procurement, sized and dated "
                      "to arrive right before the projected stock-out.")
            for i, r in enumerate(action_needed):
                c1, c2 = st.columns([5, 1])
                c1.write(f"**{r['mat_desc']}** \u00b7 {_lname(r['location'])} \u00b7 "
                        f"need **{r['recommended_qty']:g}** by **{r['required_by_date']}** "
                        f"({r['days_until_stockout']}d out, {r['pipeline_lead_time_days']}d "
                        f"lead time via {r['lead_time_source'].replace('_', ' ')})")
                if c2.button("Create PR", key=f"tp_create_pr_{i}"):
                    try:
                        pr_id = f"PR-{date.today().strftime('%Y%m%d')}-{100+i}"
                        pr_consolidation.create_pr(pr_id, requester_id="TIME-PHASED-PLANNING",
                            requester_name="Time-Phased Planning", requester_dept="Planning",
                            project_id="AUTO-PLANNING",
                            lines=[{"vendor": "", "mat_code": r["mat_code"],
                                    "mat_desc": r["mat_desc"], "uom": "pcs",
                                    "qty": r["recommended_qty"], "req_date": r["required_by_date"],
                                    "deliv_loc": r["location"], "deliv_geo": ""}])
                        st.cache_data.clear()
                        st.success(f"\u2705 {pr_id} created \u2014 required by "
                                  f"{r['required_by_date']}.")
                    except Exception as e:
                        st.error(f"\u274c {e}")
            st.divider()

        if covered:
            st.markdown("##### \u2705 Already Covered by an Existing PR/PO")
            st.caption("A real, in-progress procurement line already addresses this exact "
                      "gap \u2014 open, in RFP, or already turned into a PO \u2014 so no duplicate "
                      "gets proposed.")
            cdf = pd.DataFrame(covered)[["mat_desc", "location", "remaining_gap",
                                         "covering_pr", "stockout_date"]]
            cdf.columns = ["Material", "Location", "Gap", "Covering PR", "Needed By"]
            cdf["Location"] = cdf["Location"].map(_lname)
            st.dataframe(cdf, use_container_width=True, hide_index=True)
            st.divider()

        if not recs:
            st.success("\u2705 Nothing at risk, nothing needing a new requisition \u2014 every "
                      "current shortfall is either fully covered by a transfer or already "
                      "in hand.")

        st.markdown("##### \U0001f4c8 Drill into one material's projected position")
        all_positions = bom.project_all_positions()
        if not all_positions:
            st.info("No confirmed demand inside the planning horizon yet \u2014 nothing to project.")
        else:
            labels = {}
            for i, position in enumerate(all_positions):
                status = (
                    f" ⚠️ stock-out {position['stockout_date']}"
                    if position["stockout_date"]
                    else " ✅ covered"
                )
                labels[i] = (
                    f"{position['mat_code']} — {_lname(position['location'])}{status}"
                )
            sel = st.selectbox("Material + location", list(labels.keys()),
                format_func=lambda k: labels[k], key="tp_drill_sel")
            proj = all_positions[sel]
            tdf = pd.DataFrame(proj["trajectory"])[["date", "balance"]].set_index("date")
            st.line_chart(tdf)
            c1, c2, c3 = st.columns(3)
            c1.metric("Starting Balance", f"{proj['starting_balance']:g}")
            c2.metric("Lowest Projected Balance", f"{proj['min_projected_balance']:g}")
            c3.metric("Projected Stock-Out", proj["stockout_date"] or "None \u2014 covered")
            with st.expander("Event-by-event detail"):
                edf = pd.DataFrame(proj["trajectory"])
                edf.columns = ["Date", "Balance", "Event"]
                st.dataframe(edf, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Production
# ══════════════════════════════════════════════════════════════════════════════
def page_production():
    st.markdown("## \U0001f3ed Production")
    st.caption("Confirms a build actually happened — the event that was missing until "
               "now. Consumes every BOM component from inventory at the chosen location "
               "and produces the finished good there. Record-only, same as everywhere "
               "else here: a shortage is shown before you confirm, but doesn't block it "
               "— if you confirm anyway, the balance goes negative honestly rather than "
               "silently refusing.")
    st.divider()

    s = prod.stats()
    m1, m2 = st.columns(2)
    m1.metric("Confirmations Recorded", s["total_confirmations"])
    m2.metric("Total Units Built", f"{s['total_units_built']:g}")
    st.divider()

    tab1, tab2 = st.tabs(["\u2795 Confirm Production", "\U0001f4cb History"])

    # ── TAB 1 — confirm a build ────────────────────────────────────────────────────
    with tab1:
        fgs = bom.get_finished_goods()
        if not fgs:
            st.info("No BOMs defined yet.")
        else:
            labels = {f["code"]: f"{f['code']} — {f['desc']}" for f in fgs}
            sel_fg = st.selectbox("Finished Good / Assembly to build", list(labels.keys()),
                                  format_func=lambda k: labels[k], key="prod_fg_sel")
            qty = st.number_input("Quantity built", min_value=1, value=1, key="prod_qty")

            deliv_locs = load_delivery_locs()
            if deliv_locs:
                loc_labels = {d["id"]: d["name"] for d in deliv_locs}
                sel_loc_id = st.selectbox("Location", list(loc_labels.keys()),
                    format_func=lambda k: loc_labels[k], key="prod_loc")
                # Store the ID, not the display name — see the identical comment on
                # the BOM Explosion page's picker above. This one matters even more
                # here: it directly drives both the on-hand preview below AND the
                # actual inventory consumption/output postings on confirm.
                location = sel_loc_id
            else:
                st.warning("No delivery locations set up — add one on the S2P app first.")
                location = ""

            if location:
                preview = prod.preview_production(sel_fg, qty, location)
                shortages = [p for p in preview if p["shortage"]]
                st.markdown(f"##### Component impact — {len(preview)} line(s)")
                pdf = pd.DataFrame(preview)[["mat_code","mat_desc","gross_qty","on_hand_qty","after_qty"]]
                pdf.columns = ["Code","Description","Need","On Hand","After Build"]
                st.dataframe(pdf, use_container_width=True, hide_index=True)
                if shortages:
                    st.warning(f"\u26a0\ufe0f {len(shortages)} component(s) will go negative — "
                              "not enough on hand. Confirming will still record it honestly; "
                              "this is your chance to check before you do.")
                else:
                    st.success("\u2705 Enough stock for every component.")

            confirmed_by = st.text_input("Confirmed By", key="prod_by")
            notes = st.text_input("Notes", key="prod_notes")

            confirm_clicked = st.button("\U0001f3ed Confirm Production", type="primary",
                                        key="prod_confirm", disabled=not location)
            if confirm_clicked:
                try:
                    result = prod.confirm_production(sel_fg, qty, location, confirmed_by, notes)
                    st.cache_data.clear()
                    st.success(f"\u2705 {result['confirmation_id']} recorded — "
                              f"{result['components_consumed']} component line(s) consumed, "
                              f"{qty:g} x {sel_fg} produced at {location}.")
                    st.rerun()
                except Exception:
                    st.error("\u274c Error confirming production:")
                    st.code(traceback.format_exc())

    # ── TAB 2 — history ───────────────────────────────────────────────────────────
    with tab2:
        confs = prod.get_confirmations()
        if not confs:
            st.info("No production confirmations yet.")
        else:
            cdf = pd.DataFrame(confs)[["confirmation_id","parent_code","quantity",
                                       "location_id","confirmation_date","confirmed_by"]]
            cdf.columns = ["Confirmation","Item","Qty Built","Location","Date","Confirmed By"]
            st.dataframe(cdf, use_container_width=True, hide_index=True)

            labels = {c["confirmation_id"]: f"{c['confirmation_id']} — {c['quantity']:g} x {c['parent_code']}"
                      for c in confs}
            sel = st.selectbox("View detail", list(labels.keys()), format_func=lambda k: labels[k], key="prod_view_sel")
            detail = prod.get_confirmation_detail(sel)
            ddf = pd.DataFrame(detail)[["mat_code","mat_desc","quantity","txn_type"]]
            ddf.columns = ["Code","Description","Qty (signed)","Movement"]
            st.dataframe(ddf, use_container_width=True, hide_index=True)

            gen_clicked = st.button("\U0001f4c4 Generate Production Slip", key="prod_gen_doc")
            if gen_clicked:
                fname, fbytes = prod.generate_production_slip(sel)
                st.session_state.prod_generated_doc = {"id": sel, "filename": fname, "bytes": fbytes}
            gd = st.session_state.get("prod_generated_doc")
            if gd and gd["id"] == sel:
                st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                    file_name=gd["filename"], mime=XLSX_MIME, key=f"prod_dl_{sel}")


# ── Router ─────────────────────────────────────────────────────────────────────
if "Goods Receipt" in page:
    page_goods_receipt()
elif "Quality Inspection" in page:
    page_quality_inspection()
elif "BOM" in page:
    page_bom()
elif "Production" in page:
    page_production()
elif "Inventory" in page:
    page_inventory()
