"""
o2c_ui.py — Order-to-Cash, as its own Streamlit app.

Deliberately separate from erp_ui.py (the Source-to-Pay app), not a new
page bolted onto it. Two reasons:
  1. erp_ui.py is demo-stable and investor-facing — O2C is still being
     actively built, and shouldn't risk destabilizing it.
  2. A single 2000+ line file with 9 pages sharing one Streamlit session
     namespace is real complexity; splitting by business process (S2P vs
     O2C) keeps each app small enough to actually reason about.

What's shared: the backend modules (customer_onboarding.py, quotation.py,
vendor_onboarding.py for its GSTIN/PAN validators, po_export.py for the
item catalog) are pure Python with no Streamlit coupling — both apps
import them directly, so business logic and the underlying data.xlsx
stay single-source-of-truth. What's NOT shared: UI code. This file has
its own small item-picker widget rather than importing erp_ui.py's —
importing a Streamlit UI module from another Streamlit app executes that
module's top-level script code (its own page config, sidebar, router),
which is exactly the kind of cross-app coupling this split exists to avoid.

Run with: streamlit run o2c_ui.py
"""

import streamlit as st
import pandas as pd
import openpyxl
import os, sys, io, json, traceback
from datetime import date, datetime, timedelta

import po_export
import db
import pr_consolidation
import seed_manager as sm
import vendor_invoices as vi
import customer_onboarding as co
import quotation as qt
import vendor_onboarding as vo
import sales_order as so
import backorder as bo
import fulfillment as ful
import org_profile as op
import demo_profiles as dp
import org_defaults as od
import item_tax as it
import billing as bl
import accounting as acct
import cash_application as ca

_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(_DIR, "data.xlsx")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="O2C — Order to Cash", page_icon="\U0001f6d2", layout="wide")

from ui_theme import apply_theme
apply_theme()


# ── Shared item catalog (mirrors erp_ui.py's loader, same underlying data) ──────
@st.cache_data(ttl=300)
def load_catalog():
    return po_export.load_item_master(DATA_FILE)


def load_delivery_locs():
    """Delivery_Locations now lives in SQLite — delegates to
    pr_consolidation.py's canonical reader (also used by erp_ui.py and
    mfg_ui.py) instead of this file's own separate Excel read. That
    separate read never filtered on Active despite the column existing
    — closed now that there's exactly one implementation."""
    try:
        return pr_consolidation.get_delivery_locations(active_only=True)
    except Exception:
        return []


def _loc_display(loc_id):
    """Sales_Orders.delivery_location stores a location_id (see the
    pickers above) — this turns it back into a human-readable name for
    display. Falls back to the raw value if it doesn't match any known
    location (e.g. blank, or older data saved before this was fixed to
    store IDs instead of names)."""
    if not loc_id:
        return None
    for d in pr_consolidation.get_delivery_locations(active_only=False):
        if d["id"] == loc_id:
            return d["name"]
    return loc_id


# ── A small, self-contained item picker — simpler than erp_ui.py's on purpose ──
def item_picker(key_prefix):
    """
    Search Item Master, tick items, add to a flat session_state list at
    key_prefix + '_lines'. Deliberately flatter than the S2P app's picker
    (no separate staged/selected indirection) — fewer moving parts, easier
    to reason about, given what we just went through debugging that one.
    Returns the current list of {code, desc, uom, price, qty} dicts.
    """
    lines_key = f"{key_prefix}_lines"
    if lines_key not in st.session_state:
        st.session_state[lines_key] = []

    items = load_catalog()
    c1, c2 = st.columns([4, 1])
    with c1:
        query = st.text_input("Search items", placeholder="Type a name or code…",
                              key=f"{key_prefix}_q", label_visibility="collapsed")
    with c2:
        search_clicked = st.button("\U0001f50d Search", key=f"{key_prefix}_search",
                                   use_container_width=True)

    if search_clicked and len(query) >= 2:
        st.session_state[f"{key_prefix}_results"] = po_export.fuzzy_search(query, items, max_results=15)
    results = st.session_state.get(f"{key_prefix}_results", [])

    existing_codes = {ln["code"] for ln in st.session_state[lines_key]}
    if results:
        for item in results:
            already = item["code"] in existing_codes
            c1, c2, c3 = st.columns([5, 2, 1])
            c1.markdown(f"**{item['desc']}**  \n<small>{item['code']} · {item['uom']} · "
                       f"\u20b9{item['price']:,.2f}</small>", unsafe_allow_html=True)
            c2.write("\u2705 added" if already else "")
            if not already:
                if c3.button("Add", key=f"{key_prefix}_add_{item['code']}"):
                    st.session_state[lines_key].append({
                        "code": item["code"], "desc": item["desc"],
                        "uom": item["uom"], "price": item["price"], "qty": 1,
                    })
                    st.rerun()

    return st.session_state[lines_key]


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### \U0001f6d2 ERP Suite")
    st.caption("Order-to-Cash")
    st.divider()
    page = st.radio("", ["\U0001f465  Customer Onboarding",
                         "\U0001f4b0  Quotation",
                         "\U0001f4e6  Sales Orders",
                         "\U0001f69a  Fulfillment",
                         "\U0001f9fe  Billing & Invoicing",
                         "\U0001f4b5  Cash Application",
                         "\U0001f4d2  Accounting",
                         "\u2699\ufe0f  Settings"],
                    label_visibility="collapsed")
    st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Customer Onboarding
# ══════════════════════════════════════════════════════════════════════════════
def page_customer_onboarding():
    st.markdown("## \U0001f465 Customer Onboarding")
    st.caption("Same GSTIN/PAN checksum validation as vendor onboarding, same "
               "approve-to-activate gate. An approved customer with no credit "
               "limit on file stays on hold until one is set.")
    st.divider()

    s = co.stats()
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total Customers", s["total"])
    m2.metric("Approved", s["approved"])
    m3.metric("Format Verified", s["by_status"].get("Format Verified", 0))
    m4.metric("Credit Active", s["by_credit"].get("Active", 0))
    m5.metric("Credit On Hold", s["by_credit"].get("On Hold", 0))
    st.divider()

    st.markdown("#### \U0001f4cb Customers")
    customers = co.list_customers()
    if customers:
        cdf = pd.DataFrame(customers)
        show_cols = ["Customer_ID", "Customer_Name", "Customer_Type", "City", "GSTIN",
                     "Credit_Limit", "Credit_Status", "Active", "Onboarding_Status"]
        show_cols = [c for c in show_cols if c in cdf.columns]
        st.dataframe(cdf[show_cols], use_container_width=True, hide_index=True)

        pending = [c for c in customers if c.get("Onboarding_Status") == "Format Verified"]
        if pending:
            st.caption(f"\u23f3 {len(pending)} customer(s) format-verified and awaiting approval:")
            for c in pending:
                c1, c2 = st.columns([4, 1])
                c1.markdown(f"**{c['Customer_ID']}** — {c['Customer_Name']}")
                if c2.button("\u2705 Approve", key=f"capprove_{c['Customer_ID']}"):
                    co.approve_customer(c["Customer_ID"])
                    st.cache_data.clear()
                    st.success(f"{c['Customer_ID']} approved.")
                    st.rerun()

        held = [c for c in customers if c.get("Credit_Status") == "On Hold" and c.get("Active") == "Yes"]
        if held:
            st.caption(f"\U0001f6a7 {len(held)} approved customer(s) on credit hold:")
            for c in held:
                c1, c2, c3 = st.columns([3, 1.3, 1])
                limit_txt = "none set" if not c.get("Credit_Limit") else f"\u20b9{c['Credit_Limit']:,.0f}"
                c1.markdown(f"**{c['Customer_ID']}** — {c['Customer_Name']} (limit: {limit_txt})")
                new_limit = c2.number_input("Set limit (\u20b9)", min_value=0, step=10000,
                                            key=f"climit_{c['Customer_ID']}", label_visibility="collapsed")
                if c3.button("\u2705 Release", key=f"crelease_{c['Customer_ID']}"):
                    try:
                        if new_limit and new_limit > 0:
                            co.set_credit_limit(c["Customer_ID"], new_limit)
                        co.release_hold(c["Customer_ID"])
                        st.cache_data.clear()
                        st.success(f"{c['Customer_ID']} released.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
    else:
        st.info("No customers yet.")

    st.divider()
    st.markdown("#### \u2795 Onboard a Customer")

    left, right = st.columns(2, gap="large")
    with left:
        cid = st.text_input("Customer ID (short code, e.g. KOCHIMC)", key="co_id")
        cname = st.text_input("Customer Name", key="co_name")
        ctype = st.selectbox("Customer Type", co.list_customer_types(), key="co_type")
        gstin = st.text_input("GSTIN", key="co_gstin", placeholder="e.g. 27AABCU9603R1ZN")
        if gstin:
            ok, msg, _ = vo.validate_gstin(gstin)
            (st.success if ok else st.error)(f"GSTIN: {msg}")
        pan = st.text_input("PAN", key="co_pan", placeholder="e.g. AABCU9603R")
        if pan:
            ok, msg = vo.validate_pan_format(pan)
            (st.success if ok else st.error)(f"PAN: {msg}")
        city = st.text_input("City", key="co_city")
        country = st.text_input("Country", value="India", key="co_country")

    with right:
        address = st.text_input("Address", key="co_addr")
        geo = st.text_input("Geolocation (lat,lng)", key="co_geo", placeholder="9.9312,76.2673")
        contact_name = st.text_input("Contact Name", key="co_cname")
        contact_email = st.text_input("Contact Email", key="co_cemail")
        contact_phone = st.text_input("Contact Phone", key="co_cphone")
        payment_terms = st.selectbox("Payment Terms",
            ["Net 30", "Net 45", "Net 60", "Advance", "COD"], key="co_terms")
        credit_limit = st.number_input("Initial Credit Limit (\u20b9, optional)",
            min_value=0, step=10000, key="co_credit")

    st.markdown("")
    if st.button("\u2705  Validate & Save Customer", type="primary", use_container_width=True, key="co_save"):
        if not cid or not cname:
            st.error("Customer ID and Customer Name are required.")
        else:
            try:
                result = co.upsert_customer(cid, {
                    "Customer_Name": cname, "Customer_Type": ctype, "GSTIN": gstin, "PAN": pan,
                    "City": city, "Country": country, "Address": address, "Geolocation": geo,
                    "Contact_Name": contact_name, "Contact_Email": contact_email,
                    "Contact_Phone": contact_phone, "Payment_Terms": payment_terms,
                    "Credit_Limit": credit_limit,
                })
                st.cache_data.clear()
                for field, chk in result["checks"].items():
                    (st.success if chk["ok"] else st.error)(f"{field}: {chk['message']}")
                if result["onboarding_status"] == "Format Verified":
                    st.success(f"\u2705 {cid} saved — format verified, awaiting approval below.")
                    st.rerun()
                else:
                    st.warning(f"\u26a0\ufe0f {cid} saved as **Needs Review** — fix the "
                              "flagged field(s) above and save again.")
            except Exception:
                st.error("\u274c Error saving customer:")
                st.code(traceback.format_exc())

    if customers:
        st.divider()
        st.markdown("#### \U0001f4ce Document Log")
        dv = st.selectbox("Customer", [c["Customer_ID"] for c in customers], key="co_doc_customer")
        c1, c2, c3 = st.columns([2, 2, 3])
        with c1:
            dtype = st.selectbox("Document Type", ["GST Certificate", "PAN Card",
                "Credit Application", "Trade Reference", "Other"], key="co_doc_type")
        with c2:
            dfname = st.text_input("Filename / Reference", key="co_doc_fname")
        with c3:
            dnotes = st.text_input("Notes", key="co_doc_notes")
        if st.button("\U0001f4ce Log Document", key="co_doc_save"):
            if not dfname:
                st.warning("Enter a filename/reference first.")
            else:
                co.record_document(dv, dtype, dfname, dnotes)
                st.success(f"Logged {dtype} for {dv}.")
                st.rerun()

        docs = co.get_documents()
        if docs:
            with st.expander(f"\U0001f4c1 {len(docs)} document(s) logged"):
                st.dataframe(pd.DataFrame(docs), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Quotation
# ══════════════════════════════════════════════════════════════════════════════
def page_quotation():
    st.markdown("## \U0001f4b0 Quotation")
    st.caption("The reverse of RFQ — here we state the price, the customer accepts "
               "or declines. Suggested prices are Item Master's cost marked up "
               f"{int(qt.DEFAULT_MARKUP*100)}% as a placeholder — always editable.")
    st.divider()

    s = qt.stats()
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total Quotes", s["total"])
    m2.metric("Draft", s["by_status"].get("Draft", 0))
    m3.metric("Sent", s["by_status"].get("Sent", 0))
    m4.metric("Accepted", s["by_status"].get("Accepted", 0))
    m5.metric("Accepted Value", f"\u20b9{s['accepted_value']:,.0f}")
    st.divider()

    tab1, tab2 = st.tabs(["\u2795 Create Quote", "\U0001f4cb Manage Quotes"])

    # ── TAB 1 — build and price a new quote ──────────────────────────────────────
    with tab1:
        customers = co.list_customers()
        if not customers:
            st.info("No customers yet — onboard one first (a draft record is fine for quoting).")
        else:
            labels = {c["Customer_ID"]: f"{c['Customer_ID']} — {c['Customer_Name']} "
                      f"({c.get('Onboarding_Status') or 'Draft'})" for c in customers}
            sel_cust = st.selectbox("Customer", list(labels.keys()),
                                    format_func=lambda k: labels[k], key="qt_customer")
            customer = next(c for c in customers if c["Customer_ID"] == sel_cust)
            if customer.get("Active") != "Yes":
                st.caption("\u2139\ufe0f Not approved/credit-active yet — fine for a quote, "
                          "resolve before converting to an order.")

            c1, c2 = st.columns(2)
            with c1:
                valid_days = st.number_input("Valid for (days)", min_value=1, value=14, key="qt_valid_days")
            with c2:
                default_terms = customer.get("Payment_Terms") or "Net 30"
                st.caption(f"Payment terms default: **{default_terms}**")
            notes = st.text_area("Notes / terms shown on the quote", key="qt_notes", height=70)

            st.markdown("#### \U0001f50d Add items")
            lines = item_picker("qt")

            if lines:
                st.markdown("#### \U0001f4b5 Set quote pricing")
                st.caption("Cost is Item Master's procurement price — shown for margin "
                          "visibility only, not printed on the customer's document.")

                for i, ln in enumerate(lines):
                    c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1.3, 0.6])
                    c1.write(f"**{ln['desc']}**  \n{ln['code']} · {ln['uom']}")
                    new_qty = c2.number_input("Qty", min_value=1, value=int(ln["qty"]),
                                              key=f"qt_qty_{i}", label_visibility="collapsed")
                    c3.write(f"Cost: \u20b9{ln['price']:,.2f}")
                    default_price = qt.suggest_sale_price(ln["price"])
                    new_price = c4.number_input("Quote price", min_value=0.0,
                        value=float(ln.get("quote_price", default_price)),
                        key=f"qt_price_{i}", label_visibility="collapsed")
                    lines[i]["qty"] = new_qty
                    lines[i]["quote_price"] = new_price
                    if c5.button("\u2715", key=f"qt_rm_{i}"):
                        st.session_state["qt_lines"].pop(i)
                        st.rerun()

                quote_total = sum(ln["qty"] * ln.get("quote_price", qt.suggest_sale_price(ln["price"]))
                                  for ln in lines)
                cost_total = sum(ln["qty"] * ln["price"] for ln in lines)
                mc1, mc2, mc3 = st.columns(3)
                mc1.metric("Lines", len(lines))
                mc2.metric("Quote Total", f"\u20b9{quote_total:,.2f}")
                mc3.metric("Est. Margin", f"\u20b9{quote_total - cost_total:,.2f}")

                st.markdown("")
                create_clicked = st.button("\U0001f4b0  Create Quote", type="primary",
                                           use_container_width=True, key="qt_create")
                if create_clicked:
                    try:
                        line_items = [{"mat_code": ln["code"], "mat_desc": ln["desc"],
                            "uom": ln["uom"], "qty": ln["qty"],
                            "unit_price": ln.get("quote_price", qt.suggest_sale_price(ln["price"]))}
                            for ln in lines]
                        quote_id = qt.create_quote(sel_cust, line_items,
                            valid_days=int(valid_days), notes=notes)
                        st.session_state["qt_lines"] = []
                        st.cache_data.clear()
                        st.success(f"\u2705 {quote_id} created — \u20b9{quote_total:,.2f}. "
                                  "Head to **Manage Quotes** to generate the document.")
                    except Exception:
                        st.error("\u274c Error creating quote:")
                        st.code(traceback.format_exc())

    # ── TAB 2 — generate, send, and record the outcome ───────────────────────────
    with tab2:
        quotes = qt.get_quotes()
        if not quotes:
            st.info("No quotes yet.")
        else:
            qdf = pd.DataFrame(quotes)[["quote_id","customer_name","status","total_value",
                                         "quote_date","valid_until"]]
            qdf.columns = ["Quote","Customer","Status","Total","Date","Valid Until"]
            st.dataframe(qdf, use_container_width=True, hide_index=True)

            labels = {q["quote_id"]: f"{q['quote_id']} — {q['customer_name']} ({q['status']})" for q in quotes}
            sel = st.selectbox("Work on", list(labels.keys()), format_func=lambda k: labels[k], key="qt_work_sel")
            quote = next(q for q in quotes if q["quote_id"] == sel)
            items = qt.get_quote_items(sel)

            left, right = st.columns([3, 2], gap="large")
            with left:
                st.markdown("##### \U0001f4c4 Line Items")
                idf = pd.DataFrame(items)[["mat_code","mat_desc","uom","qty","unit_price","line_total"]]
                idf.columns = ["Code","Description","UOM","Qty","Price","Line Total"]
                st.dataframe(idf, use_container_width=True, hide_index=True)
                st.caption(f"Total: \u20b9{quote['total_value']:,.2f}  \u00b7  "
                          f"Valid until {quote['valid_until']}  \u00b7  {quote['payment_terms']}")

            with right:
                st.markdown("##### \u2696\ufe0f Actions")
                gen_clicked = st.button("\U0001f4c4 Generate Quote Document", key="qt_gen_doc")
                if gen_clicked:
                    fname, fbytes = qt.generate_quote_document(sel)
                    st.session_state.qt_generated_doc = {"quote_id": sel, "filename": fname, "bytes": fbytes}
                gd = st.session_state.get("qt_generated_doc")
                if gd and gd["quote_id"] == sel:
                    st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                        file_name=gd["filename"], mime=XLSX_MIME, key=f"qt_dl_{sel}")

                if quote["status"] == "Draft":
                    if st.button("\U0001f4e4 Mark Sent", key="qt_send"):
                        fname = gd["filename"] if (gd and gd["quote_id"] == sel) else None
                        qt.mark_sent(sel, fname)
                        st.cache_data.clear()
                        st.success(f"{sel} marked Sent.")
                        st.rerun()
                elif quote["status"] == "Sent":
                    if st.button("\U0001f3b2 Simulate Customer Response (Demo)", key="qt_sim"):
                        outcome = qt.simulate_response(sel)
                        st.cache_data.clear()
                        st.success(f"Simulated: {'Accepted' if outcome else 'Rejected'}.")
                        st.rerun()
                    st.markdown("###### Or record a real response")
                    decision = st.radio("Outcome", ["Accepted", "Rejected"], key=f"qt_dec_{sel}",
                                        horizontal=True)
                    reason = st.text_input("Reason (optional)", key=f"qt_reason_{sel}")
                    if st.button("Save Response", key=f"qt_save_dec_{sel}"):
                        qt.record_response(sel, decision == "Accepted", reason)
                        st.cache_data.clear()
                        st.success(f"{sel} recorded as {decision}.")
                        st.rerun()
                elif quote["status"] == "Accepted":
                    st.success("\u2705 Accepted — ready to convert to a Sales Order "
                              "on the Sales Orders page.")
                elif quote["status"] == "Rejected":
                    st.warning("\u274c Rejected." + (f" {quote['notes']}" if quote.get("notes") else ""))
                elif quote["status"] == "Expired":
                    st.warning("\u23f0 Expired — create a new quote at current pricing.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Sales Orders
# ══════════════════════════════════════════════════════════════════════════════
def page_sales_orders():
    st.markdown("## \U0001f4e6 Sales Orders")
    st.caption("Where credit finally gets checked. Quoting a customer never "
               "required approval — committing to an order does. A customer who "
               "fails the check still gets an order record, just held rather "
               "than confirmed, same as vendor onboarding's approval gate.")
    st.divider()

    s = so.stats()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Orders", s["total"])
    m2.metric("Confirmed", s["by_status"].get("Confirmed", 0))
    m3.metric("Credit Hold", s["by_status"].get("Credit Hold", 0))
    m4.metric("Confirmed Value", f"\u20b9{s['confirmed_value']:,.0f}")
    st.divider()

    tab1, tab2, tab3, tab4 = st.tabs(["\u2795 Create Order", "\U0001f4cb Manage Orders",
                                "\U0001f4e5 Bulk Import", "\U0001f514 Backorders"])

    # ── TAB 1 — from an accepted quote, or direct entry ──────────────────────────
    with tab1:
        st.markdown("#### \U0001f4dd From an Accepted Quote")
        accepted = qt.get_quotes(status="Accepted")
        eligible = [q for q in accepted if not so.quote_already_converted(q["quote_id"])]
        if not eligible:
            st.caption("No accepted, not-yet-converted quotes right now.")
        else:
            labels = {q["quote_id"]: f"{q['quote_id']} — {q['customer_name']} "
                      f"(\u20b9{q['total_value']:,.2f})" for q in eligible}
            sel_q = st.selectbox("Quote", list(labels.keys()), format_func=lambda k: labels[k], key="so_quote_sel")
            c1, c2 = st.columns(2)
            with c1:
                deliv_locs = load_delivery_locs()
                if deliv_locs:
                    loc_labels = {d["id"]: d["name"] for d in deliv_locs}
                    sel_loc_id = st.selectbox("Delivery Location", list(loc_labels.keys()),
                        format_func=lambda k: loc_labels[k], key="so_q_deliv")
                    # Store the ID, not the display name — GR/inventory postings key
                    # on location_id everywhere else in the app (a PO line's own
                    # Delivery_Location is an ID). Storing the name here instead
                    # made BOM's on-hand/transfer view treat this order's delivery
                    # point as a location distinct from every GR ever posted
                    # against it, even when both referred to the exact same place.
                    deliv_loc = sel_loc_id
                    deliv_geo = next(d["geo"] for d in deliv_locs if d["id"] == sel_loc_id)
                else:
                    st.warning("No delivery locations set up on the S2P app yet.")
                    deliv_loc, deliv_geo = "", ""
            with c2:
                deliv_date = st.date_input("Requested Delivery Date",
                    value=date.today()+timedelta(days=7), key="so_q_date")
            convert_clicked = st.button("\U0001f4e6 Create Order from Quote", type="primary", key="so_from_quote")
            if convert_clicked:
                try:
                    result = so.create_order_from_quote(sel_q, delivery_location=deliv_loc,
                        delivery_geo=deliv_geo, requested_delivery_date=deliv_date)
                    st.cache_data.clear()
                    if result["status"] == "Confirmed":
                        st.success(f"\u2705 {result['so_id']} confirmed — \u20b9{result['total_value']:,.2f}.")
                    else:
                        st.warning(f"\u26a0\ufe0f {result['so_id']} created on **Credit Hold** — "
                                  f"{result['credit_reason']}")
                except Exception:
                    st.error("\u274c Error creating order:")
                    st.code(traceback.format_exc())

        st.divider()
        st.markdown("#### \u26a1 Direct Order Entry")
        st.caption("For repeat/routine orders with no separate quote step. Prices from "
                  "Item Master cost with the same placeholder markup Quotation uses — "
                  "there's no customer rate-agreement module yet to price off instead.")

        customers = co.list_customers()
        if not customers:
            st.info("No customers yet.")
        else:
            clabels = {c["Customer_ID"]: f"{c['Customer_ID']} — {c['Customer_Name']} "
                       f"({c.get('Credit_Status') or 'Not Set'})" for c in customers}
            sel_cust = st.selectbox("Customer", list(clabels.keys()),
                                    format_func=lambda k: clabels[k], key="so_direct_customer")

            lines = item_picker("so")
            if lines:
                for i, ln in enumerate(lines):
                    c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1.3, 0.6])
                    c1.write(f"**{ln['desc']}**  \n{ln['code']} · {ln['uom']}")
                    new_qty = c2.number_input("Qty", min_value=1, value=int(ln["qty"]),
                                              key=f"so_qty_{i}", label_visibility="collapsed")
                    c3.write(f"Cost: \u20b9{ln['price']:,.2f}")
                    default_price = qt.suggest_sale_price(ln["price"])
                    new_price = c4.number_input("Price", min_value=0.0,
                        value=float(ln.get("quote_price", default_price)),
                        key=f"so_price_{i}", label_visibility="collapsed")
                    lines[i]["qty"] = new_qty
                    lines[i]["quote_price"] = new_price
                    if c5.button("\u2715", key=f"so_rm_{i}"):
                        st.session_state["so_lines"].pop(i)
                        st.rerun()

                order_total = sum(ln["qty"] * ln.get("quote_price", qt.suggest_sale_price(ln["price"]))
                                  for ln in lines)
                st.metric("Order Total", f"\u20b9{order_total:,.2f}")

                deliv_locs2 = load_delivery_locs()
                if deliv_locs2:
                    loc_labels2 = {d["id"]: d["name"] for d in deliv_locs2}
                    sel_loc_id2 = st.selectbox("Delivery Location", list(loc_labels2.keys()),
                        format_func=lambda k: loc_labels2[k], key="so_d_deliv")
                    deliv_loc2 = sel_loc_id2  # store the ID — see comment on the quote-based picker above
                    deliv_geo2 = next(d["geo"] for d in deliv_locs2 if d["id"] == sel_loc_id2)
                else:
                    st.warning("No delivery locations set up on the S2P app yet.")
                    deliv_loc2, deliv_geo2 = "", ""

                st.markdown("")
                direct_clicked = st.button("\U0001f4e6 Create Direct Order", type="primary", key="so_direct_create")
                if direct_clicked:
                    try:
                        line_items = [{"mat_code": ln["code"], "mat_desc": ln["desc"],
                            "uom": ln["uom"], "qty": ln["qty"],
                            "unit_price": ln.get("quote_price", qt.suggest_sale_price(ln["price"]))}
                            for ln in lines]
                        result = so.create_direct_order(sel_cust, line_items,
                            delivery_location=deliv_loc2, delivery_geo=deliv_geo2)
                        st.session_state["so_lines"] = []
                        st.cache_data.clear()
                        if result["status"] == "Confirmed":
                            st.success(f"\u2705 {result['so_id']} confirmed — \u20b9{result['total_value']:,.2f}.")
                        else:
                            st.warning(f"\u26a0\ufe0f {result['so_id']} created on **Credit Hold** — "
                                      f"{result['credit_reason']}")
                    except Exception:
                        st.error("\u274c Error creating order:")
                        st.code(traceback.format_exc())

    # ── TAB 2 — manage existing orders ────────────────────────────────────────────
    with tab2:
        orders = so.get_orders()
        if not orders:
            st.info("No orders yet.")
        else:
            odf = pd.DataFrame(orders)[["so_id","customer_name","status","total_value",
                                         "order_date","source_quote"]]
            odf.columns = ["Order","Customer","Status","Total","Date","Source Quote"]
            st.dataframe(odf, use_container_width=True, hide_index=True)

            labels = {o["so_id"]: f"{o['so_id']} — {o['customer_name']} ({o['status']})" for o in orders}
            sel = st.selectbox("Work on", list(labels.keys()), format_func=lambda k: labels[k], key="so_work_sel")
            order = next(o for o in orders if o["so_id"] == sel)
            items = so.get_order_items(sel)

            left, right = st.columns([3, 2], gap="large")
            with left:
                st.markdown("##### \U0001f4c4 Line Items")
                idf = pd.DataFrame(items)[["mat_code","mat_desc","uom","qty","unit_price","line_total"]]
                idf.columns = ["Code","Description","UOM","Qty","Price","Line Total"]
                atp_icons = {"Promised": "\u2705 Promised", "Partially Promised": "\u26a0\ufe0f Partial",
                            "Backordered": "\u274c Backordered"}
                idf["ATP"] = [atp_icons.get(i.get("atp_outcome"), "\u2014") +
                             (f" ({i['backordered_qty']:g} short)"
                              if i.get("atp_outcome") in ("Partially Promised", "Backordered")
                              and i.get("backordered_qty") else "")
                             for i in items]
                st.dataframe(idf, use_container_width=True, hide_index=True)
                st.caption(f"Total: \u20b9{order['total_value']:,.2f}  \u00b7  "
                          f"{order['payment_terms']}  \u00b7  Delivery: {_loc_display(order['delivery_location']) or 'TBD'}")

            with right:
                st.markdown("##### \u2696\ufe0f Actions")
                gen_clicked = st.button("\U0001f4c4 Generate Order Confirmation", key="so_gen_doc")
                if gen_clicked:
                    fname, fbytes = so.generate_order_confirmation(sel)
                    st.session_state.so_generated_doc = {"so_id": sel, "filename": fname, "bytes": fbytes}
                gd = st.session_state.get("so_generated_doc")
                if gd and gd["so_id"] == sel:
                    st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                        file_name=gd["filename"], mime=XLSX_MIME, key=f"so_dl_{sel}")

                if order["status"] == "Credit Hold":
                    st.warning("On credit hold.")
                    if st.button("\U0001f501 Re-check Credit & Release", key="so_release"):
                        result = so.release_credit_hold(sel)
                        st.cache_data.clear()
                        if result["status"] == "Confirmed":
                            st.success(f"\u2705 Released — {result['reason']}")
                        else:
                            st.warning(f"Still held — {result['reason']}")
                        st.rerun()
                elif order["status"] == "Confirmed":
                    st.success("\u2705 Confirmed — ready for fulfillment on the Fulfillment page.")

                if order["status"] in ("Confirmed", "Credit Hold"):
                    with st.popover("\u274c Cancel order"):
                        reason = st.text_input("Reason", key=f"so_cancel_reason_{sel}")
                        if st.button("Confirm Cancellation", key=f"so_cancel_btn_{sel}"):
                            so.cancel_order(sel, reason)
                            st.cache_data.clear()
                            st.success(f"{sel} cancelled.")
                            st.rerun()

    # ── TAB 3 — bulk import ────────────────────────────────────────────────────
    with tab3:
        import sales_order_import as soi
        st.caption("For seeding real demand history in bulk — reuses the exact same "
                  "order creation and real-time credit check as Create Order above, "
                  "so a bulk-imported order can land on Credit Hold just like a manual "
                  "one. This is what makes time-phased planning's Sales Order Based "
                  "mode work without hand-entering orders one at a time.")

        st.markdown("##### 1. Download the template")
        st.caption("Includes every active customer, material, and delivery location as "
                  "dropdown choices, plus reference sheets so you don't need to "
                  "memorize any ID. Delivery Location can be left blank on any row to "
                  "use that customer's own default location — see the Customers "
                  "reference sheet for what each customer's default resolves to.")
        template_bytes = soi.generate_template()
        st.download_button("\U0001f4e5 Download Sales Order Import Template", data=template_bytes,
            file_name="sales_order_import_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.divider()
        st.markdown("##### 2. Upload it back, filled in")
        st.caption("Rows sharing the same Order Reference become one order's separate "
                  "line items. Validation is per-order, not per-line — if any line in "
                  "an order fails, the whole order is rejected with every specific "
                  "reason shown, and the rest of the file is unaffected.")
        uploaded = st.file_uploader("Filled-in template", type=["xlsx"], key="so_bulk_upload")
        if uploaded and st.button("\U0001f4e4 Upload Sales Orders", key="so_bulk_upload_btn"):
            try:
                result = soi.import_sales_orders(uploaded.read())
                st.cache_data.clear()
                if result["accepted"]:
                    st.success(f"\u2705 {len(result['accepted'])} order(s) created.")
                    adf = pd.DataFrame(result["accepted"])
                    adf.columns = ["Order Reference", "Sales Order", "Status"]
                    st.dataframe(adf, use_container_width=True, hide_index=True)
                if result["rejected"]:
                    st.error(f"\u274c {len(result['rejected'])} order(s) rejected.")
                    for r in result["rejected"]:
                        st.write(f"**{r['order_ref']}**: " + "; ".join(r["reasons"]))
                if not result["accepted"] and not result["rejected"]:
                    st.info("No rows found in the Import sheet.")
            except Exception as e:
                st.error(f"\u274c {e}")

    # ── TAB 4 — Backorder worklist (ATP-US-03) ────────────────────────────────
    with tab4:
        st.caption("Every currently-active Backorder \u2014 quantity ATP-US-01 couldn't "
                  "promise at confirmation, automatically re-evaluated on every real "
                  "supply arrival (a Goods Receipt or a transfer Confirm Receipt), "
                  "never a batch job. Fulfilled strictly First-Confirmed-First-Served "
                  "when several compete for the same new supply.")
        active = bo.get_open_backorders()
        if not active:
            st.info("No open backorders right now.")
        else:
            for b in active:
                c1, c2 = st.columns([4, 1.3])
                status_icon = "\u26a0\ufe0f" if b["status"] == "Partially Fulfilled" else "\u274c"
                c1.write(f"**{b['backorder_id']}** {status_icon} {b['status']} \u2014 "
                        f"{b['material_desc']} \u00b7 {b['open_qty']:g} of "
                        f"{b['original_qty']:g} still open \u00b7 {_loc_display(b['location_id'])} "
                        f"\u00b7 {b['so_id']} line {b['so_line_item']} \u00b7 "
                        f"confirmed {b['created_date']}")
                if c2.button("\u274c Cancel", key=f"cancel_bo_{b['backorder_id']}"):
                    try:
                        bo.cancel_backorder(b["backorder_id"],
                            reason="Cancelled manually from Backorder worklist")
                        st.cache_data.clear()
                        st.success(f"\u2705 {b['backorder_id']} cancelled \u2014 "
                                  f"{b['open_qty']:g} units removed from future "
                                  f"fulfillment consideration.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"\u274c {e}")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Fulfillment
# ══════════════════════════════════════════════════════════════════════════════
def page_fulfillment():
    st.markdown("## \U0001f69a Fulfillment")
    st.caption("Pick \u2192 Ship \u2192 Deliver. No live inventory data exists yet "
               "(Item Master's stock column has never been populated), so this "
               "doesn't pretend to check availability automatically — it tracks "
               "what a human actually picked and shipped per line, honestly, "
               "including partial shipments.")
    st.divider()

    s = ful.stats()
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total", s["total"])
    m2.metric("Pending", s["by_status"].get("Pending", 0))
    m3.metric("Picking", s["by_status"].get("Picking", 0))
    m4.metric("Shipped", s["by_status"].get("Shipped", 0))
    m5.metric("Delivered", s["by_status"].get("Delivered", 0))
    st.divider()

    tab1, tab2 = st.tabs(["\u2795 Create Fulfillment", "\U0001f4cb Manage Fulfillments"])

    # ── TAB 1 — start fulfillment for a confirmed order ──────────────────────────
    with tab1:
        confirmed = so.get_orders(status="Confirmed")
        eligible = [o for o in confirmed if not ful.so_already_fulfilled(o["so_id"])]
        if not eligible:
            st.info("No confirmed orders waiting on fulfillment.")
        else:
            labels = {o["so_id"]: f"{o['so_id']} — {o['customer_name']} (\u20b9{o['total_value']:,.2f})"
                      for o in eligible}
            sel_so = st.selectbox("Confirmed Order", list(labels.keys()),
                                  format_func=lambda k: labels[k], key="ful_so_sel")
            order = next(o for o in eligible if o["so_id"] == sel_so)
            items = so.get_order_items(sel_so)
            idf = pd.DataFrame(items)[["mat_code","mat_desc","uom","qty"]]
            idf.columns = ["Code","Description","UOM","Qty"]
            st.dataframe(idf, use_container_width=True, hide_index=True)
            st.caption(f"Deliver to: {_loc_display(order['delivery_location']) or 'not specified'}")

            create_clicked = st.button("\U0001f69a Create Fulfillment", type="primary", key="ful_create")
            if create_clicked:
                try:
                    fid = ful.create_fulfillment(sel_so)
                    st.cache_data.clear()
                    st.success(f"\u2705 {fid} created for {sel_so} — Pending. "
                              "Head to **Manage Fulfillments** to pick and ship.")
                except Exception:
                    st.error("\u274c Error creating fulfillment:")
                    st.code(traceback.format_exc())

    # ── TAB 2 — pick, ship, deliver, or cancel ────────────────────────────────────
    with tab2:
        fulfillments = ful.get_fulfillments()
        if not fulfillments:
            st.info("No fulfillments yet.")
        else:
            fdf = pd.DataFrame(fulfillments)[["fulfillment_id","so_id","customer_name","status",
                                               "shipped_date","delivered_date"]]
            fdf.columns = ["Fulfillment","Order","Customer","Status","Shipped","Delivered"]
            st.dataframe(fdf, use_container_width=True, hide_index=True)

            labels = {f["fulfillment_id"]: f"{f['fulfillment_id']} — {f['customer_name']} ({f['status']})"
                      for f in fulfillments}
            sel = st.selectbox("Work on", list(labels.keys()), format_func=lambda k: labels[k], key="ful_work_sel")
            f = next(x for x in fulfillments if x["fulfillment_id"] == sel)
            items = ful.get_fulfillment_items(sel)

            left, right = st.columns([3, 2], gap="large")
            with left:
                st.markdown("##### \U0001f4c4 Line Items")
                if f["status"] in ("Pending", "Picking"):
                    st.caption("Enter what's actually being shipped per line before marking Shipped.")
                    ship_qtys = {}
                    for i, it in enumerate(items):
                        c1, c2 = st.columns([3, 1])
                        c1.write(f"**{it['mat_desc']}**  \n{it['mat_code']} · {it['uom']} · "
                                f"ordered {it['qty_ordered']}")
                        ship_qtys[it["mat_code"]] = c2.number_input("Qty to ship",
                            min_value=0, max_value=int(it["qty_ordered"]),
                            value=int(it["qty_ordered"]), key=f"ful_ship_{sel}_{i}")
                    st.session_state[f"ful_ship_qtys_{sel}"] = ship_qtys
                else:
                    idf = pd.DataFrame(items)[["mat_code","mat_desc","uom","qty_ordered","qty_shipped"]]
                    idf.columns = ["Code","Description","UOM","Ordered","Shipped"]
                    st.dataframe(idf, use_container_width=True, hide_index=True)

            with right:
                st.markdown("##### \u2696\ufe0f Actions")
                gen_clicked = st.button("\U0001f4c4 Generate Delivery Note", key="ful_gen_doc")
                if gen_clicked:
                    fname, fbytes = ful.generate_delivery_note(sel)
                    st.session_state.ful_generated_doc = {"fulfillment_id": sel, "filename": fname, "bytes": fbytes}
                gd = st.session_state.get("ful_generated_doc")
                if gd and gd["fulfillment_id"] == sel:
                    st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                        file_name=gd["filename"], mime=XLSX_MIME, key=f"ful_dl_{sel}")

                if f["status"] == "Pending":
                    if st.button("\U0001f4e6 Start Picking", key="ful_pick"):
                        ful.start_picking(sel)
                        st.cache_data.clear()
                        st.success(f"{sel} now Picking.")
                        st.rerun()

                if f["status"] in ("Pending", "Picking"):
                    carrier = st.text_input("Carrier", key=f"ful_carrier_{sel}")
                    tracking = st.text_input("Tracking Ref", key=f"ful_tracking_{sel}")
                    ship_clicked = st.button("\U0001f69a Mark Shipped", type="primary", key="ful_ship_btn")
                    if ship_clicked:
                        qtys = st.session_state.get(f"ful_ship_qtys_{sel}", {})
                        try:
                            result = ful.record_shipment(sel, qtys, carrier, tracking)
                            st.cache_data.clear()
                            short = [it["mat_code"] for it in items
                                    if qtys.get(it["mat_code"], it["qty_ordered"]) < it["qty_ordered"]]
                            if short:
                                st.warning(f"\u26a0\ufe0f Shipped with backorder on: {', '.join(short)}")
                            else:
                                st.success(f"\u2705 {sel} shipped in full.")
                            for w in result.get("warnings", []):
                                st.warning(f"\u26a0\ufe0f {w}")
                            st.rerun()
                        except Exception:
                            st.error("\u274c Error recording shipment:")
                            st.code(traceback.format_exc())

                elif f["status"] == "Shipped":
                    pod = st.text_input("Proof of delivery reference", key=f"ful_pod_{sel}")
                    deliver_clicked = st.button("\u2705 Mark Delivered", type="primary", key="ful_deliver")
                    if deliver_clicked:
                        ful.record_delivery(sel, pod)
                        st.cache_data.clear()
                        st.success(f"\u2705 {sel} delivered — ready for billing.")
                        try:
                            je_id = acct.post_fulfillment_entry(sel)
                            st.success(f"\U0001f4d2 Posted to ledger as {je_id} "
                                      "(COGS / Inventory Clearing).")
                        except Exception:
                            st.warning("\u26a0\ufe0f Delivery recorded, but the accounting "
                                      "entry failed to post:")
                            st.code(traceback.format_exc())
                        st.rerun()

                elif f["status"] == "Delivered":
                    st.success("\u2705 Delivered — ready for billing on the Billing & Invoicing page.")

                if f["status"] in ("Pending", "Picking", "Shipped"):
                    with st.popover("\u274c Cancel fulfillment"):
                        reason = st.text_input("Reason", key=f"ful_cancel_reason_{sel}")
                        if st.button("Confirm Cancellation", key=f"ful_cancel_btn_{sel}"):
                            ful.cancel_fulfillment(sel, reason)
                            st.cache_data.clear()
                            st.success(f"{sel} cancelled.")
                            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Settings (Org Profile + Item Tax)
# ══════════════════════════════════════════════════════════════════════════════
def page_settings():
    st.markdown("## \u2699\ufe0f Settings")
    st.caption("Setup, not a daily workflow — both of these gate Billing & Invoicing. "
               "Nothing here is guessed: an invoice can't be created until your "
               "organization has a real GSTIN on file and every line item being "
               "billed has a real HSN code and GST rate.")
    st.divider()

    tab1, tab2, tab3 = st.tabs(["\U0001f3e2 Organization Profile", "\U0001f3f7\ufe0f Item Tax (HSN/GST)",
                                "\U0001f504 Data Reset & Seed"])

    # ── TAB 1 — org profile ──────────────────────────────────────────────────────
    with tab1:
        configured = op.is_configured()
        (st.success if configured else st.warning)(
            "\u2705 Organization profile configured — invoices can be created." if configured
            else "\u26a0\ufe0f No valid GSTIN on file yet — invoicing is blocked until this is set.")

        profile = op.get_org_profile() or {}
        left, right = st.columns(2, gap="large")
        with left:
            org_id = st.text_input("Org ID", value=profile.get("Org_ID") or "", key="op_orgid",
                                   placeholder="e.g. GRB",
                                   help="A short business code, not validated against any "
                                        "format — like a company's own short name.")
            legal_name = st.text_input("Legal Name", value=profile.get("Legal_Name") or "", key="op_name")
            gstin = st.text_input("GSTIN", value=profile.get("GSTIN") or "", key="op_gstin",
                                  placeholder="e.g. 27AABCU9603R1ZN")
            if gstin:
                ok, msg, _ = vo.validate_gstin(gstin)
                (st.success if ok else st.error)(f"GSTIN: {msg}")
            pan = st.text_input("PAN", value=profile.get("PAN") or "", key="op_pan")
            if pan:
                ok, msg = vo.validate_pan_format(pan)
                (st.success if ok else st.error)(f"PAN: {msg}")
            address = st.text_input("Address", value=profile.get("Address") or "", key="op_addr")
            city = st.text_input("City", value=profile.get("City") or "", key="op_city")
            state = st.text_input("State", value=profile.get("State") or "", key="op_state")
            country = st.text_input("Country", value=profile.get("Country") or "India", key="op_country")
        with right:
            bank_acct = st.text_input("Bank Account No.", value=profile.get("Bank_Account_No") or "", key="op_bank")
            ifsc = st.text_input("IFSC", value=profile.get("IFSC") or "", key="op_ifsc")
            bank_name = st.text_input("Bank Name", value=profile.get("Bank_Name") or "", key="op_bname")
            contact_email = st.text_input("Contact Email", value=profile.get("Contact_Email") or "", key="op_cemail")
            contact_phone = st.text_input("Contact Phone", value=profile.get("Contact_Phone") or "", key="op_cphone")

        save_clicked = st.button("\u2705 Save Organization Profile", type="primary", key="op_save")
        if save_clicked:
            try:
                result = op.set_org_profile({
                    "Org_ID": org_id, "Legal_Name": legal_name, "GSTIN": gstin, "PAN": pan,
                    "Address": address, "City": city, "State": state, "Country": country,
                    "Bank_Account_No": bank_acct, "IFSC": ifsc, "Bank_Name": bank_name,
                    "Contact_Email": contact_email, "Contact_Phone": contact_phone,
                })
                for field, chk in result["checks"].items():
                    (st.success if chk["ok"] else st.error)(f"{field}: {chk['message']}")
                if result["is_configured"]:
                    st.success("\u2705 Saved — organization profile is configured.")
                else:
                    st.warning("\u26a0\ufe0f Saved, but GSTIN still isn't valid — invoicing stays blocked.")
            except Exception:
                st.error("\u274c Error saving profile:")
                st.code(traceback.format_exc())

        st.divider()
        st.markdown("##### \U0001f4e6 Demand Detection Mode")
        st.caption(
            "Controls what Inventory Position & Transfers (Manufacturing app) treats "
            "as real demand. **Manufactured Items Only** (default) — only Sales Order "
            "lines for items this org actually builds from a BOM count, right for a "
            "manufacturer. **All Items** — every confirmed Sales Order line counts, "
            "BOM or not — right for a trading/distribution business that sells finished "
            "goods it never manufactures, where the BOM-only version would show zero "
            "transfer opportunities no matter how real the stock imbalance is.")
        current_mode = od.get_default("Demand Detection Mode")
        mode = st.radio("Mode", ["Manufactured Items Only", "All Items"],
                        index=0 if current_mode != "All Items" else 1,
                        key="op_demand_mode", horizontal=True, label_visibility="collapsed")
        if st.button("\u2705 Save Demand Detection Mode", key="op_demand_mode_save"):
            od.set_org_default("Demand Detection Mode", mode)
            st.cache_data.clear()
            st.success(f"\u2705 Saved — set to '{mode}'.")

        st.divider()
        st.markdown("##### \U0001f3ed Stock Transfer Order Policy")
        st.caption(
            "Controls which action Position & Transfers offers for a shortage a Hub "
            "could resolve. **No** (default) — a direct Ship action, for "
            "organizations that move stock ad hoc. **Yes** — a named, auditable "
            "Stock Transfer Order covering every Plant competing for that Hub's "
            "stock as one allocation event, for organizations whose policy requires "
            "formal STO documentation for inter-location movement. E-way bill "
            "compliance applies either way, since it depends on the shipment "
            "itself, not this setting.")
        current_sto_policy = od.get_default("Use Stock Transfer Orders")
        sto_policy = st.radio("STO Policy", ["No", "Yes"],
                              index=0 if current_sto_policy != "Yes" else 1,
                              key="op_sto_policy", horizontal=True, label_visibility="collapsed")
        if st.button("\u2705 Save STO Policy", key="op_sto_policy_save"):
            od.set_org_default("Use Stock Transfer Orders", sto_policy)
            st.cache_data.clear()
            st.success(f"\u2705 Saved — set to '{sto_policy}'.")

    # ── TAB 2 — item tax ─────────────────────────────────────────────────────────
    with tab2:
        s = it.stats()
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Items", s["total"])
        m2.metric("Configured", s["configured"])
        m3.metric("Missing", s["missing"])

        missing = it.list_items_missing_tax_info()
        if missing:
            st.caption(f"{len(missing)} item(s) need an HSN code and GST rate before "
                      "they can be billed:")
            labels = {m["mat_code"]: f"{m['mat_code']} — {m['mat_desc']}" for m in missing}
            sel_item = st.selectbox("Item", list(labels.keys()), format_func=lambda k: labels[k], key="it_sel")

            c1, c2 = st.columns(2)
            with c1:
                hsn = st.text_input("HSN Code", key="it_hsn", placeholder="e.g. 90189000")
            with c2:
                gst_rate = st.selectbox("GST Rate (%)", it.GST_RATE_CHOICES, index=2, key="it_rate")

            set_clicked = st.button("\u2705 Set Tax Info", type="primary", key="it_set")
            if set_clicked:
                if not hsn:
                    st.error("Enter an HSN code first.")
                else:
                    it.set_item_tax_info(sel_item, hsn, gst_rate)
                    st.cache_data.clear()
                    st.success(f"\u2705 {sel_item} set — HSN {hsn}, GST {gst_rate}%.")
                    st.rerun()
        else:
            st.success("\u2705 Every item has an HSN code and GST rate on file.")

    # ── TAB 3 — data reset & seed ────────────────────────────────────────────────
    with tab3:
        st.warning("\u26a0\ufe0f This affects the shared database all three apps "
                  "(Source-to-Pay, Manufacturing, Order-to-Cash) read from — not "
                  "just this page. Restart all three after resetting.")
        st.caption(
            "For switching between industry-profile demo scenarios (e.g. Dental vs. "
            "Genrobotics). Replaces all master data — Org Profile, Org Defaults, "
            "Legal Entities, Delivery Locations, Item Master, Purchase Bundles, "
            "BOM Items, Vendor/Customer Types, Vendor Master, Customer Master, "
            "Chart of Accounts — and clears every transactional table (PRs, POs, "
            "GRs, Sales Orders, Invoices, Journal Entries, everything) since none "
            "of that history means anything once the underlying items/vendors/"
            "customers have changed out from under it.")

        uploaded = st.file_uploader("Seed file (.xlsx)", type=["xlsx"], key="seed_upload")

        if uploaded:
            tmp_path = os.path.join(_DIR, f"_seed_upload_{uploaded.name}")
            with open(tmp_path, "wb") as f:
                f.write(uploaded.getbuffer())
            try:
                result = sm.validate_seed_file(tmp_path)
            except Exception:
                st.error("\u274c Couldn't read this file:")
                st.code(traceback.format_exc())
                result = None

            if result and not result["valid"]:
                st.error(f"\u274c {len(result['errors'])} problem(s) found — fix these and "
                         "re-upload. Nothing has been changed.")
                for e in result["errors"]:
                    st.write(f"- {e}")
            elif result:
                st.success("\u2705 File is valid.")
                s = result["summary"]
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Items", s.get("Item Master", 0))
                c2.metric("Vendors", s.get("Vendor_Master", 0))
                c3.metric("Customers", s.get("Customer_Master", 0))
                c4.metric("BOM Lines", s.get("BOM_Items", 0))
                org_name = result["parsed"].get("org_profile", {}).get("Legal_Name", "?")
                st.caption(f"Organization detected: **{org_name}**")
                with st.expander("Full sheet counts"):
                    for sheet, n in s.items():
                        st.write(f"- {sheet}: {n}")

                st.divider()
                st.markdown("##### Before resetting")
                backup_bytes = sm.backup_current_state()
                st.download_button("\u2b07 Download a backup of the current data first "
                                   "(recommended, not required)",
                                   data=backup_bytes, file_name="erp_poc_backup.zip",
                                   mime="application/zip", key="seed_backup_dl")

                st.markdown("##### Confirm")
                confirm_text = st.text_input(
                    'Type RESET to confirm — this permanently replaces all master data '
                    'and clears every transaction.', key="seed_confirm")
                reset_clicked = st.button("\U0001f5d1 Reset & Reseed", type="primary",
                                          disabled=(confirm_text.strip() != "RESET"),
                                          key="seed_reset_btn")
                if reset_clicked:
                    try:
                        outcome = sm.reset_and_reseed(tmp_path)
                        st.cache_data.clear()
                        st.success(f"\u2705 Done — {outcome['tables_wiped']} table(s) reset, "
                                  f"{sum(outcome['seed_counts'].values())} row(s) seeded. "
                                  "Restart all three apps to pick up the new data.")
                        for table, n in outcome["seed_counts"].items():
                            st.write(f"- {table}: {n}")
                    except Exception:
                        st.error("\u274c Reset failed — nothing may have been fully applied. "
                                 "Restore from a backup if needed:")
                        st.code(traceback.format_exc())
            try:
                os.remove(tmp_path)
            except OSError:
                pass

        st.divider()
        DEMO_PROFILES = dp.DEMO_PROFILES
        current_org_id = (op.get_org_profile() or {}).get("Org_ID")
        profile = DEMO_PROFILES.get(current_org_id)

        st.markdown("##### \U0001f3ac Demo Scenario (optional)")
        if profile is None:
            available = ", ".join(f"{k} ({v['title']})" for k, v in DEMO_PROFILES.items())
            st.caption(f"No pre-built demo scenario for the current org "
                      f"('{current_org_id or 'not set'}'). Available for: {available}.")
        else:
            st.caption(f"**{profile['title']}** — {profile['description']} "
                      "See DEMO_GUIDE.md for the full walkthrough script.")

            import importlib
            ds = importlib.import_module(profile["module"])

            pr_count = len(pr_consolidation.get_pr_headers())
            so_count = len(so.get_orders())
            import fulfillment as ful
            fulfillment_count = len(ful.get_fulfillments())
            # fulfillment_count, not pr_count, is what actually distinguishes
            # "Setup ran, resolution hasn't yet" from "already resolved" —
            # IDS Denmed's Setup legitimately creates a real PR (the bulk
            # import now goes through a proper PR -> Consolidate -> PO chain,
            # not a direct PO insert), so a PR existing is no longer, on its
            # own, evidence that something collided. Fulfillment only ever
            # happens during resolution, for either profile, so it's the
            # reliable signal instead.

            if so_count == 0 and pr_count > 0:
                st.warning(f"\u26a0\ufe0f This database already has {pr_count} PR(s) on file but "
                          "no Sales Orders — that looks like leftover unrelated data, not "
                          "this demo's own state. Reset first (above) before loading the "
                          "demo scenario, to avoid colliding with it.")

            elif so_count > 0 and fulfillment_count == 0:
                st.info(f"\U0001f4cb Setup is loaded — {profile['setup_note']} Execute the "
                       "suggested transfers there first (that's the live moment), then "
                       "come back and click below to fast-forward everything else.")
                if st.button("\u23e9 Complete the Rest", key="demo_resolution_btn"):
                    try:
                        with st.spinner("Running procurement, fulfillment, billing, "
                                        "cash application…"):
                            ds.run_resolution()
                        st.cache_data.clear()
                        st.success(f"\u2705 Story complete — {profile['resolution_note']}")
                    except Exception:
                        st.error("\u274c Failed partway through — the database is likely in a "
                                 "mixed state now. Reset and try again:")
                        st.code(traceback.format_exc())

            elif so_count > 0 and fulfillment_count > 0:
                # The new customer wave is deliberately a separate,
                # later-triggered stage for any pilot that supports it,
                # not part of the automatic run above — see each
                # module's own new_customer_wave() docstring for why: a
                # smaller PR consolidation example is already on file
                # by this point; this button's whole purpose is to
                # create a second, much larger one live, so the
                # contrast between the two is a real, staged demo
                # moment rather than one flat reveal of everything at
                # once. Every piece of pilot-specific detail below
                # (which customer id proves the wave already ran, the
                # caption text, the resulting customer names) comes
                # from demo_profiles.py or the function's own real
                # return value — never hardcoded here again, after a
                # real bug found and fixed 2026-07-31 where this
                # section still said "SMC" and "Bandicoot bundle"
                # unconditionally, which would have been simply wrong
                # once IDS Denmed's own wave was enabled.
                supports_wave = bool(profile and profile.get("supports_customer_wave"))
                check_id = profile.get("wave_check_customer_id") if profile else None
                wave_already_run = (supports_wave and check_id
                                    and co.get_customer(check_id) is not None)
                if supports_wave and not wave_already_run:
                    st.success("\u2705 Story complete through cash application. A smaller "
                              "PR consolidation example is already on file — try it now "
                              "on S2C \u2192 Consolidate PRs \u2192 POs before adding the "
                              "wave below, so the contrast actually lands.")
                    st.markdown("**For the staged, high-impact consolidation moment**")
                    st.caption(profile.get("wave_caption", ""))
                    if st.button("\U0001f680 Add New Customer Wave", key="demo_wave_btn"):
                        try:
                            with st.spinner("Creating quotations, orders, and PRs for "
                                            "three new customers…"):
                                result = ds.new_customer_wave()
                            st.cache_data.clear()
                            names = ", ".join(r["customer"] for r in result)
                            st.success(f"\u2705 {names} are on file — three real, Open "
                                      f"PRs ready for a live PR Consolidation run.")
                        except Exception:
                            st.error("\u274c Failed partway through:")
                            st.code(traceback.format_exc())
                else:
                    st.info("\u2705 This story has already run all the way through "
                           "(fulfillment/billing/cash application are on file)"
                           + (", including the new customer wave" if wave_already_run else "")
                           + ". Reset first if you want to run it again.")

            else:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**For a live walkthrough**")
                    st.caption("Loads just the 'before' state — the imbalance, no "
                              "resolution. Go execute the transfers yourself on the "
                              "Inventory page afterward, then come back here for a "
                              "'Complete the Rest' button.")
                    if st.button("\u25b6\ufe0f Load Setup Only", key="demo_setup_btn"):
                        try:
                            with st.spinner("Setting up historical stock and demand…"):
                                ds.run_setup()
                            st.cache_data.clear()
                            st.success("\u2705 Setup loaded — head to Manufacturing app's "
                                      "Inventory \u2192 Position & Transfers to see the "
                                      "imbalance and execute the transfers live.")
                        except Exception:
                            st.error("\u274c Setup failed:")
                            st.code(traceback.format_exc())
                with c2:
                    st.markdown("**For a quick full preview**")
                    st.caption("Runs the whole story in one click, including the "
                              "transfer — no live interaction, just the finished "
                              "story. Good for a fast look, not for presenting the "
                              "transfer as a live moment.")
                    if st.button("\u23e9 Load Full Demo (quick preview)", key="demo_scenario_btn"):
                        try:
                            with st.spinner("Running the full story — historical stock, "
                                            "demand, transfers, procurement, fulfillment, "
                                            "billing, cash application…"):
                                ds.run_all()
                            st.cache_data.clear()
                            st.success("\u2705 Demo scenario loaded — head to Inventory "
                                      "Position & Transfers (Manufacturing app) to see the "
                                      "story from the top, or Accounting for the full ledger.")
                        except Exception:
                            st.error("\u274c Demo scenario failed partway through — the "
                                     "database is likely in a mixed state now. Reset and "
                                     "try again:")
                            st.code(traceback.format_exc())


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Billing & Invoicing
# ══════════════════════════════════════════════════════════════════════════════
def page_billing():
    st.markdown("## \U0001f9fe Billing & Invoicing")
    st.caption("Billed off Fulfillment, not the order — quantities reflect what "
               "actually shipped. CGST/SGST vs IGST is computed from the buyer's "
               "and seller's GSTIN state codes, no judgement call involved.")
    st.divider()

    s = bl.stats()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Invoices", s["total"])
    m2.metric("Draft", s["by_status"].get("Draft", 0))
    m3.metric("Issued", s["by_status"].get("Issued", 0))
    m4.metric("Total Value", f"\u20b9{s['total_value']:,.0f}")

    if not op.is_configured():
        st.warning("\u26a0\ufe0f No organization GSTIN on file — set that up on the "
                  "**Settings** page before creating invoices.")
    st.divider()

    tab1, tab2 = st.tabs(["\u2795 Create Invoice", "\U0001f4cb Manage Invoices"])

    # ── TAB 1 — invoice a delivered fulfillment ──────────────────────────────────
    with tab1:
        delivered = ful.get_fulfillments(status="Delivered")
        eligible = [f for f in delivered if not bl.fulfillment_already_invoiced(f["fulfillment_id"])]
        if not eligible:
            st.info("No delivered fulfillments waiting on an invoice.")
        else:
            labels = {f["fulfillment_id"]: f"{f['fulfillment_id']} — {f['customer_name']} "
                      f"({f['so_id']})" for f in eligible}
            sel_f = st.selectbox("Delivered Fulfillment", list(labels.keys()),
                                 format_func=lambda k: labels[k], key="bl_f_sel")
            items = ful.get_fulfillment_items(sel_f)
            idf = pd.DataFrame(items)[["mat_code","mat_desc","uom","qty_shipped"]]
            idf.columns = ["Code","Description","UOM","Qty Shipped"]
            st.dataframe(idf, use_container_width=True, hide_index=True)

            due_days = st.number_input("Payment due in (days)", min_value=1, value=30, key="bl_due")
            notes = st.text_input("Notes", key="bl_notes")

            create_clicked = st.button("\U0001f9fe Create Invoice", type="primary", key="bl_create")
            if create_clicked:
                try:
                    result = bl.create_invoice(sel_f, due_days=due_days, notes=notes)
                    st.cache_data.clear()
                    st.success(f"\u2705 {result['invoice_id']} created — "
                              f"\u20b9{result['grand_total']:,.2f} ({result['tax_type']}, "
                              f"place of supply: {result['place_of_supply']}).")
                except Exception:
                    st.error("\u274c Error creating invoice:")
                    st.code(traceback.format_exc())

    # ── TAB 2 — manage invoices ────────────────────────────────────────────────────
    with tab2:
        invoices = bl.get_invoices()
        if not invoices:
            st.info("No invoices yet.")
        else:
            idf = pd.DataFrame(invoices)[["invoice_id","customer_name","status","grand_total",
                                           "invoice_date","due_date"]]
            idf.columns = ["Invoice","Customer","Status","Grand Total","Date","Due"]
            st.dataframe(idf, use_container_width=True, hide_index=True)

            labels = {i["invoice_id"]: f"{i['invoice_id']} — {i['customer_name']} ({i['status']})"
                      for i in invoices}
            sel = st.selectbox("Work on", list(labels.keys()), format_func=lambda k: labels[k], key="bl_work_sel")
            inv = next(i for i in invoices if i["invoice_id"] == sel)
            items = bl.get_invoice_items(sel)

            left, right = st.columns([3, 2], gap="large")
            with left:
                st.markdown("##### \U0001f4c4 Line Items")
                ldf = pd.DataFrame(items)[["mat_code","hsn_code","qty","unit_price",
                                            "taxable_value","gst_rate","line_total"]]
                ldf.columns = ["Code","HSN","Qty","Rate","Taxable","GST%","Total"]
                st.dataframe(ldf, use_container_width=True, hide_index=True)
                st.caption(f"Subtotal \u20b9{inv['subtotal']:,.2f}  \u00b7  "
                          f"CGST \u20b9{inv['cgst_total']:,.2f}  \u00b7  "
                          f"SGST \u20b9{inv['sgst_total']:,.2f}  \u00b7  "
                          f"IGST \u20b9{inv['igst_total']:,.2f}  \u00b7  "
                          f"**Grand Total \u20b9{inv['grand_total']:,.2f}**")

            with right:
                st.markdown("##### \u2696\ufe0f Actions")
                gen_clicked = st.button("\U0001f4c4 Generate Tax Invoice", key="bl_gen_doc")
                if gen_clicked:
                    fname, fbytes = bl.generate_invoice_document(sel)
                    st.session_state.bl_generated_doc = {"invoice_id": sel, "filename": fname, "bytes": fbytes}
                gd = st.session_state.get("bl_generated_doc")
                if gd and gd["invoice_id"] == sel:
                    st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                        file_name=gd["filename"], mime=XLSX_MIME, key=f"bl_dl_{sel}")

                if inv["status"] == "Draft":
                    if st.button("\U0001f4e4 Mark Issued", type="primary", key="bl_issue"):
                        bl.mark_issued(sel)
                        st.cache_data.clear()
                        st.success(f"{sel} issued.")
                        try:
                            je_id = acct.post_invoice_entry(sel)
                            st.success(f"\U0001f4d2 Posted to ledger as {je_id} "
                                      "(AR / Revenue / GST Output).")
                        except Exception:
                            st.warning("\u26a0\ufe0f Invoice issued, but the accounting "
                                      "entry failed to post:")
                            st.code(traceback.format_exc())
                        st.rerun()
                elif inv["status"] == "Issued":
                    st.success("\u2705 Issued — ready for payment once Cash Application exists.")

                if inv["status"] in ("Draft", "Issued"):
                    with st.popover("\u274c Cancel invoice"):
                        reason = st.text_input("Reason", key=f"bl_cancel_reason_{sel}")
                        if st.button("Confirm Cancellation", key=f"bl_cancel_btn_{sel}"):
                            bl.cancel_invoice(sel, reason)
                            st.cache_data.clear()
                            st.success(f"{sel} cancelled.")
                            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Accounting
# ══════════════════════════════════════════════════════════════════════════════
def page_accounting():
    st.markdown("## \U0001f4d2 Accounting")
    st.caption("Posted automatically: a goods receipt posts Inventory/GR-IR Clearing "
              "(ex-GST) when recorded, a Vendor Invoice clears GR/IR and posts GST "
              "Input/Accounts Payable when simulated, a customer invoice posts AR/"
              "Revenue/GST Output when marked Issued, a fulfillment posts COGS/"
              "Inventory Clearing when marked Delivered. Every entry is validated "
              "to balance before it's written — an unbalanced entry is refused, "
              "not flagged.")
    st.divider()

    s = acct.stats()
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Journal Entries", s["total"])
    m2.metric("From Goods Receipts", s["by_source"].get("GR", 0))
    m3.metric("From Vendor Invoices", s["by_source"].get("VendorInvoice", 0))
    m4.metric("From Invoices", s["by_source"].get("Invoice", 0))
    m5.metric("From Fulfillment", s["by_source"].get("Fulfillment", 0))
    st.divider()

    tab1, tab2, tab3 = st.tabs(["\U0001f4d2 Journal Entries", "\U0001f4ca Chart of Accounts",
                                "\U0001f9fe Vendor Invoices"])

    # ── TAB 1 — journal entries ────────────────────────────────────────────────────
    with tab1:
        entries = acct.get_journal_entries()
        if not entries:
            st.info("No journal entries yet — post one by issuing an invoice or "
                    "marking a fulfillment delivered.")
        else:
            edf = pd.DataFrame(entries)[["je_id","entry_date","source_type","source_id",
                                          "description","total_debit"]]
            edf.columns = ["JE","Date","Source Type","Source ID","Description","Amount"]
            st.dataframe(edf, use_container_width=True, hide_index=True)

            labels = {je["je_id"]: f"{je['je_id']} — {je['description']}" for je in entries}
            sel = st.selectbox("View entry", list(labels.keys()), format_func=lambda k: labels[k], key="acct_sel")
            lines = acct.get_journal_entry_lines(sel)
            ldf = pd.DataFrame(lines)[["account_code","account_name","debit","credit","description"]]
            ldf.columns = ["Code","Account","Debit","Credit","Description"]
            st.dataframe(ldf, use_container_width=True, hide_index=True)

            gen_clicked = st.button("\U0001f4c4 Generate Journal Voucher", key="acct_gen_doc")
            if gen_clicked:
                fname, fbytes = acct.generate_journal_voucher(sel)
                st.session_state.acct_generated_doc = {"je_id": sel, "filename": fname, "bytes": fbytes}
            gd = st.session_state.get("acct_generated_doc")
            if gd and gd["je_id"] == sel:
                st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                    file_name=gd["filename"], mime=XLSX_MIME, key=f"acct_dl_{sel}")

    # ── TAB 2 — chart of accounts + trial balance ────────────────────────────────
    with tab2:
        st.markdown("#### Trial Balance")
        tb = acct.get_trial_balance()
        tdf = pd.DataFrame(tb)[["account_code","account_name","account_type",
                                 "total_debit","total_credit","balance"]]
        tdf.columns = ["Code","Account","Type","Total Debit","Total Credit","Balance"]
        st.dataframe(tdf, use_container_width=True, hide_index=True)

        total_debit = sum(a["total_debit"] for a in tb)
        total_credit = sum(a["total_credit"] for a in tb)
        c1, c2 = st.columns(2)
        c1.metric("Total Debits", f"\u20b9{total_debit:,.2f}")
        c2.metric("Total Credits", f"\u20b9{total_credit:,.2f}")
        if round(total_debit, 2) == round(total_credit, 2):
            st.success("\u2705 Ledger balances.")
        else:
            st.error("\u274c Ledger does not balance — this shouldn't be possible "
                    "given post_journal_entry()'s validation; worth investigating.")

        st.caption("Inventory Clearing nets Goods Receipt debits (stock arriving) "
                  "against Fulfillment credits (stock leaving as COGS). "
                  "Production-confirmed stock (manufactured in-house, not "
                  "purchased) still isn't modeled here, so it stays untouched by "
                  "that path — see the account's own description below.")

        st.divider()
        st.markdown("#### \u2795 Add Account")
        c1, c2 = st.columns(2)
        with c1:
            new_code = st.text_input("Account Code", key="acct_new_code")
            new_name = st.text_input("Account Name", key="acct_new_name")
        with c2:
            new_type = st.selectbox("Account Type",
                ["Asset", "Liability", "Equity", "Revenue", "Expense"], key="acct_new_type")
            new_desc = st.text_input("Description", key="acct_new_desc")
        add_clicked = st.button("\u2795 Add Account", key="acct_add_btn")
        if add_clicked:
            if not new_code or not new_name:
                st.error("Account Code and Account Name are required.")
            else:
                try:
                    acct.add_account(new_code, new_name, new_type, new_desc)
                    st.success(f"\u2705 Added {new_code} — {new_name}.")
                    st.rerun()
                except Exception:
                    st.error("\u274c Error adding account:")
                    st.code(traceback.format_exc())

    # ── TAB 3 — vendor invoices ──────────────────────────────────────────────────
    with tab3:
        st.caption("Every Vendor Invoice ('Simulate Invoice' from the Goods Receipt "
                  "page), in one place — previously only visible one at a time by "
                  "reselecting the specific GR it came from.")
        vs = vi.stats()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Invoices", vs["total_invoices"])
        m2.metric("Open", vs["open"])
        m3.metric("Partially Paid", vs["partially_paid"])
        m4.metric("Paid", vs["paid"])

        show_all = st.toggle("Show paid invoices too", value=False, key="vi_show_all")
        invoices = vi.get_invoices()
        if not show_all:
            invoices = [v for v in invoices if v["status"] != "Paid"]

        if not invoices:
            st.info("No vendor invoices" + (" outstanding." if not show_all else " yet."))
        else:
            idf = pd.DataFrame(invoices)
            idf["balance_due"] = idf["amount"] - idf["paid_amount"]
            idf = idf[["invoice_id", "gr_id", "po_number", "vendor_name", "invoice_number",
                      "invoice_date", "amount", "paid_amount", "balance_due", "status"]]
            idf.columns = ["Invoice", "GR", "PO", "Vendor", "Vendor Ref", "Date",
                          "Amount", "Paid", "Due", "Status"]
            st.dataframe(idf, use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("##### \U0001f4b8 Record a Payment")
            payable = [v for v in invoices if v["status"] != "Paid"]
            if payable:
                labels = {v["invoice_id"]: f"{v['invoice_id']} — {v['vendor_name']} — "
                         f"\u20b9{v['amount']-v['paid_amount']:,.2f} due" for v in payable}
                sel = st.selectbox("Invoice", list(labels.keys()), format_func=lambda k: labels[k],
                                   key="vi_pay_sel")
                info = vi.get_invoice_payment_info(sel)
                pay_amt = st.number_input("Payment amount", min_value=0.0,
                    max_value=float(info["balance_due"]), value=float(info["balance_due"]),
                    key="vi_pay_amt", help="Defaults to the full balance due.")
                if st.button("\U0001f4b8 Record Payment", type="primary", key="vi_pay_btn"):
                    try:
                        result = vi.record_payment(sel, pay_amt)
                        st.cache_data.clear()
                        st.success(f"\u2705 {result['payment_id']} recorded, posted as {result['je_id']}.")
                        st.rerun()
                    except Exception:
                        st.error("\u274c Error:"); st.code(traceback.format_exc())
            else:
                st.success("\u2705 Nothing outstanding.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE — Cash Application
# ══════════════════════════════════════════════════════════════════════════════
def page_cash_application():
    st.markdown("## \U0001f4b5 Cash Application")
    st.caption("Record cash first, categorize it second — a payment gets logged "
               "the moment it arrives, then matched to specific invoice(s) as a "
               "separate step, same as everywhere else in this app. Short payments "
               "(routine in India due to TDS) are tracked with a reason, not "
               "silently forced to zero.")
    st.divider()

    s = ca.stats()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Payments Received", s["total_payments"])
    m2.metric("Total Received", f"\u20b9{s['total_received']:,.0f}")
    m3.metric("Unapplied", f"\u20b9{s['total_unapplied']:,.0f}")
    m4.metric("Overdue Invoices", s["overdue_count"], delta=f"\u20b9{s['overdue_amount']:,.0f}" if s["overdue_amount"] else None)
    st.divider()

    tab1, tab2, tab3 = st.tabs(["\u2795 Record Payment", "\U0001f517 Apply Payments", "\U0001f4cb Manage & Reports"])

    # ── TAB 1 — record a payment ──────────────────────────────────────────────────
    with tab1:
        customers = co.list_customers()
        if not customers:
            st.info("No customers yet.")
        else:
            labels = {c["Customer_ID"]: f"{c['Customer_ID']} — {c['Customer_Name']}" for c in customers}
            sel_cust = st.selectbox("Customer", list(labels.keys()), format_func=lambda k: labels[k], key="ca_rec_cust")

            c1, c2 = st.columns(2)
            with c1:
                amount = st.number_input("Amount Received (\u20b9)", min_value=0.0, step=100.0, key="ca_rec_amt")
                pay_date = st.date_input("Payment Date", value=date.today(), key="ca_rec_date")
            with c2:
                method = st.selectbox("Payment Method", ca.PAYMENT_METHODS, key="ca_rec_method")
                ref_no = st.text_input("Reference No. (UTR / cheque no. / txn ID)", key="ca_rec_ref")
            notes = st.text_input("Notes", key="ca_rec_notes")

            record_clicked = st.button("\U0001f4b5 Record Payment", type="primary", key="ca_rec_btn")
            if record_clicked:
                if amount <= 0:
                    st.error("Enter an amount greater than zero.")
                else:
                    try:
                        pid = ca.record_payment(sel_cust, amount, pay_date, method, ref_no, notes)
                        st.cache_data.clear()
                        st.success(f"\u2705 {pid} recorded — \u20b9{amount:,.2f} from "
                                  f"{labels[sel_cust]}. Head to **Apply Payments** to match it.")
                    except Exception:
                        st.error("\u274c Error recording payment:")
                        st.code(traceback.format_exc())

    # ── TAB 2 — apply payments to invoices, or recognize as advance ─────────────────
    with tab2:
        unapplied_payments = [p for p in ca.get_payments() if p["unapplied_amount"] > 0]
        if not unapplied_payments:
            st.info("No payments with an unapplied balance right now.")
        else:
            labels = {p["payment_id"]: f"{p['payment_id']} — {p['customer_name']} "
                      f"(\u20b9{p['unapplied_amount']:,.2f} unapplied)" for p in unapplied_payments}
            sel_pmt = st.selectbox("Payment", list(labels.keys()), format_func=lambda k: labels[k], key="ca_app_pmt")
            payment = next(p for p in unapplied_payments if p["payment_id"] == sel_pmt)
            st.caption(f"Unapplied balance: \u20b9{payment['unapplied_amount']:,.2f}")

            open_invoices = ca.get_customer_open_invoices(payment["customer_id"])
            if open_invoices:
                st.markdown("##### Apply to an invoice")
                idf = pd.DataFrame(open_invoices)[["invoice_id","grand_total","paid_amount","balance_due"]]
                idf.columns = ["Invoice","Total","Paid So Far","Balance Due"]
                st.dataframe(idf, use_container_width=True, hide_index=True)

                inv_labels = {i["invoice_id"]: f"{i['invoice_id']} — balance \u20b9{i['balance_due']:,.2f}"
                             for i in open_invoices}
                sel_inv = st.selectbox("Invoice", list(inv_labels.keys()),
                                       format_func=lambda k: inv_labels[k], key="ca_app_inv")
                inv_row = next(i for i in open_invoices if i["invoice_id"] == sel_inv)
                default_amt = min(payment["unapplied_amount"], inv_row["balance_due"])

                c1, c2 = st.columns(2)
                with c1:
                    apply_amt = st.number_input("Amount to apply (\u20b9)", min_value=0.0,
                        value=float(default_amt), key="ca_app_amt")
                with c2:
                    reason = st.selectbox("Short payment reason, if any", ca.SHORT_PAYMENT_REASONS, key="ca_app_reason")

                apply_clicked = st.button("\U0001f517 Apply to Invoice", type="primary", key="ca_app_btn")
                if apply_clicked:
                    try:
                        je_id = ca.apply_payment(sel_pmt, sel_inv, apply_amt, reason)
                        st.cache_data.clear()
                        st.success(f"\u2705 Applied \u20b9{apply_amt:,.2f} to {sel_inv} — posted as {je_id}.")
                        st.rerun()
                    except Exception:
                        st.error("\u274c Error applying payment:")
                        st.code(traceback.format_exc())
            else:
                st.caption("This customer has no open invoices to apply against.")

            st.divider()
            st.markdown("##### Or recognize the remainder as a customer advance")
            adv_amt = st.number_input("Advance amount (\u20b9)", min_value=0.0,
                value=float(payment["unapplied_amount"]), key="ca_adv_amt")
            advance_clicked = st.button("\U0001f4b0 Record as Advance", key="ca_adv_btn")
            if advance_clicked:
                try:
                    je_id = ca.record_as_advance(sel_pmt, adv_amt)
                    st.cache_data.clear()
                    st.success(f"\u2705 \u20b9{adv_amt:,.2f} recognized as a customer advance — posted as {je_id}.")
                    st.rerun()
                except Exception:
                    st.error("\u274c Error recording advance:")
                    st.code(traceback.format_exc())

    # ── TAB 3 — manage payments + collections worklist ───────────────────────────
    with tab3:
        payments = ca.get_payments()
        if payments:
            st.markdown("##### \U0001f4b5 Payments")
            pdf = pd.DataFrame(payments)[["payment_id","customer_name","amount",
                                           "unapplied_amount","payment_date","reference_no"]]
            pdf.columns = ["Payment","Customer","Amount","Unapplied","Date","Reference"]
            st.dataframe(pdf, use_container_width=True, hide_index=True)

            labels = {p["payment_id"]: f"{p['payment_id']} — {p['customer_name']}" for p in payments}
            sel = st.selectbox("View payment", list(labels.keys()), format_func=lambda k: labels[k], key="ca_view_sel")
            apps = ca.get_payment_applications(payment_id=sel)
            if apps:
                adf = pd.DataFrame(apps)[["invoice_id","applied_amount","short_payment_reason","application_date"]]
                adf.columns = ["Applied To","Amount","Reason","Date"]
                st.dataframe(adf, use_container_width=True, hide_index=True)

            gen_clicked = st.button("\U0001f4c4 Generate Payment Receipt", key="ca_gen_doc")
            if gen_clicked:
                fname, fbytes = ca.generate_payment_receipt(sel)
                st.session_state.ca_generated_doc = {"payment_id": sel, "filename": fname, "bytes": fbytes}
            gd = st.session_state.get("ca_generated_doc")
            if gd and gd["payment_id"] == sel:
                st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                    file_name=gd["filename"], mime=XLSX_MIME, key=f"ca_dl_{sel}")
        else:
            st.info("No payments recorded yet.")

        st.divider()
        st.markdown("##### \u23f0 Overdue Invoices (Collections Worklist)")
        overdue = ca.get_overdue_invoices()
        if overdue:
            odf = pd.DataFrame(overdue)[["invoice_id","customer_name","due_date","balance_due"]]
            odf.columns = ["Invoice","Customer","Due Date","Balance Due"]
            st.dataframe(odf, use_container_width=True, hide_index=True)
        else:
            st.success("\u2705 Nothing overdue right now.")


# ── Router ─────────────────────────────────────────────────────────────────────
if "Customer Onboarding" in page:
    page_customer_onboarding()
elif "Quotation" in page:
    page_quotation()
elif "Sales Orders" in page:
    page_sales_orders()
elif "Fulfillment" in page:
    page_fulfillment()
elif "Billing" in page:
    page_billing()
elif "Cash Application" in page:
    page_cash_application()
elif "Accounting" in page:
    page_accounting()
elif "Settings" in page:
    page_settings()
