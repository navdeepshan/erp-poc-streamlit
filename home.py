"""Executive landing page for the unified ERP suite."""

import streamlit as st

import customer_onboarding as customers
import goods_receipt as receipts
import pr_consolidation as procurement
import vendor_onboarding as vendors


def safe(call, fallback):
    try:
        return call()
    except Exception:
        return fallback


vendor_stats = safe(vendors.stats, {"total": 0, "approved": 0})
customer_stats = safe(customers.stats, {"total": 0, "approved": 0})
receipt_stats = safe(receipts.stats, {"total": 0, "by_status": {}})
open_prs = safe(lambda: procurement.load_open_prs(), [])

st.markdown('<div class="erp-eyebrow">COMMAND CENTER</div>', unsafe_allow_html=True)
st.title("Good morning. Here’s your operation at a glance.")
st.caption("Live process visibility across procurement, manufacturing, fulfillment, and finance.")

cols = st.columns(4)
cols[0].metric("Open purchase requests", len(open_prs), help="Purchase requisitions awaiting downstream action")
cols[1].metric("Approved vendors", vendor_stats.get("approved", 0))
cols[2].metric("Active customers", customer_stats.get("approved", 0))
cols[3].metric("Posted receipts", receipt_stats.get("by_status", {}).get("Posted", 0))

st.markdown("### Workspaces")
st.caption("Move between end-to-end business processes without leaving the suite.")

c1, c2, c3 = st.columns(3)
with c1:
    with st.container(border=True):
        st.markdown("#### Source to Pay")
        st.write("Purchase bundles, requisitions, purchase orders, sourcing, vendors, and contracts.")
        st.page_link("erp_ui.py", label="Open procurement", icon=":material/arrow_forward:")
with c2:
    with st.container(border=True):
        st.markdown("#### Manufacturing")
        st.write("Goods receipt, quality, BOM planning, production confirmations, and inventory.")
        st.page_link("mfg_ui.py", label="Open manufacturing", icon=":material/arrow_forward:")
with c3:
    with st.container(border=True):
        st.markdown("#### Order to Cash")
        st.write("Customers, quotes, orders, fulfillment, invoicing, cash application, and accounting.")
        st.page_link("o2c_ui.py", label="Open order to cash", icon=":material/arrow_forward:")

st.markdown("### AI operations")
with st.container(border=True):
    ai, action = st.columns([4, 1])
    with ai:
        st.markdown("#### Ask, review, approve")
        st.write("Use the conversational workspace to investigate inventory, upload orders, and approve proposed actions with a human in control.")
    with action:
        st.page_link("agent_console.py", label="Open AI workspace", icon=":material/auto_awesome:", width="stretch")

st.info("This is a local demonstration environment.", icon=":material/info:")
