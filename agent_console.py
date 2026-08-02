"""
agent_console.py — Agent Console: a standalone conversational surface
demonstrating the agentic interaction model (human on the loop, not in
the loop) alongside the traditional erp_ui.py / mfg_ui.py / o2c_ui.py
screens, not merged into any of them.

Explicitly scoped to inventory optimization + order upload, per
direct steer, rather than trying to cover the whole platform. Every
action here calls the exact same backend functions the traditional
screens already call (bom.py, inventory.py, pr_consolidation.py,
sales_order_import.py, org_defaults.py) against the same shared
database — nothing approved here is a simulation of an action, it IS
the action, immediately visible in the traditional apps too.

The "conversational" layer (agent_intents.py) is rule-based pattern
matching against a curated vocabulary, not a general-purpose language
model — this app is honest about that rather than implying more than
it delivers: a visible "here's what I can help with" sits alongside
the free-text box, and genuinely unmatched input says so plainly
rather than guessing.

Nothing here ever executes an action without an explicit human
approval first — every proposal this generates is a real, structured
pending_action in session state, shown as its own confirmation card,
requiring an explicit Approve before the underlying write function is
ever called. This is the same discipline every other proposal in this
platform already follows (Ship, Create PR, RFx award); this app makes
that discipline the whole point of the surface, not an implementation
detail buried in a page.
"""
import streamlit as st
from datetime import date, datetime
import os
import traceback

import bom
import po_export
import inventory as inv
import pr_consolidation as pc
import org_defaults as od
import sales_order_import as soi
import agent_intents as ai
import demo_profiles as dp
import org_profile as op
import item_tax
import importlib
import seed_manager as sm
import sales_order as sales
import billing
import cash_application as cash
import goods_receipt as receipts
import vendor_onboarding as vendors
import customer_onboarding as customers
import purchase_bundles as bundles
from ui_theme import apply_theme

# Page configuration and the shared theme are owned by streamlit_app.py.

_DIR = os.path.dirname(os.path.abspath(__file__))

STATIC_PROMPTS = [
    "Show executive summary",
    "Show sales report",
    "Show inventory report",
    "Show finance and receivables report",
    "Show procurement report",
    "What needs action right now?",
]


def build_example_prompts(data_file=None):
    """
    Builds the ship/explain/create-PR example prompts from whatever
    Item Master and transfer/recommendation data is REALLY loaded right
    now, rather than fixed strings. Real bug found 2026-07-30: the
    original hardcoded examples ("ship the scaler to Bangalore" etc.)
    were IDS Denmed-specific — with Genrobotics data loaded instead,
    every one of those materials and locations is simply absent, so
    even the suggested buttons failed material resolution with the
    exact same "I couldn't find a material matching that" a person
    would get from a genuinely bad free-text guess. Confirmed directly
    against real Genrobotics data before concluding this was the root
    cause, not assumed. This app is meant to work against whichever
    pilot dataset happens to be loaded, so its own suggested prompts
    need to as well.

    Prefers a material with a REAL live shortfall (so the resulting
    proposal is meaningful, not just resolvable) but falls back to the
    first two active Item Master entries if nothing is currently
    flagged — either way, every generated prompt references a real
    name that WILL resolve, even in the fallback case where the
    resulting action then honestly reports no real opportunity exists,
    which is a world away from failing at entity resolution itself.
    """
    items = po_export.load_item_master(data_file, active_only=True)
    locs = pc.get_delivery_locations(active_only=True)
    if not items or not locs:
        return STATIC_PROMPTS

    recs = bom.get_procurement_recommendations(data_file)
    at_risk_or_needed = [r for r in recs if r["outcome"] in ("At Risk", "Action Needed")]
    action_needed = [r for r in recs if r["outcome"] == "Action Needed"]
    opps = bom.get_transfer_opportunities(data_file)

    prompts = list(STATIC_PROMPTS)

    if opps:
        o = opps[0]
        mat_name = _short_name(o["mat_desc"])
        loc_name = _short_name(_lname(o["to_location"]))
        prompts.insert(1, f"Ship the {mat_name} to {loc_name}")
    else:
        mat_name = _short_name(items[0]["desc"])
        loc_name = _short_name(locs[-1]["name"] if len(locs) > 1 else locs[0]["name"])
        prompts.insert(1, f"Ship the {mat_name} to {loc_name}")

    if at_risk_or_needed:
        mat_name = _short_name(at_risk_or_needed[0]["mat_desc"])
        prompts.insert(2, f"Why is the {mat_name} flagged?")
    elif len(items) > 1:
        prompts.insert(2, f"Why is the {_short_name(items[1]['desc'])} flagged?")

    if action_needed:
        mat_name = _short_name(action_needed[0]["mat_desc"])
        prompts.append(f"Create a requisition for the {mat_name}")
    elif len(items) > 2:
        prompts.append(f"Create a requisition for the {_short_name(items[2]['desc'])}")

    return prompts


def _short_name(desc):
    """A short, natural-sounding, distinctive phrase from a material/
    location description — not the whole formal name, but enough
    words to actually be specific. Taking only the single last word
    (an earlier version of this function) turned out to be too
    generic in practice — "24V" or "meter" matched several real items
    at once, trading the original "couldn't find a material" failure
    for a much higher rate of unnecessary clarification prompts.
    Two non-generic words, in their original order, does noticeably
    better without needing per-item overrides."""
    skip = {"woodpecker", "melag", "genrobotics", "office", "hub", "plant",
            "and", "hq", "-", "the", "unit"}
    words = [w for w in desc.replace("-", " ").split() if w.lower() not in skip]
    if len(words) >= 2:
        return " ".join(words[:2])
    return words[0] if words else desc.split()[0]


def _init_state():
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = [
            {"role": "assistant", "content": (
                "Hi — I'm the Agent Console for inventory optimization and order "
                "upload. I work from a curated set of things I understand (see the "
                "examples below), not open-ended conversation — if I don't recognize "
                "something, I'll say so rather than guess. Nothing I propose executes "
                "until you approve it."
            )},
        ]
    if "pending_action" not in st.session_state:
        st.session_state.pending_action = None
    if "clarify" not in st.session_state:
        st.session_state.clarify = None


def _say(content):
    st.session_state.chat_history.append({"role": "assistant", "content": content})


def _user_said(content):
    st.session_state.chat_history.append({"role": "user", "content": content})


def _friendly_failure(context="that request"):
    """Keep technical failures out of the end-user conversation."""
    _say(f"I couldn't complete {context} just now. Your data is unchanged. "
         "Please try again or choose one of the reporting prompts below.")


def _notify_shell_refresh():
    """
    Sets a flag that gets rendered as a real postMessage script at the
    TOP of the next script run, not here, directly — a real bug found
    by testing, not by inspection: calling st.components.v1.html()
    here and then immediately calling st.rerun() (which every approval
    branch already does) wipes the component out before its <script>
    tag ever gets a chance to actually execute in the browser.
    Confirmed directly: zero iframes existed in the page's DOM after
    approving a real action, meaning the component genuinely never
    rendered, not just that its message didn't arrive. Deferring to
    the next run — after the rerun has already happened — gives the
    browser real time to load and execute the script before Streamlit
    reruns again.
    """
    st.session_state.pending_refresh_notify = True


def _render_pending_refresh_notify():
    """
    Actually renders the postMessage script, called once at the very
    top of the script on every run. See window.top vs window.parent
    and the height=0 reasoning in _notify_shell_refresh()'s original
    docstring content, preserved here since this is where the render
    now actually happens.
    """
    if st.session_state.get("pending_refresh_notify"):
        st.session_state.pending_refresh_notify = False
        st.components.v1.html(
            "<script>window.top.postMessage({type: 'erp_action_approved'}, '*');</script>",
            height=0,
        )


def _lname(loc_id, locs_cache={}):
    if not locs_cache:
        for l in pc.get_delivery_locations(active_only=False):
            locs_cache[l["id"]] = l["name"]
    return locs_cache.get(loc_id, loc_id)


# ── Read-only handlers (no confirmation needed) ──────────────────────────────
def handle_check_at_risk():
    recs = bom.get_procurement_recommendations()
    at_risk = [r for r in recs if r["outcome"] == "At Risk"]
    action_needed = [r for r in recs if r["outcome"] == "Action Needed"]
    covered = [r for r in recs if r["outcome"] == "Already Covered by Existing PR/PO"]
    if not recs:
        _say("Nothing needs attention right now — every current shortfall is "
            "either fully covered by a transfer or already in hand.")
        return
    lines = [f"**{len(at_risk)} At Risk**, **{len(action_needed)} Action Needed**, "
            f"**{len(covered)} Already Covered**."]
    if at_risk:
        lines.append("\n**At Risk** (no PR proposed — not enough time by normal means):")
        for r in at_risk[:5]:
            lines.append(f"- {r['mat_desc']} at {_lname(r['location'])} — "
                        f"gap {r['remaining_gap']:g}, stock-out {r['stockout_date']}")
    if action_needed:
        lines.append("\n**Action Needed**:")
        for r in action_needed[:5]:
            lines.append(f"- {r['mat_desc']} at {_lname(r['location'])} — "
                        f"need {r['recommended_qty']:g} by {r['required_by_date']}")
    _say("\n".join(lines))


def _money(value):
    return f"₹{float(value or 0):,.2f}"


def handle_reporting(text):
    """Read-only seeded-data reports with stable, presentation-ready output."""
    lower = text.lower()
    sales_stats = sales.stats()
    inventory_stats = inv.stats()
    invoice_stats = billing.stats()
    cash_stats = cash.stats()
    receipt_stats = receipts.stats()
    vendor_stats = vendors.stats()
    customer_stats = customers.stats()
    bundle_stats = bundles.stats()
    bom_stats = bom.stats()
    item_count = len(po_export.load_item_master(active_only=True))
    pr_rows = pc.get_pr_report()
    open_prs = [r for r in pr_rows if r.get("Status") == "Open"]
    po_count = len(pc.get_all_po_headers())

    sales_lines = [
        "### Sales report",
        f"- **{sales_stats['total']}** sales orders worth **{_money(sales_stats['total_value'])}**",
        f"- Confirmed order value: **{_money(sales_stats['confirmed_value'])}**",
        f"- **{customer_stats['approved']}** active customers",
    ]
    inventory_lines = [
        "### Inventory report",
        f"- **{inventory_stats['materials_tracked']}** materials across **{inventory_stats['locations_tracked']}** locations",
        f"- Units on hand: **{inventory_stats['total_units_on_hand']:,.0f}**",
        f"- Units in transit: **{inventory_stats['total_units_in_transit']:,.0f}**",
        f"- Negative balances: **{inventory_stats['negative_balances']}**",
    ]
    finance_lines = [
        "### Finance & receivables",
        f"- **{invoice_stats['total']}** invoices worth **{_money(invoice_stats['total_value'])}**",
        f"- Cash received: **{_money(cash_stats['total_received'])}**",
        f"- Overdue: **{cash_stats['overdue_count']}** invoices / **{_money(cash_stats['overdue_amount'])}**",
        f"- Unapplied cash: **{_money(cash_stats['total_unapplied'])}**",
    ]
    procurement_lines = [
        "### Procurement report",
        f"- **{len(open_prs)}** open PR lines and **{po_count}** purchase orders",
        f"- **{vendor_stats['approved']}** approved vendors",
        f"- **{receipt_stats['by_status'].get('Posted', 0)}** posted goods receipts",
        f"- **{bundle_stats['active_bundles']}** active purchase bundles with **{bundle_stats['total_lines']}** seeded lines",
    ]

    master_lines = [
        "### Seeded master data",
        f"- **{item_count}** active materials and **{vendor_stats['approved']}** approved vendors",
        f"- **{customer_stats['approved']}** active customers",
        f"- **{bom_stats['finished_goods']}** finished goods with **{bom_stats['bom_lines']}** BOM component lines",
        f"- **{bundle_stats['active_bundles']}** active procurement bundles",
    ]

    if "sales" in lower:
        lines = sales_lines
    elif "inventory" in lower:
        lines = inventory_lines
    elif any(k in lower for k in ("finance", "financial", "receivable", "cash")):
        lines = finance_lines
    elif "procurement" in lower:
        lines = procurement_lines
    else:
        lines = ["## Executive reporting snapshot"] + master_lines[1:] + [""] + \
                sales_lines[1:] + [""] + inventory_lines[1:] + [""] + \
                finance_lines[1:] + [""] + procurement_lines[1:]
    _say("\n".join(lines) + "\n\n*Live from the seeded ERP dataset.*")


def handle_explain(text):
    resolved = ai.resolve_material(text)
    if resolved["candidates"]:
        st.session_state.clarify = {"kind": "material_for_explain",
                                    "candidates": resolved["candidates"], "raw_text": text}
        names = ", ".join(f"**{c['desc']}**" for c in resolved["candidates"])
        _say(f"I found more than one match — did you mean {names}?")
        return
    if not resolved["match"]:
        _say("I couldn't find a material matching that in the current Item Master.")
        return
    _explain_material(resolved["match"])


def _explain_material(item):
    rec = ai.find_recommendation(item["code"])
    if not rec:
        _say(f"**{item['desc']}** isn't currently flagged in any recommendation — "
            f"no unresolved shortfall against it right now.")
        return
    if rec["outcome"] == "At Risk":
        _say(f"**{item['desc']}** at {_lname(rec['location'])} is **At Risk**: "
            f"projected to run out on {rec['stockout_date']} ({rec['days_until_stockout']} "
            f"days out), but normal procurement needs {rec['pipeline_lead_time_days']} days "
            f"({rec['lead_time_source'].replace('_', ' ')}) — not enough time, so "
            f"deliberately no PR is proposed for it. This needs a human decision "
            f"(expediting, an alternate vendor, or a customer conversation), not a "
            f"system-generated document that would arrive late anyway.")
    elif rec["outcome"] == "Action Needed":
        _say(f"**{item['desc']}** at {_lname(rec['location'])} needs "
            f"**{rec['recommended_qty']:g}** units by **{rec['required_by_date']}** — "
            f"{rec['days_until_stockout']} days out, with a {rec['pipeline_lead_time_days']}-day "
            f"lead time, so there's real headroom for a timed requisition.")
    else:
        _say(f"**{item['desc']}** at {_lname(rec['location'])} has a real gap, but it's "
            f"**already covered** by Open PR {rec.get('covering_pr', '')} — no duplicate "
            f"needed.")


# ── Actionable handlers (propose, require confirmation) ─────────────────────
def handle_ship(text):
    mat = ai.resolve_material(text)
    if mat["candidates"]:
        st.session_state.clarify = {"kind": "material_for_ship",
                                    "candidates": mat["candidates"], "raw_text": text}
        names = ", ".join(f"**{c['desc']}**" for c in mat["candidates"])
        _say(f"Which material did you mean — {names}?")
        return
    if not mat["match"]:
        _say("I couldn't find a material matching that. Try naming it more "
            "specifically, or use the quick actions below.")
        return
    loc = ai.resolve_location(text)
    if loc["candidates"]:
        st.session_state.clarify = {"kind": "location_for_ship",
                                    "candidates": loc["candidates"], "raw_text": text,
                                    "mat_code": mat["match"]["code"]}
        names = ", ".join(f"**{c['name']}**" for c in loc["candidates"])
        _say(f"Which location did you mean — {names}?")
        return
    if not loc["match"]:
        _say(f"I found **{mat['match']['desc']}**, but couldn't tell which "
            f"destination you meant. Try naming the city or location.")
        return
    _propose_ship(mat["match"]["code"], mat["match"]["desc"], loc["match"]["id"], text)


def _propose_ship(mat_code, mat_desc, to_location, raw_text):
    opp = ai.find_transfer_opportunity(mat_code, to_location)
    if not opp:
        _say(f"I don't see a real transfer opportunity for **{mat_desc}** to "
            f"**{_lname(to_location)}** right now — nothing shows a shortfall "
            f"there that another location has surplus to cover.")
        return
    qty = ai.extract_quantity(raw_text) or opp["suggested_qty"]
    qty = min(qty, opp["available_at_source"])
    st.session_state.pending_action = {
        "type": "ship", "mat_code": mat_code, "mat_desc": mat_desc,
        "from_location": opp["from_location"], "to_location": to_location, "qty": qty,
        "summary": (f"Ship **{qty:g}** units of **{mat_desc}** from "
                    f"**{_lname(opp['from_location'])}** to **{_lname(to_location)}**?"),
    }
    _say(st.session_state.pending_action["summary"])


def handle_receive(text):
    mat = ai.resolve_material(text)
    if mat["candidates"]:
        st.session_state.clarify = {"kind": "material_for_receive",
                                    "candidates": mat["candidates"], "raw_text": text}
        names = ", ".join(f"**{c['desc']}**" for c in mat["candidates"])
        _say(f"Which material did you mean — {names}?")
        return
    if not mat["match"]:
        _say("I couldn't find a material matching that. Try naming it more "
            "specifically, or use the quick actions below.")
        return
    # Location is optional for receiving (unlike shipping) — if the text
    # doesn't resolve one, or resolves to something ambiguous, that's
    # fine as long as find_receivable_transfer can narrow it down on
    # material alone.
    loc = ai.resolve_location(text)
    to_location = loc["match"]["id"] if loc["match"] else None
    _propose_receive(mat["match"]["code"], mat["match"]["desc"], to_location, text)


def _propose_receive(mat_code, mat_desc, to_location, raw_text):
    matches = ai.find_receivable_transfer(mat_code, to_location)
    if not matches:
        where = f" to **{_lname(to_location)}**" if to_location else ""
        _say(f"I don't see **{mat_desc}** currently In Transit{where} — "
            f"nothing real to confirm receipt of right now.")
        return
    if len(matches) > 1:
        st.session_state.clarify = {
            "kind": "transfer_for_receive", "raw_text": raw_text,
            "candidates": [{"name": f"{mat_desc} \u2192 {_lname(t['to_location'])} "
                                    f"({t['quantity']:g} units)", **t} for t in matches],
        }
        _say(f"**{mat_desc}** has more than one shipment In Transit right now — "
            f"which one?")
        return
    t = matches[0]
    st.session_state.pending_action = {
        "type": "receive", "transfer_id": t["transfer_id"],
        "summary": (f"Confirm receipt of **{t['quantity']:g}** units of "
                    f"**{mat_desc}** at **{_lname(t['to_location'])}** "
                    f"(shipped from **{_lname(t['from_location'])}**)?"),
    }
    _say(st.session_state.pending_action["summary"])


def handle_create_pr(text):
    mat = ai.resolve_material(text)
    if mat["candidates"]:
        st.session_state.clarify = {"kind": "material_for_pr",
                                    "candidates": mat["candidates"], "raw_text": text}
        names = ", ".join(f"**{c['desc']}**" for c in mat["candidates"])
        _say(f"Which material did you mean — {names}?")
        return
    if not mat["match"]:
        _say("I couldn't find a material matching that.")
        return
    _propose_pr(mat["match"]["code"], mat["match"]["desc"])


def _propose_pr(mat_code, mat_desc):
    rec = ai.find_recommendation(mat_code)
    if not rec:
        _say(f"**{mat_desc}** doesn't have an unresolved shortfall right now — "
            f"nothing to raise a requisition against.")
        return
    if rec["outcome"] == "At Risk":
        _say(f"**{mat_desc}** is At Risk — a stock-out on {rec['stockout_date']} against "
            f"a {rec['pipeline_lead_time_days']}-day lead time. A PR raised today still "
            f"wouldn't arrive in time, so I won't propose one — this needs a human "
            f"decision (expediting, alternate sourcing), not a document that would "
            f"just arrive late.")
        return
    if rec["outcome"] == "Already Covered by Existing PR/PO":
        _say(f"**{mat_desc}** is already covered by **{rec.get('covering_pr', '')}** "
            f"— raising another would be a duplicate.")
        return
    st.session_state.pending_action = {
        "type": "create_pr", "mat_code": mat_code, "mat_desc": mat_desc,
        "location": rec["location"], "qty": rec["recommended_qty"],
        "required_by": rec["required_by_date"],
        "summary": (f"Create a PR for **{rec['recommended_qty']:g}** units of "
                    f"**{mat_desc}** at **{_lname(rec['location'])}**, required by "
                    f"**{rec['required_by_date']}**?"),
    }
    _say(st.session_state.pending_action["summary"])


def handle_switch_mode(text):
    mode = ai.resolve_mode(text)
    if not mode:
        _say("I couldn't tell which planning mode you meant — the three are "
            "*Sales Order Based*, *Optimize Existing PRs*, and *Reorder Qty Based*.")
        return
    current = od.get_default("Time-Phased Planning Mode")
    if mode == current:
        _say(f"**{mode}** is already the active mode.")
        return
    st.session_state.pending_action = {
        "type": "switch_mode", "mode": mode,
        "summary": f"Switch the Time-Phased Planning Mode to **{mode}**?",
    }
    _say(st.session_state.pending_action["summary"])


def handle_upload_orders():
    _say("You can bulk-upload Sales Orders here — download the template below, "
        "fill it in, and upload it back. This is the same import path as the "
        "Sales Orders → Bulk Import screen, so anything you upload here shows up "
        "there too, and vice versa.")
    st.session_state.pending_action = {"type": "upload_orders"}


# ── Settings / data-load handlers ────────────────────────────────────────────
def _propose_demo_action(action_type, label):
    """Shared proposal shape for the four simple demo-scenario actions
    (setup, complete rest, load full demo, add customer wave) — each
    calls into whichever demo_scenario module matches the currently
    active pilot (demo_profiles.py), not a hardcoded one, so this
    console behaves correctly regardless of which org profile is
    loaded."""
    profile, org_id = dp.current_profile()
    if not profile:
        _say(f"I can't tell which demo scenario is active — the current "
            f"organization ('{org_id or 'not set'}') doesn't match a known "
            f"demo profile. This needs the traditional Settings screen.")
        return
    st.session_state.pending_action = {
        "type": action_type, "module": profile["module"], "summary": f"{label}?",
    }
    _say(st.session_state.pending_action["summary"])


def handle_run_setup():
    _propose_demo_action("run_setup", "Load Setup Only for this demo scenario")


def handle_complete_rest():
    _propose_demo_action("complete_rest", "Complete the Rest of this demo scenario")


def handle_load_full_demo():
    _propose_demo_action("load_full_demo",
        "Load the Full Demo (setup and resolution together, no live transfer moment)")


def handle_add_customer_wave():
    profile, org_id = dp.current_profile()
    if not profile or not profile.get("supports_customer_wave"):
        _say(f"The new customer wave isn't available for the current "
            f"organization ('{org_id or 'not set'}') — this needs a demo "
            f"profile that supports it.")
        return
    # Real bug fixed here, found while testing this on IDS Denmed for
    # the first time: this used to hardcode "Surat, Guwahati, Delhi Jal
    # Board" (Genrobotics' own wave) into the proposal text regardless
    # of which pilot was actually active — wrong on IDS, where the
    # wave adds SMILEZONE-AMD/KOLKATADENTAL-CCU/PINKCITY-JAI instead.
    # Nothing about which customers get added is known until the
    # function actually runs, so the proposal stays generic; the
    # confirmation message after Approve (which already used the real
    # result) is where the actual names belong.
    _propose_demo_action("add_customer_wave", "Add the new customer wave")


def handle_reset_data():
    _say("Resetting wipes all current data and reloads from a seed file — "
        "upload the correct .xlsx for the pilot you want below. This is "
        "destructive, so it still needs your explicit confirmation, same as "
        "everything else here.")
    st.session_state.pending_action = {"type": "reset_data"}


def handle_query_org_profile():
    profile = op.get_org_profile()
    if not profile:
        _say("No organization profile is configured yet.")
        return
    _say(f"**{profile.get('Legal_Name', '(unnamed)')}** — GSTIN "
        f"**{profile.get('GSTIN', 'not set')}**, PAN "
        f"**{profile.get('PAN', 'not set')}**, {profile.get('City', '')}, "
        f"{profile.get('State', '')}.")


def handle_query_item_tax(text):
    resolved = ai.resolve_material(text)
    if resolved["candidates"]:
        names = ", ".join(f"**{c['desc']}**" for c in resolved["candidates"])
        _say(f"Which material did you mean — {names}?")
        return
    if resolved["match"]:
        info = item_tax.get_item_tax_info(resolved["match"]["code"])
        if info and info.get("hsn_code"):
            _say(f"**{resolved['match']['desc']}** — HSN **{info['hsn_code']}**, "
                f"GST **{info['gst_rate']:g}%**.")
        else:
            _say(f"**{resolved['match']['desc']}** has no tax info on file yet.")
        return
    missing = item_tax.list_items_missing_tax_info()
    if not missing:
        _say("Every active item has tax info on file.")
    else:
        names = ", ".join(m["mat_desc"] for m in missing[:8])
        _say(f"**{len(missing)} item(s)** are missing tax info: {names}" +
            ("..." if len(missing) > 8 else ""))


def handle_help():
    _say("Here's what I can help with right now:\n\n" +
        "\n".join(f"- {p}" for p in build_example_prompts()) +
        "\n\nAnything outside this I'll tell you plainly I didn't catch, rather "
        "than guess.")


# ── Router ────────────────────────────────────────────────────────────────
def process_input(text):
    _user_said(text)
    try:
        result = ai.match_intent(text)
        intent = result["intent"]
        if intent == "check_at_risk":
            handle_check_at_risk()
        elif intent == "reporting":
            handle_reporting(text)
        elif intent == "ship_transfer":
            handle_ship(text)
        elif intent == "receive_transfer":
            handle_receive(text)
        elif intent == "create_pr":
            handle_create_pr(text)
        elif intent == "switch_mode":
            handle_switch_mode(text)
        elif intent == "upload_orders":
            handle_upload_orders()
        elif intent == "run_setup":
            handle_run_setup()
        elif intent == "complete_rest":
            handle_complete_rest()
        elif intent == "load_full_demo":
            handle_load_full_demo()
        elif intent == "add_customer_wave":
            handle_add_customer_wave()
        elif intent == "reset_data":
            handle_reset_data()
        elif intent == "query_org_profile":
            handle_query_org_profile()
        elif intent == "query_item_tax":
            handle_query_item_tax(text)
        elif intent == "explain":
            handle_explain(text)
        elif intent == "help":
            handle_help()
        else:
            _say("I couldn't match that request yet. Try a reporting prompt below, "
                 "or ask what needs action right now.")
    except Exception:
        _friendly_failure("that request")


# ── UI ────────────────────────────────────────────────────────────────────
_init_state()
_render_pending_refresh_notify()

st.markdown("## \U0001f916 Agent Console")
st.caption("Ask for operational reports, investigate seeded ERP data, or propose "
           "controlled actions. Technical errors are never exposed in the chat.")
st.divider()

for msg in st.session_state.chat_history:
    with st.chat_message(msg["role"], avatar="\U0001f916" if msg["role"] == "assistant" else None):
        st.markdown(msg["content"])

# ── Clarification prompt (ambiguous entity) ──────────────────────────────────
if st.session_state.clarify:
    c = st.session_state.clarify
    st.info("Please pick one:")
    cols = st.columns(len(c["candidates"]))
    for i, cand in enumerate(c["candidates"]):
        label = cand.get("desc") or cand.get("name")
        if cols[i].button(label, key=f"clarify_{i}"):
            kind, raw_text = c["kind"], c["raw_text"]
            st.session_state.clarify = None
            if kind == "material_for_explain":
                _explain_material(cand)
            elif kind == "material_for_ship":
                loc = ai.resolve_location(raw_text)
                if loc["match"]:
                    _propose_ship(cand["code"], cand["desc"], loc["match"]["id"], raw_text)
                else:
                    _say(f"Got **{cand['desc']}** — now, which destination did you mean?")
            elif kind == "location_for_ship":
                mat_resolved = ai.resolve_material(raw_text)
                mat_desc = mat_resolved["match"]["desc"] if mat_resolved["match"] else c["mat_code"]
                _propose_ship(c["mat_code"], mat_desc, cand["id"], raw_text)
            elif kind == "material_for_pr":
                _propose_pr(cand["code"], cand["desc"])
            elif kind == "material_for_receive":
                loc = ai.resolve_location(raw_text)
                to_location = loc["match"]["id"] if loc["match"] else None
                _propose_receive(cand["code"], cand["desc"], to_location, raw_text)
            elif kind == "transfer_for_receive":
                st.session_state.pending_action = {
                    "type": "receive", "transfer_id": cand["transfer_id"],
                    "summary": (f"Confirm receipt of **{cand['quantity']:g}** units of "
                                f"**{cand['material_desc']}** at "
                                f"**{_lname(cand['to_location'])}** (shipped from "
                                f"**{_lname(cand['from_location'])}**)?"),
                }
                _say(st.session_state.pending_action["summary"])
            st.rerun()

# ── Pending action confirmation card ─────────────────────────────────────────
if st.session_state.pending_action:
    pa = st.session_state.pending_action
    with st.container(border=True):
        if pa["type"] == "upload_orders":
            st.markdown("##### \U0001f4e5 Bulk Sales Order Upload")
            template_bytes = soi.generate_template()
            st.download_button("Download Sales Order Import Template", data=template_bytes,
                file_name="sales_order_import_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            uploaded = st.file_uploader("Filled-in template", type=["xlsx"], key="agent_upload")
            cols = st.columns([1, 1, 4])
            if uploaded and cols[0].button("\U0001f4e4 Upload"):
                try:
                    result = soi.import_sales_orders(uploaded.read())
                    st.cache_data.clear()
                    parts = []
                    if result["accepted"]:
                        parts.append(f"**{len(result['accepted'])} order(s) created**: " +
                                    ", ".join(f"{a['order_ref']} \u2192 {a['so_id']} "
                                             f"({a['status']})" for a in result["accepted"]))
                    if result["rejected"]:
                        parts.append(f"**{len(result['rejected'])} order(s) rejected**: " +
                                    "; ".join(f"{r['order_ref']}: {'; '.join(r['reasons'])}"
                                             for r in result["rejected"]))
                    _say("\n\n".join(parts) if parts else "No rows found in the uploaded file.")
                except Exception:
                    _friendly_failure("the order upload")
                else:
                    if result.get("accepted"):
                        _notify_shell_refresh()
                st.session_state.pending_action = None
                st.rerun()
            if cols[1].button("Cancel"):
                st.session_state.pending_action = None
                st.rerun()
        elif pa["type"] == "reset_data":
            st.markdown("##### \u26a0\ufe0f Reset & Reseed")
            st.caption("This wipes all current data — every PR, PO, GR, Sales "
                      "Order, invoice, and journal entry — and reloads from "
                      "whichever seed file you upload. Not reversible from here.")
            uploaded = st.file_uploader("Seed file (.xlsx)", type=["xlsx"], key="agent_seed_upload")
            cols = st.columns([1, 1, 4])
            if uploaded and cols[0].button("\u26a0\ufe0f Confirm Reset", type="primary"):
                tmp_path = os.path.join(_DIR, f"_agent_seed_upload_{uploaded.name}")
                with open(tmp_path, "wb") as f:
                    f.write(uploaded.getbuffer())
                did_reset = False
                try:
                    validation = sm.validate_seed_file(tmp_path)
                    if not validation.get("valid", True):
                        _say(f"\u274c This file didn't validate: "
                            f"{'; '.join(validation.get('errors', ['unknown issue']))}")
                    else:
                        sm.reset_and_reseed(tmp_path)
                        st.cache_data.clear()
                        did_reset = True
                        _say("\u2705 Reset and reseeded. Restart the traditional "
                            "apps (or just switch tabs and back) to see fresh data.")
                except Exception:
                    _friendly_failure("the data reset")
                else:
                    if did_reset:
                        _notify_shell_refresh()
                finally:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                st.session_state.pending_action = None
                st.rerun()
            if cols[1].button("Cancel", key="agent_reset_cancel"):
                st.session_state.pending_action = None
                st.rerun()
        else:
            st.markdown(f"**Proposed action:** {pa['summary']}")
            c1, c2, _ = st.columns([1, 1, 4])
            if c1.button("\u2705 Approve", type="primary"):
                try:
                    if pa["type"] == "ship":
                        result = inv.ship_transfer(pa["mat_code"], pa["mat_desc"],
                            pa["from_location"], pa["to_location"], pa["qty"],
                            shipped_by="Agent Console")
                        st.cache_data.clear()
                        _say(f"\u2705 {result['transfer_id']} — {pa['qty']:g} units "
                            f"shipped, now in transit.")
                    elif pa["type"] == "receive":
                        inv.receive_transfer(pa["transfer_id"], received_by="Agent Console")
                        st.cache_data.clear()
                        _say(f"\u2705 {pa['transfer_id']} received — stock is now "
                            f"usable at the destination.")
                    elif pa["type"] == "create_pr":
                        pr_id = f"PR-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                        pc.create_pr(pr_id, requester_id="AGENT-CONSOLE",
                            requester_name="Agent Console", requester_dept="Planning",
                            project_id="AGENT-CONSOLE",
                            lines=[{"vendor": "", "mat_code": pa["mat_code"],
                                    "mat_desc": pa["mat_desc"], "uom": "pcs",
                                    "qty": pa["qty"], "req_date": pa["required_by"],
                                    "deliv_loc": pa["location"], "deliv_geo": ""}])
                        st.cache_data.clear()
                        _say(f"\u2705 {pr_id} created — required by {pa['required_by']}.")
                    elif pa["type"] == "switch_mode":
                        od.set_org_default("Time-Phased Planning Mode", pa["mode"])
                        st.cache_data.clear()
                        _say(f"\u2705 Switched to **{pa['mode']}**.")
                    elif pa["type"] in ("run_setup", "complete_rest", "load_full_demo",
                                       "add_customer_wave"):
                        ds_mod = importlib.import_module(pa["module"])
                        if pa["type"] == "run_setup":
                            ds_mod.run_setup()
                            st.cache_data.clear()
                            _say("\u2705 Setup loaded — head to Manufacturing app's "
                                "Inventory \u2192 Position & Transfers to see the "
                                "imbalance and execute the transfers.")
                        elif pa["type"] == "complete_rest":
                            ds_mod.run_resolution()
                            st.cache_data.clear()
                            _say("\u2705 Story complete through cash application.")
                        elif pa["type"] == "load_full_demo":
                            ds_mod.run_all()
                            st.cache_data.clear()
                            _say("\u2705 Full demo loaded, start to finish, in one shot.")
                        else:
                            result = ds_mod.new_customer_wave()
                            st.cache_data.clear()
                            names = ", ".join(r["customer"] for r in result)
                            _say(f"\u2705 New customer wave added — {names} are on "
                                f"file with real quotations, confirmed orders, and "
                                f"Open PRs, ready for a live PR Consolidation run.")
                except Exception:
                    _friendly_failure("the approved action")
                else:
                    _notify_shell_refresh()
                st.session_state.pending_action = None
                st.rerun()
            if c2.button("\u274c Reject"):
                _say("Understood — no action taken.")
                st.session_state.pending_action = None
                st.rerun()

st.divider()
st.markdown("###### Try one of these, or type your own:")
example_prompts = build_example_prompts()
cols = st.columns(3)
for i, prompt in enumerate(example_prompts):
    if cols[i % 3].button(prompt, key=f"example_{i}", use_container_width=True):
        process_input(prompt)
        st.rerun()

typed = st.chat_input("Or type your own request...")
if typed:
    process_input(typed)
    st.rerun()
