"""
nav_catalog.py — the one real, shared directory of every page (and,
for naming/matching only, every tab) across erp_ui.py/mfg_ui.py/
o2c_ui.py. Pure data, no Streamlit import — same "zero Streamlit
dependency, independently testable" discipline agent_intents.py
already holds itself to, since this is imported by that module too.

Used by:
  - agent_intents.py's own resolve_screen() — fuzzy-matches free text
    against this catalog.
  - agent_console.py — builds the real link a match resolves to.

Scope, this round, by direct instruction: page-level linking only. A
tab is listed here so it's genuinely nameable and matchable ("show me
Position & Transfers" correctly finds it), and the agent's own
response names it in plain text ("...then click the Position &
Transfers tab") — but the link itself always opens the tab's own
parent PAGE, not the tab directly. Real tab-level deep-linking was
scoped and explicitly set aside as bigger, riskier work: Streamlit's
own st.tabs() has no programmatic-selection mechanism at all, so
landing pre-selected on a specific tab would mean restructuring every
one of this codebase's ~15 st.tabs() call sites across all three
files — a real, separate piece of engineering, not a small addition.
Full fork in CONTEXT_HANDOFF_v2.md.

Deliberately mirrors each app's own inline page-options list rather
than replacing it — kept in sync by hand, not by sharing one list
object, so a change here can never alter what those lists actually
render (order, text, emoji) as a side effect. Each page's own "page"
key is exactly the substring each app's own router already keys off
(e.g. mfg_ui.py's `elif "Inventory" in page:`) — reused directly here
rather than inventing a second identifier scheme, so there is only
ever one real name for "the Inventory page" anywhere in this codebase.
"""

APPS = {
    "s2c": {"port": 8501, "app_label": "Source to Contract"},
    "mfg": {"port": 8502, "app_label": "Manufacturing"},
    "o2c": {"port": 8503, "app_label": "Order to Cash"},
}

# Tab labels are the static part only — several real tab labels carry a
# live count suffix (e.g. "Award History (34)"); that number can never
# be hardcoded here, and isn't needed for matching or naming purposes.
PAGES = {
    "s2c": [
        {"page": "Purchase Bundles",
         "tabs": ["Bundles", "Create Bundle", "Discover Candidates"]},
        {"page": "Create PR",
         "tabs": ["Individual Items", "Purchase Bundles"]},
        {"page": "Consolidate", "label": "Consolidate PRs → POs",
         "tabs": ["Flow", "Map", "POs + Create PO", "RFP"]},
        {"page": "Direct PO", "label": "Direct PO Entry", "tabs": []},
        {"page": "Auto-Categorization",
         "tabs": ["High / Medium confidence", "Needs review"]},
        {"page": "RFx", "label": "RFx Management",
         "tabs": ["Select & Invite", "Enter Quotes", "Select Winners",
                  "Issue POs", "Award History"]},
        {"page": "Vendor Onboarding",
         "tabs": ["Vendors", "VRQ Requests", "Scorecard"]},
        {"page": "Contracts",
         "tabs": ["Contracts", "Convert PO to Contract"]},
    ],
    "mfg": [
        {"page": "Goods Receipt", "tabs": ["Create GR", "Manage GRs"]},
        {"page": "Quality Inspection",
         "tabs": ["Record Inspection", "Manage Inspections",
                  "Quality Holds & RTV", "RMA Receipt & Disposition"]},
        {"page": "BOM", "label": "BOM & Explosion",
         "tabs": ["Explode & Propose PR", "View BOM"]},
        {"page": "Production", "tabs": ["Confirm Production", "History"]},
        {"page": "Inventory",
         "tabs": ["Stock by Location", "By Material", "Transaction History",
                  "Position & Transfers", "Time-Phased Planning", "Traceability"]},
    ],
    "o2c": [
        {"page": "Customer Onboarding", "tabs": []},
        {"page": "Quotation", "tabs": ["Create Quote", "Manage Quotes"]},
        {"page": "Sales Orders",
         "tabs": ["Create Order", "Manage Orders", "Bulk Import", "Backorders"]},
        {"page": "Fulfillment",
         "tabs": ["Create Fulfillment", "Manage Fulfillments"]},
        {"page": "Billing", "label": "Billing & Invoicing",
         "tabs": ["Create Invoice", "Manage Invoices"]},
        {"page": "Cash Application",
         "tabs": ["Record Payment", "Apply Payments", "Manage & Reports"]},
        {"page": "Returns", "label": "Returns (RMA)",
         "tabs": ["Authorize RMA", "Issue Credit Memo", "All RMAs"]},
        {"page": "Accounting",
         "tabs": ["Journal Entries", "Chart of Accounts", "Vendor Invoices"]},
        {"page": "Settings",
         "tabs": ["Organization Profile", "Item Tax (HSN/GST)", "Data Reset & Seed"]},
    ],
}


def build_page_url(app, page_key):
    """A real URL that opens the target app landing on the requested
    page — page["page"] is the same substring that app's own router
    matches on, passed straight through as the `page` query param."""
    from urllib.parse import quote
    port = APPS[app]["port"]
    return f"http://localhost:{port}/?page={quote(page_key)}"


def resolve_page_index(page_options, query_value):
    """
    Maps a `?page=` query param value to the right index into an app's
    own real page_options list (its st.radio(...) options, passed in
    exactly as that app already defines them, never duplicated here) —
    same substring match each app's own router already uses to decide
    which page function to call (e.g. mfg_ui.py's own
    `elif "Inventory" in page:`), so a link this module builds is
    guaranteed to land on the exact same page that link named. Returns
    0 (today's implicit default — st.radio() picks the first option
    when index is unset) whenever query_value is empty or matches
    nothing, so a normal, non-deep-linked page load behaves exactly as
    it always has.
    """
    if not query_value:
        return 0
    for i, opt in enumerate(page_options):
        if query_value in opt:
            return i
    return 0


def all_screens():
    """Flat list of every real (page) and (page, tab) pair across all
    three apps, each carrying everything a caller needs to both match
    against it and act on a match: which app, the page's own real
    dispatch substring, its display label, and — for a tab entry —
    that tab's own real name. A page with no tabs contributes exactly
    one entry (tab=None)."""
    out = []
    for app, pages in PAGES.items():
        app_label = APPS[app]["app_label"]
        for p in pages:
            page_key = p["page"]
            page_label = p.get("label", page_key)
            out.append({"app": app, "app_label": app_label, "page_key": page_key,
                        "page_label": page_label, "tab": None})
            for t in p["tabs"]:
                out.append({"app": app, "app_label": app_label, "page_key": page_key,
                            "page_label": page_label, "tab": t})
    return out
