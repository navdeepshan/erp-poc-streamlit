"""
erp_ui.py v4 — Clean rebuild fixing all reported issues.
Run: py -m streamlit run erp_ui.py
All 4 files must be in the same folder: erp_ui.py, po_export.py, pr_consolidation.py, data.xlsx
"""
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json, os, sys, io, traceback, zipfile
from datetime import date, datetime, timedelta
from collections import defaultdict

# ── Paths ────────────────────────────────────────────────────────────────────
try:
    _DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _DIR = os.getcwd()

for _p in [_DIR, os.getcwd()]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

import po_export
import db
import pr_consolidation
import purchase_bundles as pbdl
import bom
import categorization
import rfx
import vendor_onboarding as vo
import vendor_scorecard as vsc
import vrq
import contracts as ct
import org_defaults as od
import legal_entities as le
import nav_catalog as nav
from professional_theme import apply_theme

DATA_FILE = os.path.join(_DIR, "data.xlsx")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _build_zip(files):
    """files: list of {'filename': str, 'bytes': bytes}. Returns zip bytes.
    Used wherever multiple generated documents need one 'download all' button —
    st.download_button triggers a script rerun, and files only held in a local
    variable (not session_state) vanish on that rerun, so batches of generated
    documents need either persistence or a single bundled download."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.writestr(f["filename"], f["bytes"])
    buf.seek(0)
    return buf.read()

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="ERP Procurement", page_icon="\U0001f4cb",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.stApp{background:#F8FAFC}
[data-testid="stSidebar"]{background:#fff;border-right:1px solid #E2E8F0}
body,p,div,label,span{color:#1E293B!important}
h1,h2,h3{color:#0F172A!important;font-weight:600}
.stMarkdown p{color:#334155!important}
[data-testid="stSidebar"] *{color:#334155!important}
.stTextInput input,.stNumberInput input,.stDateInput input{
  background:#fff!important;color:#1E293B!important;
  border:1px solid #CBD5E1!important;border-radius:6px!important}
.stTextInput input:focus{border-color:#2563EB!important;
  box-shadow:0 0 0 3px rgba(37,99,235,.12)!important}
.stSelectbox>div>div{background:#fff!important;color:#1E293B!important;
  border:1px solid #CBD5E1!important;border-radius:6px!important}
div[data-testid="stMetric"]{background:#fff;border:1px solid #E2E8F0;
  border-radius:10px;padding:14px 18px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
div[data-testid="stMetricLabel"]{color:#64748B!important;font-size:12px!important}
div[data-testid="stMetricValue"]{color:#0F172A!important;font-weight:700!important}
.stButton>button{background:#2563EB!important;color:#fff!important;
  border:none!important;border-radius:7px!important;font-weight:500!important}
.stButton>button:hover{background:#1D4ED8!important}
.stButton>button[kind="secondary"]{background:#fff!important;
  color:#374151!important;border:1px solid #D1D5DB!important}
.stButton>button:disabled{background:#94A3B8!important;cursor:not-allowed!important}
.stDownloadButton>button{background:#059669!important;color:#fff!important;
  border:none!important;border-radius:7px!important}
.stDataFrame{border:1px solid #E2E8F0!important;border-radius:8px!important}
.stSuccess{background:#F0FDF4!important;border:1px solid #86EFAC!important;border-radius:8px!important}
.stError{background:#FEF2F2!important;border:1px solid #FECACA!important;border-radius:8px!important}
.stWarning{background:#FFFBEB!important;border:1px solid #FDE68A!important;border-radius:8px!important}
.stInfo{background:#EFF6FF!important;border:1px solid #BFDBFE!important;border-radius:8px!important}
div[data-baseweb="tab-list"]{background:#F1F5F9;border-radius:8px;
  padding:4px;border:1px solid #E2E8F0}
div[data-baseweb="tab"]{color:#64748B!important;border-radius:6px!important}
div[data-baseweb="tab"][aria-selected="true"]{background:#fff!important;
  color:#1E293B!important;box-shadow:0 1px 3px rgba(0,0,0,.1)!important}
details{background:#fff;border:1px solid #E2E8F0!important;
  border-radius:8px!important;margin-bottom:6px!important}
summary{color:#1E293B!important;font-weight:500!important;padding:12px 16px!important}
hr{border-color:#E2E8F0!important}
</style>""", unsafe_allow_html=True)
apply_theme("Source to Pay")

# ── Startup check ─────────────────────────────────────────────────────────────
# Checks erp_pilot.db, not data.xlsx — data.xlsx hasn't been a real runtime
# dependency for any of these three apps since the SQLite migration; it's
# only ever read by the one-time migrate_*.py scripts now. This used to
# gate on data.xlsx instead, which meant a fully-migrated, fully-working
# setup would still hard-refuse to start (st.stop(), the whole app) if
# that file happened to be missing or moved — even though nothing here
# actually needs it anymore.
if not os.path.exists(db.DB_FILE):
    st.error(f"\u274c erp_pilot.db not found at: {db.DB_FILE}\n\n"
             f"Run the migrate_*.py scripts first to create and seed it.\n"
             f"Current folder: {_DIR}")
    st.stop()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### \U0001f4cb AutonoVerse Garage")
    st.caption("Source to Contract")
    st.divider()
    _page_options = ["📦  Purchase Bundles",
                     "🆕  Create PR",
                     "🔄  Consolidate PRs → POs",
                     "📤  Direct PO Entry",
                     "🏷️  Auto-Categorization",
                     "📨  RFx Management",
                     "🧾  Vendor Onboarding",
                     "📜  Contracts"]
    # `?page=` deep-link support (2026-08-11) for the Agent Console's
    # own navigation links -- see nav_catalog.py's own docstring. A
    # normal load with no query param resolves to index 0, identical
    # to before this existed.
    page = st.radio("", _page_options,
                    index=nav.resolve_page_index(_page_options, st.query_params.get("page")),
                    label_visibility="collapsed")
    st.divider()


# ── Master data loaders ────────────────────────────────────────────────────────
@st.cache_data(ttl=120)
def load_catalog():
    return po_export.load_item_master(DATA_FILE)

@st.cache_data(ttl=300)
def load_delivery_locs():
    """Delivery_Locations now lives in SQLite — delegates to
    pr_consolidation.py's canonical reader (also used by mfg_ui.py and
    o2c_ui.py) instead of this file's own separate Excel read. That
    separate read never filtered on Active despite the column existing
    — closed now that there's exactly one implementation."""
    try:
        return pr_consolidation.get_delivery_locations(active_only=True)
    except Exception:
        return []

def vendor_from_tags(tags):
    # Only an ALL-CAPS first word counts as a real vendor/brand tag
    # (WOODPECKER, MELAG, etc. — the established convention every real
    # vendor tag in this item master follows). A lowercase first word
    # is a generic descriptive tag, not a vendor — found as a real bug
    # while building a new bundle: "handpiece high speed turbine
    # clinical" resolved to vendor="handpiece", which isn't a real
    # vendor id at all (confirmed: vendor_onboarding.get_vendor
    # returns None for it), silently 9 items wide across the current
    # Item Master, not just this one.
    t = tags.strip()
    first = t.split()[0] if t else ""
    return first if first.isupper() else ""


# ── AV export ─────────────────────────────────────────────────────────────────
# make_av_bytes moved to po_export.py — a pure document generator with no
# Streamlit dependency, needed to be importable from non-UI scripts
# (demo_scenario.py) without triggering this whole page's top-level code
# to execute as an import side effect.


# ── Item picker ────────────────────────────────────────────────────────────────
# FIX: search uses button click (Streamlit text_input only fires on Enter/blur).
# "Add Selected" button always visible, disabled when nothing selected.
def render_item_picker(sk):
    staged_k   = f"{sk}_staged"
    selected_k = f"{sk}_selected"
    if staged_k   not in st.session_state: st.session_state[staged_k]   = []
    if selected_k not in st.session_state: st.session_state[selected_k] = set()

    items = load_catalog()
    cats  = ["All"] + sorted(set(i["category"] for i in items if i["category"]))

    # Search row: input + button side by side
    s1, s2, s3 = st.columns([4, 1, 1])
    with s1:
        query = st.text_input("Search", placeholder="Type then click \U0001f50d",
                              key=f"{sk}_q", label_visibility="collapsed")
    with s2:
        do_search = st.button("\U0001f50d Search", key=f"{sk}_sbtn",
                              use_container_width=True)
    with s3:
        cat = st.selectbox("Cat", cats, key=f"{sk}_cat",
                           label_visibility="collapsed")

    # Store last searched query so results persist across reruns
    if do_search or st.session_state.get(f"{sk}_last_q") != query:
        st.session_state[f"{sk}_last_q"]   = query
        st.session_state[f"{sk}_results"]  = []
        if len(query) >= 2:
            res = po_export.fuzzy_search(query, items, max_results=20)
            if cat != "All":
                res = [r for r in res if r["category"] == cat]
            st.session_state[f"{sk}_results"] = res
        elif cat != "All":
            st.session_state[f"{sk}_results"] = [i for i in items if i["category"] == cat][:20]

    results       = st.session_state.get(f"{sk}_results", [])
    staged_codes  = {i["code"] for i in st.session_state[staged_k]}
    n_sel         = len(st.session_state[selected_k])

    # FIX: Add button always visible, disabled when nothing selected
    if st.button(f"\u2795  Add {n_sel} selected item(s)" if n_sel else "\u2795  Add selected item(s)",
                 type="primary", key=f"{sk}_add",
                 disabled=(n_sel == 0),
                 use_container_width=True):
        added = 0
        for item in items:
            if (item["code"] in st.session_state[selected_k]
                    and item["code"] not in staged_codes):
                st.session_state[staged_k].append({
                    "code": item["code"], "desc": item["desc"],
                    "vendor": vendor_from_tags(item["tags"]),
                    "uom": item["uom"], "price": item["price"], "qty": 1,
                })
                added += 1
        st.session_state[selected_k] = set()
        if added:
            st.success(f"Added {added} item(s) to the list.")
            st.rerun()

    if results:
        st.caption(f"{len(results)} result(s)" + (f"  ·  {n_sel} ticked" if n_sel else ""))
        for item in results:
            checked = item["code"] in st.session_state[selected_k]
            in_list = item["code"] in staged_codes
            c1, c2, c3 = st.columns([0.5, 6.5, 1.5])
            with c1:
                new_val = st.checkbox("", value=checked,
                    key=f"{sk}_c_{item['code']}",
                    label_visibility="collapsed")
                if new_val != checked:
                    if new_val: st.session_state[selected_k].add(item["code"])
                    else:       st.session_state[selected_k].discard(item["code"])
                    st.rerun()
            with c2:
                pfx = "\u2713 " if in_list else ""
                vnd = vendor_from_tags(item["tags"])
                st.markdown(
                    f"<div style='padding:3px 0'>"
                    f"<span style='color:#1E293B;font-weight:500'>{pfx}{item['desc']}</span><br>"
                    f"<span style='font-size:11px;color:#64748B'>{item['code']} · {item['uom']}</span>"
                    + (f"&nbsp;<span style='background:#EFF6FF;color:#1D4ED8;font-size:11px;"
                       f"padding:2px 6px;border-radius:4px'>{vnd}</span>" if vnd else "")
                    + "</div>", unsafe_allow_html=True)
            with c3:
                st.markdown(
                    f"<div style='text-align:right;padding-top:5px;color:#059669;"
                    f"font-family:monospace;font-weight:600'>\u20b9{item['price']:,.2f}</div>",
                    unsafe_allow_html=True)
    elif len(query) >= 2 and not do_search:
        st.caption("Press \U0001f50d Search to find items.")
    elif query or cat != "All":
        st.info("No items found.")
    else:
        st.markdown(
            "<div style='color:#94A3B8;padding:24px;text-align:center;"
            "border:1px dashed #CBD5E1;border-radius:8px'>"
            "\U0001f50d Type and click Search to find items</div>",
            unsafe_allow_html=True)


# ── Staged items table ─────────────────────────────────────────────────────────
def render_staged_table(sk, show_vendor=False):
    staged_k = f"{sk}_staged"
    staged   = st.session_state.get(staged_k, [])
    if not staged:
        return False

    rows = []
    for it in staged:
        row = {"Code": it["code"], "Description": it["desc"],
               "UOM":  it["uom"], "Qty": float(it["qty"]),
               "Rate": it["price"],
               "Line Total": round(float(it["qty"]) * it["price"], 2)}
        if show_vendor:
            row["Vendor"] = it.get("vendor","")
        rows.append(row)

    df   = pd.DataFrame(rows)
    cols = ["Code","Description","UOM","Qty","Rate","Line Total"]
    if show_vendor:
        cols = ["Code","Description","Vendor","UOM","Qty","Rate","Line Total"]

    cfg = {
        "Code":       st.column_config.TextColumn(disabled=True, width=90),
        "Description":st.column_config.TextColumn(disabled=True, width=200),
        "UOM":        st.column_config.TextColumn(disabled=True, width=55),
        "Qty":        st.column_config.NumberColumn(min_value=0, step=1, width=65),
        "Rate":       st.column_config.NumberColumn(format="\u20b9%.2f", disabled=True, width=85),
        "Line Total": st.column_config.NumberColumn(format="\u20b9%.2f", disabled=True, width=100),
    }
    if show_vendor:
        cfg["Vendor"] = st.column_config.TextColumn(disabled=True, width=100)

    edited = st.data_editor(df[cols], use_container_width=True,
                            hide_index=True, column_config=cfg, key=f"{sk}_tbl")

    # Sync qty from the edited table back to session_state.
    # Do NOT st.rerun() here — it races with button clicks
    # (e.g. Create PR) and stops them firing.
    # Compute totals from `edited` directly so they are always
    # current even before the user presses Enter to commit a cell.
    for i, row in edited.iterrows():
        if i < len(staged):
            staged[i]["qty"] = float(row["Qty"] or 0)

    # Totals from the live edited values (always up to date)
    edited_qtys = [float(edited.iloc[i]["Qty"] or 0) for i in range(len(staged))]
    total = sum(edited_qtys[i] * staged[i]["price"] for i in range(len(staged)))
    m1, m2, m3 = st.columns(3)
    m1.metric("Lines", len(staged))
    m2.metric("Items", int(sum(edited_qtys)))
    m3.metric("Total", f"\u20b9{total:,.2f}")

    # Alert: flag staged items that already have inventory coverage somewhere
    # — the point-of-action check. Doesn't block anything, just surfaces it
    # before you commit, same as every other "record honestly, let a human
    # decide" pattern in this app.
    try:
        coverage = bom.check_items_coverage([it["code"] for it in staged])
        covered = [(it, coverage[it["code"]]) for it in staged
                  if coverage.get(it["code"], {}).get("total_on_hand", 0) > 0]
        if covered:
            with st.expander(f"\u26a0\ufe0f {len(covered)} of {len(staged)} item(s) already have "
                            "inventory on hand somewhere — check before ordering more"):
                for it, cov in covered:
                    locs = ", ".join(f"{l['location_id']} ({l['balance']:g})" for l in cov["by_location"])
                    st.caption(f"**{it['code']}** — {it['desc'][:50]} — "
                              f"{cov['total_on_hand']:g} on hand: {locs}"
                              + (f" · {cov['open_po']:g} already on an open PO" if cov["open_po"] else ""))
    except Exception:
        pass  # coverage check is advisory only — never let it block staging/saving

    with st.expander("\U0001f5d1  Remove items"):
        to_rm = st.multiselect("", [i["code"] for i in staged],
                               key=f"{sk}_rm", label_visibility="collapsed")
        if to_rm and st.button("Remove", key=f"{sk}_do_rm"):
            st.session_state[staged_k] = [i for i in staged if i["code"] not in to_rm]
            st.rerun()
    return True


# ── Write PR ────────────────────────────────────────────────────────────────
# PR_Header/PR_Items now live in SQLite (pr_consolidation.py's pilot) — this
# delegates instead of writing Excel cells directly. Kept as a thin wrapper
# (same call signature as before) so page_create_pr() below didn't need to
# change its call site.
def render_bundle_picker(sk):
    """
    Purchase Bundles picker — sits alongside the individual item picker
    in Create PR (see page_create_pr()'s tabs). Picking a bundle
    explodes it (purchase_bundles.explode_bundle()) into the SAME
    staged-item shape the individual picker already produces
    (code/desc/vendor/uom/price/qty), then appends into the same
    `{sk}_staged` session-state list — so from render_staged_table()'s
    point of view, a bundle-added line and an individually-picked line
    are indistinguishable. That's the entire integration surface: once
    staged, everything downstream (qty editing, Create PR, Consolidate,
    RFx, PO) has no idea a bundle was ever involved.

    Same "skip if already staged" convention render_item_picker uses —
    if two bundles (or a bundle and an individual pick) share a
    material code, only the first one staged wins; the picker doesn't
    try to merge/sum quantities across sources.
    """
    staged_k = f"{sk}_staged"
    if staged_k not in st.session_state:
        st.session_state[staged_k] = []

    bundles = pbdl.list_bundles(active_only=True)
    if not bundles:
        st.info("No purchase bundles yet. Create one from the "
                "**\U0001f4e6 Purchase Bundles** page.")
        return

    depts = ["All"] + sorted({b["department"] for b in bundles if b["department"]})
    dept_filter = st.selectbox("Department", depts, key=f"{sk}_bdl_dept")
    visible = bundles if dept_filter == "All" else [b for b in bundles if b["department"] == dept_filter]

    labels = {b["bundle_id"]: f"{b['bundle_name']}" + (f"  ·  {b['department']}" if b["department"] else "")
              for b in visible}
    sel_id = st.selectbox("Bundle", list(labels.keys()), format_func=lambda k: labels[k],
                          key=f"{sk}_bdl_sel")
    sel_bundle = next(b for b in visible if b["bundle_id"] == sel_id)
    if sel_bundle["description"]:
        st.caption(sel_bundle["description"])

    items = pbdl.get_bundle_items(sel_id)
    st.dataframe(
        pd.DataFrame([{"Code": i["mat_code"], "Description": i["mat_desc"],
                       "UOM": i["uom"], "Default Qty": i["default_qty"]} for i in items]),
        use_container_width=True, hide_index=True, height=min(250, 42 + 35 * len(items)))

    multiplier = st.number_input("Multiply quantities by", min_value=1, value=1, step=1,
                                 key=f"{sk}_bdl_mult",
                                 help="E.g. picking this bundle for 3 sites at once.")

    if st.button(f"\u2795  Add {len(items)} bundle item(s) to PR", type="primary",
                 use_container_width=True, key=f"{sk}_bdl_add"):
        exploded = pbdl.explode_bundle(sel_id, multiplier=multiplier)
        staged_codes = {i["code"] for i in st.session_state[staged_k]}
        added = 0
        for line in exploded:
            if line["code"] not in staged_codes:
                st.session_state[staged_k].append({
                    "code": line["code"], "desc": line["desc"], "vendor": line["vendor"],
                    "uom": line["uom"], "price": line["price"] or 0.0, "qty": line["qty"],
                })
                staged_codes.add(line["code"])
                added += 1
        skipped = len(exploded) - added
        if added:
            msg = f"Added {added} item(s) from **{sel_bundle['bundle_name']}**."
            if skipped:
                msg += f" ({skipped} already in the list, skipped.)"
            st.success(msg)
            st.rerun()
        else:
            st.warning("Every item in this bundle is already in the list.")


def write_pr(pr_number, req_id, req_name, dept, project_id, req_date,
              deliv_loc, deliv_geo, po_type, legal, purch_e, purch_g,
              currency, plant, lines):
    pr_lines = [{"vendor": ln.get("vendor", ""), "mat_code": ln["code"],
                 "mat_desc": ln["desc"], "uom": ln["uom"], "qty": ln["qty"],
                 "req_date": req_date, "deliv_loc": deliv_loc, "deliv_geo": deliv_geo}
                for ln in lines]
    pr_consolidation.create_pr(pr_number, requester_id=req_id, requester_name=req_name,
                                requester_dept=dept, project_id=project_id,
                                po_type=po_type, legal_entity=legal, purchase_entity=purch_e,
                                purchasing_group=purch_g, currency=currency, plant_code=plant,
                                lines=pr_lines)
    return pr_number


# ── Next PR number ─────────────────────────────────────────────────────────────
def next_pr_num():
    try:
        return pr_consolidation.next_pr_number()
    except Exception:
        return f"PR-{date.today().strftime('%Y%m%d')}-001"


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — Create PR
# ══════════════════════════════════════════════════════════════════════════════
def page_create_pr():
    st.markdown("## \U0001f195 Create Purchase Requisition")
    st.divider()

    left, right = st.columns([5, 4], gap="large")

    with left:
        st.markdown("#### \U0001f50d Item Picker")
        tab_items, tab_bundles = st.tabs(["Individual Items", "\U0001f4e6 Purchase Bundles"])
        with tab_items:
            render_item_picker("pr")
        with tab_bundles:
            render_bundle_picker("pr")

    with right:
        st.markdown("#### \U0001f4c4 PR Details")
        if "pr_draft" not in st.session_state:
            st.session_state.pr_draft = next_pr_num()

        c1, c2 = st.columns(2)
        with c1:
            pr_num   = st.text_input("PR Number",     value=st.session_state.pr_draft, key="pr_n")
            req_id   = st.text_input("Requester ID",   value="REQ-104",  key="pr_rid")
            req_name = st.text_input("Requester Name", value="",         key="pr_rn")
            dept     = st.text_input("Department",     value="",         key="pr_d")
        with c2:
            project  = st.text_input("Project ID",    value="PROJ-GEN", key="pr_p")
            req_date = st.date_input("Required Date",  value=date.today(),key="pr_rd")
            locs     = load_delivery_locs()
            if locs:
                sel_loc  = st.selectbox("Delivery Location", [l["id"] for l in locs], key="pr_dl")
                dloc_geo = next((l["geo"] for l in locs if l["id"]==sel_loc), "")
                st.caption(f"Geo: {dloc_geo or '—'}")
            else:
                sel_loc  = st.text_input("Delivery Location", value="BD_DL_BLR_Godown 1", key="pr_dl2")
                dloc_geo = st.text_input("Geo (lat,lng)", value="12.9716,77.5946", key="pr_dg")

        with st.expander("\u2699\ufe0f Org defaults"):
            oa, ob = st.columns(2)
            with oa:
                po_type  = st.text_input("PO Type",   value=od.get_default("PO Type"),       key="pr_pt")
                legal    = st.text_input("Legal Ent.", value=od.get_default("Legal Entity"),   key="pr_le")
                purch_e  = st.text_input("Purch.Ent.",value=od.get_default("Purchasing Entity"),   key="pr_pe")
            with ob:
                purch_g  = st.text_input("Purch.Grp.",value=od.get_default("Purchasing Group"),   key="pr_pg")
                currency = st.text_input("Currency",  value=od.get_default("Currency"),      key="pr_cu")
                plant    = st.text_input("Plant",     value=od.get_default("Plant"), key="pr_pl")

        st.divider()
        staged = st.session_state.get("pr_staged", [])

        if staged:
            st.markdown("**PR Lines**")
            render_staged_table("pr", show_vendor=True)
            st.markdown("")

            if st.button("\u2705  Create PR", type="primary",
                         use_container_width=True, key="pr_create"):
                lines = [i for i in st.session_state.pr_staged if i.get("qty",0) > 0]
                if not lines:
                    st.error("Set at least one quantity > 0 before creating.")
                else:
                    try:
                        write_pr(
                            pr_number=pr_num, req_id=req_id, req_name=req_name,
                            dept=dept, project_id=project, req_date=str(req_date),
                            deliv_loc=sel_loc, deliv_geo=dloc_geo,
                            po_type=po_type, legal=legal, purch_e=purch_e,
                            purch_g=purch_g, currency=currency, plant=plant,
                            lines=lines)
                        st.success(
                            f"\u2705 **{pr_num}** created — "
                            f"{len(lines)} line(s) written.")
                        st.session_state.pr_staged   = []
                        st.session_state.pr_selected = set()
                        st.session_state.pr_draft    = next_pr_num()
                        st.session_state.pr_results  = []
                        st.cache_data.clear()
                    except Exception:
                        st.error("\u274c Error writing PR — see details below:")
                        st.code(traceback.format_exc())

            if st.button("\U0001f5d1 Clear list", use_container_width=True, key="pr_clr"):
                st.session_state.pr_staged   = []
                st.session_state.pr_selected = set()
                st.rerun()
        else:
            st.markdown(
                "<div style='color:#94A3B8;padding:36px;text-align:center;"
                "border:1px dashed #CBD5E1;border-radius:8px'>"
                "\U0001f4ed No items staged<br>"
                "<small>Search on the left, tick items, click Add</small></div>",
                unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — Consolidate
# ══════════════════════════════════════════════════════════════════════════════
def page_consolidate():
    st.markdown("## \U0001f504 PR Consolidation")
    st.divider()

    with st.expander("\U0001f5fa\ufe0f All POs Map — every PO within the selected date range, not just this session"):
        st.caption("Shows every PO within the selected date range, including Direct PO Entry. "
                   "The map and the line-item list below share the same filter, so nothing "
                   "appears in one but not the other.")

        dpo1, dpo2 = st.columns(2)
        with dpo1:
            po_from = st.date_input("Delivery date from", value=date.today() - timedelta(days=90),
                                    key="po_list_from")
        with dpo2:
            po_to = st.date_input("Delivery date to", value=date.today() + timedelta(days=270),
                                  key="po_list_to")

        all_geo = pr_consolidation.build_all_pos_geo_data(DATA_FILE, date_from=po_from, date_to=po_to)
        if all_geo["routes"]:
            _map(all_geo)
        else:
            st.info("No plottable routes in this range — either no PO has a delivery date "
                   "here, or the ones that do are missing a vendor/delivery geolocation "
                   "(check the list below, which isn't geo-gated).")

        st.markdown("##### \U0001f4cb PO Line Items")

        all_po_items = pr_consolidation.get_all_po_line_items(DATA_FILE)
        if not all_po_items:
            st.info("No PO line items yet.")
        else:
            podf = pd.DataFrame(all_po_items)
            dd = pd.to_datetime(podf["delivery_date"], errors="coerce").dt.date
            in_range = dd.notna() & (dd >= po_from) & (dd <= po_to)
            shown = podf[in_range]
            st.caption(f"{len(shown)} of {len(podf)} PO line(s) with a delivery date in this range "
                      f"(lines with no delivery date set are excluded — adjust the range above).")
            if len(shown):
                shown = shown[["po_number","supplier_name","material_code","material_desc","uom",
                               "quantity","unit_price","delivery_date","delivery_location",
                               "source_pr_number"]].rename(columns={
                    "po_number":"PO","supplier_name":"Supplier","material_code":"Code",
                    "material_desc":"Description","uom":"UOM","quantity":"Qty",
                    "unit_price":"Unit Price","delivery_date":"Delivery Date",
                    "delivery_location":"Delivery Location","source_pr_number":"Source PR"})
                st.dataframe(shown, use_container_width=True, hide_index=True, height=300)

    st.divider()

    @st.cache_data(ttl=15)
    def load_pr_data():
        rows = pr_consolidation.get_pr_report()
        return pd.DataFrame(rows, columns=["PR","Line","Vendor","Code","Description","UOM",
                                            "Qty","Required","Status","Requester","PR_Date",
                                            "PO_Number","PO_Item"])

    df = load_pr_data()

    n_prs      = df["PR"].nunique()
    n_open     = int((df["Status"]=="Open").sum())        if "Status" in df.columns else 0
    n_rfp      = int((df["Status"]=="RFP").sum())          if "Status" in df.columns else 0
    n_proposed = int((df["Status"]=="PO Proposed").sum())  if "Status" in df.columns else 0
    n_created  = int((df["Status"]=="PO Created").sum())   if "Status" in df.columns else 0
    n_no_vnd   = int((df["Vendor"]=="").sum())

    m1,m2,m3,m4,m5,m6 = st.columns(6)
    m1.metric("PRs",         n_prs)
    m2.metric("Open Lines",  n_open)
    m3.metric("In RFP",      n_rfp)
    m4.metric("PO Proposed", n_proposed)
    m5.metric("PO Created",  n_created)
    m6.metric("No Vendor",   n_no_vnd)

    with st.expander("\U0001f4cb All PR Lines", expanded=True):
        dpr1, dpr2, dpr3 = st.columns([2, 2, 1.3])
        with dpr1:
            pr_from = st.date_input("PR date from", value=date.today() - timedelta(days=90),
                                    key="pr_list_from")
        with dpr2:
            # 9 months out, not just today — PRs can genuinely be dated
            # in the future (e.g. the customer-wave PRs), and a "to"
            # filter capped at today silently hid them from this table
            # entirely, which is confusing rather than just conservative.
            pr_to = st.date_input("PR date to", value=date.today() + timedelta(days=270),
                                  key="pr_list_to")
        with dpr3:
            st.write("")  # vertical alignment with the date inputs' labels
            show_all_status = st.toggle("All POs", value=False, key="pr_list_allstatus",
                help="Off (default): hide lines whose PO has already been sent to the "
                     "vendor and Created — the ones still worth looking at (Open, RFP, "
                     "or Proposed but not yet sent). On: show every line regardless of "
                     "status, including PO Created.")

        # PR_Date is a real column (pr_header.pr_date, added 2026-07-29) — no longer
        # parsed out of the PR number's own text. A PR predating that column still
        # works: init_schema() backfills it once from the number, same rule this
        # used to apply live on every read.
        pr_dates = pd.to_datetime(df["PR_Date"], errors="coerce").dt.date
        in_range = pr_dates.notna() & (pr_dates >= pr_from) & (pr_dates <= pr_to)
        if not show_all_status:
            in_range = in_range & (df["Status"] != "PO Created")
        df_shown = df[in_range]
        st.caption(f"{df_shown['PR'].nunique()} of {n_prs} PR(s), "
                  f"{len(df_shown)} of {len(df)} line(s) in this date range"
                  + ("." if show_all_status else " and not yet PO Created."))

        show = ["PR","PR_Date","Line","Requester","Code","Description","UOM","Qty","Status","Vendor"]
        if "PO_Number" in df_shown.columns: show += ["PO_Number","PO_Item"]
        st.dataframe(df_shown[show], use_container_width=True, hide_index=True, height=250)

    st.divider()
    one_to_one = st.toggle("1:1 PR \u2192 PO Lines", value=True,
        help="ON: one PO item per PR line\nOFF: sum qty per vendor+material")
    st.caption("Only **Open** lines are consolidated. Lines already in RFP, PO Proposed, or PO Created are skipped.")


    # ── Clear PO Proposals ───────────────────────────────────────────────────
    with st.expander("\u26a0\ufe0f  Clear PO Proposals", expanded=False):
        st.caption(
            "Reverts POs still at status **Proposed** — created by Consolidate but "
            "never actually sent to the vendor — back out of existence, and the PR "
            "lines that fed them back to **Open** so they rejoin the next Consolidate "
            "run. Anything already **Created** (sent to a vendor, or a Direct PO Entry / "
            "RFx-awarded PO) is left completely untouched — those are the vendor's or "
            "the ledger's business now, not something to silently undo. RFP data is "
            "untouched too.")
        if st.button("\U0001f5d1  Clear PO Proposals", type="secondary", key="clear_po_btn"):
            try:
                result = pr_consolidation.clear_po_proposals(DATA_FILE)
                st.session_state.pop("consolidate_result", None)
                st.cache_data.clear()
                st.success(f"\u2705 Cleared {result['po_count']} proposed PO(s). "
                          f"{result['pr_lines_reset']} PR line(s) reset to Open.")
            except Exception:
                st.error("\u274c Error:"); st.code(traceback.format_exc())

    if n_open == 0 and not st.session_state.get("consolidate_result"):
        st.warning("No Open lines to consolidate.")
        return

    if n_open > 0 and st.button("\U0001f504  Consolidate Open Lines into POs & RFP", type="primary"):
        try:
            with st.spinner("Consolidating…"):
                result = pr_consolidation.run(data_file=DATA_FILE, one_to_one=one_to_one,
                    contract_lookup_fn=ct.find_active_contract_for_item)
            st.cache_data.clear()
            st.session_state.consolidate_result = result
        except Exception:
            st.error("\u274c Consolidation failed:")
            st.code(traceback.format_exc())

    # Rendered from session_state, not gated behind the button's live return
    # value — a REAL bug fixed here, not a cosmetic one. Every button inside
    # this section (particularly "Create PO") used to live inside
    # `if st.button("Consolidate..."):`, which returns True for exactly one
    # script run: the one immediately following ITS OWN click. Clicking
    # "Create PO" triggers a fresh rerun in which the Consolidate button
    # wasn't clicked, so that condition evaluates False again — the whole
    # branch containing "Create PO" never executes THAT run, so the click
    # silently does nothing. This wasn't a test-harness artifact: it's how
    # Streamlit's execution model actually works, and it meant this button
    # could never fire in a real browser session either, from the very
    # first version of this page. Persisting the result means the tabs (and their buttons) keep rendering on every
    # subsequent rerun, including the one caused by clicking a button
    # inside them.
    result = st.session_state.get("consolidate_result")
    if result:
        for msg in result["messages"]: st.success(f"\u2713 {msg}")
        st.divider()

        r1,r2,r3 = st.columns(3)
        r1.metric("POs Created",  len(result["po_summary"]))
        r2.metric("Total PO Items",sum(p["lines"] for p in result["po_summary"]))
        r3.metric("RFP Lines",    result["rfp_count"])

        tab1,tab2,tab3,tab4 = st.tabs([
            "\U0001f500 Flow", "\U0001f5fa\ufe0f Map",
            "\U0001f4e6 POs + Create PO", "\U0001f4cb RFP"])

        with tab1: _flow(result)
        with tab2: _map(result["geo_data"])
        with tab3: _po_av(result)
        with tab4:
            rfp = result["rfp_items"]
            if rfp:
                st.dataframe(pd.DataFrame(rfp)[["rfp_number","mat_code","mat_desc",
                    "uom","total_qty","req_by_date","source_prs","req_depts",
                    "closing_date","status"]],
                    use_container_width=True, hide_index=True)
            else:
                st.info("No RFP lines.")

        if st.button("\u2715 Dismiss this result", key="consolidate_dismiss"):
            del st.session_state["consolidate_result"]
            st.rerun()


def clear_consolidation_data():
    """
    Moved into pr_consolidation.py (extraction, per the pilot plan) —
    this was ad-hoc UI code before, now delegated to the module that
    owns PO_Header/PO_Items/RFP/PR_Items. Same behavior as documented
    there: full reset of PO/RFP data, PR_Items status reverted to Open
    for 'PO Proposed', 'PO Created', and 'RFP'.

    Currently unused by this UI — "Clear PO Proposals" in
    page_consolidate() calls pr_consolidation.clear_po_proposals()
    instead, a version properly scoped to only status='Proposed' POs
    (see that function's docstring). A full unfiltered wipe never
    actually served the original what-if-analysis purpose this button
    was built for. This full-reset version is left in place, still
    correct and tested, for if a genuine need for one comes up.
    """
    return pr_consolidation.clear_consolidation_data(DATA_FILE)


def _po_av(result):
    po_items_all = result["po_items"]
    po_groups    = defaultdict(list)
    for it in po_items_all: po_groups[it["po_number"]].append(it)

    st.caption("The PO already exists in the system, at status **Proposed** — it was "
              "written the moment you clicked Consolidate, but isn't usable in Goods "
              "Receipt yet. This generates the downloadable file to actually send to "
              "the vendor (or upload to whatever system consumes it), and flips the "
              "PO to **Created** — only then can it be received against.")

    if st.button("\U0001f4e4  Create All POs", type="primary", key="av_all"):
        errors = []
        created_lines = 0
        for po_num, pols in sorted(po_groups.items()):
            p = pols[0]
            try:
                fname, fdata = po_export.make_av_bytes(po_num, p["supplier_id"], p["po_type"],
                    p["legal_entity"], p["purch_entity"], p["purch_group"],
                    p["currency"], p["plant_code"], pols)
                fpath = os.path.join(_DIR, fname)
                with open(fpath, "wb") as f: f.write(fdata)
                created_lines += pr_consolidation.mark_po_created(po_num, DATA_FILE)
            except Exception as ex:
                errors.append(f"{po_num}: {ex}")
        if errors:
            for e in errors: st.error(e)
        else:
            n = len(po_groups)
            st.cache_data.clear()
            st.success(f"\u2705 {n} AV file(s) saved to: {_DIR} — {n} PO(s) now Created "
                      f"({created_lines} PR line(s) advanced).")

    st.divider()

    for po_num, pols in sorted(po_groups.items()):
        p      = pols[0]
        vendor = p["supplier_id"]
        with st.expander(f"\U0001f6d2  {po_num}  —  {vendor}  ({len(pols)} items)",
                         expanded=True):
            st.dataframe(
                pd.DataFrame(pols)[["po_item","mat_code","mat_desc","uom","qty",
                    "deliv_date","deliv_loc","source_pr","source_pr_line","req_dept"]].rename(
                    columns={"po_item":"#","mat_code":"Code","mat_desc":"Description",
                             "uom":"UOM","qty":"Qty","deliv_date":"Delivery",
                             "deliv_loc":"Location","source_pr":"Source PR",
                             "source_pr_line":"PR Line","req_dept":"Dept"}),
                use_container_width=True, hide_index=True)
            if st.button(f"\U0001f4e4  Create PO {po_num}", key=f"av_{po_num}"):
                try:
                    fname, fdata = po_export.make_av_bytes(po_num, vendor,
                        p["po_type"], p["legal_entity"], p["purch_entity"],
                        p["purch_group"], p["currency"], p["plant_code"], pols)
                    pr_consolidation.mark_po_created(po_num, DATA_FILE)
                    st.cache_data.clear()
                    st.success(f"\u2705 {fname} — {po_num} is now **Created** and can be "
                              "received in Goods Receipt.")
                    st.download_button(f"\u2b07 Download {fname}", data=fdata,
                        file_name=fname, mime=XLSX_MIME, key=f"dl_{po_num}")
                except Exception:
                    st.error("Export failed:"); st.code(traceback.format_exc())


def _flow(result):
    """Curved SVG arrow flow diagram."""
    pri  = result["pr_items"]; poi = result["po_items"]
    rfp  = result["rfp_items"]; prh = result["pr_headers"]
    prg  = defaultdict(list)
    for it in pri: prg[it["pr_number"]].append(it)
    pog  = defaultdict(list)
    for it in poi: pog[it["po_number"]].append(it)
    lnk  = {}
    for it in poi:
        for a,b in zip(str(it["source_pr"]).split(", "),
                       str(it["source_pr_line"]).split(", ")):
            lnk[(a.strip(),b.strip())] = (it["po_number"],it["po_item"])
    RH=20;PD=8;GH=32;GP=10;LX=10;LW=255;RX=525;RW=255;SW=790
    def gh(n): return GH+PD+n*RH+PD
    ppos={}; y=20
    for pr in sorted(prg):
        its=prg[pr]; ppos[pr]={"y":y,"h":gh(len(its)),"items":its}
        for i,it in enumerate(its): ppos[pr][f"L{it['pr_line']}"]=y+GH+PD+i*RH+RH//2
        y+=gh(len(its))+GP
    VC={"GARMY":"#4F46E5","HINDMED":"#059669","MANI":"#D97706",
        "Pharmalines":"#DC2626","dntl":"#7C3AED"}
    DC=["#2563EB","#0891B2","#D97706","#7C3AED","#059669"]
    opos={}; pcm={}; y=20
    for i,(po,pols) in enumerate(sorted(pog.items())):
        v=pols[0]["supplier_id"]; c=VC.get(v,DC[i%len(DC)]); pcm[po]=c
        opos[po]={"y":y,"h":gh(len(pols)),"items":pols,"v":v,"c":c}
        for j,pol in enumerate(pols): opos[po][f"#{pol['po_item']}"]=y+GH+PD+j*RH+RH//2
        y+=gh(len(pols))+GP
    ry=y; rh=gh(len(rfp))
    rly={k:ry+GH+PD+k*RH+RH//2 for k in range(len(rfp))}
    SH=max(sum(v["h"]+GP for v in ppos.values())+20,y+rh+GP)+20
    def e(s): return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    def sh(s,n=28): s=str(s); return s[:n-1]+"\u2026" if len(s)>n else s
    svg=[]
    for pr,info in ppos.items():
        for it in info["items"]:
            k=(it["pr_number"],str(it["pr_line"])); sy=info.get(f"L{it['pr_line']}",0); sx=LX+LW
            tgt=lnk.get(k)
            if tgt:
                po,pi=tgt; ty=opos.get(po,{}).get(f"#{pi}",0); tx=RX
                if sy and ty:
                    dx=(tx-sx)*0.45; col=pcm.get(po,"#059669")
                    svg.append(f'<path d="M{sx},{sy} C{sx+dx},{sy} {tx-dx},{ty} {tx},{ty}" fill="none" stroke="{col}" stroke-width="1.5" opacity="0.65" marker-end="url(#ahp)"><title>{e(pr)} L{it["pr_line"]} \u2192 {po} #{pi}</title></path>')
            else:
                ri=next((k for k,r in enumerate(rfp) if r["mat_code"]==it["mat_code"]),None)
                if ri is not None and sy:
                    ty=rly.get(ri,0); tx=RX; dx=(tx-sx)*0.45
                    svg.append(f'<path d="M{sx},{sy} C{sx+dx},{sy} {tx-dx},{ty} {tx},{ty}" fill="none" stroke="#DC2626" stroke-width="1.5" opacity="0.65" marker-end="url(#ahr)"><title>{e(pr)} L{it["pr_line"]} \u2192 RFP</title></path>')
    PRC={"PR-001":"#1D4ED8","PR-002":"#047857","PR-003":"#92400E"}
    for pr,info in ppos.items():
        col=PRC.get(pr,"#1E3A5F"); h=prh.get(pr,{}); req=h.get("req_name",""); dept=h.get("req_dept",""); y0=info["y"]; h0=info["h"]
        svg+=[f'<rect x="{LX}" y="{y0}" width="{LW}" height="{h0}" rx="6" fill="{col}10" stroke="{col}" stroke-width="1.2"/>',
              f'<text x="{LX+8}" y="{y0+14}" font-size="11" font-weight="600" fill="{col}">{e(pr)}</text>',
              f'<text x="{LX+8}" y="{y0+26}" font-size="9" fill="{col}99">{e(req)} {e(dept)}</text>',
              f'<line x1="{LX}" y1="{y0+30}" x2="{LX+LW}" y2="{y0+30}" stroke="{col}" stroke-width="0.5" stroke-dasharray="4,2"/>']
        for i,it in enumerate(info["items"]):
            iy=y0+GH+PD+i*RH; k=(it["pr_number"],str(it["pr_line"]))
            tgt=lnk.get(k); bc=pcm.get(tgt[0],"#DC2626") if tgt else "#DC2626"
            bl=tgt[0][-4:] if tgt else "RFP"; bg=f'fill="{col}08"' if i%2==0 else 'fill="none"'
            svg+=[f'<rect x="{LX+2}" y="{iy}" width="{LW-4}" height="{RH-1}" rx="2" {bg}/>',
                  f'<text x="{LX+6}" y="{iy+13}" font-size="9" font-weight="500" fill="{col}">L{it["pr_line"]}</text>',
                  f'<text x="{LX+24}" y="{iy+13}" font-size="9" fill="#334155">{e(sh(it["mat_desc"]))}</text>',
                  f'<rect x="{LX+LW-40}" y="{iy+3}" width="36" height="12" rx="3" fill="{bc}15" stroke="{bc}" stroke-width="0.5"/>',
                  f'<text x="{LX+LW-22}" y="{iy+12}" font-size="8" text-anchor="middle" fill="{bc}" font-weight="600">{e(bl)}</text>']
    for po,info in sorted(opos.items()):
        col=info["c"]; v=info["v"]; y0=info["y"]; h0=info["h"]
        svg+=[f'<rect x="{RX}" y="{y0}" width="{RW}" height="{h0}" rx="6" fill="{col}10" stroke="{col}" stroke-width="1.2"/>',
              f'<text x="{RX+8}" y="{y0+14}" font-size="11" font-weight="600" fill="{col}">{e(po)}</text>',
              f'<text x="{RX+8}" y="{y0+26}" font-size="9" fill="{col}99">{e(v)}</text>',
              f'<line x1="{RX}" y1="{y0+30}" x2="{RX+RW}" y2="{y0+30}" stroke="{col}" stroke-width="0.5" stroke-dasharray="4,2"/>']
        for j,pol in enumerate(info["items"]):
            jy=y0+GH+PD+j*RH; bg=f'fill="{col}08"' if j%2==0 else 'fill="none"'
            svg+=[f'<rect x="{RX+2}" y="{jy}" width="{RW-4}" height="{RH-1}" rx="2" {bg}/>',
                  f'<text x="{RX+6}" y="{jy+13}" font-size="9" font-weight="500" fill="{col}">#{pol["po_item"]}</text>',
                  f'<text x="{RX+26}" y="{jy+13}" font-size="9" fill="#334155">{e(sh(pol["mat_desc"]))}</text>',
                  f'<text x="{RX+RW-4}" y="{jy+13}" font-size="8" text-anchor="end" fill="{col}88">{e(pol["source_pr"])} L{e(str(pol["source_pr_line"]))}</text>']
    RC="#DC2626"
    svg+=[f'<rect x="{RX}" y="{ry}" width="{RW}" height="{rh}" rx="6" fill="{RC}10" stroke="{RC}" stroke-width="1.2"/>',
          f'<text x="{RX+8}" y="{ry+14}" font-size="11" font-weight="600" fill="{RC}">RFP ({len(rfp)} items)</text>',
          f'<line x1="{RX}" y1="{ry+30}" x2="{RX+RW}" y2="{ry+30}" stroke="{RC}" stroke-width="0.5" stroke-dasharray="4,2"/>']
    for k,ri in enumerate(rfp):
        ky=ry+GH+PD+k*RH; bg=f'fill="{RC}08"' if k%2==0 else 'fill="none"'
        svg+=[f'<rect x="{RX+2}" y="{ky}" width="{RW-4}" height="{RH-1}" rx="2" {bg}/>',
              f'<text x="{RX+6}" y="{ky+13}" font-size="9" fill="#334155">{e(sh(ri["mat_desc"],36))}</text>']
    html=(f'<div style="overflow-x:auto;background:#fff;border-radius:10px;padding:14px;border:1px solid #E2E8F0">' +
          f'<div style="display:flex;gap:10px;margin-bottom:10px;font-family:Arial,sans-serif;flex-wrap:wrap;align-items:center">' +
          f'<b style="font-size:12px;color:#0F172A">PR \u2192 PO consolidation flow</b>' +
          f'<span style="font-size:11px;background:#F0FDF4;color:#166534;padding:2px 8px;border-radius:4px;border:1px solid #86EFAC">\u25cf PO lines</span>' +
          f'<span style="font-size:11px;background:#FEF2F2;color:#991B1B;padding:2px 8px;border-radius:4px;border:1px solid #FECACA">\u25cf RFP</span>' +
          f'<span style="font-size:11px;color:#64748B">Hover arrows for details</span></div>' +
          f'<svg viewBox="0 0 {SW} {SH}" xmlns="http://www.w3.org/2000/svg" style="width:100%;min-width:{SW}px;height:{SH}px;display:block">' +
          f'<defs><marker id="ahp" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#059669"/></marker>' +
          f'<marker id="ahr" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#DC2626"/></marker></defs>' +
          "".join(svg) + '</svg></div>')
    st.components.v1.html(html, height=SH+80, scrolling=True)


def _map(geo):
    vj=json.dumps(geo["vendors"]); dj=json.dumps(geo["deliveries"]); rj=json.dumps(geo["routes"])
    html=f"""<!DOCTYPE html><html><head><meta charset="utf-8"/>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"></script>
<style>body{{margin:0}}#map{{width:100%;height:460px}}
.leg{{background:rgba(255,255,255,.95);border-radius:8px;padding:10px 14px;
border:1px solid #E2E8F0;font-size:12px;color:#374151;line-height:2}}</style></head><body>
<div id="map"></div><script>
const V={vj},D={dj},R={rj};
const map=L.map("map",{{preferCanvas:true,attributionControl:false}});
L.tileLayer("https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png",{{attribution:"© OpenStreetMap",maxZoom:18}}).addTo(map);
V.forEach(v=>{{if(!v.lat||!v.lng)return;
L.marker([v.lat,v.lng],{{icon:L.divIcon({{html:`<div style="width:14px;height:14px;background:${{v.color}};border:2px solid #fff;border-radius:3px;box-shadow:0 1px 4px rgba(0,0,0,.25)"></div>`,className:"",iconAnchor:[7,7],popupAnchor:[0,-10]}})}}).addTo(map)
.bindPopup(`<b style="color:${{v.color}}">${{v.id}}</b><br><small>${{v.lat.toFixed(4)}},${{v.lng.toFixed(4)}}</small>`);
}});
D.forEach(d=>{{if(!d.lat||!d.lng)return;
L.marker([d.lat,d.lng],{{icon:L.divIcon({{html:`<div style="width:0;height:0;border-left:9px solid transparent;border-right:9px solid transparent;border-bottom:16px solid #0284C7;filter:drop-shadow(0 1px 2px rgba(0,0,0,.25))"></div>`,className:"",iconAnchor:[9,16],popupAnchor:[0,-16]}})}}).addTo(map)
.bindPopup(`<b style="color:#0284C7">${{d.id}}</b><br><small>${{d.lat.toFixed(4)}},${{d.lng.toFixed(4)}}</small>`);
}});
function arc(a,b,c,d,n=50){{const p=[];for(let i=0;i<=n;i++){{const t=i/n;p.push([a+(c-a)*t+Math.sin(Math.PI*t)*Math.sqrt((c-a)**2+(d-b)**2)*.18,b+(d-b)*t]);}}return p;}}
R.forEach(r=>{{const v=V.find(x=>x.id===r.vendor),d=D.find(x=>x.id===r.deliv_loc);
if(!v||!d||!v.lat||!d.lat)return;
L.polyline(arc(v.lat,v.lng,d.lat,d.lng),{{color:r.color,weight:2.5,opacity:.75,dashArray:"6,4"}}).addTo(map)
.bindPopup(`<b>${{r.po_number}}</b><br>${{r.vendor}} \u2192 ${{r.deliv_loc}}<br><small>${{r.items}} item(s)</small>`);
}});
const pts=[...V.filter(v=>v.lat).map(v=>[v.lat,v.lng]),...D.filter(d=>d.lat).map(d=>[d.lat,d.lng])];
// Fit map to data points. Call multiple times with increasing delays
// because the iframe container may not be fully sized on first attempt.
function doFit(){{
  map.invalidateSize(true);
  if(pts.length){{
    map.fitBounds(pts,{{padding:[40,40],maxZoom:8}});
  }}
}}
setTimeout(doFit,200);
setTimeout(doFit,600);
setTimeout(doFit,1200);
// This map usually lives inside a Streamlit tab. Streamlit hides inactive
// tabs (zero-size container), so if this tab isn't the one open on first
// render, the timers above fire against a 0x0 box and fitBounds computes
// a wrong view that never self-corrects. Watch for the container actually
// getting real dimensions (i.e. its tab becoming visible) and re-fit once
// when that happens, on top of the fixed-delay attempts above.
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
V.filter(v=>v.lat).map(v=>`<div><span style="width:12px;height:12px;background:${{v.color}};display:inline-block;border-radius:2px;margin-right:6px;vertical-align:middle"></span>${{v.id}}</div>`).join("")+
`<div><span style="display:inline-block;width:0;height:0;border-left:7px solid transparent;border-right:7px solid transparent;border-bottom:12px solid #0284C7;margin-right:6px;vertical-align:middle"></span>Delivery</div>`;return d;}};
leg.addTo(map);
</script></body></html>"""
    st.components.v1.html(html, height=500)
    st.caption("Squares=suppliers  \u00b7  Triangles=delivery  \u00b7  Arcs=PO routes  \u00b7  "
               "Use the **Fit** button (top-right of the map) any time to re-frame every point.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — Direct PO
# ══════════════════════════════════════════════════════════════════════════════
def page_direct_po():
    st.markdown("## \U0001f4e4 Direct PO Entry")
    st.caption("For purchase orders placed outside the normal requisition flow — for "
               "example, an urgent one-off buy. Works exactly like any other PO for "
               "receiving, contracts, and accounting.")
    st.divider()
    left, right = st.columns([5,4], gap="large")
    with left:
        st.markdown("#### \U0001f50d Item Picker"); render_item_picker("po")
    with right:
        st.markdown("#### \U0001f4c4 PO Details")
        c1,c2 = st.columns(2)
        with c1:
            if "dpo_n" not in st.session_state:
                try:
                    st.session_state.dpo_n = pr_consolidation.next_po_number()
                except Exception:
                    st.session_state.dpo_n = "PO-DIRECT"
            po_num  = st.text_input("PO Number", key="dpo_n",
                                    help="Auto-suggested — edit if you need a specific number.")
            vendors = vo.list_vendors(include_inactive=False)
            if vendors:
                vendor_labels = {v["Vendor_ID"]: f"{v['Vendor_Name']}" for v in vendors}
                sup_id = st.selectbox("Supplier", list(vendor_labels.keys()),
                    format_func=lambda k: vendor_labels[k], key="dpo_s")
                sup_name = vendor_labels[sup_id]
            else:
                st.warning("No approved vendors yet — onboard one in Vendor Onboarding first.")
                sup_id, sup_name = "", ""
        with c2:
            d_date  = st.date_input("Delivery Date",  value=date.today(), key="dpo_dd")
            locs    = load_delivery_locs()
            if locs:
                loc_labels = {l["id"]: l["name"] for l in locs}
                dloc = st.selectbox("Delivery Location", list(loc_labels.keys()),
                    format_func=lambda k: loc_labels[k], key="dpo_dl")
                dloc_geo = next((l["geo"] for l in locs if l["id"] == dloc), "")
            else:
                dloc = st.text_input("Delivery Location", value="", key="dpo_dl2")
                dloc_geo = ""
        with st.expander("\u2699\ufe0f Org defaults"):
            oa,ob = st.columns(2)
            with oa:
                d_type  = st.text_input("PO Type",   value=od.get_default("PO Type"),      key="dpo_t")
                d_legal = st.text_input("Legal Ent.",value=od.get_default("Legal Entity"),  key="dpo_le")
                d_pe    = st.text_input("Purch.Ent.",value=od.get_default("Purchasing Entity"),  key="dpo_pe")
            with ob:
                d_pg    = st.text_input("Purch.Grp.",value=od.get_default("Purchasing Group"),  key="dpo_pg")
                d_cur   = st.text_input("Currency",  value=od.get_default("Currency"),     key="dpo_cur")
                d_plt   = st.text_input("Plant",     value=od.get_default("Plant"),key="dpo_plt")
        st.divider()
        staged = st.session_state.get("po_staged",[])
        if staged:
            st.markdown("**Selected Items**")
            render_staged_table("po")
            st.markdown("")
            if st.button("\U0001f4e4  Create PO", type="primary",
                         use_container_width=True, key="dpo_send",
                         disabled=not sup_id):
                lines = [i for i in staged if i.get("qty",0)>0]
                if not lines:
                    st.error("No items with qty > 0.")
                elif pr_consolidation.get_po_header(po_num):
                    st.error(f"\u274c {po_num} already exists — pick a different PO number.")
                else:
                    try:
                        # Record the actual PO first — Goods Receipt, Contracts, and
                        # Accounting all read from PO_Header/PO_Items, and until this
                        # call was added, Direct PO Entry never wrote there at all: it
                        # only ever produced the downloadable file below, so a PO
                        # created here was invisible everywhere else in the system.
                        db_lines = [{"mat_code": l["code"], "mat_desc": l["desc"],
                                    "uom": l["uom"], "qty": l["qty"], "unit_price": l["price"],
                                    "deliv_date": str(d_date), "deliv_loc": dloc,
                                    "deliv_geo": dloc_geo, "source_pr": "", "source_pr_line": "",
                                    "req_id": "", "req_dept": "", "project_id": ""} for l in lines]
                        pr_consolidation.insert_po(po_num,
                            {"po_type": d_type, "legal_entity": d_legal, "purch_entity": d_pe,
                             "purch_group": d_pg, "currency": d_cur, "plant_code": d_plt,
                             "supplier_id": sup_id, "supplier_name": sup_name,
                             "supplier_geo": ""}, db_lines)

                        po_lns = [{"mat_code":l["code"],"mat_desc":l["desc"],
                            "uom":l["uom"],"qty":l["qty"],
                            "deliv_date":str(d_date),"deliv_loc":dloc} for l in lines]
                        fname, fdata = po_export.make_av_bytes(po_num, sup_id, d_type,
                            d_legal, d_pe, d_pg, d_cur, d_plt, po_lns)
                        pr_consolidation.mark_po_created(po_num, DATA_FILE)
                        st.success(f"\u2705 {po_num} recorded ({len(lines)} line(s)), "
                                  f"now **Created**, and {fname} generated.")
                        st.download_button(f"\u2b07 Download {fname}",
                            data=fdata, file_name=fname, mime=XLSX_MIME, key="dpo_dl_btn")
                        st.session_state.po_staged=[]; st.session_state.po_selected=set()
                        del st.session_state["dpo_n"]
                        st.cache_data.clear()
                    except Exception:
                        st.error("\u274c PO creation failed:"); st.code(traceback.format_exc())
            if st.button("\U0001f5d1 Clear", use_container_width=True, key="dpo_clr"):
                st.session_state.po_staged=[]; st.session_state.po_selected=set(); st.rerun()
        else:
            st.markdown("<div style='color:#94A3B8;padding:36px;text-align:center;"
                "border:1px dashed #CBD5E1;border-radius:8px'>"
                "\U0001f4ed No items staged</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — Autonomous Categorization
# ══════════════════════════════════════════════════════════════════════════════
def page_categorization():
    st.markdown("## \U0001f3f7\ufe0f Autonomous Categorization")
    st.caption("Automatically suggests a category for each item, showing the keywords "
               "behind each suggestion so low-confidence items are easy to spot and "
               "review.")
    st.divider()

    s = categorization.stats()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Items", s["total"])
    m2.metric("Categorized", s["categorized"])
    m3.metric("Uncategorized", s["uncategorized"])
    pct = (s["categorized"] / s["total"] * 100) if s["total"] else 0
    m4.metric("Coverage", f"{pct:.0f}%")

    if s["by_category"]:
        st.markdown("#### Category distribution")
        chart_df = pd.DataFrame(
            {"Category": list(s["by_category"].keys()),
             "Items": list(s["by_category"].values())}
        ).sort_values("Items", ascending=False).set_index("Category")
        st.bar_chart(chart_df, height=220)

    st.divider()

    if s["uncategorized"] == 0:
        st.success("\u2705 Every active item already has a category.")
        return

    if st.button("\U0001f916  Run Autonomous Categorization",
                 type="primary", use_container_width=True, key="cat_run"):
        with st.spinner("Classifying uncategorized items…"):
            st.session_state.cat_preview = categorization.preview_categorization(only_blank=True)
        st.rerun()

    preview = st.session_state.get("cat_preview")
    if not preview:
        st.markdown(
            "<div style='color:#94A3B8;padding:36px;text-align:center;"
            "border:1px dashed #CBD5E1;border-radius:8px'>"
            "\U0001f916 Run the classifier to see suggested categories</div>",
            unsafe_allow_html=True)
        return

    conf_counts = defaultdict(int)
    for p in preview: conf_counts[p["confidence"]] += 1
    st.caption(
        f"{len(preview)} item(s) classified — "
        f"\U0001f7e2 {conf_counts.get('High',0)} High · "
        f"\U0001f7e1 {conf_counts.get('Medium',0)} Medium · "
        f"\U0001f534 {conf_counts.get('Low',0)} Low confidence (needs review)")

    tab_hi, tab_lo = st.tabs([
        f"\u2705 High / Medium confidence ({conf_counts.get('High',0)+conf_counts.get('Medium',0)})",
        f"\u26a0\ufe0f Needs review ({conf_counts.get('Low',0)})"])

    def _editor(rows, key):
        if not rows:
            st.info("Nothing here.")
            return []
        df = pd.DataFrame(rows)[["code","desc","category","subcategory",
                                  "confidence","matched_keywords"]]
        cfg = {
            "code":       st.column_config.TextColumn("Code", disabled=True, width=95),
            "desc":       st.column_config.TextColumn("Description", disabled=True, width=240),
            "category":   st.column_config.SelectboxColumn("Category",
                              options=categorization.CATEGORY_CHOICES, width=190),
            "subcategory":st.column_config.TextColumn("Sub-Category", width=190),
            "confidence": st.column_config.TextColumn("Conf.", disabled=True, width=70),
            "matched_keywords": st.column_config.TextColumn("Matched on", disabled=True, width=180),
        }
        edited = st.data_editor(df, use_container_width=True, hide_index=True,
                                 column_config=cfg, key=key, height=min(420, 42+35*len(rows)))
        out = []
        for i, row in edited.iterrows():
            out.append({"code": rows[i]["code"], "category": row["category"],
                        "subcategory": row["subcategory"]})
        return out

    hi_rows = [p for p in preview if p["confidence"] in ("High","Medium")]
    lo_rows = [p for p in preview if p["confidence"] == "Low"]
    with tab_hi:
        edited_hi = _editor(hi_rows, "cat_editor_hi")
    with tab_lo:
        st.caption("These didn't match closely enough — pick a category manually or leave as "
                   "Uncategorized.")
        edited_lo = _editor(lo_rows, "cat_editor_lo")

    st.divider()
    if st.button("\u2705  Apply to Item Master", type="primary",
                 use_container_width=True, key="cat_apply"):
        try:
            n = categorization.apply_categorization(edited_hi + edited_lo)
            st.success(f"\u2705 {n} item(s) categorized.")
            st.session_state.cat_preview = None
            st.cache_data.clear()
            st.rerun()
        except Exception:
            st.error("\u274c Error applying categorization:")
            st.code(traceback.format_exc())


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — RFx Management (vendor-centric)
# ══════════════════════════════════════════════════════════════════════════════
def page_rfx():
    st.markdown("## \U0001f4e8 RFx Management")
    st.caption("Invite vendors to quote on multiple items at once, compare responses, pick "
               "winners per line, and issue one consolidated purchase order per vendor.")
    st.divider()

    rx = rfx.stats()
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Open Lines", rx["open"])
    m2.metric("Pending PO", rx["pending_po"])
    m3.metric("Awarded", rx["awarded"])
    m4.metric("Quotes", rx["quotes"])
    m5.metric("Invitations", rx["invitations"])
    st.divider()

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "1\ufe0f\u20e3 Select & Invite", "2\ufe0f\u20e3 Enter Quotes",
        "3\ufe0f\u20e3 Select Winners", f"4\ufe0f\u20e3 Issue POs ({rx['pending_po']})",
        f"\U0001f4dc Award History ({rx['awarded']})"])

    # ── TAB 1 — pick RFP lines, suggest vendors, generate the RFQ document ──────
    with tab1:
        open_rfps = rfx.get_open_rfps()
        if not open_rfps:
            st.success("\u2705 No open RFP lines — everything is quoted or awarded.")
        else:
            # Category lookup per line, for the filter below — a 30+ line RFP
            # batch is much easier to work with narrowed to "just Electronics"
            # or similar before picking which lines to source together.
            cat_by_mat = {}
            for r in open_rfps:
                if r["mat_code"] not in cat_by_mat:
                    item = po_export.get_item_by_code(r["mat_code"], active_only=False)
                    cat_by_mat[r["mat_code"]] = (item["category"] if item else "") or "Uncategorized"
            categories = sorted({cat_by_mat[r["mat_code"]] for r in open_rfps})

            f1, f2 = st.columns([2, 3])
            with f1:
                cat_filter = st.multiselect("Filter by category", categories,
                    default=categories, key="rfx_cat_filter")
            filtered_rfps = [r for r in open_rfps if cat_by_mat[r["mat_code"]] in cat_filter]

            labels = {r["rfp_number"]: f"{r['rfp_number']} — {r['mat_desc']} (qty {r['total_qty']:g})"
                      for r in filtered_rfps}

            with f2:
                st.caption(f"{len(filtered_rfps)} of {len(open_rfps)} open line(s) match this filter.")
                b1, b2 = st.columns(2)
                with b1:
                    if st.button("\u2611\ufe0f Select all filtered", use_container_width=True, key="rfx_selall"):
                        st.session_state.rfx_batch_sel = list(labels.keys())
                        st.rerun()
                with b2:
                    if st.button("\u2b1c Clear selection", use_container_width=True, key="rfx_clearall"):
                        st.session_state.rfx_batch_sel = []
                        st.rerun()

            # Selection is filter-independent once made (narrowing the category
            # filter doesn't silently drop lines you already picked under a
            # wider filter) — but the widget's own options list still needs to
            # only offer currently-filtered lines, so any prior pick outside
            # the current filter is dropped from what's shown/selectable here.
            prior = st.session_state.get("rfx_batch_sel", list(labels.keys()))
            default_sel = [n for n in prior if n in labels]
            chosen_nums = st.multiselect(
                "RFP lines to source together", list(labels.keys()),
                default=default_sel, format_func=lambda k: labels[k], key="rfx_batch_sel")
            chosen = [r for r in open_rfps if r["rfp_number"] in chosen_nums]

            if chosen:
                st.markdown(f"#### \U0001f3af Suggested vendors for these {len(chosen)} line(s)")
                st.caption("Ranked by category-supply history + proximity; averaged across the selected lines.")
                suggestions = rfx.suggest_vendors_for_lines(chosen, top_n=6)
                vendor_lookup = {v["id"]: v for v in rfx.list_vendors()}
                if suggestions:
                    sdf = pd.DataFrame(suggestions)
                    sdf["scorecard_display"] = sdf.apply(
                        lambda r: (f"{r['scorecard_score']:g}" +
                                  (" (Low Vol.)" if r["scorecard_low_volume"] else ""))
                                  if r["scorecard_score"] is not None else "—", axis=1)
                    sdf = sdf[["id","name","city","lines_matched","scorecard_display","reason"]]
                    sdf.columns = ["Vendor","Name","City","Lines matched","Scorecard","Why"]
                    st.dataframe(sdf, use_container_width=True, hide_index=True)
                    st.caption("Scorecard is shown for context, not used to rank this list — "
                              "affinity + proximity answers who can genuinely supply this; the "
                              "scorecard is a separate, real signal for you to weigh yourself.")
                else:
                    st.info("No active vendors found in Vendor_Master.")

                invite_ids = st.multiselect(
                    "Vendor(s) to invite — one RFQ document per vendor, covering all selected lines",
                    [s["id"] for s in suggestions], default=[s["id"] for s in suggestions[:2]],
                    key="rfx_invite_pick")

                c1, c2 = st.columns(2)
                with c1:
                    gen = st.button("\U0001f4c4  Generate RFQ Document(s)", type="primary",
                                    use_container_width=True, key="rfx_gen",
                                    disabled=(len(invite_ids) == 0))
                with c2:
                    also_sim = st.checkbox("Also simulate their response (demo)", key="rfx_also_sim")

                if gen:
                    generated = []
                    for vid in invite_ids:
                        v = vendor_lookup.get(vid, {"name": vid, "email": ""})
                        try:
                            fname, fbytes, _ = rfx.invite_and_generate_rfq(
                                vid, v["name"], v.get("email",""), chosen)
                            generated.append({"vendor_id": vid, "vendor_name": v["name"],
                                              "filename": fname, "bytes": fbytes})
                            if also_sim:
                                rfx.simulate_quotes_batch(chosen, vid)
                        except Exception:
                            st.error(f"Failed for {vid}:"); st.code(traceback.format_exc())
                    st.session_state.rfx_generated_docs = generated
                    st.cache_data.clear()

                # Rendered from session_state (not inside `if gen`) so the download
                # buttons survive the rerun that clicking any one of them triggers —
                # otherwise every download after the first would wipe the rest.
                docs = st.session_state.get("rfx_generated_docs")
                if docs:
                    st.success(f"\u2705 {len(docs)} RFQ document(s) ready — "
                              f"{'response simulated' if also_sim else 'awaiting vendor response'}.")
                    if len(docs) > 1:
                        zip_bytes = _build_zip([{"filename": d["filename"], "bytes": d["bytes"]} for d in docs])
                        st.download_button("\U0001f4e6  Download All (ZIP)", data=zip_bytes,
                            file_name="RFQ_documents.zip", mime="application/zip",
                            use_container_width=True, key="rfx_dl_all")
                    for d in docs:
                        st.download_button(f"\u2b07 {d['filename']} — {d['vendor_name']}",
                            data=d["bytes"], file_name=d["filename"], mime=XLSX_MIME,
                            key=f"rfx_dl_{d['vendor_id']}")

                st.divider()
                with st.expander("\u26a1 Skip RFx — assign directly to a vendor"):
                    st.caption("RFx gets you a competitive price, but it isn't mandatory. For "
                              "low-value or routine items, assign straight to a vendor at list "
                              "price (editable) — it lands in **Select Winners** exactly like a "
                              "normal award, ready to issue as a PO or contract.")
                    if not suggestions:
                        st.info("No suggested vendor available, but you can still pick any "
                               "approved vendor below.")
                    all_vendors = rfx.list_vendors()
                    suggested_ids = {s["id"] for s in suggestions}
                    # Suggested vendors first (labeled with why they were suggested), then
                    # every other approved vendor — the system's pick is the default, not
                    # the only option.
                    ranked = suggestions + [
                        {"id": v["id"], "name": v["name"], "reason": "not suggested for this line"}
                        for v in all_vendors if v["id"] not in suggested_ids]
                    if ranked:
                        vendor_labels = {v["id"]: f"{v['name']} ({v['id']}) — {v['reason']}" for v in ranked}
                        sel_vendor_id = st.selectbox("Assign to", list(vendor_labels.keys()),
                            format_func=lambda k: vendor_labels[k], key="rfx_skip_vendor_sel")
                        chosen_vendor = next(v for v in ranked if v["id"] == sel_vendor_id)
                        rows = []
                        for r in chosen:
                            d = rfx.get_item_master_defaults(r["mat_code"])
                            rows.append({"RFP": r["rfp_number"], "Description": r["mat_desc"],
                                        "Qty": r["total_qty"], "List Price": d["price"] or 0.0})
                        edited = st.data_editor(pd.DataFrame(rows), use_container_width=True, hide_index=True,
                            column_config={
                                "RFP": st.column_config.TextColumn(disabled=True),
                                "Description": st.column_config.TextColumn(disabled=True, width=220),
                                "Qty": st.column_config.NumberColumn(disabled=True, width=60),
                                "List Price": st.column_config.NumberColumn(format="\u20b9%.2f"),
                            }, key="rfx_skip_editor")
                        if st.button(f"\u26a1 Assign {len(chosen)} line(s) to {chosen_vendor['name']}",
                                     type="primary", key="rfx_skip_btn"):
                            prices = {row["RFP"]: float(row["List Price"]) for _, row in edited.iterrows()}
                            rfx.quick_assign(chosen, chosen_vendor["id"], chosen_vendor["name"], prices)
                            st.cache_data.clear()
                            st.success(f"\u2705 Assigned to {chosen_vendor['name']} — already marked as "
                                      "winner. Head to **Issue POs** to finish.")
                            st.rerun()

    # ── TAB 2 — vendor responds: batch quote entry ───────────────────────────────
    with tab2:
        invited_vendor_ids = sorted({i["vendor_id"] for i in rfx.get_invitations()})
        vendor_lookup = {v["id"]: v for v in rfx.list_vendors()}
        if not invited_vendor_ids:
            st.info("No vendors invited yet — do that in the first tab.")
        else:
            vid = st.selectbox("Vendor", invited_vendor_ids,
                format_func=lambda k: vendor_lookup.get(k,{}).get("name", k), key="rfx_qvendor")
            vname = vendor_lookup.get(vid, {}).get("name", vid)
            their_invites = {i["rfp_number"] for i in rfx.get_invitations(vendor_id=vid)}
            open_nums = {r["rfp_number"] for r in rfx.get_open_rfps()}
            pending_lines = [r for r in rfx.get_open_rfps() if r["rfp_number"] in their_invites & open_nums]

            if not pending_lines:
                st.success(f"No open lines awaiting a quote from {vname}.")
            else:
                st.caption(f"{vname} was invited to quote on {len(pending_lines)} open line(s):")
                rows = [{"RFP": r["rfp_number"], "Description": r["mat_desc"], "UOM": r["uom"],
                         "Qty": r["total_qty"], "Price": None, "Lead Time (d)": None, "MOQ": 1,
                         "Notes": ""} for r in pending_lines]
                edited = st.data_editor(pd.DataFrame(rows), use_container_width=True, hide_index=True,
                    column_config={
                        "RFP": st.column_config.TextColumn(disabled=True),
                        "Description": st.column_config.TextColumn(disabled=True, width=220),
                        "UOM": st.column_config.TextColumn(disabled=True, width=60),
                        "Qty": st.column_config.NumberColumn(disabled=True, width=60),
                        "Price": st.column_config.NumberColumn(format="\u20b9%.2f"),
                        "Lead Time (d)": st.column_config.NumberColumn(min_value=0, step=1),
                        "MOQ": st.column_config.NumberColumn(min_value=1, step=1),
                    }, key="rfx_quote_editor")

                c1, c2 = st.columns(2)
                with c1:
                    if st.button("\U0001f3b2  Simulate Response (Demo)", use_container_width=True, key="rfx_sim2"):
                        rfx.simulate_quotes_batch(pending_lines, vid)
                        st.cache_data.clear()
                        st.success(f"Simulated {len(pending_lines)} line(s) for {vname}.")
                        st.rerun()
                with c2:
                    if st.button("\U0001f4be  Save Quotes", type="primary", use_container_width=True, key="rfx_save2"):
                        line_quotes = []
                        for _, row in edited.iterrows():
                            if row["Price"] is not None and float(row["Price"]) > 0:
                                line_quotes.append({"rfp_number": row["RFP"], "price": float(row["Price"]),
                                    "lead_time": int(row["Lead Time (d)"] or 0),
                                    "moq": int(row["MOQ"] or 1), "notes": row["Notes"] or ""})
                        if not line_quotes:
                            st.warning("Enter at least one price before saving.")
                        else:
                            rfx.record_quotes_batch(vid, vname, line_quotes)
                            st.cache_data.clear()
                            st.success(f"Saved {len(line_quotes)} quote(s) for {vname}.")
                            st.rerun()

    # ── TAB 3 — compare quotes per line, pick a winner (no PO yet) ──────────────
    with tab3:
        # Batch: one call for all quotes, not one call per RFP line. The old
        # per-line version (get_quotes(rfp_number=X) inside this comprehension)
        # meant 33 separate file opens for a BOM-driven RFP batch of that size
        # — each individual call was fast after the internal fix, but 33 file
        # opens still added up to 1.5s+ on every render of this tab.
        all_quotes = rfx.get_quotes()
        quoted_rfp_numbers = {q["rfp_number"] for q in all_quotes}
        quotable = [r for r in rfx.get_open_rfps() if r["rfp_number"] in quoted_rfp_numbers]
        if not quotable:
            st.info("No open lines have quotes yet.")
        else:
            cat_by_mat_w = {}
            for r in quotable:
                if r["mat_code"] not in cat_by_mat_w:
                    item = po_export.get_item_by_code(r["mat_code"], active_only=False)
                    cat_by_mat_w[r["mat_code"]] = (item["category"] if item else "") or "Uncategorized"
            categories_w = sorted({cat_by_mat_w[r["mat_code"]] for r in quotable})
            cat_filter_w = st.multiselect("Filter by category", categories_w,
                default=categories_w, key="rfx_wcat_filter")
            quotable = [r for r in quotable if cat_by_mat_w[r["mat_code"]] in cat_filter_w]

            # ── Bulk path: award many lines at their lowest bid in one action ──
            st.markdown("#### \u26a1 Bulk award — lowest price wins")
            st.caption("For 30+ line RFPs, reviewing one line at a time is slow. This table "
                      "shows every line's cheapest submitted quote — uncheck any line you'd "
                      "rather review individually below, then award the rest in one click.")
            bulk_rows = []
            for r in quotable:
                line_quotes = [q for q in all_quotes if q["rfp_number"] == r["rfp_number"]
                              and q["status"] == "Submitted"]
                if not line_quotes:
                    continue
                best = min(line_quotes, key=lambda q: q["price"])
                bulk_rows.append({"Award": True, "RFP": r["rfp_number"], "Description": r["mat_desc"],
                    "Lowest Vendor": best["vendor_name"], "Price": best["price"],
                    "Lead (d)": best["lead_time"], "Quotes": len(line_quotes),
                    "_quote_id": best["quote_id"]})

            if not bulk_rows:
                st.caption("No undecided quoted lines match this filter.")
            else:
                bulk_df = pd.DataFrame(bulk_rows)
                edited_bulk = st.data_editor(
                    bulk_df.drop(columns=["_quote_id"]), use_container_width=True, hide_index=True,
                    column_config={
                        "Award": st.column_config.CheckboxColumn(width=60),
                        "RFP": st.column_config.TextColumn(disabled=True),
                        "Description": st.column_config.TextColumn(disabled=True, width=220),
                        "Lowest Vendor": st.column_config.TextColumn(disabled=True),
                        "Price": st.column_config.NumberColumn(disabled=True, format="\u20b9%.2f"),
                        "Lead (d)": st.column_config.NumberColumn(disabled=True, width=70),
                        "Quotes": st.column_config.NumberColumn(disabled=True, width=70,
                            help="How many vendors quoted this line"),
                    }, key="rfx_bulk_editor")

                n_checked = int(edited_bulk["Award"].sum())
                if st.button(f"\u2705  Award {n_checked} line(s) at lowest price", type="primary",
                            disabled=(n_checked == 0), key="rfx_bulk_award_btn"):
                    awarded = 0
                    for i, checked in enumerate(edited_bulk["Award"]):
                        if checked:
                            row = bulk_rows[i]
                            rfx.select_winner(row["RFP"], row["_quote_id"])
                            awarded += 1
                    st.cache_data.clear()
                    st.success(f"\u2705 Awarded {awarded} line(s) at lowest price. "
                              "Head to **Issue POs** to consolidate and issue.")
                    st.rerun()

            st.divider()

            # ── Single-line path: full detail + manual override, for exceptions ──
            with st.expander("\U0001f50d Review or override a single line"):
                labels = {r["rfp_number"]: f"{r['rfp_number']} — {r['mat_desc']}" for r in quotable}
                if not labels:
                    st.caption("No lines match the current filter.")
                else:
                    sel = st.selectbox("RFP line", list(labels.keys()), format_func=lambda k: labels[k], key="rfx_wsel")
                    quotes = [q for q in all_quotes if q["rfp_number"] == sel and q["status"] == "Submitted"]
                    if not quotes:
                        st.caption("All quotes for this line already have a decision.")
                    else:
                        qdf = pd.DataFrame(quotes).sort_values("price")
                        show = qdf[["vendor_id","vendor_name","price","lead_time","moq"]].rename(
                            columns={"vendor_id":"Vendor","vendor_name":"Name","price":"Price (\u20b9)",
                                     "lead_time":"Lead (d)","moq":"MOQ"})
                        st.dataframe(show, use_container_width=True, hide_index=True)
                        st.bar_chart(qdf.set_index("vendor_id")["price"], height=160)

                        best = min(quotes, key=lambda q: q["price"])
                        st.info(f"💡 Lowest: **{best['vendor_name']}** at ₹{best['price']:,.2f} "
                                f"({best['lead_time']}d)")

                        st.markdown("###### ❓ Request More Info")
                        st.caption("Before committing to a winner, ask a vendor a follow-up "
                                  "question - a clarification on spec, MOQ, or lead time. "
                                  "Doesn't block selecting a winner; it's here so the question "
                                  "and its answer sit next to the quote they're about.")
                        existing_clr = rfx.get_clarifications(rfp_number=sel)
                        if existing_clr:
                            for c in existing_clr:
                                if c["status"] == "Answered":
                                    st.write(f"**{c['vendor_name']}** — Q: {c['question']}")
                                    st.success(f"✅ A: {c['answer']}")
                                else:
                                    st.write(f"**{c['vendor_name']}** — Q: {c['question']}")
                                    with st.form(key=f"rfx_clr_ans_{c['clarification_id']}"):
                                        ans = st.text_input("Vendor's answer",
                                            key=f"rfx_clr_ans_txt_{c['clarification_id']}")
                                        if st.form_submit_button("Record Answer") and ans:
                                            rfx.record_clarification_response(c["clarification_id"], ans)
                                            st.cache_data.clear()
                                            st.rerun()
                        with st.form(key="rfx_clr_new"):
                            clr_quote_id = st.selectbox("Ask", [q["quote_id"] for q in quotes],
                                format_func=lambda qid: next(q["vendor_name"] for q in quotes if q["quote_id"] == qid),
                                key="rfx_clr_vendor")
                            clr_question = st.text_area("Question", key="rfx_clr_question", height=68)
                            if st.form_submit_button("📨 Send Question") and clr_question.strip():
                                clr_quote = next(q for q in quotes if q["quote_id"] == clr_quote_id)
                                rfx.request_clarification(sel, clr_quote["vendor_id"], clr_quote["vendor_name"],
                                    clr_question.strip(), quote_id=clr_quote_id)
                                st.cache_data.clear()
                                st.success(f"Question logged for {clr_quote['vendor_name']}.")
                                st.rerun()
                        st.divider()

                        opts = {q["quote_id"]: f"{q['vendor_name']} — ₹{q['price']:,.2f} ({q['lead_time']}d)"
                                for q in quotes}
                        pick = st.selectbox("Select winner", list(opts.keys()), format_func=lambda k: opts[k], key="rfx_wpick")
                        if st.button("✅  Mark as Winner (doesn't issue PO yet)", type="primary", key="rfx_wbtn"):
                            rfx.select_winner(sel, pick)
                            st.cache_data.clear()
                            st.success("Marked. Head to the **Issue POs** tab to consolidate and issue.")
                            st.rerun()


    # ── TAB 4 — batch-issue consolidated POs, one per vendor ─────────────────────
    with tab4:
        preview = rfx.preview_pending_pos()
        if not preview:
            st.info("Nothing pending — select winners in the previous tab first.")
        else:
            st.caption("Every vendor below gets **one PO** covering all their winning lines.")
            for g in preview:
                with st.expander(f"\U0001f6d2 {g['vendor_name']} — {g['lines']} line(s) — "
                                  f"\u20b9{g['total_value']:,.2f}", expanded=True):
                    rows = [{"RFP": r["rfp_number"], "Description": r["mat_desc"],
                             "Qty": r["total_qty"], "Price": w["price"], "Lead (d)": w["lead_time"]}
                            for r, w in g["pairs"]]
                    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            also_contract = st.checkbox(
                "\U0001f4dc Also lock these into rate contracts (recommended for repeat items)",
                value=True, key="rfx_also_contract")
            if also_contract:
                st.caption("Skips the separate 'convert PO to contract' step — future PRs for "
                          "these items will auto-match and skip RFx entirely.")
                c1, c2 = st.columns(2)
                with c1:
                    ct_start = st.date_input("Contract Start", value=date.today(), key="rfx_ct_start")
                    ct_terms = st.selectbox("Payment Terms",
                        ["Net 30", "Net 45", "Net 60", "Net 90", "Advance"], key="rfx_ct_terms")
                    ct_autorenew = st.checkbox("Auto-Renew", key="rfx_ct_autorenew")
                with c2:
                    ct_end = st.date_input("Contract End", value=date.today()+timedelta(days=365), key="rfx_ct_end")
                    ct_sla = st.number_input("Delivery SLA (days)", min_value=1, value=7, key="rfx_ct_sla")

            btn_label = (f"\U0001f4dc  Issue {len(preview)} PO(s) & Create Contracts" if also_contract
                        else f"\U0001f3c6  Issue {len(preview)} Consolidated PO(s)")
            if st.button(btn_label, type="primary", use_container_width=True, key="rfx_issue"):
                try:
                    if also_contract:
                        results = ct.award_rfx_to_contracts(ct_start, ct_end, ct_terms,
                            ct_sla, ct_autorenew, "Created from RFx award")
                        for r in results:
                            po_num = r["po_number"]
                            hdr = pr_consolidation.get_po_header(po_num)
                            raw_items = pr_consolidation.get_po_items(po_num)
                            po_lns = [{"mat_code": it["material_code"], "mat_desc": it["material_desc"],
                                      "uom": it["uom"], "qty": it["quantity"],
                                      "deliv_date": it["delivery_date"], "deliv_loc": it["delivery_location"]}
                                     for it in raw_items]
                            fname, fdata = po_export.make_av_bytes(po_num, r["vendor_id"], hdr["po_type"],
                                hdr["legal_entity"], hdr["purchase_entity"], hdr["purchasing_group"],
                                hdr["currency"], hdr["plant_code"], po_lns)
                            base = (f"\u2705 **{po_num}** — {r['vendor_name']} — "
                                   f"{r['lines']} line(s) — \u20b9{r['total_value']:,.2f} — **Created**")
                            if r.get("contract_id"):
                                st.success(f"{base} — locked as **{r['contract_id']}** — {fname} generated.")
                            else:
                                st.warning(f"{base} — contract not created: {r.get('contract_error')} "
                                          f"— {fname} generated.")
                            st.download_button(f"\u2b07 Download {fname}", data=fdata,
                                file_name=fname, mime=XLSX_MIME, key=f"rfx_dl_{po_num}")
                    else:
                        results = rfx.generate_pos()
                        # RFx-awarded POs have no separate "send" screen the way
                        # Consolidate does — generate.pos() already left each one at
                        # status='Created'. Generating and offering the AV file here,
                        # in the same click, matches how Direct PO Entry does it —
                        # neither path has a real gap between "created" and "sent"
                        # to model as its own stage.
                        for r in results:
                            po_num = r["po_number"]
                            hdr = pr_consolidation.get_po_header(po_num)
                            raw_items = pr_consolidation.get_po_items(po_num)
                            po_lns = [{"mat_code": it["material_code"], "mat_desc": it["material_desc"],
                                      "uom": it["uom"], "qty": it["quantity"],
                                      "deliv_date": it["delivery_date"], "deliv_loc": it["delivery_location"]}
                                     for it in raw_items]
                            fname, fdata = po_export.make_av_bytes(po_num, r["vendor_id"], hdr["po_type"],
                                hdr["legal_entity"], hdr["purchase_entity"], hdr["purchasing_group"],
                                hdr["currency"], hdr["plant_code"], po_lns)
                            st.success(f"\u2705 **{po_num}** — {r['vendor_name']} — "
                                      f"{r['lines']} line(s) — \u20b9{r['total_value']:,.2f} — "
                                      f"**Created**, {fname} generated.")
                            st.download_button(f"\u2b07 Download {fname}", data=fdata,
                                file_name=fname, mime=XLSX_MIME, key=f"rfx_dl_{po_num}")
                    st.cache_data.clear()
                except Exception:
                    st.error("\u274c Issuing POs failed:")
                    st.code(traceback.format_exc())


    # ── TAB 5 — read-only history: every RFx already fully awarded ────────────────────
    with tab5:
        st.caption("Once a line's PO is issued it drops out of every tab above — this is "
                  "the one place its own award decision is still browsable: every vendor "
                  "who quoted, and why the winner won, exactly as it was at award time.")
        awarded = rfx.get_rfps_by_status(["Awarded"])
        if not awarded:
            st.info("No RFx has been awarded yet.")
        else:
            cat_by_mat_h = {}
            for r in awarded:
                if r["mat_code"] not in cat_by_mat_h:
                    item = po_export.get_item_by_code(r["mat_code"], active_only=False)
                    cat_by_mat_h[r["mat_code"]] = (item["category"] if item else "") or "Uncategorized"
            categories_h = sorted({cat_by_mat_h[r["mat_code"]] for r in awarded})
            cat_filter_h = st.multiselect("Filter by category", categories_h,
                default=categories_h, key="rfx_hcat_filter")
            awarded = [r for r in awarded if cat_by_mat_h[r["mat_code"]] in cat_filter_h]

            labels_h = {r["rfp_number"]: f"{r['rfp_number']} — {r['mat_desc']}" for r in awarded}
            sel_h = st.selectbox("Awarded RFx", list(labels_h.keys()),
                                 format_func=lambda k: labels_h[k], key="rfx_hsel")

            quotes_h = rfx.get_quotes(rfp_number=sel_h)
            if not quotes_h:
                st.caption("No quote records on file for this line — likely assigned "
                          "directly via **⚡ Skip RFx** rather than competitively quoted.")
            else:
                qdf_h = pd.DataFrame(quotes_h).sort_values("price")
                show_h = qdf_h[["vendor_id", "vendor_name", "price", "lead_time", "moq", "status"]].rename(
                    columns={"vendor_id": "Vendor", "vendor_name": "Name", "price": "Price (₹)",
                             "lead_time": "Lead (d)", "moq": "MOQ", "status": "Decision"})
                st.dataframe(show_h, use_container_width=True, hide_index=True)
                st.bar_chart(qdf_h.set_index("vendor_id")["price"], height=160)

                winner = next((q for q in quotes_h if q["status"] == "Awarded"), None)
                if winner:
                    st.success(f"🏆 Won by **{winner['vendor_name']}** at "
                              f"₹{winner['price']:,.2f} ({winner['lead_time']}d) — "
                              f"{len(quotes_h)} vendor(s) quoted, {len(quotes_h) - 1} not selected.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 6 — Vendor Onboarding
# ══════════════════════════════════════════════════════════════════════════════
def page_vendor_onboarding():
    st.markdown("## \U0001f9fe Vendor Onboarding")
    st.caption("Onboard a vendor directly, or send them a registration questionnaire to "
               "fill in and approve once it's back. Either way, GSTIN/PAN details are "
               "validated before approval.")
    st.divider()

    tab1, tab2, tab3 = st.tabs(["\U0001f4cb Vendors", "\U0001f4dd VRQ Requests",
                                "\U0001f4ca Scorecard"])

    # ── TAB 1 — vendor list + direct intake (unchanged from before) ─────────────
    with tab1:
        s = vo.stats()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Vendors", s["total"])
        m2.metric("Approved", s["approved"])
        m3.metric("Format Verified", s["by_status"].get("Format Verified", 0))
        m4.metric("Needs Review", s["by_status"].get("Needs Review", 0))
        st.divider()

        st.markdown("#### \U0001f4cb Vendors")
        vendors = vo.list_vendors()
        if vendors:
            vdf = pd.DataFrame(vendors)
            show_cols = ["Vendor_ID", "Vendor_Name", "Vendor_Type", "City", "Country", "GSTIN", "PAN",
                         "Active", "Onboarding_Status", "KYC_Flag"]
            show_cols = [c for c in show_cols if c in vdf.columns]
            st.dataframe(vdf[show_cols], use_container_width=True, hide_index=True)

            pending = [v for v in vendors if v.get("Onboarding_Status") == "Format Verified"]
            if pending:
                st.caption(f"\u23f3 {len(pending)} vendor(s) format-verified and awaiting approval:")
                for v in pending:
                    c1, c2 = st.columns([4, 1])
                    c1.markdown(f"**{v['Vendor_ID']}** — {v['Vendor_Name']}")
                    if c2.button("\u2705 Approve", key=f"approve_{v['Vendor_ID']}"):
                        vo.approve_vendor(v["Vendor_ID"])
                        st.cache_data.clear()
                        st.success(f"{v['Vendor_ID']} approved — now eligible for RFx suggestions.")
                        st.rerun()
        else:
            st.info("No vendors yet.")

        st.divider()
        st.markdown("#### \u2795 Onboard a Vendor Directly")

        left, right = st.columns(2, gap="large")
        with left:
            vid = st.text_input("Vendor ID (short code, e.g. ACMEDEN)", key="vo_id")
            vname = st.text_input("Vendor Name", key="vo_name")
            vtypes = vo.list_vendor_types()
            vtype = (st.selectbox("Vendor Type", vtypes, key="vo_type") if vtypes
                    else st.text_input("Vendor Type", key="vo_type_txt",
                        help="No Vendor_Types seeded yet — free text for now."))
            gstin = st.text_input("GSTIN", key="vo_gstin", placeholder="e.g. 27AABCU9603R1ZN")
            if gstin:
                ok, msg, _ = vo.validate_gstin(gstin)
                (st.success if ok else st.error)(f"GSTIN: {msg}")
            pan = st.text_input("PAN", key="vo_pan", placeholder="e.g. AABCU9603R")
            if pan:
                ok, msg = vo.validate_pan_format(pan)
                (st.success if ok else st.error)(f"PAN: {msg}")
            city = st.text_input("City", key="vo_city")
            country = st.text_input("Country", value="India", key="vo_country")
            address = st.text_input("Address", key="vo_addr")
            geo = st.text_input("Geolocation (lat,lng)", key="vo_geo", placeholder="9.9312,76.2673")

        with right:
            contact_name = st.text_input("Contact Name", key="vo_cname")
            contact_email = st.text_input("Contact Email", key="vo_cemail")
            bank_acct = st.text_input("Bank Account No.", key="vo_bank")
            ifsc = st.text_input("IFSC Code", key="vo_ifsc", placeholder="e.g. HDFC0001234")

            if st.button("\U0001f50d Lookup Bank from IFSC", key="vo_ifsc_lookup"):
                res = vo.lookup_ifsc(ifsc)
                if res["ok"]:
                    st.session_state.vo_bank_name = res["bank"]
                    st.session_state.vo_bank_branch = f"{res['branch']}, {res['city']}"
                    st.success(f"\u2705 {res['bank']} — {res['branch']}, {res['city']}")
                else:
                    st.warning(f"{res['error']} — enter bank details manually, or try again "
                               "once this app is running with normal internet access.")

            bank_name = st.text_input("Bank Name", value=st.session_state.get("vo_bank_name", ""), key="vo_bname")
            bank_branch = st.text_input("Bank Branch", value=st.session_state.get("vo_bank_branch", ""), key="vo_bbranch")

        st.markdown("")
        if st.button("\u2705  Validate & Save Vendor", type="primary", use_container_width=True, key="vo_save"):
            if not vid or not vname:
                st.error("Vendor ID and Vendor Name are required.")
            else:
                try:
                    result = vo.upsert_vendor(vid, {
                        "Vendor_Name": vname, "Vendor_Type": vtype, "GSTIN": gstin, "PAN": pan,
                        "City": city, "Country": country, "Address": address, "Geolocation": geo,
                        "Contact_Name": contact_name, "Contact_Email": contact_email,
                        "Bank_Account_No": bank_acct, "IFSC": ifsc,
                        "Bank_Name": bank_name, "Bank_Branch": bank_branch,
                    })
                    st.cache_data.clear()
                    for field, chk in result["checks"].items():
                        (st.success if chk["ok"] else st.error)(f"{field}: {chk['message']}")
                    if result["onboarding_status"] == "Format Verified":
                        st.success(f"\u2705 {vid} saved — format verified, awaiting approval "
                                  "(scroll up to approve).")
                    else:
                        st.warning(f"\u26a0\ufe0f {vid} saved as **Needs Review** — fix the "
                                  "flagged field(s) above and save again.")
                except Exception:
                    st.error("\u274c Error saving vendor:")
                    st.code(traceback.format_exc())

        if vendors:
            st.divider()
            st.markdown("#### \U0001f4ce Document Log")
            st.caption("Logs that a document was received for this vendor.")
            dv = st.selectbox("Vendor", [v["Vendor_ID"] for v in vendors], key="vo_doc_vendor")
            c1, c2, c3 = st.columns([2, 2, 3])
            with c1:
                dtype = st.selectbox("Document Type", ["GST Certificate", "PAN Card",
                    "Drug License", "ISO Certificate", "Cancelled Cheque", "Other"], key="vo_doc_type")
            with c2:
                dfname = st.text_input("Filename / Reference", key="vo_doc_fname")
            with c3:
                dnotes = st.text_input("Notes", key="vo_doc_notes")
            if st.button("\U0001f4ce Log Document", key="vo_doc_save"):
                if not dfname:
                    st.warning("Enter a filename/reference first.")
                else:
                    vo.record_document(dv, dtype, dfname, dnotes)
                    st.success(f"Logged {dtype} for {dv}.")
                    st.rerun()

            docs = vo.get_documents()
            if docs:
                with st.expander(f"\U0001f4c1 {len(docs)} document(s) logged"):
                    st.dataframe(pd.DataFrame(docs), use_container_width=True, hide_index=True)

    # ── TAB 2 — Vendor Registration Questionnaire ────────────────────────────────
    with tab2:
        st.caption("Mirrors how large enterprises actually run vendor intake: send a "
                   "questionnaire → vendor completes it → review → promote.")
        vs = vrq.stats()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total VRQs", vs["total"])
        m2.metric("Sent", vs["by_status"].get("Sent", 0))
        m3.metric("Received", vs["by_status"].get("Received", 0))
        m4.metric("Promoted", vs["by_status"].get("Promoted", 0))
        st.divider()

        st.markdown("#### \u2795 Send a VRQ")
        c1, c2 = st.columns(2)
        with c1:
            new_vname = st.text_input("Vendor Name (as declared by prospect)", key="vrq_new_name")
        with c2:
            new_vemail = st.text_input("Contact Email", key="vrq_new_email")
        if st.button("\U0001f4dd Generate VRQ", type="primary", key="vrq_gen"):
            if not new_vname:
                st.error("Vendor name is required.")
            else:
                vrq_id, fname, fbytes = vrq.send_vrq(new_vname, new_vemail)
                st.session_state.vrq_generated_doc = {"vrq_id": vrq_id, "filename": fname, "bytes": fbytes}
                st.cache_data.clear()

        # Rendered from session_state so the download button survives the rerun
        # that clicking it triggers (see _build_zip's docstring for why).
        gd = st.session_state.get("vrq_generated_doc")
        if gd:
            st.success(f"\u2705 {gd['vrq_id']} created")
            st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                file_name=gd["filename"], mime=XLSX_MIME, key=f"vrq_dl_{gd['vrq_id']}")

        st.divider()
        st.markdown("#### \U0001f4cb VRQ Requests")
        requests = vrq.get_vrq_requests()
        if not requests:
            st.info("No VRQs sent yet.")
        else:
            rdf = pd.DataFrame(requests)[["vrq_id","vendor_name","contact_email","sent_date","status","vendor_id"]]
            rdf.columns = ["VRQ ID","Vendor Name","Email","Sent","Status","Linked Vendor ID"]
            st.dataframe(rdf, use_container_width=True, hide_index=True)

            labels = {r["vrq_id"]: f"{r['vrq_id']} — {r['vendor_name']} ({r['status']})" for r in requests}
            sel = st.selectbox("Work on", list(labels.keys()), format_func=lambda k: labels[k], key="vrq_sel")
            vrq_row = next(r for r in requests if r["vrq_id"] == sel)

            c1, c2 = st.columns(2)
            with c1:
                if st.button("\U0001f3b2 Simulate Vendor Response (Demo)", key="vrq_sim"):
                    vrq.simulate_response(sel)
                    st.cache_data.clear()
                    st.success("Simulated response recorded.")
                    st.rerun()
            with c2:
                uploaded = st.file_uploader("Or upload the completed VRQ", type=["xlsx"], key="vrq_upload")
                if uploaded is not None:
                    upload_marker = f"{sel}:{uploaded.file_id}"
                    if st.session_state.get("vrq_last_processed_upload") != upload_marker:
                        try:
                            answers = vrq.parse_uploaded_vrq(uploaded.read())
                            n = vrq.record_responses(sel, answers)
                            st.session_state.vrq_last_processed_upload = upload_marker
                            st.cache_data.clear()
                            st.success(f"\u2705 Parsed and recorded {n} answer(s).")
                            st.rerun()
                        except ValueError as e:
                            st.session_state.vrq_last_processed_upload = upload_marker
                            st.error(str(e))

            responses = vrq.get_responses(sel)
            if responses:
                st.markdown("##### \U0001f4c4 Responses")
                rows = [{"Section": vrq._ALL_QUESTIONS.get(k,{}).get("section",""),
                         "Question": vrq._ALL_QUESTIONS.get(k,{}).get("text", k),
                         "Answer": v} for k, v in responses.items()]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                checks = vrq.validate_responses(sel)
                issues = {k: c for k, c in checks.items() if not c["ok"]}
                if issues:
                    st.warning(f"\u26a0\ufe0f {len(issues)} issue(s) found:")
                    for k, c in issues.items():
                        st.caption(f"\u2022 **{k}**: {c['message']}")
                else:
                    st.success("\u2705 No validation issues.")

                st.markdown("##### \u2696\ufe0f Decision")
                c1, c2 = st.columns(2)
                with c1:
                    new_vid = st.text_input("Assign Vendor ID", key=f"vrq_assign_{sel}",
                        value=responses.get("legal_name","").split()[0].upper()[:12] if responses.get("legal_name") else "")
                    if st.button("\u2705 Promote to Vendor Master", type="primary", key="vrq_promote"):
                        if not new_vid:
                            st.error("Assign a Vendor ID first.")
                        else:
                            try:
                                result = vrq.promote_to_vendor(sel, new_vid)
                                st.cache_data.clear()
                                st.success(f"\u2705 Promoted as **{new_vid}** — "
                                          f"{result['onboarding_status']}. Approve it in the Vendors tab.")
                                st.rerun()
                            except Exception:
                                st.error("\u274c Promotion failed:")
                                st.code(traceback.format_exc())
                with c2:
                    st.markdown("")
                    st.markdown("")
                    if st.button("\u274c Reject this VRQ", key="vrq_reject"):
                        vrq.reject_vrq(sel)
                        st.cache_data.clear()
                        st.success(f"{sel} marked Rejected.")
                        st.rerun()

    # ── TAB 3 — Vendor Scorecard (VSC-US-01/02) ──────────────────────────────────────────────────────────────
    with tab3:
        st.caption("Blends four components — on-time delivery %, quality reject rate, price "
                   "consistency, and buyer rating — into one weighted score (default "
                   "30/30/20/20, configurable). A vendor with too few closed POs is flagged "
                   "Low Volume rather than hidden. A component with no data yet is left out, "
                   "and its weight is redistributed across the rest.")

        cards = vsc.get_all_scorecards()
        if not cards:
            st.info("No active vendors yet.")
        else:
            cdf = pd.DataFrame(cards)
            show = cdf[["vendor_id", "overall_score", "low_volume", "closed_po_count",
                       "on_time_pct", "quality_reject_pct", "price_consistency_score",
                       "aggregated_rating"]].copy()
            show.columns = ["Vendor", "Score", "Low Vol.", "Closed POs", "On-Time %",
                           "Reject %", "Price Consistency", "Avg Rating"]
            show = show.sort_values("Score", ascending=False, na_position="last")
            st.dataframe(show, use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("##### 🔍 Drill into one vendor")
            vendor_names = {v["Vendor_ID"]: v["Vendor_Name"] for v in vo.list_vendors(include_inactive=False)}
            sel_v = st.selectbox("Vendor", list(vendor_names.keys()),
                format_func=lambda k: f"{k} — {vendor_names[k]}", key="vsc_drill_sel")
            card = vsc.get_vendor_scorecard(sel_v)

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Overall", card["overall_score"] if card["overall_score"] is not None else "—")
            m2.metric("On-Time %", card["on_time_pct"] if card["on_time_pct"] is not None else "—",
                      help=f"{card['on_time_sample']} real line(s) measured")
            m3.metric("Reject %", card["quality_reject_pct"] if card["quality_reject_pct"] is not None else "—",
                      help=f"{card['quality_sample']} real inspected line(s)")
            m4.metric("Price Consistency", card["price_consistency_score"] if card["price_consistency_score"] is not None else "—",
                      help=f"{card['price_sample']} material(s) with real repeat pricing")
            m5.metric("Avg Rating", f"{card['aggregated_rating']:g}/5" if card["aggregated_rating"] is not None else "—",
                      help=f"{card['rating_sample']} real rating(s)")
            if card["low_volume"]:
                st.warning(f"⚠️ Low Volume — only {card['closed_po_count']} real Closed PO(s) "
                          f"on file. This score is real, but thin.")
            if card["rtv_count"]:
                st.caption(f"📦 {card['rtv_count']} confirmed Return-to-Vendor shipment(s) on "
                          f"file — supporting context, not double-counted into the reject rate above.")

            ratings = vsc.get_ratings(vendor_id=sel_v)
            if ratings:
                st.markdown("##### 💬 Ratings & Reviews")
                for r in ratings:
                    stars = "⭐" * r["rating"]
                    st.write(f"**{stars}** ({r['po_number']}, by {r['rated_by']}, {r['rating_date']}) "
                            f"— {r['review'] or '*no review text*'}")
                    if r["vendor_response"]:
                        st.caption(f"↳ Vendor response ({r['vendor_response_date']}): {r['vendor_response']}")
                    else:
                        with st.form(key=f"vsc_resp_{r['rating_id']}"):
                            resp_text = st.text_input("Vendor's response", key=f"vsc_resp_txt_{r['rating_id']}")
                            if st.form_submit_button("Record Response") and resp_text:
                                vsc.respond_to_rating(r["rating_id"], resp_text)
                                st.cache_data.clear()
                                st.rerun()

            st.divider()
            st.markdown("##### ➕ Rate a vendor")
            st.caption("Gated on a real Closed PO — every line Fully Received. One rating "
                      "per PO per rater, immutable once submitted.")
            rater = st.text_input("Your name", key="vsc_rater")
            eligible = vsc.get_unrated_closed_pos(sel_v, rater) if rater else []
            if not rater:
                st.caption("Enter your name to see which Closed POs you haven't rated yet.")
            elif not eligible:
                st.caption("No unrated Closed POs for this vendor right now.")
            else:
                po_labels = {h["po_number"]: f"{h['po_number']} ({h['po_date']})" for h in eligible}
                sel_po = st.selectbox("Closed PO", list(po_labels.keys()),
                    format_func=lambda k: po_labels[k], key="vsc_rate_po")
                rating_val = st.slider("Rating", 1, 5, 4, key="vsc_rate_val")
                review_text = st.text_area("Review", key="vsc_rate_review", height=68)
                if st.button("✅ Submit Rating", type="primary", key="vsc_rate_submit"):
                    try:
                        vsc.rate_vendor(sel_po, sel_v, rater, rating_val, review_text)
                        st.cache_data.clear()
                        st.success(f"✅ Rated {sel_v} {rating_val}/5 for {sel_po}.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ {e}")



# ══════════════════════════════════════════════════════════════════════════════
# PAGE 7 — Contracts
# ══════════════════════════════════════════════════════════════════════════════
def page_purchase_bundles():
    st.markdown("## \U0001f4e6 Purchase Bundles")
    st.caption("Named, reusable sets of materials with default quantities — a requester "
               "picks a bundle instead of adding each item individually. Bundles are "
               "available from the Purchase Bundles tab on the Create PR page.")
    st.divider()

    s = pbdl.stats()
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Bundles", s["total_bundles"])
    m2.metric("Active", s["active_bundles"])
    m3.metric("Total Line Items", s["total_lines"])
    st.divider()

    tab_list, tab_new, tab_discover = st.tabs(["\U0001f4cb Bundles", "\u2795 Create Bundle",
                                               "\U0001f50d Discover Candidates"])

    # ── TAB 1 — list, view, edit, activate/deactivate ────────────────────────────
    with tab_list:
        bundles = pbdl.list_bundles(active_only=False)
        if not bundles:
            st.info("No purchase bundles yet — create one in the next tab.")
        else:
            bdf = pd.DataFrame(bundles)[["bundle_id", "bundle_name", "department",
                                         "created_by", "created_date", "active"]]
            bdf.columns = ["Bundle ID", "Name", "Department", "Created By", "Created", "Active"]
            st.dataframe(bdf, use_container_width=True, hide_index=True)

            labels = {b["bundle_id"]: f"{b['bundle_name']}" +
                     (f"  ({b['department']})" if b["department"] else "") +
                     ("  \u2014 inactive" if b["active"] != "Yes" else "")
                     for b in bundles}
            sel = st.selectbox("View / edit bundle", list(labels.keys()),
                               format_func=lambda k: labels[k], key="bdl_view_sel")
            bundle = next(b for b in bundles if b["bundle_id"] == sel)
            items = pbdl.get_bundle_items(sel)

            left, right = st.columns([3, 2], gap="large")
            with left:
                st.markdown("##### \U0001f4c4 Line Items")
                if items:
                    idf = pd.DataFrame(items)[["mat_code", "mat_desc", "uom", "default_qty", "notes"]]
                    idf.columns = ["Code", "Description", "UOM", "Default Qty", "Notes"]
                    st.dataframe(idf, use_container_width=True, hide_index=True)
                else:
                    st.caption("No line items.")

                with st.expander("\u2795 Add / update a line item"):
                    c1, c2, c3 = st.columns([2, 1, 2])
                    with c1:
                        add_code = st.text_input("Material Code", key=f"bdl_addcode_{sel}")
                    with c2:
                        add_qty = st.number_input("Default Qty", min_value=0.0, value=1.0,
                                                  step=1.0, key=f"bdl_addqty_{sel}")
                    with c3:
                        add_notes = st.text_input("Notes (optional)", key=f"bdl_addnotes_{sel}")
                    if st.button("Save line item", key=f"bdl_addbtn_{sel}"):
                        try:
                            pbdl.add_or_update_bundle_item(sel, add_code.strip(), add_qty, add_notes)
                            st.success(f"Saved {add_code.strip()}.")
                            st.rerun()
                        except ValueError as e:
                            st.error(str(e))

                if items:
                    with st.expander("\U0001f5d1 Remove a line item"):
                        rm_code = st.selectbox("Item to remove", [i["mat_code"] for i in items],
                                               key=f"bdl_rmcode_{sel}")
                        if st.button("Remove", key=f"bdl_rmbtn_{sel}"):
                            pbdl.remove_bundle_item(sel, rm_code)
                            st.success(f"Removed {rm_code}.")
                            st.rerun()

            with right:
                st.markdown("##### \u2696\ufe0f Actions")
                if bundle["description"]:
                    st.caption(bundle["description"])
                if bundle["active"] == "Yes":
                    if st.button("\U0001f6ab Deactivate", key=f"bdl_deact_{sel}",
                                 help="Removes it from the Create PR picker. Doesn't affect "
                                      "PRs already created from it, and can be reversed."):
                        pbdl.set_bundle_active(sel, False)
                        st.success(f"{sel} deactivated.")
                        st.rerun()
                else:
                    if st.button("\u2705 Reactivate", key=f"bdl_react_{sel}"):
                        pbdl.set_bundle_active(sel, True)
                        st.success(f"{sel} reactivated.")
                        st.rerun()

    # ── TAB 2 — create a new bundle ──────────────────────────────────────────────
    with tab_new:
        if "bdl_new_staged" not in st.session_state:
            st.session_state.bdl_new_staged = []

        c1, c2 = st.columns(2)
        with c1:
            new_name = st.text_input("Bundle Name", key="bdl_new_name",
                                     placeholder="e.g. New Site HSE Kit")
            new_dept = st.text_input("Department (optional)", key="bdl_new_dept")
        with c2:
            new_desc = st.text_area("Description (optional)", key="bdl_new_desc", height=68)
            new_by = st.text_input("Created By", value="REQ-104", key="bdl_new_by")

        st.markdown("##### \U0001f50d Add Items")
        s1, s2 = st.columns([4, 1])
        with s1:
            q = st.text_input("Search catalog", key="bdl_new_q", label_visibility="collapsed",
                              placeholder="Type then click Search")
        with s2:
            do_search = st.button("\U0001f50d Search", key="bdl_new_sbtn", use_container_width=True)

        if do_search and len(q) >= 2:
            st.session_state.bdl_new_results = po_export.fuzzy_search(q, load_catalog(), max_results=15)

        results = st.session_state.get("bdl_new_results", [])
        staged_codes = {i["mat_code"] for i in st.session_state.bdl_new_staged}
        for item in results:
            c1, c2, c3 = st.columns([5, 1.2, 1])
            with c1:
                pfx = "\u2713 " if item["code"] in staged_codes else ""
                st.caption(f"{pfx}**{item['desc']}** — {item['code']} \u00b7 {item['uom']}")
            with c2:
                qty = st.number_input("Qty", min_value=0.0, value=1.0, step=1.0,
                                      key=f"bdl_new_qty_{item['code']}", label_visibility="collapsed")
            with c3:
                if item["code"] not in staged_codes:
                    if st.button("Add", key=f"bdl_new_add_{item['code']}"):
                        st.session_state.bdl_new_staged.append({
                            "mat_code": item["code"], "desc": item["desc"],
                            "uom": item["uom"], "qty": qty})
                        st.rerun()

        if st.session_state.bdl_new_staged:
            st.markdown("##### \U0001f4cb Staged Items")
            sdf = pd.DataFrame(st.session_state.bdl_new_staged)[["mat_code", "desc", "uom", "qty"]]
            sdf.columns = ["Code", "Description", "UOM", "Qty"]
            st.dataframe(sdf, use_container_width=True, hide_index=True)
            to_rm = st.multiselect("Remove staged item(s)",
                                   [i["mat_code"] for i in st.session_state.bdl_new_staged],
                                   key="bdl_new_rm")
            if to_rm and st.button("Remove selected", key="bdl_new_rm_btn"):
                st.session_state.bdl_new_staged = [
                    i for i in st.session_state.bdl_new_staged if i["mat_code"] not in to_rm]
                st.rerun()

        st.divider()
        if st.button("\u2705 Create Bundle", type="primary", use_container_width=True,
                     disabled=not (new_name.strip() and st.session_state.bdl_new_staged),
                     key="bdl_new_create"):
            try:
                bundle_id = pbdl.create_bundle(
                    new_name, description=new_desc, department=new_dept, created_by=new_by,
                    items=[{"mat_code": i["mat_code"], "qty": i["qty"]}
                           for i in st.session_state.bdl_new_staged])
                st.success(f"\u2705 **{bundle_id}** created with "
                          f"{len(st.session_state.bdl_new_staged)} line(s).")
                st.session_state.bdl_new_staged = []
                st.session_state.bdl_new_results = []
                st.rerun()
            except ValueError as e:
                st.error(str(e))

    # ── TAB 3 — agent-assisted discovery from real PO history ───────────────────
    with tab_discover:
        st.caption("Finds materials that repeatedly appear together on the same purchase order "
                   "and proposes candidate bundles. Nothing is created automatically — each "
                   "candidate is a proposal you can review and turn into a real bundle.")
        min_count = st.slider("Minimum repeat count to count as a pattern", 2, 5, 2,
                              key="bdl_discover_min")
        candidates = pbdl.discover_bundle_candidates(min_po_count=min_count)

        if not candidates:
            st.success("\u2705 No new patterns found above this threshold — either "
                      "nothing repeats that often, or it's already captured in an "
                      "existing bundle.")
        else:
            for i, c in enumerate(candidates):
                with st.container(border=True):
                    st.markdown(f"**{c['mat_descs'][0]}** + **{c['mat_descs'][1]}**")
                    st.caption(f"Ordered together on {c['po_count']} separate POs: " +
                              ", ".join(c["po_refs"]))
                    c1, c2 = st.columns([2, 1])
                    default_name = f"{c['mat_descs'][0].split(chr(32))[0]} + " \
                                   f"{c['mat_descs'][1].split(chr(32))[0]} Kit"
                    bundle_name = c1.text_input("Bundle name", value=default_name,
                                                key=f"bdl_discover_name_{i}")
                    if c2.button("\u2795 Create Bundle From This", key=f"bdl_discover_create_{i}"):
                        try:
                            bundle_id = pbdl.create_bundle(
                                bundle_name, description=f"Proposed from PO co-occurrence "
                                f"({c['po_count']} POs: {', '.join(c['po_refs'])})",
                                department="", created_by="AGENT-DISCOVERY",
                                items=[{"mat_code": code, "qty": 1} for code in c["mat_codes"]])
                            st.success(f"\u2705 **{bundle_id}** created with "
                                      f"{len(c['mat_codes'])} line(s) \u2014 edit "
                                      f"quantities or add more items in the Bundles tab.")
                            st.rerun()
                        except ValueError as e:
                            st.error(str(e))


def page_contracts():
    st.markdown("## \U0001f4dc Contract Management")
    st.caption("Convert an awarded purchase order into a rate contract with locked pricing "
               "for a validity window. Future requisitions for a contracted item go "
               "straight to a priced PO with the same vendor, no quoting needed.")
    st.divider()

    s = ct.stats()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Contracts", s["total"])
    m2.metric("Active", s["by_status"].get("Active", 0))
    m3.metric("Expired", s["by_status"].get("Expired", 0))
    m4.metric("Renewals Due (30d)", s["renewals_due_30d"])

    due = ct.renewals_due(30)
    if due:
        st.warning(f"\u23f0 {len(due)} contract(s) expiring within 30 days: " +
                  ", ".join(f"**{d['contract_id']}** ({d['vendor_name']}, ends {d['end_date']})" for d in due))
    st.divider()

    tab1, tab2 = st.tabs(["\U0001f4cb Contracts", "\u2795 Convert PO to Contract"])

    # ── TAB 1 — list, detail, lifecycle actions ──────────────────────────────────
    with tab1:
        contracts = ct.get_contracts()
        if not contracts:
            st.info("No contracts yet — convert an awarded PO in the next tab.")
        else:
            cdf = pd.DataFrame(contracts)[["contract_id","vendor_id","vendor_name","status",
                "start_date","end_date","payment_terms","delivery_sla_days","auto_renew"]]
            cdf.columns = ["Contract","Vendor ID","Vendor","Status","Start","End",
                          "Payment Terms","SLA (d)","Auto-Renew"]
            st.dataframe(cdf, use_container_width=True, hide_index=True)

            labels = {c["contract_id"]: f"{c['contract_id']} — {c['vendor_name']} ({c['status']})"
                      for c in contracts}
            sel = st.selectbox("View contract", list(labels.keys()), format_func=lambda k: labels[k], key="ct_sel")
            contract = next(c for c in contracts if c["contract_id"] == sel)
            items = ct.get_contract_items(sel)

            left, right = st.columns([3,2], gap="large")
            with left:
                st.markdown("##### \U0001f4c4 Contracted Line Items")
                idf = pd.DataFrame(items)[["mat_code","mat_desc","uom","unit_price",
                    "min_order_qty","lead_time_days"]]
                idf.columns = ["Code","Description","UOM","Unit Price","MOQ","Lead (d)"]
                st.dataframe(idf, use_container_width=True, hide_index=True)
                total_potential = sum(i["unit_price"] for i in items)
                st.caption(f"Source PO: {contract['source_po']}  \u00b7  "
                          f"{len(items)} line(s)  \u00b7  Sum of unit prices: \u20b9{total_potential:,.2f}")

            with right:
                st.markdown("##### \u2696\ufe0f Actions")
                if st.button("\U0001f4c4 Generate Contract Document", key="ct_gen_doc"):
                    fname, fbytes = ct.generate_contract_document(sel)
                    st.session_state.ct_generated_doc = {"contract_id": sel, "filename": fname, "bytes": fbytes}
                gd = st.session_state.get("ct_generated_doc")
                if gd and gd["contract_id"] == sel:
                    st.download_button(f"\u2b07 Download {gd['filename']}", data=gd["bytes"],
                        file_name=gd["filename"], mime=XLSX_MIME, key=f"ct_dl_{sel}")

                if contract["status"] in ("Active", "Scheduled"):
                    with st.popover("\U0001f501 Renew"):
                        new_end = st.date_input("New End Date",
                            value=date.today()+timedelta(days=365), key=f"ct_renew_date_{sel}")
                        if st.button("Confirm Renewal", key=f"ct_renew_btn_{sel}"):
                            ct.renew_contract(sel, new_end)
                            st.cache_data.clear()
                            st.success(f"{sel} renewed to {new_end}.")
                            st.rerun()
                    with st.popover("\u274c Terminate"):
                        reason = st.text_input("Reason", key=f"ct_term_reason_{sel}")
                        if st.button("Confirm Termination", key=f"ct_term_btn_{sel}"):
                            ct.terminate_contract(sel, reason)
                            st.cache_data.clear()
                            st.success(f"{sel} terminated.")
                            st.rerun()
                else:
                    st.caption(f"No lifecycle actions available for a **{contract['status']}** contract.")

    # ── TAB 2 — convert an awarded PO into a contract ───────────────────────────
    with tab2:
        candidates = ct.list_pos_available_for_contract()
        if not candidates:
            st.info("No priced POs available to convert yet. POs need at least one line "
                    "with a known Unit_Price — award something via RFx, or auto-match an "
                    "existing contract, first.")
        else:
            labels = {c["po_number"]: f"{c['po_number']} — {c['vendor_name']} ({c['lines']} line(s))"
                      for c in candidates}
            sel_po = st.selectbox("Source PO", list(labels.keys()), format_func=lambda k: labels[k], key="ct_new_po")

            c1, c2 = st.columns(2)
            with c1:
                start = st.date_input("Start Date", value=date.today(), key="ct_new_start")
                payment_terms = st.selectbox("Payment Terms",
                    ["Net 30", "Net 45", "Net 60", "Net 90", "Advance"], key="ct_new_terms")
                auto_renew = st.checkbox("Auto-Renew", key="ct_new_autorenew")
            with c2:
                end = st.date_input("End Date", value=date.today()+timedelta(days=365), key="ct_new_end")
                delivery_sla = st.number_input("Delivery SLA (days)", min_value=1, value=7, key="ct_new_sla")
            notes = st.text_input("Notes", key="ct_new_notes")

            if st.button("\U0001f4dc  Create Contract", type="primary", key="ct_create_btn"):
                if end <= start:
                    st.error("End date must be after start date.")
                else:
                    try:
                        result = ct.create_contract_from_po(sel_po, start, end, payment_terms,
                            delivery_sla, auto_renew, notes)
                        st.cache_data.clear()
                        msg = f"\u2705 {result['contract_id']} created with {result['lines']} line(s)."
                        if result["skipped_unpriced"]:
                            msg += f" ({result['skipped_unpriced']} unpriced line(s) skipped.)"
                        st.success(msg)
                        st.rerun()
                    except Exception:
                        st.error("\u274c Contract creation failed:")
                        st.code(traceback.format_exc())


# ── Router ─────────────────────────────────────────────────────────────────────
if   "Purchase Bundles" in page: page_purchase_bundles()
elif "Create PR"   in page: page_create_pr()
elif "Consolidate" in page: page_consolidate()
elif "Direct PO"   in page: page_direct_po()
elif "Auto-Categorization" in page: page_categorization()
elif "RFx" in page: page_rfx()
elif "Vendor Onboarding" in page: page_vendor_onboarding()
elif "Contracts" in page: page_contracts()
