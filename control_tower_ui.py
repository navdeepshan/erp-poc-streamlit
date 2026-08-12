"""
control_tower_ui.py — Exception Control Tower, as its own Streamlit app.

Same reasoning as every other split in this project (see mfg_ui.py's own
docstring): this is a genuinely distinct concern — a read-only worklist
that reads across S2C/MFG/O2C — not a natural extension of any one of
those three apps, and folding it into one of them would be exactly the
cross-app coupling this project's own CLAUDE.md forbids. This app never
writes anything; every fix still happens on its owning screen in one of
the other three apps. Reuses control_tower.py's registry, no UI code
shared with any other app.

v1 is text-only navigation instructions (owner_app / owner_screen shown
as plain text) rather than clickable cross-app deep links — a real
"jump to X" needs each of the other three apps to read a launch query
param, a small change to already-working code, correctly sequenced
*after* this registry pattern is proven live, not before.

Run with: streamlit run control_tower_ui.py
"""

import streamlit as st
import pandas as pd

import control_tower as ct
from professional_theme import apply_theme

st.set_page_config(page_title="Exception Control Tower", page_icon="\U0001f6e2", layout="wide")
apply_theme("Exception Control Tower")

st.title("\U0001f6e2 Exception Control Tower")
st.caption(
    "One worklist across the whole platform — Source to Contract, Manufacturing/"
    "Inventory, and Order to Cash — every row here is a real, currently-open "
    "exception read live from its own owning screen. Fixing anything still happens "
    "over there; this page only tells you it needs attention and where to go."
)

all_ex, failed_sources = ct.get_all_exceptions()

if failed_sources:
    st.warning(
        "Some exception sources could not be read this refresh, so the list below is "
        "incomplete:\n" + "\n".join(f"- `{name}`: {err}" for name, err in failed_sources)
    )

stats = ct.stats()
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Open Exceptions", stats["total"])
c2.metric("\U0001f534 Critical", stats["by_severity"].get("Critical", 0))
c3.metric("\U0001f7e0 Attention", stats["by_severity"].get("Attention", 0))
c4.metric("\U0001f7e1 Info", stats["by_severity"].get("Info", 0))

st.divider()

f1, f2, f3, f4 = st.columns(4)
categories = ct.get_categories()
sel_categories = f1.multiselect("Category", categories)
sel_severities = f2.multiselect("Severity", ["Critical", "Attention", "Info"])
sel_owners = f3.multiselect("Owner App", ["S2C", "MFG", "O2C"])
if f4.button("\U0001f504 Refresh", use_container_width=True):
    st.rerun()

filtered = all_ex
if sel_categories:
    filtered = [e for e in filtered if e["category"] in sel_categories]
if sel_severities:
    filtered = [e for e in filtered if e["severity"] in sel_severities]
if sel_owners:
    filtered = [e for e in filtered if e["owner_app"] in sel_owners]

st.subheader(f"Worklist ({len(filtered)} of {len(all_ex)})")

if not filtered:
    st.success("Nothing matches the current filters. If unfiltered, the pilot is clean.")
else:
    df = pd.DataFrame([
        {
            "Severity": e["severity"],
            "Category": e["category"],
            "ID": e["id"],
            "Age (days)": e["age_days"],
            "Details": e["title"],
            "Go to": e["owner_app"],
            "Screen": e["owner_screen"],
        }
        for e in filtered
    ])
    severity_icon = {"Critical": "\U0001f534", "Attention": "\U0001f7e0", "Info": "\U0001f7e1"}
    df["Severity"] = df["Severity"].map(lambda s: f"{severity_icon.get(s, '')} {s}")
    st.dataframe(df, use_container_width=True, hide_index=True, height=min(700, 60 + 35 * len(df)))

st.caption(
    f"{stats['sources_registered']} exception sources registered "
    f"({stats['sources_registered'] - stats['sources_failed']} read successfully this refresh). "
    "Nothing here is cached — every refresh recomputes from the same live data the owning "
    "app itself would show."
)
