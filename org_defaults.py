"""
org_defaults.py — Default field values for PR/PO creation forms.

Replaces 12 hardcoded value="..." strings previously scattered across
erp_ui.py's Create PR and Direct PO Entry pages (PO Type, Legal Entity,
Purchasing Entity, Purchasing Group, Currency, Plant — each duplicated
across both pages). Those defaults now live here as real, editable
master data instead of literals baked into the UI code — changing one
for a different industry profile no longer means hunting through two
pages' worth of text_input calls.

Deliberately a flat Org_Element -> Default_Value lookup, not a richer
structured table — every one of these fields is a single string the
UI needs to pre-fill, nothing more. If a field ever needs validation
or a foreign-key relationship (e.g. Legal Entity should really
reference legal_entities.le_id), that's a natural extension later;
not built now because nothing enforces it downstream yet either.

New table:
  Org_Defaults   Org_Element | Default_Value
"""

import os

import db

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

# Fallback values if a key is missing from the table entirely (fresh
# install before any seed data has been loaded) — matches what used to
# be hardcoded directly in erp_ui.py, so behavior doesn't change for
# anyone who hasn't touched this yet.
_FALLBACKS = {
    "PO Type": "NB", "Legal Entity": "LE-001", "Purchasing Entity": "PE-001",
    "Purchasing Group": "PG-001", "Currency": "INR", "Plant": "PLANT-01",
    "Demand Detection Mode": "Manufactured Items Only",
    "Default Transfer Lead Time Days": "3",
    "Time-Phased Planning Mode": "Sales Order Based",
    "Use Stock Transfer Orders": "No",
    "Concurrency Guarantee Level": "Strict",
    "ATP Sourcing Scope": "Single-Plant",
    "Reservation/Backorder Priority": "First-Confirmed-First-Served",
    "Reservation Granularity": "Quantity-Only",
    "Reservation Visibility to Planning": "Reserved Replaces Confirmed-SO Signal",
    "Vendor Scorecard Weight - On-Time": "30",
    "Vendor Scorecard Weight - Quality": "30",
    "Vendor Scorecard Weight - Price Consistency": "20",
    "Vendor Scorecard Weight - Rating": "20",
    "Vendor Scorecard Min Transactions": "3",
}

# Values other than each element's own real, currently-built default
# are shown in Settings (so the intended design surface is visible —
# these are real, named policies from ATP-US-01's own Default
# Configuration Values table, not placeholders) but locked from
# selection there: each one names a genuine, larger piece of future
# engineering (Network ATP sourcing consulting the transfer network,
# Customer Priority Tier requiring a real tier field on Customer
# Master, lot-level reservation binding, an alternate planning-signal
# mode) that was deliberately not built alongside ATP-US-01/02/03,
# not an oversight. get_default() and set_org_default() below don't
# enforce this themselves -- the lock lives in the Settings UI, so a
# direct call (e.g. from a script) can still set an unsupported value.
# validate_atp_policy() exists so any future caller that reads one of
# these can check explicitly, and fail with a clear, real reason,
# rather than a decision point silently behaving as if the
# unsupported mode were honored.
ATP_POLICY_OPTIONS = {
    "Concurrency Guarantee Level": ["Strict", "Relaxed"],
    "ATP Sourcing Scope": ["Single-Plant", "Network"],
    "Reservation/Backorder Priority": ["First-Confirmed-First-Served", "Customer Priority Tier"],
    "Reservation Granularity": ["Quantity-Only", "Lot-Level"],
    "Reservation Visibility to Planning": ["Reserved Replaces Confirmed-SO Signal",
                                           "Confirmed-SO Signal Only (pre-ATP)"],
}


def ensure_sheets(wb=None):
    db.init_schema()


def get_org_defaults(data_file=None):
    """{Org_Element: Default_Value} for every row currently set."""
    conn = db.get_connection()
    try:
        rows = conn.execute("SELECT * FROM org_defaults").fetchall()
    finally:
        conn.close()
    return {r["org_element"]: r["default_value"] for r in rows}


def get_default(element, data_file=None):
    """Single-value convenience getter for the UI — always returns
    something usable, falling back to the pre-this-feature hardcoded
    value if the table has nothing set for this element yet."""
    return get_org_defaults(data_file).get(element, _FALLBACKS.get(element, ""))


def set_org_default(element, value, data_file=None):
    db.init_schema()
    conn = db.get_connection()
    try:
        conn.execute(
            "INSERT INTO org_defaults (org_element, default_value) VALUES (?, ?) "
            "ON CONFLICT(org_element) DO UPDATE SET default_value = excluded.default_value",
            (element, value),
        )
        conn.commit()
    finally:
        conn.close()


def set_org_defaults(fields, data_file=None):
    """fields: {Org_Element: Default_Value} — sets several at once."""
    for element, value in fields.items():
        set_org_default(element, value, data_file)


def validate_atp_policy(element, data_file=None):
    """
    For a real ATP policy element (one of the five keys in
    ATP_POLICY_OPTIONS): returns the current value if it's the one
    real, currently-built default; raises a clear NotImplementedError
    otherwise. The Settings UI already locks these to their own real
    default, so this should never actually trigger through normal use
    — it exists for a direct set_org_default() call (a script, a future
    caller) that bypasses that lock, so a decision point in atp.py/
    reservation.py/backorder.py/bom.py fails with a real, specific
    reason instead of silently behaving as if an unbuilt mode were
    honored.
    """
    value = get_default(element, data_file)
    built_default = _FALLBACKS.get(element)
    if element in ATP_POLICY_OPTIONS and value != built_default:
        raise NotImplementedError(
            f"'{element}' is set to '{value}', but only '{built_default}' is "
            f"actually built in this release — the alternative is a real, "
            f"named future extension (see ATP_POLICY_OPTIONS), not yet "
            f"implemented at any decision point. Locked in Settings for "
            f"exactly this reason; this value must have been set directly.")
    return value


if __name__ == "__main__":
    print("Org defaults:", get_org_defaults())
