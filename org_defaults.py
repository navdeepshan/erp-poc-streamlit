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


if __name__ == "__main__":
    print("Org defaults:", get_org_defaults())
