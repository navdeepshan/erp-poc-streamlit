"""
item_tax.py — HSN code + GST rate on Item Master (O2C, prerequisite for Billing).

Deliberately does NOT auto-classify HSN codes or guess GST rates. HSN
classification is real tax work — it depends on the specific product,
and getting it wrong has real compliance consequences. This module gives
items a place to record that classification and a way to see which ones
are missing it (mirroring categorization.py's honest "needs review"
list), but a human — ideally with a CA's input — has to actually supply
the values. billing.py refuses to invoice a line with no GST rate set,
same enforcement principle as org_profile.py's is_configured() gate.

Item Master now lives in SQLite (item_master table, HSN_Code/GST_Rate
columns already part of its schema). Reads here are deliberately
active_only=False (via po_export.load_item_master) — an item shouldn't
lose its tax classification, or become invisible to the "needs review"
list, just because it's currently marked inactive; Active gates
selectability for new transactions, not whether master data is allowed
to be complete.
"""

import os

import db
import po_export

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

GST_RATE_CHOICES = [0, 5, 12, 18, 28]  # standard Indian GST slabs


def get_item_tax_info(material_code, data_file=None):
    item = po_export.get_item_by_code(material_code, active_only=False)
    if not item:
        return {"hsn_code": None, "gst_rate": None, "mat_desc": None}
    return {"hsn_code": item["hsn_code"], "gst_rate": item["gst_rate"], "mat_desc": item["desc"]}


def set_item_tax_info(material_code, hsn_code, gst_rate, data_file=None):
    conn = db.get_connection()
    try:
        cur = conn.execute(
            "UPDATE item_master SET hsn_code = ?, gst_rate = ? WHERE item_code = ?",
            (hsn_code, gst_rate, material_code),
        )
        found = cur.rowcount > 0
        conn.commit()
    finally:
        conn.close()
    if not found:
        raise ValueError(f"Material code {material_code} not found in Item Master.")
    return True


def list_items_missing_tax_info(data_file=None):
    """Mirrors categorization.py's 'needs review' pattern — items with no
    HSN code or no GST rate set, so they're visible rather than silently
    invoiced wrong."""
    items = po_export.load_item_master(data_file, active_only=False)
    out = []
    for item in items:
        if not item["hsn_code"] or item["gst_rate"] is None:
            out.append({"mat_code": item["code"], "mat_desc": item["desc"],
                        "hsn_code": item["hsn_code"], "gst_rate": item["gst_rate"]})
    return out


def stats(data_file=None):
    items = po_export.load_item_master(data_file, active_only=False)
    total = len(items)
    configured = sum(1 for item in items if item["hsn_code"] and item["gst_rate"] is not None)
    return {"total": total, "configured": configured, "missing": total - configured}


if __name__ == "__main__":
    print("Item tax stats:", stats())
    missing = list_items_missing_tax_info()
    print(f"\n{len(missing)} item(s) missing HSN/GST info (showing first 5):")
    for m in missing[:5]:
        print(" ", m)
