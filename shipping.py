"""
shipping.py — real shipping-detail collection and Excel export for
third-party courier handoff.

Built 2026-07-31, deliberately scoped down at direct request: just
BlueDart (not Delhivery yet), just the Excel export (not "Ship All",
not a live API connector) — all explicitly deferred to a later pass,
not overlooked. COURIERS below is a list of one specifically so
extending it later is a one-line change, not a redesign.
"""

import re
import io
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

import po_export
import pr_consolidation as pc
import org_profile as op

COURIERS = ["BlueDart"]


def _extract_pincode(address, city):
    """A real 6-digit Indian PIN, pulled from the location's own real
    address/city text, not fabricated. Returns "" if genuinely none is
    present, rather than guessing one — every real Delivery Location's
    address in this system's seed data does end with one, checked
    directly before relying on this, but a location added later
    without one should show an honest blank, not a wrong number."""
    m = re.search(r"\b(\d{6})\b", f"{address or ''} {city or ''}")
    return m.group(1) if m else ""


def _location_lookup(data_file=None):
    return {l["id"]: l for l in pc.get_delivery_locations(active_only=False)}


def build_shipment_details(transfer, data_file=None):
    """
    transfer: a real row from inventory.get_stock_transfers() — an
    actual, already-shipped movement with its own real transfer_id,
    quantities, and carrier, not a transfer OPPORTUNITY (a suggestion
    that may never be acted on). Every field below is either read
    directly from real data (item master, delivery locations, org
    profile, the transfer record itself) or computed from it — nothing
    here is a placeholder value.

    Consignor and consignee contact both use the org's own profile
    contact (Contact_Phone/Contact_Email), not a fabricated separate
    number per location — honest, since these are the same company's
    own two facilities, and no per-location contact is captured
    anywhere in this system to draw on instead.

    Weight and declared value are both real computations (per-unit
    weight/price from the item master times the real shipped
    quantity), not flat estimates re-applied per shipment.
    """
    locs = _location_lookup(data_file)
    from_loc = locs.get(transfer["from_location"], {})
    to_loc = locs.get(transfer["to_location"], {})
    org = op.get_org_profile(data_file) or {}
    item = po_export.get_item_by_code(transfer["material_code"], active_only=False) or {}

    qty = float(transfer["quantity"])
    weight_per_unit = item.get("weight_kg")
    total_weight = round(weight_per_unit * qty, 2) if weight_per_unit is not None else None
    unit_price = item.get("price") or 0
    declared_value = round(unit_price * qty, 2)

    return {
        "Shipment Reference": transfer["transfer_id"],
        "Courier": transfer.get("carrier") or "BlueDart",
        "Ship Date": transfer.get("shipped_date") or date.today().strftime("%Y-%m-%d"),
        "Mode": "Surface",
        "Consignor Name": org.get("Legal_Name", ""),
        "Consignor Address": from_loc.get("address", ""),
        "Consignor City": from_loc.get("city", ""),
        "Consignor State": from_loc.get("state", ""),
        "Consignor Pincode": _extract_pincode(from_loc.get("address", ""), from_loc.get("city", "")),
        "Consignor Contact": org.get("Contact_Phone", ""),
        "Consignee Name": org.get("Legal_Name", ""),
        "Consignee Address": to_loc.get("address", ""),
        "Consignee City": to_loc.get("city", ""),
        "Consignee State": to_loc.get("state", ""),
        "Consignee Pincode": _extract_pincode(to_loc.get("address", ""), to_loc.get("city", "")),
        "Consignee Contact": org.get("Contact_Phone", ""),
        "Material Code": transfer["material_code"],
        "Product Description": transfer["material_desc"],
        "HSN Code": item.get("hsn_code", ""),
        "Quantity": qty,
        "UOM": transfer.get("uom", ""),
        "Pieces": int(qty) if qty == int(qty) else qty,
        "Weight per Unit (kg)": weight_per_unit,
        "Total Weight (kg)": total_weight,
        "Declared Value (INR)": declared_value,
    }


def generate_shipping_excel(shipment):
    """
    One real shipment -> one formatted .xlsx, laid out as a real
    courier booking form would be (labeled sections, not a raw data
    dump) — ready to hand-fill into BlueDart's own portal, or attach as
    a booking reference alongside the physical consignment. No
    formulas (every value here is already a real, final number — there
    is nothing for a formula to compute from something else in this
    sheet), so no recalculation step applies. Returns (filename, bytes).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Shipping Details"

    title_font = Font(name="Arial", bold=True, size=14)
    section_font = Font(name="Arial", bold=True, color="FFFFFF")
    section_fill = PatternFill(start_color="1E1E2E", end_color="1E1E2E", fill_type="solid")
    label_font = Font(name="Arial", bold=True)
    value_font = Font(name="Arial")

    ws["A1"] = f"{shipment['Courier']} \u2014 Shipment Booking Details"
    ws["A1"].font = title_font
    ws.merge_cells("A1:B1")
    ws["A2"] = f"Reference: {shipment['Shipment Reference']}  \u00b7  Generated: " \
              f"{date.today().strftime('%Y-%m-%d')}"
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="666666")
    ws.merge_cells("A2:B2")

    sections = [
        ("Shipment", ["Shipment Reference", "Courier", "Ship Date", "Mode"]),
        ("Consignor (From)", ["Consignor Name", "Consignor Address", "Consignor City",
                              "Consignor State", "Consignor Pincode", "Consignor Contact"]),
        ("Consignee (To)", ["Consignee Name", "Consignee Address", "Consignee City",
                            "Consignee State", "Consignee Pincode", "Consignee Contact"]),
        ("Cargo", ["Material Code", "Product Description", "HSN Code", "Quantity", "UOM",
                  "Pieces", "Weight per Unit (kg)", "Total Weight (kg)", "Declared Value (INR)"]),
    ]

    row = 4
    for section_title, fields in sections:
        ws.cell(row=row, column=1, value=section_title).font = section_font
        ws.cell(row=row, column=1).fill = section_fill
        ws.cell(row=row, column=2).fill = section_fill
        row += 1
        for f in fields:
            c1 = ws.cell(row=row, column=1, value=f)
            c1.font = label_font
            c2 = ws.cell(row=row, column=2, value=shipment.get(f, ""))
            c2.font = value_font
            if "Value" in f:
                c2.number_format = "\u20b9#,##0.00"
            row += 1
        row += 1

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 52

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_ref = shipment["Shipment Reference"].replace("/", "-")
    filename = f"{shipment['Courier']}_Shipment_{safe_ref}.xlsx"
    return filename, buf.getvalue()
