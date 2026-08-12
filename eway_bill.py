"""
eway_bill.py — E-Way Bill generation for inter-state goods movement.

A real, clearly-labeled simulation, same honesty standard as
shipping.py's mock courier connectors: no live GSTN E-Way Bill API is
called here (no real credentials/access exists for this either).
Structured to be swappable for a real integration later without
changing anything that calls generate_eway_bill() — only its own
internal simulation would need to be replaced.

Built for STO-US-02 first (inter-state Hub-to-Plant STOs), but
deliberately general — is_eway_bill_required() and generate_eway_bill()
take a from/to location pair and a value, nothing STO-specific, so a
real future ad hoc-transfer or O2C dispatch integration (FUL-US-05)
can reuse this directly rather than duplicating the same real
threshold/validity logic.
"""

import math
import random
from datetime import date, timedelta

EWAY_BILL_VALUE_THRESHOLD = 50000  # INR — the real, standard national
                                   # threshold most states use for
                                   # inter-state movement. A few states
                                   # set their own lower threshold for
                                   # intra-state movement specifically —
                                   # not modeled here, since this
                                   # system only ever considers
                                   # inter-state movement for this check.

KM_PER_VALIDITY_DAY = 200  # Real government rule for "regular" cargo
                           # (not over-dimensional): one additional day
                           # of validity per 200km travelled, minimum
                           # one day regardless of how short the route.


def _haversine_km(geo1, geo2):
    """Real great-circle distance between two real 'lat,lon' strings,
    not a flat assumption — every delivery location in this system
    already carries real geo-coordinates."""
    lat1, lon1 = (float(x) for x in geo1.split(","))
    lat2, lon2 = (float(x) for x in geo2.split(","))
    R = 6371  # Earth radius, km
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def is_inter_state(from_location, to_location, data_file=None):
    """
    Real, reusable state comparison, not duplicated per caller: both
    is_eway_bill_required() below and inventory.py's own GST-deemed-supply
    GL posting need to know whether a movement crosses a real state
    boundary. E-way bill additionally applies its own value threshold on
    top of this; GST's own deemed-supply IGST treatment applies to any
    inter-state movement regardless of value, so this stays a pure state
    comparison with no threshold baked in.
    """
    import pr_consolidation as pc
    locs = {l["id"]: l for l in pc.get_delivery_locations(active_only=False)}
    from_state = locs.get(from_location, {}).get("state")
    to_state = locs.get(to_location, {}).get("state")
    if not from_state or not to_state:
        return False
    return from_state != to_state


def is_eway_bill_required(from_location, to_location, declared_value, data_file=None):
    """
    A real rule, not a placeholder that always returns True for
    every STO regardless of whether the real rule would actually
    require one: required only for genuine inter-state movement
    (comparing each location's own real State) at or above the real
    national value threshold. Intra-state movement, or any movement
    below threshold, correctly returns False.
    """
    if not is_inter_state(from_location, to_location, data_file=data_file):
        return False
    return declared_value >= EWAY_BILL_VALUE_THRESHOLD


def generate_eway_bill(from_location, to_location, declared_value, transporter_id="",
                       vehicle_no="", data_file=None):
    """
    SIMULATED — no real GSTN API call. Real EWB number shape (12
    digits), real distance-based validity computed from the haversine
    distance between the two real delivery locations' own geo-
    coordinates (not a flat assumption applied to every shipment
    regardless of how far apart the two locations actually are).

    Not idempotent by itself — a real e-way bill is generated once per
    real dispatch event, and the caller is responsible for persisting
    the result and not calling this again for the same shipment; this
    function is a pure generator, the same separation of concerns
    shipping.py's own payload builders use.
    """
    import pr_consolidation as pc
    locs = {l["id"]: l for l in pc.get_delivery_locations(active_only=False)}
    from_loc = locs.get(from_location, {})
    to_loc = locs.get(to_location, {})
    distance_km = round(_haversine_km(from_loc.get("geo", "0,0"), to_loc.get("geo", "0,0")))

    validity_days = max(1, math.ceil(distance_km / KM_PER_VALIDITY_DAY))
    generated_date = date.today()
    valid_until = generated_date + timedelta(days=validity_days)

    ewb_number = f"{random.randint(10**11, 10**12 - 1)}"

    return {
        "ewb_number": ewb_number,
        "generated_date": str(generated_date),
        "valid_until": str(valid_until),
        "distance_km": distance_km,
        "validity_days": validity_days,
        "declared_value": declared_value,
        "from_location": from_location, "to_location": to_location,
        "from_state": from_loc.get("state"), "to_state": to_loc.get("state"),
        "transporter_id": transporter_id, "vehicle_no": vehicle_no,
    }


def generate_eway_bill_document(transfer_id, data_file=None):
    """
    A real, downloadable PDF for a transfer that already has a real
    e-way bill number on file — a real gap found and fixed directly:
    the number was visible on screen, but there was never anything an
    actual carrier driver could be handed. Recomputes the supporting
    details (distance, declared value, addresses) from the transfer's
    own real, persisted data rather than storing them a second time —
    the only two values that must be read back exactly as originally
    generated are the transfer's own real eway_bill_number and eway_
    bill_valid_until, both already on the stock_transfers row.

    Raises ValueError if this transfer never actually got a real
    e-way bill (nothing to document) — a caller shouldn't be able to
    generate a document claiming compliance that was never triggered.

    Returns (filename, pdf_bytes).
    """
    import inventory as inv
    import pr_consolidation as pc
    import po_export
    import org_profile as op

    t = inv.get_stock_transfer(transfer_id, data_file)
    if t is None:
        raise ValueError(f"{transfer_id} not found.")
    if not t.get("eway_bill_number"):
        raise ValueError(f"{transfer_id} has no real e-way bill on file — nothing to document.")

    locs = {l["id"]: l for l in pc.get_delivery_locations(active_only=False)}
    from_loc = locs.get(t["from_location"], {})
    to_loc = locs.get(t["to_location"], {})
    item = po_export.get_item_by_code(t["material_code"], active_only=False)
    org = op.get_org_profile(data_file) or {}

    distance_km = round(_haversine_km(from_loc.get("geo", "0,0"), to_loc.get("geo", "0,0")))
    unit_price = item["price"] if item else 0
    declared_value = round(unit_price * t["quantity"], 2)
    hsn_code = item["hsn_code"] if item else ""

    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    import io

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=18 * mm, bottomMargin=18 * mm,
                            leftMargin=18 * mm, rightMargin=18 * mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("EWBTitle", parent=styles["Title"], fontSize=16, spaceAfter=2)
    small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=8, textColor=colors.grey)
    label_style = ParagraphStyle("Label", parent=styles["Normal"], fontSize=9,
                                 textColor=colors.HexColor("#555555"))
    value_style = ParagraphStyle("Value", parent=styles["Normal"], fontSize=11)

    elements = [
        Paragraph("E-WAY BILL", title_style),
        Paragraph("Simulated document \u2014 this PoC generates a real-shaped EWB number and "
                 "distance-based validity period, but never calls the real GSTN e-way bill "
                 "API. Not valid for actual movement of goods.", small),
        Spacer(1, 8 * mm),
    ]

    header_table = Table([
        [Paragraph("<b>E-Way Bill No.</b>", label_style), Paragraph(t["eway_bill_number"], value_style)],
        [Paragraph("<b>Generated</b>", label_style), Paragraph(t["shipped_date"] or "", value_style)],
        [Paragraph("<b>Valid Until</b>", label_style), Paragraph(t["eway_bill_valid_until"] or "", value_style)],
        [Paragraph("<b>Distance</b>", label_style), Paragraph(f"{distance_km} km", value_style)],
    ], colWidths=[45 * mm, 120 * mm])
    header_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#1E3A5F")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#DCE6F1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 8 * mm))

    def _party_block(title, name, loc, gstin=None):
        addr = loc.get("address", "")
        city_state = f"{loc.get('city', '')}, {loc.get('state', '')}"
        gstin_line = f"GSTIN: {gstin}" if gstin else "GSTIN: (own location \u2014 stock transfer, not a sale)"
        return [Paragraph(f"<b>{title}</b>", label_style),
               Paragraph(name, value_style),
               Paragraph(addr, styles["Normal"]),
               Paragraph(city_state, styles["Normal"]),
               Paragraph(gstin_line, small)]

    party_table = Table([
        [_party_block("From (Consignor)", from_loc.get("name", t["from_location"]), from_loc,
                      gstin=org.get("GSTIN")),
         _party_block("To (Consignee)", to_loc.get("name", t["to_location"]), to_loc)],
    ], colWidths=[82 * mm, 82 * mm])
    party_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#1E3A5F")),
        ("LINEAFTER", (0, 0), (0, -1), 0.75, colors.HexColor("#1E3A5F")),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    elements.append(party_table)
    elements.append(Spacer(1, 8 * mm))

    elements.append(Paragraph("<b>Goods Details</b>", label_style))
    elements.append(Spacer(1, 2 * mm))
    goods_table = Table([
        ["Description", "HSN", "Qty", "UOM", "Declared Value (INR)"],
        [t["material_desc"], hsn_code, f"{t['quantity']:g}", t.get("uom", "pcs"),
         f"{declared_value:,.2f}"],
    ], colWidths=[62 * mm, 24 * mm, 20 * mm, 20 * mm, 38 * mm])
    goods_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A5F")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(goods_table)
    elements.append(Spacer(1, 8 * mm))

    transport_table = Table([
        [Paragraph("<b>Transporter</b>", label_style), Paragraph(t.get("carrier") or "\u2014", value_style)],
        [Paragraph("<b>Vehicle No.</b>", label_style), Paragraph("_______________________ "
                  "(to be filled at dispatch)", styles["Normal"])],
        [Paragraph("<b>Mode</b>", label_style), Paragraph("Road", value_style)],
    ], colWidths=[45 * mm, 120 * mm])
    transport_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#1E3A5F")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#DCE6F1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    elements.append(transport_table)
    elements.append(Spacer(1, 10 * mm))
    elements.append(Paragraph(f"Real transfer reference: {transfer_id}"
                              + (f" \u00b7 STO {t['source_doc']}" if t.get("source_type") == "STO"
                                 and t.get("source_doc") else ""), small))

    doc.build(elements)
    pdf_bytes = buf.getvalue()
    buf.close()
    filename = f"EWayBill_{t['eway_bill_number']}.pdf"
    return filename, pdf_bytes
