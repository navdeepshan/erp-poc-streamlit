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
