"""
demo_profiles.py — the demo scenario registry, one source instead of two.

Extracted 2026-07-31 from o2c_ui.py's page_settings(), which previously
defined this dict inline. Now imported by both o2c_ui.py and
agent_console.py, so which pilot supports which demo action, and the
pilot-specific details of that action, are decided in exactly one
place — found a real bug from NOT doing this consistently enough the
first time: o2c_ui.py's own "Add New Customer Wave" button still had
Genrobotics-specific text and a hardcoded check for customer "SMC"
baked directly into the page, which would have been simply wrong once
IDS Denmed's own wave (\"SMC\" doesn't exist there) was enabled. Fixed
by moving the pilot-specific pieces (which customer ID proves the wave
already ran, and the caption describing what it does) into each
profile here, same as everything else pilot-specific.
"""

DEMO_PROFILES = {
    "GRB": {
        "module": "demo_scenario", "title": "\"Two City Contracts, One Smart Transfer\"",
        "description": (
            "R&D has surplus battery/servo-motor stock from a wrapped-up pilot; "
            "Factory is short on both once Mysore and Indore place real fleet "
            "orders. A transfer fully closes the battery gap and mostly closes "
            "the servo-motor one, competitive RFx sources the honest remainder, "
            "Purchase-Bundle procurement covers the rest of the BOM, 2 of the 4 "
            "units get built and carried through to cash (Mysore paid in full, "
            "Indore partially — the other 2 units stay visibly in progress)."),
        "setup_note": ("Mysore and Indore's orders are on file, and the stock "
            "imbalance should be visible now on the Manufacturing app's "
            "Inventory \u2192 Position & Transfers page."),
        "resolution_note": ("head to Accounting for the full ledger, or Cash "
            "Application to see Mysore's full payment and Indore's partial one."),
        "supports_customer_wave": True,
        "wave_check_customer_id": "SMC",
        "wave_caption": (
            "Three new municipal customers — Surat, Guwahati, Delhi Jal Board — "
            "each get a real quotation, a confirmed Sales Order, and a full "
            "bundle-based PR (the same ~50-line Bandicoot bundle each time), left "
            "genuinely Open. Run PR Consolidation again afterward and the merge "
            "spans potentially the whole BOM, three real customer orders deep."),
    },
    "IDS": {
        "module": "demo_scenario_ids", "title": "\"One Container, Four Cities\"",
        "description": (
            "A bulk import consignment lands entirely at Chennai. Four real "
            "customers in four different cities then place orders — Bangalore, "
            "Delhi, Mumbai, Hyderabad — while the stock still sits at the port. "
            "Transfers resolve three items completely; sterilization pouches "
            "genuinely run short even after the transfer, triggering a direct "
            "reorder to MELAG (not competitive RFx — a distributor reordering an "
            "exclusive branded line doesn't shop it around). Delhi is IDS "
            "Denmed's own HQ state (CGST+SGST); the other three are inter-state "
            "(IGST) — both tax paths in one scenario."),
        "setup_note": ("Four orders are on file across four cities, and the stock "
            "imbalance should be visible now on the Manufacturing app's "
            "Inventory \u2192 Position & Transfers page — Chennai fanning out to "
            "all four."),
        "resolution_note": ("head to Accounting for the full ledger, or Cash "
            "Application to see the mix of full and partial payments."),
        "supports_customer_wave": True,
        "wave_check_customer_id": "SMILEZONE-AMD",
        "wave_caption": (
            "Three new dental clinics — Ahmedabad, Kolkata, Jaipur — each get a "
            "real quotation, a confirmed Sales Order, and a full bundle-based PR "
            "(a new \"Dental Clinic Starter Kit\" bundle each time — chair, "
            "handpiece, scaler, endo motor, curing light, X-ray, sterilization "
            "gear — routed to Chennai, the import hub, not each clinic's own "
            "city, so they can actually merge), left genuinely Open. Run PR "
            "Consolidation again afterward and watch the merge span the whole "
            "kit, three real customer orders deep."),
    },
}


def current_profile(data_file=None):
    """Returns the active profile dict (or None if Org_ID doesn't match
    a known profile) based on the currently configured Org Profile."""
    import org_profile as op
    org_id = (op.get_org_profile(data_file) or {}).get("Org_ID")
    return DEMO_PROFILES.get(org_id), org_id
