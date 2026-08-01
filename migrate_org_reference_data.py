"""
migrate_org_reference_data.py — seeds Org_Defaults, Legal_Entities,
Vendor_Types, and Customer_Types with real Genrobotics values.

One-time seed script, matching the established pattern of every other
migrate_*.py in this project (e.g. migrate_o2c_accounting.py's
SEED_ACCOUNTS). Run once against a fresh or existing erp_pilot.db;
safe to re-run — every insert is an upsert, nothing gets duplicated.

Values here match the SeedData_GRB.xlsx sheet the same names describe
— this is the SQLite-side seed for the running PoC, not the upload
pipeline itself (that's a separate, larger feature — see Item 5 in the
industry-profile-switching conversation).
"""

import org_defaults as od
import legal_entities as le
import vendor_onboarding as vo
import customer_onboarding as co
import db

ORG_DEFAULTS = {
    "PO Type": "NB",
    "Purchasing Group": "GRB_PG",
    "Legal Entity": "GRB_001",
    "Currency": "INR",
    "Purchasing Entity": "GRB_001",
    "Plant": "GRB_DL_PKD_Factory",
    "Demand Detection Mode": "Manufactured Items Only",
}

LEGAL_ENTITY = {
    "le_id": "GRB_001",
    "fields": {
        "LE_Name": "Genrobotic Innovations Private Limited",
        "GSTIN": "32AAGCG8901P1ZA",
        "PAN": "AAGCG8901P",
        "Address": "Genrobotic Innovations PVT LTD, CDAC Building, Technopark Campus",
        "City": "Thiruvananthapuram 695581", "State": "Kerala", "Country": "India",
        "Bank_Account_No": "50100200400800", "IFSC": "HDFC0000996", "Bank_Name": "HDFC Bank",
        "Contact_Email": "info@genrobotics.org", "Contact_Phone": "+91 99616 16166",
    },
}

VENDOR_TYPES = ["Electronics & Controls", "Mechanical & Fabrication", "Power & Battery Systems",
                "Pneumatics & Fluid Power", "Fasteners & Hardware", "Cables & Connectors",
                "International Sourcing", "Distributor / Reseller", "Service Provider", "Other"]

CUSTOMER_TYPES = ["Municipal Corporation", "Water & Sewerage Board", "Government / PSU",
                  "Industrial Plant", "Sanitation Contractor", "Distributor", "Other"]

VENDOR_TYPE_ASSIGNMENTS = {
    "IPCS": "Electronics & Controls", "ARKPWR": "Electronics & Controls",
    "KELTRON": "Electronics & Controls", "PLINTRO": "Electronics & Controls",
    "RAGHAV": "Pneumatics & Fluid Power", "PROBOTS": "Mechanical & Fabrication",
    "COSTAPWR": "Power & Battery Systems", "SINGHANIA": "Fasteners & Hardware",
    "SERIALSYS": "International Sourcing", "CREATIVEAUTO": "International Sourcing",
    "VIASION": "Cables & Connectors",
}


def run():
    db.init_schema()

    od.set_org_defaults(ORG_DEFAULTS)
    print(f"Org_Defaults: seeded {len(ORG_DEFAULTS)} value(s).")

    le.upsert_legal_entity(LEGAL_ENTITY["le_id"], LEGAL_ENTITY["fields"])
    print(f"Legal_Entities: seeded {LEGAL_ENTITY['le_id']}.")

    conn = db.get_connection()
    try:
        for vt in VENDOR_TYPES:
            conn.execute("INSERT OR IGNORE INTO vendor_types (vendor_type) VALUES (?)", (vt,))
        for ct_ in CUSTOMER_TYPES:
            conn.execute("INSERT OR IGNORE INTO customer_types (customer_type) VALUES (?)", (ct_,))
        conn.commit()
    finally:
        conn.close()
    print(f"Vendor_Types: seeded {len(VENDOR_TYPES)} value(s).")
    print(f"Customer_Types: seeded {len(CUSTOMER_TYPES)} value(s).")

    backfilled = 0
    for vendor_id, vtype in VENDOR_TYPE_ASSIGNMENTS.items():
        existing = vo.get_vendor(vendor_id)
        if existing is not None:
            vo.upsert_vendor(vendor_id, {"Vendor_Type": vtype})
            backfilled += 1
    print(f"Vendor_Master: backfilled Vendor_Type on {backfilled} existing vendor(s).")


if __name__ == "__main__":
    run()
