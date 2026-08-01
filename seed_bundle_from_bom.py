"""
seed_bundle_from_bom.py — one-time seed: creates a Purchase Bundle from
the full leaf-level explosion of the Bandicoot BOM (GRB-FG-0001).

Demo/seed purpose: shows Purchase Bundles and BOM working side by side
on the same underlying material data, without the two capabilities
having any actual code dependency on each other (purchase_bundles.py
never imports bom.py, and bom.py's explosion is only used here, once,
to source the bundle's initial contents — after creation the bundle is
just ordinary purchase_bundle_items rows, same as one built by hand
through the UI).

Idempotent by name: if a bundle with this exact name already exists
(active or not), it's left alone and this script is a no-op — safe to
re-run after a fresh migration without creating duplicates.
"""

import bom
import purchase_bundles as pb

BUNDLE_NAME = "Bandicoot Standard Unit — Full Build Kit"
PARENT_CODE = "GRB-FG-0001"
DEPARTMENT = "MFG-Production"
CREATED_BY = "SYSTEM-SEED"


def seed(quantity=1):
    existing = [b for b in pb.list_bundles(active_only=False) if b["bundle_name"] == BUNDLE_NAME]
    if existing:
        print(f"'{BUNDLE_NAME}' already exists ({existing[0]['bundle_id']}) — skipping.")
        return existing[0]["bundle_id"]

    detailed = bom.explode_bom_detailed(PARENT_CODE, quantity)
    if not detailed:
        raise RuntimeError(f"{PARENT_CODE} has no BOM to seed from — is bom migration data loaded?")

    items = [{"mat_code": d["mat_code"], "qty": d["gross_qty"]} for d in detailed]
    bundle_id = pb.create_bundle(
        BUNDLE_NAME,
        description=(f"Every leaf-level component needed to build one "
                     f"{PARENT_CODE} (Bandicoot Robotic Scavenger - Standard Unit), "
                     f"seeded from its full BOM explosion — {len(items)} materials."),
        department=DEPARTMENT, created_by=CREATED_BY, items=items,
    )
    print(f"Created {bundle_id}: '{BUNDLE_NAME}' with {len(items)} line item(s).")
    return bundle_id


if __name__ == "__main__":
    seed()
