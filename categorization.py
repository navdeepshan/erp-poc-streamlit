"""
categorization.py — Autonomous item categorization.

Keyword/taxonomy-based classifier for the Item Master (Category / Sub-Category
columns are currently blank for all 609 items). Deterministic and fully
explainable (returns matched keywords + a confidence band) rather than a
black-box call — important for a procurement audit trail, and it means the
PoC has zero external dependency / API cost.

Design note: this is intentionally rule-based rather than LLM-based for v1.
It's instant, free, and 100% reproducible for a demo. A future version could
swap classify_item() for an LLM call (e.g. via the Anthropic API) for the
"Uncategorized" long tail without changing anything else in this module's
contract — preview_categorization()/apply_categorization() would not need
to change.
"""

import openpyxl
import os, re
from collections import defaultdict

import db
import po_export

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

ITEM_MASTER_SHEET = "Item Master"

# ── Taxonomy: Category -> Sub-Category -> keyword list ────────────────────────
# Keywords are matched against normalized (lowercased) description + tags text.
#
# The dental/medical categories below were substantially expanded from their
# original small example set after being run against a real ~4,000-item
# dental/oral-surgery/ENT supply catalog and finding a ~7-8% hit rate. Every
# keyword added below came directly from that real data (brand names, generic
# terms, and abbreviations actually seen in it) — not guessed. Re-run against
# a representative sample after expansion: hit rate went from 0/24 to 20/24
# on the specific sample used to spot the gap (see conversation history /
# commit notes — not reproduced as a comment here since the exact sample
# was one-off). The Robotics categories are untouched — they were built
# for a different demo dataset and this catalog never touches them.
TAXONOMY = {
    "Endodontics": {
        "Root Canal Files & Reamers": [
            "k-file", "k file", "k files", "h-file", "h file", "h files", "flexile", "reamer",
            "gated drill", "gates drill", "protaper", "pro taper", "rotary file", "fo-32", "iso 257",
            "hyflex", "neoendo", "finger spreader", "finger plugger", "paste carrier",
            "broach", "broaches", "niti", "path file", "glide file", "glide flex",
            "c+file", "endo access bur", "endo z", "root canal file", "tri angular file",
            "smart file", "unitaper", "powertaper", "heatflex", "little files",
            "paper point", "proiso", "dia gp", "dia pp",
        ],
        "Endo Motors & Apex Locators": [
            "endo motor", "endomotor", "apex locator", "root zx", "canal pro", "canalpro",
            "i-pex", "ipex", "endo activator", "obturation unit", "x-smart", "cordless endomotor",
            "woodpex", "j morita", "smartlite", "endodontic probe",
        ],
        "Irrigation & Chelating Agents": [
            "edta", "sodium hypochlorite", "hypochlor", "rc prep", "rcprep", "glyde",
            "parcan", "chlorhexidine", "chloro hx", "canal prep", "endo l ", "irrigator",
            "irrigant", "camphenol",
        ],
        "Obturation Materials": [
            "gutta", "gp point", "gp cutter", "sealer", "ah plus", "diacomp", "mta",
            "biodentine", "bio-c", "bioceramic", "calcium hydroxide", "metapex",
            "calplus", "calx", "calapex", "rc cal", "rc help", "rc seal", "rootfyx",
            "d pulp", "d vitz", "vitapex", "endoseal", "endoflas", "obturation",
        ],
    },
    "Orthodontics": {
        "Brackets & Tubes": ["bracket", "molar tube", "buccal tube", "mbt", "lingual button"],
        "Wires & Auxiliaries": [
            "arch wire", "archwire", "elastic chain", "crimpable hook", "ligature", "coil spring",
            "niti", "cu-niti", "cu niti", "garmy", "elastic module", "power chain",
            "ipr", "proximal strip", "rotation wedge",
        ],
        "Bands & Separators": ["band", "separator"],
        "Retainers & Appliances": [
            "retainer", "myobrace", "trainer for kids", "lingual retainer", "lingual sheath",
        ],
    },
    "Restorative & Aesthetic": {
        "Composites & Bonding": [
            "composite", "filtek", "bond", "adper", "etchant", "etch", "z250", "z350",
            "tetric", "empress", "beautifil", "opalescence", "variolink", "multilink",
            "adhese", "primer", "flowable", "flow refill", "te econom", "brill flow",
            "compoflo", "compomax", "helioseal", "optragate", "optrasculpt",
            "ivoclean", "avue",
        ],
        "Cements & Liners": [
            "cement", "gic", "liner", "cold cure", "self cure", "tmp", "relyx", "ketac",
            "fuji", "glass ionomer", "poly f", "zinc phosphate", "dycal", "calcimol",
            "ionomer", "luting", "lute", "zinc oxide", "zol ", "riva",
        ],
        "Impression & Bite Materials": [
            "impression", "alginate", "putty", "algitex", "algitray", "zetaplus",
            "zeta plus", "oranwash", "elite hd", "aquasil", "imprint", "bite registration",
            "occlufast", "indurent", "zeta 5", "picosep", "separating agent",
            "bite wax", "bite wafer",
        ],
    },
    "Surgical Instruments": {
        "Hand Instruments": [
            "forceps", "elevator", "curved", "castroviejo", "bone file",
            "cutting knife", "scalpel", "dontriks", "curette", "curett", "elevator",
            "retractor", "scissor", "rongeur", "clamp", "mouth gag", "probe",
            "excavator", "spatula", "burnisher", "condenser", "plugger", "carver",
            "explorer", "tweezer", "hemostat", "artery forcep", "crile",
        ],
        "Bone Surgery & Osteotomy": [
            "chisel", "osteotome", "gouge", "bone scoop", "bone awl", "bone hammer",
            "mallet", "bone plate", "bone graft", "rongeur",
        ],
        "Craniomaxillofacial & ENT": [
            "distractor", "distraction osteogenesis", "mandibular", "cleft", "nasal septum",
            "mastoid", "tonsil", "septum", "nasal", "orbital floor", "condyle",
        ],
        "Implant & Bone Graft": ["b-ostin", "bone graft", "implant", "btcp", "ht ", "osseograft", "periograft"],
    },
    "Diagnostic & Monitoring": {
        "Vital Signs": [
            "littman", "stethoscope", "bp monitor", "bp apparatus", "thermometer",
            "pulse oximeter", "nebulizer", "oxygen concentrator", "oxygen cylinder",
            "digital weighing scale",
        ],
        "Test Strips & Diagnostics": ["accu-check", "accu check", "strips", "glucose", "glucometer", "lancet"],
        "Imaging & Radiography": [
            "x-ray film", "x ray film", "rvg", "developer", "fixer", "sensor sleeve",
            "panoramic", "vistascan", "intraoral camera",
        ],
    },
    "Injection & Syringes": {
        "Needles & Syringes": [
            "dispovan", "needle", "syringe", "unolock", "insulin syr",
        ],
        "Local Anesthetics": [
            "lignox", "lignospan", "septanest", "lox 2%", "xylocaine",
        ],
    },
    "Infection Control & Sterilization": {
        "Disinfectants & Sterilants": [
            "bacillol", "disinfect", "sterile", "autoclave", "formoline", "formo-cresol",
            "formocresol", "cidex", "korsolex", "glutaraldehyde", "hypochlorite",
            "hand sanitizer", "antiseptic", "surfacept",
        ],
        "Sterilization Pouches & Reels": [
            "sterilization pouch", "sterilization reel", "cssd", "sterilization tape",
            "sterilization cassette", "steam indicator",
        ],
        "PPE & Barrier": [
            "glove", "mask", "gown", "barrier", "bouffant cap", "shoe cover",
            "patient apron", "patient bib", "eyewear", "protective eyewear",
        ],
    },
    "Equipment & Handpieces": {
        "Handpieces & Motors": [
            "handpiece", "hand piece", "airotor", "coupling", "fx-65", "contra angle",
            "micro motor", "micromotor", "marathon", "sprint contra angle",
        ],
        "Fixtures & Stands": ["stand", "dispenser"],
    },
    "General Consumables": {
        "Cotton & Gauze": ["cotton", "gauze", "sponge"],
        "Articulating & Misc": [
            "articulating paper", "arti fol", "arti paper", "shimstock", "mouthwash",
            "mint", "gum mono", "clean tray",
        ],
    },
    "Dental Lab & Prosthetics": {
        "Waxes": [
            "modelling wax", "modeling wax", "carving wax", "inlay wax", "boxing wax",
            "relief wax", "sticky wax", "utility wax", "casting wax", "beading wax",
            "wax knife", "wax spatula", "wax pot",
        ],
        "Plaster, Stone & Casting": [
            "dental stone", "plaster", "die stone", "casting ring", "crucible",
            "articulator", "dental flask", "die spacer", "die hardener", "die lubricant",
            "investment", "acry pol",
        ],
        "Alloys & Acrylics": [
            "nickel chrome", "ni-cr", "heat cure", "cold cure", "acrylic trimmer",
            "denture", "self cure acrylic", "chrome alloy", "dpi ",
        ],
    },
    "Sutures & Wound Closure": {
        "Sutures": [
            "vicryl", "mersilk", "prolene", "silk", "suture", "lifeline", "lifesilk",
            "lifechrome", "lifepga",
        ],
    },
    "Matrices, Wedges & Rubber Dam": {
        "Matrices & Wedges": [
            "matrix retainer", "matrix band", "saddle matrix", "wedge", "palodent",
            "toffelmire",
        ],
        "Rubber Dam": ["rubber dam", "dam clamp", "dam punch", "dam frame", "dam forcep"],
    },
    "Rotary Cutting Instruments": {
        "Burs": [
            "carbide bur", "diamond bur", "diamond rotary instrument", "fg bur",
            "ra bur", "hp bur", "endo access bur", "trimmer bur",
        ],
        "Discs & Polishers": [
            "cutting disc", "polishing disc", "finishing", "polishing kit", "polishing paste",
            "polishing cup", "abrasive disc",
        ],
    },
    "Robotics - Finished Goods": {
        "Robots & Major Assemblies": [
            "bandicoot", "robotic scavenger", "robotic arm assembly", "sensor module - complete",
            "chassis assembly", "power module assembly", "pneumatic drive assembly",
        ],
    },
    "Robotics - Mechanical": {
        "Actuation & Drives": [
            "servo motor", "gear motor", "planetary gearbox", "linear actuator", "gearbox",
        ],
        "Arm & Gripper": [
            "gripper", "robotic arm linkage", "universal joint", "rotary joint bearing",
            "coupling precision",
        ],
        "Chassis & Mobility": [
            "caterpillar track", "track roller", "chassis frame", "drive wheel hub",
            "suspension damper", "wheel motor mount", "cnc machined",
        ],
        "Fasteners & Hardware": [
            "hex bolt", "wing nut", "mounting bracket", "washer set", "l-shape aluminum",
        ],
    },
    "Robotics - Electronics & Controls": {
        "Sensors & Vision": [
            "camera module", "thermal imaging", "ultrasonic distance sensor", "imu",
            "inertial measurement", "pressure sensor", "proximity sensor",
        ],
        "Control Systems": [
            "motor driver board", "microcontroller board", "control pcb", "relay module",
            "motion controller",
        ],
        "Power Systems": [
            "battery pack", "battery charger", "voltage regulator", "power distribution board",
        ],
        "Communication": [
            "rf transmitter", "rf receiver", "antenna", "lte communication",
        ],
    },
    "Robotics - Pneumatics": {
        "Pneumatic Components": [
            "pneumatic cylinder", "solenoid valve", "air compressor", "pressure regulator",
            "pneumatic hose", "quick connect fitting",
        ],
    },
    "Robotics - Safety & Environmental": {
        "Gas Detection": ["gas detector", "gas sensor module", "air quality monitor"],
        "Lighting": ["led floodlight", "floodlight", "headlamp module"],
    },
    "Robotics - Enclosures & Materials": {
        "Structural & Sealing": [
            "aluminum sheet", "enclosure box", "gasket seal", "cable gland",
        ],
    },
    "Robotics - Cables & Connectors": {
        "Wiring": ["wiring harness", "waterproof connector", "power cable"],
    },
}

# Flat list for building selectboxes in the UI, e.g. for manual overrides.
CATEGORY_CHOICES = sorted(TAXONOMY.keys()) + ["Uncategorized"]


def _subcat_choices(category):
    return sorted(TAXONOMY.get(category, {}).keys()) or ["Needs Review"]


def _normalize(text):
    return re.sub(r"[^a-z0-9\s\-]", " ", str(text).lower())


def classify_item(desc, tags=""):
    """Return (category, subcategory, confidence, matched_keywords list)."""
    hay = _normalize(desc) + " " + _normalize(tags)
    best = None  # (score, category, subcategory, matches)
    for cat, subcats in TAXONOMY.items():
        for subcat, keywords in subcats.items():
            matches = [kw for kw in keywords if kw in hay]
            score = len(matches)
            if score and (best is None or score > best[0]):
                best = (score, cat, subcat, matches)
    if not best:
        return ("Uncategorized", "Needs Review", "Low", [])
    score, cat, subcat, matches = best
    confidence = "High" if score >= 2 else "Medium"
    return (cat, subcat, confidence, matches)


def preview_categorization(data_file=None, only_blank=True):
    """
    Run the classifier over the Item Master without writing anything.
    Returns a list of row dicts ready to render in an editable table.

    Item Master now lives in SQLite. Deliberately reads with
    active_only=False (via po_export.load_item_master) — this module's
    job is making sure every item HAS a category, regardless of whether
    it's currently active; Active gates what's newly selectable, not
    whether master data is allowed to be complete. Keyed by item_code
    now instead of an Excel row number (SQLite rows have no stable row
    position) — apply_categorization() below writes back by code.
    """
    items = po_export.load_item_master(data_file, active_only=False)
    out = []
    for item in items:
        code = item["code"]
        cat_existing = item["category"]
        if only_blank and cat_existing:
            continue
        desc = item["desc"]
        tags = item["tags"]
        cat, subcat, conf, matches = classify_item(desc, tags)
        out.append({
            "code": code, "desc": desc, "tags": tags,
            "category": cat, "subcategory": subcat,
            "confidence": conf, "matched_keywords": ", ".join(matches),
        })
    return out


def apply_categorization(rows, data_file=None):
    """
    Write category/subcategory back to the Item Master.
    rows: list of dicts with 'code', 'category', 'subcategory' (as
    returned/edited from preview_categorization()). Writes by item_code
    now (SQLite UPDATE ... WHERE item_code = ?), replacing the old
    Excel-row-number write.
    """
    conn = db.get_connection()
    try:
        n = 0
        for r in rows:
            conn.execute(
                "UPDATE item_master SET category = ?, subcategory = ? WHERE item_code = ?",
                (r["category"], r["subcategory"], r["code"]),
            )
            n += 1
        conn.commit()
    finally:
        conn.close()
    return n


def stats(data_file=None):
    """Summary counts for the dashboard: total / categorized / by-category breakdown.
    Deliberately active_only=False, same reasoning as preview_categorization()."""
    items = po_export.load_item_master(data_file, active_only=False)
    total = len(items)
    categorized = 0
    by_cat = defaultdict(int)
    for item in items:
        cat = item["category"]
        if cat:
            categorized += 1
            by_cat[cat] += 1
    return {"total": total, "categorized": categorized,
            "uncategorized": total - categorized, "by_category": dict(by_cat)}


if __name__ == "__main__":
    s = stats()
    print(f"Item Master: {s['total']} items, {s['categorized']} categorized, "
          f"{s['uncategorized']} pending.")
    preview = preview_categorization()
    print(f"\nClassifying {len(preview)} uncategorized item(s)...")
    for p in preview[:15]:
        print(f"  {p['code']:14s} [{p['confidence']:6s}] {p['category']} / "
              f"{p['subcategory']:28s} <- {p['desc'][:45]}")
    conf_counts = defaultdict(int)
    for p in preview:
        conf_counts[p["confidence"]] += 1
    print(f"\nConfidence breakdown: {dict(conf_counts)}")
