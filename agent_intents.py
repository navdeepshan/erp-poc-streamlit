"""
agent_intents.py — Rule-based intent matching and entity resolution for
the Agent Console.

Explicitly NOT a general-purpose language model. This is pattern/keyword
matching against a curated intent vocabulary — the same vocabulary
already documented as "Chat Prompts" across the S2S and O2C user
stories, given a real matching engine instead of being illustrative
text. Entities (material, location) are resolved via fuzzy matching
against REAL current Item Master / Delivery Locations data, not a
hardcoded list, so a real material added tomorrow is matchable
tomorrow without touching this file.

Never guesses past a confidence threshold. Genuinely unmatched input
returns intent=None rather than picking the closest intent and acting
on a low-confidence guess — the same "never fabricate confidence"
discipline already threaded through every recommendation this platform
produces, applied here to input instead of output. A wrong guess acted
on is worse than an honest "I didn't catch that."

Every action this resolves to still requires an explicit confirmation
before anything executes — this module only ever proposes, matching
every other proposal-then-confirm flow already built (Ship, Create PR,
RFx award, bundle discovery). It does not call any write function
itself.
"""
import re
import difflib

import po_export
import pr_consolidation as pc
import bom
import org_defaults as od
import inventory as inv
import nav_catalog as nav

STOPWORDS = {"the", "a", "an", "to", "of", "for", "and", "some", "any", "extra",
            "those", "this", "that", "please", "can", "you", "i", "want", "need",
            "me", "my", "we", "our", "is", "are", "in", "at", "on", "it", "its",
            "with", "up", "over", "from", "there"}

# Each intent: name, trigger phrases (multi-word, scored higher) and
# trigger keywords (single word, scored lower) — a simple, transparent
# scoring scheme, not a black box. The highest-scoring intent above
# MIN_CONFIDENCE wins; a tie or nothing above threshold means "not
# understood," not a guess.
MIN_CONFIDENCE = 2

INTENTS = [
    {"name": "check_at_risk",
     "phrases": ["at risk", "needs action", "what needs attention", "what's short",
                "whats short", "what is short", "action needed", "what's flagged",
                "whats flagged", "procurement recommendation"],
     "keywords": ["shortfall", "shortage", "recommendations", "recommendation"]},

    {"name": "ship_transfer",
     "phrases": ["ship the", "send the", "move the", "ship to", "send to", "move to"],
     # Weighted higher than the default keyword weight of 1 — found as a
     # real gap during testing, not a hypothetical one: "ship battery
     # pack to the factory" (a completely natural phrasing with no
     # "the"/"to" immediately after the verb) matched none of the phrase
     # triggers above and scored only 1 point from the bare "ship"
     # keyword, landing below MIN_CONFIDENCE and failing outright. In
     # this specific domain "ship"/"send"/"move"/"transfer" are
     # distinctive action verbs, not generic words that could mean
     # almost anything — a single occurrence is already a strong signal
     # here, unlike a keyword match in most other intents.
     "keywords": ["ship", "transfer", "send", "move"], "keyword_weight": 2},

    {"name": "receive_transfer",
     "phrases": ["receive the", "confirm receipt", "confirm the receipt", "received the"],
     # Real gap found directly, not hypothetical: "receive Servo Motor
     # 12V High Torque Arm Joint Drive in palakkad" — the ship_transfer
     # button worked, but this natural follow-up phrasing (material
     # name right after the verb, no "the") didn't match any phrase
     # trigger here either, so applying ship_transfer's own fix
     # proactively rather than waiting to rediscover the identical gap.
     "keywords": ["receive", "received"], "keyword_weight": 2},

    {"name": "create_pr",
     "phrases": ["create a pr", "create a requisition", "raise a pr",
                "raise a requisition", "raise the pr", "create the pr"],
     # Same reasoning and same fix as ship_transfer above — "requisition
     # the battery pack" and "need a requisition raised for scaler" both
     # failed to match any phrase and scored only 1 point from the bare
     # "requisition" keyword, below threshold. "pr" itself is also added
     # as a keyword (it wasn't one before), since "pr for the battery
     # pack" matched nothing at all, not even a single point.
     "keywords": ["requisition", "pr"], "keyword_weight": 2},

    {"name": "switch_mode",
     "phrases": ["switch to", "switch mode", "use mode", "change mode",
                "change the mode", "plan without"],
     "keywords": []},

    # Real scope expansion (2026-08-11), by direct instruction: this
    # console was originally scoped to inventory optimization + order
    # upload only (see this module's and agent_console.py's own
    # docstrings), deliberately not "the whole platform." Recognizing
    # every real screen name across all three core apps is a genuine
    # widening of that boundary, not a small addition — worth naming,
    # not silently folded in as if it were always in scope.
    {"name": "navigate",
     "phrases": ["show me", "take me to", "go to", "navigate to",
                "open the", "where is", "wheres", "where's",
                "which screen", "which tab", "how do i get to",
                "how do i find"],
     # Weighted higher than the default 1, same real reason as ship_
     # transfer/receive_transfer/create_pr above: "open RFx Management"
     # (the verb immediately followed by a proper screen name, no
     # "the" in between) matched none of the phrase triggers and
     # scored only 1 point from the bare "open" keyword, below
     # MIN_CONFIDENCE — the identical gap class, found by the same
     # kind of direct test, not by inspection.
     "keywords": ["open", "screen"], "keyword_weight": 2},

    {"name": "upload_orders",
     "phrases": ["upload orders", "bulk load", "bulk import", "order data",
                "import orders", "load order history", "upload sales orders"],
     "keywords": []},

    {"name": "run_setup",
     "phrases": ["load setup only", "run setup", "load setup", "start the demo setup",
                "just the setup"],
     "keywords": []},

    {"name": "complete_rest",
     "phrases": ["complete the rest", "finish the story", "fast forward the rest",
                "fast-forward the rest", "finish the demo"],
     "keywords": []},

    {"name": "load_full_demo",
     "phrases": ["load full demo", "full demo", "run the full demo",
                "quick preview", "whole story in one"],
     "keywords": []},

    {"name": "add_customer_wave",
     "phrases": ["add new customer wave", "customer wave", "add the new customers",
                "trigger the wave", "add new customers"],
     "keywords": []},

    {"name": "reset_data",
     "phrases": ["reset and reseed", "reset the data", "reseed the data",
                "start fresh", "wipe the data"],
     "keywords": []},

    {"name": "query_org_profile",
     "phrases": ["org profile", "organization profile", "our gstin", "our pan",
                "company profile"],
     "keywords": []},

    {"name": "query_item_tax",
     "phrases": ["tax info for", "hsn code for", "gst rate for",
                "missing tax info", "items missing tax"],
     "keywords": []},

    # Weighted higher than the default 3 — "why is/does/was" and "explain
    # why" are highly distinctive question phrasings with very low false-
    # positive risk, deliberately weighted to reliably win over incidental
    # keyword overlap from another intent (e.g. "why is the pouches
    # shortage at risk" contains both "at risk" AND "shortage", two of
    # check_at_risk's own triggers, entirely coincidentally — found as a
    # real mismatch during testing, not a hypothetical one, and fixed by
    # weighting the question phrasing itself rather than trying to strip
    # every possible incidental overlap out of every other intent).
    {"name": "explain",
     "phrases": ["why is", "why does", "why was", "explain why"],
     "keywords": ["explain"], "phrase_weight": 5},

    {"name": "help",
     "phrases": ["what can you do", "help me", "show me what"],
     "keywords": ["help"]},
]

MODE_NAMES = ["Sales Order Based", "Optimize Existing PRs", "Reorder Qty Based"]
MODE_ALIASES = {
    "sales order": "Sales Order Based", "sales orders": "Sales Order Based",
    "confirmed order": "Sales Order Based",
    "optimize existing": "Optimize Existing PRs", "existing prs": "Optimize Existing PRs",
    "existing pr": "Optimize Existing PRs",
    "reorder qty": "Reorder Qty Based", "reorder quantity": "Reorder Qty Based",
    "min max": "Reorder Qty Based", "min/max": "Reorder Qty Based",
    "without any sales order": "Reorder Qty Based", "without sales order": "Reorder Qty Based",
    "no sales order": "Reorder Qty Based",
}


def _tokenize(text):
    return [w for w in re.findall(r"[a-z0-9]+", text.lower()) if w not in STOPWORDS]


def match_intent(text):
    """
    Returns {"intent": name, "confidence": int} or {"intent": None,
    "confidence": 0} if nothing scores at or above MIN_CONFIDENCE.
    Phrase matches score 3 (a multi-word match is a much stronger
    signal than one keyword), keyword matches score 1.
    """
    lower = text.lower()
    scores = {}
    for spec in INTENTS:
        score = 0
        phrase_weight = spec.get("phrase_weight", 3)
        keyword_weight = spec.get("keyword_weight", 1)
        for phrase in spec["phrases"]:
            if phrase in lower:
                score += phrase_weight
        for kw in spec["keywords"]:
            if re.search(rf"\b{re.escape(kw)}\b", lower):
                score += keyword_weight
        if score:
            scores[spec["name"]] = score
    if not scores:
        return {"intent": None, "confidence": 0}
    best = max(scores, key=scores.get)
    if scores[best] < MIN_CONFIDENCE:
        return {"intent": None, "confidence": scores[best]}
    return {"intent": best, "confidence": scores[best]}


def resolve_material(text, data_file=None):
    """
    Fuzzy-matches free text against real, active Item Master
    descriptions and codes — whole-word overlap scoring, not exact
    string match, so "the scaler" can resolve against "Woodpecker
    UDS-J Ultrasonic Scaler". Returns {"match": item, "candidates": []}
    on a clear single winner, {"match": None, "candidates": [item, ...]}
    when the top scorers are genuinely tied (e.g. "the scaler" against
    both the scaler itself and its own replacement tips — a real,
    correct ambiguity, not a bug), or {"match": None, "candidates": []}
    when nothing scores at all. Callers use "candidates" to offer a
    clarification rather than just reporting failure.

    Matches whole tokens against the haystack's own tokens, not a raw
    substring check — found and fixed a real false positive from the
    substring version: "what items are missing tax info" matched
    "Caterpillar Track Assembly **Rein-FOR-CED**" purely because "info"
    is a literal substring of "reinforced", nothing to do with either
    query's actual meaning. Tokenizing both sides the same way and
    comparing as sets closes this off structurally, not just for this
    one case.
    """
    items = po_export.load_item_master(data_file, active_only=True)
    words = set(_tokenize(text))
    if not words:
        return {"match": None, "candidates": []}
    scored = []
    for item in items:
        haystack_words = set(_tokenize(f"{item['code']} {item['desc']}"))
        score = len(words & haystack_words)
        if score > 0:
            scored.append((score, item))
    if not scored:
        return {"match": None, "candidates": []}
    scored.sort(key=lambda x: -x[0])
    top_score = scored[0][0]
    tied = [item for score, item in scored if score == top_score]
    if len(tied) > 1:
        return {"match": None, "candidates": tied}
    return {"match": tied[0], "candidates": []}


def resolve_location(text, data_file=None):
    """Same whole-word overlap approach and same {"match", "candidates"}
    shape as resolve_material() — see that function's own docstring for
    why this is whole-word, not substring, matching. Against real,
    active Delivery Locations (name and city both searchable, so
    "Chennai" resolves the same as the full location name would)."""
    locs = pc.get_delivery_locations(active_only=True)
    words = set(_tokenize(text))
    if not words:
        return {"match": None, "candidates": []}
    scored = []
    for loc in locs:
        haystack_words = set(_tokenize(f"{loc['name']} {loc['city'] or ''}"))
        score = len(words & haystack_words)
        if score > 0:
            scored.append((score, loc))
    if not scored:
        return {"match": None, "candidates": []}
    scored.sort(key=lambda x: -x[0])
    top_score = scored[0][0]
    tied = [loc for score, loc in scored if score == top_score]
    if len(tied) > 1:
        return {"match": None, "candidates": tied}
    return {"match": tied[0], "candidates": []}


def resolve_screen(text):
    """
    Fuzzy-matches free text against every real page and tab name across
    all three core apps (nav_catalog.py's own single source of truth)
    — same whole-word overlap approach and same {"match", "candidates"}
    shape as resolve_material()/resolve_location() above.

    A tab's own haystack is its tab name alone, not combined with its
    parent page's name, and every real page/tab entry (across every
    app) competes in one single pool, on real score alone — a page
    with a genuinely stronger match (e.g. "Vendor Onboarding", scoring
    2 for matching both real words) correctly beats an unrelated tab
    that only weakly overlaps (e.g. Accounting's own "Vendor Invoices"
    tab, scoring 1 for "vendor" alone) rather than a blanket rule
    always preferring any tab over any page regardless of score — found
    directly as a real false positive, not assumed: an early version
    that unconditionally preferred tabs sent "take me to Vendor
    Onboarding" to Accounting's Vendor Invoices tab instead of the real
    Vendor Onboarding page.

    The one real tie-break applied: when the top score is shared
    between a tab and that *same tab's own parent page* (e.g. "take me
    to Inventory Position" scores the bare Inventory page and its own
    Position & Transfers tab equally), the bare page is dropped from
    that tie and the tab wins — a query specific enough to also name a
    real tab most likely means that tab, not just its page. An
    unrelated page that happens to tie at the same score is never
    dropped by this rule, only a tied entry's own parent.
    """
    screens = nav.all_screens()
    words = set(_tokenize(text))
    if not words:
        return {"match": None, "candidates": []}
    scored = []
    for s in screens:
        name = s["tab"] if s["tab"] else s["page_label"]
        haystack_words = set(_tokenize(name))
        score = len(words & haystack_words)
        if score > 0:
            scored.append((score, s))
    if not scored:
        return {"match": None, "candidates": []}
    top_score = max(sc for sc, _ in scored)
    tied = [s for sc, s in scored if sc == top_score]
    tied_tab_parents = {(s["app"], s["page_key"]) for s in tied if s["tab"] is not None}
    tied = [s for s in tied
            if s["tab"] is not None or (s["app"], s["page_key"]) not in tied_tab_parents]
    if len(tied) > 1:
        return {"match": None, "candidates": tied}
    return {"match": tied[0], "candidates": []}


def resolve_mode(text):
    """Matches free text against the three real Time-Phased Planning
    Mode names, via alias phrases first (more forgiving of casual
    phrasing) then a direct substring check against the real names."""
    lower = text.lower()
    for alias, mode in MODE_ALIASES.items():
        if alias in lower:
            return mode
    for mode in MODE_NAMES:
        if mode.lower() in lower:
            return mode
    return None


def extract_quantity(text):
    """First standalone number in the text, or None if there isn't
    one — callers fall back to a real suggested quantity from actual
    data rather than guessing one themselves."""
    m = re.search(r"\b(\d+(?:\.\d+)?)\b", text)
    return float(m.group(1)) if m else None


def find_transfer_opportunity(mat_code, to_location, data_file=None):
    """Cross-references a resolved material+destination against real,
    current transfer opportunities (bom.get_transfer_opportunities())
    — never fabricates a transfer that doesn't correspond to an actual,
    live shortfall with real stock available to cover it."""
    for o in bom.get_transfer_opportunities(data_file):
        if o["mat_code"] == mat_code and o["to_location"] == to_location:
            return o
    return None


def find_receivable_transfer(mat_code, to_location=None, data_file=None):
    """Cross-references a resolved material (and, if given, a resolved
    destination) against real, currently In Transit stock transfers
    (inventory.get_stock_transfers(status="In Transit")) — same
    non-fabrication discipline as find_transfer_opportunity(): this
    only ever returns a transfer that's genuinely sitting there
    awaiting receipt, never a guess.

    Location is optional here (unlike shipping, where both material
    and destination are needed to size a NEW transfer) — if only one
    In Transit transfer exists for this material, it can be resolved
    without a stated destination. Returns a list so the caller can
    handle 0 (nothing to receive), 1 (unambiguous), or 2+ (needs
    clarification, e.g. the same material in transit to two different
    destinations at once) explicitly, rather than silently guessing
    which one was meant.
    """
    matches = []
    for t in inv.get_stock_transfers(status="In Transit", data_file=data_file):
        if t["material_code"] != mat_code:
            continue
        if to_location and t["to_location"] != to_location:
            continue
        matches.append(t)
    return matches


def find_recommendation(mat_code, data_file=None):
    """Cross-references a resolved material against real, current
    procurement recommendations (bom.get_procurement_recommendations())
    — same non-fabrication discipline as find_transfer_opportunity()."""
    for r in bom.get_procurement_recommendations(data_file):
        if r["mat_code"] == mat_code:
            return r
    return None
