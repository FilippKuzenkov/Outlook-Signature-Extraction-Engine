"""
spaCy-based PERSON name extraction with caching per sender.

Loads en_core_web_sm and de_core_news_sm (best-effort).
Tries both models; whichever gives PERSON entities first wins.

Adds a plausibility filter so we avoid garbage names like
"UNSUBSCRIBE HERE", "Amtsgericht Tostedt", addresses, etc.
"""

from __future__ import annotations

import spacy
from typing import Optional, List, Dict

_nlp_en = None
_nlp_de = None
_name_cache: Dict[str, Optional[str]] = {}

# Known-bad or internal names that should not be used as detected_name.
# Extend this list as you see recurring false positives in the CSV.
_NAME_BLACKLIST_EXACT = {
    "unsubscribe here",
    "unsubscribe",
    "any",
    "personally",
    "client success",
    "project manager",  # title, not name
    "translation en de",
    "hallo frau",
    "vielen dank",
    "uber videogeraet",
    "über videogerät",
    # Internal signature examples (so we don't mis-assign your own name):
    "person a",
    "person b",
}


def _try_load(model: str):
    try:
        return spacy.load(model)
    except Exception:
        return None


def _ensure_models_loaded():
    global _nlp_en, _nlp_de
    if _nlp_en is None:
        _nlp_en = _try_load("en_core_web_sm")
    if _nlp_de is None:
        _nlp_de = _try_load("de_core_news_sm")


def _is_plausible_name(text: str) -> bool:
    """
    Heuristic filter to decide if a detected string looks like a person name.
    We want to avoid things like "UNSUBSCRIBE HERE","London SE13", etc.
    """
    if not text:
        return False

    # Normalize newlines and whitespace
    t = text.replace("\n", " ").strip()
    if not t:
        return False

    lower = t.lower()

    # Reject if explicitly blacklisted
    if lower in _NAME_BLACKLIST_EXACT:
        return False

    # Too short or too long
    if len(t) < 3 or len(t) > 60:
        return False

    # Names rarely contain digits (skip anything with numbers)
    if any(ch.isdigit() for ch in t):
        return False

    # Tokenize
    tokens = [tok for tok in t.split() if tok.strip(",.;:!\"'()[]")]
    if len(tokens) < 2 or len(tokens) > 5:
        # usually expect at least First + Last, at most some middle names
        return False

    # Reject greetings/articles as first token
    first_lower = tokens[0].strip(",.;:!\"'()[]").lower()
    if first_lower in {
        "hi",
        "hallo",
        "liebe",
        "lieber",
        "dear",
        "der",
        "die",
        "das",
        "the",
        "this",
        "diese",
        "dieser",
        "dieses",
    }:
        return False

    # Company/legal/address-ish patterns
    org_suffixes = {
        "gmbh",
        "ag",
        "kg",
        "llc",
        "ltd",
        "inc",
        "sarl",
        "sas",
        "s.a.",
        "s.a",
        "s.p.a",
        "ug",
        "gbr",
        "e.k.",
    }
    bad_substrings = {
        "gericht",          # Amtsgericht, Registergericht, etc.
        "rechte",           # Alle Rechte ...
        "rights",
        "washington",
        "london",
        "registergericht",
        "amtsgericht",
        "consult",
        "consulting",
        "translation",
        "translators",
        "str",              # Str, Str., Virchowstr, etc.
        "straße",
        "strasse",
        "platz",
        "plaza",
        "city",
        "road",
        "avenue",
        "street",
        "gmbh",
    }

    good_tokens = 0

    for tok in tokens:
        tok_clean = tok.strip(",.;:!\"'()[]")
        if not tok_clean:
            continue
        tl = tok_clean.lower()

        # Company/legal suffix?
        if tl in org_suffixes:
            return False

        # Contains obvious non-person markers?
        if any(b in tl for b in bad_substrings):
            return False

        # For counting good tokens: Require capitalized and alphabetic
        if not tok_clean[0].isupper():
            continue

        letters = [ch for ch in tok_clean if ch.isalpha()]
        if not letters:
            continue

        good_tokens += 1

    if good_tokens < 2:
        return False

    # Avoid all-caps "names" like "FOO BAR" (usually non-person content)
    all_caps = all(
        tok.strip(",.;:!\"'()[]").isupper()
        for tok in tokens
        if tok.strip(",.;:!\"'()[]").isalpha()
    )
    if all_caps and len(tokens) > 1:
        return False

    return True


def _postprocess_candidate(name: Optional[str]) -> Optional[str]:
    if not name:
        return None
    candidate = name.strip()
    if not candidate:
        return None
    if not _is_plausible_name(candidate):
        return None
    return candidate


def extract_name_cached(sender_email: str, lines: List[str]) -> Optional[str]:
    key = sender_email.lower()
    if key in _name_cache:
        return _name_cache[key]

    name = extract_name(lines)
    name = _postprocess_candidate(name)
    _name_cache[key] = name
    return name


def extract_name(lines: List[str]) -> Optional[str]:
    _ensure_models_loaded()

    text = "\n".join(lines)
    if not text.strip():
        return None

    # First try German
    if _nlp_de:
        doc = _nlp_de(text)
        for ent in doc.ents:
            if ent.label_ == "PER":
                return ent.text.strip()

    # Then English
    if _nlp_en:
        doc = _nlp_en(text)
        for ent in doc.ents:
            if ent.label_ == "PERSON":
                return ent.text.strip()

    # Fallback: heuristic → first line with 2 capitalized tokens
    for line in lines:
        tokens = line.strip().split()
        caps = [t for t in tokens if t[:1].isupper()]
        if len(caps) >= 2:
            return " ".join(caps)

    return None

# Backwards-compatible alias for older code (e.g. signature_extractor)

def extract_name_candidates(lines: List[str]) -> Optional[str]:
    """
    Backwards-compatible wrapper used by older modules.
    Currently just forwards to extract_name(lines).
    """
    return extract_name(lines)
