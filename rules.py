"""
Keyword-based detection for job titles and departments.
Reads keyword lists from config/
Uses simple substring matching in core_lines, with heuristics
to skip disclaimers, URLs, legal footers and courtesy lines.
"""

from __future__ import annotations
from pathlib import Path
from typing import List


BASE_DIR = Path(__file__).resolve().parent
CONFIG_DIR = BASE_DIR / "config"


def _read_lines(path: Path) -> List[str]:
    if not path.exists():
        return []
    out = []
    with path.open("r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            out.append(line.lower())
    return out


# Load keyword files once
JOB_TITLES = (
    _read_lines(CONFIG_DIR / "job_title_keywords.txt")
    + _read_lines(CONFIG_DIR / "extra_title_tokens.txt")
)

DEPARTMENTS = (
    _read_lines(CONFIG_DIR / "department_keywords.txt")
    + _read_lines(CONFIG_DIR / "extra_department_tokens.txt")
)

# Common disclaimer / footer phrases we want to ignore
_DISCLAIMER_SUBSTRINGS = [
    "please consider the environment",
    "before printing this e-mail",
    "before printing this email",
    "this email and any files",
    "this e-mail and any files",
    "if you are not the intended recipient",
    "unauthorized use",
    "unauthorised use",
    "unauthorized disclosure",
    "confidential",
    "virus",
    "viruses",
    "malware",
    "unsolicited",
    "unsubscribe",
    "this message may contain",
    "datenschutz",
    "privacy",
    "imprint",
    "impressum",
    "handelsregister",
    "registergericht",
    "amtsgericht",
    "agb",
    "a.g.b.",
    "u.st.-id",
    "ust.-id",
    "ust id",
    "ust-id",
    "vat ",
    "vat no.",
    "rcs ",
    "hrb ",
    "sitz der gesellschaft",
    "postfach",
    "steuer",
    "tax id",
    "copyright",
    "newsletter",
    "update subscription preferences",
    "visit our blog",
    "follow us on",
    "linkedin",
    "instagram",
    "facebook",
    "twitter",
]

# Courtesy / generic phrases that are not titles
_NON_TITLE_SUBSTRINGS = [
    "thank you very much",
    "thank you for your",
    "thank you for your cooperation",
    "many thanks",
    "best regards",
    "kind regards",
    "freundliche grüße",
    "freundliche grüsse",
    "mit freundlichen grüßen",
    "mit freundlichen gruessen",
]


def _looks_like_disclaimer(line: str) -> bool:
    ll = line.lower()
    if not ll.strip():
        return True

    # Email addresses / URLs
    if "@" in ll or "http://" in ll or "https://" in ll or "www." in ll:
        return True

    # Separator / original message markers etc.
    if ll.startswith("-----"):
        return True

    # Known disclaimer / footer fragments
    if any(pat in ll for pat in _DISCLAIMER_SUBSTRINGS):
        return True

    # Very long "sentences" are typically disclaimers or free text
    words = ll.split()
    if len(words) > 18:
        return True

    return False


def _looks_like_nontitle(line: str) -> bool:
    ll = line.lower()
    return any(pat in ll for pat in _NON_TITLE_SUBSTRINGS)


def detect_position(lines: List[str]) -> str | None:
    """
    Return first matching job title line, bottom-up.
    We skip lines that look like disclaimers or pure courtesy phrases,
    and we also avoid lines that contain lots of digits.
    """
    for line in reversed(lines):
        if _looks_like_disclaimer(line) or _looks_like_nontitle(line):
            continue

        # Avoid phone-number/zipcode heavy lines as titles
        digits = sum(ch.isdigit() for ch in line)
        if digits >= 4:
            continue

        ll = line.lower()
        for kw in JOB_TITLES:
            if kw in ll:
                return line
    return None


def detect_department(lines: List[str]) -> str | None:
    """
    Return first matching department line, bottom-up.
    We skip lines that look like disclaimers / legal footers / URLs,
    and we avoid numeric-heavy lines (addresses, HRB, etc.).
    """
    for line in reversed(lines):
        if _looks_like_disclaimer(line):
            continue

        digits = sum(ch.isdigit() for ch in line)
        if digits >= 4:
            continue

        ll = line.lower()
        for kw in DEPARTMENTS:
            if kw in ll:
                return line
    return None

# Backwards-compatible aliases for older code (e.g. signature_extractor)

def extract_position(lines: List[str]) -> str | None:
    """
    Backwards-compatible wrapper for legacy code.
    Forwards to detect_position(lines).
    """
    return detect_position(lines)


def extract_department(lines: List[str]) -> str | None:
    """
    Backwards-compatible wrapper for legacy code.
    Forwards to detect_department(lines).
    """
    return detect_department(lines)
