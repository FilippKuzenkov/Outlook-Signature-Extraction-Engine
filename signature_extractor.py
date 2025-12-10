# signature_extractor.py
"""
Signature extraction helpers.

This module supports BOTH:
- Phase 1: line trimming + refining for core_lines extraction
- Phase 2: full signature extraction from HTML (NLP + rules)

Phase 1:
    phase1_scan_folder.py imports:
        trim_signature_lines()
        refine_signature_lines()

Phase 2 (legacy / direct-from-HTML):
    Some code paths may still use the SignatureExtractor class, which
    takes a raw HTML body and returns a dict with:
        email, name, position, department

The new implementation introduces a slightly smarter Phase-1 trimming
algorithm without changing the public API:

    1) Preferred "anchor window" behaviour:
       - Find the last sign-off line (e.g. "Best regards", "Viele Grüße").
       - From the line *after* that, scan a small window of lines below.
       - Within that window, detect:
            * email address (highest priority), or
            * phone number, or
            * URL / website.
       - If a sign-off and such an "end" line are found, the signature block
         is defined as all lines between them (exclusive of the sign-off
         line, inclusive of the end line).

    2) Fallback to the older heuristic:
       - Start collecting all lines after the sign-off phrase until the end.
       - If no sign-off is found at all, take the bottom N lines.

This gives us a narrower, more deterministic core_lines block whenever
the anchor cues are available, while preserving previous behaviour for
the remaining cases.
"""

from __future__ import annotations

import re
from typing import List, Optional

from html_cleaner import html_to_clean_lines
from nlp_extractor import extract_name_candidates
from rules import extract_position, extract_department

# PHASE 1 – sign-off and contact heuristics

# Sign-off / finishing phrases in EN/DE which typically precede the signature.
# We deliberately keep this list simple and transparent.
_SIGNOFF_PHRASES: tuple[str, ...] = (
    # English
    "best regards",
    "kind regards",
    "regards",
    "thanks and regards",
    "many thanks",
    "with best regards",
    "sincerely",
    "yours sincerely",
    "yours faithfully",
    # German
    "mit freundlichen grüßen",
    "mit freundlichen gruessen",
    "mit freundlichen grussen",
    "freundliche grüße",
    "freundliche grüsse",
    "viele grüße",
    "viele grüsse",
    "beste grüße",
    "beste grüsse",
    "herzliche grüße",
    "herzliche grüsse",
    # very short variants
    "mfg",
    "vhb",
    "br,",  # "best regards" style abbreviations
)

# Basic detectors for email / phone / URL like lines
_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


_PHONE_HINT_WORDS: tuple[str, ...] = (
    "tel",
    "telefon",
    "phone",
    "mobil",
    "mobile",
    "handy",
    "cell",
    "fax",
    "fon",
    "t:",
    "m:",
    "f:",
)


_URL_HINT_SUBSTRINGS: tuple[str, ...] = (
    "http://",
    "https://",
    "www.",
)


def _find_last_signoff_index(lines: List[str]) -> Optional[int]:
    """
    Return the index of the *last* line that contains a sign-off phrase.

    We search for any of the configured _SIGNOFF_PHRASES as substring in the
    lower-cased line. Using the last index is robust against occasional
    older content that slipped through from reply chains.
    """
    last_idx: Optional[int] = None
    for idx, line in enumerate(lines):
        ll = line.lower()
        if any(phrase in ll for phrase in _SIGNOFF_PHRASES):
            last_idx = idx
    return last_idx


def _line_has_email(line: str) -> bool:
    return bool(_EMAIL_RE.search(line))


def _line_has_phone(line: str) -> bool:
    ll = line.lower()
    # Quick hint check to avoid treating random numbers as phones.
    if not any(hint in ll for hint in _PHONE_HINT_WORDS):
        # Still allow generic numeric lines to be phones if they look like numbers.
        digits = sum(ch.isdigit() for ch in line)
        if digits < 6:
            return False
    return True


def _line_has_url(line: str) -> bool:
    ll = line.lower()
    return any(sub in ll for sub in _URL_HINT_SUBSTRINGS)


def _find_contact_end_index(
    lines: List[str],
    start_from: int,
    max_span: int = 12,
) -> Optional[int]:
    """
    Within a window of up to `max_span` lines starting at `start_from`,
    find the end index of the signature block based on contact-like lines.

    Priority:
        1) email address
        2) phone number
        3) URL / website
    """
    email_idx: Optional[int] = None
    phone_idx: Optional[int] = None
    url_idx: Optional[int] = None

    n = len(lines)
    end_limit = min(n, start_from + max_span)

    for idx in range(start_from, end_limit):
        line = lines[idx]
        if not line:
            continue

        if _line_has_email(line):
            email_idx = idx
        elif _line_has_phone(line):
            phone_idx = idx
        elif _line_has_url(line):
            url_idx = idx

    if email_idx is not None:
        return email_idx
    if phone_idx is not None:
        return phone_idx
    if url_idx is not None:
        return url_idx
    return None


# Fallback – approximate reconstruction of the older behaviour

def _old_trim_signature_lines(lines: List[str]) -> List[str]:
    """
    Previous Phase-1 behaviour, kept as a fallback:

        - look for a sign-off phrase,
        - once found, start collecting all *following* lines as signature.

    If no sign-off is found at all, fall back to a simple
    "last N lines" heuristic.
    """
    if not lines:
        return []

    collected: List[str] = []
    started = False

    for line in lines:
        ll = line.lower()

        if any(phrase in ll for phrase in _SIGNOFF_PHRASES):
            started = True
            # we *skip* the sign-off line itself, same as before
            continue

        if started:
            collected.append(line)

    # Fallback: if nothing was collected, take the bottom N lines.
    if not collected:
        n = len(lines)
        if n <= 13:
            return list(lines)
        return lines[-13:]

    return collected


# PHASE 1 – public API

def trim_signature_lines(lines: List[str]) -> List[str]:
    """
    Phase-1 trimming with a smarter, more deterministic algorithm.

    Algorithm:
        1) Try *new* anchor-based behaviour:
            - locate the last sign-off line,
            - from the line *after* that, scan a small window of lines
              below and determine the "end" line by looking for email /
              phone / URL lines,
            - if both indices exist and there is at least one line in
              between, return that slice.

        2) If no sign-off is found or no contact-style end line exists
           in the window, fall back to the older behaviour implemented
           in _old_trim_signature_lines().
    """
    if not lines:
        return []

    # 1) locate sign-off line
    signoff_idx = _find_last_signoff_index(lines)

    if signoff_idx is not None:
        start_from = signoff_idx + 1  # signature starts *after* the greeting
        if start_from < len(lines):
            end_idx = _find_contact_end_index(lines, start_from=start_from)
            if end_idx is not None and end_idx >= start_from:
                candidate = lines[start_from : end_idx + 1]
                if candidate:
                    return candidate

    # 2) fallback
    return _old_trim_signature_lines(lines)


def refine_signature_lines(lines: List[str]) -> List[str]:
    """
    Phase-1 post-processing:

    - remove trivial separators and obvious reply markers,
    - drop typical unsubscribe / disclaimer fragments,
    - normalize whitespace by stripping, but keep the original line text.

    This function is deliberately light-weight; heavier logic such as
    NLP scoring happens in Phase 2.
    """
    refined: List[str] = []

    for line in lines:
        if line is None:
            continue
        stripped = line.strip()
        if not stripped:
            continue

        ll = stripped.lower()

        # Obvious separators / reply markers
        if ll.startswith("-----"):
            continue
        if ll.startswith("from:") or ll.startswith("von:"):
            continue
        if ll.startswith("sent:") or ll.startswith("gesendet:"):
            continue
        if "unsubscribe" in ll:
            continue
        if "this email" in ll or "this e-mail" in ll:
            # typical disclaimer beginnings
            continue
        if "original message" in ll or "ursprüngliche nachricht" in ll:
            continue

        refined.append(stripped)

    return refined


# PHASE 2 – legacy direct-HTML extractor (kept for compatibility)

class SignatureExtractor:
    """
    Simple Phase-2 style helper that extracts a signature directly from
    an HTML body using the NLP and rules modules.

    Modern Phase-2 code should prefer signature_pipeline.py working on
    cached core_lines. This class remains for backwards compatibility.
    """

    def __init__(self, logger) -> None:
        self.logger = logger

    def extract(self, sender_email: str, html_body: str) -> dict:
        """
        Extract name, position, department from a raw HTML email body.

        Returns a dict with keys:
            email, name, position, department
        """
        try:
            lines = html_to_clean_lines(html_body)
            self.logger.debug(f"[{sender_email}] Cleaned signature lines: {lines}")

            name = extract_name_candidates(lines)
            position = extract_position(lines)
            department = extract_department(lines)

            return {
                "email": sender_email,
                "name": name or "",
                "position": position or "",
                "department": department or "",
            }

        except Exception as exc:  # pragma: no cover - defensive
            self.logger.error(f"Signature extraction failed for {sender_email}: {exc}")
            return {
                "email": sender_email,
                "name": "",
                "position": "",
                "department": "",
            }
