"""
Signature extraction pipeline for Phase 2 (cache-based).

Phase 2 reads per-email records from the core_lines cache (JSONL),
and for each record uses this module to obtain an EmailSignatureResult,
which is then fed into aggregation.SenderAggregator.

The idea is:
    cache record  →  EmailSignatureResult  →  SenderAggregator  →  export
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, List, Optional

from aggregation import EmailSignatureResult
from nlp_extractor import extract_name_cached
from rules import detect_position, detect_department


def _parse_received_time(value: Any) -> Optional[datetime]:
    """
    Parse the 'received_time' field from cache into a datetime, if possible.

    I expect an ISO 8601 string (e.g. '2025-11-01T10:30:00'), but accept
    a few relaxed formats. If parsing fails, returns None.
    """
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    if not isinstance(value, str):
        return None

    value = value.strip()
    if not value:
        return None

    # Try ISO 8601 first
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            continue

    # Very last fallback: just the date
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except Exception:
        return None


def _ensure_core_lines(value: Any) -> List[str]:
    """
    Normalize core_lines from cache into a list of strings.
    """
    if value is None:
        return []

    if isinstance(value, list):
        # Coerce each element to str for safety
        return [str(x) for x in value if str(x).strip()]

    # If it's a single string with separators, split on common delimiters
    if isinstance(value, str):
        # Try to split on pipes (our typical debug joiner) or newlines
        if "|" in value:
            parts = [p.strip() for p in value.split("|")]
        else:
            parts = [p.strip() for p in value.splitlines()]

        return [p for p in parts if p]

    return []


def build_result_from_cache_record(record: Dict[str, Any]) -> EmailSignatureResult:
    """
    Convert a cache record (JSON dict) into an EmailSignatureResult, including
    running NLP-based name extraction and rules-based position/department.

    Expected keys in `record` (Phase 1 should provide these):

        sender_email: str
        sender_name: str
        received_time: str (ISO) or datetime
        entry_id: str
        subject: str
        folder_path: str
        core_lines: list[str] or similar

    Unknown extra keys are ignored.
    """
    sender_email = str(record.get("sender_email") or "").strip()
    sender_name = str(record.get("sender_name") or "").strip()

    received_time_raw = record.get("received_time")
    received_time = _parse_received_time(received_time_raw)

    entry_id = str(record.get("entry_id") or "")
    subject = str(record.get("subject") or "")
    folder_path = str(record.get("folder_path") or "")

    core_lines = _ensure_core_lines(record.get("core_lines"))

    # NLP + rules
    detected_name = extract_name_cached(sender_email, core_lines)
    position = detect_position(core_lines)
    department = detect_department(core_lines)

    # Build result; score will be computed lazily by EmailSignatureResult.ensure_score()
    return EmailSignatureResult(
        sender_email=sender_email,
        sender_name=sender_name,
        received_time=received_time,
        detected_name=detected_name,
        position=position,
        department=department,
        core_lines=core_lines,
        entry_id=entry_id,
        subject=subject,
        folder_path=folder_path,
    )
