"""
Per-sender aggregation logic for email signature extraction.

This module is pure Python (no Outlook, no spaCy), and can be reused in
different phases. Phase 2 will typically:

    - read per-email records (from cache or directly from Outlook),
    - run NLP + rules to obtain detected_name / position / department,
    - feed those into SenderAggregator,
    - then export the aggregated rows to CSV.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional


@dataclass
class EmailSignatureResult:
    """
    Represents the extraction result for a single email.
    Fields are deliberately simple dict-like types so that we can easily
    JSON-serialize or convert to rows later.
    """

    sender_email: str
    sender_name: str
    received_time: Optional[datetime]

    detected_name: Optional[str] = None
    position: Optional[str] = None
    department: Optional[str] = None

    core_lines: List[str] = field(default_factory=list)

    entry_id: str = ""
    subject: str = ""
    folder_path: str = ""

    score: Optional[int] = None  # computed if None

    def ensure_score(self) -> int:
        """
        Compute a simple completeness-based score if not already set.

        Current scoring:
          - +1 if detected_name is non-empty
          - +1 if position is non-empty
          - +1 if department is non-empty

        So the score is in the range [0, 3].
        """
        if self.score is not None:
            return self.score

        score = 0
        if self.detected_name:
            score += 1
        if self.position:
            score += 1
        if self.department:
            score += 1

        self.score = score
        return score


@dataclass
class SenderAggregate:
    """
    Aggregated information for a sender across multiple emails.

    What is being saved:
      - the best-scoring result,
      - debug info from that best email.
    """

    sender_email: str
    sender_name: str = ""
    detected_name: Optional[str] = None
    position: Optional[str] = None
    department: Optional[str] = None
    received_time: Optional[datetime] = None
    score: int = 0

    debug_core_lines: str = ""


class SenderAggregator:
    """
    Aggregate multiple EmailSignatureResult objects per sender.
    Strategy:
      - group by sender_email (case-insensitive),
      - for each new email:
          * compute its score (if not pre-computed),
          * if its score is higher than the existing aggregate → replace,
          * if score is equal:
                - if both have received_time, keep the newer,
                - otherwise keep existing.

    At the end, to_export_rows() yields dicts that can be fed into exporter.export_contacts().
    """

    def __init__(self, year: Optional[int] = None) -> None:
        self._by_sender: Dict[str, SenderAggregate] = {}
        self._year = year

    # Public API
    def add_result(self, result: EmailSignatureResult) -> None:
        """
        Add a single EmailSignatureResult into the aggregation.
        """
        email_lower = (result.sender_email or "").lower()
        if not email_lower:
            return

        score = result.ensure_score()

        existing = self._by_sender.get(email_lower)
        if existing is None:
            self._by_sender[email_lower] = SenderAggregate(
                sender_email=result.sender_email,
                sender_name=result.sender_name or "",
                detected_name=result.detected_name,
                position=result.position,
                department=result.department,
                received_time=result.received_time,
                score=score,
                debug_core_lines=" | ".join(result.core_lines) if result.core_lines else "",
            )
            return

        # Decide whether this result should replace the existing aggregate.
        if self._should_replace(existing, score, result.received_time):
            existing.sender_email = result.sender_email
            if result.sender_name:
                existing.sender_name = result.sender_name
            existing.detected_name = result.detected_name
            existing.position = result.position
            existing.department = result.department
            existing.received_time = result.received_time
            existing.score = score
            existing.debug_core_lines = (
                " | ".join(result.core_lines) if result.core_lines else ""
            )

    def to_export_rows(self) -> List[Dict[str, object]]:
        """
        Convert all aggregates into row dicts suitable for exporter.export_contacts().
        The export schema is kept lean:
            sender_email
            sender_name
            detected_name
            position
            department
            debug_core_lines
        """
        rows: List[Dict[str, object]] = []

        for agg in self._by_sender.values():
            if agg.score <= 0:
                detected_name = "no_signature"
                position = "no_signature"
                department = "no_signature"
            else:
                detected_name = agg.detected_name or ""
                position = agg.position or ""
                department = agg.department or ""

            rows.append(
                {
                    "sender_email": agg.sender_email,
                    "sender_name": agg.sender_name,
                    "detected_name": detected_name,
                    "position": position,
                    "department": department,
                    "debug_core_lines": agg.debug_core_lines,
                }
            )

        return rows

    # Internal helpers
    @staticmethod
    def _should_replace(
        existing: SenderAggregate,
        cand_score: int,
        cand_time: Optional[datetime],
    ) -> bool:
        """
        Decide if a new candidate email should replace the existing aggregate.
        Rules:
          - if candidate score > existing score → replace
          - if scores equal:
                - if both have datetime, newer wins
                - else keep existing
        """
        ex_score = existing.score
        if cand_score > ex_score:
            return True
        if cand_score < ex_score:
            return False

        if isinstance(existing.received_time, datetime) and isinstance(cand_time, datetime):
            return cand_time > existing.received_time

        return False
