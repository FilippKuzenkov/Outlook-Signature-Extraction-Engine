"""
Helpers for iterating MailItems in a specific Outlook folder.

External dependency (install via pip):
pywin32    # provides win32com.client for Outlook automation

This module focuses on *reading* from a given folder:
    - Sorting by ReceivedTime (newest → oldest)
    - Applying a simple date filter (>= since)
    - Skipping non-mail items
    - Skipping ignored senders (e.g., your own address or suppliers)
    - Skipping "notification-like" senders based on patterns in
      config/notification_patterns.txt

It deliberately does NOT know about signature extraction or caching; it just
yields MailItem COM objects.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Generator, Iterable, List, Optional, Set, Tuple
import re  # NEW: for email normalization

# For type hints only; this module does not import win32com directly,
# it operates on folder/items objects given from outlook_client.
try:  # pragma: no cover - for type checking only
    from win32com.client import CDispatch  # type: ignore
except Exception:  # pragma: no cover - import may fail in non-Windows env
    CDispatch = object  # fall back type


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_DIR = BASE_DIR / "config"
DEFAULT_NOTIFICATION_PATTERNS_FILE = DEFAULT_CONFIG_DIR / "notification_patterns.txt"

# Email normalization

_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


def _normalize_email_identifier(value: str) -> str:
    """
    Normalize any sender identifier (Outlook SenderEmailAddress or CSV entry)
    to a canonical email-like string:

        - lowercase
        - strip whitespace
        - if any email pattern is found → return first match
        - otherwise return the lowercased trimmed value

    This allows us to match:
        'SMTP:supplier@foo.de'
        'Supplier GmbH <supplier@foo.de>'
        'supplier@foo.de'
    all as 'supplier@foo.de'.
    """
    if not value:
        return ""
    raw = str(value).strip().lower()
    m = _EMAIL_RE.search(raw)
    if m:
        return m.group(0)
    return raw


# Notification-like sender classification
def _load_notification_patterns(path: Path) -> Tuple[List[str], List[str]]:
    """
    Load notification-like patterns from a text file.

    File format:

        # Comments start with '#'
        [local_part]
        no-reply
        do-not-reply
        noreply

        [domain_part]
        github.com
        amazon.de

    Both sections are optional. If the file is missing or empty, this simply
    returns two empty lists.
    """
    local_patterns: List[str] = []
    domain_patterns: List[str] = []

    if not path.exists():
        return local_patterns, domain_patterns

    current_section = "local"
    with path.open("r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue

            lower = line.lower()
            if lower == "[local_part]":
                current_section = "local"
                continue
            if lower == "[domain_part]":
                current_section = "domain"
                continue

            if current_section == "local":
                local_patterns.append(lower)
            else:
                domain_patterns.append(lower)

    return local_patterns, domain_patterns


def _is_notification_like(
    sender_email: str,
    local_patterns: Iterable[str],
    domain_patterns: Iterable[str],
) -> bool:
    """
    Apply simple pattern matching to classify notification-like senders.

    - local_patterns are matched against the local-part (before '@').
    - domain_patterns are matched against the domain part (after '@').
    """
    if not sender_email:
        return False

    email_lower = sender_email.lower()
    if "@" in email_lower:
        local_part, domain_part = email_lower.split("@", 1)
    else:
        local_part, domain_part = email_lower, ""

    for pat in local_patterns:
        if pat and pat in local_part:
            return True

    for pat in domain_patterns:
        if pat and pat in domain_part:
            return True

    return False


def _normalize_for_compare(
    dt1: Optional[datetime],
    dt2: Optional[datetime],
) -> Tuple[Optional[datetime], Optional[datetime]]:
    """
    Normalize two datetimes so they can be compared without raising
    'can't compare offset-naive and offset-aware datetimes'.

    Strategy:
      - If both are None or not datetime → return as-is.
      - If one is aware and the other is naive → strip tzinfo from the aware one.
      - If both are aware or both naive → return unchanged.
    """
    if not isinstance(dt1, datetime) or not isinstance(dt2, datetime):
        return dt1, dt2

    aware1 = dt1.tzinfo is not None
    aware2 = dt2.tzinfo is not None

    if aware1 != aware2:
        # Make both naive for comparison
        if aware1:
            dt1 = dt1.replace(tzinfo=None)
        if aware2:
            dt2 = dt2.replace(tzinfo=None)

    return dt1, dt2


# Iteration helper
def iter_mail_items_in_folder(
    folder: "CDispatch",
    since: Optional[datetime] = None,
    ignore_senders: Optional[Set[str]] = None,
    notification_patterns_file: Optional[Path] = None,
) -> Generator["CDispatch", None, None]:
    """
    Iterate over MailItems in the given Outlook folder.

    This yields MailItem objects (Class == 43), sorted by ReceivedTime
    descending, optionally restricted to mails received after `since`
    and filtered by ignored / notification-like senders.

    Parameters
    ----------
    folder:
        Outlook folder COM object (e.g. returned by OutlookClient.resolve_folder_from_config_path).

    since:
        If given, only mail items with ReceivedTime >= since are considered
        (subject to Outlook's Restrict behavior; if Restrict fails, we fall
        back to manual filtering).

    ignore_senders:
        A set of email addresses (lowercased) that should be skipped entirely,
        e.g. "sales@sales.com".

    notification_patterns_file:
        Path to the notification_patterns.txt file. If None, the default
        path "<repo>/config/notification_patterns.txt" is used.
    """
    ignore_senders = ignore_senders or set()

    # Normalize ignore set to canonical email-like identifiers
    ignore_senders_norm: Set[str] = {
        _normalize_email_identifier(v) for v in ignore_senders if v
    }

    notification_path = (
        notification_patterns_file
        if notification_patterns_file is not None
        else DEFAULT_NOTIFICATION_PATTERNS_FILE
    )

    local_patterns, domain_patterns = _load_notification_patterns(notification_path)

    try:
        items = folder.Items
    except Exception:
        try:
            folder_path = folder.FolderPath
        except Exception:
            folder_path = "<unknown>"
        print(f"[iter_mail_items_in_folder] Cannot read Items for folder: {folder_path}")
        return

    # Sort by ReceivedTime descending
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    # Optional date filter
    restricted_items = items
    if since is not None:
        try:
            filter_str = f"[ReceivedTime] >= '{since.strftime('%Y-%m-%d %H:%M')}'"
            restricted_items = items.Restrict(filter_str)
        except Exception:
            # Fall back to no Restrict
            restricted_items = items

    try:
        item = restricted_items.GetFirst()
    except Exception:
        item = None

    while item is not None:
        try:
            item_class = getattr(item, "Class", None)
        except Exception:
            item_class = None

        if item_class == 43:  # olMailItem
            should_yield = True

            # ReceivedTime filter (in case Restrict failed silently)
            if since is not None:
                try:
                    received_time = getattr(item, "ReceivedTime", None)
                except Exception:
                    received_time = None

                if isinstance(received_time, datetime):
                    rt_cmp, since_cmp = _normalize_for_compare(received_time, since)
                    if isinstance(rt_cmp, datetime) and isinstance(since_cmp, datetime):
                        if rt_cmp < since_cmp:
                            # Because items are sorted newest → oldest, we can break early
                            break

            # Sender-based filters
            try:
                sender_email = getattr(item, "SenderEmailAddress", "") or ""
            except Exception:
                sender_email = ""

            sender_email_lower = sender_email.lower().strip()
            sender_email_norm = _normalize_email_identifier(sender_email_lower)

            # Hard ignore list (e.g. your own address, suppliers, linguists)
            if sender_email_norm and sender_email_norm in ignore_senders_norm:
                should_yield = False

            # Notification-like (no-reply, system, etc.)
            elif _is_notification_like(sender_email_norm, local_patterns, domain_patterns):
                should_yield = False

            if should_yield:
                yield item

        try:
            item = restricted_items.GetNext()
        except Exception:
            break
