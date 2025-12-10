"""
Phase 1: scan Outlook folder(s) and cache per-email signature core_lines.

Usage:
    - Edit ACTIVE_FOLDER_INDICES and YEAR below.
    - Make sure config/folders.txt lists the folder roots you care about.
    - Optional: put additional senders to ignore in config/ignored_senders.csv.
    - Run:
        python phase1_scan_folder.py

Result:
    For each selected folder root, this script will create a JSONL file in
    the cache/ directory

Each line in that file is a JSON object with:

    {
      "sender_email": "...",
      "sender_name": "...",
      "received_time": "2025-11-12T10:34:00",
      "entry_id": "...",
      "subject": "...",
      "folder_path": "...",
      "core_lines": ["Best regards", "Foo Bar", "Projektmanagerin", "Company GmbH"]
    }

Phase 2 will later read these records, run NLP+rules, aggregate per sender,
and export CSV.
"""

from __future__ import annotations

import time
from datetime import datetime
from pathlib import Path
from typing import List, Set
import csv

from outlook_client import OutlookClient, OutlookConfig
from outlook_iterators import iter_mail_items_in_folder

# CLEAN HTML + REPLY REMOVAL (HTML + line-level)
from html_cleaner import html_to_clean_lines, strip_reply_history_lines

# Trim + refine only (strip_reply_history removed!)
from signature_extractor import (
    trim_signature_lines,
    refine_signature_lines,
)

from cache_io import CoreLinesCacheWriter, make_cache_path_for_folder

# CONFIG KNOBS (EDIT HERE)

YEAR = 2025

ACTIVE_FOLDER_INDICES: List[int] = [12]

SHARED_MAILBOX_NAME = "your_mailbox_name"
INBOX_FOLDER_NAME = "your_folder_name"

BASE_IGNORE_SENDER_EMAILS: Set[str] = {"your_ignored_email"}

ENABLE_TIMING: bool = True

BASE_DIR = Path(__file__).resolve().parent
CONFIG_DIR = BASE_DIR / "config"
FOLDERS_FILE = CONFIG_DIR / "folders.txt"
NOTIFICATION_PATTERNS_FILE = CONFIG_DIR / "notification_patterns.txt"
IGNORED_SENDERS_CSV = CONFIG_DIR / "ignored_senders.csv"



# INTERNAL HELPERS
def _load_folder_roots(config_path: Path) -> List[str]:
    if not config_path.exists():
        raise FileNotFoundError(f"Folder config file not found: {config_path!s}")

    roots: List[str] = []
    with config_path.open("r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            roots.append(line)

    if not roots:
        raise ValueError(f"No usable folder roots found in {config_path!s}")

    return roots


def _print_available_folders(roots: List[str]) -> None:
    print("\n[Phase 1] Available folder roots (from config/folders.txt):")
    for idx, root in enumerate(roots):
        print(f"  [{idx}] {root}")
    print()


def _load_ignored_senders_from_csv(path: Path) -> Set[str]:
    ignore: Set[str] = set()

    if not path.exists():
        return ignore

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        first_row = True
        for row in reader:
            if not row:
                continue
            value = row[0].strip()
            if not value:
                continue
            if first_row and ("@" not in value):
                first_row = False
                continue
            first_row = False
            if "@" in value:
                ignore.add(value.lower())

    return ignore


def _process_single_folder(
    client: OutlookClient,
    folder_root: str,
    year: int,
    ignore_senders: Set[str],
) -> None:

    folder = client.resolve_folder_from_config_path(folder_root)

    try:
        folder_path = folder.FolderPath
    except Exception:
        folder_path = "<unknown>"

    print(f"\n[Phase 1] Processing folder root : {folder_root}")
    print(f"[Phase 1] Outlook folder path    : {folder_path}")

    since = datetime(year, 1, 1, 0, 0, 0)

    cache_path = make_cache_path_for_folder(folder_root)
    print(f"[Phase 1] Cache file             : {cache_path}")

    total_mails = 0
    total_records = 0

    folder_start_time = time.time()
    batch_start_time = folder_start_time
    batch_start_count = 0

    with CoreLinesCacheWriter(cache_path) as writer:

        mail_iter = iter_mail_items_in_folder(
            folder=folder,
            since=since,
            ignore_senders=ignore_senders,
            notification_patterns_file=NOTIFICATION_PATTERNS_FILE,
        )

        for mail in mail_iter:
            total_mails += 1

            # Timing output
            if ENABLE_TIMING and total_mails % 100 == 0:
                now = time.time()
                batch_time = now - batch_start_time
                batch_count = total_mails - batch_start_count
                avg = batch_time / batch_count if batch_count else 0.0
                elapsed = now - folder_start_time
                print(
                    f"[Phase 1]   Processed {total_mails} mails "
                    f"(last {batch_count} took {batch_time:.2f}s, "
                    f"{avg:.3f}s/mail, total {elapsed:.2f}s)",
                    flush=True,
                )
                batch_start_time = now
                batch_start_count = total_mails
            elif not ENABLE_TIMING and total_mails % 100 == 0:
                print(f"[Phase 1]   Processed {total_mails} mails...", flush=True)

            # Extract Outlook fields safely
            try:
                received_time = getattr(mail, "ReceivedTime", None)
            except Exception:
                received_time = None

            if isinstance(received_time, datetime) and received_time.year != year:
                continue

            try:
                sender_email = getattr(mail, "SenderEmailAddress", "") or ""
            except Exception:
                sender_email = ""

            try:
                sender_name = getattr(mail, "SenderName", "") or ""
            except Exception:
                sender_name = ""

            try:
                entry_id = getattr(mail, "EntryID", "") or ""
            except Exception:
                entry_id = ""

            try:
                subject = getattr(mail, "Subject", "") or ""
            except Exception:
                subject = ""

            try:
                mail_folder_path = mail.Parent.FolderPath
            except Exception:
                mail_folder_path = folder_path

            try:
                html_body = getattr(mail, "HTMLBody", "") or ""
            except Exception:
                html_body = ""

            # SIGNATURE PIPELINE — FIXED VERSION

            # 1) Convert HTML → clean lines (INCLUDES HTML-level reply stripping)
            all_lines = html_to_clean_lines(html_body)

            # 2) Additional line-level reply stripping
            current_lines = strip_reply_history_lines(all_lines)

            # 3) Existing line trimming / refining
            sig_lines = trim_signature_lines(current_lines)
            core_lines = refine_signature_lines(sig_lines)

            rec = {
                "sender_email": sender_email,
                "sender_name": sender_name,
                "received_time": (
                    received_time.isoformat()
                    if isinstance(received_time, datetime)
                    else None
                ),
                "entry_id": entry_id,
                "subject": subject,
                "folder_path": mail_folder_path,
                "core_lines": core_lines,
            }

            writer.write_record(rec)
            total_records += 1

    folder_end = time.time()
    total_elapsed = folder_end - folder_start_time

    print(
        f"[Phase 1] Done folder '{folder_root}'. "
        f"Scanned mails: {total_mails}, cached records: {total_records}"
    )

    if ENABLE_TIMING:
        avg = total_elapsed / total_mails if total_mails > 0 else 0.0
        print(
            f"[Phase 1] Timing for folder '{folder_root}': "
            f"{total_elapsed:.2f}s total, {avg:.3f}s/mail on average",
            flush=True,
        )

# MAIN
def main() -> None:
    roots = _load_folder_roots(FOLDERS_FILE)
    _print_available_folders(roots)

    for idx in ACTIVE_FOLDER_INDICES:
        if idx < 0 or idx >= len(roots):
            raise IndexError(
                f"ACTIVE_FOLDER_INDICES contains {idx}, but only {len(roots)} entries available."
            )

    extra_ignored = _load_ignored_senders_from_csv(IGNORED_SENDERS_CSV)
    all_ignored = set(s.lower() for s in BASE_IGNORE_SENDER_EMAILS) | extra_ignored

    print("[Phase 1] Settings:")
    print(f"  YEAR                  : {YEAR}")
    print(f"  ACTIVE_FOLDER_INDICES : {ACTIVE_FOLDER_INDICES}")
    print(f"  SHARED_MAILBOX_NAME   : {SHARED_MAILBOX_NAME}")
    print(f"  INBOX_FOLDER_NAME     : {INBOX_FOLDER_NAME}")
    print(f"  Base IGNORE senders   : {', '.join(sorted(BASE_IGNORE_SENDER_EMAILS))}")
    print(f"  Extra IGNORE CSV      : {IGNORED_SENDERS_CSV} ({len(extra_ignored)} loaded)")
    print(f"  Total IGNORE senders  : {len(all_ignored)}")
    print(f"  FOLDERS_FILE          : {FOLDERS_FILE}")
    print(f"  NOTIFICATION_PATTERNS : {NOTIFICATION_PATTERNS_FILE}")
    print(f"  ENABLE_TIMING         : {ENABLE_TIMING}")
    print()

    config = OutlookConfig(
        mailbox_name=SHARED_MAILBOX_NAME,
        inbox_folder_name=INBOX_FOLDER_NAME,
    )
    client = OutlookClient(config=config)

    overall_start = time.time()

    for idx in ACTIVE_FOLDER_INDICES:
        folder_root = roots[idx]
        _process_single_folder(
            client=client,
            folder_root=folder_root,
            year=YEAR,
            ignore_senders=all_ignored,
        )

    overall_end = time.time()

    if ENABLE_TIMING:
        print(f"\n[Phase 1] All folders done in {overall_end - overall_start:.2f}s.", flush=True)
    else:
        print("\n[Phase 1] All done.")


if __name__ == "__main__":
    main()
