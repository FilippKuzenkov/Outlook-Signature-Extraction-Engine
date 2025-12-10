"""
Phase 2: read cached core_lines and extract per-sender signatures.

Usage:
    - Ensure Phase 1 has run and produced cache/corelines_*.jsonl files.
    - Edit ACTIVE_FOLDER_INDICES and YEAR below (indices match config/folders.txt).
    - Optionally put additional senders to ignore in config/ignored_senders.csv.
    - Run:
        python phase2_signatures_from_cache.py

Result:
    For each selected folder root,
    this script will:
      - read cache
      - run NLP + rules on core_lines per email
      - aggregate best result per sender
      - apply an ignore list of senders
      - write output

The CSV contains (as defined by exporter.py):
    sender_email
    sender_name
    detected_name
    position
    department
    debug_core_lines
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Set
import csv

from cache_io import (
    iter_corelines_cache,
    make_cache_path_for_folder,
    slugify_folder_root,
)
from signature_pipeline import build_result_from_cache_record
from aggregation import SenderAggregator
from exporter import export_contacts

# CONFIG KNOBS (EDIT HERE)

# Year for sanity/logging (not strictly required, but useful for checks)
YEAR = 2025

# Index/indices into config/folders.txt (0-based)
# Must be consistent with Phase 1 (same folder roots, same indices).
ACTIVE_FOLDER_INDICES: List[int] = [12]

# Paths
BASE_DIR = Path(__file__).resolve().parent
CONFIG_DIR = BASE_DIR / "config"
FOLDERS_FILE = CONFIG_DIR / "folders.txt"
IGNORED_SENDERS_CSV = CONFIG_DIR / "ignored_senders.csv"

OUTPUT_DIR = BASE_DIR / "output"

# INTERNAL HELPERS
def _load_folder_roots(config_path: Path) -> List[str]:
    """
    Read Outlook folder roots from config/folders.txt.

    Same logic as in Phase 1 (duplicated here for independence).
    """
    if not config_path.exists():
        raise FileNotFoundError(
            f"Folder config file not found: {config_path!s}"
        )

    roots: List[str] = []
    with config_path.open("r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line:
                continue
            if line.startswith("#"):
                continue
            roots.append(line)

    if not roots:
        raise ValueError(f"No usable folder roots found in {config_path!s}")

    return roots


def _print_available_folders(roots: List[str]) -> None:
    print("\n[Phase 2] Available folder roots (from config/folders.txt):")
    for idx, root in enumerate(roots):
        print(f"  [{idx}] {root}")
    print()


def _make_output_path_for_folder(folder_root: str) -> Path:
    """
    Build the output CSV path for a given folder root
    """
    slug = slugify_folder_root(folder_root)
    filename = f"contacts_{slug}.csv"
    return OUTPUT_DIR / filename


def _load_ignored_senders_from_csv(path: Path) -> Set[str]:
    """
    Load sender email addresses to ignore from a CSV file.

    **Very simple, "one email per line" semantics:**
      - reads only the FIRST column
      - trims whitespace
      - ignores empty lines
      - ignores a header line named 'sender_email' (case-insensitive)
    """
    ignore: Set[str] = set()

    if not path.exists():
        return ignore

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        for row in reader:
            if not row:
                continue
            value = row[0].strip()
            if not value:
                continue

            # Optional header line
            if value.lower() == "sender_email":
                continue

            ignore.add(value)

    return ignore


def _process_single_folder_root(folder_root: str, ignored_senders: Set[str]) -> None:
    """
    Process a single folder root by:
      - locating the corresponding cache file,
      - reading all per-email records,
      - running NLP + rules to build EmailSignatureResult objects,
      - aggregating per sender,
      - filtering ignored senders (hard filter based on sender_email),
      - exporting a CSV.
    """
    cache_path = make_cache_path_for_folder(folder_root)
    if not cache_path.exists():
        print(
            f"\n[Phase 2] Skipping folder '{folder_root}' – "
            f"cache file does not exist:\n  {cache_path}"
        )
        return

    print(f"\n[Phase 2] Processing folder root  : {folder_root}")
    print(f"[Phase 2] Cache file              : {cache_path}")

    aggregator = SenderAggregator(year=YEAR)

    total_records = 0

    for record in iter_corelines_cache(cache_path):
        total_records += 1
        result = build_result_from_cache_record(record)
        aggregator.add_result(result)

        if total_records % 500 == 0:
            print(f"[Phase 2]   Processed {total_records} cached emails...")

    rows = aggregator.to_export_rows()
    pre_filter_count = len(rows)
    """
     FINAL HARD FILTER:
     Re-check every aggregated row against ignored_senders.csv.
     Semantics:
       - Look at row["sender_email"] exactly as it will be written.
       - Strip whitespace.
       - If that string is found in ignored_senders, drop the row.
    """
    if ignored_senders:
        filtered_rows = []
        dropped = 0

        # Normalize ignore entries once (strip only, keep case-sensitive value)
        ignore_set = {s.strip() for s in ignored_senders if s.strip()}

        for row in rows:
            raw_sender = str(row.get("sender_email") or "").strip()
            if raw_sender and raw_sender in ignore_set:
                dropped += 1
                continue
            filtered_rows.append(row)

        rows = filtered_rows

        print(
            f"[Phase 2] Ignore list removed {dropped} sender(s) "
            f"based on exact sender_email match."
        )

    unique_senders = len(rows)

    print(
        f"[Phase 2] Done cache '{cache_path.name}'. "
        f"Cached emails: {total_records}, "
        f"aggregated senders (before ignore): {pre_filter_count}, "
        f"after ignore filter: {unique_senders}"
    )

    if not rows:
        print("[Phase 2] No rows collected after filtering. Nothing to export for this folder.")
    else:
        output_path = _make_output_path_for_folder(folder_root)
        full_path = export_contacts(rows, output_path=output_path)
        print(f"[Phase 2] Exported contacts to:\n  {full_path}")

# MAIN

def main() -> None:
    roots = _load_folder_roots(FOLDERS_FILE)
    _print_available_folders(roots)

    # Basic validation
    for idx in ACTIVE_FOLDER_INDICES:
        if idx < 0 or idx >= len(roots):
            raise IndexError(
                f"ACTIVE_FOLDER_INDICES contains {idx}, but config/folders.txt "
                f"only has {len(roots)} usable entries (0..{len(roots) - 1})."
            )

    ignored_senders = _load_ignored_senders_from_csv(IGNORED_SENDERS_CSV)

    print("[Phase 2] Settings:")
    print(f"  YEAR                  : {YEAR}")
    print(f"  ACTIVE_FOLDER_INDICES : {ACTIVE_FOLDER_INDICES}")
    print(f"  FOLDERS_FILE          : {FOLDERS_FILE}")
    print(f"  IGNORED_SENDERS_CSV   : {IGNORED_SENDERS_CSV} ({len(ignored_senders)} addresses loaded)")
    print(f"  OUTPUT_DIR            : {OUTPUT_DIR}")
    print()

    # Ensure output directory exists
    OUTPUT_DIR.mkdir(exist_ok=True, parents=True)

    for idx in ACTIVE_FOLDER_INDICES:
        folder_root = roots[idx]
        _process_single_folder_root(folder_root, ignored_senders=ignored_senders)

    print("\n[Phase 2] All done.")


if __name__ == "__main__":
    main()
