"""
Final CSV writer.
Takes aggregated rows (dicts) from aggregation.SenderAggregator
and writes a semicolon-separated CSV with UTF-8 BOM.
Columns:
    sender_email
    sender_name
    detected_name
    position
    department
    debug_core_lines
"""

from __future__ import annotations
import csv
from pathlib import Path
from typing import List, Dict


def export_contacts(rows: List[Dict], output_path: str | Path) -> str:
    output_path = Path(output_path)
    output_path.parent.mkdir(exist_ok=True, parents=True)

    fieldnames = [
        "sender_email",
        "sender_name",
        "detected_name",
        "position",
        "department",
        "debug_core_lines",
    ]

    with output_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    return str(output_path)
