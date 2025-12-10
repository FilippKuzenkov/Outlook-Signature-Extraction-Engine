"""
Helpers for reading/writing per-email core_lines cache files (JSONL).

Phase 1 will:
    - iterate all MailItems in a folder,
    - compute signature core_lines per email,
    - write one JSON object per email to a cache file.

Phase 2 will:
    - read the JSONL file,
    - run NLP + rules + aggregation on those records.

File naming convention:

    cache/corelines_<slugified_folder_root>.jsonl

where the slug is derived from the folder root path from config/folders.txt
"""

from __future__ import annotations

import json
import os
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, Mapping, Union


BASE_DIR = Path(__file__).resolve().parent
CACHE_DIR = BASE_DIR / "cache"
CACHE_DIR.mkdir(exist_ok=True)


def slugify_folder_root(folder_root: str) -> str:
    """
    Convert a folder root into a safe slug.

    Common German umlauts are normalized for portability.
    """
    slug = folder_root.replace("\\", "_").replace("/", "_")
    slug = slug.replace(" ", "")

    # German umlauts
    slug = (
        slug.replace("ä", "ae").replace("Ä", "Ae")
            .replace("ö", "oe").replace("Ö", "Oe")
            .replace("ü", "ue").replace("Ü", "Ue")
            .replace("ß", "ss")
    )
    return slug


def make_cache_path_for_folder(folder_root: str) -> Path:
    """
    Builds the cache file path for a given folder root
    """
    slug = slugify_folder_root(folder_root)
    filename = f"corelines_{slug}.jsonl"
    return CACHE_DIR / filename


class CoreLinesCacheWriter:
    """
    Simple streaming writer for core_lines cache files (JSONL).

    Usage:

        from cache_io import CoreLinesCacheWriter, make_cache_path_for_folder

        path = make_cache_path_for_folder(folder_root)
        with CoreLinesCacheWriter(path) as writer:
            writer.write_record({
                "sender_email": "...",
                "sender_name": "...",
                "received_time": "2025-11-01T10:30:00",
                "entry_id": "...",
                "subject": "...",
                "folder_path": "...",
                "core_lines": ["Best regards", "Foo Bar", "Projektmanagerin"],
            })
    """

    def __init__(self, path: Union[str, Path]) -> None:
        self._path = Path(path)
        self._fh = None

    @property
    def path(self) -> Path:
        return self._path

    def __enter__(self) -> "CoreLinesCacheWriter":
        # Ensure directory
        if self._path.parent and not self._path.parent.exists():
            os.makedirs(self._path.parent, exist_ok=True)

        # Open in text mode, overwrite existing
        self._fh = self._path.open("w", encoding="utf-8")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._fh is not None:
            self._fh.close()
            self._fh = None

    def write_record(self, record: Union[Mapping[str, Any], Any]) -> None:
        """
        Writes a single record as one JSON line.
        If `record` is a dataclass, it is converted via asdict().
        """
        if self._fh is None:
            raise RuntimeError("CoreLinesCacheWriter is not open. Use as a context manager.")

        if is_dataclass(record):
            data = asdict(record)
        elif isinstance(record, Mapping):
            data = dict(record)
        else:
            raise TypeError(f"Unsupported record type for cache writing: {type(record)!r}")

        json_line = json.dumps(data, ensure_ascii=False)
        self._fh.write(json_line + "\n")


def iter_corelines_cache(path: Union[str, Path]) -> Iterator[Dict[str, Any]]:
    """
    Iterates records from a core_lines cache file (JSONL).
    Each line is parsed as JSON and yielded as a dict. Lines that fail
    to parse are skipped with a simple warning.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Core-lines cache file not found: {path!s}")

    with path.open("r", encoding="utf-8") as fh:
        for lineno, raw in enumerate(fh, start=1):
            line = raw.strip()
            if not line:
                continue
            try:
                data = json.loads(line)
            except Exception as exc:
                print(
                    f"[iter_corelines_cache] Warning: failed to parse JSON on "
                    f"{path.name}:{lineno}: {exc}"
                )
                continue

            # Ensure we always return a dict
            if isinstance(data, dict):
                yield data
            else:
                yield {"value": data}
