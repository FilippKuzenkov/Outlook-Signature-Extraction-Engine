"""
Outlook client helpers for connecting to a (shared) mailbox and resolving folders.
External dependency (install via pip):

pywin32  # provides win32com.client for Outlook automation

This module is intentionally small and generic so it can be reused by both
Phase 1 (slow Outlook scan → cache) and Phase 2 variants if needed.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, List

try:
    import win32com.client as win32com_client
except ImportError as exc:  # pragma: no cover - environment-specific
    win32com_client = None
    _import_error = exc
else:  # pragma: no cover - normal environment
    _import_error = None


@dataclass
class OutlookConfig:
    """
    Basic configuration for connecting to Outlook.

    Attributes
    ----------
    mailbox_name:
        Display name (or substring) of the mailbox to open.
        For a shared mailbox, this should
        typically be set to that string. If None, the default
        mailbox (current Outlook profile) is used.

    inbox_folder_name:
        Name of the Inbox folder in this mailbox, e.g. "Inbox".
    """

    mailbox_name: Optional[str] = None
    inbox_folder_name: str = "your_folder_name"


class OutlookClient:
    """
    Small wrapper around the Outlook COM API.

    Responsibilities:
      - Ensure pywin32 is available.
      - Connect to the Outlook namespace.
      - Locate the target mailbox (default or shared).
      - Resolve an Inbox folder.
      - Resolve "config-style" folder paths like:
            "Inbox"
            "Inbox\\2025" etc
        to a concrete Outlook folder object.

    This class deliberately does NOT implement any higher-level iteration or
    filtering; that is handled in outlook_iterators.py.
    """

    def __init__(self, config: OutlookConfig) -> None:
        if win32com_client is None:
            raise RuntimeError(
                "win32com.client (pywin32) is not available. "
                "Install pywin32 and ensure you are running on Windows "
                "with Outlook Desktop installed."
            ) from _import_error

        self._config = config
        self._outlook = win32com_client.Dispatch("Outlook.Application")
        self._namespace = self._outlook.GetNamespace("MAPI")

        self._mailbox_root = self._get_mailbox_root()
        self._inbox_folder = self._get_inbox_folder()

    # Internal helpers

    def _get_mailbox_root(self):
        """
        Locate the mailbox root folder.

        If config.mailbox_name is None, use the default Inbox's parent store.
        Otherwise, search all Stores for a DisplayName that either equals or
        contains the mailbox_name (case-insensitive).
        """
        if not self._config.mailbox_name:
            # Default store: take the parent of the default Inbox
            inbox = self._namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            return inbox.Parent

        target = self._config.mailbox_name.lower()
        target_store = None

        for store in self._namespace.Stores:
            try:
                name = store.DisplayName
            except Exception:
                continue

            if not name:
                continue

            lower_name = str(name).lower()
            if lower_name == target or target in lower_name:
                target_store = store
                break

        if target_store is None:
            available: List[str] = []
            for store in self._namespace.Stores:
                try:
                    available.append(str(store.DisplayName))
                except Exception:
                    continue

            raise ValueError(
                f"Mailbox '{self._config.mailbox_name}' not found in Outlook profiles.\n"
                f"Available stores:\n  " + "\n  ".join(available)
            )

        return target_store.GetRootFolder()

    def _get_inbox_folder(self):
        """
        Resolve the Inbox folder inside the mailbox root.

        We first try the configured name directly, then a couple of common
        fallbacks ("Inbox").
        """
        candidates: List[str] = []
        if self._config.inbox_folder_name:
            candidates.append(self._config.inbox_folder_name)

        # fallback list, in case the configured name is not present
        for name in ("Inbox"):
            if name not in candidates:
                candidates.append(name)

        for name in candidates:
            try:
                return self._mailbox_root.Folders[name]
            except Exception:
                continue

        available: List[str] = []
        for folder in self._mailbox_root.Folders:
            try:
                available.append(str(folder.Name))
            except Exception:
                continue

        raise ValueError(
            f"Could not find an Inbox folder in mailbox '{self._mailbox_root.Name}'.\n"
            f"Tried names: {candidates}\n"
            f"Available top-level folders:\n  " + "\n  ".join(available)
        )

    # --------------------------------------------------------------------- #
    # Public API
    # --------------------------------------------------------------------- #

    @property
    def mailbox_root(self):
        """Return the mailbox root folder (COM object)."""
        return self._mailbox_root

    @property
    def inbox_folder(self):
        """Return the Inbox/Posteingang folder (COM object)."""
        return self._inbox_folder

    def resolve_folder_from_config_path(self, config_path: str):
        """
        Resolve a folder path like:

            "Inbox"
            "Inbox\\2025"

        relative to this mailbox's Inbox/Posteingang folder.

        Rules:
          - The path segments are separated by backslashes ("\\").
          - If the first segment equals the Inbox name (case-insensitive) or
            the English variant ("Inbox"), it is *skipped* and the path is
            interpreted as subfolders of the Inbox.
          - Otherwise, the path is still interpreted as subfolders of the Inbox.

        Returns A COM folder object.

        Raises ValueError if any subfolder in the path cannot be found.
        """
        if not config_path:
            raise ValueError("Empty folder path in config")

        parts = [p for p in config_path.split("\\") if p]
        if not parts:
            raise ValueError(f"Invalid folder path in config: {config_path!r}")

        inbox_name = str(self._inbox_folder.Name)
        inbox_name_lower = inbox_name.lower()

        first_lower = parts[0].lower()
        if first_lower in (inbox_name_lower, "inbox"):
            parts = parts[1:]

        current = self._inbox_folder

        for sub in parts:
            try:
                current = current.Folders[sub]
            except Exception:
                # Case-insensitive fallback
                found = None
                for f in current.Folders:
                    try:
                        if str(f.Name).lower() == sub.lower():
                            found = f
                            break
                    except Exception:
                        continue

                if found is None:
                    try:
                        current_path = current.FolderPath
                    except Exception:
                        current_path = f"<{current.Name}>"

                    raise ValueError(
                        f"Could not find subfolder '{sub}' under {current_path} "
                        f"for path '{config_path}'."
                    )

                current = found

        return current
