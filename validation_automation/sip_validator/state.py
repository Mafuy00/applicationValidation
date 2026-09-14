"""
Tracks which Outlook mail items (by their stable EntryID) have already been
inspected by the validator, so that:

  - run_watcher.py can run an automatic catch-up scan at startup (covering
    mail that arrived while it wasn't running) without re-logging emails
    that were already processed in a previous run/scan, and
  - run_scan_existing.py can be re-run safely without creating duplicate
    rows in the CSV for the same email.

Storage is a plain text file, one Outlook EntryID per line. This is simple
and robust for a single-machine, single-process tool; no database needed.
"""
import os

from . import config


def _ensure_file():
    os.makedirs(os.path.dirname(config.PROCESSED_IDS_PATH), exist_ok=True)
    if not os.path.exists(config.PROCESSED_IDS_PATH):
        open(config.PROCESSED_IDS_PATH, "a", encoding="utf-8").close()


def load_processed_ids():
    """Returns the set of Outlook EntryIDs already marked as processed."""
    _ensure_file()
    with open(config.PROCESSED_IDS_PATH, "r", encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def mark_processed(entry_id):
    """Appends an EntryID to the processed-ids file (flushed immediately)."""
    if not entry_id:
        return
    _ensure_file()
    with open(config.PROCESSED_IDS_PATH, "a", encoding="utf-8") as f:
        f.write(entry_id + "\n")
        f.flush()
