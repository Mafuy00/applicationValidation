"""
Writes one row per processed email/attachment to a CSV log file for staff
review.

The file is kept sorted newest-first (most recent scan at the top, and
within a single scan the most recently received email first), so whoever
opens it in Excel sees today's submissions before yesterday's without
having to sort the sheet manually.
"""
import csv
import os
import time
from datetime import datetime, timezone

from . import config

FIELDNAMES = [
    "processed_at",
    "received_time",
    "sender_name",
    "sender_email",
    "subject",
    "attachment_filename",
    "org_name",
    "uen",
    "contact_name",
    "contact_email",
    "passed",
    "missing_count",
    "missing_details",
    "status",
    "error",
]

# Sort key used for rows whose timestamp is missing or unparseable, so they
# sink to the bottom instead of breaking the comparison.
_OLDEST = datetime.min.replace(tzinfo=timezone.utc)

# How long to wait for another process (e.g. a manual scan running while the
# watcher is live) to finish rewriting the log before assuming a stale lock.
_LOCK_TIMEOUT_SECONDS = 10


class _LogLock:
    """
    Cross-process lock guarding the read-modify-write of the CSV, so a manual
    scan and the live watcher can't clobber each other's rows.
    """

    def __init__(self, path):
        self.lock_path = f"{path}.lock"
        self.fd = None

    def __enter__(self):
        deadline = time.monotonic() + _LOCK_TIMEOUT_SECONDS
        while True:
            try:
                self.fd = os.open(self.lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                return self
            except FileExistsError:
                if time.monotonic() >= deadline:
                    # Lock left behind by a crashed run - reclaim it.
                    try:
                        os.remove(self.lock_path)
                    except OSError:
                        pass
                    deadline = time.monotonic() + _LOCK_TIMEOUT_SECONDS
                time.sleep(0.05)

    def __exit__(self, *exc_info):
        if self.fd is not None:
            os.close(self.fd)
            self.fd = None
            try:
                os.remove(self.lock_path)
            except OSError:
                pass
        return False


def _parse_timestamp(value):
    """Parse a logged timestamp into an aware datetime, or _OLDEST if unusable."""
    if not value:
        return _OLDEST
    try:
        parsed = datetime.fromisoformat(str(value).strip())
    except (TypeError, ValueError):
        return _OLDEST
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def _sort_key(row):
    # Newest scan first; within one scan, the most recently received email first.
    return (
        _parse_timestamp(row.get("processed_at")),
        _parse_timestamp(row.get("received_time")),
    )


def sort_rows(rows):
    """Returns rows ordered newest-first. Stable, so same-timestamp rows keep order."""
    return sorted(rows, key=_sort_key, reverse=True)


def read_rows(path=None):
    """Reads existing log rows, normalised to FIELDNAMES. Returns [] if no log yet."""
    path = path or config.LOG_CSV_PATH
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return []
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        return [{name: (row.get(name) or "") for name in FIELDNAMES} for row in csv.DictReader(f)]


def write_rows(rows, path=None):
    """
    Rewrites the log with the given rows. Writes to a temp file first and
    swaps it in, so an interrupted or failed write can't truncate the log.
    """
    path = path or config.LOG_CSV_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp_path = f"{path}.tmp"
    try:
        with open(tmp_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
            writer.writeheader()
            writer.writerows(rows)
        os.replace(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise


def log_result(
    *,
    received_time="",
    sender_name="",
    sender_email="",
    subject="",
    attachment_filename="",
    identifying_values=None,
    validation_result=None,
    status="processed",
    error="",
    path=None,
):
    """
    status: "processed" | "skipped_not_sip_form" | "error"
    validation_result: dict from validator.validate(), or None if not applicable
    identifying_values: dict of {field_name: value} for org_name/uen/contact_name/contact_email
    """
    path = path or config.LOG_CSV_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)

    identifying_values = identifying_values or {}
    passed = ""
    missing_count = ""
    missing_details = ""
    if validation_result is not None:
        passed = validation_result.get("passed")
        missing = validation_result.get("missing", [])
        missing_count = len(missing)
        missing_details = " | ".join(
            f"[{m['section']}] {m['label']}: {m['reason']}" for m in missing
        )

    row = {
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "received_time": str(received_time),
        "sender_name": sender_name,
        "sender_email": sender_email,
        "subject": subject,
        "attachment_filename": attachment_filename,
        "org_name": identifying_values.get("org_name", ""),
        "uen": identifying_values.get("uen", ""),
        "contact_name": identifying_values.get("contact_name", ""),
        "contact_email": identifying_values.get("contact_email", ""),
        "passed": passed,
        "missing_count": missing_count,
        "missing_details": missing_details,
        "status": status,
        "error": error,
    }

    # Re-sorting the whole file on each write also repairs any older log that
    # was written before newest-first ordering existed.
    with _LogLock(path):
        rows = read_rows(path)
        rows.append(row)
        write_rows(sort_rows(rows), path)
