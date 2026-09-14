"""
Persistent "already handled" tracking for the upload automation.

Two separate records are kept:

  - processed Outlook EntryIDs, so the startup catch-up scan and repeated
    manual scans never download the same email's attachment twice, and
  - successfully uploaded file paths, so a retry after a partial failure
    never uploads the same form to the portal a second time.

Both are plain text files (one entry per line), which is enough for a
single-machine tool and keeps the state trivially inspectable and repairable.
These files are deliberately distinct from validation_automation's.
"""
import os

from . import config


def _ensure_file(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if not os.path.exists(path):
        open(path, "a", encoding="utf-8").close()


def _load(path):
    _ensure_file(path)
    with open(path, "r", encoding="utf-8") as f:
        return {line.strip() for line in f if line.strip()}


def _append(path, value):
    if not value:
        return
    _ensure_file(path)
    with open(path, "a", encoding="utf-8") as f:
        f.write(value + "\n")
        f.flush()


def load_processed_ids():
    """Outlook EntryIDs whose attachments have already been considered."""
    return _load(config.PROCESSED_IDS_PATH)


def mark_processed(entry_id):
    _append(config.PROCESSED_IDS_PATH, entry_id)


def _normalise(path):
    return os.path.normcase(os.path.abspath(path))


def load_uploaded_files():
    """Normalised absolute paths of forms already uploaded successfully."""
    return {_normalise(p) for p in _load(config.UPLOADED_FILES_PATH)}


def mark_uploaded(path):
    _append(config.UPLOADED_FILES_PATH, _normalise(path))


def is_uploaded(path, uploaded=None):
    uploaded = load_uploaded_files() if uploaded is None else uploaded
    return _normalise(path) in uploaded
