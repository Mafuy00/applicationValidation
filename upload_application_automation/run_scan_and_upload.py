#!/usr/bin/env python
"""
One-off pass over mail already sitting in the IITSIP shared mailbox: download
any verified SIP application forms and upload them to the portal.

run_upload_watcher.py already catches up on unread mail at startup, so this is
mainly for a deliberate re-scan (for example after fixing selectors.json).

Run from the upload_application_automation folder:

    python run_scan_and_upload.py --limit 20
    python run_scan_and_upload.py --include-read --force
"""
import argparse

from sip_uploader.watcher import scan_existing


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--limit", type=int, default=50, help="Maximum number of emails to scan (newest first)."
    )
    parser.add_argument(
        "--include-read",
        action="store_true",
        help="Also scan mail that has already been read (default: unread only).",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-handle emails and re-upload forms even if a previous run already did.",
    )
    parser.add_argument(
        "--batch",
        default=None,
        help="Override the SIP batch to file the uploads under.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Rehearse only: fill in the portal form but never click Upload, so no records "
        "are created. Emails are left unmarked so a real run still processes them.",
    )
    args = parser.parse_args()

    scan_existing(
        limit=args.limit,
        force=args.force,
        unread_only=not args.include_read,
        batch_label=args.batch,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
