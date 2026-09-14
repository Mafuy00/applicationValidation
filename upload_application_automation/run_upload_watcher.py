#!/usr/bin/env python
"""
Main entry point for the SIP application upload automation.

Watches the IITSIP shared mailbox and, for every new or unread email carrying
a verified SIP application form, downloads the form and uploads it to the SIP
intranet portal under the configured batch.

On startup it first catches up on unread mail that arrived while it wasn't
running (skip with --no-catch-up), then stays running and handles new mail as
it arrives. Press Ctrl+C to stop.

Run from the upload_application_automation folder:

    python run_upload_watcher.py
"""
import argparse

from sip_uploader.watcher import run_watcher


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--no-catch-up",
        action="store_true",
        help="Skip the startup catch-up scan; only handle mail arriving from now on.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Rehearse only: fill in the portal form for each arrival but never click "
        "Upload, so no records are created. Emails are left unmarked so a real run "
        "still processes them.",
    )
    args = parser.parse_args()
    run_watcher(catch_up=not args.no_catch_up, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
