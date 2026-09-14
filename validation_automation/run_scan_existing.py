#!/usr/bin/env python
"""
One-off / manual catch-up scan: processes emails ALREADY sitting in the
IITSIP > Inbox folder (not just new arrivals). Useful for:
  - an initial backlog run
  - catching up on emails received while run_watcher.py wasn't running
    (though run_watcher.py now does this automatically on startup too -
    see sip_validator/state.py)
  - re-testing the validator without waiting for new mail

By default, emails already logged in a previous run/scan are skipped (no
duplicate CSV rows). Use --force to re-process everything regardless.

Usage:
    python run_scan_existing.py            # scans the most recent 50 emails
    python run_scan_existing.py --limit 200
    python run_scan_existing.py --force    # re-process even already-logged emails
"""
import argparse

import win32com.client

from sip_validator import config, state
from sip_validator.watcher import process_mail_item, get_watched_folder
from sip_validator.validator import load_rules


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=50, help="Max number of most-recent emails to scan")
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-process items even if already logged in a previous run/scan.",
    )
    args = parser.parse_args()

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = get_watched_folder(namespace)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # newest first

    rules = load_rules()
    seen_ids = None if args.force else state.load_processed_ids()

    scanned = 0
    for item in items:
        if scanned >= args.limit:
            break
        scanned += 1
        process_mail_item(item, rules=rules, seen_ids=seen_ids)

    print(f"Scanned {scanned} email(s) in '{folder.FolderPath}'.")
    print(f"Results logged to: {config.LOG_CSV_PATH}")


if __name__ == "__main__":
    main()
