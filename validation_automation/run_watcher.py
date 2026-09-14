#!/usr/bin/env python
"""
Entry point: watches the IITSIP shared mailbox's Inbox for NEW emails and
automatically validates any SIP application form (.docx) attachments,
logging results to logs/validation_log.csv.

On startup, this also runs a catch-up scan for any *unread* mail that
arrived while the watcher wasn't running (see sip_validator/state.py) -
already-processed and already-read items are skipped. Use --no-catch-up
to skip it and only watch for new / unread mail from this point on.

Requirements:
  - Desktop Outlook must be installed, running, and signed in, with the
    IITSIP shared mailbox already added to this Outlook profile.
  - This script must keep running for validation to happen (it listens for
    the "new mail arrived" event). Leave the terminal window open, or set it
    up to run at logon (see README.md).

Usage:
    python run_watcher.py
    python run_watcher.py --no-catch-up
"""
import argparse

from sip_validator.watcher import run_watcher


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--no-catch-up",
        action="store_true",
        help="Skip the startup catch-up scan; only watch for new mail arriving from now on.",
    )
    args = parser.parse_args()
    run_watcher(catch_up=not args.no_catch_up)


if __name__ == "__main__":
    main()
