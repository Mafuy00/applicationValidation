#!/usr/bin/env python
"""
Re-sorts an existing logs/validation_log.csv so the newest scans are at the
top. New rows are already written in sorted order, so this is only needed
once to fix a log created before newest-first ordering existed.

Usage:
    python tools/sort_log.py
    python tools/sort_log.py --path "C:\\some\\other\\validation_log.csv"
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from sip_validator import config, csv_logger


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--path", default=config.LOG_CSV_PATH, help="CSV log file to sort")
    args = parser.parse_args()

    if not os.path.exists(args.path):
        print(f"No log file at: {args.path}")
        return

    with csv_logger._LogLock(args.path):
        rows = csv_logger.read_rows(args.path)
        csv_logger.write_rows(csv_logger.sort_rows(rows), args.path)

    print(f"Sorted {len(rows)} row(s), newest first: {args.path}")


if __name__ == "__main__":
    main()
