#!/usr/bin/env python
"""
Uploads an application form that has already been downloaded, without going
near Outlook. Use this to retry a row from Failed_upload.xlsx once the cause
has been fixed.

Run from the upload_application_automation folder:

    python run_upload_file.py                       # newest form not yet uploaded
    python run_upload_file.py --file "C:\\path\\form.docx"
    python run_upload_file.py --list

To test the whole flow WITHOUT creating a record in the portal, add --dry-run:
it logs in, navigates, attaches the form and picks the batch, then stops
before clicking Upload. Add --keep-open to inspect the filled dialog yourself.

    python run_upload_file.py --dry-run --keep-open
"""
import argparse
import os
import sys

from sip_uploader import config, downloader
from sip_uploader.watcher import form_from_path, upload_forms, upload_latest


def _list_forms():
    paths = downloader.list_downloaded_forms()
    if not paths:
        print(f"No .docx forms in {config.DOWNLOAD_DIR}")
        return

    from sip_uploader import state

    uploaded = state.load_uploaded_files()
    print(f"Forms in {config.DOWNLOAD_DIR} (newest first):\n")
    for path in paths:
        mark = "uploaded" if state.is_uploaded(path, uploaded) else "pending "
        print(f"  [{mark}] {os.path.basename(path)}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--file", default=None, help="Path of the .docx form to upload.")
    parser.add_argument(
        "--list", action="store_true", help="List downloaded forms and their upload status."
    )
    parser.add_argument(
        "--force", action="store_true", help="Upload even if this form was already uploaded."
    )
    parser.add_argument("--batch", default=None, help="Override the SIP batch.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Rehearse only: log in, navigate, attach the form and select the batch, "
        "then stop WITHOUT clicking Upload, so no record is created in the portal.",
    )
    parser.add_argument(
        "--keep-open",
        action="store_true",
        help="Leave the (visible) browser open at the end so you can check the filled-in "
        "dialog by hand. Best combined with --dry-run.",
    )
    args = parser.parse_args()

    if args.list:
        _list_forms()
        return

    if args.file:
        path = os.path.abspath(args.file)
        if not os.path.isfile(path):
            print(f"No such file: {path}")
            sys.exit(1)
        succeeded, failed = upload_forms(
            [form_from_path(path)],
            batch_label=args.batch,
            skip_already_uploaded=not args.force,
            dry_run=args.dry_run,
            pause_before_close=args.keep_open,
        )
    else:
        succeeded, failed = upload_latest(
            batch_label=args.batch,
            force=args.force,
            dry_run=args.dry_run,
            pause_before_close=args.keep_open,
        )

    verb = "Rehearsed" if args.dry_run else "Uploaded"
    print(f"\n{verb} {succeeded} form(s); {failed} failure(s).")
    if failed:
        workbook = (
            config.FAILED_UPLOAD_DRYRUN_XLSX_PATH if args.dry_run else config.FAILED_UPLOAD_XLSX_PATH
        )
        print(f"See {workbook} for details.")
        sys.exit(1)


if __name__ == "__main__":
    main()
