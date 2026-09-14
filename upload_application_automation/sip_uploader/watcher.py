"""
Ties the automation together: watch the IITSIP mailbox, download verified SIP
application forms (step 1), and upload each one to the intranet portal
(steps 2-7), recording any failure in Failed_upload.xlsx.

A browser session is opened per email that actually contains forms, rather
than being held open across an idle watch, so a portal session never expires
between arrivals.
"""
import os
import time
import traceback

import pythoncom
import win32com.client

from . import config, downloader, failure_log, state
from .downloader import DownloadedForm, download_sip_forms, get_watched_folder, unread_items
from .uploader import IntranetUploadSession, UploadError

OL_MAIL_ITEM_CLASS = 43  # olMail


def form_from_path(path):
    """Wraps an already-downloaded file so it can go through the same path."""
    path = os.path.abspath(path)
    return DownloadedForm(path=path, original_filename=os.path.basename(path))


def _record_failure(form, step, error, screenshot_path="", dry_run=False):
    failure_log.record_failure(
        step=step,
        error=error,
        file_path=form.path,
        email_subject=form.subject,
        sender=form.sender,
        screenshot_path=screenshot_path,
        dry_run=dry_run,
    )


def upload_forms(
    forms, batch_label=None, skip_already_uploaded=True, dry_run=None, pause_before_close=False
):
    """
    Runs steps 2-7 for each form in one browser session.

    Returns (succeeded, failed) counts. Individual failures are logged to
    Failed_upload.xlsx and reported in the terminal; they never stop the
    remaining forms from being attempted.

    dry_run: stop short of the click that commits each record, so the whole
        flow is exercised without creating anything in the portal. Nothing is
        recorded as uploaded, so a later real run still picks the form up.
    pause_before_close: wait for Enter before closing the browser, so the
        filled-in dialog can be inspected by hand.
    """
    dry_run = config.DRY_RUN if dry_run is None else dry_run
    forms = list(forms)
    if skip_already_uploaded:
        uploaded = state.load_uploaded_files()
        kept = []
        for form in forms:
            if state.is_uploaded(form.path, uploaded):
                print(f"  - already uploaded, skipping: {form.filename}")
            else:
                kept.append(form)
        forms = kept

    if not forms:
        return 0, 0

    if dry_run:
        print("DRY RUN: no record will be created in the portal.\n")

    succeeded = failed = 0
    session = IntranetUploadSession(
        dry_run=dry_run,
        # Leaving the browser open only helps if it is actually visible.
        headless=False if pause_before_close else None,
    )

    try:
        try:
            session.start()
            session.login()
        except UploadError as e:
            # Nothing can be uploaded this run - make sure every pending form
            # is accounted for rather than silently dropped.
            for form in forms:
                _record_failure(form, e.step, e, e.screenshot_path, dry_run=dry_run)
            return 0, len(forms)
        except Exception as e:
            for form in forms:
                _record_failure(form, "start_session", e, dry_run=dry_run)
            return 0, len(forms)

        for form in forms:
            try:
                result = session.upload_form(form.path, batch_label=batch_label)
            except UploadError as e:
                failed += 1
                _record_failure(form, e.step, e, e.screenshot_path, dry_run=dry_run)
                continue
            except Exception as e:
                failed += 1
                _record_failure(
                    form,
                    "upload",
                    f"{type(e).__name__}: {e}\n{traceback.format_exc()}",
                    dry_run=dry_run,
                )
                continue

            if result.success:
                succeeded += 1
                # A dry run never committed anything, so the form must stay
                # pending - otherwise the real run would skip it.
                if not result.dry_run:
                    state.mark_uploaded(form.path)
                    print(f"  Upload complete: {form.filename}")
                if result.message:
                    print(f"    {result.message}")
            else:
                failed += 1
                _record_failure(
                    form, "upload", result.message, result.screenshot_path, dry_run=dry_run
                )

        if pause_before_close:
            input("\nBrowser left open for inspection. Press Enter here to close it...")
    finally:
        session.close()

    return succeeded, failed


def process_mail_item(
    item, seen_ids=None, mark_seen=True, unread_only=False, batch_label=None, dry_run=None
):
    """
    Handles one Outlook mail item end to end: download any SIP application
    forms it carries, then upload them.

    seen_ids: EntryIDs already handled in a previous run; matching items are
        skipped so nothing is downloaded or uploaded twice. When mark_seen is
        True the item is recorded as handled once inspected, whatever the
        outcome (failures are recoverable from Failed_upload.xlsx).
    unread_only: skip mail already marked read in Outlook.
    dry_run: rehearse without committing anything. Implies the item is NOT
        recorded as handled, so the real run still processes it.
    """
    dry_run = config.DRY_RUN if dry_run is None else dry_run
    if dry_run:
        mark_seen = False
    try:
        if getattr(item, "Class", None) != OL_MAIL_ITEM_CLASS:
            return 0, 0
    except Exception:
        return 0, 0

    if unread_only:
        try:
            if not item.UnRead:
                return 0, 0
        except Exception:
            return 0, 0

    try:
        entry_id = item.EntryID
    except Exception:
        entry_id = None

    if seen_ids is not None and entry_id and entry_id in seen_ids:
        return 0, 0

    succeeded = failed = 0
    try:
        try:
            subject = item.Subject or "(no subject)"
        except Exception:
            subject = "(no subject)"

        forms = download_sip_forms(item)
        if forms:
            print(f"\nSIP form(s) found in: {subject}")
            succeeded, failed = upload_forms(forms, batch_label=batch_label, dry_run=dry_run)
    except Exception:
        # One bad email must never abort the scan or stop the watcher.
        traceback.print_exc()
    finally:
        if mark_seen and entry_id:
            state.mark_processed(entry_id)
            if seen_ids is not None:
                seen_ids.add(entry_id)

    return succeeded, failed


class _ItemsEventHandler:
    """win32com event sink for a Folder.Items collection's ItemAdd event."""

    def OnItemAdd(self, item):
        try:
            process_mail_item(
                item,
                seen_ids=getattr(self, "seen_ids", None),
                unread_only=True,
                dry_run=getattr(self, "dry_run", None),
            )
        except Exception:
            traceback.print_exc()


def run_watcher(poll_interval_seconds=0.5, catch_up=True, dry_run=None):
    """
    Watches the configured Outlook folder and uploads SIP forms as they
    arrive. Blocks until Ctrl+C.

    catch_up: before listening, process unread mail already sitting in the
        folder that hasn't been handled before - covering anything that
        arrived while the watcher wasn't running. Read mail is ignored.
    dry_run: rehearse every arrival without creating portal records.
    """
    dry_run = config.DRY_RUN if dry_run is None else dry_run
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = get_watched_folder(namespace)
        seen_ids = state.load_processed_ids()

        unread = unread_items(folder)
        print(f"Watching: {folder.FolderPath}  ({unread.Count} unread)")
        print(f"Downloading forms to: {config.DOWNLOAD_DIR}")
        print(f"Uploading to        : {config.LOGON_URL}")
        print(f"SIP batch           : {config.SIP_BATCH_LABEL}")
        if dry_run:
            print(f"Failures logged to  : {config.FAILED_UPLOAD_DRYRUN_XLSX_PATH}")
            print("\n*** DRY RUN: forms will be downloaded and the portal form filled in,")
            print("*** but the Upload button will NOT be clicked and no record created.")
            print("*** Emails are left unmarked, so a real run will still process them.\n")
        else:
            print(f"Failures logged to  : {config.FAILED_UPLOAD_XLSX_PATH}\n")

        if catch_up:
            print("Running catch-up scan on unread mail only...")
            # Snapshot before listening so the scan and the live handler can't
            # double-handle an item arriving mid-scan (seen_ids also guards it).
            snapshot = list(unread)
            total_ok = total_fail = 0
            for it in snapshot:
                ok, bad = process_mail_item(
                    it, seen_ids=seen_ids, unread_only=True, dry_run=dry_run
                )
                total_ok += ok
                total_fail += bad
            print(
                f"Catch-up scan complete ({len(snapshot)} unread item(s) checked, "
                f"{total_ok} {'rehearsed' if dry_run else 'uploaded'}, {total_fail} failed).\n"
            )

        verb = "rehearsing" if dry_run else "uploading"
        print(f"Live mode: {verb} SIP forms from new / unread emails as they arrive.")
        print("Press Ctrl+C to stop.\n")

        items = folder.Items
        # Keep the handler referenced for the lifetime of the watcher.
        handler = win32com.client.WithEvents(items, _ItemsEventHandler)  # noqa: F841
        handler.seen_ids = seen_ids
        handler.dry_run = dry_run

        while True:
            pythoncom.PumpWaitingMessages()
            time.sleep(poll_interval_seconds)
    except KeyboardInterrupt:
        print("\nStopping watcher.")
    finally:
        pythoncom.CoUninitialize()


def scan_existing(limit=50, force=False, unread_only=True, batch_label=None, dry_run=None):
    """
    One-off pass over mail already in the folder (newest first).

    force: re-handle items even if a previous run already processed them.
    dry_run: rehearse without creating portal records.
    """
    dry_run = config.DRY_RUN if dry_run is None else dry_run
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = get_watched_folder(namespace)

        items = unread_items(folder) if unread_only else folder.Items
        items.Sort("[ReceivedTime]", True)  # newest first

        seen_ids = None if force else state.load_processed_ids()

        scanned = total_ok = total_fail = 0
        for item in items:
            if scanned >= limit:
                break
            scanned += 1
            ok, bad = process_mail_item(
                item,
                seen_ids=seen_ids,
                mark_seen=not force,
                unread_only=unread_only,
                batch_label=batch_label,
                dry_run=dry_run,
            )
            total_ok += ok
            total_fail += bad

        print(f"\nScanned {scanned} email(s) in '{folder.FolderPath}'.")
        print(
            f"{'Rehearsed' if dry_run else 'Uploaded'} {total_ok} form(s); "
            f"{total_fail} failure(s)."
        )
        if total_fail:
            print(
                f"See "
                f"{config.FAILED_UPLOAD_DRYRUN_XLSX_PATH if dry_run else config.FAILED_UPLOAD_XLSX_PATH}"
                f" for details."
            )
        return total_ok, total_fail
    finally:
        pythoncom.CoUninitialize()


def upload_latest(batch_label=None, force=False, dry_run=None, pause_before_close=False):
    """Uploads the most recent form in the download folder."""
    path = downloader.latest_downloaded_form(exclude_uploaded=not force)
    if path is None:
        print(f"No application forms awaiting upload in {config.DOWNLOAD_DIR}")
        return 0, 0
    return upload_forms(
        [form_from_path(path)],
        batch_label=batch_label,
        skip_already_uploaded=not force,
        dry_run=dry_run,
        pause_before_close=pause_before_close,
    )
