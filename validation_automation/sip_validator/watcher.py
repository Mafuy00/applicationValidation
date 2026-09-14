"""
Core logic to process a single Outlook MailItem: find .docx attachment(s)
that look like a SIP application form, extract + validate their fields,
and log the result. Also provides an Outlook "ItemAdd" event listener that
runs this automatically whenever new mail arrives in the watched folder.
"""
import os
import time
import traceback

import pythoncom
import win32com.client

from . import config, csv_logger, state
from .docx_form_reader import read_form_fields, DocxFormReadError
from .validator import load_rules, validate

OL_MAIL_ITEM_CLASS = 43  # olMail
OL_BY_VALUE = 1  # olAttachmentType: normal file attachment (not OLE / embedded)


def _looks_like_sip_form(fields):
    known = set(config.KNOWN_SIP_FORM_FIELD_NAMES)
    matches = sum(1 for name in fields if name in known)
    return matches >= config.MIN_KNOWN_FIELDS_TO_TREAT_AS_SIP_FORM


def _extract_identifying_values(fields):
    out = {}
    for name in config.IDENTIFYING_FIELDS:
        entry = fields.get(name)
        out[name] = entry.get("value") if entry else ""
    return out


def _safe_attachment_filename(att):
    """
    Returns the attachment's file name, or None if Outlook cannot expose one
    (common for embedded images, OLE objects, and some inline attachments).
    """
    try:
        # Skip non-file attachment types (OLE / embedded items raise COM errors
        # on FileName / SaveAsFile with "cannot perform this action on this
        # type of attachment").
        if int(getattr(att, "Type", OL_BY_VALUE)) != OL_BY_VALUE:
            return None
    except Exception:
        pass

    for attr in ("FileName", "DisplayName"):
        try:
            name = getattr(att, attr, None)
            if name:
                return str(name)
        except Exception:
            continue
    return None


def _iter_docx_attachments(attachments):
    """Yields (attachment, filename) for saveable .docx file attachments only."""
    for att in attachments:
        name = _safe_attachment_filename(att)
        if name and name.lower().endswith(".docx"):
            yield att, name


def process_mail_item(item, rules=None, seen_ids=None, mark_seen=True, unread_only=False):
    """
    Inspects one Outlook mail item's attachments; for each .docx attachment
    that looks like a SIP form, extracts + validates fields and logs a row.
    Safe to call on any Outlook item type (non-mail items are skipped).

    seen_ids: optional set of Outlook EntryIDs already processed in a
        previous run/scan. If the item's EntryID is in this set, it's
        skipped entirely (prevents duplicate CSV rows when the startup
        catch-up scan and manual re-scans overlap with already-processed
        mail). If mark_seen is True (default), the item's EntryID is
        recorded (both in `seen_ids` and on disk via `state.mark_processed`)
        once it's been inspected, regardless of outcome.

    unread_only: if True, skip mail that is already marked as read in Outlook.
    """
    try:
        if getattr(item, "Class", None) != OL_MAIL_ITEM_CLASS:
            return
    except Exception:
        return

    if unread_only:
        try:
            if not item.UnRead:
                return
        except Exception:
            return

    entry_id = None
    try:
        entry_id = item.EntryID
    except Exception:
        entry_id = None

    if seen_ids is not None and entry_id and entry_id in seen_ids:
        return  # already inspected in a previous run/scan

    try:
        _process_mail_item_body(item, rules=rules)
    except Exception:
        # Never let one bad email abort the catch-up scan / watcher.
        traceback.print_exc()
    finally:
        if mark_seen and entry_id:
            state.mark_processed(entry_id)
            if seen_ids is not None:
                seen_ids.add(entry_id)


def _process_mail_item_body(item, rules=None):
    subject = ""
    sender_name = ""
    sender_email = ""
    received_time = ""
    try:
        subject = item.Subject or ""
        sender_name = item.SenderName or ""
        try:
            sender_email = item.SenderEmailAddress or ""
        except Exception:
            sender_email = ""
        received_time = item.ReceivedTime
    except Exception:
        pass

    try:
        attachments = list(item.Attachments)
    except Exception:
        attachments = []

    docx_attachments = list(_iter_docx_attachments(attachments))
    if not docx_attachments:
        return  # nothing to validate on this email

    if rules is None:
        rules = load_rules()

    for att, filename in docx_attachments:
        # Sanitize the filename for the temp path (strip path separators etc.).
        safe_name = os.path.basename(filename).replace("\\", "_").replace("/", "_")
        save_path = os.path.abspath(
            os.path.join(config.TEMP_ATTACHMENT_DIR, f"proc_{int(time.time() * 1000)}_{safe_name}")
        )
        try:
            att.SaveAsFile(save_path)
            fields = read_form_fields(save_path)

            if not _looks_like_sip_form(fields):
                csv_logger.log_result(
                    received_time=received_time,
                    sender_name=sender_name,
                    sender_email=sender_email,
                    subject=subject,
                    attachment_filename=filename,
                    status="skipped_not_sip_form",
                )
                continue

            result = validate(fields, rules=rules)
            identifying = _extract_identifying_values(fields)
            csv_logger.log_result(
                received_time=received_time,
                sender_name=sender_name,
                sender_email=sender_email,
                subject=subject,
                attachment_filename=filename,
                identifying_values=identifying,
                validation_result=result,
                status="processed",
            )
        except DocxFormReadError as e:
            csv_logger.log_result(
                received_time=received_time,
                sender_name=sender_name,
                sender_email=sender_email,
                subject=subject,
                attachment_filename=filename,
                status="error",
                error=str(e),
            )
        except Exception as e:
            csv_logger.log_result(
                received_time=received_time,
                sender_name=sender_name,
                sender_email=sender_email,
                subject=subject,
                attachment_filename=filename,
                status="error",
                error=f"{type(e).__name__}: {e}\n{traceback.format_exc()}",
            )
        finally:
            if os.path.exists(save_path):
                try:
                    os.remove(save_path)
                except OSError:
                    pass


def get_watched_folder(namespace):
    store = next((s for s in namespace.Folders if s.Name == config.OUTLOOK_STORE_NAME), None)
    if store is None:
        raise RuntimeError(
            f"Could not find Outlook store named '{config.OUTLOOK_STORE_NAME}'. "
            f"Make sure the shared mailbox is added to this Outlook profile."
        )
    folder = store
    for part in config.OUTLOOK_FOLDER_PATH:
        folder = next((f for f in folder.Folders if f.Name == part), None)
        if folder is None:
            raise RuntimeError(f"Could not find folder path {config.OUTLOOK_FOLDER_PATH} under store '{config.OUTLOOK_STORE_NAME}'.")
    return folder


class _ItemsEventHandler:
    """win32com event sink for a Folder.Items collection's ItemAdd event."""

    def OnItemAdd(self, item):
        try:
            process_mail_item(
                item,
                rules=getattr(self, "rules", None),
                seen_ids=getattr(self, "seen_ids", None),
                unread_only=True,
            )
        except Exception:
            traceback.print_exc()


def _unread_items(folder):
    """Return Outlook Items restricted to unread mail only."""
    # Restrict is far cheaper than iterating the whole mailbox.
    return folder.Items.Restrict("[UnRead] = True")


def run_watcher(poll_interval_seconds=0.5, catch_up=True):
    """
    Blocks and watches the configured Outlook folder for new mail, running
    the validator automatically on arrival. Press Ctrl+C to stop.

    catch_up: if True (default), before starting the live listener this
        scans unread items currently in the folder and validates any that
        haven't been processed before (tracked via sip_validator.state).
        Read mail is ignored. This covers unread mail that arrived while
        the watcher wasn't running.
    """
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = get_watched_folder(namespace)
        rules = load_rules()
        seen_ids = state.load_processed_ids()

        unread = _unread_items(folder)
        print(
            f"Watching: {folder.FolderPath}  "
            f"(currently {folder.Items.Count} items, {unread.Count} unread)"
        )
        print(f"Logging results to: {config.LOG_CSV_PATH}")

        if catch_up:
            print("Running catch-up scan on unread mail only...")
            # Snapshot unread items before we start listening for ItemAdd,
            # so the catch-up scan and the live listener don't double-handle
            # an item that arrives mid-scan (it'll just be picked up by
            # whichever one sees it first; seen_ids prevents a duplicate).
            snapshot = list(unread)
            for it in snapshot:
                process_mail_item(
                    it, rules=rules, seen_ids=seen_ids, unread_only=True
                )
            print(f"Catch-up scan complete ({len(snapshot)} unread item(s) checked).\n")

        print("Live mode: validating new / unread emails as they arrive.")
        print("Press Ctrl+C to stop.\n")

        items = folder.Items
        # Keep a reference to the handler object alive for the lifetime of the watcher.
        handler = win32com.client.WithEvents(items, _ItemsEventHandler)  # noqa: F841
        handler.rules = rules
        handler.seen_ids = seen_ids

        while True:
            pythoncom.PumpWaitingMessages()
            time.sleep(poll_interval_seconds)
    except KeyboardInterrupt:
        print("\nStopping watcher.")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    run_watcher()
