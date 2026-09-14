"""
Step 1: pull SIP application forms out of the IITSIP shared mailbox.

Every .docx attachment is saved to a scratch folder and parsed first; only
attachments that are genuinely SIP application forms (per the shared
detection in `validation_bridge`) are kept and moved into the download
folder. Anything else is discarded, so unrelated .docx attachments never
reach the portal.
"""
import os
import time
from dataclasses import dataclass, field

from . import config, state
from .validation_bridge import (
    DocxFormReadError,
    count_known_fields,
    extract_identifying_values,
    looks_like_sip_form,
    read_form_fields,
)

OL_MAIL_ITEM_CLASS = 43  # olMail
OL_BY_VALUE = 1  # olAttachmentType: a normal file attachment (not OLE/embedded)


@dataclass
class DownloadedForm:
    """A verified SIP application form saved to the download folder."""

    path: str
    original_filename: str
    subject: str = ""
    sender: str = ""
    received_time: str = ""
    identifying_values: dict = field(default_factory=dict)

    @property
    def filename(self):
        return os.path.basename(self.path)


def get_watched_folder(namespace):
    """Resolves the configured Outlook store + folder path to a MAPI folder."""
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
            raise RuntimeError(
                f"Could not find folder path {config.OUTLOOK_FOLDER_PATH} under "
                f"store '{config.OUTLOOK_STORE_NAME}'."
            )
    return folder


def unread_items(folder):
    """Outlook Items restricted to unread mail (far cheaper than iterating all)."""
    return folder.Items.Restrict("[UnRead] = True")


def _safe_attachment_filename(att):
    """
    The attachment's file name, or None when Outlook cannot expose one.

    Embedded images and OLE objects raise a COM error on FileName/SaveAsFile
    ("Outlook cannot perform this action on this type of attachment"), so
    they are filtered out here rather than blowing up mid-scan.
    """
    try:
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


def iter_docx_attachments(attachments):
    """Yields (attachment, filename) for saveable .docx file attachments only."""
    for att in attachments:
        name = _safe_attachment_filename(att)
        if name and name.lower().endswith(".docx"):
            yield att, name


def _sanitise(filename):
    """Strips path separators so an attachment name can't escape its folder."""
    return os.path.basename(filename).replace("\\", "_").replace("/", "_")


def _unique_destination(directory, filename):
    """
    A non-colliding path inside `directory`, preserving the original name
    where possible so staff can still recognise the file on the portal.
    """
    stem, ext = os.path.splitext(filename)
    candidate = os.path.join(directory, filename)
    if not os.path.exists(candidate):
        return candidate

    stamp = time.strftime("%Y%m%d_%H%M%S")
    candidate = os.path.join(directory, f"{stem}_{stamp}{ext}")
    counter = 2
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{stem}_{stamp}_{counter}{ext}")
        counter += 1
    return candidate


def _mail_metadata(item):
    subject = sender = received_time = ""
    try:
        subject = item.Subject or ""
    except Exception:
        pass
    try:
        sender = item.SenderName or ""
    except Exception:
        pass
    try:
        received_time = str(item.ReceivedTime)
    except Exception:
        pass
    return subject, sender, received_time


def download_sip_forms(item, download_dir=None, verbose=True):
    """
    Saves every attachment on `item` that is a verified SIP application form.

    Returns a list of DownloadedForm. Attachments that aren't SIP forms, or
    that can't be parsed, are skipped (and reported) rather than downloaded.
    """
    download_dir = download_dir or config.DOWNLOAD_DIR
    os.makedirs(download_dir, exist_ok=True)

    try:
        if getattr(item, "Class", None) != OL_MAIL_ITEM_CLASS:
            return []
    except Exception:
        return []

    try:
        attachments = list(item.Attachments)
    except Exception:
        attachments = []

    candidates = list(iter_docx_attachments(attachments))
    if not candidates:
        return []

    subject, sender, received_time = _mail_metadata(item)
    downloaded = []

    for att, filename in candidates:
        safe_name = _sanitise(filename)
        temp_path = os.path.abspath(
            os.path.join(
                config.TEMP_ATTACHMENT_DIR,
                f"check_{int(time.time() * 1000)}_{safe_name}",
            )
        )
        moved_out_of_temp = False
        try:
            att.SaveAsFile(temp_path)

            try:
                fields = read_form_fields(temp_path)
            except DocxFormReadError as e:
                if verbose:
                    print(f"  - skipped '{filename}': not a readable .docx ({e})")
                continue

            if not looks_like_sip_form(fields):
                if verbose:
                    print(
                        f"  - skipped '{filename}': not a SIP application form "
                        f"({count_known_fields(fields)} known field(s) found)"
                    )
                continue

            destination = _unique_destination(download_dir, safe_name)
            os.replace(temp_path, destination)
            moved_out_of_temp = True

            form = DownloadedForm(
                path=destination,
                original_filename=filename,
                subject=subject,
                sender=sender,
                received_time=received_time,
                identifying_values=extract_identifying_values(fields),
            )
            downloaded.append(form)
            if verbose:
                org = form.identifying_values.get("org_name") or "unknown organisation"
                print(f"  - downloaded '{os.path.basename(destination)}' ({org})")
        except Exception as e:
            if verbose:
                print(f"  - skipped '{filename}': {type(e).__name__}: {e}")
        finally:
            if not moved_out_of_temp and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    return downloaded


def list_downloaded_forms(download_dir=None):
    """All .docx files in the download folder, newest first."""
    download_dir = download_dir or config.DOWNLOAD_DIR
    if not os.path.isdir(download_dir):
        return []

    paths = [
        os.path.join(download_dir, name)
        for name in os.listdir(download_dir)
        if name.lower().endswith(".docx") and not name.startswith("~$")
    ]
    paths = [p for p in paths if os.path.isfile(p)]
    return sorted(paths, key=os.path.getmtime, reverse=True)


def latest_downloaded_form(download_dir=None, exclude_uploaded=False):
    """
    The most recently saved application form, or None.

    exclude_uploaded: skip forms already recorded as successfully uploaded.
    """
    paths = list_downloaded_forms(download_dir)
    if exclude_uploaded:
        uploaded = state.load_uploaded_files()
        paths = [p for p in paths if not state.is_uploaded(p, uploaded)]
    return paths[0] if paths else None
