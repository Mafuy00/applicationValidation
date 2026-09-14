"""
Central configuration for the SIP application upload automation.

Everything that an operator might reasonably need to change lives here or in
`config/selectors.json`. Secrets are never stored in this file - they are read
from the environment (optionally populated from a gitignored `.env` at the
repo root).
"""
import ctypes
import os
from ctypes import wintypes

# This package lives at <repo>/upload_application_automation/sip_uploader/.
# AUTOMATION_ROOT owns everything specific to this automation; REPO_ROOT is
# shared with the repo's other automations (notably the logs/ folder staff
# read, and the validation_automation whose form detection we reuse).
AUTOMATION_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
REPO_ROOT = os.path.abspath(os.path.join(AUTOMATION_ROOT, ".."))
LOGS_DIR = os.path.join(REPO_ROOT, "logs")
VALIDATION_AUTOMATION_ROOT = os.path.join(REPO_ROOT, "validation_automation")
ENV_FILE_PATH = os.path.join(REPO_ROOT, ".env")


def _load_env_file():
    """
    Populates os.environ from the repo-root .env, if python-dotenv is
    installed. Absent either the package or the file, real environment
    variables still work - so this is best-effort by design.
    """
    try:
        from dotenv import load_dotenv
    except ImportError:
        return False
    return load_dotenv(ENV_FILE_PATH, override=False)


_load_env_file()


def _env_bool(name, default=False):
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in ("1", "true", "yes", "on")


def _env_int(name, default):
    try:
        return int(os.environ.get(name, "").strip())
    except (AttributeError, ValueError):
        return default


_FOLDERID_DOCUMENTS = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"


def _documents_dir():
    """
    Resolves the user's real Documents folder via the Windows known-folder
    API. This matters here because Documents is redirected into OneDrive on
    this machine, so os.path.expanduser("~/Documents") points somewhere else.
    """

    class _GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", wintypes.DWORD),
            ("Data2", wintypes.WORD),
            ("Data3", wintypes.WORD),
            ("Data4", ctypes.c_ubyte * 8),
        ]

    try:
        guid = _GUID()
        if ctypes.windll.ole32.CLSIDFromString(_FOLDERID_DOCUMENTS, ctypes.byref(guid)) != 0:
            raise OSError("CLSIDFromString failed for FOLDERID_Documents")

        path_ptr = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(guid), 0, None, ctypes.byref(path_ptr)
        ) != 0:
            raise OSError("SHGetKnownFolderPath failed for FOLDERID_Documents")
        try:
            return path_ptr.value
        finally:
            ctypes.windll.ole32.CoTaskMemFree(path_ptr)
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Documents")


# ---------------------------------------------------------------------------
# Step 1 - mailbox + downloads
# ---------------------------------------------------------------------------

# Same shared mailbox the validation automation watches. Kept as its own
# constant (rather than imported) so this automation can be pointed at a
# different folder without touching validation_automation.
OUTLOOK_STORE_NAME = "IITSIP"
OUTLOOK_FOLDER_PATH = ["Inbox"]  # folder names from the store root

# Where verified SIP application forms are saved. Override with
# SIP_UPLOAD_DOWNLOAD_DIR if the folder ever moves.
DOWNLOAD_DIR = os.environ.get("SIP_UPLOAD_DOWNLOAD_DIR") or os.path.join(
    _documents_dir(), "Applications"
)

# Attachments land here first so a .docx can be inspected before we decide
# whether it belongs in DOWNLOAD_DIR. Cleared after every check.
TEMP_ATTACHMENT_DIR = os.path.join(AUTOMATION_ROOT, "_temp_attachments")

# ---------------------------------------------------------------------------
# Steps 2-7 - intranet portal
# ---------------------------------------------------------------------------

LOGON_URL = os.environ.get(
    "SIP_INTRANET_LOGON_URL", "https://intranetapps.tp.edu.sg/sip/Account/LogOn"
)

USERNAME_ENV_VAR = "SIP_INTRANET_USERNAME"
PASSWORD_ENV_VAR = "SIP_INTRANET_PASSWORD"

# Credentials are preferably kept in the OS credential store (Windows
# Credential Manager) rather than in a file. That matters here because this
# repo lives inside a synced OneDrive folder, so a plaintext .env would be
# copied to the cloud. Populate the store with: python set_credentials.py
KEYRING_SERVICE_NAME = "sip_intranet_upload"
KEYRING_USERNAME_KEY = "username"
KEYRING_PASSWORD_KEY = "password"

# The SIP batch to file every upload under (step 6). It must match one of the
# portal dropdown's options exactly. The portal names the academic year twice,
# as "AY<start><end>-AY<start><end>", so AY26/27 is "AY2627-AY2627" - not
# "AY26-AY2627". A --dry-run lists every option the portal currently offers if
# the label doesn't match, which is the quickest way to check next year's name.
SIP_BATCH_LABEL = os.environ.get("SIP_UPLOAD_BATCH", "AY2627-AY2627")

SELECTORS_PATH = os.path.join(AUTOMATION_ROOT, "config", "selectors.json")

# Headed by default: this portal sits behind corporate SSO and a visible
# browser makes the occasional manual intervention (and first-run selector
# tuning) far easier. Set SIP_UPLOAD_HEADLESS=1 for unattended runs.
HEADLESS = _env_bool("SIP_UPLOAD_HEADLESS", False)

# Google Chrome is the browser this automation is meant to run on. Driving an
# already-installed browser also sidesteps `playwright install chromium`, which
# TP's TLS-inspecting proxy blocks with SELF_SIGNED_CERT_IN_CHAIN.
#
# Setting SIP_UPLOAD_BROWSER_CHANNEL pins one browser and disables the
# fallback; leaving it empty tries the order below (None = Playwright's own
# bundled Chromium, used only as a last resort).
BROWSER_CHANNEL = os.environ.get("SIP_UPLOAD_BROWSER_CHANNEL", "").strip()
BROWSER_CHANNEL_ORDER = ["chrome", "msedge", None]

# Apply Playwright Stealth, which masks the signals that mark a browser as
# automated (navigator.webdriver, missing plugins, WebGL vendor and similar).
# It is applied to the browser context, so every page inherits it. If the
# package is missing or fails to apply, the run continues without it rather
# than aborting.
USE_STEALTH = _env_bool("SIP_UPLOAD_STEALTH", True)

# Slows every Playwright action down by N ms - useful when tuning selectors.
SLOW_MO_MS = _env_int("SIP_UPLOAD_SLOW_MO_MS", 0)

# Per-action / per-navigation timeout.
PAGE_TIMEOUT_MS = _env_int("SIP_UPLOAD_TIMEOUT_MS", 30000)

# Step 7: how long to let the upload settle before deciding the outcome.
UPLOAD_WAIT_SECONDS = _env_int("SIP_UPLOAD_WAIT_SECONDS", 8)

# After the wait, an upload is treated as failed if the portal shows an error.
# If it shows neither an error nor a recognised success banner, the default is
# to accept it (and say so in the run output). Set this to 1 once
# `result.success` in selectors.json matches the portal's real confirmation
# banner, to turn "no confirmation" into a logged failure instead.
REQUIRE_SUCCESS_CONFIRMATION = _env_bool("SIP_UPLOAD_REQUIRE_SUCCESS_CONFIRMATION", False)

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

# Dry run: walk the whole portal flow - log in, navigate, attach the form,
# select the batch - and stop immediately before the click that commits the
# record, so nothing is created in the live system. Everything up to that point
# is client-side only. Enable per-run with --dry-run on any entry point.
DRY_RUN = _env_bool("SIP_UPLOAD_DRY_RUN", False)

# Step 7: failures are recorded here for staff to review.
FAILED_UPLOAD_XLSX_PATH = os.path.join(LOGS_DIR, "Failed_upload.xlsx")

# Dry-run failures go to their own workbook so rehearsals never contaminate
# the real failure log that staff triage from.
FAILED_UPLOAD_DRYRUN_XLSX_PATH = os.path.join(LOGS_DIR, "Failed_upload_dryrun.xlsx")

# Screenshots of the page at the moment a failure occurred, so a failed row
# can be diagnosed after the fact.
SCREENSHOT_DIR = os.path.join(LOGS_DIR, "upload_screenshots")

# Outlook EntryIDs already handled by THIS automation. Deliberately a
# different file from the validation automation's, so the two never interfere.
PROCESSED_IDS_PATH = os.path.join(LOGS_DIR, "upload_processed_entry_ids.txt")

# Absolute paths of forms already uploaded successfully, so a re-run or a
# retry never uploads the same form to the portal twice.
UPLOADED_FILES_PATH = os.path.join(LOGS_DIR, "uploaded_files.txt")


def _credentials_from_env():
    """Environment variables, which also covers .env and any CI runner."""
    return (
        os.environ.get(USERNAME_ENV_VAR, "").strip(),
        os.environ.get(PASSWORD_ENV_VAR, "") or "",
    )


def _credentials_from_keyring():
    """
    The OS credential store (Windows Credential Manager on this machine).

    Returns ("", "") when keyring isn't installed or nothing is stored, so
    callers can simply fall through to the next source.
    """
    try:
        import keyring
    except ImportError:
        return "", ""
    try:
        username = keyring.get_password(KEYRING_SERVICE_NAME, KEYRING_USERNAME_KEY) or ""
        password = keyring.get_password(KEYRING_SERVICE_NAME, KEYRING_PASSWORD_KEY) or ""
    except Exception:
        # A locked or unavailable credential store must not crash a run before
        # the environment has even been considered.
        return "", ""
    return username.strip(), password


def credentials_source():
    """
    Which store the credentials would come from right now: "environment",
    "keyring", or None if neither is configured. Used for diagnostics.
    """
    if all(_credentials_from_env()):
        return "environment"
    if all(_credentials_from_keyring()):
        return "keyring"
    return None


def get_credentials():
    """
    Returns (username, password).

    Sources are tried in order:
      1. environment variables (also covers .env, and a CI runner's secrets)
      2. the OS credential store, so nothing need be written to disk at all

    Raises RuntimeError with actionable guidance rather than letting the
    automation fail later on a blank login form.
    """
    for reader in (_credentials_from_env, _credentials_from_keyring):
        username, password = reader()
        if username and password:
            return username, password

    raise RuntimeError(
        "Missing intranet credentials. Either:\n"
        "  1. (recommended) store them once in Windows Credential Manager, so no\n"
        "     password is ever written to a file:\n"
        "         cd upload_application_automation\n"
        "         python set_credentials.py\n"
        f"  2. or set {USERNAME_ENV_VAR} and {PASSWORD_ENV_VAR} as environment\n"
        f"     variables, or fill them into '{ENV_FILE_PATH}' (see .env.example).\n"
        "\n"
        "Option 1 needs the 'keyring' package; option 2's .env needs "
        "'python-dotenv'. Both are in requirements.txt."
    )


os.makedirs(LOGS_DIR, exist_ok=True)
os.makedirs(TEMP_ATTACHMENT_DIR, exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
