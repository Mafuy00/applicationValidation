# SIP Application Upload Automation

Watches the **IITSIP** shared mailbox, downloads every attachment that is a
genuine SIP application form, and uploads it to the SIP intranet portal under
the configured batch. Anything that fails is recorded in
`logs/Failed_upload.xlsx` and announced in the terminal.

> **Before the first real run:** the portal's CSS selectors in
> `config/selectors.json` are educated guesses — the site is behind the TP
> intranet and could not be inspected while this was built. Run
> `python tools/inspect_page.py` once to capture the real ones (see
> [First-run setup](#first-run-setup)), then rehearse with
> `python run_upload_file.py --dry-run --keep-open`, which exercises the whole
> flow without creating a record (see
> [Testing without creating a record](#testing-without-creating-a-record)).

## What it does

| Step | Action |
|---|---|
| 1 | An email arrives in IITSIP > Inbox. Each `.docx` attachment is parsed and checked against the official form's field names; only real application forms are saved to `Documents\Applications`. |
| 2 | Logs in at `https://intranetapps.tp.edu.sg/sip/Account/LogOn` using credentials from Windows Credential Manager (see [Credentials](#credentials)). |
| 3 | Opens the **Application** tab, then **Manage Application** from its dropdown. |
| 4 | Clicks **Upload New Applications**. |
| 5 | Attaches the form just downloaded via the dialog's **Choose a file** control. |
| 6 | Selects **AY2627-AY2627** in the dialog's SIP batch dropdown, then clicks the dialog's **Save** button. |
| 7 | Waits for the upload to settle and confirms the outcome. On failure: writes a row to `logs/Failed_upload.xlsx`, saves a screenshot, and prints a notice in the terminal. |

## Setup

```powershell
# from the repo root
pip install -r requirements.txt
```

### Credentials

Store them once in **Windows Credential Manager**. The password is prompted
for without echo and encrypted against your Windows account, so it never
touches a file:

```powershell
cd upload_application_automation
python set_credentials.py           # prompt and store
python set_credentials.py --show    # what's configured (never prints the password)
python set_credentials.py --clear   # remove it again
```

This is the recommended route because **the repo lives inside a synced
OneDrive folder** — a password typed into a `.env` here would be uploaded to
the cloud along with everything else.

`get_credentials()` in `sip_uploader/config.py` resolves in this order:

| Order | Source | When to use it |
|---|---|---|
| 1 | `SIP_INTRANET_USERNAME` / `SIP_INTRANET_PASSWORD` environment variables, including anything loaded from a `.env` | A one-off override, or a CI runner injecting its own secrets |
| 2 | Windows Credential Manager, service name `sip_intranet_upload` | Normal use on this machine |

Because the environment wins, a leftover `.env` will silently shadow the
credential store — `set_credentials.py` warns you if that's the case.

#### Why GitHub environment secrets can't be used here

GitHub secrets are decrypted only inside a GitHub Actions job, and there is no
API to read a value back out (`gh secret list` shows names only). They are
therefore unavailable to a script started from your own terminal. Running this
automation on a GitHub-hosted runner isn't an option either: it needs the
desktop Outlook client for the shared mailbox and network access to
`intranetapps.tp.edu.sg`, and a cloud runner has neither. A self-hosted runner
on this PC would receive the secrets, but it lets anyone who can push to the
repo run code on this machine, which is a worse trade than Credential Manager.

### Browser

The automation drives the **Google Chrome already installed on the machine**,
through Playwright, with **Playwright Stealth** applied to the browser context
so that every page it opens masks the usual automation fingerprints
(`navigator.webdriver`, missing plugins, WebGL vendor).

Using the installed Chrome also avoids `python -m playwright install chromium`,
which fails on this network with `SELF_SIGNED_CERT_IN_CHAIN` — the corporate
proxy re-signs HTTPS traffic with its own root CA and Playwright's Node-based
downloader doesn't trust it. Nothing needs to be done about that: no download
is involved, and Chrome is already approved software.

If Chrome is missing, it falls back to Edge, then to Playwright's bundled
Chromium. Set `SIP_UPLOAD_BROWSER_CHANNEL=chrome` (or `msedge`) to pin one and
disable the fallback, and `SIP_UPLOAD_STEALTH=0` to turn stealth off.

Stealth is applied best-effort: if the package is missing or fails, the run
prints a warning and continues with a plain browser rather than aborting.

> To be clear about what stealth does and doesn't do here: it patches
> fingerprinting signals inside a browser that is already running, so it helps
> only if the portal or an SSO gateway rejects the session *because* it looks
> automated. It has no effect on the `SELF_SIGNED_CERT_IN_CHAIN` download
> error above, which happens in a Node HTTPS request before any browser exists.
> If you do want Playwright's own Chromium, the clean fix is to trust the
> corporate CA rather than work around it:
>
> ```powershell
> $env:NODE_EXTRA_CA_CERTS = "C:\path\to\corporate-root-ca.cer"
> python -m playwright install chromium
> ```
>
> Ask IT for the certificate, or export it from `certmgr.msc` under
> *Trusted Root Certification Authorities*.

## First-run setup

`config/selectors.json` holds every CSS selector the automation uses. Each
entry is a **list of candidates tried in order**, so you can put the portal's
real selector at the front and leave the guesses behind it as fallbacks.

To capture the real ones:

```powershell
cd upload_application_automation
python tools/inspect_page.py --out ..\logs\page_inspection.txt
```

It logs in, walks the flow as far as it can, and for each stage prints every
interactive element with a suggested selector, plus an `OK` / `MISSING`
verdict for each configured key. Paste the correct selectors into
`config/selectors.json` and re-run until every key reports `OK`.

Once `result.success` matches the portal's real confirmation banner, set
`SIP_UPLOAD_REQUIRE_SUCCESS_CONFIRMATION=1` so that an upload with no visible
confirmation is treated as a failure instead of a pass.

## Testing without creating a record

Add `--dry-run` to any entry point. It performs every step against the real
portal — logs in, opens Application > Manage Application, clicks *Upload New
Applications*, attaches the form, selects the batch — then **locates the Upload
button and deliberately does not click it**. Everything up to that click is
client-side only, so nothing reaches the system.

```powershell
cd upload_application_automation

# rehearse the newest downloaded form, and leave the browser open to inspect
python run_upload_file.py --dry-run --keep-open

# rehearse a specific form
python run_upload_file.py --dry-run --file "C:\...\Documents\Applications\form.docx"

# rehearse against real unread mail
python run_scan_and_upload.py --dry-run --limit 5

# rehearse the live watcher
python run_upload_watcher.py --dry-run
```

A dry run reads the dialog back and reports what it actually contained, which
is the evidence worth checking:

```
DRY RUN: found the Upload button (Upload) but did NOT click it - no record was created.
DRY RUN: attached file  = SIP Application Form (BCS).docx
DRY RUN: selected batch = AY2627-AY2627
DRY RUN: screenshot     = ...\logs\upload_screenshots\..._dryrun_dialog_ready.png
```

If the file or batch doesn't read back as expected, the run is reported as a
failure even though nothing was submitted — that is the point of the rehearsal.

Dry runs are deliberately side-effect free:

- forms are **not** recorded as uploaded, so a later real run still picks them up
- emails are **not** marked as handled, so the watcher will process them properly later
- failures go to `logs/Failed_upload_dryrun.xlsx`, never to the real
  `Failed_upload.xlsx` that staff triage from

There are two other options depending on how cautious you want to be:

| Approach | Touches the portal? | Covers |
|---|---|---|
| `python tools/inspect_page.py` | Navigates and opens the dialog only; never attaches a file | Login and all navigation selectors |
| `--dry-run` (recommended) | Fills the form in the browser; submits nothing | Everything except the final click |
| `--keep-open` with `--dry-run` | Same, but pauses so you can look | Same, plus your own eyeball check |

> One caveat worth confirming on your first dry run: this is safe because the
> portal uploads on the Upload click. If it turns out to upload the moment a
> file is chosen (some pages do this via AJAX), a dry run would no longer be
> side-effect free. The screenshot and the portal's own record list will tell
> you — check once, and if that is the case use `tools/inspect_page.py`
> instead, which never attaches a file.

## Running it

```powershell
cd upload_application_automation

# the automation: catch up on unread mail, then handle new mail as it arrives
python run_upload_watcher.py

# skip the startup catch-up
python run_upload_watcher.py --no-catch-up
```

Leave the window open — it stays running until Ctrl+C. Outlook must be running
and signed in, with the IITSIP shared mailbox added to the profile.

### Other entry points

```powershell
# deliberate re-scan of mail already in the folder
python run_scan_and_upload.py --limit 20

# retry a form that's already downloaded (e.g. a row from Failed_upload.xlsx)
python run_upload_file.py --list
python run_upload_file.py --file "C:\...\Documents\Applications\form.docx"
python run_upload_file.py            # newest form not yet uploaded
```

## Output

| Path | Contents |
|---|---|
| `Documents\Applications\` | Downloaded application forms (original filenames; a timestamp is appended only on a name clash). |
| `logs/Failed_upload.xlsx` | One row per failed upload, newest first: when, which step failed, the form, the SIP batch, the email, the error, and the screenshot path. |
| `logs/Failed_upload_dryrun.xlsx` | The same, for `--dry-run` rehearsals only, so tests never contaminate the real log. |
| `logs/upload_screenshots/` | Full-page screenshot taken at the moment of each failure. |
| `logs/upload_processed_entry_ids.txt` | Emails already handled, so nothing is downloaded twice. |
| `logs/uploaded_files.txt` | Forms already uploaded, so nothing reaches the portal twice. |

The `step` column in `Failed_upload.xlsx` is the quickest triage signal —
`login` means credentials or portal access, `select_sip_batch` means
`SIP_UPLOAD_BATCH` is missing from the dropdown or matches more than one option
(the error lists every option, so a typo is easy to spot), `click_upload_button`
means the dialog's Upload
button couldn't be found or was the wrong one, and anything else containing
`locate`/`click` usually means a selector in `config/selectors.json` needs
updating.

## How it decides a `.docx` is an application form

It reuses the validation automation's detection rather than keeping a second
copy: `sip_uploader/validation_bridge.py` imports `read_form_fields` and the
known-field list from `validation_automation/sip_validator/`, and applies the
same threshold (at least 10 of the official template's ~97 named form fields).
Both automations therefore always agree on what an application form is, and
nothing in `validation_automation/` is modified or written to by this one.

If that folder is ever moved or renamed, `validation_bridge.py` is the single
place to fix.

## Configuration

Everything below is optional and set through the environment (or `.env`);
defaults are in `sip_uploader/config.py`.

| Variable | Default | Purpose |
|---|---|---|
| `SIP_UPLOAD_BATCH` | `AY2627-AY2627` | SIP batch to file uploads under. Must match a dropdown option exactly; a `--dry-run` lists them all if it doesn't. |
| `SIP_UPLOAD_DOWNLOAD_DIR` | `<Documents>\Applications` | Where forms are saved |
| `SIP_INTRANET_LOGON_URL` | the portal's log-on page | Portal entry point |
| `SIP_UPLOAD_BROWSER_CHANNEL` | auto (Chrome → Edge → bundled Chromium) | Pin the browser to drive |
| `SIP_UPLOAD_STEALTH` | `1` | Apply Playwright Stealth to the browser context |
| `SIP_UPLOAD_HEADLESS` | `0` | Run without a visible window |
| `SIP_UPLOAD_SLOW_MO_MS` | `0` | Slow each action down when tuning selectors |
| `SIP_UPLOAD_TIMEOUT_MS` | `30000` | Per-action / navigation timeout |
| `SIP_UPLOAD_WAIT_SECONDS` | `8` | Step 7's settle time before judging the outcome |
| `SIP_UPLOAD_REQUIRE_SUCCESS_CONFIRMATION` | `0` | Treat "no success banner" as a failure |
| `SIP_UPLOAD_DRY_RUN` | `0` | Rehearse without ever clicking Upload (same as `--dry-run`) |

## Project structure

```
upload_application_automation/
├── config/
│   └── selectors.json          # every CSS selector, with fallbacks
├── sip_uploader/
│   ├── config.py               # paths, URLs, timeouts, credentials
│   ├── validation_bridge.py    # the only link to validation_automation
│   ├── downloader.py           # step 1: Outlook -> Documents\Applications
│   ├── selectors.py            # resolves a logical name to a locator
│   ├── uploader.py             # steps 2-7: the Playwright portal flow
│   ├── failure_log.py          # Failed_upload.xlsx + terminal notice
│   ├── state.py                # what's been downloaded / uploaded already
│   └── watcher.py              # ties it together; Outlook ItemAdd listener
├── tools/
│   └── inspect_page.py         # discover the portal's real selectors
├── set_credentials.py          # one-time credential setup (Credential Manager)
├── run_upload_watcher.py       # main entry point
├── run_scan_and_upload.py      # one-off mailbox re-scan
└── run_upload_file.py          # upload / retry a specific form
```

## How the dialog is located

Everything under `upload_dialog` in `selectors.json` is resolved **inside the
dialog**, never against the whole page. This is not a detail: the Manage
Application table behind the dialog has its own **SIP Batch** filter listing
every historical batch, so a page-wide `select[name*='Batch']` matches that
filter instead of the dialog's dropdown. Selecting a batch there reports success
while the dialog's own dropdown stays empty — a silent wrong-data failure.

The dialog is found by locating the file input and walking up to the nearest
enclosing container that is dialog-like *and* contains the dialog's buttons,
rather than by matching framework-specific class names. That container is tagged
with a `data-sip-upload-dialog` attribute and used as the scope. The walk-up
deliberately keeps going past a `<form>` or a `.modal-body`, because a dialog's
Save/Close pair usually sits in a footer outside both. Child frames are searched
too, so a modal whose body is an `<iframe>` works.

Two guards back this up: the batch selection is read back from the page and
fails loudly if the dropdown doesn't show it, and the submit button is refused
if its label is `Close`/`Cancel` or the `Upload New Applications` trigger.

## Known limitations

- **Selectors are partly verified.** Login, navigation, the dialog, the file
  attach, the batch dropdown and the Save button have been exercised against
  the live portal in a dry run. `result.success` has not. Run
  `tools/inspect_page.py` if the portal's markup changes.
- **Login assumes a username/password form.** If the portal redirects to
  single sign-on or asks for MFA, the run stops at the `login` step; run
  headed (the default) so you can complete sign-in by hand.
- **Success detection depends on `result.success`.** Until that matches the
  portal's real banner, an upload that shows neither an error nor a
  confirmation is reported as a pass with a note saying so.
- **A failed email is still marked as handled**, so the watcher doesn't retry
  it in a loop. Retry from `Failed_upload.xlsx` with `run_upload_file.py`.
- **Outlook must be running** and the IITSIP mailbox present in the profile.
- **Windows only** — Outlook COM and the Documents known-folder lookup.
