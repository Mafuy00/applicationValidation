# SIP Application Form Validator

Watches the **IITSIP** shared Outlook mailbox for new emails and automatically
validates any attached **SIP Application Form (IIT)** `.docx` against the
required (`*`) fields defined by the official template, logging the result
for staff review. No auto-reply is sent to applicants; this is purely a
staff-side triage log.

## How it works

1. The form template (`sip-application-form-iit.docx`) uses legacy Word
   **form fields** (FORMTEXT / FORMCHECKBOX / FORMDROPDOWN), each with an
   internal field name (e.g. `org_name`, `uen`, `contact_email`). These
   values are embedded directly in the `.docx`'s XML and don't require Word
   to be opened/automated to read.
2. `sip_validator/docx_form_reader.py` parses that XML directly and returns
   every field's current value.
3. `sip_validator/validator.py` checks the extracted values against
   `config/field_rules.json`, which encodes every field marked with `*` on
   the template (i.e. "required").
4. `sip_validator/watcher.py` listens for new mail arriving in
   `IITSIP > Inbox` (via Outlook's `ItemAdd` event) and, for each `.docx`
   attachment that looks like a SIP form, runs steps 2–3 and appends one row
   to `logs/validation_log.csv`.
5. On startup, it also runs a **catch-up scan over unread mail only**
   (covering unread items that arrived while it wasn't running), before
   switching to live listening for new / unread arrivals. Every mail item
   it inspects is recorded (by Outlook `EntryID`) in
   `logs/processed_entry_ids.txt`, so restarting the watcher - or re-running
   a manual scan - never creates duplicate rows for an email that's already
   been logged. Already-read mail is ignored.

## Requirements

- Windows PC with **desktop Outlook** installed, running, and signed in.
- The `IITSIP` shared mailbox already added to this Outlook profile (so it
  shows up as its own store in the folder list).
- Python 3.9+ with the packages in the repo-level `requirements.txt`:

  ```powershell
  pip install -r ..\requirements.txt
  ```

## Running it

All commands below are run **from this `validation_automation/` folder**:

```powershell
cd validation_automation
```

### Option A — live, real-time validation (recommended)

Watches for new mail as it arrives, running the validator via Outlook's
`ItemAdd` event. **Must be left running** for live detection to work — if
you close it, new mail queues up untouched until you start it again.

```powershell
python run_watcher.py
```

Every time it starts, it **automatically catches up on unread mail** that
arrived while it wasn't running (already-read mail is skipped), then
switches to live mode for new / unread arrivals. Already-processed emails
are also skipped via the EntryID state file. Pass `--no-catch-up` if you
want it to skip the startup unread scan and only watch for new mail from
that point on:

```powershell
python run_watcher.py --no-catch-up
```

Leave the terminal window open, or set it up to run automatically:
- Task Scheduler → "Run at log on" → Action: `python.exe` with argument
  `run_watcher.py`, and `Start in` set to this `validation_automation`
  folder (the `Start in` path matters — it's how Python finds the
  `sip_validator` package).

### Option B — manual scan

Processes emails that are already sitting in the Inbox — useful for a
bounded/manual re-check (e.g. only the most recent N emails), independent
of the watcher. Safe to re-run anytime; it doesn't move or modify any
email, and by default it skips emails already logged in a previous
run/scan (no duplicate CSV rows).

```powershell
python run_scan_existing.py --limit 50
python run_scan_existing.py --force   # re-process even already-logged emails
```

## Where results go

`logs/validation_log.csv` — one row per SIP-form attachment processed, kept
sorted **newest first**: the most recent scan is at the top, and within a
single scan the most recently received email comes first. So today's
submissions sit above yesterday's the moment you open the sheet, with no
manual sorting. Rows whose timestamp is missing or unreadable sink to the
bottom rather than being dropped.

| Column | Meaning |
|---|---|
| `processed_at` | When the validator ran (UTC) |
| `received_time` | When the email arrived |
| `sender_name` / `sender_email` | Who sent it |
| `subject` | Email subject |
| `attachment_filename` | The `.docx` filename |
| `org_name` / `uen` / `contact_name` / `contact_email` | Pulled straight from the form, to help you identify which submission this is without opening the file |
| `passed` | `True` if every required field is filled in correctly |
| `missing_count` / `missing_details` | What's missing, e.g. `[Section 2: Main Contact] Main Contact - Tel no.: required but empty` |
| `status` | `processed`, `skipped_not_sip_form` (docx attachment that isn't a recognised SIP form), or `error` |
| `error` | Details if something went wrong reading the file |

Open the CSV in Excel and filter `passed = FALSE` to see which submissions
need follow-up.

## Required fields currently validated

See `config/field_rules.json` — this is the single source of truth. It
mirrors every `*`-marked field on the official template:

- **Organisation**: UEN, Organisation Name, Postal Code, Country, Tel no.,
  Type of Company, Nature of Business, No. of Staff
- **Main Contact**: Salutation, Name, Designation, Email, Tel no.
- **No. of Interns / Job Scope** (row 1): Diploma, No. of Interns, Job Scope
- **Internship Details**: Monthly Allowance, No. of Days per Week,
  Interview Requirement (at least one checkbox ticked)
- **Others**: Future-engagement contact preference

Checkbox-based requirements (e.g. Interview Requirement) are enforced
**strictly** — if the org didn't tick a box, it's flagged as missing, even
though in practice many real submissions currently leave this blank.

To add/change a required field, edit `config/field_rules.json` — no code
changes needed. Each field entry needs the internal form-field `name` (see
`tools/dump_form_fields.py` to inspect any `.docx` and list its field
names/types) and a human-readable `label`.

## Known limitations

- **Only catches unread mail while `run_watcher.py` is running, or via its
  startup unread catch-up scan / a manual scan.** Already-read mail is
  ignored. There's no cloud-side trigger (this uses desktop Outlook COM
  automation, not Microsoft Graph), so if the PC is off/asleep, unread mail
  sits in the Inbox until the watcher is started again (at which point its
  unread catch-up scan will pick it up) or you run `run_scan_existing.py`
  yourself.
- **Attachment filenames vary** (senders don't follow a consistent naming
  convention), so the tool identifies "is this a SIP form" by checking
  whether the `.docx` contains enough of the known internal field names
  (see `MIN_KNOWN_FIELDS_TO_TREAT_AS_SIP_FORM` in `sip_validator/config.py`),
  not by filename.
- **If someone edits the template's structure** (renames/removes a form
  field, or fills it in outside Word in a way that breaks the field code),
  extraction for that field may silently return empty/missing — matching
  the template's own instruction #3 ("Do not modify this form as it will
  affect the automated data extraction process").
- No auto-reply/notification is sent to applicants or staff; this is a
  log-only tool by design. If you later want an email/Teams notification or
  to move flagged emails into a subfolder, that logic would go in
  `sip_validator/watcher.py`'s `process_mail_item()`.

## Project structure

This automation lives in `validation_automation/`. It writes to the
repo-level `logs/` folder, which is shared with the repo's other
automations — see the [repo README](../README.md).

```
applicationValidation/
├── logs/                            # SHARED across automations (repo level)
│   ├── validation_log.csv           # output (created on first run)
│   └── processed_entry_ids.txt      # dedup state (created on first run)
├── requirements.txt                 # shared dependencies
└── validation_automation/           # <- this automation
    ├── config/
    │   └── field_rules.json         # required-field definitions (edit to change validation rules)
    ├── sip_validator/
    │   ├── config.py                # mailbox/folder/paths configuration
    │   ├── docx_form_reader.py      # reads legacy Word form-field values from a .docx
    │   ├── validator.py             # checks extracted values against field_rules.json
    │   ├── csv_logger.py            # writes results to logs/validation_log.csv, newest first
    │   ├── state.py                 # tracks already-processed EntryIDs (prevents duplicate rows)
    │   └── watcher.py               # Outlook ItemAdd listener + startup catch-up + per-email processing
    ├── tools/
    │   ├── dump_form_fields.py      # inspect any .docx's internal form-field names/types
    │   ├── list_outlook_folders.py  # list all Outlook stores/folders in this profile
    │   └── sort_log.py              # one-off re-sort of an existing log (newest first)
    ├── run_watcher.py               # entry point: live watcher
    └── run_scan_existing.py         # entry point: manual/catch-up scan
```

### Paths in code

`sip_validator/config.py` defines two roots so the split stays explicit:

- `AUTOMATION_ROOT` — this folder. Owns `config/field_rules.json` and the
  transient `_temp_attachments/` scratch directory.
- `REPO_ROOT` / `LOGS_DIR` — the repo root and its shared `logs/` folder.

If you move this automation, those are the only two constants to check.
