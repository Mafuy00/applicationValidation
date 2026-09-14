# SIP Application Automations

Automations supporting the Student Internship Programme (SIP) intake for the
**IITSIP** shared mailbox. Each automation is a self-contained folder with
its own entry points, configuration and README; they share this repo's
dependencies and the top-level `logs/` folder.

## Automations

| Folder | What it does |
|---|---|
| [`validation_automation/`](validation_automation/README.md) | Watches the IITSIP mailbox for incoming SIP application forms (`.docx`), checks every required (`*`) field, and logs a pass/fail row per submission for staff review. |
| [`upload_application_automation/`](upload_application_automation/README.md) | Watches the same mailbox, downloads verified application forms to `Documents\Applications`, and uploads them to the SIP intranet portal under the current batch. Failures land in `logs/Failed_upload.xlsx`. |

The two are independent and can be run together or separately. The upload
automation reuses the validator's "is this really an application form?"
detection (through `sip_uploader/validation_bridge.py`) rather than keeping a
second copy of it, but it never writes to the validator's logs or state.

## Repo layout

```
applicationValidation/
├── logs/                          # shared output - written by the automations, read by staff
│   ├── validation_log.csv
│   ├── processed_entry_ids.txt
│   ├── Failed_upload.xlsx
│   ├── upload_processed_entry_ids.txt
│   ├── uploaded_files.txt
│   └── upload_screenshots/
├── validation_automation/         # form validation + logging      (see its README)
├── upload_application_automation/ # download + portal upload       (see its README)
├── .env.example                   # template for the credentials in .env (gitignored)
├── requirements.txt               # shared Python dependencies
└── README.md                      # you are here
```

## Getting started

```powershell
pip install -r requirements.txt

# validation: check incoming forms and log the results
cd validation_automation
python run_watcher.py
```

```powershell
# upload: file incoming forms on the intranet portal
# (first copy .env.example to .env and fill in the credentials)
cd upload_application_automation
python tools/inspect_page.py    # once, to confirm the portal's CSS selectors
python run_upload_watcher.py
```

Each automation's README covers its own setup, Outlook requirements, output
format and known limitations.

## Conventions for adding a new automation

Follow the same shape so the repo stays predictable:

1. Create a top-level folder named `<purpose>_automation/`.
2. Keep its importable code in a package inside that folder (as
   `validation_automation/sip_validator/` does), with entry-point scripts
   `run_*.py` at the folder's top level.
3. Resolve paths from two explicit roots rather than hardcoding or relying
   on the current working directory — see
   `validation_automation/sip_validator/config.py`:
   - `AUTOMATION_ROOT` for anything the automation owns (its own config,
     scratch space)
   - `REPO_ROOT` / `LOGS_DIR` for shared output
4. Write output into the shared `logs/` folder using a distinct filename, so
   staff have one place to look and nothing collides.
5. Add any new dependencies to the shared `requirements.txt`.
6. Give the folder a README and link it in the table above.

## Data handling

`logs/` and `_temp_attachments/` contain real applicant and organisation
data, and `.env` holds intranet credentials. All are gitignored and must never
be committed. Downloaded application forms live outside the repo, in
`Documents\Applications`.
