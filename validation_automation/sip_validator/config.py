"""
Central configuration for the SIP application form validator.
Edit these constants if the mailbox/folder name ever changes.
"""
import os

# This package lives at <repo>/validation_automation/sip_validator/.
# AUTOMATION_ROOT holds everything owned by this automation (its rules,
# scratch space); REPO_ROOT is shared with the repo's other automations,
# which is where the logs staff read live.
AUTOMATION_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
REPO_ROOT = os.path.abspath(os.path.join(AUTOMATION_ROOT, ".."))
LOGS_DIR = os.path.join(REPO_ROOT, "logs")

# Outlook shared mailbox + folder to watch for new SIP application emails.
OUTLOOK_STORE_NAME = "IITSIP"
OUTLOOK_FOLDER_PATH = ["Inbox"]  # path of folder names from the store root

# Where results are logged.
LOG_CSV_PATH = os.path.join(LOGS_DIR, "validation_log.csv")

# Tracks which mail items (by Outlook EntryID) have already been inspected,
# so the startup catch-up scan and repeated manual scans don't create
# duplicate rows in the CSV for the same email.
PROCESSED_IDS_PATH = os.path.join(LOGS_DIR, "processed_entry_ids.txt")

# Where attachments are temporarily saved for reading (deleted immediately after).
TEMP_ATTACHMENT_DIR = os.path.join(AUTOMATION_ROOT, "_temp_attachments")

# Field rules file.
FIELD_RULES_PATH = os.path.join(AUTOMATION_ROOT, "config", "field_rules.json")

# Identifying fields (non-validated) pulled into the log purely to help staff
# triage which organisation/contact a flagged email belongs to.
IDENTIFYING_FIELDS = ["org_name", "uen", "contact_name", "contact_email"]

# A .docx attachment is treated as a SIP application form if at least this
# many of the known form-field names are found in it (guards against random
# unrelated .docx attachments, since real senders don't use a consistent
# attachment filename convention).
MIN_KNOWN_FIELDS_TO_TREAT_AS_SIP_FORM = 10

# All legacy form-field names known to exist in the official sip-application-form-iit
# template, used only for the "is this actually a SIP form" heuristic above.
KNOWN_SIP_FORM_FIELD_NAMES = [
    "uen", "org_name", "block", "street", "floor", "unit_no", "building", "state",
    "org_country", "org_postal", "org_tel", "org_faxno", "org_email", "org_email2",
    "org_url", "org_type", "biz_nature", "biz_nature_others", "org_staff_strength",
    "org_staff_strength2", "org_address2", "org_postal2", "contact_pref",
    "contact_name", "contact_dept", "contact_designation", "contact_email",
    "contact_tel", "contact_mobile", "contact_fax", "dip1", "dip1_no", "dip1_job",
    "dip1_prj_mentor", "dip2", "dip2_no", "dip2_job", "dip2_prj_mentor", "dip3",
    "dip3_no", "dip3_job", "dip3_prj_mentor", "dip4", "dip4_no", "dip4_job",
    "dip4_prj_mentor", "dip5", "dip5_no", "dip5_job", "dip5_lead", "Allowance",
    "other_allow", "OT_pay", "OT_off", "OT_na", "trsp_yes", "trsp_ot", "trsp_na",
    "no_day_perwk", "wkday_fr", "wkday_to", "sat_fr", "sat_to", "sun_fr", "sun_to",
    "shift_req", "interview_face", "interview_phone", "interview_na",
    "interview_remote", "special_req", "sipo_flag", "overseatrip_period",
    "overseastrip_country", "osip_country", "osip_country_other", "osip_state",
    "osip_city", "osip_airfare", "osip_others", "airport_pickup", "osip_hotel",
    "osip_meal", "osip_insurance", "relation_parent", "relation_sister",
    "relation_subsidy", "relation_partner", "relation_na", "know_email",
    "know_referral", "know_website", "know_others", "know_others_details",
    "tp_staff_name", "tp_staff_diploma", "future_engagemt",
]

os.makedirs(TEMP_ATTACHMENT_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)
