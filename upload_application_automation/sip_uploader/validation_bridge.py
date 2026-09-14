"""
Single point of coupling to the validation automation.

Step 1 of this automation has to answer exactly the question
`validation_automation` already answers well: "is this .docx actually a SIP
application form?". Rather than copy that logic (and let the two copies drift
apart the next time the official template changes), we import it.

`validation_automation` is a sibling folder, not an installed package, so its
folder is put on sys.path here - in this one module - so the import stays
visible and easy to fix if the repo layout ever changes.

Nothing in this module writes to the validation automation's logs or state;
only its pure form-reading helpers are used.
"""
import sys

from . import config


def _ensure_validation_importable():
    root = config.VALIDATION_AUTOMATION_ROOT
    if root not in sys.path:
        sys.path.insert(0, root)


_ensure_validation_importable()

try:
    # Importing sip_validator.config creates that automation's own logs/ and
    # _temp_attachments/ folders as a side effect. Both already exist in a
    # working checkout and the operation is idempotent, so it is harmless.
    from sip_validator import config as _validation_config
    from sip_validator.docx_form_reader import (  # noqa: F401  (re-exported)
        DocxFormReadError,
        read_form_fields,
    )
except ImportError as e:  # pragma: no cover - only hit on a broken checkout
    raise ImportError(
        f"Could not import the SIP form reader from "
        f"'{config.VALIDATION_AUTOMATION_ROOT}'. The upload automation reuses "
        f"validation_automation's form detection; make sure that folder is "
        f"present alongside this one. Original error: {e}"
    ) from e


KNOWN_SIP_FORM_FIELD_NAMES = frozenset(_validation_config.KNOWN_SIP_FORM_FIELD_NAMES)
MIN_KNOWN_FIELDS_TO_TREAT_AS_SIP_FORM = _validation_config.MIN_KNOWN_FIELDS_TO_TREAT_AS_SIP_FORM
IDENTIFYING_FIELDS = list(_validation_config.IDENTIFYING_FIELDS)


def looks_like_sip_form(fields):
    """
    True if the parsed form fields match enough of the official template's
    known field names to be confident this is a SIP application form.

    Mirrors validation_automation's heuristic, using the same field list and
    threshold, so both automations agree on what counts as an application form.
    """
    matches = sum(1 for name in fields if name in KNOWN_SIP_FORM_FIELD_NAMES)
    return matches >= MIN_KNOWN_FIELDS_TO_TREAT_AS_SIP_FORM


def count_known_fields(fields):
    """How many known template fields were found (useful for diagnostics)."""
    return sum(1 for name in fields if name in KNOWN_SIP_FORM_FIELD_NAMES)


def extract_identifying_values(fields):
    """Organisation/contact values, used to label log rows and downloads."""
    out = {}
    for name in IDENTIFYING_FIELDS:
        entry = fields.get(name)
        out[name] = entry.get("value") if entry else ""
    return out
