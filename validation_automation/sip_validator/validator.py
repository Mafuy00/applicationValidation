"""
Validates a dict of extracted form-field values (from docx_form_reader)
against the required-field rules in config/field_rules.json.

Returns a list of missing/invalid required-field labels, grouped by section.
"""
import json
import re

from . import config

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
NUMERIC_RE = re.compile(r"^[0-9]+(\.[0-9]+)?$")


def load_rules(path=None):
    path = path or config.FIELD_RULES_PATH
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _is_placeholder(value, placeholders):
    if value is None:
        return True
    if isinstance(value, str) and value.strip() in placeholders:
        return True
    return False


def validate(fields, rules=None):
    """
    fields: dict from docx_form_reader.read_form_fields()
    rules: parsed field_rules.json (loaded fresh if not given)

    Returns dict:
      {
        "passed": bool,
        "missing": [ {"section":..., "label":..., "reason":...}, ... ],
        "checked_field_count": int,
      }
    """
    if rules is None:
        rules = load_rules()

    placeholders = set(rules.get("placeholder_values", []))
    missing = []
    checked_count = 0

    for section in rules.get("sections", []):
        section_name = section.get("section", "")

        for field_rule in section.get("fields", []):
            checked_count += 1
            name = field_rule["name"]
            label = field_rule["label"]
            ftype = field_rule["type"]

            entry = fields.get(name)
            value = entry.get("value") if entry else None

            if entry is None:
                missing.append({"section": section_name, "label": label, "reason": "field not found in document (template may have been modified)"})
                continue

            if ftype in ("text", "dropdown"):
                if isinstance(value, str):
                    value = value.strip()
                if not value or _is_placeholder(value, placeholders):
                    missing.append({"section": section_name, "label": label, "reason": "required but empty"})
                    continue

                if field_rule.get("numeric") and not NUMERIC_RE.match(value):
                    missing.append({"section": section_name, "label": label, "reason": f"expected a number, got '{value}'"})
                    continue

                if field_rule.get("email") and not EMAIL_RE.match(value):
                    missing.append({"section": section_name, "label": label, "reason": f"does not look like a valid email address: '{value}'"})
                    continue

        for group in section.get("checkbox_groups", []):
            checked_count += 1
            group_label = group["label"]
            group_fields = group["fields"]
            any_checked = any(
                fields.get(fname, {}).get("value") is True for fname in group_fields
            )
            if not any_checked:
                missing.append({"section": section_name, "label": group_label, "reason": "no option selected"})

    return {
        "passed": len(missing) == 0,
        "missing": missing,
        "checked_field_count": checked_count,
    }


if __name__ == "__main__":
    # Run from the validation_automation folder:
    #   python -m sip_validator.validator "path\to\form.docx"
    import sys

    from .docx_form_reader import read_form_fields

    path = sys.argv[1]
    fields = read_form_fields(path)
    result = validate(fields)
    print(json.dumps(result, indent=2, ensure_ascii=False))
