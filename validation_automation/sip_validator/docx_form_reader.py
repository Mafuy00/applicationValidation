"""
Reads legacy Word "form field" (FORMTEXT / FORMCHECKBOX / FORMDROPDOWN)
values directly out of a .docx file's word/document.xml.

This does NOT require Word to be installed/automated - it works purely by
parsing the OOXML, which is fast and safe to run on many attachments.

Returns a dict: field_name -> {"type": "text"|"checkbox"|"dropdown", "value": ...}
  - text:     stripped string (the visible FORMTEXT result)
  - checkbox: bool (True if checked)
  - dropdown: the selected list-entry string (or None if it couldn't be resolved)
"""
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return f"{{{W_NS}}}{tag}"


def _is_on_off(el):
    """
    Interpret an OOXML CT_OnOff element (e.g. <w:checked/>, <w:default w:val="0"/>).

    Per the OOXML spec (ST_OnOff): if the element is present and @w:val is
    omitted, the value defaults to true. Explicit on values are
    "1" / "true" / "on"; off values are "0" / "false" / "off".
    """
    if el is None:
        return False
    val = el.get(qn("val"))
    if val is None:
        return True
    return val.strip().lower() in ("1", "true", "on")


class DocxFormReadError(Exception):
    pass


def read_form_fields(docx_path):
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml_bytes = z.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError, FileNotFoundError) as e:
        raise DocxFormReadError(f"Could not read '{docx_path}' as a .docx file: {e}")

    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        raise DocxFormReadError(f"Could not parse document.xml in '{docx_path}': {e}")

    fields = {}
    current = None  # dict describing the form field currently being parsed

    def flush_current():
        nonlocal current
        if current is None:
            return
        name = current.get("name")
        if not name:
            current = None
            return
        ftype = current["type"]
        if ftype == "checkBox":
            fields[name] = {"type": "checkbox", "value": current.get("checked", False)}
        elif ftype == "ddList":
            idx = current.get("result_index", 0)
            entries = current.get("entries", [])
            value = entries[idx] if 0 <= idx < len(entries) else None
            fields[name] = {"type": "dropdown", "value": value}
        else:  # textInput
            text = "".join(current.get("text_parts", [])).strip()
            fields[name] = {"type": "text", "value": text}
        current = None

    for el in root.iter():
        tag = el.tag

        if tag == qn("fldChar"):
            fld_type = el.get(qn("fldCharType"))

            if fld_type == "begin":
                ffdata = el.find(qn("ffData"))
                if ffdata is None:
                    continue  # not a form field (e.g. TOC/other field code)

                name_el = ffdata.find(qn("name"))
                name = name_el.get(qn("val")) if name_el is not None else None

                cb = ffdata.find(qn("checkBox"))
                dd = ffdata.find(qn("ddList"))

                if cb is not None:
                    # <w:checked> holds the CURRENT state (set when the user
                    # ticks/unticks the box). <w:default> is only the value
                    # used when the form field is reset, and is NOT updated
                    # when the user interacts with the checkbox - so it must
                    # not be used to determine whether the box is ticked
                    # when <w:checked> is present.
                    #
                    # Important: Word commonly writes a bare <w:checked/>
                    # (no w:val attribute) for a ticked box. Per OOXML
                    # ST_OnOff, omitted val means true - so we must NOT
                    # require w:val="1".
                    checked_el = cb.find(qn("checked"))
                    default_el = cb.find(qn("default"))
                    if checked_el is not None:
                        checked = _is_on_off(checked_el)
                    else:
                        checked = _is_on_off(default_el)
                    current = {"name": name, "type": "checkBox", "checked": checked}
                    flush_current()
                elif dd is not None:
                    result_el = dd.find(qn("result"))
                    try:
                        result_index = int(result_el.get(qn("val"))) if result_el is not None else 0
                    except (TypeError, ValueError):
                        result_index = 0
                    entries = [e.get(qn("val")) for e in dd.findall(qn("listEntry"))]
                    current = {"name": name, "type": "ddList", "result_index": result_index, "entries": entries}
                    flush_current()
                else:
                    current = {"name": name, "type": "textInput", "text_parts": [], "capturing": False}

            elif fld_type == "separate":
                if current is not None and current.get("type") == "textInput":
                    current["capturing"] = True

            elif fld_type == "end":
                if current is not None and current.get("type") == "textInput":
                    flush_current()
                current = None

        elif tag == qn("t"):
            if current is not None and current.get("type") == "textInput" and current.get("capturing"):
                current["text_parts"].append(el.text or "")

    return fields


if __name__ == "__main__":
    import sys
    import json

    path = sys.argv[1]
    result = read_form_fields(path)
    print(json.dumps(result, indent=2, ensure_ascii=False))
