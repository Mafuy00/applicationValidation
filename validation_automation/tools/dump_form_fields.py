#!/usr/bin/env python
"""
Diagnostic: list every legacy Word form field (name, type, dropdown options)
embedded in a .docx, so you can look up the internal field name to reference
in config/field_rules.json.

Usage:
    python tools/dump_form_fields.py "path\\to\\sip-application-form-iit.docx"
"""
import argparse
import re
import zipfile


def dump(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")

    blocks = re.findall(r"<w:ffData>.*?</w:ffData>", xml, re.S)
    print(f"Total form fields: {len(blocks)}\n")

    for i, block in enumerate(blocks):
        name = re.search(r'<w:name w:val="([^"]*)"', block)
        if "textInput" in block:
            ftype = "textInput"
        elif "checkBox" in block:
            ftype = "checkBox"
        elif "ddList" in block:
            ftype = "ddList"
        else:
            ftype = "unknown"
        maxlen = re.search(r'<w:maxLength w:val="([^"]*)"', block)
        default = re.search(r'<w:default w:val="([^"]*)"', block)
        entries = re.findall(r'<w:listEntry w:val="([^"]*)"', block)

        print(
            i,
            name.group(1) if name else None,
            ftype,
            "maxlen=", maxlen.group(1) if maxlen else None,
            "default=", default.group(1) if default else None,
            "list=", entries if entries else "",
        )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", help="Path to the .docx file to inspect")
    args = parser.parse_args()
    dump(args.docx)


if __name__ == "__main__":
    main()
