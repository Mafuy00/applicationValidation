#!/usr/bin/env python
"""
Discovers the SIP portal's real CSS selectors.

The defaults in config/selectors.json are educated guesses, because the portal
sits behind the TP intranet and cannot be inspected from outside. This tool
logs in with your credentials, walks the upload flow as far as it can, and
prints every interactive element it finds at each stage along with a suggested
CSS selector - so you can paste the right ones into selectors.json.

It also reports which configured keys currently resolve, so you can see at a
glance what still needs fixing. Nothing is uploaded.

Run from the upload_application_automation folder:

    python tools/inspect_page.py
    python tools/inspect_page.py --pause --out ..\\logs\\page_inspection.txt
"""
import argparse
import io
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from sip_uploader import config, selectors  # noqa: E402
from sip_uploader.uploader import IntranetUploadSession, UploadError  # noqa: E402

_COLLECT_JS = """
() => {
  const pick = (el) => {
    if (el.id) return '#' + CSS.escape(el.id);
    const name = el.getAttribute('name');
    if (name) return el.tagName.toLowerCase() + "[name='" + name + "']";
    const cls = (el.getAttribute('class') || '').trim().split(/\\s+/).filter(Boolean);
    if (cls.length) return el.tagName.toLowerCase() + '.' + cls.map(c => CSS.escape(c)).join('.');
    const text = (el.innerText || '').trim().split('\\n')[0];
    if (text) return el.tagName.toLowerCase() + ":has-text('" + text.slice(0, 40) + "')";
    return el.tagName.toLowerCase();
  };
  const nodes = Array.from(document.querySelectorAll(
    "a, button, input, select, textarea, label, [role='button'], [role='dialog'], .modal"
  ));
  return nodes.slice(0, 250).map(el => ({
    tag: el.tagName.toLowerCase(),
    type: el.getAttribute('type') || '',
    id: el.id || '',
    name: el.getAttribute('name') || '',
    text: ((el.innerText || el.value || '').trim().split('\\n')[0] || '').slice(0, 55),
    visible: el.getClientRects().length > 0,
    suggested: pick(el),
    options: el.tagName === 'SELECT'
      ? Array.from(el.options).map(o => (o.label || o.textContent || '').trim()).slice(0, 40)
      : null,
  }));
}
"""


def dump_page(page, title, out):
    out.write(f"\n{'=' * 78}\n{title}\n  URL: {page.url}\n{'=' * 78}\n")
    try:
        elements = page.evaluate(_COLLECT_JS)
    except Exception as e:
        out.write(f"  (could not inspect the page: {e})\n")
        return

    visible = [e for e in elements if e["visible"]]
    hidden_inputs = [e for e in elements if not e["visible"] and e["tag"] in ("input", "select")]

    out.write(f"\n  {len(visible)} visible interactive element(s):\n\n")
    out.write(f"  {'TAG/TYPE':<16} {'TEXT':<40} SUGGESTED SELECTOR\n")
    out.write(f"  {'-' * 16} {'-' * 40} {'-' * 40}\n")
    for e in visible:
        label = f"{e['tag']}" + (f"/{e['type']}" if e["type"] else "")
        out.write(f"  {label:<16} {e['text']:<40} {e['suggested']}\n")
        if e["options"]:
            out.write(f"  {'':<16} {'options:':<40} {e['options']}\n")

    if hidden_inputs:
        out.write("\n  hidden input/select elements (file inputs are often hidden):\n")
        for e in hidden_inputs:
            label = f"{e['tag']}" + (f"/{e['type']}" if e["type"] else "")
            out.write(f"    {label:<16} {e['suggested']}\n")
            if e["options"]:
                out.write(f"    {'':<16} options: {e['options']}\n")


def report_configured(page, keys, out):
    out.write("\n  Configured selectors, checked against this page:\n")
    for key in keys:
        require_visible = not key.endswith("file_input")
        found = selectors.find(
            page, key, require_visible=require_visible, timeout_ms=0, optional=True
        )
        status = "OK     " if found is not None else "MISSING"
        out.write(f"    [{status}] {key}\n")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=None, help="Also write the report to this file.")
    parser.add_argument(
        "--pause",
        action="store_true",
        help="Keep the browser open at the end so you can inspect it by hand.",
    )
    parser.add_argument(
        "--headless", action="store_true", help="Run without a visible browser window."
    )
    args = parser.parse_args()

    buffer = io.StringIO()

    def out_write(text):
        sys.stdout.write(text)
        buffer.write(text)

    out = type("Out", (), {"write": staticmethod(out_write)})()

    out.write(f"SIP portal inspection\n  Logon URL: {config.LOGON_URL}\n")
    out.write(f"  Selector config: {config.SELECTORS_PATH}\n")

    session = IntranetUploadSession(headless=args.headless or config.HEADLESS)
    try:
        session.start()
        page = session.page

        page.goto(config.LOGON_URL, wait_until="domcontentloaded")
        dump_page(page, "STEP 2 - Log-on page", out)
        report_configured(page, ["login.username", "login.password", "login.submit"], out)

        try:
            session.login()
        except UploadError as e:
            out.write(f"\n  Could not log in ({e.step}): {e}\n")
            out.write("  Fix the login.* selectors or credentials, then re-run.\n")
            return

        dump_page(page, "STEP 3 - After log-on (find the 'Application' tab here)", out)
        report_configured(page, ["nav.application_tab"], out)

        try:
            session.open_manage_application()
        except UploadError as e:
            out.write(f"\n  Could not reach Manage Application ({e.step}): {e}\n")
            dump_page(page, "Page state when navigation failed", out)
            return

        dump_page(page, "STEP 4 - Manage Application (find 'Upload New Applications' here)", out)
        report_configured(page, ["manage.upload_new_applications"], out)

        try:
            session.open_upload_dialog()
        except UploadError as e:
            out.write(f"\n  Could not open the upload dialog ({e.step}): {e}\n")
            dump_page(page, "Page state when the dialog failed to open", out)
            return

        dump_page(page, "STEPS 5-6 - Upload dialog (file input + SIP batch dropdown)", out)
        report_configured(
            page,
            [
                "upload_dialog.container",
                "upload_dialog.file_input",
                "upload_dialog.sip_batch_select",
                "upload_dialog.upload_button",
            ],
            out,
        )

        out.write(
            "\nNext: copy the correct selectors above into the matching lists in\n"
            f"{config.SELECTORS_PATH} (put each real selector first, keep the rest\n"
            "as fallbacks), then run: python run_upload_file.py --list\n"
        )

        if args.pause:
            input("\nBrowser left open. Press Enter here to close it...")
    finally:
        session.close()
        if args.out:
            path = os.path.abspath(args.out)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                f.write(buffer.getvalue())
            print(f"\nReport written to: {path}")


if __name__ == "__main__":
    main()
