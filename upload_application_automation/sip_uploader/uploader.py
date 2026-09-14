"""
Steps 2-7: drive the TP SIP intranet portal with Playwright to upload an
application form.

The flow is:

  2. log on at config.LOGON_URL with the credentials from the environment
  3. open the "Application" menu, then "Manage Application"
  4. click "Upload New Applications"
  5. choose the downloaded form in the dialog that appears
  6. select the SIP batch (config.SIP_BATCH_LABEL), then click Upload
  7. wait for the upload to settle, then confirm the outcome

Every element is looked up through `selectors.py`, so no CSS lives in this
module. Any failure raises UploadError carrying the step that failed and a
screenshot of the page at that moment.
"""
import difflib
import os
import time
from dataclasses import dataclass

from . import config, selectors
from .selectors import SelectorNotFoundError


class UploadError(RuntimeError):
    """A step of the portal flow failed."""

    def __init__(self, step, message, screenshot_path=None):
        super().__init__(message)
        self.step = step
        self.screenshot_path = screenshot_path or ""


@dataclass
class UploadResult:
    success: bool
    file_path: str
    message: str = ""
    screenshot_path: str = ""
    # True when the run stopped short of committing the record (dry run), so
    # callers know not to record the form as uploaded.
    dry_run: bool = False


def _import_playwright():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise ImportError(
            "Playwright is required for the upload automation. Install it with:\n"
            "  pip install -r requirements.txt\n"
            "  python -m playwright install chromium"
        ) from e
    return sync_playwright


class IntranetUploadSession:
    """
    One browser session against the SIP portal.

    Used as a context manager; log in once and upload one or more forms:

        with IntranetUploadSession() as session:
            session.login()
            result = session.upload_form(path)
    """

    def __init__(
        self,
        headless=None,
        slow_mo_ms=None,
        timeout_ms=None,
        channel=None,
        use_stealth=None,
        dry_run=None,
    ):
        self.headless = config.HEADLESS if headless is None else headless
        self.slow_mo_ms = config.SLOW_MO_MS if slow_mo_ms is None else slow_mo_ms
        self.timeout_ms = config.PAGE_TIMEOUT_MS if timeout_ms is None else timeout_ms
        self.channel = config.BROWSER_CHANNEL if channel is None else channel
        self.use_stealth = config.USE_STEALTH if use_stealth is None else use_stealth
        self.dry_run = config.DRY_RUN if dry_run is None else dry_run
        self.stealth_applied = False
        # The option label actually chosen in step 6, which can differ in
        # padding or case from what was asked for.
        self.selected_batch_label = ""
        self._playwright = None
        self._browser = None
        self._context = None
        self.page = None
        self.logged_in = False

    # -- lifecycle ---------------------------------------------------------

    def _launch_candidates(self):
        """(channel, description) pairs to try, in order."""
        if self.channel:
            return [(self.channel, self.channel)]
        return [(name, name or "bundled Chromium") for name in config.BROWSER_CHANNEL_ORDER]

    def _launch_browser(self):
        attempts = []
        for channel, description in self._launch_candidates():
            try:
                browser = self._playwright.chromium.launch(
                    headless=self.headless,
                    slow_mo=self.slow_mo_ms or 0,
                    **({"channel": channel} if channel else {}),
                )
            except Exception as e:
                attempts.append(f"  - {description}: {str(e).splitlines()[0]}")
                continue
            print(f"  Browser: {description}")
            return browser

        raise UploadError(
            "launch_browser",
            "Could not start a browser. Tried:\n"
            + "\n".join(attempts)
            + "\nInstall Google Chrome, or set SIP_UPLOAD_BROWSER_CHANNEL to 'msedge', "
            "or run 'python -m playwright install chromium' to use Playwright's own build.",
        )

    def _apply_stealth(self, context):
        """
        Applies Playwright Stealth to the context so every page inherits it.

        Best-effort: a missing or failing stealth package degrades to a plain
        browser rather than stopping an upload.
        """
        if not self.use_stealth:
            return False
        try:
            from playwright_stealth import Stealth
        except ImportError:
            print(
                "  [warn] playwright-stealth is not installed; continuing without it "
                "(pip install -r requirements.txt)."
            )
            return False

        try:
            Stealth().apply_stealth_sync(context)
        except Exception as e:
            print(f"  [warn] Could not apply Playwright Stealth ({e}); continuing without it.")
            return False

        print("  Playwright Stealth: applied")
        return True

    def start(self):
        sync_playwright = _import_playwright()
        self._playwright = sync_playwright().start()
        try:
            self._browser = self._launch_browser()
        except Exception:
            self.close()
            raise

        self._context = self._browser.new_context(accept_downloads=False)
        self.stealth_applied = self._apply_stealth(self._context)
        self._context.set_default_timeout(self.timeout_ms)
        self._context.set_default_navigation_timeout(self.timeout_ms)
        self.page = self._context.new_page()
        return self

    def close(self):
        for closer in (
            getattr(self._context, "close", None),
            getattr(self._browser, "close", None),
            getattr(self._playwright, "stop", None),
        ):
            if closer is None:
                continue
            try:
                closer()
            except Exception:
                pass
        self._context = self._browser = self._playwright = None
        self.page = None

    def __enter__(self):
        return self.start()

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False

    # -- helpers -----------------------------------------------------------

    def screenshot(self, label):
        """Captures the current page; returns the path, or "" if unavailable."""
        if self.page is None:
            return ""
        stamp = time.strftime("%Y%m%d_%H%M%S")
        path = os.path.join(config.SCREENSHOT_DIR, f"{stamp}_{label}.png")
        try:
            os.makedirs(config.SCREENSHOT_DIR, exist_ok=True)
            self.page.screenshot(path=path, full_page=True)
            return path
        except Exception:
            return ""

    def _fail(self, step, message, cause=None):
        shot = self.screenshot(f"fail_{step}")
        raise UploadError(step, message, screenshot_path=shot) from cause

    def _find(self, key_path, **kwargs):
        return selectors.find(self.page, key_path, **kwargs)

    def _settle(self):
        """Best-effort wait for the page to stop working."""
        for state_name in ("domcontentloaded", "networkidle"):
            try:
                self.page.wait_for_load_state(state_name, timeout=self.timeout_ms)
            except Exception:
                pass

    def _visible_message(self, key_path, scope=None):
        return selectors.first_visible_text(scope or self.page, key_path)

    # -- step 2: log on ----------------------------------------------------

    def login(self):
        username, password = config.get_credentials()

        try:
            self.page.goto(config.LOGON_URL, wait_until="domcontentloaded")
        except Exception as e:
            self._fail("open_logon_page", f"Could not open {config.LOGON_URL}: {e}", e)

        try:
            self._find("login.username").fill(username)
            self._find("login.password").fill(password)
        except SelectorNotFoundError as e:
            self._fail("locate_login_fields", str(e), e)

        try:
            self._find("login.submit").click()
        except SelectorNotFoundError:
            # Some login forms submit on Enter with no visible button.
            try:
                self._find("login.password").press("Enter")
            except SelectorNotFoundError as e:
                self._fail("submit_login", str(e), e)

        self._settle()

        error_text = self._visible_message("login.error")
        still_on_login = "/Account/LogOn".lower() in (self.page.url or "").lower()
        password_still_shown = (
            selectors.find(self.page, "login.password", timeout_ms=0, optional=True) is not None
        )

        if error_text:
            self._fail("login", f"Login rejected by the portal: {error_text}")
        if still_on_login and password_still_shown:
            self._fail(
                "login",
                "Still on the log-on page after submitting credentials. Check "
                f"{config.USERNAME_ENV_VAR} / {config.PASSWORD_ENV_VAR}, or whether "
                f"the portal requires additional sign-in steps.",
            )

        self.logged_in = True
        print(f"  Logged in to the SIP portal as '{username}'.")
        return True

    # -- step 3: Application > Manage Application --------------------------

    def open_manage_application(self):
        try:
            tab = self._find("nav.application_tab")
        except SelectorNotFoundError as e:
            self._fail("open_application_tab", str(e), e)

        try:
            tab.click()
        except Exception as e:
            self._fail("open_application_tab", f"Could not click the Application tab: {e}", e)

        # The dropdown may open on click or on hover, depending on the theme.
        manage = selectors.find(self.page, "nav.manage_application", timeout_ms=3000, optional=True)
        if manage is None:
            try:
                tab.hover()
            except Exception:
                pass
            try:
                manage = self._find("nav.manage_application")
            except SelectorNotFoundError as e:
                self._fail("open_manage_application", str(e), e)

        try:
            manage.click()
        except Exception as e:
            self._fail("open_manage_application", f"Could not click 'Manage Application': {e}", e)

        self._settle()
        print("  Opened Application > Manage Application.")
        return True

    # -- step 4: open the upload dialog ------------------------------------

    # Set on the dialog container we resolve, so it can then be addressed as an
    # ordinary locator regardless of what the portal's own markup calls it.
    DIALOG_MARKER_ATTR = "data-sip-upload-dialog"

    def open_upload_dialog(self):
        try:
            self._find("manage.upload_new_applications").click()
        except SelectorNotFoundError as e:
            self._fail("click_upload_new_applications", str(e), e)
        except Exception as e:
            self._fail(
                "click_upload_new_applications",
                f"Could not click 'Upload New Applications': {e}",
                e,
            )

        # Getting this scope right is critical. With the whole page as the
        # scope, a lookup can match the Manage Application table's own filter
        # controls instead of the dialog's - the table has a SIP Batch filter
        # listing every historical batch, so selecting a batch there "succeeds"
        # while the dialog's own dropdown stays empty. That fails silently,
        # which is worse than not working at all.
        #
        # The file input appearing is the real signal that the dialog is open,
        # and it doubles as the anchor for working out the dialog's boundaries.
        frame = self._wait_for_file_input_frame(timeout_ms=5000)

        configured = selectors.find(
            self.page, "upload_dialog.container", timeout_ms=0, optional=True
        )
        if configured is not None:
            if self._contains_file_input(configured):
                print("  Opened the upload dialog.")
                return configured
            # Matching a container that doesn't hold the file input would put us
            # right back to searching the wrong part of the page.
            print("  [note] 'upload_dialog.container' matched an element that does not "
                  "contain the file input; deriving the dialog from the file input instead.")

        if frame is None:
            print("  [warn] No file input found, so the dialog could not be isolated - "
                  "falling back to the whole page, which may match the wrong controls.")
            return self.page

        in_iframe = frame is not self.page.main_frame
        where = " (its body is an iframe)" if in_iframe else ""

        container, description = self._tag_dialog_container(frame)
        if container is not None:
            print(f"  Opened the upload dialog{where}: <{description}>.")
            return container

        print(f"  Opened the upload form{where}.")
        return frame

    def _frame_with_file_input(self):
        """
        The frame holding the upload dialog's file input. Checked across child
        frames too, because a portal modal's body is sometimes a separate
        document in an <iframe>, whose contents a page-level locator cannot see.
        """
        frames = [self.page.main_frame]
        frames.extend(f for f in self.page.frames if f is not self.page.main_frame)
        for frame in frames:
            try:
                if frame.locator("input[type='file']").count() > 0:
                    return frame
            except Exception:
                continue  # a frame can detach mid-check
        return None

    def _wait_for_file_input_frame(self, timeout_ms=None):
        """Polls for the dialog's file input to turn up, returning its frame."""
        timeout_ms = self.timeout_ms if timeout_ms is None else timeout_ms
        deadline = time.monotonic() + (timeout_ms / 1000.0)
        while True:
            frame = self._frame_with_file_input()
            if frame is not None:
                return frame
            if time.monotonic() >= deadline:
                return None
            time.sleep(0.25)

    @staticmethod
    def _contains_file_input(locator):
        try:
            return locator.locator("input[type='file']").count() > 0
        except Exception:
            return False

    def _tag_dialog_container(self, frame):
        """
        Walks up from the file input to the nearest dialog-like ancestor, tags
        it with a marker attribute and returns (locator, description).

        Anchoring on the file input beats a list of framework-specific class
        names: whatever the portal's modal is built from, the file input is
        unambiguously inside it, so an ancestor found this way is certain to be
        the right region of the page.

        The container must also contain the dialog's buttons. Stopping at the
        <form> is not enough - a dialog's Save/Close pair frequently sits in a
        footer outside the form - and stopping at .modal-body would miss
        .modal-footer for the same reason.
        """
        try:
            description = frame.evaluate(
                """(marker) => {
                    const input = document.querySelector("input[type='file']");
                    if (!input) return null;
                    document.querySelectorAll('[' + marker + ']')
                        .forEach(el => el.removeAttribute(marker));

                    const DIALOGISH = /modal|dialog|popup|lightbox|window/i;
                    const BUTTONS =
                        "button, input[type='submit'], input[type='button'], a[onclick]";

                    // Ancestors of the file input, innermost first.
                    const chain = [];
                    for (let el = input.parentElement;
                         el && el !== document.body;
                         el = el.parentElement) {
                        chain.push(el);
                    }
                    if (!chain.length) return null;

                    // Expanding as far as the button that OPENED the dialog means
                    // we have escaped the dialog and are back to the whole page.
                    const escaped = el =>
                        /upload\\s+new\\s+application/i.test(el.textContent || '');
                    const dialogish = el =>
                        el.getAttribute('role') === 'dialog'
                        || DIALOGISH.test(typeof el.className === 'string' ? el.className : '')
                        || DIALOGISH.test(el.id || '');
                    const hasButtons = el => !!el.querySelector(BUTTONS);

                    const usable = chain.filter(el => !escaped(el));
                    const target =
                        usable.find(el => dialogish(el) && hasButtons(el))
                        || usable.find(hasButtons)
                        || usable.find(dialogish)
                        || usable.find(el => el.tagName === 'FORM');

                    if (!target) return null;
                    target.setAttribute(marker, '1');

                    const cls = typeof target.className === 'string' ? target.className.trim() : '';
                    let described = target.tagName.toLowerCase();
                    if (target.id) described += '#' + target.id;
                    else if (cls) described += '.' + cls.split(/\\s+/)[0];
                    return described;
                }""",
                self.DIALOG_MARKER_ATTR,
            )
        except Exception:
            return None, ""

        if not description:
            return None, ""
        return frame.locator(f"[{self.DIALOG_MARKER_ATTR}]"), description

    # -- step 5: choose the file -------------------------------------------

    def choose_file(self, scope, file_path):
        file_path = os.path.abspath(file_path)
        if not os.path.isfile(file_path):
            self._fail("choose_file", f"File to upload does not exist: {file_path}")

        # A file input is often hidden behind a styled "Choose a file" label,
        # so visibility must not be required here.
        file_input = selectors.find(
            scope, "upload_dialog.file_input", require_visible=False, timeout_ms=5000, optional=True
        )
        if file_input is not None:
            try:
                file_input.set_input_files(file_path)
                print(f"  Selected file: {os.path.basename(file_path)}")
                return True
            except Exception as e:
                self._fail("choose_file", f"Could not attach '{file_path}': {e}", e)

        # No input in the DOM: the button must open a native file chooser.
        try:
            button = selectors.find(scope, "upload_dialog.choose_file_button")
            with self.page.expect_file_chooser(timeout=self.timeout_ms) as fc:
                button.click()
            fc.value.set_files(file_path)
            print(f"  Selected file via file chooser: {os.path.basename(file_path)}")
            return True
        except SelectorNotFoundError as e:
            self._fail("choose_file", str(e), e)
        except Exception as e:
            self._fail("choose_file", f"File chooser did not accept '{file_path}': {e}", e)

    # -- step 6: select the SIP batch --------------------------------------

    def select_sip_batch(self, scope, batch_label=None):
        batch_label = batch_label or config.SIP_BATCH_LABEL

        try:
            dropdown = selectors.find(scope, "upload_dialog.sip_batch_select")
        except SelectorNotFoundError as e:
            self._fail("select_sip_batch", str(e), e)

        try:
            tag = (dropdown.evaluate("el => el.tagName") or "").lower()
        except Exception:
            tag = ""

        if tag == "select":
            options = self._native_options(dropdown)
            matches = self._match_batch_options(options, batch_label)

            if len(matches) > 1:
                # Never guess: the batch decides which academic year the
                # application is filed under, so the wrong pick is worse than
                # stopping. "AY2324" alone, for instance, matches three options.
                self._fail(
                    "select_sip_batch",
                    f"SIP batch '{batch_label}' is ambiguous - it matches "
                    f"{[label for _, label in matches]}. Set SIP_UPLOAD_BATCH to the "
                    f"exact label you want.",
                )
            if not matches:
                self._fail(
                    "select_sip_batch",
                    f"SIP batch '{batch_label}' is not one of the dropdown's options."
                    f"{self._batch_hint(options, batch_label)}",
                )

            value, label = matches[0]
            try:
                # Selecting by value rather than label sidesteps any padding or
                # reformatting the portal applies to what it displays.
                dropdown.select_option(value=value)
            except Exception as e:
                self._fail(
                    "select_sip_batch", f"Could not select SIP batch '{label}': {e}", e
                )
            # Read back what the dropdown now displays. The portal repopulates
            # this list over AJAX, so a selection can be silently discarded -
            # and an empty dropdown is the symptom of having driven the wrong
            # element entirely.
            shown = self._selected_option_label(dropdown)
            if shown and self._normalise_batch(shown) != self._normalise_batch(label):
                self._fail(
                    "select_sip_batch",
                    f"Selected '{label}' but the dropdown now reads '{shown}'. The portal "
                    f"reset it, or 'upload_dialog.sip_batch_select' is not the dialog's own "
                    f"dropdown.",
                )
            if not shown:
                self._fail(
                    "select_sip_batch",
                    f"Selected '{label}' but the dropdown still shows nothing. It is "
                    f"probably not the dialog's own SIP Batch dropdown - the Manage "
                    f"Application table behind the dialog has a filter of the same name.",
                )

            self.selected_batch_label = label
            print(f"  Selected SIP batch: {label}")
            return True

        # Custom (non-<select>) dropdown widget: open it, then click the option.
        try:
            dropdown.click()
            option = scope.locator(f"li:has-text('{batch_label}'), a:has-text('{batch_label}'), "
                                   f"option:has-text('{batch_label}')").first
            option.click(timeout=self.timeout_ms)
            self.selected_batch_label = batch_label
            print(f"  Selected SIP batch: {batch_label}")
            return True
        except Exception as e:
            self._fail(
                "select_sip_batch",
                f"Could not select SIP batch '{batch_label}' from the dropdown: {e}",
                e,
            )

    def _native_options(self, dropdown):
        """[(value, label)] for every option of a <select>, in document order."""
        try:
            return dropdown.evaluate(
                "el => Array.from(el.options).map(o => "
                "[o.value, (o.label || o.textContent || '').trim()])"
            ) or []
        except Exception:
            return []

    @staticmethod
    def _normalise_batch(text):
        """Comparison form for batch labels: case-folded, padding collapsed."""
        return " ".join((text or "").split()).lower()

    def _match_batch_options(self, options, batch_label):
        """
        The options matching `batch_label`, preferring an exact match on either
        the label or the underlying value and only falling back to a substring
        match when nothing matches exactly.

        Returns every match rather than the first, so the caller can refuse an
        ambiguous one instead of picking the wrong academic year.
        """
        wanted = self._normalise_batch(batch_label)
        if not wanted:
            return []

        exact = [
            (value, label)
            for value, label in options
            if wanted in (self._normalise_batch(label), self._normalise_batch(value))
        ]
        if exact:
            return exact

        return [
            (value, label)
            for value, label in options
            if label and wanted in self._normalise_batch(label)
        ]

    def _batch_hint(self, options, batch_label):
        """
        The closest option labels followed by the full list, so a typo in
        SIP_UPLOAD_BATCH is obvious from the error alone.
        """
        labels = [label for _, label in options if label]
        if not labels:
            return " The dropdown reported no options at all."
        close = difflib.get_close_matches(batch_label, labels, n=3, cutoff=0.4)
        hint = f" Did you mean {close}?" if close else ""
        return f"{hint} Available: {labels}"

    # -- step 6b: click Upload ---------------------------------------------

    def click_upload_button(self, scope):
        """
        Clicks the button that submits the dialog - labelled "Save" on this
        portal. Required: the form is never assumed to submit on its own.
        """
        try:
            button = selectors.find(scope, "upload_dialog.upload_button")
        except SelectorNotFoundError as e:
            # Listing what the dialog really contains turns a selector miss into
            # a one-step fix instead of a separate inspection run.
            self._fail(
                "click_upload_button", f"{e}\n{self._describe_dialog_controls(scope)}", e
            )

        try:
            label = " ".join(((button.inner_text() or "") or button.get_attribute("value") or "").split())
        except Exception:
            label = ""
        lowered = label.lower()

        # Step 4 clicked "Upload New Applications", and that trigger is usually
        # still in the DOM behind the dialog. Matching it here would silently
        # reopen the dialog instead of uploading, so refuse rather than guess.
        if "new application" in lowered:
            self._fail(
                "click_upload_button",
                f"The 'upload_button' selector matched '{label}', which is the button that "
                f"opens the dialog, not the one that submits it. Point "
                f"'upload_dialog.upload_button' in {config.SELECTORS_PATH} at the dialog's "
                f"own submit button.",
            )

        # The dialog's footer also holds Close/Cancel. Clicking one of those
        # would discard the form while looking like a successful submit.
        glyph_only = bool(lowered) and lowered.strip("×✕✖x ") == ""
        if glyph_only or lowered in ("close", "cancel", "dismiss", "back"):
            self._fail(
                "click_upload_button",
                f"The 'upload_button' selector matched '{label or '(no label)'}', which "
                f"dismisses the dialog rather than submitting it. Point "
                f"'upload_dialog.upload_button' in {config.SELECTORS_PATH} at the dialog's "
                f"Save button.\n{self._describe_dialog_controls(scope)}",
            )

        if self.dry_run:
            # Everything up to this point is client-side: the dialog is filled
            # in but nothing has been sent to the portal. Stopping here is what
            # makes a dry run safe.
            print(f"  DRY RUN: found the Upload button{f' ({label})' if label else ''} "
                  f"but did NOT click it - no record was created.")
            return False

        try:
            button.click()
        except Exception as e:
            self._fail("click_upload_button", f"Could not click the Upload button: {e}", e)

        print(f"  Clicked Upload{f' ({label})' if label else ''}.")
        return True

    def _describe_dialog_controls(self, scope):
        """
        Lists the buttons and dropdowns actually present in the dialog, so a
        selector miss reports what the portal really offers.
        """
        try:
            controls = scope.locator(
                "button, input[type='submit'], input[type='button'], select, a[onclick]"
            )
            total = controls.count()
        except Exception as e:
            return f"(could not list the dialog's controls: {e})"

        described = []
        for index in range(min(total, 25)):
            try:
                described.append(controls.nth(index).evaluate(
                    """el => {
                        const bits = [el.tagName.toLowerCase()];
                        if (el.type) bits.push('type=' + el.type);
                        if (el.id) bits.push('id=' + el.id);
                        if (el.name) bits.push('name=' + el.name);
                        const cls = typeof el.className === 'string' ? el.className.trim() : '';
                        if (cls) bits.push('class=' + cls);
                        const text = (el.value || el.textContent || '').trim().slice(0, 40);
                        if (text) bits.push('text=' + JSON.stringify(text));
                        return bits.join(' ');
                    }"""
                ))
            except Exception:
                continue

        if not described:
            return "The dialog reported no buttons or dropdowns at all."
        listed = "\n  - ".join(described)
        more = f"\n  ... and {total - len(described)} more" if total > len(described) else ""
        return f"Controls actually present in the dialog:\n  - {listed}{more}"

    @staticmethod
    def _report_form_state(form_state):
        """Prints what the dialog is holding, as read back from the page."""
        print(f"  DRY RUN: attached file  = "
              f"{form_state.get('attached_file') or '(unreadable)'}")
        print(f"  DRY RUN: selected batch = "
              f"{form_state.get('selected_batch') or '(unreadable)'}")

    @staticmethod
    def _selected_option_label(dropdown):
        """The label a <select> currently displays, or "" if unreadable."""
        try:
            return (dropdown.evaluate(
                """el => {
                    if (!el.options || el.selectedIndex < 0) return '';
                    const o = el.options[el.selectedIndex];
                    return (o.label || o.textContent || '').trim();
                }"""
            ) or "").strip()
        except Exception:
            return ""

    def describe_filled_form(self, scope):
        """
        Reads back what the dialog actually holds: the attached file name and
        the selected batch. This is the evidence a dry run is judged on - it
        proves the form was filled correctly, without submitting it.
        """
        attached = batch = ""

        file_input = selectors.find(
            scope, "upload_dialog.file_input", require_visible=False, timeout_ms=0, optional=True
        )
        if file_input is not None:
            try:
                attached = file_input.evaluate(
                    "el => (el.files && el.files.length) ? el.files[0].name : ''"
                ) or ""
            except Exception:
                pass

        dropdown = selectors.find(
            scope, "upload_dialog.sip_batch_select", timeout_ms=0, optional=True
        )
        if dropdown is not None:
            batch = self._selected_option_label(dropdown)

        return {"attached_file": attached, "selected_batch": batch}

    # -- step 7: wait, confirm ---------------------------------------------

    def confirm_upload(self, file_path):
        # Scoped to the whole page rather than the dialog: a confirmation
        # banner is as likely to replace the dialog as to appear inside it.
        print(f"  Waiting {config.UPLOAD_WAIT_SECONDS}s for the upload to complete...")
        time.sleep(config.UPLOAD_WAIT_SECONDS)
        self._settle()

        error_text = self._visible_message("result.error")
        if error_text:
            self._fail("upload", f"The portal reported an error: {error_text}")

        success_text = self._visible_message("result.success")
        if success_text:
            return UploadResult(True, file_path, message=success_text)

        if config.REQUIRE_SUCCESS_CONFIRMATION:
            self._fail(
                "confirm_upload",
                "No success confirmation was found on the page after uploading. "
                "Either the upload failed silently, or 'result.success' in "
                f"{config.SELECTORS_PATH} does not match the portal's confirmation banner.",
            )

        return UploadResult(
            True,
            file_path,
            message=(
                "Upload submitted; no error was shown, but no success banner was "
                "recognised either. Confirm the entry on the portal, then add the "
                "portal's confirmation selector to 'result.success' in selectors.json."
            ),
        )

    # -- steps 3-7 for one file --------------------------------------------

    def upload_form(self, file_path, batch_label=None):
        """Runs steps 3-7 for a single form. Assumes login() already ran."""
        if not self.logged_in:
            self.login()

        # One session uploads many forms; a batch resolved for an earlier form
        # must not be reported against this one.
        self.selected_batch_label = ""

        verb = "DRY RUN: preparing" if self.dry_run else "Uploading"
        print(f"{verb} '{os.path.basename(file_path)}' {'in' if self.dry_run else 'to'} the SIP portal...")
        self.open_manage_application()
        scope = self.open_upload_dialog()
        self.choose_file(scope, file_path)
        self.select_sip_batch(scope, batch_label)

        if self.dry_run:
            state_before = self.describe_filled_form(scope)
            # Reported now rather than with the result: if locating the submit
            # button fails, this is exactly the evidence needed to diagnose it.
            self._report_form_state(state_before)
            self.click_upload_button(scope)  # locates and verifies, does not click
            shot = self.screenshot("dryrun_dialog_ready")
            return self._dry_run_result(file_path, batch_label, state_before, shot)

        self.click_upload_button(scope)
        return self.confirm_upload(file_path)

    def _dry_run_result(self, file_path, batch_label, form_state, screenshot_path):
        """Turns a completed dry run into a result, checking what was filled in."""
        # Compare against the option that was actually chosen where we know it,
        # so a label that differs only in padding isn't reported as a mismatch.
        expected_batch = self.selected_batch_label or batch_label or config.SIP_BATCH_LABEL
        attached = form_state.get("attached_file") or ""
        selected = form_state.get("selected_batch") or ""

        problems = []
        if attached and attached != os.path.basename(file_path):
            problems.append(f"the dialog holds '{attached}', not '{os.path.basename(file_path)}'")
        elif not attached:
            problems.append("could not read back an attached file from the dialog")
        if selected and expected_batch.strip().lower() not in selected.strip().lower():
            problems.append(f"batch reads '{selected}', expected '{expected_batch}'")
        elif not selected:
            problems.append("could not read back the selected SIP batch")

        if screenshot_path:
            print(f"  DRY RUN: screenshot     = {screenshot_path}")

        if problems:
            # Not raised as a failure: the flow completed, but the form was not
            # filled as intended, which is exactly what a rehearsal is for.
            message = "Dry run completed, but the form was not filled as expected - " + "; ".join(problems)
            print(f"  DRY RUN: WARNING - {message}")
            return UploadResult(
                False, file_path, message=message, screenshot_path=screenshot_path, dry_run=True
            )

        return UploadResult(
            True,
            file_path,
            message=(
                f"Dry run OK - logged in, navigated, attached '{attached}' and selected "
                f"'{selected}'. The Upload button was located but not clicked, so no record "
                f"was created."
            ),
            screenshot_path=screenshot_path,
            dry_run=True,
        )
