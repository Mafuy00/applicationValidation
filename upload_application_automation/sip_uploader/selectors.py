"""
Loads `config/selectors.json` and resolves a logical element name (for
example "login.username") to a Playwright locator.

Each logical name maps to a list of candidate selectors. All candidates are
polled in order on every pass until one matches or the timeout expires, so a
slow-rendering page and a slightly-different-markup page are both handled,
and a failure can report exactly what was tried.
"""
import json
import time

from . import config

_cache = {}


class SelectorNotFoundError(RuntimeError):
    """No configured candidate selector matched within the timeout."""


class SelectorConfigError(RuntimeError):
    """selectors.json is missing, malformed, or missing a requested key."""


def load_selectors(path=None, refresh=False):
    path = path or config.SELECTORS_PATH
    if refresh:
        _cache.pop(path, None)
    if path in _cache:
        return _cache[path]

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError as e:
        raise SelectorConfigError(f"Selector config not found: {path}") from e
    except json.JSONDecodeError as e:
        raise SelectorConfigError(f"Selector config is not valid JSON ({path}): {e}") from e

    _cache[path] = data
    return data


def candidates_for(key_path, path=None):
    """
    Returns the candidate selector list for a dotted key such as
    "upload_dialog.file_input".
    """
    node = load_selectors(path)
    for part in key_path.split("."):
        if not isinstance(node, dict) or part not in node:
            raise SelectorConfigError(
                f"'{key_path}' is not defined in {path or config.SELECTORS_PATH}"
            )
        node = node[part]

    if isinstance(node, str):
        return [node]
    if isinstance(node, list) and all(isinstance(s, str) for s in node):
        return list(node)
    raise SelectorConfigError(
        f"'{key_path}' must be a selector string or a list of selector strings."
    )


def _matches(locator, require_visible):
    try:
        if locator.count() == 0:
            return False
        return locator.first.is_visible() if require_visible else True
    except Exception:
        # A locator can transiently throw while the page is navigating.
        return False


def find(
    scope,
    key_path,
    require_visible=True,
    timeout_ms=None,
    optional=False,
    poll_interval_ms=250,
):
    """
    Returns the first matching locator for `key_path`.

    scope: a Playwright Page, Frame or Locator - anything with .locator().
    require_visible: file inputs are frequently hidden behind a styled label,
        so visibility is opt-out rather than assumed.
    optional: return None instead of raising when nothing matches.
    """
    timeout_ms = config.PAGE_TIMEOUT_MS if timeout_ms is None else timeout_ms
    selectors = candidates_for(key_path)
    deadline = time.monotonic() + (timeout_ms / 1000.0)

    while True:
        for selector in selectors:
            try:
                locator = scope.locator(selector)
            except Exception:
                continue  # invalid selector - reported below if nothing matches
            if _matches(locator, require_visible):
                return locator.first
        if time.monotonic() >= deadline:
            break
        time.sleep(poll_interval_ms / 1000.0)

    if optional:
        return None

    where = ""
    url = getattr(scope, "url", None)
    if url:
        where = f" on {url}"
    raise SelectorNotFoundError(
        f"Could not find '{key_path}'{where} after {timeout_ms} ms.\n"
        f"Tried: {selectors}\n"
        f"Fix: add the portal's real selector to '{key_path}' in "
        f"{config.SELECTORS_PATH} (run tools/inspect_page.py to discover it)."
    )


def first_visible_text(scope, key_path, timeout_ms=0):
    """
    The collapsed visible text of the first element matching `key_path`, or ""
    if nothing matches or the match is empty.

    Used for outcome detection: a status container usually exists in the DOM
    from the start and only becomes visible with text once there is something
    to report, so presence alone is not evidence.
    """
    locator = find(scope, key_path, require_visible=True, timeout_ms=timeout_ms, optional=True)
    if locator is None:
        return ""
    try:
        return " ".join((locator.inner_text() or "").split())
    except Exception:
        return ""
