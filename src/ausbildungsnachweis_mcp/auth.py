"""Microsoft Graph token acquisition via MSAL.

Uses a public client (default: the pre-consented "Microsoft Graph Command
Line Tools" app) with a persistent token cache: after one interactive
browser login (or device-code fallback) the server refreshes tokens
silently for months. The browser flow runs on the Windows side so that
device-based Conditional Access policies are satisfied.
"""

from __future__ import annotations

import os
import subprocess
import threading
import time
import webbrowser
from pathlib import Path

import msal
from msal_extensions import (
    FilePersistence,
    PersistedTokenCache,
    build_encrypted_persistence,
)

# ---------------------------------------------------------------------------
# Browser bridge for WSL: open URLs in the *Windows* default browser so that
# the sign-in happens on the managed (domain-joined) device - required to
# satisfy device-based Conditional Access policies.
# ---------------------------------------------------------------------------

_POWERSHELL = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"


class _WindowsBrowser:
    """webbrowser controller that delegates to the Windows default browser."""

    def open(self, url: str, new: int = 0, autoraise: bool = True) -> bool:
        subprocess.Popen(
            [_POWERSHELL, "-NoProfile", "-Command", f'Start-Process "{url}"'],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True

    open_new = open
    open_new_tab = open


if os.path.exists(_POWERSHELL):
    try:
        webbrowser.register("windows-default", None, _WindowsBrowser(), preferred=True)
    except Exception:
        pass

# "Microsoft Graph Command Line Tools" (formerly Graph PowerShell) - a
# well-known public client that is pre-consented in many tenants.
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
DEFAULT_TENANT = "organizations"
SCOPES = ["Calendars.Read"]

# Project root (the ausbildungsnachweis/ folder): everything - token cache,
# credentials - lives in one place, excluded from git via .gitignore.
PROJECT_DIR = Path(__file__).resolve().parents[2]

CACHE_DIR = Path(os.environ.get("AN_CACHE_DIR", str(PROJECT_DIR)))
CACHE_FILE = CACHE_DIR / ".msal_cache.bin"

_lock = threading.Lock()
_app: msal.PublicClientApplication | None = None

# State of a pending device-code login (set by start_device_login).
_pending: dict = {}


def _build_cache() -> PersistedTokenCache:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        persistence = build_encrypted_persistence(str(CACHE_FILE))
        # Probe: encryption backends can fail lazily on first use.
        persistence.load()
    except Exception:
        persistence = FilePersistence(str(CACHE_FILE))
    return PersistedTokenCache(persistence)


def _get_app() -> msal.PublicClientApplication:
    global _app
    with _lock:
        if _app is None:
            client_id = os.environ.get("AN_CLIENT_ID", DEFAULT_CLIENT_ID)
            tenant = os.environ.get("AN_TENANT_ID", DEFAULT_TENANT)
            _app = msal.PublicClientApplication(
                client_id,
                authority=f"https://login.microsoftonline.com/{tenant}",
                token_cache=_build_cache(),
            )
        return _app


def get_token_silent() -> str | None:
    """Return an access token from the cache (refreshing if needed), or None."""
    app = _get_app()
    accounts = app.get_accounts()
    if not accounts:
        return None
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        return result["access_token"]
    return None


def get_token() -> str:
    """Return a token or raise with instructions to run the login tool."""
    token = get_token_silent()
    if token:
        return token
    raise PermissionError(
        "No cached Microsoft Graph credentials. Run the 'login' tool first, "
        "or pass an access_token argument explicitly."
    )


def start_interactive_login(timeout: int = 300) -> dict:
    """Begin an interactive (auth-code) login via the Windows browser.

    Unlike the device-code flow, the whole authentication happens inside the
    browser on the managed Windows device, so device-based Conditional
    Access policies (error 530033) are satisfied - same as Graph Explorer.

    Runs in a background thread; poll 'auth_status' for the result.
    """
    app = _get_app()

    _pending.clear()
    _pending.update({
        "status": "pending",
        "method": "interactive",
        "started_at": time.time(),
    })

    def _run() -> None:
        try:
            result = app.acquire_token_interactive(
                SCOPES,
                timeout=timeout,
                prompt="select_account",
                success_template=(
                    "<html><body><h3>Anmeldung erfolgreich.</h3>"
                    "Du kannst dieses Fenster schliessen.</body></html>"
                ),
            )
        except Exception as exc:  # noqa: BLE001
            _pending["status"] = "failed"
            _pending["error"] = str(exc)
            return
        if result and "access_token" in result:
            _pending["status"] = "completed"
            _pending["account"] = (result.get("id_token_claims") or {}).get(
                "preferred_username"
            )
        else:
            _pending["status"] = "failed"
            _pending["error"] = (
                f"{(result or {}).get('error')}: "
                f"{(result or {}).get('error_description')}"
            )

    threading.Thread(target=_run, daemon=True).start()

    return {
        "method": "interactive",
        "message": (
            "A sign-in window is opening in your Windows browser. Complete "
            "the login there, then check 'auth_status'."
        ),
    }


def start_device_login() -> dict:
    """Begin a device-code login; poll for completion in a background thread.

    Returns the verification URL + user code immediately. The background
    thread writes the token to the persistent cache once the user completes
    the browser step, so subsequent tool calls succeed silently.
    """
    app = _get_app()

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(
            "Failed to start device flow: "
            f"{flow.get('error')}: {flow.get('error_description')}"
        )

    _pending.clear()
    _pending.update({
        "status": "pending",
        "method": "device",
        "started_at": time.time(),
        "user_code": flow["user_code"],
        "verification_uri": flow.get("verification_uri")
        or flow.get("verification_url"),
        "expires_at": time.time() + flow.get("expires_in", 900),
    })

    def _poll() -> None:
        result = app.acquire_token_by_device_flow(flow)  # blocks until done
        if result and "access_token" in result:
            _pending["status"] = "completed"
            _pending["account"] = (result.get("id_token_claims") or {}).get(
                "preferred_username"
            )
        else:
            _pending["status"] = "failed"
            _pending["error"] = (
                f"{(result or {}).get('error')}: "
                f"{(result or {}).get('error_description')}"
            )

    threading.Thread(target=_poll, daemon=True).start()

    return {
        "verification_uri": _pending["verification_uri"],
        "user_code": _pending["user_code"],
        "expires_in_minutes": round((_pending["expires_at"] - time.time()) / 60),
        "message": (
            f"Open {_pending['verification_uri']} and enter code "
            f"{_pending['user_code']} to sign in. Then check 'auth_status'."
        ),
    }


def auth_status() -> dict:
    """Report cached account, silent-token health and any pending login."""
    app = _get_app()
    accounts = app.get_accounts()
    silent_ok = get_token_silent() is not None

    status: dict = {
        "client_id": os.environ.get("AN_CLIENT_ID", DEFAULT_CLIENT_ID),
        "tenant": os.environ.get("AN_TENANT_ID", DEFAULT_TENANT),
        "cache_file": str(CACHE_FILE),
        "accounts": [a.get("username") for a in accounts],
        "silent_token_available": silent_ok,
    }
    if _pending:
        status["pending_login"] = {
            k: _pending[k]
            for k in (
                "status", "method", "user_code", "verification_uri",
                "account", "error",
            )
            if k in _pending
        }
    return status


def logout() -> dict:
    """Remove all cached accounts/tokens."""
    app = _get_app()
    removed = []
    for account in app.get_accounts():
        app.remove_account(account)
        removed.append(account.get("username"))
    return {"removed_accounts": removed}
