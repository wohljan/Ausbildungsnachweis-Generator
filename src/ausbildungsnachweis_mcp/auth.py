"""Microsoft Graph token acquisition via MSAL.

Uses a public client (default: the pre-consented "Microsoft Graph Command
Line Tools" app) with a persistent token cache: after one interactive
browser login (or device-code fallback) the server refreshes tokens
silently for months. The browser flow runs on the Windows side so that
device-based Conditional Access policies are satisfied.
"""

from __future__ import annotations

import base64
import json
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
# Calendars for events/work plan, Files.ReadWrite for the SharePoint upload
# of finished reports (user-consentable, no admin approval needed).
SCOPES = ["Calendars.Read", "Files.ReadWrite"]

# Project root (the ausbildungsnachweis/ folder): everything - token cache,
# credentials - lives in one place, excluded from git via .gitignore.
PROJECT_DIR = Path(__file__).resolve().parents[2]

CACHE_DIR = Path(os.environ.get("AN_CACHE_DIR", str(PROJECT_DIR)))
CACHE_FILE = CACHE_DIR / ".msal_cache.bin"
# Manual-token fallback (a bearer token pasted from Graph Explorer). Valid
# for about an hour; refreshed by re-pasting. Kept separate from the MSAL
# cache because these tokens have no refresh capability.
RAW_TOKEN_FILE = CACHE_DIR / ".raw_token.json"

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
    """Return an access token (raw fallback first, then MSAL cache) or None."""
    raw = _load_raw_token()
    if raw:
        return raw
    app = _get_app()
    accounts = app.get_accounts()
    if not accounts:
        return None
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        return result["access_token"]
    return None


# ---------------------------------------------------------------------------
# Manual-token fallback (Graph Explorer paste, hourly)
# ---------------------------------------------------------------------------

# 60s safety margin so we don't hand out a token that will die mid-request.
_RAW_TOKEN_LEEWAY = 60


def _decode_jwt_expiry(token: str) -> int:
    """Return the exp claim (unix seconds) of a JWT, or 0 if unparseable."""
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        data = json.loads(base64.urlsafe_b64decode(payload))
        return int(data.get("exp") or 0)
    except Exception:
        return 0


def _load_raw_token() -> str | None:
    if not RAW_TOKEN_FILE.exists():
        return None
    try:
        data = json.loads(RAW_TOKEN_FILE.read_text())
    except (OSError, json.JSONDecodeError):
        return None
    token = data.get("access_token")
    expires_at = int(data.get("expires_at") or 0)
    if not token or expires_at - _RAW_TOKEN_LEEWAY < time.time():
        return None
    return token


def store_raw_token(access_token: str) -> dict:
    """Persist a bearer token from Graph Explorer for silent use until expiry."""
    token = access_token.strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    if not token:
        raise ValueError("access_token is empty.")

    expires_at = _decode_jwt_expiry(token)
    if not expires_at:
        # Fall back to a conservative 55-minute lifetime if the JWT can't be
        # decoded (defensive - Graph tokens are always JWTs in practice).
        expires_at = int(time.time()) + 55 * 60

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    RAW_TOKEN_FILE.write_text(
        json.dumps({"access_token": token, "expires_at": expires_at})
    )
    os.chmod(RAW_TOKEN_FILE, 0o600)
    return {
        "expires_at": expires_at,
        "expires_in_minutes": max(0, (expires_at - int(time.time())) // 60),
    }


def clear_raw_token() -> bool:
    if RAW_TOKEN_FILE.exists():
        RAW_TOKEN_FILE.unlink()
        return True
    return False


def get_token() -> str:
    """Return a token or raise with instructions to run the login tool."""
    token = get_token_silent()
    if token:
        return token
    # If a previous interactive login was blocked by the tenant, surface that.
    if _pending.get("admin_consent_required"):
        raise PermissionError(
            "No usable Microsoft Graph token. "
            + (_pending.get("help") or _admin_consent_help())
        )
    raise PermissionError(
        "No cached Microsoft Graph credentials. Run the 'login' tool "
        "(method='interactive'), or paste a Graph Explorer token via "
        "login(method='token', access_token='...') for a one-hour fallback."
    )


# Error codes that indicate the tenant blocks user consent for our app.
# AADSTS65001 = user/admin consent required; AADSTS900971 similar.
_ADMIN_CONSENT_ERRORS = ("AADSTS65001", "AADSTS900971", "consent_required")


def _admin_consent_help() -> str:
    """Human-readable instructions for requesting tenant admin consent."""
    client_id = os.environ.get("AN_CLIENT_ID", DEFAULT_CLIENT_ID)
    tenant = os.environ.get("AN_TENANT_ID", DEFAULT_TENANT)
    scopes = " ".join(SCOPES)
    admin_consent_url = (
        f"https://login.microsoftonline.com/{tenant}/adminconsent"
        f"?client_id={client_id}"
    )
    return (
        "The tenant requires admin approval for this app - ask your IT admin "
        f"to grant tenant-wide consent for:\n"
        f"  App name : Microsoft Graph Command Line Tools\n"
        f"  Client ID: {client_id}\n"
        f"  Scopes   : {scopes}\n"
        f"  Consent  : {admin_consent_url}\n"
        "As an interim, use login(method='token', access_token='<token from "
        "Graph Explorer>') to unblock report generation for one hour."
    )


def _looks_like_consent_error(text: str) -> bool:
    if not text:
        return False
    low = text.lower()
    return any(m.lower() in low for m in _ADMIN_CONSENT_ERRORS)


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
            if _looks_like_consent_error(str(exc)):
                _pending["admin_consent_required"] = True
                _pending["help"] = _admin_consent_help()
            return
        if result and "access_token" in result:
            _pending["status"] = "completed"
            _pending["account"] = (result.get("id_token_claims") or {}).get(
                "preferred_username"
            )
        else:
            _pending["status"] = "failed"
            err = (result or {}).get("error")
            desc = (result or {}).get("error_description") or ""
            _pending["error"] = f"{err}: {desc}"
            if _looks_like_consent_error(f"{err} {desc}"):
                _pending["admin_consent_required"] = True
                _pending["help"] = _admin_consent_help()

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
            err = (result or {}).get("error")
            desc = (result or {}).get("error_description") or ""
            _pending["error"] = f"{err}: {desc}"
            if _looks_like_consent_error(f"{err} {desc}"):
                _pending["admin_consent_required"] = True
                _pending["help"] = _admin_consent_help()

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

    # Manual-token fallback state.
    raw_expiry = 0
    if RAW_TOKEN_FILE.exists():
        try:
            raw_expiry = int(json.loads(RAW_TOKEN_FILE.read_text()).get("expires_at") or 0)
        except (OSError, json.JSONDecodeError):
            raw_expiry = 0
    if raw_expiry:
        status["manual_token"] = {
            "present": True,
            "expires_in_minutes": max(0, (raw_expiry - int(time.time())) // 60),
            "expired": raw_expiry - _RAW_TOKEN_LEEWAY < time.time(),
        }

    if _pending:
        status["pending_login"] = {
            k: _pending[k]
            for k in (
                "status", "method", "user_code", "verification_uri",
                "account", "error", "admin_consent_required", "help",
            )
            if k in _pending
        }
    return status


def logout() -> dict:
    """Remove all cached accounts/tokens (MSAL cache + manual token)."""
    app = _get_app()
    removed = []
    for account in app.get_accounts():
        app.remove_account(account)
        removed.append(account.get("username"))
    raw_cleared = clear_raw_token()
    return {"removed_accounts": removed, "manual_token_cleared": raw_cleared}
