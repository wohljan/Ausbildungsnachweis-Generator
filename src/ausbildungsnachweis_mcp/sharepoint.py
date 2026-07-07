"""Upload finished reports to a SharePoint folder via Microsoft Graph.

The submission target is a normal SharePoint folder URL (profile setting
``submit_url``). It is resolved once through the ``/shares`` endpoint to a
(driveId, itemId) pair, which is cached in the credential store; uploads
then PUT the file content directly. Reports are far below the 4 MB simple-
upload limit.

Requires the ``Files.ReadWrite`` delegated scope, which is part of the
standard login.
"""

from __future__ import annotations

import base64
import os
from pathlib import Path

import httpx

from . import credentials

GRAPH = "https://graph.microsoft.com/v1.0"


class SharePointError(RuntimeError):
    pass


def _share_id(url: str) -> str:
    return "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")


async def _resolve_folder(
    client: httpx.AsyncClient, token: str, url: str
) -> tuple[str, str]:
    """Resolve a SharePoint folder URL to (drive_id, item_id), with caching."""
    cached = credentials.load().get("sharepoint", {})
    if cached.get("url") == url and cached.get("drive_id") and cached.get("item_id"):
        return cached["drive_id"], cached["item_id"]

    resp = await client.get(
        f"{GRAPH}/shares/{_share_id(url)}/driveItem",
        headers={"Authorization": f"Bearer {token}"},
    )
    if resp.status_code != 200:
        raise SharePointError(
            f"Could not resolve the submit folder ({resp.status_code}): "
            f"{(resp.json().get('error') or {}).get('message', resp.text[:200])}"
        )
    item = resp.json()
    drive_id = (item.get("parentReference") or {}).get("driveId")
    item_id = item.get("id")
    if not drive_id or not item_id:
        raise SharePointError("Submit folder resolution returned no drive/item id.")
    credentials.update("sharepoint", {
        "url": url, "drive_id": drive_id, "item_id": item_id,
    })
    return drive_id, item_id


async def upload_report(file_path: str, token: str) -> str:
    """Upload a file into the configured submit folder; returns its webUrl."""
    url = credentials.get_profile()["submit_url"]
    if not url:
        raise SharePointError("No submit_url configured.")

    name = os.path.basename(file_path)
    content = Path(file_path).read_bytes()

    async with httpx.AsyncClient(timeout=60) as client:
        drive_id, item_id = await _resolve_folder(client, token, url)
        resp = await client.put(
            f"{GRAPH}/drives/{drive_id}/items/{item_id}:/{name}:/content",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/pdf",
            },
            content=content,
        )
        if resp.status_code == 404:
            # Cached ids may be stale (folder moved/recreated) - re-resolve once.
            credentials.update("sharepoint", {"drive_id": "", "item_id": ""})
            drive_id, item_id = await _resolve_folder(client, token, url)
            resp = await client.put(
                f"{GRAPH}/drives/{drive_id}/items/{item_id}:/{name}:/content",
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/pdf",
                },
                content=content,
            )
        if resp.status_code not in (200, 201):
            raise SharePointError(
                f"Upload failed ({resp.status_code}): "
                f"{(resp.json().get('error') or {}).get('message', resp.text[:200])}"
            )
        return resp.json().get("webUrl", "")
