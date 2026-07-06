"""Microsoft Graph calendar access.

Thin async wrappers around the ``/me/calendarView`` endpoint (with the
paging behaviour of the original PowerShell script) and the beta
work-hours-and-locations API used to determine per-day Büro/Homeoffice/
Abwesenheit statuses.
"""

from __future__ import annotations

from datetime import date, timedelta

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"

# Fields we actually need for the report; keeps the response small.
_SELECT = "subject,start,end,isAllDay,showAs,categories"


async def fetch_calendar_events(
    access_token: str,
    start_iso: str,
    end_iso: str,
    timezone: str = "W. Europe Standard Time",
    *,
    timeout: float = 30.0,
) -> list[dict]:
    """Fetch all calendar events between two ISO timestamps.

    Parameters mirror the PowerShell script: a Bearer token, a start/end
    datetime and an Outlook timezone (sent via the ``Prefer`` header so that
    the returned ``dateTime`` values are already localised).
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "Prefer": f'outlook.timezone="{timezone}"',
    }
    params = {
        "startdatetime": start_iso,
        "enddatetime": end_iso,
        "$select": _SELECT,
        "$top": "100",
        "$orderby": "start/dateTime",
    }

    url = f"{GRAPH_BASE}/me/calendarView"
    events: list[dict] = []

    async with httpx.AsyncClient(timeout=timeout) as client:
        # First request uses query params; subsequent ones follow nextLink.
        next_url: str | None = None
        while True:
            if next_url:
                resp = await client.get(next_url, headers=headers)
            else:
                resp = await client.get(url, headers=headers, params=params)

            if resp.status_code == 401:
                raise PermissionError(
                    "Microsoft Graph returned 401 Unauthorized - the access "
                    "token is missing, expired or lacks Calendars.Read scope."
                )
            resp.raise_for_status()
            data = resp.json()

            value = data.get("value") or []
            events.extend(value)

            next_url = data.get("@odata.nextLink")
            if not next_url:
                break

    return events


# Priority when a day has several work-plan occurrences (e.g. a base
# "office" recurrence plus an explicit "timeOff" occurrence).
_WORK_LOCATION_RANK = {"timeoff": 3, "remote": 2, "office": 1}


async def fetch_work_locations(
    access_token: str,
    start_date: date,
    end_date: date,
    timezone: str = "W. Europe Standard Time",
    *,
    timeout: float = 30.0,
) -> dict[date, str]:
    """Fetch per-day work locations from the Outlook work plan.

    Uses the beta ``occurrencesView`` of ``/me/settings/workHoursAndLocations``
    (readable with Calendars.Read). Returns a mapping of day -> location type
    (``office`` / ``remote`` / ``timeOff``); when a day has multiple
    occurrences the most specific one wins (timeOff > remote > office).

    Fail-soft: any error (API unavailable, permission change, beta removal)
    yields an empty mapping so day statuses fall back to defaults.
    """
    url = (
        f"{GRAPH_BETA}/me/settings/workHoursAndLocations/"
        f"occurrencesView(startDateTime='{start_date:%Y-%m-%d}T00:00:00',"
        f"endDateTime='{end_date + timedelta(days=1):%Y-%m-%d}T00:00:00')"
    )
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "Prefer": f'outlook.timezone="{timezone}"',
    }

    try:
        async with httpx.AsyncClient(timeout=timeout) as client:
            resp = await client.get(url, headers=headers)
            resp.raise_for_status()
            data = resp.json()
    except Exception:
        return {}

    locations: dict[date, str] = {}
    for occ in data.get("value") or []:
        loc = (occ.get("workLocationType") or "").strip()
        start_str = ((occ.get("start") or {}).get("dateTime") or "")[:10]
        if not loc or not start_str:
            continue
        try:
            day = date.fromisoformat(start_str)
        except ValueError:
            continue
        rank = _WORK_LOCATION_RANK.get(loc.lower(), 0)
        if rank == 0:
            continue
        current = locations.get(day)
        if current is None or rank > _WORK_LOCATION_RANK.get(current.lower(), 0):
            locations[day] = loc
    return locations
