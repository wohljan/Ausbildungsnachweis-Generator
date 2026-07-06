"""WebUntis access: school timetable + Lehrstoff for the report's column 3.

Uses the WebUntis JSON-RPC API (login, timetable) plus the internal REST API
of the web frontend (JWT via ``/api/token/new``, then
``/api/rest/view/v2/calendar-entry/detail``) to fetch the per-lesson
"Lehrstoff" (teaching content) that student accounts see in the web UI.

Everything is fail-soft: callers get empty results when WebUntis is not
configured or unreachable, so report generation keeps working.
"""

from __future__ import annotations

import asyncio
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

import httpx

from . import credentials
from .processing import format_hours_de

USERAGENT = "ausbildungsnachweis-mcp"

# Period codes that mean the lesson did not take place.
_SKIP_CODES = {"cancelled"}


class UntisError(RuntimeError):
    pass


@dataclass
class SchoolLesson:
    day: date
    subject: str
    subject_long: str
    start: datetime
    end: datetime
    minutes: int
    topic: str = ""
    cancelled: bool = False


@dataclass
class SchoolWeek:
    lessons: list[SchoolLesson] = field(default_factory=list)

    @property
    def school_days(self) -> list[date]:
        return sorted({l.day for l in self.lessons if not l.cancelled})

    @property
    def total_minutes(self) -> int:
        return sum(l.minutes for l in self.lessons if not l.cancelled)

    def rows(self) -> list[dict]:
        """Aggregate into report rows: one row per subject with topics + hours."""
        by_subject: dict[str, dict] = {}
        for lesson in self.lessons:
            if lesson.cancelled:
                continue
            entry = by_subject.setdefault(
                lesson.subject,
                {"minutes": 0, "topics": [], "long": lesson.subject_long,
                 "first": lesson.start},
            )
            entry["minutes"] += lesson.minutes
            entry["first"] = min(entry["first"], lesson.start)
            if lesson.topic and lesson.topic not in entry["topics"]:
                entry["topics"].append(lesson.topic)

        rows = []
        for subject, entry in by_subject.items():
            long_name = (entry["long"] or "").strip()
            # LF01..LF05 all carry the generic long name "Lernfeld" - prefer
            # the code there; real names (Deutsch/Kommunikation, ...) win.
            display = long_name if long_name and long_name.lower() != "lernfeld" else subject
            if display in ("", "?"):
                # Periods without a subject (e.g. exam blocks): topic only.
                display = ""
            if entry["topics"]:
                topics = "; ".join(entry["topics"])
                text = f"{display}: {topics}" if display else topics
            else:
                text = display or "Unterricht"
            rows.append({
                "text": text,
                "hours": format_hours_de(entry["minutes"] / 60),
                "_minutes": entry["minutes"],
                "_first": entry["first"],
            })
        rows.sort(key=lambda r: (-r["_minutes"], r["_first"]))
        for r in rows:
            r.pop("_minutes"), r.pop("_first")
        return rows


def _parse_period_dt(day_int: int, time_int: int) -> datetime:
    return datetime(
        day_int // 10000, day_int // 100 % 100, day_int % 100,
        time_int // 100, time_int % 100,
    )


async def _authenticate(client: httpx.AsyncClient, cfg: dict) -> dict:
    resp = await client.post(
        f"https://{cfg['server']}/WebUntis/jsonrpc.do",
        params={"school": cfg["school"]},
        json={
            "id": "auth", "method": "authenticate", "jsonrpc": "2.0",
            "params": {
                "user": cfg["username"], "password": cfg["password"],
                "client": USERAGENT,
            },
        },
    )
    resp.raise_for_status()
    data = resp.json()
    if "error" in data or "result" not in data or "personId" not in data["result"]:
        err = (data.get("error") or {}).get("message", "unknown error")
        raise UntisError(f"WebUntis login failed: {err}")
    return data["result"]


async def verify_login(cfg: dict) -> dict:
    """Validate credentials; returns login info or raises.

    Besides personId/klasseId the result includes ``klasse`` - the *current*
    class name resolved live (e.g. "ITA1"), which changes automatically
    when the school moves the student to a new class.
    """
    async with httpx.AsyncClient(timeout=30, headers={"User-Agent": USERAGENT}) as client:
        result = await _authenticate(client, cfg)
        result["klasse"] = await _resolve_klasse_name(
            client, cfg, result.get("klasseId")
        )
        await _logout(client, cfg)
        return result


async def _resolve_klasse_name(
    client: httpx.AsyncClient, cfg: dict, klasse_id: int | None
) -> str | None:
    """Map the klasseId from the login response to its name (best-effort)."""
    if not klasse_id:
        return None
    try:
        resp = await client.post(
            f"https://{cfg['server']}/WebUntis/jsonrpc.do",
            params={"school": cfg["school"]},
            json={"id": "kl", "method": "getKlassen", "params": {}, "jsonrpc": "2.0"},
        )
        for klasse in resp.json().get("result") or []:
            if klasse.get("id") == klasse_id:
                return klasse.get("name")
    except Exception:
        pass
    return None


async def _logout(client: httpx.AsyncClient, cfg: dict) -> None:
    try:
        await client.post(
            f"https://{cfg['server']}/WebUntis/jsonrpc.do",
            params={"school": cfg["school"]},
            json={"id": "out", "method": "logout", "params": {}, "jsonrpc": "2.0"},
        )
    except Exception:
        pass


async def _fetch_topic(
    client: httpx.AsyncClient,
    cfg: dict,
    bearer: dict,
    person_id: int,
    person_type: int,
    lesson: SchoolLesson,
    period_id: int,
    sem: asyncio.Semaphore,
) -> None:
    """Fetch the Lehrstoff for one period (best-effort)."""
    async with sem:
        try:
            resp = await client.get(
                f"https://{cfg['server']}/WebUntis/api/rest/view/v2/calendar-entry/detail",
                params={
                    "elementId": person_id,
                    "elementType": person_type,
                    "periodId": period_id,
                    "startDateTime": lesson.start.strftime("%Y-%m-%dT%H:%M:%S"),
                    "endDateTime": lesson.end.strftime("%Y-%m-%dT%H:%M:%S"),
                },
                headers=bearer,
            )
            if resp.status_code != 200:
                return
            for entry in resp.json().get("calendarEntries") or []:
                topic = (entry.get("teachingContent") or "").strip()
                if topic:
                    lesson.topic = topic
                    return
        except Exception:
            return


async def fetch_school_week(
    start_date: date,
    end_date: date,
    cfg: dict | None = None,
    *,
    with_topics: bool = True,
) -> SchoolWeek:
    """Fetch the school timetable (and Lehrstoff) for a date range.

    Raises UntisError when credentials are missing or the login fails;
    network errors bubble up as httpx exceptions.
    """
    cfg = cfg or credentials.get_untis_config()
    if not all(cfg.get(k) for k in ("server", "school", "username", "password")):
        raise UntisError(
            "WebUntis is not configured - run the 'initialise' or "
            "'untis_login' tool."
        )

    week = SchoolWeek()
    async with httpx.AsyncClient(timeout=30, headers={"User-Agent": USERAGENT}) as client:
        auth = await _authenticate(client, cfg)
        person_id, person_type = auth["personId"], auth["personType"]

        resp = await client.post(
            f"https://{cfg['server']}/WebUntis/jsonrpc.do",
            params={"school": cfg["school"]},
            json={
                "id": "tt", "method": "getTimetable", "jsonrpc": "2.0",
                "params": {"options": {
                    "element": {"id": person_id, "type": person_type},
                    "startDate": int(start_date.strftime("%Y%m%d")),
                    "endDate": int(end_date.strftime("%Y%m%d")),
                    "subjectFields": ["name", "longname"],
                }},
            },
        )
        resp.raise_for_status()
        periods = resp.json().get("result") or []

        lessons: list[tuple[SchoolLesson, int]] = []
        for p in periods:
            subjects = p.get("su") or []
            subject = (subjects[0].get("name") if subjects else "") or "?"
            subject_long = (subjects[0].get("longname") if subjects else "") or ""
            start = _parse_period_dt(p["date"], p["startTime"])
            end = _parse_period_dt(p["date"], p["endTime"])
            lesson = SchoolLesson(
                day=start.date(),
                subject=subject,
                subject_long=subject_long,
                start=start,
                end=end,
                minutes=round((end - start).total_seconds() / 60),
                cancelled=(p.get("code") or "").lower() in _SKIP_CODES,
            )
            lessons.append((lesson, p.get("id", 0)))

        lessons.sort(key=lambda item: item[0].start)
        week.lessons = [l for l, _pid in lessons]

        if with_topics and week.lessons:
            try:
                token = (
                    await client.get(f"https://{cfg['server']}/WebUntis/api/token/new")
                ).text.strip()
                bearer = {"Authorization": f"Bearer {token}"}
                sem = asyncio.Semaphore(5)
                await asyncio.gather(*[
                    _fetch_topic(
                        client, cfg, bearer, person_id, person_type,
                        lesson, pid, sem,
                    )
                    for lesson, pid in lessons
                    if not lesson.cancelled and pid
                ])
            except Exception:
                pass  # topics are best-effort

        await _logout(client, cfg)

    return week
