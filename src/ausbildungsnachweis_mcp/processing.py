"""Event processing logic ported from the original PowerShell script.

Takes raw Microsoft Graph calendar events, computes per-event durations,
merges events that share the same subject, and prepares the data used to
fill the weekly training report (Ausbildungsnachweis) PDF form.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Iterable

# Events tagged with one of these Outlook categories go into the second
# column of the report ("Unterweisungen bzw. überbetriebliche Unterweisungen,
# betrieblicher Unterricht, sonstige Schulungen") instead of column 1.
UNTERWEISUNG_CATEGORIES = ("onboarding",)

# Keyword groups used to infer the status of a given weekday.
ABSENCE_KEYWORDS = (
    "urlaub",
    "krank",
    "feiertag",
    "abwesend",
    "abwesenheit",
    "frei",
    "ganztägig frei",
)
SCHOOL_KEYWORDS = (
    "berufsschule",
    "berufschule",
    "berufskolleg",
    "schule",
    "hochschule",
    "unterricht",
    "blockunterricht",
)

# Hours credited per covered weekday for all-day events (Urlaub, Feiertag,
# Berufsschule blocks, ...).
ALL_DAY_MINUTES_PER_WEEKDAY = 480  # 8h

# The four day-status columns of the form, in x-position (left->right) order.
STATUS_BUERO = "Büro"
STATUS_HOMEOFFICE = "Homeoffice"
STATUS_BERUFSSCHULE = "Berufsschule"
STATUS_ABWESENHEIT = "Abwesenheit"

STATUS_COLUMN_ORDER = (
    STATUS_BUERO,
    STATUS_HOMEOFFICE,
    STATUS_BERUFSSCHULE,
    STATUS_ABWESENHEIT,
)

# German weekday names as used for the PDF radio-group field names.
WEEKDAY_FIELD_NAMES = (
    "Montag",
    "Dienstag",
    "Mittwoch",
    "Donnerstag",
    "Freitag",
    "Samstag",
    "Sonntag",
)


# ---------------------------------------------------------------------------
# Date / time parsing helpers
# ---------------------------------------------------------------------------

def parse_flexible_date(value: str) -> date:
    """Parse a date given as dd/mm/yy, dd.mm.yyyy, dd-mm-yyyy or ISO yyyy-mm-dd."""
    value = value.strip()
    fmts = (
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%d.%m.%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d-%m-%Y",
        "%d-%m-%y",
    )
    for fmt in fmts:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(
        f"Unrecognised date '{value}'. Use dd/mm/yy, dd.mm.yyyy or yyyy-mm-dd."
    )


def to_graph_iso(d: date, *, end_of_day: bool = False) -> str:
    """Format a date as the ISO timestamp expected by the Graph calendarView API."""
    if end_of_day:
        return f"{d:%Y-%m-%d}T23:59:59.999Z"
    return f"{d:%Y-%m-%d}T00:00:00.000Z"


def parse_graph_datetime(value: str | None) -> datetime | None:
    """Parse a Graph dateTime string (e.g. '2026-07-03T09:00:00.0000000')."""
    if not value:
        return None
    v = value.strip()
    if v.endswith("Z"):
        v = v[:-1]
    # Normalise fractional seconds to at most 6 digits for %f.
    m = re.match(r"^(\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2})(?:\.(\d+))?", v)
    if not m:
        try:
            return datetime.fromisoformat(v)
        except ValueError:
            return None
    base, frac = m.group(1), m.group(2)
    base = base.replace(" ", "T")
    if frac:
        frac = (frac + "000000")[:6]
        return datetime.strptime(f"{base}.{frac}", "%Y-%m-%dT%H:%M:%S.%f")
    return datetime.strptime(base, "%Y-%m-%dT%H:%M:%S")


def format_hours_de(hours: float) -> str:
    """Format a number of hours as German text with two decimals: 18.0 -> '18,00'."""
    return f"{hours:.2f}".replace(".", ",")


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Activity:
    """A merged calendar activity (one row of the report table)."""

    text: str
    minutes: int
    hours: float
    start: datetime
    end: datetime

    def hours_de(self) -> str:
        return format_hours_de(self.hours)

    def as_row(self) -> dict:
        """The {"text", "hours"} shape used to fill the PDF columns."""
        return {"text": self.text, "hours": self.hours_de()}

    def as_dict(self) -> dict:
        return {
            "text": self.text,
            "hours": self.hours_de(),
            "hours_value": self.hours,
            "minutes": self.minutes,
            "start": self.start.isoformat(),
            "end": self.end.isoformat(),
        }


@dataclass
class DayInfo:
    day_name: str
    day_date: date
    status: str
    subjects: list[str] = field(default_factory=list)
    work_location: str | None = None


@dataclass
class ProcessedEvents:
    activities: list[Activity]
    day_infos: list[DayInfo]
    start_date: date
    end_date: date
    # Column 2: Unterweisungen / Schulungen (e.g. events tagged "Onboarding").
    unterweisungen: list[Activity] = field(default_factory=list)

    @property
    def total_minutes(self) -> int:
        return sum(a.minutes for a in self.activities) + sum(
            a.minutes for a in self.unterweisungen
        )

    @property
    def total_hours(self) -> float:
        return round(self.total_minutes / 60, 2)

    def date_range_de(self) -> str:
        return f"{self.start_date:%d.%m.%Y} - {self.end_date:%d.%m.%Y}"

    def as_dict(self) -> dict:
        return {
            "date_range": self.date_range_de(),
            "start_date": self.start_date.isoformat(),
            "end_date": self.end_date.isoformat(),
            "total_hours": format_hours_de(self.total_hours),
            "total_minutes": self.total_minutes,
            "activities": [a.as_dict() for a in self.activities],
            "unterweisungen": [a.as_dict() for a in self.unterweisungen],
            "day_statuses": {d.day_name: d.status for d in self.day_infos},
            "days": [
                {
                    "day": d.day_name,
                    "date": d.day_date.isoformat(),
                    "status": d.status,
                    "work_location": d.work_location,
                    "subjects": d.subjects,
                }
                for d in self.day_infos
            ],
        }


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def _event_subject(ev: dict) -> str:
    return (ev.get("subject") or "").strip()


def _matches_any(text: str, keywords: Iterable[str]) -> bool:
    low = text.lower()
    return any(k in low for k in keywords)


def infer_day_status(
    events_of_day: list[dict],
    is_weekend: bool,
    work_location: str | None = None,
) -> tuple[str, list[str]]:
    """Determine the status column for a single weekday.

    Primary source is the Outlook work plan (``work_location``:
    office / remote / timeOff); calendar events refine it.

    Rules (in priority order):
      * timeOff in the work plan, an out-of-office event (showAs == 'oof')
        or an all-day absence event                       -> Abwesenheit
      * a Berufsschule / Hochschule event                 -> Berufsschule
      * work location 'remote'                            -> Homeoffice
      * work location 'office' or no data                 -> Büro
    Weekend days without events default to Abwesenheit.
    """
    subjects = [_event_subject(e) for e in events_of_day if _event_subject(e)]
    loc = (work_location or "").lower()

    has_absence = loc == "timeoff"
    has_school = False
    for ev in events_of_day:
        subject = _event_subject(ev)
        if _matches_any(subject, SCHOOL_KEYWORDS):
            # School events (often marked oof) mean Berufsschule, not absence.
            has_school = True
            continue
        show_as = (ev.get("showAs") or "").lower()
        is_all_day = bool(ev.get("isAllDay"))
        if show_as == "oof" or (is_all_day and _matches_any(subject, ABSENCE_KEYWORDS)):
            has_absence = True

    if has_absence:
        return STATUS_ABWESENHEIT, subjects
    if has_school:
        return STATUS_BERUFSSCHULE, subjects
    if loc == "remote":
        return STATUS_HOMEOFFICE, subjects
    if loc == "office":
        return STATUS_BUERO, subjects
    if is_weekend and not events_of_day:
        return STATUS_ABWESENHEIT, subjects
    return STATUS_BUERO, subjects


def _is_unterweisung(ev: dict) -> bool:
    """True if the event is tagged with an Unterweisung category (e.g. 'Onboarding')."""
    categories = ev.get("categories") or []
    return any((c or "").strip().lower() in UNTERWEISUNG_CATEGORIES for c in categories)


def _merge_by_subject(parsed: list[tuple]) -> list[Activity]:
    """Merge (subject, start, end, minutes) tuples that share the same subject."""
    merged: dict[str, dict] = {}
    for subject, start, end, minutes in parsed:
        m = merged.get(subject)
        if m is None:
            merged[subject] = {"minutes": minutes, "start": start, "end": end}
        else:
            m["minutes"] += minutes
            m["start"] = min(m["start"], start)
            m["end"] = max(m["end"], end)

    activities = [
        Activity(
            text=subject,
            minutes=data["minutes"],
            hours=round(data["minutes"] / 60, 2),
            start=data["start"],
            end=data["end"],
        )
        for subject, data in merged.items()
    ]
    # Sort by duration descending (matches the layout of recent reports).
    activities.sort(key=lambda a: a.minutes, reverse=True)
    return activities


def process_events(
    raw_events: list[dict],
    start_date: date,
    end_date: date,
    work_locations: dict[date, str] | None = None,
) -> ProcessedEvents:
    """Compute durations, merge by subject and infer per-day statuses.

    ``work_locations`` maps days to the Outlook work-plan location
    (office / remote / timeOff) and drives the Büro/Homeoffice/Abwesenheit
    columns; calendar events refine it (oof, Berufsschule).

    Durations are clipped to the requested date range (multi-week events
    only count the covered part). All-day events are credited with 8h per
    covered weekday. Events matching SCHOOL_KEYWORDS drive the Berufsschule
    day status but are not listed as activities - the school timetable
    (column 3) covers them.
    """
    work_locations = work_locations or {}
    window_start = datetime.combine(start_date, datetime.min.time())
    window_end = datetime.combine(end_date + timedelta(days=1), datetime.min.time())

    # ---- Per-event parsing + duration (ported from Try-ParseDateTime) ----
    parsed_col1 = []  # (subject, start, end, minutes) -> Betriebliche Tätigkeiten
    parsed_col2 = []  # events tagged 'Onboarding'      -> Unterweisungen
    events_by_day: dict[date, list[dict]] = {}

    for ev in raw_events:
        subject = _event_subject(ev)
        start = parse_graph_datetime((ev.get("start") or {}).get("dateTime"))
        end = parse_graph_datetime((ev.get("end") or {}).get("dateTime"))
        if start is None or end is None:
            continue
        if end < start:
            end = end + timedelta(days=1)

        # Keep only events overlapping the requested local date range (the
        # fetch window is widened by a day on each side because the Graph
        # calendarView filter works in UTC).
        if end <= window_start or start >= window_end:
            continue

        # Days covered by the event within the range (end is exclusive for
        # midnight-aligned events such as all-day entries).
        first_day = max(start.date(), start_date)
        last_moment = max(end - timedelta(microseconds=1), start)
        last_day = min(last_moment.date(), end_date)

        cursor = first_day
        while cursor <= last_day:
            events_by_day.setdefault(cursor, []).append(ev)
            cursor += timedelta(days=1)

        if not subject:
            continue
        # School-block events are markers for the day status only.
        if _matches_any(subject, SCHOOL_KEYWORDS):
            continue

        if ev.get("isAllDay"):
            weekdays = sum(
                1
                for i in range((last_day - first_day).days + 1)
                if (first_day + timedelta(days=i)).weekday() < 5
            )
            minutes = ALL_DAY_MINUTES_PER_WEEKDAY * weekdays
        else:
            overlap = min(end, window_end) - max(start, window_start)
            minutes = round(overlap.total_seconds() / 60)
        if minutes <= 0:
            continue

        if _is_unterweisung(ev):
            parsed_col2.append((subject, start, end, minutes))
        else:
            parsed_col1.append((subject, start, end, minutes))

    activities = _merge_by_subject(parsed_col1)
    unterweisungen = _merge_by_subject(parsed_col2)

    # ---- Per-day statuses for the whole Monday..Sunday range ----
    day_infos: list[DayInfo] = []
    cursor = start_date
    while cursor <= end_date:
        weekday_idx = cursor.weekday()  # 0 = Monday
        name = WEEKDAY_FIELD_NAMES[weekday_idx]
        is_weekend = weekday_idx >= 5
        location = work_locations.get(cursor)
        status, subjects = infer_day_status(
            events_by_day.get(cursor, []), is_weekend, location
        )
        day_infos.append(DayInfo(
            day_name=name,
            day_date=cursor,
            status=status,
            subjects=subjects,
            work_location=location,
        ))
        cursor += timedelta(days=1)

    return ProcessedEvents(
        activities=activities,
        day_infos=day_infos,
        start_date=start_date,
        end_date=end_date,
        unterweisungen=unterweisungen,
    )
