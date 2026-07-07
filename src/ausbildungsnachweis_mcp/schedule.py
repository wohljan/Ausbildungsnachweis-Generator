"""Training-year and department-rotation lookups.

* The Ausbildungsjahr is derived from the training start date (from the
  profile, set via the 'initialise' tool): the year rolls over every
  anniversary of the start.
* The report number is anchored at the week containing the training start
  (= report 001); override with AN_NUMBER_ANCHOR ("YYYY-MM-DD=N").
* The Abteilung for a given week comes from the Ausbildungseinsatzplan
  Excel files (see ``einsatzplan.py``), optionally overridden per week via
  the 'set_department' tool (stored locally, the xlsx stays untouched).
  Weeks missing from both yield ``None`` plus a warning.
"""

from __future__ import annotations

import os
from datetime import date, timedelta

from . import credentials, einsatzplan


def training_start() -> date:
    return credentials.require_training_start()


def training_year(d: date) -> int:
    """Return the Ausbildungsjahr (1-based) for a given date."""
    start = training_start()
    if d < start:
        return 1
    years = d.year - start.year
    if (d.month, d.day) < (start.month, start.day):
        years -= 1
    return years + 1


def week_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())


# ---------------------------------------------------------------------------
# Local per-week department overrides (never written into the xlsx)
# ---------------------------------------------------------------------------

def get_overrides() -> dict[date, str]:
    stored = credentials.load().get("rotation_overrides", {})
    out = {}
    for k, v in stored.items():
        try:
            out[date.fromisoformat(k)] = v
        except ValueError:
            continue
    return out


def set_override(start: date, end: date, department: str) -> list[date]:
    """Set (or clear, when department is empty) overrides for all weeks in range.

    Returns the affected week Mondays.
    """
    data = credentials.load()
    overrides = data.setdefault("rotation_overrides", {})
    affected = []
    monday = week_monday(start)
    last = week_monday(end)
    while monday <= last:
        key = monday.isoformat()
        if department:
            overrides[key] = department
        else:
            overrides.pop(key, None)
        affected.append(monday)
        monday += timedelta(days=7)
    credentials.save(data)
    return affected


def _load_rotation() -> tuple[dict[date, str], str]:
    """Return (rotation, source) from the Einsatzplan files + local overrides."""
    name = credentials.get_profile()["name"]
    plan = einsatzplan.load_rotation(name) if name else {}
    plan.update(get_overrides())
    return plan, "Einsatzplan"


def expected_report_number(d: date) -> int:
    """Return the report number that belongs to the week containing ``d``.

    Anchored at the week containing the training start date (= report 001);
    override with AN_NUMBER_ANCHOR ("YYYY-MM-DD=N", meaning the week of
    that date is report N).
    """
    raw = os.environ.get("AN_NUMBER_ANCHOR", "")
    if raw:
        date_part, num_part = raw.split("=")
        anchor_week, anchor_number = date.fromisoformat(date_part.strip()), int(num_part)
    else:
        anchor_week, anchor_number = training_start(), 1
    weeks = (week_monday(d) - week_monday(anchor_week)).days // 7
    return anchor_number + weeks


def department_for_week(week_start: date) -> tuple[str | None, str | None]:
    """Resolve the department for the week containing ``week_start``.

    Returns ``(department, warning)``. ``department`` is ``None`` when the
    week is missing from the rotation schedule.
    """
    rotation, source = _load_rotation()
    monday = week_monday(week_start)

    if monday in rotation:
        return rotation[monday], None

    if not rotation:
        return None, (
            "No rotation data found - check the Einsatzplan folder and the "
            "trainee name (run 'initialise'), or pass 'department' explicitly."
        )

    if monday < min(rotation):
        return None, (
            f"Week {monday:%d.%m.%Y} is before the first entry of the "
            f"{source} ({min(rotation):%d.%m.%Y}); department unknown."
        )

    if monday > max(rotation):
        return None, (
            f"Week {monday:%d.%m.%Y} is after the last entry of the "
            f"{source} ({max(rotation):%d.%m.%Y}); department unknown."
        )

    return None, (
        f"Week {monday:%d.%m.%Y} is missing from the {source}. "
        "Add it there or pass 'department' explicitly."
    )
