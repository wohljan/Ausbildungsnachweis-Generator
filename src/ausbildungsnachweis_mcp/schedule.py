"""Training-year and department-rotation lookups.

* The Ausbildungsjahr is derived from the training start date (from the
  profile, set via the 'initialise' tool): the year rolls over every
  anniversary of the start.
* The report number is anchored at the week containing the training start
  (= report 001); override with AN_NUMBER_ANCHOR ("YYYY-MM-DD=N").
* The Abteilung for a given week comes from the Ausbildungseinsatzplan
  Excel files (see ``einsatzplan.py``). Weeks missing from the plan yield
  ``None`` plus a warning.
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


def _load_rotation() -> tuple[dict[date, str], str]:
    """Return (rotation, source) from the Ausbildungseinsatzplan files."""
    name = credentials.get_profile()["name"]
    if name:
        plan = einsatzplan.load_rotation(name)
        if plan:
            return plan, "Einsatzplan"
    return {}, "Einsatzplan"


def week_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())


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
