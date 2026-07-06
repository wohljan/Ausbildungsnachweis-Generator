"""MCP server for generating weekly training reports (Ausbildungsnachweis).

Fills the official AcroForm PDF template from three data sources: Microsoft
Graph calendar events (column 1 + Onboarding-tagged events for column 2),
the Outlook work plan (per-day Büro/Homeoffice/Abwesenheit statuses) and
the WebUntis school timetable incl. Lehrstoff (column 3).

Tools
-----
initialise           - one-stop setup: profile, Einsatzplan check, logins
generate_report      - end-to-end: fetch week -> fill PDF
fetch_events         - Graph calendar events + day statuses (JSON)
fetch_school_lessons - WebUntis timetable + Lehrstoff (JSON)
fill_report          - fill the PDF from data you provide
inspect_template     - list the fields / day options of a template PDF
login / logout / auth_status - Microsoft sign-in (cached refresh token)
untis_login          - validate + store WebUntis credentials
setup_status         - what is configured, what is missing
"""

from __future__ import annotations

import os
import re
import shutil
from datetime import date, timedelta

from mcp.server.fastmcp import FastMCP

from . import auth, credentials, einsatzplan, graph, pdffill, schedule, untis
from .processing import (
    WEEKDAY_FIELD_NAMES,
    ProcessedEvents,
    format_hours_de,
    parse_flexible_date,
    process_events,
    to_graph_iso,
)

mcp = FastMCP("ausbildungsnachweis")

DEFAULT_TIMEZONE = "W. Europe Standard Time"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _profile_value(explicit: str, key: str, hint: str) -> str:
    """Resolve a setting: explicit argument > env var > profile; else error."""
    if explicit:
        return explicit
    value = credentials.get_profile()[key]
    if not value:
        raise ValueError(f"No {hint} configured - run the 'initialise' tool first.")
    return value

def _next_report_number(output_dir: str) -> int:
    """Return the next report number based on NNN_*.pdf files already present."""
    highest = 0
    if output_dir and os.path.isdir(output_dir):
        for fn in os.listdir(output_dir):
            m = re.match(r"^(\d{1,4})[_ ]", fn)
            if m and fn.lower().endswith(".pdf"):
                highest = max(highest, int(m.group(1)))
    return highest + 1


def _output_filename(number: int | str, name: str) -> str:
    safe = re.sub(r"\s+", "_", name.strip())
    try:
        num = f"{int(number):03d}"
    except (TypeError, ValueError):
        num = str(number)
    return f"{num}_{safe}.pdf"


def _resolve_token(access_token: str) -> str:
    """Use the provided token, or fall back to the MSAL cache."""
    return access_token if access_token else auth.get_token()


async def _collect_week(
    access_token: str, start: date, end: date, timezone: str
) -> tuple[ProcessedEvents, dict[date, str], int]:
    """Fetch calendar events + work locations and process them.

    Returns (processed, work_locations, events_fetched). The Graph UTC
    window is widened by a day on each side so local-midnight events on the
    first day are not clipped; process_events filters back to the exact
    local date range.
    """
    token = _resolve_token(access_token)
    raw = await graph.fetch_calendar_events(
        token,
        to_graph_iso(start - timedelta(days=1)),
        to_graph_iso(end + timedelta(days=1), end_of_day=True),
        timezone=timezone,
    )
    work_locations = await graph.fetch_work_locations(
        token, start, end, timezone=timezone
    )
    return process_events(raw, start, end, work_locations), work_locations, len(raw)


def _derive_year_and_department(
    week_start: date, year: str, department: str
) -> tuple[str, str]:
    """Fill in Ausbildungsjahr / Abteilung from the schedule when omitted."""
    if not year:
        year = str(schedule.training_year(week_start))
    if not department:
        dept, warning = schedule.department_for_week(week_start)
        if dept is None:
            raise ValueError(
                warning or "Department for this week is unknown; pass 'department'."
            )
        department = dept
    return year, department


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def inspect_template(pdf_path: str) -> dict:
    """Inspect an Ausbildungsnachweis PDF template.

    Returns the text fields (with their current values), the day radio-groups
    and which export value maps to each column
    (Büro / Homeoffice / Berufsschule / Abwesenheit).

    Args:
        pdf_path: Absolute path to the template (or an already filled) PDF.
    """
    return pdffill.inspect_pdf(pdf_path)


@mcp.tool()
def login(method: str = "interactive") -> dict:
    """Start a Microsoft sign-in for Graph access.

    method="interactive" (default): opens the sign-in page in your Windows
    browser (auth-code flow). The authentication runs on the managed device,
    which satisfies device-based Conditional Access policies - use this if
    the device-code flow fails with error 530033.

    method="device": classic device-code flow - returns a URL and a code to
    enter manually.

    The resulting refresh token is cached locally, so subsequent calls work
    without any token for months. Check progress with 'auth_status'.
    """
    if method == "device":
        return auth.start_device_login()
    return auth.start_interactive_login()


@mcp.tool()
def auth_status() -> dict:
    """Show Microsoft Graph auth state: cached accounts, whether a silent
    token is available, and the progress of a pending device-code login."""
    return auth.auth_status()


@mcp.tool()
def logout() -> dict:
    """Remove all cached Microsoft Graph accounts and tokens."""
    return auth.logout()


@mcp.tool()
async def untis_login(
    username: str,
    password: str,
    server: str = "",
    school: str = "",
) -> dict:
    """Save WebUntis credentials after validating them with a live login.

    The credentials are stored in the repo-local .credentials.json
    (gitignored, chmod 600). Nothing is saved if the login fails.

    Args:
        username: WebUntis username; typically the first 6 letters of the
            surname + first 2 of the first name (Max Mustermann -> "musterma").
        password: WebUntis password.
        server: WebUntis server; defaults to the stored value.
        school: WebUntis school login name; defaults likewise.
    """
    cfg = credentials.get_untis_config()
    if server:
        cfg["server"] = server
    if school:
        cfg["school"] = school
    if not cfg["server"] or not cfg["school"]:
        raise ValueError(
            "server and school are required on first setup (e.g. "
            "server='myschool.webuntis.com', school='myschool')."
        )
    cfg["username"] = username
    cfg["password"] = password

    info = await untis.verify_login(cfg)

    credentials.update("untis", {
        "server": cfg["server"],
        "school": cfg["school"],
        "username": username,
        "password": password,
    })
    return {
        "status": "ok",
        "server": cfg["server"],
        "school": cfg["school"],
        "person_id": info.get("personId"),
        "klasse": info.get("klasse"),
        "stored_in": str(credentials.CREDENTIALS_FILE),
    }


@mcp.tool()
async def initialise(
    name: str,
    training_start: str,
    output_dir: str,
    einsatzplan_dir: str,
    template_path: str = "",
    submit_dir: str = "",
    untis_username: str = "",
    untis_password: str = "",
    untis_server: str = "",
    untis_school: str = "",
) -> dict:
    """One-stop setup for a fresh clone: saves the profile and validates it.

    Stores everything personal in the repo-local .credentials.json
    (gitignored, chmod 600): trainee name, training start date (drives both
    the Ausbildungsjahr and the report numbering - the week containing it
    becomes report 001), output/Einsatzplan/template paths and optionally
    the WebUntis credentials.

    Validates the name against the Ausbildungseinsatzplan Excel files (the
    trainee's column must exist) and reports the covered weeks and
    departments. Also starts the Microsoft browser login if no cached
    account exists yet.

    Args:
        name: Trainee name exactly as in the Einsatzplan, e.g. "Max Muster".
        training_start: First day of the Ausbildung (dd.mm.yyyy or ISO).
        output_dir: Folder where the generated NNN_*.pdf reports are saved.
        einsatzplan_dir: Folder containing the Ausbildungseinsatzplan*.xlsx
            files (the rotation source).
        template_path: Template PDF; defaults to template.pdf in the repo
            folder if present.
        submit_dir: Optional submission folder (e.g. a Teams/OneDrive
            "Eingereicht" folder) - every generated report is copied there
            as well.
        untis_username / untis_password: Optional WebUntis credentials
            (validated live before saving).
        untis_server / untis_school: WebUntis server + school login name
            (required together with the credentials on first setup).
    """
    start = parse_flexible_date(training_start)

    if not os.path.isdir(output_dir):
        raise ValueError(f"output_dir does not exist: {output_dir}")
    if not os.path.isdir(einsatzplan_dir):
        raise ValueError(f"einsatzplan_dir does not exist: {einsatzplan_dir}")
    if submit_dir and not os.path.isdir(submit_dir):
        raise ValueError(f"submit_dir does not exist: {submit_dir}")

    if not template_path:
        candidate = credentials.PROJECT_DIR / "template.pdf"
        template_path = str(candidate) if candidate.is_file() else ""

    credentials.update("profile", {
        "name": name.strip(),
        "training_start": start.isoformat(),
        "output_dir": output_dir,
        "template_path": template_path,
        "einsatzplan_dir": einsatzplan_dir,
        "submit_dir": submit_dir,
    })

    result: dict = {
        "profile_stored_in": str(credentials.CREDENTIALS_FILE),
        "name": name.strip(),
        "training_start": start.isoformat(),
    }

    # Validate the name against the Einsatzplan and summarise the rotation.
    plan = einsatzplan.load_rotation(name.strip())
    if plan:
        result["einsatzplan"] = {
            "ok": True,
            "weeks": len(plan),
            "from": min(plan).isoformat(),
            "to": max(plan).isoformat(),
            "departments": sorted(set(plan.values())),
        }
    else:
        result["einsatzplan"] = {
            "ok": False,
            "error": (
                f"No column for '{name}' found in any Lehrjahr sheet under "
                f"{einsatzplan.plan_dir()} - department lookup will fall "
                "back to rotation.json."
            ),
        }

    # Optional WebUntis setup (validated live; nothing saved on failure).
    if untis_username and untis_password:
        cfg = credentials.get_untis_config()
        cfg.update({
            "username": untis_username,
            "password": untis_password,
            "server": untis_server or cfg["server"],
            "school": untis_school or cfg["school"],
        })
        if not cfg["server"] or not cfg["school"]:
            raise ValueError("untis_server and untis_school are required on first setup.")
        info = await untis.verify_login(cfg)
        credentials.update("untis", {
            "server": cfg["server"],
            "school": cfg["school"],
            "username": untis_username,
            "password": untis_password,
        })
        result["webuntis"] = {"ok": True, "klasse": info.get("klasse")}

    # Microsoft login: start the browser flow only if nothing is cached.
    if auth.get_token_silent():
        result["microsoft"] = {"ok": True, "note": "cached login reused"}
    else:
        result["microsoft"] = auth.start_interactive_login()

    return result


@mcp.tool()
async def setup_status() -> dict:
    """Check the complete setup: profile, Einsatzplan, Microsoft login,
    WebUntis login, template and output directory. Tells you exactly what
    is still missing."""
    profile = credentials.get_profile()
    profile_ok = bool(profile["name"] and profile["training_start"])

    ms = auth.auth_status()
    ms_ok = bool(ms.get("silent_token_available"))

    untis_cfg = credentials.get_untis_config()
    untis_ok = False
    untis_error = None
    untis_klasse = None
    if untis_cfg["username"] and untis_cfg["password"]:
        try:
            info = await untis.verify_login(untis_cfg)
            untis_ok = True
            untis_klasse = info.get("klasse")
        except Exception as exc:  # noqa: BLE001
            untis_error = str(exc)
    else:
        untis_error = "no credentials stored - run 'initialise' or 'untis_login'"

    plan = einsatzplan.load_rotation(profile["name"]) if profile["name"] else {}
    template_ok = bool(profile["template_path"]) and os.path.isfile(profile["template_path"])
    output_ok = bool(profile["output_dir"]) and os.path.isdir(profile["output_dir"])

    missing = []
    if not profile_ok:
        missing.append("Profile: run the 'initialise' tool")
    if not ms_ok:
        missing.append("Microsoft: run the 'login' tool")
    if not untis_ok:
        missing.append("WebUntis: run 'initialise' or 'untis_login'")
    if not template_ok:
        missing.append("Template: pass template_path to 'initialise'")
    if not output_ok:
        missing.append("Output dir: pass output_dir to 'initialise'")

    return {
        "profile": {
            "ok": profile_ok,
            "name": profile["name"] or None,
            "training_start": profile["training_start"] or None,
        },
        "einsatzplan": {
            "ok": bool(plan),
            "weeks": len(plan),
            "from": min(plan).isoformat() if plan else None,
            "to": max(plan).isoformat() if plan else None,
            "source_dir": str(einsatzplan.plan_dir()) or None,
        },
        "microsoft": {"ok": ms_ok, "accounts": ms.get("accounts")},
        "webuntis": {
            "ok": untis_ok,
            "server": untis_cfg["server"] or None,
            "school": untis_cfg["school"] or None,
            "username": untis_cfg["username"] or None,
            "klasse": untis_klasse,
            "error": untis_error,
        },
        "template": {"ok": template_ok, "path": profile["template_path"] or None},
        "output_dir": {"ok": output_ok, "path": profile["output_dir"] or None},
        "submit_dir": {
            "configured": bool(profile["submit_dir"]),
            "ok": not profile["submit_dir"] or os.path.isdir(profile["submit_dir"]),
            "path": profile["submit_dir"] or None,
        },
        "ready": not missing,
        "missing": missing,
    }


@mcp.tool()
async def fetch_school_lessons(start_date: str, end_date: str) -> dict:
    """Fetch the WebUntis school timetable (with Lehrstoff) for a date range.

    Returns per-subject rows for column 3 of the report (Themen des
    Berufsschulunterrichts) plus the individual lessons.

    Args:
        start_date: Range start (dd/mm/yy, dd.mm.yyyy or yyyy-mm-dd).
        end_date: Range end, inclusive.
    """
    start = parse_flexible_date(start_date)
    end = parse_flexible_date(end_date)
    week = await untis.fetch_school_week(start, end)
    return {
        "rows": week.rows(),
        "school_days": [d.isoformat() for d in week.school_days],
        "lessons": [
            {
                "date": l.day.isoformat(),
                "subject": l.subject,
                "subject_long": l.subject_long,
                "start": l.start.strftime("%H:%M"),
                "end": l.end.strftime("%H:%M"),
                "minutes": l.minutes,
                "topic": l.topic or None,
                "cancelled": l.cancelled,
            }
            for l in week.lessons
        ],
    }


@mcp.tool()
async def fetch_events(
    start_date: str,
    end_date: str,
    access_token: str = "",
    timezone: str = DEFAULT_TIMEZONE,
) -> dict:
    """Fetch and process Microsoft Graph calendar events for a week.

    Fetches ``/me/calendarView`` (with paging), computes each event's
    duration, merges events that share the same subject, and infers each
    weekday's status (Büro / Homeoffice / Berufsschule / Abwesenheit).

    Day statuses come from the Outlook work plan (work hours & locations:
    office -> Büro, remote -> Homeoffice, timeOff -> Abwesenheit), refined
    by calendar events (showAs = oof or all-day absences -> Abwesenheit,
    school events -> Berufsschule). Events tagged with the Outlook category
    "Onboarding" are returned separately as ``unterweisungen`` (column 2).

    Args:
        start_date: Week start, e.g. 06/07/26, 06.07.2026 or 2026-07-06.
        end_date: Week end (inclusive), same accepted formats.
        access_token: Optional Graph Bearer token; if omitted, the cached
            MSAL login is used (run the 'login' tool once).
        timezone: Outlook timezone name for localised times.

    Returns:
        A dict with the merged ``activities``, per-day ``day_statuses`` and
        totals, ready to hand to ``fill_report``.
    """
    start = parse_flexible_date(start_date)
    end = parse_flexible_date(end_date)

    processed, work_locations, events_fetched = await _collect_week(
        access_token, start, end, timezone
    )
    result = processed.as_dict()
    result["events_fetched"] = events_fetched
    result["work_locations"] = {
        d.isoformat(): loc for d, loc in sorted(work_locations.items())
    }
    return result


@mcp.tool()
def fill_report(
    activities: list[dict],
    day_statuses: dict[str, str],
    date_range: str,
    number: str | int,
    template_path: str = "",
    output_dir: str = "",
    output_path: str = "",
    name: str = "",
    department: str = "",
    year: str = "",
    unterweisungen: list[dict] | None = None,
    berufsschule: list[dict] | None = None,
) -> dict:
    """Fill the training-report PDF template with the supplied data.

    Args:
        activities: List of {"text": str, "hours": "18,00"} rows (column 1,
            Betriebliche Tätigkeiten).
        day_statuses: Map of German weekday -> status, e.g.
            {"Montag": "Homeoffice", "Samstag": "Abwesenheit"}.
        date_range: Week text for the form, e.g. "06.07.2026 - 10.07.2026".
        number: Report number (used in the field and the file name).
        template_path: Template PDF path (defaults to AN_TEMPLATE_PATH env).
        output_dir: Directory for the output file (defaults to AN_OUTPUT_DIR).
        output_path: Explicit output path; overrides output_dir/name.
        name: Trainee name (defaults to AN_NAME env).
        department: Abteilung; auto-derived from the rotation schedule via
            the first date in date_range when omitted.
        year: Ausbildungsjahr; auto-derived from date_range when omitted.
        unterweisungen: Rows for column 2 (Unterweisungen / Schulungen).
        berufsschule: Rows for column 3 (Themen des Berufsschulunterrichts).

    Returns:
        Details of the fields that were set and the output path.
    """
    template = _profile_value(template_path, "template_path", "template path")
    name = _profile_value(name, "name", "trainee name")

    if not department or not year:
        week_start = parse_flexible_date(date_range.split("-")[0].strip())
        year, department = _derive_year_and_department(week_start, year, department)

    if not output_path:
        out_dir = _profile_value(output_dir, "output_dir", "output directory")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, _output_filename(number, name))

    data = pdffill.ReportData(
        name=name,
        number=str(number),
        department=department,
        year=str(year),
        date_range=date_range,
        activities=activities,
        day_statuses=day_statuses,
        activities_col2=unterweisungen,
        activities_col3=berufsschule,
    )
    return pdffill.fill_pdf(template, output_path, data)


@mcp.tool()
async def generate_report(
    start_date: str,
    end_date: str,
    access_token: str = "",
    number: str | int = "",
    template_path: str = "",
    output_dir: str = "",
    name: str = "",
    department: str = "",
    year: str = "",
    timezone: str = DEFAULT_TIMEZONE,
    day_status_overrides: dict[str, str] | None = None,
) -> dict:
    """End-to-end: fetch the week's data and write the filled PDF.

    Strict source separation based on the rotation schedule: Berufsschule
    weeks use only WebUntis (column 3, weekdays marked Berufsschule);
    Betrieb weeks use only the Outlook calendar + work plan (column 1,
    Onboarding-tagged events in column 2, day statuses from work
    locations). The report number is derived from the week
    (001 = week of 01.08.2025) when omitted.

    Args:
        start_date: Week start (dd/mm/yy, dd.mm.yyyy or yyyy-mm-dd).
        end_date: Week end, inclusive.
        access_token: Optional Graph Bearer token; if omitted, the cached
            MSAL login is used (run the 'login' tool once).
        number: Report number; derived from the week when empty
            (001 = week of 01.08.2025). If given and it does not match the
            week, the tool raises an error.
        template_path: Template PDF path (defaults to AN_TEMPLATE_PATH).
        output_dir: Output directory (defaults to AN_OUTPUT_DIR).
        name: Trainee name (defaults to AN_NAME env).
        department: Abteilung; auto-derived from the rotation schedule
            (rotation.json) when omitted. Berufsschule blocks are part of the
            schedule; unknown weeks raise an error.
        year: Ausbildungsjahr; auto-derived from the week when omitted
            (01.08.2025-31.07.2026 -> 1, 01.08.2026-31.07.2027 -> 2, ...).
        timezone: Outlook timezone name.
        day_status_overrides: Optional manual overrides, e.g.
            {"Freitag": "Berufsschule"}.

    Returns:
        A dict combining the processed event data and the fill result.
    """
    template = _profile_value(template_path, "template_path", "template path")
    out_dir = _profile_value(output_dir, "output_dir", "output directory")
    name = _profile_value(name, "name", "trainee name")

    start = parse_flexible_date(start_date)
    end = parse_flexible_date(end_date)
    date_range = f"{start:%d.%m.%Y} - {end:%d.%m.%Y}"

    # Ausbildungsjahr + Abteilung from the schedule unless given explicitly.
    year, department = _derive_year_and_department(start, year, department)

    # Strict source separation: Berufsschule weeks use only WebUntis,
    # Betrieb weeks use only the Outlook calendar + work plan.
    school_rows: list[dict] = []
    untis_note = None
    activities: list[dict] = []
    unterweisungen: list[dict] = []
    work_locations: dict = {}
    events_fetched = 0

    if department == "Berufsschule":
        total_hours = 0.0
        try:
            school_week = await untis.fetch_school_week(start, end)
            school_rows = school_week.rows()
            total_hours = round(school_week.total_minutes / 60, 2)
            if not school_rows:
                untis_note = (
                    "Rotation says Berufsschule, but WebUntis returned no lessons "
                    "- timetable not published yet, school year rolled over, or "
                    "wrong week. Column 3 is empty; fill it manually if needed."
                )
        except untis.UntisError as exc:
            untis_note = f"Berufsschule week, but WebUntis failed: {exc}"
        except Exception as exc:  # noqa: BLE001
            untis_note = f"WebUntis unavailable: {exc}"

        day_statuses = {
            day: ("Abwesenheit" if day in ("Samstag", "Sonntag") else "Berufsschule")
            for day in WEEKDAY_FIELD_NAMES
        }
    else:
        processed, work_locations, events_fetched = await _collect_week(
            access_token, start, end, timezone
        )
        activities = [a.as_row() for a in processed.activities]
        unterweisungen = [a.as_row() for a in processed.unterweisungen]
        total_hours = processed.total_hours

        day_statuses = {d.day_name: d.status for d in processed.day_infos}
        # Weekend days outside the fetched range default to Abwesenheit.
        for weekend_day in ("Samstag", "Sonntag"):
            day_statuses.setdefault(weekend_day, "Abwesenheit")

    if day_status_overrides:
        day_statuses.update(day_status_overrides)

    # Number check: the report number must match the week
    # (001 = week of 01.08.2025 -> 002 = 04.08.2025, ..., 050 = 06.07.2026).
    expected_number = schedule.expected_report_number(start)
    number_note = None
    if number == "" or number is None:
        number = expected_number
        next_by_files = _next_report_number(out_dir)
        if next_by_files != expected_number:
            number_note = (
                f"Note: week-based number is {expected_number}, but the "
                f"files in the output directory suggest {next_by_files} - "
                "check for missing or duplicate reports."
            )
    elif int(number) != expected_number:
        raise ValueError(
            f"Report number {number} does not match the week "
            f"{date_range} (expected {expected_number}, "
            "001 = week of 01.08.2025). Pass the correct number or omit it."
        )

    os.makedirs(out_dir, exist_ok=True)
    output_path = os.path.join(out_dir, _output_filename(number, name))

    data = pdffill.ReportData(
        name=name,
        number=str(number),
        department=department,
        year=str(year),
        date_range=date_range,
        activities=activities,
        day_statuses=day_statuses,
        activities_col2=unterweisungen or None,
        activities_col3=school_rows or None,
    )
    fill_result = pdffill.fill_pdf(template, output_path, data)

    # Copy to the submission folder (e.g. Teams "Eingereicht") if configured.
    submitted_to = None
    submit_note = None
    submit_dir = credentials.get_profile()["submit_dir"]
    if submit_dir:
        try:
            submitted_to = shutil.copy2(
                output_path, os.path.join(submit_dir, os.path.basename(output_path))
            )
        except OSError as exc:
            submit_note = f"Copy to submit_dir failed: {exc}"

    return {
        "report_number": number,
        "expected_number_for_week": expected_number,
        "number_note": number_note,
        "untis_note": untis_note,
        "submit_note": submit_note,
        "date_range": date_range,
        "training_year": year,
        "department": department,
        "events_fetched": events_fetched,
        "activities": activities,
        "unterweisungen": unterweisungen,
        "berufsschule": school_rows,
        "total_hours": format_hours_de(total_hours),
        "day_statuses": day_statuses,
        "work_locations": {
            d.isoformat(): loc for d, loc in sorted(work_locations.items())
        },
        "output_path": output_path,
        "submitted_to": submitted_to,
        "fill_result": fill_result,
    }


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
