# Ausbildungsnachweis MCP Server

MCP server that turns Microsoft Graph calendar events, the
Ausbildungseinsatzplan Excel files and WebUntis school data into filled
weekly training-report (Ausbildungsnachweis) PDFs. Python port of the
original PowerShell/Excel workflow, writing directly into the official
AcroForm PDF template. Nothing personal is hardcoded - everything comes
from a local profile created by the `initialise` tool.

## Tools

| Tool | Purpose |
|------|---------|
| `initialise` | One-stop setup: profile, Einsatzplan check, logins |
| `generate_report` | End-to-end: fetch the week's data -> fill PDF |
| `fetch_events` | Fetch + process Graph events only (returns JSON) |
| `fetch_school_lessons` | Fetch the WebUntis timetable incl. Lehrstoff |
| `fill_report` | Fill the PDF template from data you provide |
| `inspect_template` | Show fields / day-column mapping of a template PDF |
| `login` | Microsoft sign-in (browser / device code) |
| `untis_login` | Validate + store WebUntis credentials |
| `setup_status` | Check what is configured and what is missing |
| `auth_status` | Show cached Microsoft account / pending login state |
| `logout` | Clear cached Microsoft tokens |

## Setup (fresh clone)

```bash
python3 -m venv .venv
.venv/bin/pip install -e .
```

Then run the `initialise` tool once via the MCP client with:

- `name` - trainee name exactly as in the Einsatzplan (also used on the PDF)
- `training_start` - first day of the Ausbildung; drives the Ausbildungsjahr
  and the report numbering (the week containing it = report 001)
- `output_dir` - folder where the generated `NNN_*.pdf` reports are saved
- `einsatzplan_dir` - folder containing the `Ausbildungseinsatzplan*.xlsx`
  files (the rotation source)
- `template_path` - template PDF (defaults to `template.pdf` in this folder)
- `submit_dir` - optional submission folder (e.g. a Teams/OneDrive
  "Eingereicht" folder); every generated report is also copied there
- optionally `untis_username` / `untis_password` / `untis_server` /
  `untis_school`

`initialise` validates the name against the Einsatzplan, checks the WebUntis
login and starts the Microsoft browser sign-in (auth-code flow on the
managed device satisfies Conditional Access; device-code fails with 530033).
`setup_status` shows what is still missing.

All local state lives **inside this folder** and is gitignored:
`.credentials.json` (profile + WebUntis credentials, chmod 600),
`.msal_cache.bin` (Microsoft refresh-token cache), `.venv/`, `template.pdf`.

## Data sources

- **Abteilung / rotation:** the `Ausbildungseinsatzplan*.xlsx` files (one
  sheet per Lehrjahr, one column per trainee, column A = week; empty cells
  = Berufsschule). All workbooks in the folder are merged, so adding next
  year's file is enough.
- **Column 1 (Betriebliche Tätigkeiten):** Graph `/me/calendarView`; events
  with the same subject are merged, durations summed as German decimal
  hours (`18,00`). Durations are clipped to the week; all-day events count
  7,6h per covered weekday; school-block events are excluded.
- **Column 2 (Unterweisungen / Schulungen):** calendar events tagged with
  the Outlook category **Onboarding**. If you want Unterweisungen or
  Schulungen to appear in column 2 of the report, mark those calendar
  events with the category `Onboarding` - everything untagged goes to
  column 1.
- **Column 3 (Themen des Berufsschulunterrichts):** WebUntis timetable via
  JSON-RPC plus the internal REST API for the per-lesson **Lehrstoff**.
  Strict source separation: Berufsschule weeks use only WebUntis (columns
  1+2 empty, weekdays marked Berufsschule); Betrieb weeks use only the
  Outlook calendar. Cancelled lessons (Entfall) are skipped, hours summed
  per subject, topics appended (`LF02: Cisco Projektwoche`). The timetable
  is queried via the student's personId, so yearly class changes need no
  reconfiguration.
- **Day statuses (Betrieb weeks):** Outlook work plan (work hours &
  locations, beta Graph API): `office` -> **Büro**, `remote` ->
  **Homeoffice**, `timeOff` -> **Abwesenheit**. Calendar events refine this
  (`showAs = oof` / all-day absences -> Abwesenheit). Weekends default to
  Abwesenheit.

## Conventions

- **Ausbildungsjahr** rolls over on each anniversary of `training_start`.
- **Report number** = weeks since the week containing `training_start` + 1;
  override the anchor with `AN_NUMBER_ANCHOR` ("YYYY-MM-DD=N"). Passing a
  number that does not match the week raises an error.
- Output files are named `NNN_Vorname_Nachname.pdf`.

## MCP client configuration

```json
{
  "mcpServers": {
    "ausbildungsnachweis": {
      "command": "/path/to/ausbildungsnachweis/.venv/bin/ausbildungsnachweis-mcp"
    }
  }
}
```

No env vars required - the profile covers everything. Env overrides exist
for all settings: `AN_NAME`, `AN_TRAINING_START`, `AN_OUTPUT_DIR`,
`AN_TEMPLATE_PATH`, `AN_EINSATZPLAN_DIR`, `AN_SUBMIT_DIR`,
`AN_NUMBER_ANCHOR`, `AN_UNTIS_SERVER`, `AN_UNTIS_SCHOOL`, `AN_UNTIS_USER`,
`AN_UNTIS_PASSWORD`, `AN_CLIENT_ID`, `AN_TENANT_ID`, `AN_CACHE_DIR`.

## License

Copyright (C) 2026 wohljan

GPL-3.0 - see [LICENSE](LICENSE). This program comes with ABSOLUTELY NO
WARRANTY; it is free software, and you are welcome to redistribute it
under the conditions of the GNU General Public License v3.
