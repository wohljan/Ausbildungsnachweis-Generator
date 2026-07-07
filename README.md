# Ausbildungsnachweis MCP Server

MCP server that turns Microsoft Graph calendar events, the
Ausbildungseinsatzplan Excel files and WebUntis school data into filled
weekly Ausbildungsnachweis PDFs. Nothing personal is hardcoded, everything comes
from a local profile created by the `initialise` tool.

## Tools

| Tool | Purpose |
|------|---------|
| `initialise` | One-stop setup: profile, Einsatzplan check, logins |
| `generate_report` | End-to-end: fetch the week's data -> fill PDF |
| `fetch_events` | Fetch + process Graph events only (returns JSON) |
| `fetch_school_lessons` | Fetch the WebUntis timetable incl. Lehrstoff |
| `fetch_rotation` | Show the parsed rotation plan (week -> Abteilung) |
| `set_department` | Override the Abteilung for a time frame (stored locally) |
| `fill_report` | Fill the PDF template from data you provide |
| `inspect_template` | Show fields / day-column mapping of a template PDF |
| `login` | Microsoft sign-in (browser / device code) |
| `untis_login` | Validate + store WebUntis credentials |
| `setup_status` | Check what is configured and what is missing |
| `auth_status` | Show cached Microsoft account / pending login state |
| `logout` | Clear cached Microsoft tokens |

## Setup on a new device

```bash
git clone https://github.com/wohljan/Ausbildungsnachweis-Generator.git
cd Ausbildungsnachweis-Generator
python3 -m venv .venv
.\.venv\Scripts\pip install -e .
```

For privacy, portability, and environment-specific reasons, the following files and folders are **not** part of this repository:

1. **`template.pdf`**  
   The PDF template is excluded via `.gitignore`. Add a copy of any existing report PDF and rename it to `template.pdf`. The application overwrites all form fields when generating reports.

2. **Data Directories**  
   Output, Einsatzplan, and optional submission directories are not tracked. These locations must exist locally, either through OneDrive/Teams synchronization or accessible network shares. Path validation is performed during `initialise`.

3. **Ausbildungseinsatzplan**  
   The Ausbildungseinsatzplan is not included and must be provided separately. It is used to generate a CSV that enables:
   - Automatic Berufsschul week detection (Untis-only, Outlook check skipped)
   - Automatic department detection (Outlook-only, Untis check skipped)


Register the server in your MCP client (VS Code: Command Palette ->
">MCP: Open User Configuration"):

```json
{
  "servers": {
    "ausbildungsnachweis": {
      "type": "stdio",
      "command": "/path/to/Ausbildungsnachweis-Generator/.venv/Scripts/ausbildungsnachweis-mcp.exe"
    }
  }
}
```

Then run the `initialise` tool once via the MCP client with:

- `name` - trainee name exactly as in the Einsatzplan (also used on the PDF)
- `training_start` - first day of the Ausbildung; drives the Ausbildungsjahr
  and the report numbering (the week containing it = report 001)
- `output_dir` - folder where the generated `NNN_*.pdf` reports are saved
- `einsatzplan_dir` - folder containing the `Ausbildungseinsatzplan*.xlsx`
  files (the rotation source)
- `template_path` - template PDF (defaults to `template.pdf` in this folder)
- `submit_url` - optional SharePoint folder URL (e.g. the Teams
  "Eingereicht" library folder); every generated report is uploaded there
  directly via Microsoft Graph (`Files.ReadWrite` delegated scope, part of
  the standard login - no admin consent needed)
- optionally `untis_username` / `untis_password` / `untis_server` /
  `untis_school`

`initialise` validates the name against the Einsatzplan, checks the WebUntis
login and starts the Microsoft browser sign-in (auth-code flow on the
managed device satisfies Conditional Access; device-code fails with 530033).
`setup_status` shows what is still missing - once it reports `ready: true`,
report generation works.

All local state lives **inside this folder** and is gitignored:
`.credentials.json` (profile + WebUntis credentials, chmod 600),
`.msal_cache.bin` (Microsoft refresh-token cache), `.venv/`, `template.pdf`.

## Data sources

- **Abteilung / rotation:** the `Ausbildungseinsatzplan*.xlsx` files (one
  sheet per Lehrjahr, one column per trainee, column A = week; empty cells
  = Berufsschule). All workbooks in the folder are merged, so adding next
  year's file is enough. Inspect the parsed plan with `fetch_rotation`;
  deviations (e.g. a swapped week) are set with `set_department` for any
  time frame - overrides live in `.credentials.json`, the Excel files are
  never modified.
- **Column 1 (Betriebliche Tätigkeiten):** Graph `/me/calendarView`; events
  with the same subject are merged, durations summed as German decimal
  hours (`18,00`). Durations are clipped to the week; all-day events count
  8h per covered weekday; school-block events are excluded.
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

## Corporate tenant: Microsoft admin consent

Interactive login uses the pre-built Microsoft "Graph Command Line Tools"
app (client `14d82eec-204b-4c2f-b7e8-296a70dab67e`), which asks the user
for consent to `Calendars.Read` and `Files.ReadWrite`. Many corporate
tenants block user consent for third-party apps - the sign-in page then
shows "Need admin approval" and blocks the flow.

Two paths:

1. **Ask IT to grant tenant-wide admin consent** (one-time, permanent
   fix). Send them the admin-consent URL that `login` returns on failure;
   equivalent Azure Portal path: Enterprise Applications -> Microsoft
   Graph Command Line Tools -> Permissions -> Grant admin consent. After
   that any user can run `login` normally.
2. **Interim: paste a token from Graph Explorer** - visit
   <https://developer.microsoft.com/graph/graph-explorer>, sign in, open
   the *Access token* tab, copy the token, then run
   `login(method="token", access_token="<paste>")`. It lasts ~1h and the
   server uses it silently until expiry; re-paste to extend.

`auth_status` shows both states (cached MSAL account, manual-token
expiry) and includes the admin-consent instructions when the last login
was blocked.

## Configuration

No env vars required - the profile covers everything. Env overrides exist
for all settings: `AN_NAME`, `AN_TRAINING_START`, `AN_OUTPUT_DIR`,
`AN_TEMPLATE_PATH`, `AN_EINSATZPLAN_DIR`, `AN_SUBMIT_URL`,
`AN_NUMBER_ANCHOR`, `AN_UNTIS_SERVER`, `AN_UNTIS_SCHOOL`, `AN_UNTIS_USER`,
`AN_UNTIS_PASSWORD`, `AN_CLIENT_ID`, `AN_TENANT_ID`, `AN_CACHE_DIR`.

## License

Copyright (C) 2026 wohljan

GPL-3.0 - see [LICENSE](LICENSE). This program comes with ABSOLUTELY NO
WARRANTY; it is free software, and you are welcome to redistribute it
under the conditions of the GNU General Public License v3.
