"""Fill (and inspect) the weekly Ausbildungsnachweis AcroForm PDF.

The current template exposes these fields:

  Text fields
    Name1        - trainee name
    Number1      - report number
    Department1  - department
    Year1        - training year
    Date1        - week range, e.g. "29.06.2026 - 03.07.2026"
    Text1/2/3    - activity descriptions, one per line (CR separated)
    Hours1/2/3   - hours per activity, one per line (CR separated)

  Day radio-groups (Montag..Sonntag)
    Each is a radio group with four widgets, one per column:
    Büro / Homeoffice / Berufsschule/Hochschule / Abwesenheit.
    Export values differ between fields (e.g. day1..day4 vs 1/0/2/3),
    so the correct value is resolved from each widget's x-position.
"""

from __future__ import annotations

from dataclasses import dataclass

from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject

from .processing import STATUS_COLUMN_ORDER, WEEKDAY_FIELD_NAMES

# Line separator used inside multi-line text fields of the template.
LINE_SEP = "\r"


@dataclass
class ReportData:
    name: str
    number: str
    department: str
    year: str
    date_range: str
    activities: list[dict]          # [{"text": str, "hours": str}, ...] -> column 1
    day_statuses: dict[str, str]    # {"Montag": "Büro", ...}
    # Optional extra columns (Unterweisungen / Berufsschulunterricht).
    activities_col2: list[dict] | None = None
    activities_col3: list[dict] | None = None


# ---------------------------------------------------------------------------
# Inspection
# ---------------------------------------------------------------------------

def _day_widget_options(reader: PdfReader) -> dict[str, list[str]]:
    """Return, per day field, the export values ordered by column (left->right)."""
    groups: dict[str, list[tuple[float, str]]] = {}
    for page in reader.pages:
        for annot in page.get("/Annots") or []:
            obj = annot.get_object()
            parent = obj.get("/Parent")
            if not parent:
                continue
            name = parent.get_object().get("/T")
            if name not in WEEKDAY_FIELD_NAMES:
                continue
            ap = obj.get("/AP")
            if not ap or "/N" not in ap:
                continue
            on_states = [s for s in ap["/N"].keys() if s != "/Off"]
            if not on_states:
                continue
            rect = [float(x) for x in obj["/Rect"]]
            groups.setdefault(name, []).append((rect[0], on_states[0]))
    return {
        name: [val for _x, val in sorted(items)]
        for name, items in groups.items()
    }


def inspect_pdf(pdf_path: str) -> dict:
    """Return the fields, current values and day-column options of a template."""
    reader = PdfReader(pdf_path)
    fields = reader.get_fields() or {}

    text_fields = {}
    for name, f in fields.items():
        if f.get("/FT") == "/Btn":
            continue
        text_fields[name] = f.get("/V")

    day_options = _day_widget_options(reader)
    day_columns = {}
    for name, exports in day_options.items():
        day_columns[name] = {
            STATUS_COLUMN_ORDER[i]: exports[i]
            for i in range(min(len(exports), len(STATUS_COLUMN_ORDER)))
        }

    return {
        "path": pdf_path,
        "pages": len(reader.pages),
        "text_fields": text_fields,
        "day_fields": day_columns,
        "current_day_values": {
            name: fields[name].get("/V")
            for name in WEEKDAY_FIELD_NAMES
            if name in fields
        },
    }


# ---------------------------------------------------------------------------
# Filling
# ---------------------------------------------------------------------------

def _join(items: list[dict] | None) -> tuple[str, str]:
    """Turn a list of {text, hours} into two CR-separated column strings."""
    if not items:
        return "", ""
    texts = LINE_SEP.join(str(i.get("text", "")) for i in items)
    hours = LINE_SEP.join(str(i.get("hours", "")) for i in items)
    # Trailing separator matches the style produced by the template.
    return texts + LINE_SEP, hours + LINE_SEP


def _resolve_day_export(day_options: dict[str, list[str]], day: str, status: str) -> str | None:
    """Map a human status ('Homeoffice') to the widget export value for that day."""
    exports = day_options.get(day)
    if not exports:
        return None
    try:
        col = STATUS_COLUMN_ORDER.index(status)
    except ValueError:
        return None
    if col >= len(exports):
        return None
    return exports[col]


def fill_pdf(template_path: str, output_path: str, data: ReportData) -> dict:
    """Fill the template with ``data`` and write the result to ``output_path``."""
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)

    text1, hours1 = _join(data.activities)
    text2, hours2 = _join(data.activities_col2)
    text3, hours3 = _join(data.activities_col3)

    text_updates = {
        "Name1": data.name,
        "Number1": str(data.number),
        "Department1": data.department,
        "Year1": str(data.year),
        "Date1": data.date_range,
        "Text1": text1,
        "Hours1": hours1,
        "Text2": text2,
        "Hours2": hours2,
        "Text3": text3,
        "Hours3": hours3,
    }
    for page in writer.pages:
        writer.update_page_form_field_values(page, text_updates)

    # Resolve + apply the day radio groups.
    day_options = _day_widget_options(reader)
    resolved_days: dict[str, str] = {}
    for day, status in (data.day_statuses or {}).items():
        export = _resolve_day_export(day_options, day, status)
        if export:
            resolved_days[day] = export

    acro = writer._root_object["/AcroForm"]
    for f in acro["/Fields"]:
        obj = f.get_object()
        name = obj.get("/T")
        if name not in resolved_days:
            continue
        target = resolved_days[name]
        obj[NameObject("/V")] = NameObject(target)
        for kid in obj.get("/Kids", []):
            k = kid.get_object()
            ap = k.get("/AP")
            states = list(ap["/N"].keys()) if ap and "/N" in ap else []
            k[NameObject("/AS")] = NameObject(target if target in states else "/Off")

    # Ask viewers to (re)generate field appearances so values are visible.
    acro[NameObject("/NeedAppearances")] = BooleanObject(True)

    with open(output_path, "wb") as fh:
        writer.write(fh)

    return {
        "output_path": output_path,
        "text_fields_set": {k: v for k, v in text_updates.items() if v},
        "day_statuses_set": {d: data.day_statuses[d] for d in resolved_days},
    }
