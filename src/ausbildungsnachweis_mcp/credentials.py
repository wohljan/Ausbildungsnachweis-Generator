"""Repo-local configuration and credential store.

Everything personal lives in ``.credentials.json`` next to the project
(gitignored, chmod 600) - profile (name, training start, paths), WebUntis
credentials. A fresh clone runs the ``initialise`` tool once and is set up;
nothing personal is hardcoded in the repository.

Precedence when reading configuration: explicit arguments > environment
variables (``AN_*``) > credential file.
"""

from __future__ import annotations

import json
import os
from datetime import date
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[2]
CREDENTIALS_FILE = PROJECT_DIR / ".credentials.json"


def load() -> dict:
    if not CREDENTIALS_FILE.exists():
        return {}
    try:
        with open(CREDENTIALS_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return {}


def save(data: dict) -> None:
    CREDENTIALS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CREDENTIALS_FILE, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
    os.chmod(CREDENTIALS_FILE, 0o600)


def update(section: str, values: dict) -> None:
    data = load()
    data.setdefault(section, {}).update(values)
    save(data)


# ---------------------------------------------------------------------------
# Profile (set once via the 'initialise' tool)
# ---------------------------------------------------------------------------

def get_profile() -> dict:
    """Profile values with env-var overrides applied."""
    stored = load().get("profile", {})
    return {
        "name": os.environ.get("AN_NAME") or stored.get("name") or "",
        "training_start": os.environ.get("AN_TRAINING_START")
        or stored.get("training_start")
        or "",
        "output_dir": os.environ.get("AN_OUTPUT_DIR")
        or stored.get("output_dir")
        or "",
        "template_path": os.environ.get("AN_TEMPLATE_PATH")
        or stored.get("template_path")
        or "",
        "einsatzplan_dir": os.environ.get("AN_EINSATZPLAN_DIR")
        or stored.get("einsatzplan_dir")
        or "",
    }


def require_name() -> str:
    name = get_profile()["name"]
    if not name:
        raise ValueError(
            "No trainee name configured - run the 'initialise' tool first."
        )
    return name


def require_training_start() -> date:
    raw = get_profile()["training_start"]
    if not raw:
        raise ValueError(
            "No training start date configured - run the 'initialise' tool first."
        )
    return date.fromisoformat(raw)


# ---------------------------------------------------------------------------
# WebUntis
# ---------------------------------------------------------------------------

def get_untis_config() -> dict:
    """Resolve the WebUntis configuration (env vars override the file).

    No class name is needed: timetables are queried via the personId
    returned at login, which follows the student across class changes.
    """
    stored = load().get("untis", {})
    return {
        "server": os.environ.get("AN_UNTIS_SERVER") or stored.get("server") or "",
        "school": os.environ.get("AN_UNTIS_SCHOOL") or stored.get("school") or "",
        "username": os.environ.get("AN_UNTIS_USER") or stored.get("username") or "",
        "password": os.environ.get("AN_UNTIS_PASSWORD") or stored.get("password") or "",
    }
