"""User preferences, kept in $XDG_CONFIG_HOME/tse/settings.json (~/.config/... by default)."""
from __future__ import annotations

import json
import os


def path() -> str:
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "tse", "settings.json")


def load() -> dict:
    try:
        with open(path(), encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def update(**values) -> None:
    """Merge values into the settings file. Failures to write are ignored (preferences are optional)."""
    data = load()
    data.update(values)
    p = path()
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        tmp = p + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
            f.write("\n")
        os.replace(tmp, p)
    except OSError:
        pass
