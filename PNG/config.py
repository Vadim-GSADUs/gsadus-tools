"""Persist last-used tab settings to config.json next to this file."""
import json
from pathlib import Path

_CONFIG_PATH = Path(__file__).parent / "config.json"


def load() -> dict:
    try:
        return json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save(data: dict) -> None:
    _CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
