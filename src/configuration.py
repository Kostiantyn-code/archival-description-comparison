"""Read configuration and dictionary YAML with one validation policy."""
from __future__ import annotations

from pathlib import Path
from typing import Any


def safe_load_yaml(path: Path, yaml_module) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        data = yaml_module.safe_load(stream)
    if not isinstance(data, dict):
        raise ValueError(f"Корінь YAML має бути словником: {path}")
    return data
