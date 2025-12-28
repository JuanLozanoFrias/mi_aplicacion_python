from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List

from .models import ValidationIssue

RULES_PATH = Path("data/cuartos_industriales/thermal_load_validation_rules.json")


def load_rules() -> Dict:
    if RULES_PATH.exists():
        return json.load(open(RULES_PATH, encoding="utf-8"))
    return {}


def validate_room(room: Dict, rules: Dict) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    limits = rules.get("limits", {})
    def add(msg, level="warning"):
        issues.append(ValidationIssue(level=level, message=msg))

    max_dim = limits.get("max_dim_m")
    min_dim = limits.get("min_dim_m")
    for key in ("largo_m", "ancho_m", "altura_m"):
        val = room.get(key, 0)
        if max_dim and val > max_dim:
            add(f"{key.upper()} supera {max_dim} m")
        if min_dim and val < min_dim:
            add(f"{key.upper()} es menor a {min_dim} m")
    return issues

