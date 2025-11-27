# logic/programacion_loader.py
from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Tuple, Any


def _as_dict(x: Any) -> Dict[str, Any]:
    return x if isinstance(x, dict) else {}


def load_programacion_snapshot(path: str | Path) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """
    Lee un snapshot .ecalc.json y devuelve (step2, globs).

    Formatos aceptados (todos retornan lo mismo):
      1) Nuevo:
         {
           "format": "electrocalc.programacion.v1",
           "globs": {...},
           "step2": {...}
         }
      2) Alternativos / legados:
         { "globals": {...}, "step2_state": {...} }
         { "data": { "globs": {...}, "step2": {...} } }
         { "payload": { "globs": {...}, "step2": {...} } }
    """
    p = Path(path)
    with p.open("r", encoding="utf-8") as f:
        raw = json.load(f)

    # Caso 1: estructura nueva (recomendada)
    if isinstance(raw, dict) and ("globs" in raw or "step2" in raw):
        globs = _as_dict(raw.get("globs"))
        step2 = _as_dict(raw.get("step2"))
        return step2, globs

    # Caso 2: dentro de "data" o "payload"
    for key in ("data", "payload"):
        cont = _as_dict(raw.get(key)) if isinstance(raw, dict) else {}
        if cont:
            globs = _as_dict(cont.get("globs"))
            step2 = _as_dict(cont.get("step2")) or _as_dict(cont.get("step2_state"))
            if globs or step2:
                return step2, globs

    # Caso 3: nombres alternativos directos
    if isinstance(raw, dict):
        globs = _as_dict(raw.get("globals")) or _as_dict(raw.get("global")) or {}
        step2 = _as_dict(raw.get("step2_state")) or _as_dict(raw.get("step2")) or {}
        if globs or step2:
            return step2, globs

    # Fallback: nada reconocible → devolvemos vacíos
    return {}, {}

