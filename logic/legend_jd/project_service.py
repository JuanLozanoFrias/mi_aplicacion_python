from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any

DEFAULT_FILE = "legend_project.json"

DEFAULT_META = {
    "nombre_proyecto": "PROYECTO",
    "fecha": datetime.now().strftime("%Y-%m-%d"),
    "version": "1.0",
}

DEFAULT_SPECS = {
    "voltaje": "",
    "tevap": "",
    "tcond": "",
    "refrigerante": "",
    "controlador": "",
}

DEFAULT_ROW = {
    "loop": 1,
    "ramal": 1,
    "dim_m": "",
    "equipo": "",
    "uso": "",
    "carga_btu_h": 0.0,
    "tevap_f": 0.0,
    "evap_qty": 0,
    "evap_modelo": "",
    "control": "",
    "succion": "",
    "liquida": "",
    "direccion": "",
    "deshielo": "",
}


def _normalize_row(row: Dict[str, Any]) -> Dict[str, Any]:
    out = DEFAULT_ROW.copy()
    # Compat: permitir datos antiguos con clave "tevap"
    compat = dict(row)
    if "tevap_f" not in compat and "tevap" in compat:
        compat["tevap_f"] = compat.get("tevap")
    for k in out.keys():
        out[k] = compat.get(k, out[k])
    return out


class LegendProjectService:
    @staticmethod
    def _path(project_dir: Path) -> Path:
        return Path(project_dir) / DEFAULT_FILE

    @staticmethod
    def load(project_dir: Path) -> Dict[str, Any]:
        project_dir = Path(project_dir)
        if not project_dir.exists():
            project_dir.mkdir(parents=True, exist_ok=True)
        path = LegendProjectService._path(project_dir)
        if not path.exists():
            return {
                "meta": DEFAULT_META.copy(),
                "specs": DEFAULT_SPECS.copy(),
                "bt_items": [],
                "mt_items": [],
                "bt_ramales": 1,
                "mt_ramales": 1,
            }
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return {
                "meta": DEFAULT_META.copy(),
                "specs": DEFAULT_SPECS.copy(),
                "bt_items": [],
                "mt_items": [],
                "bt_ramales": 1,
                "mt_ramales": 1,
            }
        meta = DEFAULT_META.copy()
        meta.update(data.get("meta", {}))
        specs = DEFAULT_SPECS.copy()
        specs.update(data.get("specs", {}))
        bt_items = [_normalize_row(r) for r in data.get("bt_items", []) if isinstance(r, dict)]
        mt_items = [_normalize_row(r) for r in data.get("mt_items", []) if isinstance(r, dict)]
        bt_ramales = int(data.get("bt_ramales", 1) or 1)
        mt_ramales = int(data.get("mt_ramales", 1) or 1)
        return {
            "meta": meta,
            "specs": specs,
            "bt_items": bt_items,
            "mt_items": mt_items,
            "bt_ramales": bt_ramales,
            "mt_ramales": mt_ramales,
        }

    @staticmethod
    def save(project_dir: Path, payload: Dict[str, Any]) -> None:
        project_dir = Path(project_dir)
        project_dir.mkdir(parents=True, exist_ok=True)
        path = LegendProjectService._path(project_dir)
        # Normalizar antes de guardar
        meta = DEFAULT_META.copy()
        meta.update(payload.get("meta", {}))
        specs = DEFAULT_SPECS.copy()
        specs.update(payload.get("specs", {}))
        bt_items = [_normalize_row(r) for r in payload.get("bt_items", []) if isinstance(r, dict)]
        mt_items = [_normalize_row(r) for r in payload.get("mt_items", []) if isinstance(r, dict)]
        out = {
            "meta": meta,
            "specs": specs,
            "bt_items": bt_items,
            "mt_items": mt_items,
            "bt_ramales": int(payload.get("bt_ramales", 1) or 1),
            "mt_ramales": int(payload.get("mt_ramales", 1) or 1),
        }
        path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    import tempfile
    tmp = Path(tempfile.gettempdir()) / "legend_project_test"
    svc = LegendProjectService()
    data = svc.load(tmp)
    print("Loaded default:", data["meta"])
    svc.save(tmp, data)
    print("Saved to", svc._path(tmp))
