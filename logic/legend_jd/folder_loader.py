from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Tuple

from .models import (
    EquipoCatalogItem,
    LegendConfig,
    PlantillaBOMItem,
    VariadorItem,
    WcrItem,
)


def _safe_read_json(path: Path):
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def load_config(dir_path: Path) -> Tuple[LegendConfig | None, bool]:
    data = _safe_read_json(dir_path / "config.json")
    if not isinstance(data, dict):
        return None, False
    cfg = LegendConfig(
        proyecto=data.get("proyecto"),
        ciudad=data.get("ciudad"),
        tipo_sistema=data.get("tipo_sistema"),
        refrigerante=data.get("refrigerante"),
        tcond=data.get("tcond"),
        tevap_bt=data.get("tevap_bt"),
        tevap_mt=data.get("tevap_mt"),
        marca_compresores=data.get("marca_compresores"),
        tipo_instalacion=data.get("tipo_instalacion"),
        factor_seguridad=data.get("factor_seguridad"),
        extras={k: v for k, v in data.items() if k not in LegendConfig().__dict__},
    )
    return cfg, True


def load_equipos(dir_path: Path) -> Tuple[List[EquipoCatalogItem], bool]:
    data = _safe_read_json(dir_path / "equipos.json") or []
    items: List[EquipoCatalogItem] = []
    if isinstance(data, list):
        for d in data:
            try:
                items.append(EquipoCatalogItem(str(d.get("equipo", "")), float(d.get("btu_hr_ft", 0.0))))
            except Exception:
                continue
    return items, bool(data)


def load_usos(dir_path: Path) -> Tuple[Dict[str, List[str]], bool]:
    data = _safe_read_json(dir_path / "usos.json") or {}
    usos = {"BT": [], "MT": []}
    if isinstance(data, dict):
        for k in ("BT", "MT"):
            val = data.get(k, [])
            if isinstance(val, list):
                usos[k] = [str(x) for x in val]
    return usos, bool(data)


def load_variadores(dir_path: Path) -> Tuple[List[VariadorItem], bool]:
    data = _safe_read_json(dir_path / "variadores.json") or []
    items: List[VariadorItem] = []
    if isinstance(data, list):
        for d in data:
            try:
                items.append(VariadorItem(str(d.get("modelo", "")), float(d.get("potencia", 0.0))))
            except Exception:
                continue
    return items, bool(data)


def load_wcr(dir_path: Path) -> Tuple[List[WcrItem], bool]:
    data = _safe_read_json(dir_path / "wcr.json") or []
    items: List[WcrItem] = []
    if isinstance(data, list):
        for d in data:
            try:
                items.append(WcrItem(str(d.get("modelo", "")), float(d.get("capacidad", 0.0))))
            except Exception:
                continue
    return items, bool(data)


def load_plantillas(dir_path: Path) -> Tuple[Dict[str, List[PlantillaBOMItem]], bool]:
    data = _safe_read_json(dir_path / "plantillas.json") or {}
    names = ["Rack Loop", "Minisistema", "Rack Americano"]
    plantillas: Dict[str, List[PlantillaBOMItem]] = {n: [] for n in names}
    found = False
    if isinstance(data, dict):
        for name in names:
            lista = data.get(name, [])
            if isinstance(lista, list):
                found = True
                for item in lista:
                    if isinstance(item, dict):
                        qty = item.get("qty")
                        try:
                            qty_val = float(qty) if qty is not None else None
                        except Exception:
                            qty_val = None
                        desc = str(item.get("descripcion", ""))
                        plantillas[name].append(PlantillaBOMItem(qty=qty_val, descripcion=desc))
    return plantillas, found
