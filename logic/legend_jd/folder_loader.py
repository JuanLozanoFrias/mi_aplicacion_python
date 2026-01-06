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


def _load_analysis_info(dir_path: Path) -> Dict[str, object]:
    data = _safe_read_json(dir_path / "legend_jd_analysis.json")
    if not isinstance(data, dict):
        return {}
    info = data.get("INFO")
    return info if isinstance(info, dict) else {}


def _split_equipos_by_row(rows: List[dict]) -> Dict[str, List[dict]]:
    items = [r for r in rows if isinstance(r, dict) and isinstance(r.get("row"), int)]
    if not items:
        return {"BT": [], "MT": []}
    items = sorted(items, key=lambda x: x.get("row", 0))
    blocks: List[List[dict]] = []
    current: List[dict] = []
    prev = None
    for item in items:
        row = item.get("row", 0)
        if prev is not None and row - prev > 1:
            blocks.append(current)
            current = []
        current.append(item)
        prev = row
    if current:
        blocks.append(current)
    if len(blocks) == 1:
        return {"BT": blocks[0], "MT": []}
    return {"BT": blocks[0], "MT": blocks[1]}


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
    seen: set[str] = set()
    if isinstance(data, list):
        for d in data:
            try:
                nombre = str(d.get("equipo", "")).strip()
                if not nombre:
                    continue
                key = nombre.upper()
                if key in seen:
                    continue
                seen.add(key)
                items.append(EquipoCatalogItem(nombre, float(d.get("btu_hr_ft", 0.0))))
            except Exception:
                continue
    info = _load_analysis_info(dir_path)
    info_items = info.get("equipos", [])
    if isinstance(info_items, list):
        for d in info_items:
            try:
                nombre = str(d.get("equipo", "")).strip()
                if not nombre:
                    continue
                key = nombre.upper()
                if key in seen:
                    continue
                seen.add(key)
                items.append(EquipoCatalogItem(nombre, float(d.get("btu_hr_ft", 0.0))))
            except Exception:
                continue
    return items, bool(items)


def load_equipos_split(dir_path: Path) -> Tuple[Dict[str, List[EquipoCatalogItem]], bool]:
    info = _load_analysis_info(dir_path)
    info_items = info.get("equipos", [])
    out: Dict[str, List[EquipoCatalogItem]] = {"BT": [], "MT": []}
    seen_bt: set[str] = set()
    seen_mt: set[str] = set()

    if isinstance(info_items, list):
        split = _split_equipos_by_row(info_items)
        for d in split.get("BT", []):
            try:
                nombre = str(d.get("equipo", "")).strip()
                if not nombre:
                    continue
                u = nombre.upper()
                if u in seen_bt:
                    continue
                seen_bt.add(u)
                out["BT"].append(EquipoCatalogItem(nombre, float(d.get("btu_hr_ft", 0.0))))
            except Exception:
                continue
        for d in split.get("MT", []):
            try:
                nombre = str(d.get("equipo", "")).strip()
                if not nombre:
                    continue
                u = nombre.upper()
                if u in seen_mt:
                    continue
                seen_mt.add(u)
                out["MT"].append(EquipoCatalogItem(nombre, float(d.get("btu_hr_ft", 0.0))))
            except Exception:
                continue

    # Mezclar equipos.json con clasificación simple por texto
    data = _safe_read_json(dir_path / "equipos.json") or []
    if isinstance(data, list):
        for d in data:
            try:
                nombre = str(d.get("equipo", "")).strip()
                if not nombre:
                    continue
                u = nombre.upper()
                bt = "CONGEL" in u
                mt = "CONSERV" in u
                if not bt and not mt:
                    bt = True
                    mt = True
                if bt and u not in seen_bt:
                    seen_bt.add(u)
                    out["BT"].append(EquipoCatalogItem(nombre, float(d.get("btu_hr_ft", 0.0))))
                if mt and u not in seen_mt:
                    seen_mt.add(u)
                    out["MT"].append(EquipoCatalogItem(nombre, float(d.get("btu_hr_ft", 0.0))))
            except Exception:
                continue

    found = bool(out.get("BT")) or bool(out.get("MT"))
    return out, found


def load_usos(dir_path: Path) -> Tuple[Dict[str, List[str]], bool]:
    data = _safe_read_json(dir_path / "usos.json") or {}
    usos = {"BT": [], "MT": []}
    if isinstance(data, dict):
        for k in ("BT", "MT"):
            val = data.get(k, [])
            if isinstance(val, list):
                usos[k] = [str(x) for x in val]
    info = _load_analysis_info(dir_path)
    info_usos = info.get("usos", {})
    if isinstance(info_usos, dict):
        for k in ("BT", "MT"):
            val = info_usos.get(k, [])
            if isinstance(val, list):
                existing = [str(x) for x in usos.get(k, [])]
                seen = {e.upper() for e in existing}
                for x in val:
                    name = str(x)
                    if name.upper() in seen:
                        continue
                    seen.add(name.upper())
                    existing.append(name)
                usos[k] = existing
    return usos, bool(usos.get("BT")) or bool(usos.get("MT"))


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
