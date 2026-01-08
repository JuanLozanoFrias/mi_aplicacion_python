from __future__ import annotations

from pathlib import Path
import json
import re
from typing import Any, Dict, List, Tuple


def _load_json(path: Path) -> Dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


_COST_PROFILE: Dict[str, Any] | None = None
_COST_PROFILE_MTIME: float | None = None


def _load_cost_profile() -> Dict[str, Any]:
    global _COST_PROFILE, _COST_PROFILE_MTIME
    paths = [
        Path("data/LEGEND/eev_cost_profile.json"),
        Path("data/legend/eev_cost_profile.json"),
    ]
    for path in paths:
        if path.exists():
            mtime = path.stat().st_mtime
            if _COST_PROFILE is None or _COST_PROFILE_MTIME != mtime:
                _COST_PROFILE = _load_json(path)
                _COST_PROFILE_MTIME = mtime
            return _COST_PROFILE or {}
    _COST_PROFILE = {}
    _COST_PROFILE_MTIME = None
    return _COST_PROFILE


def _norm(val: Any) -> str:
    return " ".join(str(val or "").strip().upper().split())


def _to_float(val: Any, default: float = 0.0) -> float:
    try:
        if isinstance(val, str):
            val = val.replace(",", ".")
        return float(val)
    except Exception:
        return default


def _to_int(val: Any, default: int = 0) -> int:
    try:
        return int(val)
    except Exception:
        return default


def _select_family_by_load(rules: List[Dict[str, Any]], load_btu_hr: float) -> str:
    for rule in rules:
        if "lt" in rule and load_btu_hr < float(rule["lt"]):
            return str(rule.get("family", "AKVP"))
        if "gte" in rule and load_btu_hr >= float(rule["gte"]):
            return str(rule.get("family", "AKVP"))
    return "AKVP"


def _select_orifice(table_rows: List[Dict[str, Any]], load_btu_hr: float) -> int | None:
    for row in table_rows:
        cap = _to_float(row.get("cap_selection_btu_hr"))
        if cap >= load_btu_hr:
            return _to_int(row.get("orifice"), None)
    return None


def _is_room_row(row: Dict[str, Any]) -> bool:
    equipo = _norm(row.get("equipo", ""))
    if "CUARTO" in equipo:
        return True
    largo = _to_float(row.get("largo_m", row.get("dim_m", 0)))
    ancho = _to_float(row.get("ancho_m", 0))
    alto = _to_float(row.get("alto_m", 0))
    return largo > 0 and ancho > 0 and alto > 0


def _room_id(row: Dict[str, Any]) -> str:
    equipo = _norm(row.get("equipo", ""))
    if "CUARTO" in equipo and equipo:
        return f"E:{equipo}"
    uso = _norm(row.get("uso", ""))
    largo = round(_to_float(row.get("largo_m", row.get("dim_m", 0))), 2)
    ancho = round(_to_float(row.get("ancho_m", 0)), 2)
    alto = round(_to_float(row.get("alto_m", 0)), 2)
    tevap = round(_to_float(row.get("tevap_f", row.get("tevap", 0))), 2)
    return f"U:{uso}|L:{largo}|W:{ancho}|H:{alto}|T:{tevap}"


def compute_eev(
    project_data: Dict[str, Any],
    bt_items: List[Dict[str, Any]],
    mt_items: List[Dict[str, Any]],
    *,
    mt_ramal_offset: int = 0,
) -> Dict[str, Any]:
    data_dir = Path("data/LEGEND")
    tables_path = data_dir / "eev_orifice_tables.json"
    catalog_path = data_dir / "eev_danfoss_catalog.json"
    tables = _load_json(tables_path)
    catalog = _load_json(catalog_path)
    if not tables or not catalog:
        return {"detail_rows": [], "bom_rows": [], "warnings": ["FALTAN ARCHIVOS EEV"]}

    specs = project_data.get("specs", {}) if isinstance(project_data, dict) else {}
    refrigerant = specs.get("refrigerante", "")
    family_map = tables.get("refrigerant_family_map", {})
    family_map_norm = {_norm(k).replace("-", ""): v for k, v in family_map.items()}
    norm_refrig = _norm(refrigerant).replace("-", "")
    if "CO2" in norm_refrig or "R744" in norm_refrig:
        refrigerant_family = "CO2"
    else:
        refrigerant_family = family_map_norm.get(norm_refrig, family_map_norm.get("DEFAULT", "FR"))
        refrigerant_family = str(refrigerant_family or "FR")

    rules = catalog.get("rules", {})
    family_rules = rules.get("valve_family_by_load_per_valve_btu_hr", [])

    def _pick_table(tevap_f: float) -> List[Dict[str, Any]]:
        block = "CONG" if tevap_f < 0 else "LAC"
        fam_tables = tables.get("tables", {}).get(refrigerant_family, {})
        return list(fam_tables.get(block, []))

    detail_rows: List[Dict[str, Any]] = []
    warnings: List[str] = []

    def _append_rows(items: List[Dict[str, Any]], suction: str, ramal_offset: int) -> None:
        for it in items:
            load = _to_float(it.get("btu_hr", it.get("carga_btu_h", 0.0)))
            if load <= 0:
                continue
            tevap = _to_float(it.get("tevap_f", it.get("tevap", 0.0)))
            valve_family = _select_family_by_load(family_rules, load)
            orifice = None
            table_rows = _pick_table(tevap)
            if table_rows:
                orifice = _select_orifice(table_rows, load)
            if valve_family == "AKVP" and not orifice:
                warnings.append("SIN SELECCIÓN DE ORIFICIO (AKVP)")
            if valve_family == "AKVP":
                model_sel = f"AKV 10P{orifice}" if orifice else "SIN SELECCIÓN"
            else:
                model_sel = "CCM 10"
            orifice_label = orifice if orifice else ("SIN SELECCIÓN" if valve_family == "AKVP" else "")
            detail_rows.append(
                {
                    "suction": suction,
                    "loop": _to_int(it.get("loop", "")),
                    "ramal": _to_int(it.get("ramal", "")) + ramal_offset,
                    "equipo": it.get("equipo", ""),
                    "uso": it.get("uso", ""),
                    "btu_hr": load,
                    "tevap_f": tevap,
                    "familia": valve_family,
                    "orifice": orifice_label,
                    "model": model_sel,
                }
            )

    _append_rows(bt_items, "BAJA", 0)
    _append_rows(mt_items, "MEDIA", mt_ramal_offset)

    akvp_rows = [r for r in detail_rows if _norm(r.get("familia")) == "AKVP"]
    ccm_rows = [r for r in detail_rows if _norm(r.get("familia")) == "CCM"]
    akvp_valve_count = len(akvp_rows)
    ccm_valve_count = len(ccm_rows)

    room_ids = set()
    for it in list(bt_items) + list(mt_items):
        if isinstance(it, dict) and _is_room_row(it):
            room_ids.add(_room_id(it))
    n_cuartos = len(room_ids)

    orifice_counts: Dict[int, int] = {i: 0 for i in range(1, 9)}
    for r in akvp_rows:
        if r.get("orifice"):
            orifice_counts[int(r["orifice"])] += 1

    sensors_rule = rules.get("sensors_per_valve", {})
    sensors_thresh = _to_float(sensors_rule.get("if_tevap_f_gte", 20))
    sensors_gte = _to_int(sensors_rule.get("qty_if_gte", 3), 3)
    sensors_lt = _to_int(sensors_rule.get("qty_if_lt", 4), 4)
    sensors_total = 0
    for r in akvp_rows:
        if _to_float(r.get("tevap_f", 0.0)) >= sensors_thresh:
            sensors_total += sensors_gte
        else:
            sensors_total += sensors_lt

    transducer_groups = set()
    for r in akvp_rows:
        transducer_groups.add((_norm(r.get("suction")), r.get("loop"), r.get("ramal")))
    transducer_count = len(transducer_groups)

    bom_rows: List[Dict[str, Any]] = []

    def _add_bom(model: str, desc: str, qty: int) -> None:
        if qty <= 0:
            return
        bom_rows.append({"model": model, "description": desc, "qty": qty})

    families = catalog.get("families", {})
    akvp = families.get("AKVP", {})
    for part in akvp.get("orifice_parts", []):
        orif = _to_int(part.get("orifice"), 0)
        qty = orifice_counts.get(orif, 0)
        _add_bom(part.get("model", ""), part.get("description", ""), qty)

    for part in akvp.get("common_parts", []):
        qty_rule = part.get("qty", {})
        qtype = _norm(qty_rule.get("type"))
        if qtype == "PER_VALVE":
            qty = akvp_valve_count * _to_int(qty_rule.get("multiplier", 1), 1)
        elif qtype == "RULE" and _norm(qty_rule.get("rule")) == "SENSORS_PER_VALVE":
            qty = sensors_total
        elif qtype == "PER_RAMAL":
            qty = transducer_count * _to_int(qty_rule.get("multiplier", 1), 1)
        elif qtype == "PER_TRANSDUCER":
            qty = transducer_count * _to_int(qty_rule.get("multiplier", 1), 1)
        else:
            qty = 0
        _add_bom(part.get("model", ""), part.get("description", ""), int(qty))

    for part in akvp.get("system_parts", []):
        qty_rule = part.get("qty", {})
        if _norm(qty_rule.get("type")) == "FIXED" and _norm(qty_rule.get("when")) == "AT_LEAST_ONE_VALVE":
            qty = _to_int(qty_rule.get("value", 1), 1) if akvp_valve_count > 0 else 0
            _add_bom(part.get("model", ""), part.get("description", ""), qty)

    ccm = families.get("CCM", {})
    for part in ccm.get("parts_per_valve", []):
        qty_rule = part.get("qty", {})
        qty = ccm_valve_count * _to_int(qty_rule.get("multiplier", 1), 1)
        _add_bom(part.get("model", ""), part.get("description", ""), qty)
    for part in ccm.get("system_parts", []):
        qty_rule = part.get("qty", {})
        if _norm(qty_rule.get("type")) == "FIXED" and _norm(qty_rule.get("when")) == "AT_LEAST_ONE_VALVE":
            qty = _to_int(qty_rule.get("value", 1), 1) if ccm_valve_count > 0 else 0
            _add_bom(part.get("model", ""), part.get("description", ""), qty)

    total_valves = akvp_valve_count + ccm_valve_count
    total_evap = 0
    for it in list(bt_items) + list(mt_items):
        if not isinstance(it, dict):
            continue
        total_evap += _to_int(it.get("evap_qty", 0), 0)
    if total_valves > 0:
        _add_bom("CAJAS ELECTRICAS", "CAJAS ELECTRICAS", total_evap + 1)
    co2_qty = 0
    if refrigerant_family == "CO2" and total_valves > 0:
        co2_qty = 1 + n_cuartos
        _add_bom(
            "DGS-IR CO2 5m+B&L",
            "Detector basic fugas para CO2 incluido B&L",
            co2_qty,
        )

    cost_profile = _load_cost_profile()
    overrides = {}
    if isinstance(project_data, dict):
        overrides = project_data.get("eev_cost_overrides", {}) or {}
    factor_default = _to_float(cost_profile.get("factor_default", 0.0))
    factor_override = overrides.get("factor")
    factor_final = _to_float(factor_override, factor_default) if factor_override is not None else factor_default
    parts = cost_profile.get("parts", {}) if isinstance(cost_profile, dict) else {}
    part_rules = cost_profile.get("model_to_part_rules", []) if isinstance(cost_profile, dict) else []
    parts_override = overrides.get("parts_base_cost", {}) if isinstance(overrides, dict) else {}
    models_override = overrides.get("models_unit_cost", {}) if isinstance(overrides, dict) else {}
    currency = cost_profile.get("currency", "") if isinstance(cost_profile, dict) else ""

    grand_total = 0.0
    for row in bom_rows:
        model = str(row.get("model", "") or "")
        model_norm = _norm(model)
        qty = _to_int(row.get("qty", 0), 0)
        unit_cost = None
        if isinstance(models_override, dict):
            if model in models_override:
                unit_cost = _to_float(models_override.get(model))
            else:
                override_norm = {_norm(k): v for k, v in models_override.items()}
                if model_norm in override_norm:
                    unit_cost = _to_float(override_norm.get(model_norm))
        if unit_cost is None:
            part_key = None
            for rule in part_rules:
                regex = rule.get("match_regex")
                if not regex:
                    continue
                try:
                    if re.search(regex, model_norm, re.IGNORECASE):
                        part_key = rule.get("part_key")
                        break
                except Exception:
                    continue
            if not part_key and isinstance(parts, dict):
                for key, part in parts.items():
                    label = _norm(part.get("label", ""))
                    if label and label == model_norm:
                        part_key = key
                        break
            if part_key and isinstance(parts, dict) and part_key in parts:
                base_cost = _to_float(parts.get(part_key, {}).get("base_cost", 0.0))
                if isinstance(parts_override, dict) and part_key in parts_override:
                    base_cost = _to_float(parts_override.get(part_key, base_cost))
                unit_cost = base_cost * (1.0 + factor_final)
        total_cost = None
        if unit_cost is not None:
            total_cost = float(qty) * float(unit_cost)
            grand_total += total_cost
        row["unit_cost"] = unit_cost
        row["total_cost"] = total_cost
        row["currency"] = currency

    package_counts = {
        "VALVULA": total_valves,
        "CONTROL": total_valves,
        "TRANSDUCTOR": transducer_count,
        "SENSOR": sensors_total,
        "CAJAS": (total_evap + 1) if total_valves > 0 else 0,
        "SENSORES CO2": co2_qty,
    }

    return {
        "detail_rows": detail_rows,
        "bom_rows": bom_rows,
        "warnings": warnings,
        "cost_currency": currency,
        "cost_total": grand_total,
        "package_counts": package_counts,
    }
