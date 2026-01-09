from __future__ import annotations

from copy import copy
from pathlib import Path
import re
import os
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, Iterable, List, Tuple

import json
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import unicodedata

try:
    from logic.legend.eev_calc import compute_eev
except Exception:
    compute_eev = None  # type: ignore

def _to_float(val: Any, default: float = 0.0) -> float:
    try:
        if isinstance(val, str):
            val = val.replace(",", ".").strip()
        return float(val)
    except Exception:
        return default


def _to_int(val: Any, default: int = 0) -> int:
    try:
        if isinstance(val, str):
            val = val.strip()
        return int(float(val))
    except Exception:
        return default


def _norm_label(text: Any) -> str:
    s = str(text or "").strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace(":", "").replace("°", "").replace("º", "")
    s = " ".join(s.split())
    return s


def _find_label_cell(ws, label: str, max_row: int = 20) -> Tuple[int, int] | None:
    target = _norm_label(label)
    for r in range(1, max_row + 1):
        for c in range(1, ws.max_column + 1):
            if _norm_label(ws.cell(r, c).value) == target:
                return r, c
    return None


def _col_after_label(ws, row: int, col: int) -> Tuple[int, int]:
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return row, rng.max_col + 1
    return row, col + 1


def _find_header_row(ws, max_row: int = 30) -> int | None:
    for r in range(1, max_row + 1):
        row_vals = [_norm_label(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "RAMAL" in row_vals and "EQUIPO" in row_vals and "USO" in row_vals:
            return r
    return None


def _find_comp_section_row(ws, max_row: int = 200) -> int | None:
    for r in range(1, min(max_row, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            val = _norm_label(ws.cell(r, c).value)
            if not val:
                continue
            if "COMPRESORES" in val:
                return r
    return None


def _map_header_cols(ws, header_row: int) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        val = _norm_label(ws.cell(header_row, c).value)
        if not val:
            continue
        mapping.setdefault(val, c)
    return mapping


def _find_subheader_cols(ws, header_row: int, look_ahead: int = 3) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for r in range(header_row + 1, header_row + 1 + look_ahead):
        for c in range(1, ws.max_column + 1):
            val = _norm_label(ws.cell(r, c).value)
            if val in ("CANTIDAD", "MODELO"):
                mapping[val] = c
    return mapping


def _col_for(mapping: Dict[str, int], labels: Iterable[str]) -> int | None:
    for lab in labels:
        key = _norm_label(lab)
        if key in mapping:
            return mapping[key]
    return None


def _copy_row_style(ws, src_row: int, dst_row: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)


def _row_height(ws, row: int) -> float | None:
    return ws.row_dimensions[row].height


def _set_row_height(ws, row: int, height: float | None) -> None:
    if height:
        ws.row_dimensions[row].height = height


def _safe_set(ws, row: int, col: int, value: Any) -> None:
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
    cell.value = value


def _merge_span(ws, row: int, col: int) -> Tuple[int, int] | None:
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_col, rng.max_col
    return None


def _apply_borders(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    thin = Side(border_style="thin", color="000000")
    thick = Side(border_style="medium", color="000000")
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            left = thick if c == min_col else thin
            right = thick if c == max_col else thin
            top = thick if r == min_row else thin
            bottom = thick if r == max_row else thin
            ws.cell(r, c).border = Border(left=left, right=right, top=top, bottom=bottom)


def _ensure_sheet_drawing(sheet_xml: bytes) -> bytes:
    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ET.register_namespace("", ns_main)
    ET.register_namespace("r", ns_rel)
    root = ET.fromstring(sheet_xml)
    tag = f"{{{ns_main}}}drawing"
    has_drawing = root.find(tag) is not None
    if not has_drawing:
        drawing = ET.SubElement(root, tag)
        drawing.set(f"{{{ns_rel}}}id", "rId1")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_sheet_rels(rels_xml: bytes | None) -> bytes:
    ns_pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", ns_pkg)
    if rels_xml:
        root = ET.fromstring(rels_xml)
    else:
        root = ET.Element(f"{{{ns_pkg}}}Relationships")
    has_rel = False
    for rel in root.findall(f"{{{ns_pkg}}}Relationship"):
        if rel.get("Type", "").endswith("/drawing"):
            has_rel = True
            break
    if not has_rel:
        rel = ET.SubElement(root, f"{{{ns_pkg}}}Relationship")
        rel.set("Id", "rId1")
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
        rel.set("Target", "../drawings/drawing1.xml")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _merge_content_types(out_xml: bytes, tpl_xml: bytes) -> bytes:
    ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", ns)
    out_root = ET.fromstring(out_xml)
    tpl_root = ET.fromstring(tpl_xml)
    out_defaults = {(el.get("Extension"), el.get("ContentType")) for el in out_root.findall(f"{{{ns}}}Default")}
    out_overrides = {(el.get("PartName"), el.get("ContentType")) for el in out_root.findall(f"{{{ns}}}Override")}
    for el in tpl_root.findall(f"{{{ns}}}Default"):
        key = (el.get("Extension"), el.get("ContentType"))
        if key not in out_defaults:
            out_root.append(el)
            out_defaults.add(key)
    for el in tpl_root.findall(f"{{{ns}}}Override"):
        key = (el.get("PartName"), el.get("ContentType"))
        if key not in out_overrides:
            out_root.append(el)
            out_overrides.add(key)
    return ET.tostring(out_root, encoding="utf-8", xml_declaration=True)


def restore_template_assets(template_path: Path, output_path: Path, logo_override: Path | None = None) -> bool:
    try:
        with zipfile.ZipFile(template_path) as ztpl, zipfile.ZipFile(output_path) as zout:
            tpl_names = ztpl.namelist()
            media_names = [n for n in tpl_names if n.startswith("xl/media/")]
            drawing_names = [n for n in tpl_names if n.startswith("xl/drawings/")]
            sheet_xml = zout.read("xl/worksheets/sheet1.xml")
            new_sheet_xml = _ensure_sheet_drawing(sheet_xml)
            rels_path = "xl/worksheets/_rels/sheet1.xml.rels"
            rels_xml = zout.read(rels_path) if rels_path in zout.namelist() else None
            new_rels_xml = _ensure_sheet_rels(rels_xml)
            ct_xml = zout.read("[Content_Types].xml")
            tpl_ct_xml = ztpl.read("[Content_Types].xml")
            new_ct_xml = _merge_content_types(ct_xml, tpl_ct_xml)

            tmp_fd, tmp_name = tempfile.mkstemp(suffix=".xlsx")
            os.close(tmp_fd)
            Path(tmp_name).unlink(missing_ok=True)
            tmp_path = Path(tmp_name)
            with zipfile.ZipFile(tmp_path, "w") as znew:
                for item in zout.infolist():
                    name = item.filename
                    if name in ("xl/worksheets/sheet1.xml", rels_path, "[Content_Types].xml"):
                        continue
                    if name.startswith("xl/media/") or name.startswith("xl/drawings/"):
                        continue
                    znew.writestr(item, zout.read(name))
                znew.writestr("xl/worksheets/sheet1.xml", new_sheet_xml)
                znew.writestr(rels_path, new_rels_xml)
                znew.writestr("[Content_Types].xml", new_ct_xml)
                for name in media_names:
                    if logo_override and media_names and name == media_names[0]:
                        znew.writestr(name, logo_override.read_bytes())
                    else:
                        znew.writestr(name, ztpl.read(name))
                for name in drawing_names:
                    znew.writestr(name, ztpl.read(name))
        output_path.write_bytes(tmp_path.read_bytes())
        return True
    except Exception:
        return False


def _normalize_rows(items: Iterable[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows = []
    for it in items:
        if not isinstance(it, dict):
            continue
        rows.append(it)
    return rows


def _sort_rows(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return sorted(items, key=lambda r: (_to_int(r.get("ramal", 0)), _to_int(r.get("loop", 0))))


def _normalize_deshielo_type(val: Any) -> str:
    s = _norm_label(val)
    if not s:
        return ""
    if "GAS CALIENTE" in s or "GASCALIENTE" in s:
        return "GAS CALIENTE"
    if "GAS TIBIO" in s or "GASTIBIO" in s:
        return "GAS TIBIO"
    if "ELECTR" in s:
        return "ELECTRICO"
    return s


def _format_deshielo(default_type: Any) -> str:
    dtype = _normalize_deshielo_type(default_type)
    if not dtype:
        return "DESHIELO"
    return f"DESHIELO {dtype}"


def _load_deshielo_por_uso() -> Dict[str, Dict[str, str]]:
    candidates = [
        Path("data/legend/deshielo_por_uso.json"),
        Path("data/LEGEND/deshielo_por_uso.json"),
    ]
    for path in candidates:
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return data
            except Exception:
                return {}
    return {}


def _load_eev_cost_profile() -> Dict[str, Any]:
    for path in (
        Path("data/LEGEND/eev_cost_profile.json"),
        Path("data/legend/eev_cost_profile.json"),
    ):
        if path.exists():
            try:
                return json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                return {}
    return {}


def _compute_eev_sets_for_export(project_data: Dict[str, Any], result: Dict[str, Any]) -> List[Dict[str, Any]]:
    profile = _load_eev_cost_profile()
    if not profile:
        return []
    overrides = project_data.get("eev_cost_overrides", {}) if isinstance(project_data, dict) else {}
    factor_default = _to_float(profile.get("factor_default", 0.0))
    factor_override = overrides.get("factor") if isinstance(overrides, dict) else None
    factor_val = _to_float(factor_override, factor_default) if factor_override is not None else factor_default
    parts = profile.get("parts", {}) if isinstance(profile, dict) else {}
    parts_override = overrides.get("parts_base_cost", {}) if isinstance(overrides, dict) else {}
    sets_cfg = profile.get("sets", {}) if isinstance(profile, dict) else {}
    currency = profile.get("currency", "") if isinstance(profile, dict) else ""
    package_counts = result.get("package_counts", {}) if isinstance(result, dict) else {}

    def _unit_cost_for_part(part_key: str) -> float:
        base_cost = _to_float(parts.get(part_key, {}).get("base_cost", 0.0))
        if isinstance(parts_override, dict) and part_key in parts_override:
            base_cost = _to_float(parts_override.get(part_key, base_cost))
        return base_cost * (1.0 + factor_val)

    def _set_unit_cost(set_key: str) -> float | None:
        cfg = sets_cfg.get(set_key, {}) if isinstance(sets_cfg, dict) else {}
        parts_list = cfg.get("parts", []) if isinstance(cfg, dict) else []
        if not parts_list:
            return None
        total = 0.0
        for part_key in parts_list:
            total += _unit_cost_for_part(str(part_key))
        return total

    order = ["VALVULA", "CONTROL", "SENSOR", "TRANSDUCTOR", "CAJAS", "SENSORES CO2"]
    rows: List[Dict[str, Any]] = []
    for key in order:
        qty = _to_int(package_counts.get(key, 0), 0)
        if key in ("CAJAS", "SENSORES CO2"):
            part_key = "CAJAS_ELECTRICAS" if key == "CAJAS" else "DGS_IR_CO2"
            unit_cost = _unit_cost_for_part(part_key)
        else:
            unit_cost = _set_unit_cost(key)
        total_cost = None if unit_cost is None else qty * unit_cost
        rows.append(
            {
                "category": key,
                "qty": qty,
                "unit_cost": unit_cost,
                "total_cost": total_cost,
                "currency": currency,
            }
        )
    return rows


def _write_eev_sheet(wb, project_data: Dict[str, Any], mt_ramal_offset: int) -> None:
    ws = wb["EEV"] if "EEV" in wb.sheetnames else wb.create_sheet("EEV")
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    if ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)
    if ws.max_column > 0:
        ws.delete_cols(1, ws.max_column)

    specs = project_data.get("specs", {}) if isinstance(project_data, dict) else {}
    if _norm_label(specs.get("expansion", "")) != "ELECTRONICA":
        return
    if not compute_eev:
        return

    bt_items = _normalize_rows(project_data.get("bt_items", []))
    mt_items = _normalize_rows(project_data.get("mt_items", []))
    result = compute_eev(project_data, bt_items, mt_items, mt_ramal_offset=mt_ramal_offset)
    detail_rows = result.get("detail_rows", []) if isinstance(result, dict) else []
    bom_rows = result.get("bom_rows", []) if isinstance(result, dict) else []
    set_rows = _compute_eev_sets_for_export(project_data, result if isinstance(result, dict) else {})
    if not detail_rows and not bom_rows and not set_rows:
        return

    header_fill = PatternFill("solid", fgColor="C6EFCE")
    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    def _upper(val: Any) -> str:
        return str(val or "").upper()

    def _write_headers(row: int, headers: List[str], start_col: int = 1) -> None:
        for idx, h in enumerate(headers):
            cell = ws.cell(row, start_col + idx, h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center

    def _set_widths(widths: Dict[int, float]) -> None:
        for col_idx, w in widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = w

    row = 1
    detail_headers = [
        "SUCCION",
        "LOOP",
        "RAMAL",
        "EQUIPO",
        "USO",
        "CARGA (BTU/HR)",
        "TEVAP (F)",
        "FAMILIA (EEV)",
        "ORIFICIO",
        "MODELO",
    ]
    max_detail_col = len(detail_headers)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max_detail_col)
    title = ws.cell(row, 1, "EEV - EXPANSION ELECTRONICA")
    title.font = title_font
    title.alignment = left
    row += 2
    _write_headers(row, detail_headers)
    row += 1
    start_detail = row - 1
    for drow in detail_rows:
        ws.cell(row, 1, _upper(drow.get("suction", ""))).alignment = center
        ws.cell(row, 2, drow.get("loop", "")).alignment = center
        ws.cell(row, 3, drow.get("ramal", "")).alignment = center
        ws.cell(row, 4, _upper(drow.get("equipo", ""))).alignment = left
        ws.cell(row, 5, _upper(drow.get("uso", ""))).alignment = left
        cell_btu = ws.cell(row, 6, int(round(_to_float(drow.get("btu_hr", 0)))))
        cell_btu.number_format = "#,##0"
        cell_btu.alignment = right
        cell_t = ws.cell(row, 7, _to_float(drow.get("tevap_f", 0.0)))
        cell_t.number_format = "0.0"
        cell_t.alignment = right
        ws.cell(row, 8, _upper(drow.get("familia", ""))).alignment = center
        ws.cell(row, 9, _upper(drow.get("orifice", ""))).alignment = center
        ws.cell(row, 10, _upper(drow.get("model", ""))).alignment = left
        row += 1
    end_detail = row - 1
    _apply_borders(ws, start_detail, end_detail, 1, max_detail_col)

    row += 1
    bom_headers = ["MODELO", "DESCRIPCION", "CANTIDAD", "COSTO UNITARIO", "COSTO TOTAL", "MONEDA"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(bom_headers))
    bom_title = ws.cell(row, 1, "RESUMEN EEV")
    bom_title.font = title_font
    bom_title.alignment = left
    row += 1
    _write_headers(row, bom_headers)
    row += 1
    start_bom = row - 1
    currency = result.get("cost_currency", "") if isinstance(result, dict) else ""
    for brow in bom_rows:
        ws.cell(row, 1, _upper(brow.get("model", ""))).alignment = left
        ws.cell(row, 2, _upper(brow.get("description", ""))).alignment = left
        cell_qty = ws.cell(row, 3, _to_int(brow.get("qty", 0)))
        cell_qty.alignment = right
        unit_cost = brow.get("unit_cost")
        total_cost = brow.get("total_cost")
        if unit_cost is not None:
            c_unit = ws.cell(row, 4, float(unit_cost))
            c_unit.number_format = "#,##0.00"
            c_unit.alignment = right
        if total_cost is not None:
            c_tot = ws.cell(row, 5, float(total_cost))
            c_tot.number_format = "#,##0.00"
            c_tot.alignment = right
        ws.cell(row, 6, _upper(brow.get("currency", currency))).alignment = center
        row += 1
    end_bom = row - 1
    _apply_borders(ws, start_bom, end_bom, 1, len(bom_headers))

    if set_rows:
        row += 1
        set_headers = ["CATEGORIA", "CANTIDAD", "COSTO UNITARIO", "COSTO TOTAL", "MONEDA"]
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(set_headers))
        set_title = ws.cell(row, 1, "PAQUETES")
        set_title.font = title_font
        set_title.alignment = left
        row += 1
        _write_headers(row, set_headers)
        row += 1
        start_set = row - 1
        total_set_cost = 0.0
        for srow in set_rows:
            ws.cell(row, 1, _upper(srow.get("category", ""))).alignment = left
            cell_qty = ws.cell(row, 2, _to_int(srow.get("qty", 0)))
            cell_qty.alignment = right
            unit_cost = srow.get("unit_cost")
            total_cost = srow.get("total_cost")
            if unit_cost is not None:
                c_unit = ws.cell(row, 3, float(unit_cost))
                c_unit.number_format = "#,##0.00"
                c_unit.alignment = right
            if total_cost is not None:
                c_tot = ws.cell(row, 4, float(total_cost))
                c_tot.number_format = "#,##0.00"
                c_tot.alignment = right
                total_set_cost += float(total_cost)
            ws.cell(row, 5, _upper(srow.get("currency", currency))).alignment = center
            row += 1
        ws.cell(row, 1, "TOTAL").font = header_font
        tot_cell = ws.cell(row, 4, total_set_cost)
        tot_cell.number_format = "#,##0.00"
        tot_cell.alignment = right
        ws.cell(row, 5, _upper(currency)).alignment = center
        end_set = row
        _apply_borders(ws, start_set, end_set, 1, len(set_headers))
        row += 1

    def _autofit_columns(min_col: int, max_col: int) -> None:
        for col_idx in range(min_col, max_col + 1):
            max_len = 0
            for r in range(1, ws.max_row + 1):
                val = ws.cell(r, col_idx).value
                if val is None:
                    continue
                max_len = max(max_len, len(str(val)))
            if max_len <= 0:
                continue
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    _autofit_columns(1, max_detail_col)


def _make_block_rows(
    items: List[Dict[str, Any]],
    default_deshielo: str,
    deshielo_map: Dict[str, str],
    ramal_offset: int = 0,
) -> Tuple[List[Dict[str, Any]], float]:
    rows = []
    total = 0.0
    for it in _sort_rows(items):
        btu_hr = int(round(_to_float(it.get("btu_hr", it.get("carga_btu_h", 0.0)))))
        total += btu_hr
        model = str(it.get("evap_modelo", "") or "")
        if model:
            model = re.sub(r"(?i)\s*-?\s*(DUAL|FRONTAL)\b", "", model)
            model = re.sub(r"\s{2,}", " ", model)
            model = model.strip(" -")
        uso = str(it.get("uso", "") or "")
        uso_key = _norm_label(uso)
        flag = deshielo_map.get(uso_key)
        if flag is not None:
            if _norm_label(flag) == "NO":
                deshielo_val = "DESHIELO POR TIEMPO"
            else:
                deshielo_val = _format_deshielo(default_deshielo)
        else:
            deshielo_val = str(it.get("deshielo", "") or "")
            if not deshielo_val:
                deshielo_val = "DESHIELO POR TIEMPO"
        ramal_val = _to_int(it.get("ramal", ""))
        if ramal_offset:
            ramal_val += ramal_offset
        rows.append(
            {
                "loop": _to_int(it.get("loop", "")),
                "ramal": ramal_val,
                "dim_ft": it.get("dim_ft", ""),
                "equipo": it.get("equipo", ""),
                "uso": uso,
                "btu_hr": btu_hr,
                "tevap": it.get("tevap_f", ""),
                "evap_qty": _to_int(it.get("evap_qty", "")),
                "evap_modelo": model,
                "deshielo": deshielo_val,
            }
        )
    return rows, total


def build_legend_workbook(template_path: Path, project_data: Dict[str, Any]):
    wb = load_workbook(template_path)
    ws = wb.active

    specs = project_data.get("specs", {}) if isinstance(project_data, dict) else {}

    def _write_spec(label: str, value: Any) -> None:
        pos = _find_label_cell(ws, label, max_row=6)
        if not pos:
            return
        row, col = pos
        row, col = _col_after_label(ws, row, col)
        _safe_set(ws, row, col, value)

    _write_spec("PROYECTO", specs.get("proyecto", ""))
    _write_spec("CIUDAD", specs.get("ciudad", ""))
    _write_spec("TIPO SISTEMA", specs.get("tipo_sistema", ""))
    _write_spec("VOLTAJE PRINCIPAL", specs.get("voltaje_principal", ""))
    _write_spec("VOLTAJE CONTROL", specs.get("voltaje_control", ""))
    _write_spec("T. CONDENSACION", specs.get("tcond_f", ""))
    _write_spec("REFRIGERANTE", specs.get("refrigerante", ""))
    _write_spec("CONTROLADOR", specs.get("controlador", ""))
    _write_spec("MEDIDAS", "FT")

    default_deshielo = specs.get("deshielos", "")
    deshielo_por_uso = _load_deshielo_por_uso()

    header_row = _find_header_row(ws) or 7
    header_map = _map_header_cols(ws, header_row)
    sub_map = _find_subheader_cols(ws, header_row)

    suction_col = _col_for(header_map, ["SUCCION"])
    loop_col = _col_for(header_map, ["LOOP"])
    ramal_col = _col_for(header_map, ["RAMAL"])
    dim_col = _col_for(header_map, ["DIM (FT)", "DIM"])
    equipo_col = _col_for(header_map, ["EQUIPO"])
    uso_col = _col_for(header_map, ["USO"])
    btu_col = _col_for(header_map, ["CARGA (BTU/HR)", "CARGA (BTU/H)", "CARGA (BTU/HR)"])
    tevap_col = _col_for(header_map, ["TEVAP (°F)", "TEVAP (F)", "TEVAP"])
    evap_hdr_col = _col_for(header_map, ["EVAPORADORES"])
    qty_col = sub_map.get("CANTIDAD") or evap_hdr_col
    model_col = sub_map.get("MODELO") or (evap_hdr_col + 1 if evap_hdr_col else None)
    deshielo_col = _col_for(header_map, ["DESHIELO"])
    equipo_span = _merge_span(ws, header_row, equipo_col) if equipo_col else None
    deshielo_span = _merge_span(ws, header_row, deshielo_col) if deshielo_col else None

    start_row = header_row + 2

    bt_items = _normalize_rows(project_data.get("bt_items", []))
    mt_items = _normalize_rows(project_data.get("mt_items", []))

    bt_count = _to_int(project_data.get("bt_ramales", 0))
    if bt_count <= 0:
        bt_count = max((_to_int(r.get("ramal", 0)) for r in bt_items), default=0)
    bt_rows, bt_total = _make_block_rows(
        bt_items, default_deshielo, deshielo_por_uso.get("BT", {}), ramal_offset=0
    )
    mt_rows, mt_total = _make_block_rows(
        mt_items, default_deshielo, deshielo_por_uso.get("MT", {}), ramal_offset=bt_count
    )
    needed_rows = len(bt_rows) + len(mt_rows) + 3
    if needed_rows < 1:
        needed_rows = 1

    comp_row = _find_comp_section_row(ws)
    reserva_pos = _find_label_cell(ws, "RESERVA %", max_row=200) or _find_label_cell(
        ws, "RESERVA%", max_row=200
    )
    summary_rows: List[int] = []
    for label in ("CAP", "CARGA", "RESERVA %", "RESERVA%"):
        pos = _find_label_cell(ws, label, max_row=200)
        if pos:
            summary_rows.append(pos[0])
    block_start = None
    if summary_rows:
        block_start = min(summary_rows) - 1
    if block_start is not None:
        block_start = max(start_row, block_start)
    anchor_row = block_start or comp_row

    if anchor_row and anchor_row > start_row:
        available_rows = anchor_row - start_row
        bt_spacer = True
        mt_spacer = True
        required_rows = (
            len(bt_rows)
            + 1
            + (1 if bt_spacer else 0)
            + len(mt_rows)
            + 1
            + (1 if mt_spacer else 0)
        )
        if available_rows < required_rows:
            ws.insert_rows(anchor_row, required_rows - available_rows)
        elif available_rows > required_rows:
            ws.delete_rows(start_row + required_rows, available_rows - required_rows)
    else:
        if ws.max_row >= start_row:
            ws.delete_rows(start_row, ws.max_row - start_row + 1)
        rows_to_insert = max(needed_rows, 0)
        if rows_to_insert:
            ws.insert_rows(start_row, rows_to_insert)

    data_style_row = start_row
    total_style_row = header_row
    max_col = ws.max_column
    data_row_height = _row_height(ws, data_style_row)
    total_row_height = _row_height(ws, total_style_row)

    cur = start_row
    total_ranges: List[Tuple[int, int, int]] = []
    spacer_rows: List[int] = []

    def _is_deshielo_por_tiempo(val: Any) -> bool:
        return _norm_label(val) in ("DESHIELO POR TIEMPO", "POR TIEMPO", "NO")

    def _set_block_label(label: str, start: int, end: int) -> None:
        if not suction_col or start > end:
            return
        ws.merge_cells(start_row=start, start_column=suction_col, end_row=end, end_column=suction_col)
        cell = ws.cell(start, suction_col)
        cell.value = label
        cell.alignment = cell.alignment.copy(horizontal="center", vertical="center")

    def _merge_loop_block_empty(start: int, end: int) -> None:
        if not loop_col or start > end:
            return
        for r in range(start, end + 1):
            ws.cell(r, loop_col).value = ""
        if end > start:
            ws.merge_cells(start_row=start, start_column=loop_col, end_row=end, end_column=loop_col)
        cell = ws.cell(start, loop_col)
        cell.alignment = cell.alignment.copy(horizontal="center", vertical="center")

    def _merge_loop_ranges(start: int, end: int) -> None:
        if not loop_col or start > end:
            return
        row = start
        while row <= end:
            val = ws.cell(row, loop_col).value
            if val is None or str(val).strip() == "":
                row += 1
                continue
            next_row = row + 1
            while next_row <= end:
                next_val = ws.cell(next_row, loop_col).value
                if next_val == val:
                    next_row += 1
                    continue
                break
            if next_row - row > 1:
                ws.merge_cells(start_row=row, start_column=loop_col, end_row=next_row - 1, end_column=loop_col)
                cell = ws.cell(row, loop_col)
                cell.alignment = cell.alignment.copy(horizontal="center", vertical="center")
            row = next_row

    def _write_rows(rows: List[Dict[str, Any]]) -> None:
        nonlocal cur
        for row in rows:
            _copy_row_style(ws, data_style_row, cur, max_col)
            _set_row_height(ws, cur, data_row_height)
            if equipo_span and equipo_span[1] > equipo_span[0]:
                ws.merge_cells(start_row=cur, start_column=equipo_span[0], end_row=cur, end_column=equipo_span[1])
            if deshielo_span and deshielo_span[1] > deshielo_span[0]:
                ws.merge_cells(start_row=cur, start_column=deshielo_span[0], end_row=cur, end_column=deshielo_span[1])
            if loop_col:
                _safe_set(ws, cur, loop_col, row.get("loop", ""))
                lcell = ws.cell(cur, loop_col)
                lcell.alignment = lcell.alignment.copy(horizontal="center", vertical="center")
            if ramal_col:
                _safe_set(ws, cur, ramal_col, row.get("ramal", ""))
            if dim_col:
                _safe_set(ws, cur, dim_col, row.get("dim_ft", ""))
            if equipo_col:
                _safe_set(ws, cur, equipo_col, row.get("equipo", ""))
            if uso_col:
                _safe_set(ws, cur, uso_col, row.get("uso", ""))
            if btu_col:
                _safe_set(ws, cur, btu_col, int(round(_to_float(row.get("btu_hr", 0)))))
            if tevap_col:
                _safe_set(ws, cur, tevap_col, row.get("tevap", ""))
            if qty_col:
                _safe_set(ws, cur, qty_col, row.get("evap_qty", ""))
            if model_col:
                _safe_set(ws, cur, model_col, row.get("evap_modelo", ""))
            if deshielo_col:
                _safe_set(ws, cur, deshielo_col, row.get("deshielo", ""))
            is_pt = _is_deshielo_por_tiempo(row.get("deshielo", default_deshielo))
            if tevap_col:
                cell = ws.cell(cur, tevap_col)
                cell.font = cell.font.copy(color="000000" if is_pt else "FF0000")
            if deshielo_col:
                dcell = ws.cell(cur, deshielo_col)
                dcell.font = dcell.font.copy(color="000000" if is_pt else "FF0000")
            cur += 1

    def _write_total(label: str, total: float, *, add_spacer: bool) -> None:
        nonlocal cur
        _copy_row_style(ws, total_style_row, cur, max_col)
        _set_row_height(ws, cur, total_row_height)
        label_col = suction_col or 1
        end_col = deshielo_span[1] if deshielo_span else (deshielo_col or max_col)
        if btu_col and btu_col > label_col + 1:
            ws.merge_cells(start_row=cur, start_column=label_col, end_row=cur, end_column=btu_col - 1)
        if tevap_col and end_col and end_col >= tevap_col:
            ws.merge_cells(start_row=cur, start_column=tevap_col, end_row=cur, end_column=end_col)
        _safe_set(ws, cur, label_col, label)
        if btu_col:
            _safe_set(ws, cur, btu_col, int(round(total)))
        clear_fill = PatternFill()
        for c in range(1, max_col + 1):
            ws.cell(cur, c).fill = clear_fill
        if btu_col:
            total_ranges.append((cur, label_col, btu_col - 1))
            total_ranges.append((cur, btu_col, btu_col))
        if tevap_col and end_col and end_col >= tevap_col:
            total_ranges.append((cur, tevap_col, end_col))
        if add_spacer:
            spacer_row = cur + 1
            ws.merge_cells(start_row=spacer_row, start_column=1, end_row=spacer_row, end_column=max_col)
            spacer_rows.append(spacer_row)
            cur += 2
        else:
            cur += 1

    bt_start = cur
    _write_rows(bt_rows)
    bt_end = cur - 1
    if bt_rows:
        _set_block_label("BAJA", bt_start, bt_end)
        if _norm_label(specs.get("distribucion_tuberia", "")) == "AMERICANA":
            _merge_loop_block_empty(bt_start, bt_end)
        else:
            _merge_loop_ranges(bt_start, bt_end)
    _write_total("CARGA TOTAL BAJA", bt_total, add_spacer=True)
    mt_start = cur
    _write_rows(mt_rows)
    mt_end = cur - 1
    if mt_rows:
        _set_block_label("MEDIA", mt_start, mt_end)
        if _norm_label(specs.get("distribucion_tuberia", "")) == "AMERICANA":
            _merge_loop_block_empty(mt_start, mt_end)
        else:
            _merge_loop_ranges(mt_start, mt_end)
    _write_total("CARGA TOTAL MEDIA", mt_total, add_spacer=True if anchor_row else False)

    last_row = cur - 1
    if header_row and last_row >= header_row:
        _apply_borders(ws, header_row, last_row, 1, max_col)
        for row_idx, min_col, max_col in total_ranges:
            _apply_borders(ws, row_idx, row_idx, min_col, max_col)
        if spacer_rows:
            empty_border = Border()
            for row_idx in spacer_rows:
                for c in range(1, max_col + 1):
                    ws.cell(row_idx, c).border = empty_border
    ws.calculate_dimension()
    try:
        _write_eev_sheet(wb, project_data, bt_count)
    except Exception:
        pass
    return wb
