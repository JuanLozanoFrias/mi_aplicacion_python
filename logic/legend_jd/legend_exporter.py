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
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.utils import get_column_letter
import unicodedata


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


def _make_block_rows(
    items: List[Dict[str, Any]],
    default_deshielo: str,
    deshielo_map: Dict[str, str],
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
        rows.append(
            {
                "loop": _to_int(it.get("loop", "")),
                "ramal": _to_int(it.get("ramal", "")),
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
    if ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)

    data_style_row = start_row
    total_style_row = header_row
    max_col = ws.max_column
    data_row_height = _row_height(ws, data_style_row)
    total_row_height = _row_height(ws, total_style_row)

    bt_items = _normalize_rows(project_data.get("bt_items", []))
    mt_items = _normalize_rows(project_data.get("mt_items", []))

    bt_rows, bt_total = _make_block_rows(bt_items, default_deshielo, deshielo_por_uso.get("BT", {}))
    mt_rows, mt_total = _make_block_rows(mt_items, default_deshielo, deshielo_por_uso.get("MT", {}))

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
    _write_total("CARGA TOTAL BAJA", bt_total, add_spacer=True)
    mt_start = cur
    _write_rows(mt_rows)
    mt_end = cur - 1
    if mt_rows:
        _set_block_label("MEDIA", mt_start, mt_end)
    _write_total("CARGA TOTAL MEDIA", mt_total, add_spacer=False)

    last_row = cur - 1
    if header_row and last_row >= header_row:
        _apply_borders(ws, header_row, last_row, 1, max_col)
        for row_idx, min_col, max_col in total_ranges:
            _apply_borders(ws, row_idx, row_idx, min_col, max_col)
        for row_idx in spacer_rows:
            _apply_borders(ws, row_idx, row_idx, 1, max_col)
    ws.calculate_dimension()
    return wb
