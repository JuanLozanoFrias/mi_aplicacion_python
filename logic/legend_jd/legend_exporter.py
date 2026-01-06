from __future__ import annotations

from copy import copy
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Tuple

import json
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
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


def _add_logo(ws) -> None:
    base = Path(__file__).resolve()
    logo_path = None
    for p in base.parents:
        candidate = p / "resources" / "logo.png"
        if candidate.exists():
            logo_path = candidate
            break
    if not logo_path:
        return
    try:
        img = XLImage(str(logo_path))
        img.width = 220
        img.height = 80
        ws.add_image(img, "A1")
    except Exception:
        pass


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
            model = re.sub(r"(?i)\s*-?\s*FRONTAL\b", "", model).strip(" -")
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
    _add_logo(ws)

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
    total_ranges: List[Tuple[int, int, int, int]] = []

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

    def _write_total(label: str, total: float) -> None:
        nonlocal cur
        _copy_row_style(ws, total_style_row, cur, max_col)
        _set_row_height(ws, cur, total_row_height)
        label_col = suction_col or 1
        end_col = deshielo_span[1] if deshielo_span else (deshielo_col or max_col)
        if btu_col and btu_col > label_col + 1:
            ws.merge_cells(start_row=cur, start_column=label_col, end_row=cur, end_column=btu_col - 1)
        if btu_col and end_col and end_col >= btu_col:
            ws.merge_cells(start_row=cur, start_column=btu_col, end_row=cur, end_column=end_col)
        _safe_set(ws, cur, label_col, label)
        if btu_col:
            _safe_set(ws, cur, btu_col, int(round(total)))
        clear_fill = PatternFill()
        for c in range(1, max_col + 1):
            ws.cell(cur, c).fill = clear_fill
        if btu_col:
            total_ranges.append((cur, label_col, btu_col, end_col))
        cur += 2  # espacio

    bt_start = cur
    _write_rows(bt_rows)
    bt_end = cur - 1
    if bt_rows:
        _set_block_label("BAJA", bt_start, bt_end)
    _write_total("CARGA TOTAL BAJA", bt_total)
    mt_start = cur
    _write_rows(mt_rows)
    mt_end = cur - 1
    if mt_rows:
        _set_block_label("MEDIA", mt_start, mt_end)
    _write_total("CARGA TOTAL MEDIA", mt_total)

    last_row = cur - 1
    if header_row and last_row >= header_row:
        _apply_borders(ws, header_row, last_row, 1, max_col)
        for row_idx, label_col, btu_col, end_col in total_ranges:
            _apply_borders(ws, row_idx, row_idx, label_col, btu_col - 1)
            _apply_borders(ws, row_idx, row_idx, btu_col, end_col)
    ws.calculate_dimension()
    return wb
