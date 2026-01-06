from __future__ import annotations

from copy import copy
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


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


def _normalize_rows(items: Iterable[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows = []
    for it in items:
        if not isinstance(it, dict):
            continue
        rows.append(it)
    return rows


def _sort_rows(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return sorted(items, key=lambda r: (_to_int(r.get("ramal", 0)), _to_int(r.get("loop", 0))))


def _make_block_rows(items: List[Dict[str, Any]], default_deshielo: str) -> Tuple[List[Dict[str, Any]], float]:
    rows = []
    total = 0.0
    for it in _sort_rows(items):
        btu_hr = _to_float(it.get("btu_hr", it.get("carga_btu_h", 0.0)))
        total += btu_hr
        rows.append(
            {
                "loop": _to_int(it.get("loop", "")),
                "ramal": _to_int(it.get("ramal", "")),
                "dim_ft": it.get("dim_ft", ""),
                "equipo": it.get("equipo", ""),
                "uso": it.get("uso", ""),
                "btu_hr": btu_hr,
                "tevap": it.get("tevap_f", ""),
                "evap_qty": _to_int(it.get("evap_qty", "")),
                "evap_modelo": it.get("evap_modelo", ""),
                "deshielo": it.get("deshielo", default_deshielo),
            }
        )
    return rows, total


def build_legend_workbook(template_path: Path, project_data: Dict[str, Any]):
    wb = load_workbook(template_path)
    ws = wb.active

    specs = project_data.get("specs", {}) if isinstance(project_data, dict) else {}
    # Encabezados (celdas según plantilla)
    ws["H3"] = specs.get("proyecto", "")
    ws["L3"] = specs.get("tipo_sistema", "")
    ws["P3"] = specs.get("voltaje_principal", "")

    ws["H4"] = specs.get("ciudad", "")
    ws["L4"] = "FT"
    ws["P4"] = specs.get("voltaje_control", "")

    tcond_f = specs.get("tcond_f", "")
    ws["H5"] = tcond_f
    ws["L5"] = specs.get("refrigerante", "")
    ws["P5"] = specs.get("controlador", "")

    default_deshielo = specs.get("deshielos", "")

    # limpiar filas debajo del encabezado
    start_row = 10
    if ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)

    data_style_row = 10
    total_style_row = 7
    max_col = 16  # A..P
    data_row_height = _row_height(ws, data_style_row)
    total_row_height = _row_height(ws, total_style_row)

    bt_items = _normalize_rows(project_data.get("bt_items", []))
    mt_items = _normalize_rows(project_data.get("mt_items", []))

    bt_rows, bt_total = _make_block_rows(bt_items, default_deshielo)
    mt_rows, mt_total = _make_block_rows(mt_items, default_deshielo)

    cur = start_row

    def _write_rows(rows: List[Dict[str, Any]]) -> None:
        nonlocal cur
        for row in rows:
            _copy_row_style(ws, data_style_row, cur, max_col)
            _set_row_height(ws, cur, data_row_height)
            # merge EQUIPO (D:E)
            ws.merge_cells(start_row=cur, start_column=4, end_row=cur, end_column=5)
            # escribir valores
            ws.cell(cur, 1).value = row.get("loop", "")
            ws.cell(cur, 2).value = row.get("ramal", "")
            ws.cell(cur, 3).value = row.get("dim_ft", "")
            ws.cell(cur, 4).value = row.get("equipo", "")
            ws.cell(cur, 6).value = row.get("uso", "")
            ws.cell(cur, 7).value = row.get("btu_hr", "")
            ws.cell(cur, 8).value = row.get("tevap", "")
            ws.cell(cur, 9).value = row.get("evap_qty", "")
            ws.cell(cur, 10).value = row.get("evap_modelo", "")
            ws.cell(cur, 15).value = row.get("deshielo", "")
            cur += 1

    def _write_total(label: str, total: float) -> None:
        nonlocal cur
        _copy_row_style(ws, total_style_row, cur, max_col)
        _set_row_height(ws, cur, total_row_height)
        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=6)
        ws.cell(cur, 1).value = label
        ws.cell(cur, 7).value = total
        cur += 2  # espacio

    _write_rows(bt_rows)
    _write_total("CARGA TOTAL BAJA", bt_total)
    _write_rows(mt_rows)
    _write_total("CARGA TOTAL MEDIA", mt_total)

    # ajustar dimensión de impresión
    ws.calculate_dimension()
    return wb

