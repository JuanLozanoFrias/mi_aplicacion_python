# logic/step4_borneras.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import pandas as pd

from .util_excel import letter_to_index as col, cell, to_int

# Interfaz pública
__all__ = ["borneras_compresores_totales"]


def _arr_code(txt: str) -> Optional[str]:
    """
    Normaliza el tipo de arranque a código: V (Variador), P (Partido), D (Directo).
    """
    u = (txt or "").strip().upper()
    if not u:
        return None
    if u[:1] in ("V", "P", "D"):
        # ya viene como V/P/D o "VFD..." etc.
        return u[:1]
    if "VARIADOR" in u or "VFD" in u or "INVERTER" in u:
        return "V"
    if "PARTIDO" in u or "ESTRELLA" in u or "TRIANGULO" in u:
        return "P"
    if "DIRECTO" in u:
        return "D"
    return None


def _count_by_type_from_tables(tables: List) -> Dict[str, int]:
    """
    Cuenta cuántos compresores hay por tipo de arranque a partir de ResumenTable.
    Espera que cada tabla tenga t.arranque (ej: 'VARIADOR' / 'DIRECTO' / 'PARTIDO').
    """
    counts = {"V": 0, "P": 0, "D": 0}
    for t in tables or []:
        code = _arr_code(getattr(t, "arranque", "") or "")
        if code in counts:
            counts[code] += 1
    return counts


def _count_by_type_from_step2(step2_state: Dict[str, Dict[str, object]]) -> Dict[str, int]:
    """
    Alternativa si no tenemos 'tables'. Lee el 'arranque' del estado del Paso 2.
    Admite valores simples o dicts con 'value'.
    """
    counts = {"V": 0, "P": 0, "D": 0}
    for _, st in (step2_state or {}).items():
        arr = ""
        for k, v in (st or {}).items():
            key = (k or "").strip().upper()
            if "ARRANQUE" in key or key in ("ARRANQUE", "TIPOARRANQUE"):
                if isinstance(v, dict) and "value" in v:
                    arr = str(v.get("value") or "")
                else:
                    arr = "" if v is None else str(v)
                break
        code = _arr_code(arr)
        if code in counts:
            counts[code] += 1
    return counts


def borneras_compresores_totales(
    basedatos_path: Path | str,
    *,
    tables: Optional[List] = None,
    step2_state: Optional[Dict[str, Dict[str, object]]] = None,
) -> Tuple[int, int, int]:
    """
    Suma AX/AY/AZ de 'OPCIONES CO2' por tipo BA (V/P/D) y los multiplica por
    la cantidad de compresores de cada tipo (derivada de 'tables' o 'step2_state').

    Retorna: (fase_total, neutro_total, tierra_total)
    """
    book = Path(basedatos_path)
    df_ops = pd.read_excel(book, sheet_name="OPCIONES CO2", header=None, dtype=str)

    ax = col("AX")
    ay = col("AY")
    az = col("AZ")
    ba = col("BA")

    # 1) Sumas base por tipo en OPCIONES CO2
    sums_by_type: Dict[str, Tuple[int, int, int]] = {"V": (0, 0, 0), "P": (0, 0, 0), "D": (0, 0, 0)}
    for i in range(1, df_ops.shape[0]):  # incluye fila 2
        tcode = (cell(df_ops, i, ba) or "").strip().upper()[:1]
        if tcode not in sums_by_type:
            continue
        f0, n0, t0 = sums_by_type[tcode]
        f = to_int(cell(df_ops, i, ax))
        n = to_int(cell(df_ops, i, ay))
        t = to_int(cell(df_ops, i, az))
        sums_by_type[tcode] = (f0 + f, n0 + n, t0 + t)

    # 2) Conteo de compresores por tipo
    if tables is not None:
        count_by_type = _count_by_type_from_tables(tables)
    elif step2_state is not None:
        count_by_type = _count_by_type_from_step2(step2_state)
    else:
        # Fallback conservador: si no sabemos cuántos compresores hay, asumimos 0
        count_by_type = {"V": 0, "P": 0, "D": 0}

    # 3) Multiplicación
    tot_f = (
        count_by_type["V"] * sums_by_type["V"][0]
        + count_by_type["P"] * sums_by_type["P"][0]
        + count_by_type["D"] * sums_by_type["D"][0]
    )
    tot_n = (
        count_by_type["V"] * sums_by_type["V"][1]
        + count_by_type["P"] * sums_by_type["P"][1]
        + count_by_type["D"] * sums_by_type["D"][1]
    )
    tot_t = (
        count_by_type["V"] * sums_by_type["V"][2]
        + count_by_type["P"] * sums_by_type["P"][2]
        + count_by_type["D"] * sums_by_type["D"][2]
    )

    return tot_f, tot_n, tot_t
