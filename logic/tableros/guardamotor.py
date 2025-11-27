# -*- coding: utf-8 -*-
# logic/guardamotor.py
from __future__ import annotations

from typing import Any, Dict, Iterable, List, Optional
import re
import pandas as pd


# ---------------- util ----------------
def _to_number(val) -> Optional[float]:
    """Convierte '12,3 A' -> 12.3; None si no numérico."""
    if pd.isna(val):
        return None
    s = str(val).strip().replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)  # quita unidades y texto
    try:
        return float(s) if s else None
    except ValueError:
        return None


def _normaliza_texto(x: Any) -> str:
    return str(x).strip().upper() if not pd.isna(x) else ""


# ---------------- selección clásica (se mantiene por compatibilidad) ----------------
def seleccionar_guardamotor_abb(
    ruta_excel: str,
    corriente_compresor_a: float,
    tipo_arranque: str,
) -> Dict[str, Any]:
    """
    Selecciona guardamotor ABB según la regla conocida:
    - DIRECTO: Iaj = I * 1.15        -> cantidad 1
    - PARTIDO: Iaj = I * 1.15 / 2    -> cantidad 2

    Búsqueda en hoja 'ABB' (header=None):
      A (0): código que empieza por 'G'
      C (2): MODELO
      E (4): 'CAP MAX' (primer >= Iaj)

    Devuelve dict con 'aplica', 'modelo', 'cantidad', etc.
    """
    t = _normaliza_texto(tipo_arranque)
    if t not in {"DIRECTO", "PARTIDO"}:
        return {
            "aplica": False,
            "motivo": f"Tipo de arranque '{tipo_arranque}' no requiere guardamotor.",
        }

    cantidad = 1 if t == "DIRECTO" else 2
    i_ajustada = corriente_compresor_a * 1.15
    if t == "PARTIDO":
        i_ajustada /= 2.0

    # Leemos sin encabezados para indexar por letra (A=0, C=2, E=4)
    df = pd.read_excel(ruta_excel, sheet_name="ABB", header=None, engine="openpyxl")

    # 1) Primera fila con 'G' en col-A
    colA = df[0].apply(_normaliza_texto)
    idx_g = colA[colA.str.startswith("G")]
    if idx_g.empty:
        return {"aplica": False, "error": "No se encontraron 'G' en col-A de hoja ABB."}
    start_idx = int(idx_g.index.min())

    # 2) Buscar en col-E >= Iaj desde start_idx
    df_sub = df.loc[start_idx:].copy()
    df_sub["E_num"] = df_sub[4].apply(_to_number)
    match = df_sub[df_sub["E_num"].ge(i_ajustada)].head(1)

    if match.empty:
        return {
            "aplica": False,
            "error": f"No hay valor en col-E >= {i_ajustada:.2f} A desde la primera 'G'.",
            "corriente_ajustada": i_ajustada,
        }

    row = match.iloc[0]
    fila_excel_1based = int(row.name) + 1  # número de fila como en Excel
    return {
        "aplica": True,
        "tipo_arranque": t,
        "cantidad": cantidad,
        "corriente_compresor": float(corriente_compresor_a),
        "corriente_ajustada": float(i_ajustada),
        "dispositivo": str(row[0]).strip(),    # col-A (código, debe empezar por G)
        "modelo": str(row[2]).strip(),         # col-C
        "valor_col_E": float(row["E_num"]),    # col-E numérico usado en el match
        "fila_excel": fila_excel_1based,
        "hoja": "ABB",
    }


def seleccionar_guardamotor_abb_batch(
    ruta_excel: str,
    corrientes_a: Iterable[float],
    tipos_arranque: Iterable[str],
) -> List[Dict[str, Any]]:
    """Ayuda para varios compresores (G1, G2, B1, etc.)."""
    resultados = []
    for idx, (i, t) in enumerate(zip(corrientes_a, tipos_arranque), start=1):
        r = seleccionar_guardamotor_abb(ruta_excel, float(i), t)
        r["compresor_idx"] = idx
        resultados.append(r)
    return resultados


# ---------------- NUEVO: lista de candidatos con Q/R (PUENTES) ----------------
def listar_guardamotores_abb(
    ruta_excel: str,
    corriente_compresor_a: float,
    tipo_arranque: str,
) -> List[Dict[str, Any]]:
    """
    Devuelve una lista ordenada de candidatos de guardamotor ABB que cumplen:
    col-E >= Iaj (directo) o >= Iaj/2 (partido), incluyendo Q (PUENTES) y R (CÓDIGO PUENTES).

    Hoja 'ABB' con header=None:
      A=0 (código inicia con 'G'), C=2 (MODELO), E=4 (CAP MAX),
      Q=16 (PUENTES), R=17 (CÓDIGO PUENTES).
    """
    t = _normaliza_texto(tipo_arranque)
    if t not in {"DIRECTO", "PARTIDO"}:
        return []

    i_ajustada = corriente_compresor_a * 1.15
    if t == "PARTIDO":
        i_ajustada /= 2.0

    df = pd.read_excel(ruta_excel, sheet_name="ABB", header=None, engine="openpyxl")

    colA = df[0].apply(_normaliza_texto)
    mask_g = colA.str.startswith("G")

    df_g = df[mask_g].copy()
    df_g["E_num"] = df_g[4].apply(_to_number)
    cand = df_g[df_g["E_num"].ge(i_ajustada)].sort_values("E_num")

    out: List[Dict[str, Any]] = []
    for idx, row in cand.iterrows():
        q = "" if pd.isna(df.iat[idx, 16]) else str(df.iat[idx, 16]).strip()
        r = "" if pd.isna(df.iat[idx, 17]) else str(df.iat[idx, 17]).strip()
        out.append({
            "modelo": str(row[2]).strip(),      # col-C
            "valor_col_E": float(row["E_num"]) if row["E_num"] is not None else None,
            "fila_excel": int(idx) + 1,
            "puente_modelo": q,                  # Q
            "puente_codigo": r,                  # R (código interno)
            "cantidad": 1 if t == "DIRECTO" else 2,
        })
    return out

