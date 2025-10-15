# logic/breaker_vfd.py
from __future__ import annotations

from typing import Any, Dict, Iterable, List, Optional
import re
import pandas as pd


def _to_float(val) -> Optional[float]:
    """Convierte '123,4 A' -> 123.4 ; None si no numérico."""
    if pd.isna(val):
        return None
    s = str(val).strip().replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s) if s not in ("", ".", "-", None) else None
    except Exception:
        return None


def _to_int(val) -> Optional[int]:
    """Convierte a int tolerante (admite '3', '3.0', ' 3 ', NaN -> None)."""
    f = _to_float(val)
    if f is None:
        return None
    try:
        return int(round(f))
    except Exception:
        return None


def seleccionar_breaker_vfd_abb(
    ruta_excel: str,
    corriente_compresor_a: float,
    factor: float = 1.5,
) -> Dict[str, Any]:
    """
    Selecciona el BREAKER para VFD (marca ABB) con reglas:
      - Columna A: sólo familia que empiece por 'B'
      - Columna K == 3
      - Columna D >= (corriente_compresor * factor), por defecto factor=1.5

    Devuelve:
      {
        aplica: bool,
        modelo: str (columna C) cuando aplica,
        dispositivo: str (col-A, ej. B...),
        valor_col_D: float,
        k: int,
        corriente_objetivo: float,
        fila_excel: int (1-based),
        hoja: 'ABB',
        error|motivo: str (cuando no aplica)
      }
    """
    try:
        amps = float(corriente_compresor_a)
    except Exception:
        return {"aplica": False, "motivo": "Corriente inválida."}

    if amps <= 0:
        return {"aplica": False, "motivo": "Corriente inválida (<= 0)."}

    objetivo = amps * float(factor)

    # Leemos sin encabezado para indexar por letra
    df = pd.read_excel(ruta_excel, sheet_name="ABB", header=None)

    # Buscar el primer bloque cuya col-A empiece por 'B'
    colA = df[0].astype(str).str.strip().str.upper()
    idx_b = colA[colA.str.startswith("B")]
    if idx_b.empty:
        return {"aplica": False, "error": "No se encontró familia 'B' en col-A de hoja ABB."}
    start_idx = int(idx_b.index.min())

    # Subconjunto desde el primer 'B'
    df_sub = df.loc[start_idx:].copy()
    df_sub["A_is_B"] = df_sub[0].astype(str).str.strip().str.upper().str.startswith("B")
    df_sub["D_num"] = df_sub[3].map(_to_float)   # Columna D numérica
    df_sub["K_int"] = df_sub[10].map(_to_int)    # Columna K como entero seguro

    # Filtro completo: A=B, K==3, D >= objetivo
    mask = (
        df_sub["A_is_B"]
        & df_sub["D_num"].notna()
        & (df_sub["D_num"] >= objetivo)
        & (df_sub["K_int"] == 3)
    )

    cand = df_sub[mask]
    if cand.empty:
        return {
            "aplica": False,
            "error": f"Sin BREAKER (A='B*', K=3) con D ≥ {objetivo:.2f} A.",
            "corriente_objetivo": float(objetivo),
        }

    # Primer valor mayor o igual -> el de menor D que cumpla
    best = cand.sort_values("D_num", kind="mergesort").iloc[0]

    return {
        "aplica": True,
        "modelo": str(best[2]).strip(),         # Columna C
        "dispositivo": str(best[0]).strip(),    # Columna A (ej. B...)
        "valor_col_D": float(best["D_num"]),    # Corriente de D usada
        "k": int(best["K_int"]) if pd.notna(best["K_int"]) else None,
        "corriente_objetivo": float(objetivo),
        "fila_excel": int(best.name) + 1,       # 1-based como Excel
        "hoja": "ABB",
    }


# (Opcional) batch helper si lo llegas a necesitar
def seleccionar_breaker_vfd_abb_batch(
    ruta_excel: str,
    corrientes_a: Iterable[float],
    factor: float = 1.5,
) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for i, amps in enumerate(corrientes_a, start=1):
        r = seleccionar_breaker_vfd_abb(ruta_excel, float(amps), factor=factor)
        r["compresor_idx"] = i
        out.append(r)
    return out
