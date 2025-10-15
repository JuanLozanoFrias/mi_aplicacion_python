# logic/contactor.py
from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional

import pandas as pd
import re


def _num(x) -> Optional[float]:
    """Extrae el primer número de un string (acepta , o .)."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x)
    m = re.search(r"[-+]?\d+(?:[.,]\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", "."))
    except Exception:
        return None


def seleccionar_contactor_abb(book_path: str, corriente_compresor: float, tipo_arranque: str) -> Dict:
    """
    Contactor:
      - Sólo DIRECTO o PARTIDO.
      - Buscar filas con 'C...' en col A.
      - En col D: primer valor >= I*1.15 (directo) o >= I*1.15/2 (partido).
      - Devolver col C (modelo). También reporta el valor de col D y cantidad sugerida.
    """
    arr = (tipo_arranque or "").strip().upper()
    if arr not in ("DIRECTO", "PARTIDO"):
        return {"aplica": False, "motivo": "NO APLICA"}

    try:
        book = Path(book_path)

        # Hoja ABB o fallback a la primera hoja que tenga códigos 'C' en la columna A
        try:
            df = pd.read_excel(book, sheet_name="ABB", header=None, dtype=object)
        except Exception:
            xls = pd.ExcelFile(book)
            chosen = None
            for s in xls.sheet_names:
                tmp = pd.read_excel(xls, sheet_name=s, header=None, dtype=object)
                colA = tmp.iloc[:, 0].astype(str).str.upper().str.strip()
                if colA.str.startswith("C").any():
                    chosen = s
                    break
            if chosen is None:
                return {"aplica": False, "error": "No se encontró hoja con contactores (C) en col A."}
            df = pd.read_excel(book, sheet_name=chosen, header=None, dtype=object)

        # Corriente objetivo
        objetivo = corriente_compresor * 1.15
        if arr == "PARTIDO":
            objetivo /= 2.0

        # Recorre TODAS las filas donde A empieza con 'C'
        best_model = None
        best_val = None

        nrows = len(df.index)
        for i in range(nrows):
            a = str(df.iat[i, 0]).strip().upper() if i < len(df.index) else ""
            if not a.startswith("C"):
                continue  # ignorar filas que no sean contactores

            # Columna D (índice 3)
            if df.shape[1] <= 3:
                continue
            d_val = _num(df.iat[i, 3])
            if d_val is None:
                continue

            if d_val >= objetivo:
                # Columna C (índice 2) -> modelo
                model_cell = df.iat[i, 2] if df.shape[1] > 2 else ""
                best_model = "" if (model_cell is None or (isinstance(model_cell, float) and pd.isna(model_cell))) else str(model_cell).strip()
                best_val = d_val
                break

        if best_model:
            return {
                "aplica": True,
                "modelo": best_model,
                "cantidad": 1 if arr == "DIRECTO" else 2,
                "valor_col_D": float(best_val),
            }
        else:
            return {"aplica": False, "motivo": "Sin modelo que cumpla en columna D"}

    except Exception as e:
        return {"aplica": False, "error": str(e)}
