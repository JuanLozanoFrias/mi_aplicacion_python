# -*- coding: utf-8 -*-
# logic/contactor.py
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional
import pandas as pd
import re


# ---------------- util ----------------
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


# ---------------- selección clásica (se mantiene por compatibilidad) ----------------
def seleccionar_contactor_abb(book_path: str, corriente_compresor: float, tipo_arranque: str) -> Dict:
    """
    Contactor:
      - DIRECTO o PARTIDO.
      - Filas con 'C...' en col A.
      - En col D: primer valor >= I*1.15 (directo) o >= I*1.15/2 (partido).
      - Devuelve col C (modelo). Reporta valor de col D y cantidad sugerida.
    """
    arr = (tipo_arranque or "").strip().upper()
    if arr not in ("DIRECTO", "PARTIDO"):
        return {"aplica": False, "motivo": "NO APLICA"}

    try:
        book = Path(book_path)

        # Hoja ABB o fallback a la primera hoja con 'C' en col A
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

        # Recorre filas donde A empieza con 'C'
        best_model = None
        best_val = None

        nrows = len(df.index)
        for i in range(nrows):
            a = str(df.iat[i, 0]).strip().upper() if i < len(df.index) else ""
            if not a.startswith("C"):
                continue

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


# ---------------- NUEVO: lista de candidatos con Q/R (PUENTES) ----------------
def listar_contactores_abb(
    book_path: str,
    corriente_compresor: float,
    tipo_arranque: str,
) -> List[Dict[str, object]]:
    """
    Devuelve lista ordenada de contactores ABB que cumplen col-D >= objetivo
    con Q (PUENTES) y R (CÓDIGO PUENTES).

    Hoja con header=None:
      A=0 (código inicia con 'C'), C=2 (MODELO), D=3 (capacidad),
      Q=16 (PUENTES), R=17 (CÓDIGO PUENTES).
    """
    arr = (tipo_arranque or "").strip().upper()
    if arr not in ("DIRECTO", "PARTIDO"):
        return []

    try:
        # Hoja ABB o fallback a la primera hoja con 'C' en col A
        try:
            df = pd.read_excel(book_path, sheet_name="ABB", header=None, dtype=object)
        except Exception:
            xls = pd.ExcelFile(book_path)
            chosen = None
            for s in xls.sheet_names:
                tmp = pd.read_excel(xls, sheet_name=s, header=None, dtype=object)
                colA = tmp.iloc[:, 0].astype(str).str.upper().str.strip()
                if colA.str.startswith("C").any():
                    chosen = s
                    break
            if chosen is None:
                return []
            df = pd.read_excel(book_path, sheet_name=chosen, header=None, dtype=object)

        objetivo = corriente_compresor * 1.15
        if arr == "PARTIDO":
            objetivo /= 2.0

        colA = df.iloc[:, 0].astype(str).str.upper().str.strip()
        mask_c = colA.str.startswith("C")
        df_c = df[mask_c].copy()

        dnum = df_c.iloc[:, 3].apply(_num)
        df_c = df_c.assign(D_num=dnum)
        cand = df_c[df_c["D_num"].ge(objetivo)].sort_values("D_num")

        out: List[Dict[str, object]] = []
        for idx, row in cand.iterrows():
            model_cell = row.iloc[2] if df_c.shape[1] > 2 else ""
            modelo = "" if (model_cell is None or (isinstance(model_cell, float) and pd.isna(model_cell))) else str(model_cell).strip()
            q = "" if pd.isna(df.iat[idx, 16]) else str(df.iat[idx, 16]).strip()
            r = "" if pd.isna(df.iat[idx, 17]) else str(df.iat[idx, 17]).strip()
            out.append({
                "modelo": modelo,                 # col-C
                "valor_col_D": float(row["D_num"]) if row["D_num"] is not None else None,
                "fila_excel": int(idx) + 1,
                "puente_modelo": q,               # Q
                "puente_codigo": r,               # R (código interno)
                "cantidad": 1 if arr == "DIRECTO" else 2,
            })
        return out

    except Exception:
        return []

