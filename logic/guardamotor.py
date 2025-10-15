# logic/guardamotor.py
from __future__ import annotations
from typing import Any, Dict, List, Iterable, Optional
import re
import pandas as pd

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

def seleccionar_guardamotor_abb(
    ruta_excel: str,
    corriente_compresor_a: float,
    tipo_arranque: str,
) -> Dict[str, Any]:
    """
    Selecciona guardamotor para ABB según tu regla:
    - DIRECTO: Iaj = I * 1.15        -> cantidad 1
    - PARTIDO: Iaj = I * 1.15 / 2    -> cantidad 2
    Búsqueda:
      1) Encontrar primera fila con col-A que empiece por 'G'
      2) Desde ahí, en col-E encontrar primer valor >= Iaj
      3) Devolver modelo (col-C) y datos útiles
    Columnas: A=0, C=2, E=4 (sin encabezados).
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
