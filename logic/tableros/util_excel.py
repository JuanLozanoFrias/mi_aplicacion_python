# logic/util_excel.py
from __future__ import annotations
from pathlib import Path
from typing import Any, Optional
import pandas as pd
import math

__all__ = [
    "letter_to_index",
    "index_to_letter",
    "cell",
    "only_digits",
    "to_int",
    "read_excel_sheet",
]

def letter_to_index(letter: str) -> int:
    """
    Convierte letras de columna de Excel a índice base 0.
    Ej: A->0, Z->25, AA->26, BA->52.
    """
    lt = (letter or "").strip().upper()
    n = 0
    for ch in lt:
        if not ("A" <= ch <= "Z"):
            return 0
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return max(0, n - 1)

def index_to_letter(index: int) -> str:
    """Convierte índice base 0 a letras de columna de Excel (0->A, 25->Z, 26->AA)."""
    if index < 0:
        return "A"
    index += 1
    out = []
    while index:
        index, rem = divmod(index - 1, 26)
        out.append(chr(rem + ord("A")))
    return "".join(reversed(out))

def _is_nan(x: Any) -> bool:
    try:
        return bool(pd.isna(x))
    except Exception:
        try:
            return math.isnan(x)  # type: ignore[arg-type]
        except Exception:
            return False

def cell(df: pd.DataFrame, i: int, j: int) -> str:
    """
    Devuelve el valor de la celda como string limpio ('' si está fuera de rango o NaN).
    """
    if df is None or i < 0 or j < 0:
        return ""
    nrows, ncols = df.shape
    if i >= nrows or j >= ncols:
        return ""
    v = df.iat[i, j]
    return "" if _is_nan(v) else str(v).strip()

def only_digits(s: Any) -> str:
    """Extrae solo dígitos (útil para tensiones tipo '220 V')."""
    return "".join(ch for ch in str(s or "") if ch.isdigit())

def to_int(s: Any) -> int:
    """Convierte a entero tomando solo dígitos (''->0)."""
    t = only_digits(s)
    return int(t) if t else 0

def read_excel_sheet(book: Path | str, sheet_name: str, *, header: Optional[int] = None, dtype: Any = str) -> pd.DataFrame:
    """
    Helper para leer una hoja de Excel con defaults que usamos en el proyecto.
    """
    path = Path(book)
    return pd.read_excel(path, sheet_name=sheet_name, header=header, dtype=dtype)  # type: ignore[call-arg]

