# logic/cable_selector.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional, Tuple

import math
import pandas as pd


@dataclass
class CableResult:
    comp_key: str
    regla_arranque: str         # VARIADOR | DIRECTO | PARTIDO
    i_nominal: float            # A
    i_ajustada: float           # A (tras regla de arranque)
    cable: str                  # Columna A
    ampacidad: float            # Columna D (A)
    fila_excel: int             # Índice 1-based en Excel (sólo informativo)


class CableSelector:
    """
    Selecciona el cable de potencia a partir de la hoja 'CABLE POTENCIA'.

    Reglas de corriente ajustada por tipo de arranque:
      - VARIADOR: I * 1.25
      - DIRECTO : I * 1.25
      - PARTIDO : (I * 1.15) / 2

    Búsqueda:
      - Compara I_ajustada contra la columna D (ampacidad).
      - Selecciona la PRIMERA fila con D >= I_ajustada.
      - ***Recorre desde la fila 3 de Excel*** (es decir, índice 0-based = 2).
      - Retorna columna A (cable) y D (ampacidad).
    """

    def __init__(self, basedatos_path: Path | str):
        self._book = Path(basedatos_path)
        self._df_cache: Dict[str, Optional[pd.DataFrame]] = {}

    # --------------------------- API pública ---------------------------

    def for_comp_key(self, comp_key: str, comp_state: Dict[str, str]) -> Optional[CableResult]:
        """
        comp_state: diccionario del Paso 2 para ese compresor (espera 'arranque' y 'corriente').
        """
        arr_txt = (comp_state.get("arranque") or comp_state.get("tipo_arranque") or "").strip().upper()
        regla = self._normalize_arranque(arr_txt)
        if not regla:
            return None

        i_nom = self._to_float(comp_state.get("corriente"))
        if i_nom is None or i_nom <= 0:
            return None

        i_adj = self._corriente_ajustada(i_nom, regla)
        pick = self._pick_cable(i_adj)
        if pick is None:
            return None

        cable, amp, idx = pick
        return CableResult(
            comp_key=comp_key,
            regla_arranque=regla,
            i_nominal=i_nom,
            i_ajustada=i_adj,
            cable=cable,
            ampacidad=amp,
            fila_excel=idx
        )

    def batch_from_step2(self, step2: Dict[str, Dict[str, str]]) -> Dict[str, CableResult]:
        out: Dict[str, CableResult] = {}
        for k, st in (step2 or {}).items():
            res = self.for_comp_key(k, st or {})
            if res:
                out[k] = res
        return out

    # ---------------------- helpers de negocio ------------------------

    @staticmethod
    def _normalize_arranque(v: str) -> Optional[str]:
        t = v.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
        t = t.strip().upper()
        if t in ("V", "VARIADOR", "VFD", "DRIVE"):  return "VARIADOR"
        if t in ("D", "DIRECTO", "DIRECT"):         return "DIRECTO"
        if t in ("P", "PARTIDO", "STAR-DELTA", "ESTRELLA-TRIANGULO", "Y-Δ", "Y-DELTA"):  return "PARTIDO"
        return None

    @staticmethod
    def _corriente_ajustada(i: float, regla: str) -> float:
        if regla == "VARIADOR":
            return i * 1.25
        if regla == "DIRECTO":
            return i * 1.25
        if regla == "PARTIDO":
            return (i * 1.15) / 2.0
        return i

    def _pick_cable(self, i_ajustada: float) -> Optional[Tuple[str, float, int]]:
        df = self._get_sheet("CABLE POTENCIA")
        if df is None or df.shape[1] < 4:
            return None

        colA, colD = 0, 3  # A=0, D=3 (0-based)
        START = 2          # <<< Fila 3 de Excel (0-based index)

        best_row = None

        # Recorre desde fila 3 (index 2)
        for i in range(START, df.shape[0]):
            raw_amp = df.iat[i, colD]
            amp = self._to_float(raw_amp)
            if amp is None:
                continue
            if amp >= i_ajustada:
                best_row = i
                break

        if best_row is None:
            # Si no existe >=, toma la última fila válida (mayor disponible), empezando también en fila 3.
            last_idx, last_amp = None, -math.inf
            for i in range(START, df.shape[0]):
                amp = self._to_float(df.iat[i, colD])
                if amp is not None and amp > last_amp:
                    last_amp = amp
                    last_idx = i
            if last_idx is None:
                return None
            best_row = last_idx

        cable = str(df.iat[best_row, colA]).strip() if not pd.isna(df.iat[best_row, colA]) else ""
        amp = float(self._to_float(df.iat[best_row, colD]) or 0.0)
        return (cable, amp, best_row + 1)  # +1 para índice 1-based en Excel

    # ----------------------------- Excel ------------------------------

    def _get_sheet(self, name: str) -> Optional[pd.DataFrame]:
        if name in self._df_cache:
            return self._df_cache[name]
        try:
            df = pd.read_excel(self._book, sheet_name=name, header=None, dtype=str)
            self._df_cache[name] = df
            return df
        except Exception:
            self._df_cache[name] = None
            return None

    # ---------------------------- util num ----------------------------

    @staticmethod
    def _to_float(v) -> Optional[float]:
        if v is None:
            return None
        s = str(v).strip()
        if s == "" or s == "—":
            return None
        # admite coma decimal
        s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            # extrae dígitos como "82.8 A"
            import re
            m = re.search(r"[-+]?\d+(?:[.,]\d+)?", s)
            if not m:
                return None
            try:
                return float(m.group(0).replace(",", "."))
            except Exception:
                return None

