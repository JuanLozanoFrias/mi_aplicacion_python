# logic/step4_compresores.py
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

from .opciones_co2_engine import OpcionesCO2Engine, ResumenTable


def _norm(s: str) -> str:
    t = (s or "").strip().upper()
    trans = str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN")
    return t.translate(trans)


def _normalize_number(s: str) -> str:
    t = (s or "").strip()
    if not t:
        return ""
    t = t.replace(",", ".")
    m = re.search(r"[-+]?\d+(?:\.\d+)?", t)
    return m.group(0) if m else ""


def extract_comp_meta(st: Dict[str, Any], globs: Dict[str, Any]) -> Tuple[str, str, str]:
    """
    Extrae (marca, modelo, corrienteA) desde el dict del compresor y/o globals.
    Acepta tanto valores planos como dicts {'value': ...}.
    """

    def find_any(d: Dict[str, Any], candidates: List[str]) -> str:
        if not isinstance(d, dict):
            return ""
        for ck in candidates:
            for k, v in d.items():
                if _norm(k) == _norm(ck):
                    if isinstance(v, dict) and "value" in v:
                        v = v.get("value")
                    s = ("" if v is None else str(v)).strip()
                    if s and s != "—":
                        return s
        return ""

    model = (find_any(st, ["modelo_compresor", "modelo del compresor", "modelo compresor", "modelo", "model", "compressor model"])
             or find_any(globs, ["modelo_compresor", "compressor_model"]))
    brand = (find_any(st, ["marca_compresor", "marca del compresor", "marca", "brand", "fabricante"])
             or find_any(globs, ["marca_compresor", "compressor_brand"]))
    amps_raw = (find_any(st, ["corriente", "corriente nominal", "i nominal", "amperaje", "amps", "amp", "i"])
                or find_any(globs, ["corriente_compresor", "amps_compresor"]))
    return brand, model, _normalize_number(amps_raw)


def build_compresores_tables(
    basedatos_path: Path | str,
    step2_state: Dict[str, Dict[str, str]],
    ctx: Optional[object] = None,
) -> List[ResumenTable]:
    """
    Devuelve las tablas de compresores usando OpcionesCO2Engine.
    Construye los 'globs' que requiere el engine a partir del contexto.
    """
    book = Path(basedatos_path)
    engine = OpcionesCO2Engine(book)

    def _get(obj, *names, default=""):
        if obj is None:
            return default
        for n in names:
            if hasattr(obj, n):
                v = getattr(obj, n)
                return "" if v is None else v
        return default

    globs = {
        # marca de elementos (ABB/GENÉRICO/...)
        "marca_elem": str(_get(ctx, "marca_elementos", "marca_elem", default="")).strip(),
        # norma aplicada (UL/IEC)
        "norma_ap": str(_get(ctx, "norma_ap", default="")).strip().upper(),
        # tensiones que algunas fórmulas requieren
        "t_ctl": "".join(ch for ch in str(_get(ctx, "tension_control", "t_ctl", default="")) if ch.isdigit()),
        "t_alim": "".join(ch for ch in str(_get(ctx, "tension_alimentacion", "t_alim", default="")) if ch.isdigit()),
        # flags del paso 3 normalizados
        "step3_state": getattr(ctx, "step3_state", {}) if ctx is not None else {},
    }

    return engine.build(step2_state, globs)
