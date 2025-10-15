# logic/resumen_cables.py
from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional

from .cable_selector import CableSelector, CableResult


def build_cable_badges(step2: Dict[str, Dict[str, str]], basedatos_path: Path | str) -> Dict[str, str]:
    """
    Retorna un diccionario: comp_key -> " | CABLE: <A> | CAP: <D> A | ICOND: <x.x> A"
    (Texto listo para concatenar al rótulo de resumen).
    """
    selector = CableSelector(basedatos_path)
    picks = selector.batch_from_step2(step2)
    out: Dict[str, str] = {}
    for k, r in picks.items():
        # ICOND = corriente condicionada (la que se usó para el cálculo)
        out[k] = f"  |  CABLE: {r.cable}  |  CAP: {r.ampacidad:g} A  |  ICOND: {r.i_ajustada:.1f} A"
    return out


def pick_for_one(step2: Dict[str, Dict[str, str]], comp_key: str, basedatos_path: Path | str) -> Optional[str]:
    """
    Solo un compresor → devuelve el mismo badge de arriba o None si no encuentra.
    """
    selector = CableSelector(basedatos_path)
    res: Optional[CableResult] = selector.for_comp_key(comp_key, step2.get(comp_key, {}) or {})
    if not res:
        return None
    return f"  |  CABLE: {res.cable}  |  CAP: {res.ampacidad:g} A  |  ICOND: {res.i_ajustada:.1f} A"
