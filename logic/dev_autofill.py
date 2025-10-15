# logic/dev_autofill.py
from __future__ import annotations

import random
import re
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import pandas as pd

# ← cambia a False cuando termines de probar
DEV_AUTOFILL: bool = False

# ---------------------------------------------------------------------
# Página 1 (globals)
# ---------------------------------------------------------------------
def apply_autofill_globals(existing: Optional[Dict[str, object]]) -> Dict[str, object]:
    """
    Devuelve un dict de globals con valores por defecto de desarrollo.
    Si DEV_AUTOFILL=False, devuelve el dict original sin tocar.
    No pisa valores existentes: solo completa los que falten o estén vacíos.
    """
    globs = dict(existing or {})
    if not DEV_AUTOFILL:
        return globs

    defaults = {
        "tension_alimentacion": "220",
        "refrigerante": "R744",
        "norma_ap": "IEC",            # tu motor interpreta 'UL' vs 'no UL'; IEC entra como 'no UL'
        "tipo_compresores": "BITZER",
        "num_comp_paralelo": 2,
        "num_comp_media": 2,
        "num_comp_baja": 2,
        "marca_elem": "ABB",
        "marca_var": "SCHNEIDER",
        # por compatibilidad con lógicas existentes:
        "tension": "220",
        "t_ctl": "220",               # usado por 'TENSION CONTROL:' (tu motor toma solo dígitos)
        "step3_state": globs.get("step3_state", {}) or {},
    }

    for k, v in defaults.items():
        if _is_empty(globs.get(k)):
            globs[k] = v

    return globs


def _is_empty(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip() or v.strip() == "—"
    return False


# ---------------------------------------------------------------------
# Página 2 (step2) – compresores
# ---------------------------------------------------------------------
def apply_autofill_step2(existing: Optional[Dict[str, Dict[str, str]]],
                         basedatos_path: Path,
                         marca_compresor: str = "BITZER") -> Dict[str, Dict[str, str]]:
    """
    Si step2 está vacío y DEV_AUTOFILL=True:
      - Crea 6 grupos: 2 PARTIDO, 2 VARIADOR, 2 DIRECTO
      - Asigna modelos aleatorios tomados de la base de datos (si es posible)
      - Asigna corriente aleatoria (A) solo para display de Paso 4
    Si ya hay contenido, lo devuelve tal cual.
    """
    step2 = dict(existing or {})
    if not DEV_AUTOFILL or step2:
        return step2

    # plan de grupos (orden natural G, B, F)
    plan: List[Tuple[str, str]] = [
        ("G1", "PARTIDO"),
        ("G2", "PARTIDO"),
        ("B1", "VARIADOR"),
        ("B2", "VARIADOR"),
        ("F1", "DIRECTO"),
        ("G3", "DIRECTO"),
    ]

    models = _load_models_from_db(basedatos_path, prefer_brand=marca_compresor) or _FALLBACK_MODELS

    used = set()
    def pick_model() -> str:
        # elige sin repetir mientras haya disponibilidad
        pool = [m for m in models if m not in used] or models
        m = random.choice(pool)
        used.add(m)
        return m

    for key, arranque in plan:
        step2[key] = {
            "arranque": arranque,              # lo usa tu engine
            "marca_compresor": marca_compresor,
            "modelo_compresor": pick_model(),
            "corriente": _rnd_amp(),           # solo para header del paso 4/4.2
            # Puedes añadir aquí otros campos si tu UI los muestra por defecto
        }

    return step2


def _rnd_amp() -> str:
    return f"{random.uniform(6.0, 16.0):.1f}"  # A


# ---------------------------------------------------------------------
# Lectura heurística de modelos desde data/basedatos.xlsx
# ---------------------------------------------------------------------
_FALLBACK_MODELS = [
    "2JME-3", "4NES-12", "4PES-15", "6FE-50Y",
    "CSH8573-125Y", "CE6-25", "BSE32-40", "HSP15"
]

def _load_models_from_db(book: Path, prefer_brand: str = "") -> List[str]:
    """
    Intenta encontrar modelos de compresor en basedatos.xlsx.
    Estrategia:
      1) Busca hojas cuyo nombre sugiera compresores (COMPRES, BITZER, COMP).
      2) Recorre todas las celdas y extrae tokens "modelo" plausibles [A-Z0-9-_/\.],
         con al menos una letra y un dígito.
      3) Si detecta una columna 'MARCA' o similar (no suele haber headers en tu xlsx),
         no la usaremos; por eso opcionalmente filtramos por heurística del nombre de hoja.
    Si falla, devuelve [] para que el caller use el fallback.
    """
    try:
        xls = pd.ExcelFile(book)
    except Exception:
        return []

    # prioriza hojas con nombres que apunten a compresores o a la marca
    preferred: List[str] = []
    secondary: List[str] = []
    want = prefer_brand.upper().strip()
    for sn in xls.sheet_names:
        name_u = sn.upper()
        if any(tag in name_u for tag in ("COMPRES", "COMP", "BITZER", "COMPRESORES")):
            preferred.append(sn)
        else:
            secondary.append(sn)

    pool = preferred + secondary

    models: List[str] = []
    rx_model = re.compile(r"^[A-Z0-9][A-Z0-9\-\_/\.]{2,}$", re.IGNORECASE)

    for sn in pool:
        try:
            df = pd.read_excel(book, sheet_name=sn, header=None, dtype=str)
        except Exception:
            continue

        # si la hoja parece de la marca, mejor aún
        sheet_u = sn.upper()
        brand_bias = 2 if (want and want in sheet_u) else 1

        # escanea celdas
        local: Dict[str, int] = {}
        for v in df.values.flatten():
            s = "" if v is None else str(v).strip()
            if not s or s.upper() in ("SI", "NO", "UL", "IEC", "*"):
                continue
            if rx_model.match(s) and any(ch.isdigit() for ch in s) and any(ch.isalpha() for ch in s):
                # favorece si la hoja coincide con la marca preferida
                local[s] = local.get(s, 0) + brand_bias

        # añade al conjunto global, priorizando los más “votados”
        if local:
            models.extend(sorted(local, key=lambda k: -local[k]))

        # si ya juntamos bastantes, paramos
        if len(models) >= 60:
            break

    # dedup conservando orden
    dedup: List[str] = []
    seen = set()
    for m in models:
        if m not in seen:
            seen.add(m); dedup.append(m)
    return dedup
