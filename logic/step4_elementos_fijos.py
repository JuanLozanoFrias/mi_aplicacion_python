# logic/step4_elementos_fijos.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple, Any
import pandas as pd

from .util_excel import letter_to_index as col, cell, to_int
from .materiales_core import MaterialesCore  # núcleo que ya tienes

# Columnas de salida para la tabla "OTROS ELEMENTOS (FIJOS)"
COLUMNS_OUT = [
    "ITEM", "CÓDIGO", "MODELO", "NOMBRE", "DESCRIPCIÓN",
    "ICC 240 V (kA)", "ICC 480 V (kA)", "REFERENCIA", "TORQUE"
]


def _only_digits(s: object) -> str:
    return "".join(ch for ch in str(s or "") if ch.isdigit())


def _get_attr(ctx: Any, name: str, *alts: str, default: str = "") -> str:
    """Lee atributos del contexto siendo tolerante con nombres alternos (t_ctl/tension_control)."""
    for key in (name, *alts):
        if hasattr(ctx, key):
            val = getattr(ctx, key)
            return "" if val is None else str(val)
    return default


def cargar_elementos_fijos(
    basedatos_path: Path | str,
    ctx: Any,
) -> Tuple[List[List[str]], Dict[str, int], pd.DataFrame]:
    """
    Lee 'OPCIONES CO2' y genera:
      - out_rows: filas para la tabla 'OTROS ELEMENTOS (FIJOS)' (lista de listas COLUMNS_OUT)
      - totales_wxy: dict {'fase':int, 'neutro':int, 'tierra':int}
      - debug_df: DataFrame con trazas por fila (útil para diagnóstico)

    Reglas:
      - Recorre TODAS las filas desde la 2 (índice 1) hasta el final, sin filtrar por BA/BB.
      - Soporta fórmulas 'MARCA DE ELEMENTOS: ...' con '//' y multiplicadores '*n' o '*#'.
      - Aplica Cond1 (AD/AE/AF) y Cond2 (AG/AH/AI).
      - Si un comando trae flag de norma ["UL"] / ["IEC"], la norma efectiva del comando puede
        venir del Paso 3 (si existe la pregunta UL); si no, se usa norma_ap (Paso 1).
      - El ITEM mask (Z para IEC, AA para UL) se decide por **comando**, con fallback cruzado.
    """
    book = Path(basedatos_path)
    df_ops = pd.read_excel(book, sheet_name="OPCIONES CO2", header=None, dtype=str)
    core = MaterialesCore(book)

    # Índices relevantes en OPCIONES CO2
    w, x, y = col("W"), col("X"), col("Y")
    z, aa = col("Z"), col("AA")
    ab, ac = col("AB"), col("AC")
    ad, ae, af = col("AD"), col("AE"), col("AF")
    ag, ah, ai = col("AG"), col("AH"), col("AI")

    # Contexto tolerante
    norma = (_get_attr(ctx, "norma_ap") or "").strip().upper()
    marca_elementos = (_get_attr(ctx, "marca_elementos", "marca_elem") or "").strip()
    t_ctl = _only_digits(_get_attr(ctx, "t_ctl", "tension_control"))
    t_alim = _only_digits(_get_attr(ctx, "t_alim", "tension_alimentacion"))
    step3 = getattr(ctx, "step3_state", {}) or {}

    out_rows: List[List[str]] = []
    tot = {"fase": 0, "neutro": 0, "tierra": 0}
    debug_rows: List[Tuple[int, str, str, str]] = []

    def _resolve_item_mask_for_cmd(row_idx: int, norma_cmd: str) -> str:
        """Devuelve mask de ITEM para el comando. Si la celda primaria está vacía, intenta la alterna."""
        primary = (cell(df_ops, row_idx, aa) if norma_cmd == "UL" else cell(df_ops, row_idx, z)) or ""
        primary = primary.strip()
        if primary:
            return primary
        # Fallback cruzado por si llenaron la otra columna
        alt = (cell(df_ops, row_idx, z) if norma_cmd == "UL" else cell(df_ops, row_idx, aa)) or ""
        return alt.strip()

    # <<< Recorremos todas las filas >>>
    for i in range(1, df_ops.shape[0]):  # fila 2 (índice 1) hasta el final
        nombre = (cell(df_ops, i, ab) or "").strip()
        raw_formula = (cell(df_ops, i, ac) or "").strip()

        # Si la fila está totalmente vacía (ni nombre, ni fórmula, ni W/X/Y), saltamos
        if not (nombre or raw_formula or cell(df_ops, i, w) or cell(df_ops, i, x) or cell(df_ops, i, y)):
            continue

        # 1) Condición 1 (filtra fila): AD etiqueta, AE acción cuando sí, AF cuando no
        act1 = core.pick_action(df_ops, i, ad, ae, af, step3)
        skip1, brand1 = core.resolve_actions([act1])
        if skip1:
            debug_rows.append((i, nombre, "FILTRADA", "Cond1: *BORRA"))
            continue

        # 2) Si no hay fórmula, sumamos W/X/Y directos (si hay algo que sumar)
        if not raw_formula:
            fase   = to_int(cell(df_ops, i, w))
            neutro = to_int(cell(df_ops, i, x))
            tierra = to_int(cell(df_ops, i, y))
            if fase or neutro or tierra:
                tot["fase"]   += fase
                tot["neutro"] += neutro
                tot["tierra"] += tierra
                debug_rows.append((i, nombre, "WXY", "Sin fórmula, sumados directos"))
            else:
                debug_rows.append((i, nombre, "VACIA", "Sin fórmula ni W/X/Y"))
            continue

        # 3) Condición 2: AG etiqueta, AH acción si sí, AI acción si no
        c2_label = (cell(df_ops, i, ag) or "").strip()
        a_yes = (cell(df_ops, i, ah) or "").strip()
        a_no  = (cell(df_ops, i, ai) or "").strip()
        cond2_mult, brand2 = core.eval_cond2(c2_label, a_yes, a_no, step3, norma)
        cond2_uses_number = (a_yes == "#" or a_no == "#")

        # 4) Marca a usar
        brand_to_use = (brand2 or brand1 or marca_elementos).strip()
        if not brand_to_use:
            debug_rows.append((i, nombre, "SIN MARCA", "No se pudo resolver marca"))
            continue

        # 5) Ejecutar comandos (separados por //)
        commands = core.split_commands(raw_formula)
        if not commands:
            debug_rows.append((i, nombre, "CMD-VACIO", "No hay comandos tras split"))
            continue

        multiplicadores_sumados = 0
        appended_any = False

        for cmd_raw in commands:
            # --- split core y multiplicador
            expr_core, mult_tok = core.split_core_and_multiplier(cmd_raw)

            # --- detectar flags de norma dentro del comando y decidir la norma “efectiva”
            expr_clean, norma_flags = core.strip_norma_flags(expr_core)
            norma_for_cmd = norma
            if norma_flags:
                # Si existe la pregunta UL en el Paso 3, usamos esa elección; si no, caemos al Paso 1
                if isinstance(step3, dict) and any(MaterialesCore.norm(k) == "UL" for k in step3.keys()):
                    norma_for_cmd = "UL" if MaterialesCore._step3_yes(step3, "UL") else "IEC"
                else:
                    norma_for_cmd = norma  # fallback normal

            # --- ITEM mask por comando (según norma efectiva) con fallback cruzado
            item_mask_cmd = _resolve_item_mask_for_cmd(i, norma_for_cmd)

            # --- Ejecuta 'MARCA DE ELEMENTOS' con la norma efectiva
            rows = core.execute_command(
                expr=expr_clean,
                marca_elem=brand_to_use,
                norma_ap=norma_for_cmd,
                t_ctl=t_ctl,
                t_alim=t_alim,
            )
            if not rows:
                debug_rows.append((i, nombre, "CMD-SINMATCH", expr_clean))
                continue

            # --- Multiplicador robusto
            base_mult = 1
            if isinstance(mult_tok, int):
                base_mult = max(1, mult_tok)

            cond2_mult_int = max(0, int(cond2_mult))
            if mult_tok == "#":
                multiplicador = cond2_mult_int
            else:
                multiplicador = base_mult * (cond2_mult_int if cond2_uses_number else 1)

            if multiplicador <= 0:
                debug_rows.append((i, nombre, "MULT=0", expr_clean))
                continue

            # --- Volcado de filas
            for d in rows:
                for k in range(1, int(multiplicador) + 1):
                    item_code = (
                        item_mask_cmd.replace("#", str(k))
                        if item_mask_cmd else ""
                    )
                    out_rows.append([
                        item_code,        # ITEM (Z o AA según norma del comando; con fallback)
                        d.get("B", ""),   # CÓDIGO
                        d.get("C", ""),   # MODELO
                        nombre,           # NOMBRE (col AB)
                        d.get("H", ""),   # DESCRIPCIÓN
                        d.get("F", ""),   # ICC 240 V (kA)
                        d.get("G", ""),   # ICC 480 V (kA)
                        d.get("I", ""),   # REFERENCIA (siempre la de Excel)
                        d.get("L", ""),   # TORQUE
                    ])
                    appended_any = True

            multiplicadores_sumados += max(1, int(multiplicador))

        # 6) W/X/Y multiplicados por la suma de multiplicadores
        if appended_any:
            m = max(1, int(multiplicadores_sumados))
            tot["fase"]   += m * to_int(cell(df_ops, i, w))
            tot["neutro"] += m * to_int(cell(df_ops, i, x))
            tot["tierra"] += m * to_int(cell(df_ops, i, y))
            debug_rows.append((i, nombre, "OK", f"mult={m}"))

    dbg = pd.DataFrame(debug_rows, columns=["ROW", "NOMBRE", "ESTADO", "DETALLE"])
    return out_rows, tot, dbg
