# logic/step4_elementos_fijos.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Tuple, Any
import json
import pandas as pd
import re

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




def _load_ops_df(book: Path) -> pd.DataFrame:
    """
    Carga solo desde el cache JSON pre-generado (preguntas_opciones_co2.json).
    Ya no se lee el Excel ni se regenera el cache aquí.
    """
    cache_path = book.parent / "preguntas_opciones_co2.json"

    def _df_from_compact(payload: dict) -> pd.DataFrame:
        data = payload.get("data", [])
        ncols = int(payload.get("ncols", 0)) if payload.get("ncols") else 0
        if data and ncols:
            # asegura ancho homogéneo
            norm_rows = []
            for row in data:
                r = list(row) + [""] * max(0, ncols - len(row))
                norm_rows.append(r[:ncols])
            return pd.DataFrame(norm_rows)
        return pd.DataFrame()

    def _df_from_rows(rows, ncols):
        data = []
        for row in rows:
            r = []
            for c in range(ncols):
                r.append(row.get(str(c), ""))
            data.append(r)
        return pd.DataFrame(data)

    # leer solo cache (formato compacto preferido)
    try:
        data = json.loads(cache_path.read_text(encoding="utf-8"))
        if "data" in data and "ncols" in data:
            df = _df_from_compact(data)
            if not df.empty:
                return df
        elif "rows" in data and "cols" in data:
            rows = data.get("rows", [])
            ncols = int(data.get("cols", 0)) if data.get("cols") else 0
            if rows and ncols:
                return _df_from_rows(rows, ncols)
    except Exception:
        pass

    return pd.DataFrame()


def _df_from_reglas_json(reglas_path: Path, default_cols: int = 29) -> pd.DataFrame:
    """
    Construye un DataFrame mínimo a partir de reglas estructuradas.
    Solo usa columnas 0 (pregunta) y 1/2 (opciones) para convivir con el parser actual.
    """
    if not reglas_path.exists():
        return pd.DataFrame()
    try:
        data = json.loads(reglas_path.read_text(encoding="utf-8"))
        rows = []
        for bloque in data.get("bloques", []):
            for r in bloque.get("reglas", []):
                pregunta = r.get("pregunta", "").strip()
                if not pregunta:
                    continue
                alternativas = r.get("alternativas", [])
                if not alternativas:
                    # fallback simple
                    row = [""] * default_cols
                    row[0] = pregunta
                    row[28] = r.get("raw", "") or ""
                    rows.append(row)
                    continue
                for alt in alternativas:
                    row = [""] * default_cols
                    row[0] = pregunta
                    raw = alt.get("raw", "").strip()
                    if not raw:
                        # construir raw básico a partir de tokens y permitidos
                        parts = []
                        for t in alt.get("tokens", []):
                            op = t.get("op", "")
                            val = t.get("val", "")
                            if op and val:
                                if '"' in str(val):
                                    val = val.replace('"', '\\"')
                                parts.append(f"{op}({val})")
                        raw = "MARCA DE ELEMENTOS:" + "".join(parts)
                        perms = alt.get("permitidos", [])
                        if perms:
                            raw += "=" + ",".join(perms)
                        if alt.get("mult") == "#":
                            raw = f"({raw})*#"
                    row[28] = raw
                    rows.append(row)
        if rows:
            return pd.DataFrame(rows)
    except Exception:
        return pd.DataFrame()
    return pd.DataFrame()


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
    # 1) Intentar DataFrame principal desde Excel/cache
    df_ops = _load_ops_df(book)
    # 2) Complementar con reglas estructuradas (solo añade preguntas al df_ops si no existen)
    df_reglas = _df_from_reglas_json(book.parent / "opciones_co2_reglas.json")
    if not df_reglas.empty:
        if df_ops.empty:
            df_ops = df_reglas
        else:
            # agregar filas que no estén en df_ops (comparando columna 0: pregunta)
            existentes = set(str(x).strip().upper() for x in df_ops.iloc[:, 0].fillna(""))
            nuevos = []
            for _, row in df_reglas.iterrows():
                preg = str(row.iloc[0]).strip().upper()
                if preg and preg not in existentes:
                    nuevos.append(row.tolist())
            if nuevos:
                df_ops = pd.concat([df_ops, pd.DataFrame(nuevos)], ignore_index=True)

    # Normalizar: quitar prefijo literal "MARCA DE ELEMENTOS:" en fórmulas (col AC ~ índice 28)
    try:
        if not df_ops.empty and df_ops.shape[1] > 28:
            df_ops.iloc[:, 28] = df_ops.iloc[:, 28].astype(str).str.replace(r"^\\s*MARCA DE ELEMENTOS:\\s*", "", regex=True)
    except Exception:
        pass
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
    seen_json: set[tuple[str, str, str, str]] = set()

    # ------------ Reglas JSON estructuradas (opcional) ------------
    def _apply_json_rules() -> None:
        reglas_path = book.parent / "opciones_co2_reglas.json"
        if not reglas_path.exists():
            return
        try:
            data = json.loads(reglas_path.read_text(encoding="utf-8"))
        except Exception:
            return

        # norma efectiva: usar la del Paso 1 (ctx.norma_ap)
        norma_ctx = (norma or "IEC").strip().upper() or "IEC"

        for bloque in data.get("bloques", []):
            for regla in bloque.get("reglas", []):
                pregunta = (regla.get("pregunta") or "").strip()
                if not pregunta:
                    continue
                display_name = (regla.get("nombre") or bloque.get("id") or pregunta).strip()
                siempre_flag = bool(regla.get("siempre"))
                if siempre_flag or MaterialesCore.norm(pregunta) == "SIEMPRE":
                    es_si = True
                else:
                    es_si = MaterialesCore._step3_yes(step3, pregunta)

                alts = regla.get("alternativas", []) or []
                alts_match_norma = []
                alts_sin_norma = []
                for alt in alts:
                    n_alt = (alt.get("norma") or "").strip().upper()
                    if n_alt == norma_ctx:
                        alts_match_norma.append(alt)
                    elif not n_alt:
                        alts_sin_norma.append(alt)
                alts_to_apply = alts_match_norma or alts_sin_norma or alts
                if not alts_to_apply:
                    continue

                for chosen in alts_to_apply:
                    expr = (chosen.get("raw") or "").strip()
                    if not expr:
                        tokens = chosen.get("tokens", [])
                        parts = []
                        for t in tokens:
                            op = t.get("op", "")
                            val = t.get("val", "")
                            if op and val:
                                parts.append(f"{op}({val})")
                        if parts:
                            perms = chosen.get("permitidos", [])
                            expr = "".join(parts)
                            if perms:
                                expr += "=" + ",".join(perms)
                    if expr and not re.match(r"^\s*MARCA\s+DE\s+ELEMENTOS\s*:", expr, flags=re.IGNORECASE):
                        expr = "MARCA DE ELEMENTOS:" + expr

                    item_rule = regla.get("item", "") or ""
                    item_alt = chosen.get("item", "") or ""
                    item_no_rule = regla.get("item_no", "") or ""
                    item_no_alt = chosen.get("item_no", "") or ""
                    marca_override = (chosen.get("marca") or regla.get("marca") or "").strip()

                    accion = (regla.get("accion") or "pone").lower()
                    accion_no = (regla.get("accion_no") or "").lower()

                    item_use = item_alt or item_rule
                    item_no_use = item_no_alt or item_no_rule

                    if es_si:
                        if accion == "borra" or not item_use:
                            continue
                        marca_use = (marca_override or marca_elementos or "").strip().upper()
                        rows_exec = core.execute_command(
                            expr=expr,
                            marca_elem=marca_use,
                            norma_ap=norma_ctx,
                            t_ctl=t_ctl,
                            t_alim=t_alim,
                        )
                        if not rows_exec and marca_override:
                            rows_exec = core.execute_command(
                                expr=expr,
                                marca_elem=(marca_elementos or "").strip().upper(),
                                norma_ap=norma_ctx,
                                t_ctl=t_ctl,
                                t_alim=t_alim,
                            )
                        if rows_exec:
                            allow_multi = bool(chosen.get('allow_multi') or regla.get('allow_multi'))
                            if not allow_multi:
                                rows_exec = rows_exec[:1]
                        if not rows_exec:
                            debug_rows.append(
                                (-1, display_name, "JSON-SINMATCH", expr)
                            )
                            continue
                        multiplicador = 1
                        if chosen.get("mult") == "#":
                            label_mult = (
                                regla.get("pregunta_mult")
                                or chosen.get("pregunta_mult")
                                or (pregunta if MaterialesCore.norm(pregunta) != "SIEMPRE" else "")
                                or (bloque.get("id") if isinstance(bloque, dict) else "")
                                or ""
                            ).strip()
                            def _find_number(lbl: str) -> int | None:
                                if not lbl:
                                    return None
                                n0 = MaterialesCore._step3_number(step3, lbl)
                                if n0 is not None:
                                    return n0
                                if isinstance(step3, dict):
                                    want = MaterialesCore.norm(lbl)
                                    for k, v in step3.items():
                                        kn = MaterialesCore.norm(k)
                                        if want in kn or kn in want:
                                            try:
                                                s = str(v.get("value") if isinstance(v, dict) else v)
                                                return int(float(s.replace(",", ".")))
                                            except Exception:
                                                continue
                                    want_sub = "ILUMIN"
                                    if want_sub in MaterialesCore.norm(lbl):
                                        for k, v in step3.items():
                                            if want_sub in MaterialesCore.norm(k):
                                                try:
                                                    s = str(v.get("value") if isinstance(v, dict) else v)
                                                    return int(float(s.replace(",", ".")))
                                                except Exception:
                                                    continue
                                return None

                            n = _find_number(label_mult)
                            multiplicador = max(1, n if n is not None else 1)
                        for d in rows_exec:
                            for k in range(1, multiplicador + 1):
                                item_k = item_use.replace("#", str(k)) if "#" in item_use else item_use
                                key = (item_k, d.get("B", ""), d.get("C", ""), display_name)
                                if key in seen_json:
                                    continue
                                seen_json.add(key)
                                out_rows.append([
                                    item_k,
                                    d.get("B", ""),
                                    d.get("C", ""),
                                    display_name,
                                    d.get("H", ""),
                                    d.get("F", ""),
                                    d.get("G", ""),
                                    d.get("I", ""),
                                    d.get("L", ""),
                                ])
                        bor = chosen.get("bornera") or regla.get("bornera") or {}
                        tot["fase"]   += int(bor.get("fase", 0)) * multiplicador
                        tot["neutro"] += int(bor.get("neutro", 0)) * multiplicador
                        tot["tierra"] += int(bor.get("tierra", 0)) * multiplicador
                        debug_rows.append((-1, display_name, "JSON-OK", f"{expr} x{multiplicador}"))
                    else:
                        if accion_no != "borra" and item_no_use:
                            out_rows.append([item_no_use, "", "", pregunta, "", "", "", "", ""])
                        debug_rows.append((-1, display_name, "JSON-NO", accion_no or "skip"))

    _apply_json_rules()
    # Solo usamos reglas JSON: devolver resultados sin procesar la hoja de Excel
    dbg = pd.DataFrame(debug_rows, columns=["ROW", "NOMBRE", "ESTADO", "DETALLE"])
    return out_rows, tot, dbg

    def _resolve_item_mask_for_cmd(row_idx: int, norma_cmd: str) -> str:
        """Devuelve mask de ITEM para el comando. Si la celda primaria está vacía, intenta la alterna."""
        primary = (cell(df_ops, row_idx, aa) if norma_cmd == "UL" else cell(df_ops, row_idx, z)) or ""
        primary = primary.strip()
        if primary:
            return primary
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
            expr_core, mult_tok = core.split_core_and_multiplier(cmd_raw)
            expr_clean, norma_flags = core.strip_norma_flags(expr_core)
            norma_for_cmd = norma
            if norma_flags:
                if isinstance(step3, dict) and any(MaterialesCore.norm(k) == "UL" for k in step3.keys()):
                    norma_for_cmd = "UL" if MaterialesCore._step3_yes(step3, "UL") else "IEC"
                else:
                    norma_for_cmd = norma

            item_mask_cmd = _resolve_item_mask_for_cmd(i, norma_for_cmd)
            if not item_mask_cmd:
                debug_rows.append((i, nombre, "SIN ITEM", f"{expr_clean}"))
                continue

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

            for d in rows:
                for k in range(1, int(multiplicador) + 1):
                    item_code = item_mask_cmd.replace("#", str(k)) if item_mask_cmd else ""
                    out_rows.append([
                        item_code,
                        d.get("B", ""),
                        d.get("C", ""),
                        nombre,
                        d.get("H", ""),
                        d.get("F", ""),
                        d.get("G", ""),
                        d.get("I", ""),
                        d.get("L", ""),
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
