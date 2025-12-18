from pathlib import Path
import json, re

path = Path('logic/tableros/step4_elementos_fijos.py')
lines = path.read_text(encoding='utf-8').splitlines()
start = None
end = None
for i,l in enumerate(lines):
    if '# ------------ Reglas JSON estructuradas (opcional) ------------' in l:
        start = i
        break
for i,l in enumerate(lines):
    if l.strip().startswith('def _resolve_item_mask_for_cmd'):
        end = i
        break
if start is None or end is None:
    raise SystemExit('start/end not found')

block = """    # ------------ Reglas JSON estructuradas (opcional) ------------
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
"""

new_lines = lines[:start] + block.splitlines() + lines[end:]
path.write_text('\n'.join(new_lines), encoding='utf-8')
print('done', len(lines), '->', len(new_lines))
