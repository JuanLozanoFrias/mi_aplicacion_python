# logic/opciones_co2_engine.py
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


@dataclass
class ResumenTable:
    comp_key: str
    arranque: str
    title: str  # "<G#> — Arranque: <...>"
    rows: List[List[str]]  # [ITEM, CÓDIGO, MODELO, NOMBRE, DESCRIPCIÓN, ICC240, ICC480, REF, TORQUE]


class OpcionesCO2Engine:
    """
    Genera tablas del Paso 4 desde:
      - Hoja 'OPCIONES CO2' (BE/BF/BG/BH)
      - Estado del Paso 2 (step2)
      - Globals (marca_elem, norma_ap, t_ctl, t_alim, step3_state)

    Soporta:
      - *PONE / *BORRA
      - Overrides de marca (ABB / GENERICO / SCHNEIDER / RITTAL) vía BG/BH
      - BE con varios comandos separados por //
      - Comandos BE:
          • MARCA DE ELEMENTOS: ...
          • (VARIADOR CALCULADO) [con o sin "= B,C,..."]
          • (GUARDAMOTOR|CONTACTOR|BREAKER) CALCULADO [con o sin "= B,C,..."]

    Para (VARIADOR CALCULADO) busca en hojas VAR* (prioriza VARSCHNEIDER)
    y mapea SIEMPRE: B=Q/AQ (código), C=A/AA (modelo), respetando 220/460.

    Placeholders de 'MARCA DE ELEMENTOS':
      - D(@BREAKER_A)  -> usa el amperaje del breaker seleccionado (comparación ≥).
      - D(@GM_A)       -> amperaje guardamotor (comparación ≥).

    NUEVO:
      - Si hay placeholder en las condiciones, se devuelve **solo la primera coincidencia**
        (mínimo que cumpla) y se detiene la búsqueda.
    """

    def __init__(self, basedatos_path: Path):
        self._book = Path(basedatos_path)
        self._sheet_cache: Dict[str, Optional[pd.DataFrame]] = {}

        # Columnas en 'OPCIONES CO2'
        self.col_BA = self._col_letter_to_index("BA")  # Tipo arranque (V/D/P)
        self.col_BB = self._col_letter_to_index("BB")  # Ítem / código IEC
        self.col_BD = self._col_letter_to_index("BD")  # Nombre (etiqueta)
        self.col_BE = self._col_letter_to_index("BE")  # Fórmula/expresión BE
        self.col_BF = self._col_letter_to_index("BF")  # Etiqueta condicional
        self.col_BG = self._col_letter_to_index("BG")  # Acción si elige B
        self.col_BH = self._col_letter_to_index("BH")  # Acción si elige C

        self.arr_to_letter = {"VARIADOR": "V", "DIRECTO": "D", "PARTIDO": "P"}

    # =============================== API ===============================
    def build(self, step2: Dict[str, Dict[str, str]], globs: Dict[str, object]) -> List[ResumenTable]:
        try:
            df_ops = pd.read_excel(self._book, sheet_name="OPCIONES CO2", header=None, dtype=str)
        except Exception as e:
            raise RuntimeError(f"No pude leer la hoja 'OPCIONES CO2'. Detalle: {e}")

        marca_elem_global = (globs.get("marca_elem") or globs.get("marca_elementos") or "").strip()
        norma_ap = (globs.get("norma_ap") or "").strip().upper()
        t_ctl = self._only_digits(globs.get("t_ctl") or "")
        t_alim = self._only_digits(globs.get("t_alim") or "")
        step3_state = globs.get("step3_state") or {}

        if not step2:
            return []

        # orden natural G#, B#, F#
        def sort_key(k: str) -> Tuple[int, int]:
            g = {"G": 0, "B": 1, "F": 2}.get(k[:1].upper(), 9)
            try:
                n = int("".join(ch for ch in k[1:] if ch.isdigit()))
            except Exception:
                n = 999
            return (g, n)

        tables: List[ResumenTable] = []

        for comp_key in sorted(step2.keys(), key=sort_key):
            st = step2.get(comp_key, {}) or {}
            arranque = (st.get("arranque") or "").upper().strip()
            letter = self.arr_to_letter.get(arranque, "")
            if not letter:
                continue

            # --- variables dinámicas por compresor (p.ej. @BREAKER_A)
            dyn_vars = self._compute_dynamic_vars(st, marca_elem_global)

            rows_out: List[List[str]] = []

            # Recorrer desde fila 2 (index 1)
            for i in range(1, df_ops.shape[0]):  # desde fila 2 (0-based)
                if (self._cell(df_ops, i, self.col_BA).upper() or "") != letter:
                    continue

                item_code  = self._cell(df_ops, i, self.col_BB).strip()
                nombre_det = self._cell(df_ops, i, self.col_BD).strip()
                be_raw     = self._cell(df_ops, i, self.col_BE)
                if not be_raw:
                    continue

                # BG/BH según BF
                brand_override: Optional[str] = None
                bf_label = (self._cell(df_ops, i, self.col_BF) or "").strip()
                if bf_label:
                    choose_B = self._step3_selected_B(step3_state, bf_label)  # True->BG ; False->BH
                    directive_raw = self._cell(df_ops, i, self.col_BG if choose_B else self.col_BH)
                    action, payload = self._parse_directive(directive_raw)
                    if action == "BORRA":
                        continue
                    elif action == "BRAND":
                        brand_override = payload  # sólo aplica a 'MARCA DE ELEMENTOS'

                # Soportar múltiples comandos en BE (separador //)
                commands = [p.strip() for p in (be_raw or "").split("//") if p.strip()]
                for cmd in commands:
                    for d in self._be_to_rows(
                        cmd,
                        comp_key=comp_key,
                        comp_state=st,
                        marca_elem_base=marca_elem_global,
                        brand_override=brand_override,
                        norma_ap=norma_ap,
                        t_ctl=t_ctl,
                        t_alim=t_alim,
                        dyn_vars=dyn_vars,
                    ):
                        rows_out.append([
                            f"{comp_key} {item_code}",
                            d.get("B",""), d.get("C",""), nombre_det, d.get("H",""),
                            d.get("F",""), d.get("G",""), d.get("I",""), d.get("L",""),
                        ])

            if rows_out:
                tables.append(ResumenTable(
                    comp_key=comp_key,
                    arranque=arranque,
                    title=f"{comp_key} — Arranque: {arranque}",
                    rows=rows_out,
                ))

        return tables

    # ============================ BE resolver ============================
    def _be_to_rows(
        self,
        be: str,
        *,
        comp_key: str,
        comp_state: Dict[str, str],
        marca_elem_base: str,
        brand_override: Optional[str],
        norma_ap: str,
        t_ctl: str,
        t_alim: str,
        dyn_vars: Dict[str, str],
    ) -> List[Dict[str, str]]:
        out: List[Dict[str, str]] = []
        rule = (be or "").strip()
        if not rule:
            return out
        rule_u = rule.upper()

        # ---- (VARIADOR CALCULADO)  (con o sin "= B,C,...") ----
        m_vfd = re.match(
            r'^\(?\s*VARIADOR\s+CALCULADO\s*\)?\s*(?:=\s*([A-Z](?:\s*,\s*[A-Z])*)\s*)?$',
            rule_u
        )
        if m_vfd:
            # Nota: ret_cols no gobierna B/C; siempre mapeamos B=CÓDIGO y C=MODELO
            ret_cols = self._parse_col_list(m_vfd.group(1)) or ["A","Q"]

            var_sel = (comp_state.get("variador_sel") or "").strip()
            var_model = (comp_state.get("variador1") if var_sel == "V1"
                         else comp_state.get("variador2") if var_sel == "V2"
                         else comp_state.get("variador1") or comp_state.get("variador2") or "")
            var_model = self._clean_model_text(var_model)
            if not var_model:
                return out

            # Prioriza Schneider; incluye sinónimos
            candidate_var_sheets = ["VARSCHNEIDER", "VAR SCHNEIDER", "VARABB", "VAR ABB"]

            for sheet in candidate_var_sheets:
                d = self._lookup_vfd_in_var_sheet(sheet, var_model, t_alim, ret_cols)
                if d:
                    out.append(d)
                    return out
            return out

        # ---- (GM / CONTACTOR / BREAKER CALCULADO) (con o sin "= B,C,...) ----
        m_calc = re.match(
            r'^\(?\s*(GUARDAMOTOR|CONTACTOR|BREAKER)\s+CALCULADO\s*\)?\s*(?:=\s*([A-Z](?:\s*,\s*[A-Z])*)\s*)?$',
            rule_u
        )
        if m_calc:
            kind = m_calc.group(1)
            ret_cols = self._parse_col_list(m_calc.group(2)) or ["B","C","F","G","H","I","L"]

            if kind == "BREAKER":
                df = self._get_sheet("ABB")
                if df is None:
                    return out
                modelo_txt = (comp_state.get("breaker") or comp_state.get("breaker_vfd") or "")
                modelo = self._clean_model_text(modelo_txt)
                if not modelo:
                    return out
                # primero, con familia B
                idx = self._find_by_family_and_model_relaxed(df, "B", modelo)
                # si no, sin familia (por si la col-A no está exactamente en 'B')
                if idx is None:
                    idx = self._find_by_family_and_model_relaxed(df, None, modelo)
                if idx is None:
                    return out
                out.append(self._row_dict(df, idx, ret_cols))
                return out

            # GM / CONTACTOR: marca base, con fallback a ABB
            sheet = (marca_elem_base or "").strip() or "ABB"
            df = self._get_sheet(sheet)
            if df is None:
                return out
            modelo_txt = (comp_state.get("guardamotor") if kind == "GUARDAMOTOR" else comp_state.get("contactor")) or ""
            modelo = self._clean_model_text(modelo_txt)
            if not modelo:
                return out
            fam = "G" if kind == "GUARDAMOTOR" else "C"

            idx = self._find_by_family_and_model_relaxed(df, fam, modelo)
            df_used = df
            if idx is None and sheet.upper() != "ABB":
                df2 = self._get_sheet("ABB")
                if df2 is not None:
                    idx2 = self._find_by_family_and_model_relaxed(df2, fam, modelo)
                    if idx2 is not None:
                        idx = idx2
                        df_used = df2

            if idx is None:
                return out
            out.append(self._row_dict(df_used, idx, ret_cols))
            return out

        # ---- MARCA DE ELEMENTOS: A(...) ... = B,C,F,G,H,I,L ----
        if rule_u.startswith("MARCA DE ELEMENTOS:"):
            before_eq, cols_part = self._split_once(rule[len("MARCA DE ELEMENTOS:"):], "=")
            conds   = before_eq.strip()
            retcols = self._parse_col_list(cols_part) or ["B","C","F","G","H","I","L"]

            # Fallback robusto a ABB si no vino marca
            brand = (brand_override or marca_elem_base or "ABB").strip()
            df = self._get_sheet(brand)
            if df is None:
                return out

            cond_list = self._parse_conditions(conds)
            # ¿Alguna condición usa placeholder @... ?
            has_placeholder = any((v or "").strip().startswith("@") for _, v in cond_list)

            def append_matches(parsed_conds, ret_cols_local):
                # desde la fila 2 (0-based)
                for i in range(1, df.shape[0]):
                    if self._row_matches(df, i, parsed_conds, norma_ap, t_ctl, dyn_vars):
                        out.append(self._row_dict(df, i, ret_cols_local))
                        if has_placeholder:
                            # devolver solo la primera coincidencia cuando hay placeholder (p.ej. D(@BREAKER_A))
                            return True
                return False

            stop = append_matches(cond_list, retcols)

            # Regla A(R) ⇒ añadir A(RB) solo si no se detuvo por placeholder
            if not stop:
                has_A_R  = any(c == "A" and (v or "").strip().upper() == "R"  for c, v in cond_list)
                has_A_RB = any(c == "A" and (v or "").strip().upper() == "RB" for c, v in cond_list)
                if has_A_R and not has_A_RB:
                    append_matches([("A", "RB")], ["B","C","F","G","H","I","L"])

            return out

        return out  # regla desconocida

    # ===================== util/directivas/step3 ======================
    def _parse_directive(self, raw: str) -> Tuple[str, Optional[str]]:
        v = (raw or "").strip()
        if not v:
            return ("NONE", None)
        t = self._norm(v)
        t0 = re.sub(r"\s+", "", t)

        if re.search(r"\*\s*BORRA", t, re.IGNORECASE):   return ("BORRA", None)
        if re.search(r"\*\s*PONE",  t, re.IGNORECASE):   return ("PONE",  None)

        if re.search(r"\bABB\b", t):                     return ("BRAND", "ABB")
        if "GENERICO" in t or "GENÉRICO" in t or "GENERICO" in t0:
            return ("BRAND", "GENERICO")
        if re.search(r"\bSCHNEIDER(\s+ELECTRIC)?\b", t):
            return ("BRAND", "SCHNEIDER")
        if re.search(r"\bRITTAL\b", t):                  # <<< NUEVO
            return ("BRAND", "RITTAL")
        return ("NONE", None)

    def _step3_selected_B(self, st: Dict[str, str], label: str) -> bool:
        rec = self._get_step3_record(st, label)

        if isinstance(rec, dict):
            if isinstance(rec.get("value_bool"), bool): return bool(rec["value_bool"])
            if isinstance(rec.get("bool"), bool):       return bool(rec["bool"])
            col = str(rec.get("col", "")).strip().upper()
            if col in ("B", "C"): return col == "B"
            for k in ("index","idx","selected"):
                if k in rec:
                    try: return int(rec[k]) == 0
                    except Exception: pass
            if "value" in rec and isinstance(rec["value"], str):
                v = self._norm(rec["value"])
                if v in ("SI","SÍ","TRUE","1","PRIMERO","PRIMERA","OPCION 1","OPTION 1","B"): return True
                if v in ("NO","FALSE","0","SEGUNDO","SEGUNDA","OPCION 2","OPTION 2","C"):     return False

        if isinstance(rec, str):
            v = self._norm(rec)
            if v in ("SI","SÍ","TRUE","1","B","COL B","COL_B","LEFT","IZQUIERDA","OPCION 1","OPTION 1","PRIMERO","PRIMERA"): return True
            if v in ("NO","FALSE","0","C","COL C","COL_C","RIGHT","DERECHA","OPCION 2","OPTION 2","2","SEGUNDO","SEGUNDA"):   return False

        parts = [p.strip() for p in re.split(r"[\/|]| O ", label, flags=re.IGNORECASE) if p.strip()]
        if len(parts) >= 2 and isinstance(rec, str):
            v = self._norm(rec)
            if v == self._norm(parts[0]): return True
            if v == self._norm(parts[1]): return False

        return True  # por defecto B

    def _get_step3_record(self, st: Dict[str, str], label: str):
        if not isinstance(st, dict):
            return None
        want = self._norm(label)
        for k, v in st.items():
            if self._norm(k) == want:
                return v
        return None

    # ============================== Excel ==============================
    def _get_sheet(self, name: str) -> Optional[pd.DataFrame]:
        """Carga la hoja `name` con fallback por coincidencia normalizada."""
        key = (name or "").strip()
        if not key:
            return None
        if key in self._sheet_cache:
            return self._sheet_cache[key]

        # 1) intento directo
        try:
            df = pd.read_excel(self._book, sheet_name=key, header=None, dtype=str)
            self._sheet_cache[key] = df
            return df
        except Exception:
            pass

        # 2) intento por coincidencia normalizada (p.ej. "SCHNEIDER" vs "SCHNEIDER ELECTRIC")
        try:
            xls = pd.ExcelFile(self._book)
            key_n = self._norm_key(key)
            found_name: Optional[str] = None

            # prioridad: igualdad normalizada; luego "contiene"
            for s in xls.sheet_names:
                sn = self._norm_key(s)
                if sn == key_n:
                    found_name = s
                    break
            if not found_name:
                for s in xls.sheet_names:
                    sn = self._norm_key(s)
                    if key_n in sn or sn in key_n:
                        found_name = s
                        break

            if found_name:
                df = pd.read_excel(xls, sheet_name=found_name, header=None, dtype=str)
                self._sheet_cache[key] = df
                return df
        except Exception:
            pass

        self._sheet_cache[key] = None
        return None

    @staticmethod
    def _norm_key(s: str) -> str:
        t = (s or "").upper()
        trans = str.maketrans("ÁÉÍÓÚÜÑáéíóúüñ", "AEIOUUNaeiouun")
        t = t.translate(trans)
        return re.sub(r"[^A-Z0-9]", "", t)

    @staticmethod
    def _col_letter_to_index(letter: str) -> int:
        lt = (letter or "").upper().strip()
        n = 0
        for ch in lt:
            if not ('A' <= ch <= 'Z'):
                return 0
            n = n * 26 + (ord(ch) - ord('A') + 1)
        return max(0, n - 1)

    @staticmethod
    def _cell(df: pd.DataFrame, i: int, j: int) -> str:
        if i < 0 or j < 0 or i >= df.shape[0] or j >= df.shape[1]:
            return ""
        v = df.iat[i, j]
        return "" if pd.isna(v) else str(v).strip()

    @staticmethod
    def _split_once(s: str, sep: str) -> Tuple[str, str]:
        parts = s.split(sep, 1)
        return (parts[0], parts[1] if len(parts) > 1 else "")

    @staticmethod
    def _parse_col_list(s: str) -> List[str]:
        s = (s or "").strip()
        if not s:
            return []
        return [p.strip().upper() for p in s.split(",") if p.strip()]

    @staticmethod
    def _parse_conditions(s: str) -> List[Tuple[str, str]]:
        out: List[Tuple[str, str]] = []
        for m in re.finditer(r'([A-Z]+)\s*\(\s*(".*?"|[^)]+)\s*\)', s or ""):
            col = m.group(1).upper()
            val = (m.group(2) or "").strip()
            if len(val) >= 2 and val[0] in "\"'" and val[-1] == val[0]:
                val = val[1:-1]
            out.append((col, val))
        return out

    def _row_matches(
        self,
        df: pd.DataFrame,
        i: int,
        conds: List[Tuple[str, str]],
        norma_ap: str,
        t_ctl: str,
        dyn_vars: Optional[Dict[str, str]] = None,
    ) -> bool:
        """
        Coincidencia de fila para 'MARCA DE ELEMENTOS'.
        - Soporta etiquetas UL/IEC y 'TENSION CONTROL:'.
        - Si el selector viene de @VAR, la comparación numérica usa >=,
          si es literal numérico la comparación es por igualdad exacta.
        """
        dyn_vars = dyn_vars or {}

        for col_letter, selector in conds:
            j = self._col_letter_to_index(col_letter)
            cell = (self._cell(df, i, j) or "").strip()
            if cell == "*":
                continue

            sel_raw = (selector or "").strip()
            sel_u = sel_raw.upper()

            # Tensión control
            if sel_u == "TENSION CONTROL:":
                if self._only_digits(cell) == self._only_digits(t_ctl):
                    continue
                return False

            if sel_u in ("UL", "IEC"):
                want = "SI" if (norma_ap or "").upper() == "UL" else "NO"
                if cell.upper() == want:
                    continue
                return False

            # --- Placeholders @VAR ---
            from_var = False
            if sel_raw.startswith("@"):
                key = sel_raw[1:].strip().upper()
                sel_val = dyn_vars.get(key, "")
                from_var = True
            else:
                sel_val = sel_raw

            sd = self._only_digits(sel_val)
            cd = self._only_digits(cell)

            if sd and cd:
                try:
                    sdv = float(sd)
                    cdv = float(cd)
                except Exception:
                    return False
                if from_var:
                    # ceiling (>=) cuando el valor proviene de un placeholder
                    if cdv >= sdv:
                        continue
                    return False
                else:
                    # literal numérico → igualdad exacta
                    if cdv == sdv:
                        continue
                    return False

            if cell.upper() == sel_u:
                continue
            return False
        return True

    def _row_dict(self, df: pd.DataFrame, i: int, ret_cols: List[str]) -> Dict[str, str]:
        d: Dict[str, str] = {}
        for col in ("B","C","F","G","H","I","L"):
            j = self._col_letter_to_index(col)
            d[col] = self._cell(df, i, j) if col in ret_cols else ""
        return d

    # ========================== matching modelos ==========================
    @staticmethod
    def _clean_model_text(lbl: str) -> str:
        s = (lbl or "").strip()
        if not s or s == "—":
            return ""
        s = s.split("(", 1)[0]
        s = s.split(" x", 1)[0]
        return s.strip()

    @staticmethod
    def _norm_model(s: str) -> str:
        # sólo letras y números, sin espacios/guiones/puntos
        return re.sub(r"[^A-Z0-9]", "", (s or "").upper())

    def _find_by_family_and_model_relaxed(
        self,
        df: pd.DataFrame,
        family_letter: Optional[str],
        model: str
    ) -> Optional[int]:
        """
        Busca por modelo en col-C con matching RELAJADO:
          - normaliza col-C y modelo (A-Z0-9)
          - acepta contains en ambos sentidos (colC in model o model in colC)
        y, si se indica, exige familia (col-A) que empiece por family_letter.
        Recorre desde la fila 2 (index 1) hacia abajo.
        """
        if not model:
            return None

        target = self._norm_model(model)
        if not target:
            return None

        for i in range(1, df.shape[0]):  # desde fila 2 (0-based)
            fam_val = str(df.iat[i, 0]).strip().upper()         # A
            c_val_norm = self._norm_model(str(df.iat[i, 2]))    # C normalizada
            if not c_val_norm:
                continue
            if (target in c_val_norm) or (c_val_norm in target):
                if family_letter:
                    if fam_val.startswith(family_letter.upper()):
                        return int(i)
                else:
                    return int(i)
        return None

    # ---------- lookup variador en hoja VAR* (fila 5 en adelante) ----------
    def _lookup_vfd_in_var_sheet(
        self,
        sheet_name: str,
        var_model: str,
        t_alim: str,
        ret_cols: List[str],
    ) -> Optional[Dict[str, str]]:
        df = self._get_sheet(sheet_name)
        if df is None:
            return None

        # Columnas según tensión
        t = self._only_digits(t_alim) or "460"
        model_col = "A" if t == "220" else "AA"
        code_col  = "Q" if t == "220" else "AQ"

        j_model = self._col_letter_to_index(model_col)
        j_code  = self._col_letter_to_index(code_col)

        target = (var_model or "").strip().upper()
        if not target:
            return None

        # Recorrer desde la fila 5 (index 4)
        for i in range(4, df.shape[0]):
            mv = self._cell(df, i, j_model)
            if not mv:
                continue
            if mv.strip().upper() == target:
                return {
                    "B": self._cell(df, i, j_code) or "",
                    "C": mv,
                    "F": "", "G": "", "H": "", "I": "", "L": ""
                }

        # Fallback: matching relajado
        target_norm = self._norm_model(target)
        for i in range(4, df.shape[0]):
            mv = self._cell(df, i, j_model)
            if not mv:
                continue
            mvn = self._norm_model(mv)
            if (mvn == target_norm) or (target_norm in mvn) or (mvn in target_norm):
                return {
                    "B": self._cell(df, i, j_code) or "",
                    "C": mv,
                    "F": "", "G": "", "H": "", "I": "", "L": ""
                }
        return None

    @staticmethod
    def _norm(s: str) -> str:
        t = (s or "").strip().upper()
        trans = str.maketrans("ÁÉÍÓÚÜÑáéíóúüñ", "AEIOUUNaeiouun")
        return t.translate(trans)

    @staticmethod
    def _only_digits(s: str) -> str:
        return "".join(ch for ch in (s or "") if ch.isdigit())

    # ======================= dinámicos (@VAR) =======================
    def _compute_dynamic_vars(self, comp_state: Dict[str, str], marca_elem_base: str) -> Dict[str, str]:
        """
        Construye variables dinámicas por compresor:
          - BREAKER_A: corriente (col D) del breaker seleccionado en ABB.
          - GM_A:      corriente (col D) del guardamotor seleccionado (marca base o ABB).
        Devuelve valores como strings de dígitos (e.g., "125").
        """
        vars_map: Dict[str, str] = {}

        # --- BREAKER_A ---
        brk_model_txt = (comp_state.get("breaker") or comp_state.get("breaker_vfd") or "")
        brk_model = self._clean_model_text(brk_model_txt)
        if brk_model:
            df_abb = self._get_sheet("ABB")
            if df_abb is not None:
                idx = self._find_by_family_and_model_relaxed(df_abb, "B", brk_model)
                if idx is None:
                    idx = self._find_by_family_and_model_relaxed(df_abb, None, brk_model)
                if idx is not None:
                    jD = self._col_letter_to_index("D")
                    val = self._cell(df_abb, idx, jD)
                    vars_map["BREAKER_A"] = self._only_digits(val)

        # --- GM_A (opcional) ---
        gm_model_txt = comp_state.get("guardamotor") or ""
        gm_model = self._clean_model_text(gm_model_txt)
        if gm_model:
            brand = (marca_elem_base or "").strip() or "ABB"
            df_brand = self._get_sheet(brand)
            if df_brand is None:
                df_brand = self._get_sheet("ABB")
            if df_brand is not None:
                idx = self._find_by_family_and_model_relaxed(df_brand, "G", gm_model)
                if idx is None and (brand or "").upper() != "ABB":
                    df2 = self._get_sheet("ABB")
                    if df2 is not None:
                        idx2 = self._find_by_family_and_model_relaxed(df2, "G", gm_model)
                        if idx2 is not None:
                            df_brand, idx = df2, idx2
                if idx is not None:
                    jD = self._col_letter_to_index("D")
                    val = self._cell(df_brand, idx, jD)
                    vars_map["GM_A"] = self._only_digits(val)

        return vars_map

