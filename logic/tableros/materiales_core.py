# logic/materiales_core.py
from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pandas as pd


class MaterialesCore:
    """
    Núcleo utilitario para Paso 4:
    - utilidades Excel
    - parser de 'MARCA DE ELEMENTOS'
    - condicionales (Cond1/Cond2)
    - conteo de borneras
    - lookup de breaker ABB por modelo (para '(BREAKER CALCULADO)')
    """

    def __init__(self, book: Path) -> None:
        self.book = Path(book)
        self._sheet_cache: Dict[str, Optional[pd.DataFrame]] = {}

    # --------------------------- Excel utils ---------------------------
    @staticmethod
    def col(letter: str) -> int:
        lt = (letter or "").strip().upper()
        n = 0
        for ch in lt:
            if not ("A" <= ch <= "Z"):
                return 0
            n = n * 26 + (ord(ch) - ord("A") + 1)
        return max(0, n - 1)

    @staticmethod
    def cell(df: pd.DataFrame, i: int, j: int) -> str:
        if i < 0 or j < 0 or i >= df.shape[0] or j >= df.shape[1]:
            return ""
        v = df.iat[i, j]
        return "" if pd.isna(v) else str(v).strip()

    @staticmethod
    def only_digits(s: str) -> str:
        return "".join(ch for ch in (s or "") if ch.isdigit())

    @staticmethod
    def to_int(s: str) -> int:
        t = "".join(ch for ch in (s or "") if ch.isdigit())
        return int(t) if t else 0

    @staticmethod
    def norm(s: str) -> str:
        """Mayúsculas, sin acentos y con espacios colapsados (compat con _to_flags)."""
        t = unicodedata.normalize("NFD", (s or "").strip())
        t = t.encode("ascii", "ignore").decode("ascii")
        t = t.upper()
        t = re.sub(r"\s+", " ", t)
        return t.strip()

    # --------------------------- Condicionales ---------------------------
    def pick_action(
        self,
        df_ops: pd.DataFrame,
        i: int,
        col_lab: int,
        col_yes: int,
        col_no: int,
        step3_state: Dict[str, object],
    ) -> str:
        label = (self.cell(df_ops, i, col_lab) or "").strip()
        if not label:
            return ""
        return (self.cell(df_ops, i, col_yes if self._step3_yes(step3_state, label) else col_no) or "").strip()

    @staticmethod
    def _normalize_action(txt: str) -> str:
        t = (txt or "").strip()
        u = MaterialesCore.norm(t)
        if "*BORRA" in u:
            return "BORRA"
        if "*PONE" in u:
            return "PONE"
        if u == "GENERICO":
            return "GENERICO"
        return u

    def resolve_actions(self, actions: List[str]) -> Tuple[bool, Optional[str]]:
        skip = False
        brand: Optional[str] = None
        for a in actions:
            aa = self._normalize_action(a)
            if aa == "BORRA":
                skip = True
            elif aa not in ("", "PONE"):
                brand = aa
        return (skip, brand)

    def eval_cond2(
        self,
        label: str,
        yes: str,
        no: str,
        step3_state: Dict[str, object],
        norma_ap: str,
    ) -> Tuple[int, Optional[str]]:
        """
        Devuelve (multiplicador, brand_override).

        - Si la rama elegida es '#': devuelve (n, None).
          **Fix**: si 'n' no existe o no es válido, asumimos **1** (antes 0).
        - Si es '*BORRA': devuelve (0, None).
        - Si es marca (ABB/GENÉRICO/...): (1, 'ABB').
        """
        if not label:
            return (1, None)

        labn = self.norm(label)
        yn = (yes or "").strip()
        nn = (no or "").strip()

        if labn == "UL":
            has_ul = isinstance(step3_state, dict) and any(self.norm(k) == "UL" for k in step3_state.keys())
            if has_ul:
                chosen = yn if self._step3_yes(step3_state, "UL") else nn
            else:
                chosen = yn if (norma_ap or "").strip().upper() == "UL" else nn
        else:
            chosen = yn if self._step3_yes(step3_state, label) else nn

        if chosen.strip() == "#":
            n = self._step3_number(step3_state, label)
            # ---- cambio clave: default a 1 si no hay valor en step3 ----
            if n is None:
                return (1, None)
            try:
                mult = int(float(str(n)))
            except Exception:
                mult = 1
            mult = max(0, mult)  # por si el usuario puso 0 o negativo
            return (mult, None)

        skip, brand = self.resolve_actions([chosen])
        if skip:
            return (0, None)
        return (1, brand)

    @staticmethod
    def _step3_yes(st: Dict[str, object], label: str) -> bool:
        if not isinstance(st, dict):
            return False
        want = MaterialesCore.norm(label)
        for k, v in st.items():
            if MaterialesCore.norm(k) == want:
                if isinstance(v, dict) and "value" in v:
                    v = v["value"]
                vv = MaterialesCore.norm(str(v))
                return vv in ("SI", "SÍ", "1", "TRUE")
        return False

    @staticmethod
    def _step3_number(st: Dict[str, object], label: str) -> Optional[int]:
        if not isinstance(st, dict):
            return None
        want = MaterialesCore.norm(label)
        for k, v in st.items():
            if MaterialesCore.norm(k) == want:
                if isinstance(v, dict) and "value" in v:
                    v = v["value"]
                try:
                    s = str(v).replace(",", ".")
                    return int(float(s))
                except Exception:
                    return None
        return None

    # --------------------------- Fórmulas ---------------------------
    @staticmethod
    def split_commands(raw: str) -> List[str]:
        return [p.strip() for p in (raw or "").split("//") if p.strip()]

    @staticmethod
    def split_core_and_multiplier(raw_expr: str) -> Tuple[str, Optional[object]]:
        s = (raw_expr or "").strip()
        m = re.match(r"^\(\s*(.*)\s*\)\s*(?:\s*\*\s*(#|\d+))?\s*$", s, flags=re.IGNORECASE | re.DOTALL)
        if m:
            core = (m.group(1) or "").strip()
            tok = m.group(2)
        else:
            m2 = re.match(r"^(.*?)(?:\s*\*\s*(#|\d+))?\s*$", s, flags=re.IGNORECASE | re.DOTALL)
            core = (m2.group(1) if m2 else s).strip()
            tok = (m2.group(2) if m2 else None)

        if tok == "#":
            return core, "#"
        if tok and tok.isdigit():
            try:
                return core, int(tok)
            except Exception:
                return core, None
        return core, None

    # ---------- flags ["UL"] / ["IEC"] ----------
    @staticmethod
    def _strip_norma_flags(expr: str) -> Tuple[str, List[str]]:
        flags: List[str] = []

        def repl(m):
            val = (m.group(1) or "").strip().upper()
            if val in ("UL", "IEC"):
                flags.append(val)
            return ""

        clean = re.sub(r'\[\s*"([^"]+)"\s*\]', repl, expr)
        clean = re.sub(r"\s+", " ", clean).strip()
        return clean, flags

    # Wrapper público (para usar desde step4_elementos_fijos)
    def strip_norma_flags(self, expr: str) -> Tuple[str, List[str]]:
        return self._strip_norma_flags(expr)

    # ----------------------- 'MARCA DE ELEMENTOS' -----------------------
    def execute_command(
        self,
        expr: str,
        marca_elem: str,
        norma_ap: str,
        t_ctl: str,
        t_alim: str,
    ) -> List[Dict[str, str]]:
        out: List[Dict[str, str]] = []
        rule = (expr or "").strip()
        if not rule:
            return out

        m = re.match(r"^\s*MARCA\s+DE\s+ELEMENTOS\s*:\s*(.*)$", rule, flags=re.IGNORECASE | re.DOTALL)
        if not m:
            return out

        rest = m.group(1).strip()

        # Si el comando trae flags UL/IEC, respetarlos aquí también
        rest, norma_flags = self._strip_norma_flags(rest)
        if norma_flags:
            want = (norma_ap or "").strip().upper() or "IEC"
            if want not in norma_flags:
                return out

        before_eq, cols_part = self._split_once(rest, "=")
        conds = before_eq.strip()
        ret_cols = self._parse_col_list(cols_part) or ["B", "C", "F", "G", "H", "I", "L"]

        brand = (marca_elem or "").strip()
        if not brand:
            return out

        df_brand = self._get_sheet(brand)
        if df_brand is None:
            return out

        cond_list = self._parse_conditions(conds)
        for i in range(1, df_brand.shape[0]):  # incluir fila 2
            if self._row_matches(df_brand, i, cond_list, norma_ap, t_ctl, t_alim):
                d: Dict[str, str] = {}
                for col in ("B", "C", "F", "G", "H", "I", "L"):
                    j = self.col(col)
                    d[col] = self.cell(df_brand, i, j) if col in ret_cols else ""
                out.append(d)
        return out

    # -------------------- ABB: lookup breaker por modelo --------------------
    def get_abb_breaker_row(self, model: str) -> Dict[str, str]:
        """
        Devuelve columnas B,C,F,G,H,I de la hoja ABB para el breaker cuyo MODELO (col-C)
        coincide con 'model'. Solo acepta filas con col-A == 'B'.
        """
        model = (model or "").strip()
        if not model:
            return {}

        df = self._get_sheet("ABB")
        if df is None:
            return {}

        try:
            colA = df.iloc[:, 0].astype(str).str.strip().str.upper()  # A
            colC = df.iloc[:, 2].astype(str).str.strip()               # C

            mask = colA.eq("B") & colC.str.upper().eq(model.upper())
            idxs = df.index[mask].tolist()

            # fallback: contiene
            if not idxs:
                mask = colA.eq("B") & colC.str.upper().str.contains(model.upper(), na=False)
                idxs = df.index[mask].tolist()
            if not idxs:
                return {}

            i = idxs[0]
            def g(letter: str) -> str:
                return self.cell(df, i, self.col(letter))

            return {k: g(k) for k in ["B", "C", "F", "G", "H", "I"]}
        except Exception:
            return {}

    # --------------------------- Borneras ---------------------------
    def first_data_row_ops(self, df_ops: pd.DataFrame, col_AB: int, col_AC: int) -> int:
        for i in range(1, df_ops.shape[0]):
            if (self.cell(df_ops, i, col_AB) or self.cell(df_ops, i, col_AC)).strip():
                return i
        return 1

    def borneras_por_compresor_arranque(
        self, df_ops: pd.DataFrame, col_BA: int, col_AX: int, col_AY: int, col_AZ: int, arranque: str
    ) -> Tuple[int, int, int]:
        want = (arranque or "").strip().upper()
        if want not in ("V", "P", "D"):
            return (0, 0, 0)
        tot_f = tot_n = tot_t = 0
        for i in range(1, df_ops.shape[0]):
            raw = (self.cell(df_ops, i, col_BA) or "").strip().upper()
            if raw[:1] != want:
                continue
            tot_f += self.to_int(self.cell(df_ops, i, col_AX))
            tot_n += self.to_int(self.cell(df_ops, i, col_AY))
            tot_t += self.to_int(self.cell(df_ops, i, col_AZ))
        return (tot_f, tot_n, tot_t)

    def compute_borneras_totales_WXY(
        self,
        df_ops: pd.DataFrame,
        start_ops: int,
        *,
        col_W: int,
        col_X: int,
        col_Y: int,
        col_BA: int,
        col_AC: int,
        col_AD: int,
        col_AE: int,
        col_AF: int,
        col_AG: int,
        col_AH: int,
        col_AI: int,
        norma_ap: str,
        step3_state: Dict[str, object],
    ) -> Dict[str, int]:
        tot_fase = tot_neutro = tot_tierra = 0

        for i in range(start_ops, df_ops.shape[0]):
            ba_raw = (self.cell(df_ops, i, col_BA) or "").strip().upper()
            if ba_raw[:1] in ("V", "P", "D"):
                continue

            raw_formula = (self.cell(df_ops, i, col_AC) or "").strip()
            if not raw_formula:
                continue

            act1 = self.pick_action(df_ops, i, col_AD, col_AE, col_AF, step3_state)
            skip1, _ = self.resolve_actions([act1])
            if skip1:
                continue

            commands = self.split_commands(raw_formula)
            if not commands:
                continue

            c2_label = (self.cell(df_ops, i, col_AG) or "").strip()
            a_yes = (self.cell(df_ops, i, col_AH) or "").strip()
            a_no  = (self.cell(df_ops, i, col_AI) or "").strip()
            cond2_mult, _brand2 = self.eval_cond2(c2_label, a_yes, a_no, step3_state, norma_ap)
            cond2_uses_number = (a_yes == "#" or a_no == "#")

            total_multipliers = 0
            for cmd in commands:
                _, mult = self.split_core_and_multiplier(cmd)
                base_mult = 1
                requires_c2 = False
                if isinstance(mult, int):
                    base_mult = max(1, mult)
                elif mult == "#":
                    requires_c2 = True

                if requires_c2:
                    m = max(0, int(cond2_mult))
                else:
                    m = base_mult
                    if cond2_uses_number:
                        m *= max(1, int(cond2_mult))

                if m <= 0:
                    continue
                total_multipliers += int(m)

            tot_fase   += total_multipliers * self.to_int(self.cell(df_ops, i, col_W))
            tot_neutro += total_multipliers * self.to_int(self.cell(df_ops, i, col_X))
            tot_tierra += total_multipliers * self.to_int(self.cell(df_ops, i, col_Y))

        return {"fase": tot_fase, "neutro": tot_neutro, "tierra": tot_tierra}

    # --------------------------- Parser internals ---------------------------
    def _get_sheet(self, name: str) -> Optional[pd.DataFrame]:
        key = (name or "").strip()
        if not key:
            return None
        if key in self._sheet_cache:
            return self._sheet_cache[key]
        try:
            df = pd.read_excel(self.book, sheet_name=key, header=None, dtype=str)
            self._sheet_cache[key] = df
            return df
        except Exception:
            self._sheet_cache[key] = None
            return None

    @staticmethod
    def _split_once(s: str, sep: str) -> Tuple[str, str]:
        parts = (s or "").split(sep, 1)
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
        t_alim: str,
    ) -> bool:
        """
        Coincidencia de fila con soporte especial para norma UL/IEC:
        - Para condiciones UL/IEC NO se acepta '*' ni vacío como comodín.
        - Columna con 'SI/NO' interpreta SI⇒UL, NO⇒IEC.
        - Columna con 'UL', 'IEC', 'UL/IEC' selecciona por etiqueta.
        """
        norma_want = (norma_ap or "").strip().upper() or "IEC"

        for col_letter, selector in conds:
            j = self.col(col_letter)
            cell_val = (self.cell(df, i, j) or "").strip()
            sel = (selector or "").strip()
            up_sel = sel.upper()
            cell_u = (cell_val or "").strip().upper()

            # ---------- Caso especial: chequeo de norma UL/IEC ----------
            if up_sel in ("UL", "IEC"):
                if cell_u in ("", "*"):
                    return False

                if cell_u in ("SI", "NO"):
                    want = "SI" if norma_want == "UL" else "NO"
                    if cell_u == want:
                        continue
                    return False

                has_ul  = "UL"  in cell_u
                has_iec = "IEC" in cell_u
                if (norma_want == "UL" and has_ul) or (norma_want != "UL" and has_iec):
                    continue
                return False

            if cell_val == "*":
                continue

            if up_sel == "TENSION CONTROL:":
                if self.only_digits(cell_val) == self.only_digits(t_ctl):
                    continue
                return False

            if up_sel == "TENSION ALIMENTACION:":
                if self.only_digits(cell_val) == self.only_digits(t_alim):
                    continue
                return False

            sd = self.only_digits(sel)
            cd = self.only_digits(cell_val)
            if sd and cd:
                if sd == cd:
                    continue
                return False

            if cell_u == up_sel:
                continue

            return False

        return True

