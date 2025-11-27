# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from typing import Dict, Any

from .step4_compresores import build_compresores_tables
from .step4_elementos_fijos import cargar_elementos_fijos
from .step4_borneras import borneras_compresores_totales

# ---------------- Normalización Paso 3 / util ----------------
import unicodedata, re

def _n_txt(s: Any) -> str:
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).upper().strip()

def _to_flags(step3_raw: Dict[str, Any] | None) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for k, v in (step3_raw or {}).items():
        kk = _n_txt(k)
        if isinstance(v, bool):
            out[kk] = "SI" if v else "NO"; continue
        if isinstance(v, (int, float)):
            out[kk] = str(int(v)) if isinstance(v, float) and v.is_integer() else str(v); continue
        vv = _n_txt(v)
        if vv in ("SI","S","YES","TRUE","1"): out[kk] = "SI"
        elif vv in ("NO","N","FALSE","0"):    out[kk] = "NO"
        else:                                  out[kk] = vv
    if "UL" in out and "UL/IEC" not in out:
        out["UL/IEC"] = "SI" if out["UL"] == "SI" else "NO"
    return out

def _only_digits(s: Any) -> str:
    return "".join(ch for ch in str(s or "") if ch.isdigit())

def _get_first(globs: Dict[str, Any], *keys: str, default: str = "") -> str:
    for k in keys:
        if k in globs and globs[k] not in (None, ""):
            return str(globs[k])
    return default


class Step4Engine:
    """
    Orquestador del Paso 4. Devuelve:
      - tables_compresores: List[ResumenTable]
      - otros_rows: List[List[str]]
      - borneras_compresores / borneras_otros / borneras_total
    """

    def __init__(self, basedatos_path: Path | str) -> None:
        self.book = Path(basedatos_path)
        self.ctx = None  # objeto con atributos usados por los engines

    def set_contexto(self, ctx) -> None:
        self.ctx = ctx

    def _ensure_ctx_from_globs(self, globs: Dict[str, Any] | None) -> None:
        class _Dummy: pass
        if self.ctx is None:
            self.ctx = _Dummy()
        if globs is None:
            globs = {}

        # REFRESCAR SIEMPRE
        self.ctx.marca_elementos = (
            _get_first(globs, "marca_elementos", "marca_elem", "marca", default="")
        ).strip()

        self.ctx.norma_ap = str(globs.get("norma_ap") or "").strip().upper()

        t_ctl  = _get_first(globs, "t_ctl",  "tension_control",       default="")
        t_alim = _get_first(globs, "t_alim", "tension_alimentacion",  default="")
        self.ctx.tension_control       = _only_digits(t_ctl)
        self.ctx.tension_alimentacion  = _only_digits(t_alim)

        raw_step3 = globs.get("step3_state") or globs.get("step3") or {}
        self.ctx.step3_state_raw = raw_step3
        self.ctx.step3_state     = _to_flags(raw_step3)

        # Fallback UL/IEC si no vino norma_ap explícita
        if not self.ctx.norma_ap:
            ul_flag = self.ctx.step3_state.get("UL", self.ctx.step3_state.get("UL/IEC", "NO"))
            self.ctx.norma_ap = "UL" if ul_flag == "SI" else "IEC"

    # ===== NUEVO: inyección de fila(s) para PUENTE =====
    def _inject_puente_rows(self, tables, step2_state: Dict[str, Dict[str, str]]) -> None:
        """
        Si el Paso 2 dejó 'puente_modelo'/'puente_codigo', agregamos una o dos filas
        a la tabla del compresor. Reglas:
          - CÓDIGO: SIEMPRE el valor de R (puente_codigo).
          - MODELO: el valor de Q (puente_modelo).
          - NOMBRE: 'PUENTE GM-CONTACTOR'.
          - DESCRIPCIÓN: 'PUENTE GUARDAMOTOR CONTACTOR'.
          - REFERENCIA: también R (puente_codigo).
          - Cantidad: si el arranque del compresor es PARTIDO -> 2 filas; si no -> 1 fila.
        Formato de fila: [ITEM, CÓDIGO, MODELO, NOMBRE, DESCRIPCIÓN, C240, C480, REFERENCIA, TORQUE]
        """
        for t in tables or []:
            st = step2_state.get(getattr(t, "comp_key", ""), {}) or {}
            p_model = (st.get("puente_modelo") or "").strip()  # Q (ej. BEA38-4)
            p_code  = (st.get("puente_codigo") or "").strip()  # R (ej. 130511)

            # si no hay ni Q ni R, no agregamos nada
            if not p_model and not p_code:
                continue

            # cantidad por tipo de arranque
            arr = str(getattr(t, "arranque", "") or "").strip().upper()
            qty = 2 if arr == "PARTIDO" else 1

            row = [
                "PUENTE GM-CONTACTOR",          # ITEM
                p_code,                         # CÓDIGO -> ***R***
                p_model,                        # MODELO -> Q
                "PUENTE GM-CONTACTOR",          # NOMBRE
                "PUENTE GUARDAMOTOR CONTACTOR", # DESCRIPCIÓN
                "", "",                         # C 240 / C 480
                p_code,                         # REFERENCIA -> también R
                ""                              # TORQUE
            ]

            # repetir según cantidad (PARTIDO = 2)
            for _ in range(max(1, qty)):
                getattr(t, "rows", []).append(list(row))  # copia defensiva

    def calcular(self, step2_state: Dict[str, Dict[str, str]], globs: Dict[str, Any] | None = None) -> Dict[str, Any]:
        self._ensure_ctx_from_globs(globs)

        tables = build_compresores_tables(self.book, step2_state, self.ctx)

        # NUEVO: inyectar fila(s) del puente si vienen del Paso 2
        try:
            self._inject_puente_rows(tables, step2_state)
        except Exception:
            # si falla, no rompemos el resto del cálculo
            pass

        otros_rows, otros_wxy, _dbg = cargar_elementos_fijos(self.book, self.ctx)

        fase_c, neutro_c, tierra_c = borneras_compresores_totales(
            self.book, tables=tables, step2_state=step2_state
        )
        borneras_comp = {"fase": fase_c, "neutro": neutro_c, "tierra": tierra_c}

        borneras_otros = {
            "fase": int(otros_wxy.get("fase", 0)),
            "neutro": int(otros_wxy.get("neutro", 0)),
            "tierra": int(otros_wxy.get("tierra", 0)),
        }
        borneras_total = {
            "fase":   borneras_comp["fase"]   + borneras_otros["fase"],
            "neutro": borneras_comp["neutro"] + borneras_otros["neutro"],
            "tierra": borneras_comp["tierra"] + borneras_otros["tierra"],
        }

        return {
            "tables_compresores": tables,
            "otros_rows": otros_rows,
            "borneras_compresores": borneras_comp,
            "borneras_otros": borneras_otros,
            "borneras_total": borneras_total,
        }

