# gui/pages/materiales_step2.py
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QFrame,
    QLabel, QLineEdit, QComboBox, QScrollArea, QCompleter,
    QRadioButton, QButtonGroup
)

# ---- imports de lógica con fallback a stubs para que nunca rompa la GUI ----
try:
    from logic.guardamotor import seleccionar_guardamotor_abb as _sel_gm
except Exception:  # pragma: no cover
    def _sel_gm(book_path: str, corriente: float, arranque: str) -> Dict:
        return {"aplica": False, "motivo": "GM stub", "modelo": "", "cantidad": 0, "valor_col_E": 0.0}

try:
    from logic.contactor import seleccionar_contactor_abb as _sel_cont
except Exception:  # pragma: no cover
    def _sel_cont(book_path: str, corriente: float, arranque: str) -> Dict:
        return {"aplica": False, "motivo": "CONTACTOR stub", "modelo": "", "cantidad": 0, "valor_col_D": 0.0}

# >>> NUEVO: breaker para variador (ABB) <<<
try:
    from logic.breaker_vfd import seleccionar_breaker_vfd_abb as _sel_brk_vfd
except Exception:  # pragma: no cover
    def _sel_brk_vfd(book_path: str, corriente: float) -> Dict:
        return {"aplica": False, "motivo": "BREAKER stub", "modelo": "", "valor_col_D": 0.0}


class Step2Panel(QWidget):
    """
    Paso 2 — Configuración de Compresores.
    Guarda/restaura la selección del variador (bolita) y no la pierde al navegar.
    """

    def __init__(self) -> None:
        super().__init__()
        self._sheet_cache_comp: Dict[str, Optional[pd.DataFrame]] = {}
        self._sheet_cache_var: Dict[str, Optional[pd.DataFrame]] = {}
        self._step2_state: Dict[str, Dict[str, str]] = {}     # último estado exportado
        self._sel_state: Dict[str, str] = {}                   # **recordatorio de V1/V2 por grupo**
        self._step2_widgets: Dict[str, Dict[str, object]] = {}

        # fuentes globales (combos del Paso 1)
        self._cb_tipo_comp: Optional[QComboBox] = None
        self._cb_t_alim: Optional[QComboBox] = None
        self._cb_refrig: Optional[QComboBox] = None
        self._cb_marca_var: Optional[QComboBox] = None

        # UI base
        root = QVBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(8)
        self.scroll = QScrollArea(); self.scroll.setWidgetResizable(True)
        self.host = QFrame()
        self.host_layout = QVBoxLayout(self.host); self.host_layout.setContentsMargins(0,0,0,0); self.host_layout.setSpacing(10)
        self.scroll.setWidget(self.host)
        root.addWidget(self.scroll, 1)

    # ---------- API pública ----------
    def set_globals(self, tipo_comp: QComboBox, t_alim: QComboBox, refrig: QComboBox, marca_var: QComboBox) -> None:
        self._cb_tipo_comp = tipo_comp
        self._cb_t_alim = t_alim
        self._cb_refrig = refrig
        self._cb_marca_var = marca_var

    def rebuild(self, nb: int, nm: int, np: int, modelos: List[str]) -> None:
        names_g = [f"G{i}" for i in range(1, nb + 1)]
        names_b = [f"B{i}" for i in range(1, nm + 1)]
        names_f = [f"F{i}" for i in range(1, np + 1)]
        names_all = names_g + names_b + names_f

        # conserva estado/selección de nombres existentes
        self._step2_state = {k: v for k, v in self._step2_state.items() if k in names_all}
        self._sel_state   = {k: v for k, v in self._sel_state.items()   if k in names_all}

        # limpia UI
        while self.host_layout.count():
            w = self.host_layout.takeAt(0).widget()
            if w: w.deleteLater()
        self._step2_widgets.clear()

        # secciones
        self._build_comp_section("COMPRESORES BAJA TEMPERATURA", names_g, modelos)
        self._build_comp_section("COMPRESORES MEDIA TEMPERATURA", names_b, modelos)
        self._build_comp_section("COMPRESORES PARALELO", names_f, modelos)

        # refresca corrientes y cálculos
        self.refresh_all()

    def refresh_all(self) -> None:
        for name in list(self._step2_widgets.keys()):
            self._update_corriente_for(name)

    def export_state(self) -> Dict[str, Dict[str, str]]:
        # captura del UI actual a dict
        for name, w in self._step2_widgets.items():
            # prioriza el recordatorio interno; si no, lee la UI
            remembered = self._sel_state.get(name, "")
            ui_sel = "V1" if w["rb_v1"].isChecked() else "V2" if w["rb_v2"].isChecked() else ""
            variador_sel = remembered or ui_sel

            self._step2_state[name] = {
                "modelo": w["modelo"].currentText().strip(),
                "arranque": w["arranque"].currentText().strip(),
                "corriente": w["corriente"].text().strip(),
                "guardamotor": w["gm_lbl"].text().strip(),
                "variador1": w["v1_lbl"].text().strip(),
                "variador2": w["v2_lbl"].text().strip(),
                "variador_sel": variador_sel,               # <-- guarda V1/V2
                "contactor": w["cont_lbl"].text().strip(),
                "breaker": w["brk_lbl"].text().strip(),     # breaker VFD
            }
        return dict(self._step2_state)

    def import_state(self, state: Dict[str, Dict[str, str]]) -> None:
        """Restaura modelo/arranque/corriente/variadores y re-marca la bolita."""
        if not isinstance(state, dict):
            return
        for name, rec in state.items():
            w = self._step2_widgets.get(name)
            if not w:
                continue

            # Modelo
            mdl = (rec.get("modelo") or "").strip()
            if mdl:
                self._set_combo_text(w["modelo"], mdl)

            # Corriente
            curr = (rec.get("corriente") or "").strip()
            if curr:
                w["corriente"].setText(curr)

            # Arranque
            arr = (rec.get("arranque") or "").strip()
            if arr:
                self._set_combo_text(w["arranque"], arr)

            # Variadores (textos de referencia)
            v1 = (rec.get("variador1") or "").strip()
            v2 = (rec.get("variador2") or "").strip()
            if v1:
                w["v1_lbl"].setText(v1)
            if v2:
                w["v2_lbl"].setText(v2)

            # Breaker/GM/Contactor (si venían)
            if rec.get("guardamotor"):
                w["gm_lbl"].setText(rec["guardamotor"].strip() or "—")
            if rec.get("contactor"):
                w["cont_lbl"].setText(rec["contactor"].strip() or "—")
            if rec.get("breaker"):
                w["brk_lbl"].setText(rec["breaker"].strip() or "—")

            # Selección del variador
            vsel = self._normalize_vsel(rec.get("variador_sel", ""))
            if vsel:
                self._sel_state[name] = vsel
                # se marcará al habilitar radios:
                self._apply_remembered_selection(name)

        # No llamo a refresh_all() para no sobrescribir labels; solo
        # aseguro habilitación/selección acorde a arranque actual:
        for name in state.keys():
            self._toggle_var_radios(name)

    def clear(self) -> None:
        self._step2_state.clear()
        self._sel_state.clear()
        self._step2_widgets.clear()
        while self.host_layout.count():
            w = self.host_layout.takeAt(0).widget()
            if w: w.deleteLater()

    # ---------- construcción de secciones ----------
    def _build_comp_section(self, titulo: str, names: List[str], modelos: List[str]) -> None:
        if not names: return
        box = QFrame(); box.setStyleSheet("QFrame{background:#2f3242;border:1px solid #42465a;border-radius:8px;}")
        bl = QVBoxLayout(box); bl.setContentsMargins(12, 12, 12, 12); bl.setSpacing(8)
        lab = QLabel(titulo); lab.setStyleSheet("color:#e0e0e0;font-weight:700;"); bl.addWidget(lab)

        grid = QGridLayout(); grid.setHorizontalSpacing(16); grid.setVerticalSpacing(8); bl.addLayout(grid)
        headers = [
            "", "MODELO", "CORRIENTE", "ARRANQUE",
            "VARIADOR 1", "VARIADOR 2", "GUARDAMOTOR",
            "MODELO CONTACTOR", "BREAKER (VFD)"
        ]
        for c, h in enumerate(headers):
            hl = QLabel(h); hl.setStyleSheet("color:#aeb7d6;font-weight:700;"); grid.addWidget(hl, 0, c)

        for i, name in enumerate(names, start=1):
            grid.addWidget(QLabel(name + ":"), i, 0)

            # modelo
            cb_modelo = QComboBox(); cb_modelo.addItem(""); cb_modelo.addItems(modelos)
            cb_modelo.setEditable(True); cb_modelo.setInsertPolicy(QComboBox.NoInsert)
            compl = QCompleter(modelos); compl.setFilterMode(Qt.MatchContains); compl.setCaseSensitivity(Qt.CaseInsensitive)
            cb_modelo.setCompleter(compl); grid.addWidget(cb_modelo, i, 1)

            # corriente
            le_corr = QLineEdit(); le_corr.setReadOnly(True); le_corr.setPlaceholderText("A")
            le_corr.setFixedWidth(120)
            le_corr.setStyleSheet("QLineEdit{background:#202334;color:#e0e0e0;border:1px solid #3a3d4f;border-radius:6px;padding:4px 6px;}")
            grid.addWidget(le_corr, i, 2)

            # arranque
            cb_arranque = QComboBox(); cb_arranque.addItems(["", "VARIADOR", "DIRECTO", "PARTIDO"])
            grid.addWidget(cb_arranque, i, 3)

            # variador 1
            v1_layout = QHBoxLayout(); v1_layout.setSpacing(6)
            rb_v1 = QRadioButton(); rb_v1.setEnabled(False)
            v1_lbl = QLabel("—"); v1_lbl.setStyleSheet("color:#e0e0e0;")
            v1_w = QFrame(); v1_w.setLayout(v1_layout); v1_layout.addWidget(rb_v1); v1_layout.addWidget(v1_lbl, 1)
            grid.addWidget(v1_w, i, 4)

            # variador 2
            v2_layout = QHBoxLayout(); v2_layout.setSpacing(6)
            rb_v2 = QRadioButton(); rb_v2.setEnabled(False)
            v2_lbl = QLabel("—"); v2_lbl.setStyleSheet("color:#e0e0e0;")
            v2_w = QFrame(); v2_w.setLayout(v2_layout); v2_layout.addWidget(rb_v2); v2_layout.addWidget(v2_lbl, 1)
            grid.addWidget(v2_w, i, 5)
            grp = QButtonGroup(self); grp.addButton(rb_v1); grp.addButton(rb_v2)

            gm_lbl = QLabel("—"); gm_lbl.setStyleSheet("color:#e0e0e0;")
            cont_lbl = QLabel("—"); cont_lbl.setStyleSheet("color:#e0e0e0;")
            brk_lbl = QLabel("—"); brk_lbl.setStyleSheet("color:#e0e0e0;")

            grid.addWidget(gm_lbl, i, 6)
            grid.addWidget(cont_lbl, i, 7)
            grid.addWidget(brk_lbl, i, 8)

            # restaura estado si existe
            if name in self._step2_state:
                prev = self._step2_state[name]
                ixm = cb_modelo.findText(prev.get("modelo",""), Qt.MatchFixedString)
                if ixm >= 0: cb_modelo.setCurrentIndex(ixm); cb_modelo.lineEdit().setText(cb_modelo.itemText(ixm))
                ixa = cb_arranque.findText(prev.get("arranque",""), Qt.MatchFixedString)
                if ixa >= 0: cb_arranque.setCurrentIndex(ixa)
                if prev.get("corriente"): le_corr.setText(prev["corriente"])
                if prev.get("guardamotor"): gm_lbl.setText(prev["guardamotor"])
                if prev.get("variador1"): v1_lbl.setText(prev["variador1"])
                if prev.get("variador2"): v2_lbl.setText(prev["variador2"])
                if prev.get("contactor"): cont_lbl.setText(prev["contactor"])
                if prev.get("breaker"): brk_lbl.setText(prev["breaker"])

                # Normaliza y recuerda selección
                vsel = self._normalize_vsel(prev.get("variador_sel", ""))
                if vsel:
                    self._sel_state[name] = vsel
                    if vsel == "V1": rb_v1.setChecked(True)
                    elif vsel == "V2": rb_v2.setChecked(True)

            # eventos
            def _ensure_valid_and_current(cmb=cb_modelo, n=name, le=le_corr):
                txt = cmb.lineEdit().text().strip()
                if not txt:
                    cmb.setCurrentIndex(0); cmb.lineEdit().clear(); le.setText("")
                    self._set_gm(n, ""); self._set_cont(n, ""); self._set_vars(n, "", "")
                    self._set_breaker(n, ""); self._toggle_var_radios(n); return
                ix = cmb.findText(txt, Qt.MatchFixedString)
                if ix < 0:
                    for j in range(1, cmb.count()):
                        if txt.lower() in cmb.itemText(j).lower(): ix = j; break
                if ix >= 0: cmb.setCurrentIndex(ix); cmb.lineEdit().setText(cmb.itemText(ix))
                else: cmb.setCurrentIndex(0); cmb.lineEdit().clear()
                self._update_corriente_for(n)

            cb_modelo.lineEdit().editingFinished.connect(_ensure_valid_and_current)
            cb_modelo.currentTextChanged.connect(lambda _=None, n=name: self._update_corriente_for(n))
            cb_arranque.currentIndexChanged.connect(lambda _=None, n=name: self._on_arranque_changed(n))
            rb_v1.toggled.connect(lambda _=None, n=name: self._on_var_selection_changed(n))
            rb_v2.toggled.connect(lambda _=None, n=name: self._on_var_selection_changed(n))

            self._step2_widgets[name] = {
                "modelo": cb_modelo, "corriente": le_corr, "arranque": cb_arranque,
                "rb_v1": rb_v1, "rb_v2": rb_v2, "v1_lbl": v1_lbl, "v2_lbl": v2_lbl, "grp": grp,
                "gm_lbl": gm_lbl, "cont_lbl": cont_lbl, "brk_lbl": brk_lbl
            }

        self.host_layout.addWidget(box)
        grid.setColumnStretch(1, 2)
        for c in (2,3,4,5,6,7,8):
            grid.setColumnStretch(c, 1)

    # ---------- cálculos ----------
    def _on_arranque_changed(self, name: str) -> None:
        self._recalc_guardamotor(name)
        self._recalc_contactor(name)
        self._recalc_variadores(name)
        self._recalc_breaker_vfd(name)

    def _on_var_selection_changed(self, name: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return
        sel = "V1" if w["rb_v1"].isChecked() else "V2" if w["rb_v2"].isChecked() else ""
        if sel:
            self._sel_state[name] = sel

    def _update_corriente_for(self, name: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return

        modelo = w["modelo"].currentText().strip()
        le = w["corriente"]
        if not modelo:
            le.setText("")
            self._set_gm(name,""); self._set_cont(name,""); self._set_vars(name,"","")
            self._set_breaker(name,""); self._toggle_var_radios(name); return

        marca = (self._cb_tipo_comp.currentText() or "").upper().strip() if self._cb_tipo_comp else ""
        tension = (self._cb_t_alim.currentText() or "").strip() if self._cb_t_alim else ""
        refrig = (self._cb_refrig.currentText() or "").upper().strip() if self._cb_refrig else ""

        if refrig != "R744":
            le.setText("")
            self._set_gm(name,""); self._set_cont(name,""); self._set_vars(name,"","")
            self._set_breaker(name,""); self._toggle_var_radios(name); return

        amps_raw = self._buscar_corriente_en_hoja(marca, modelo, tension)
        if not amps_raw:
            le.setText("")
            self._set_gm(name,""); self._set_cont(name,""); self._set_vars(name,"","")
            self._set_breaker(name,""); self._toggle_var_radios(name); return

        try:
            val = float(str(amps_raw).replace(",", ".")); le.setText(f"{val:.1f} A")
        except ValueError:
            le.setText(""); self._set_gm(name,""); self._set_cont(name,""); self._set_vars(name,"","")
            self._set_breaker(name,""); self._toggle_var_radios(name); return

        # después de corriente: recalcular
        self._recalc_guardamotor(name)
        self._recalc_contactor(name)
        self._recalc_variadores(name)
        self._recalc_breaker_vfd(name)

    def _recalc_guardamotor(self, name: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return
        arr = (w["arranque"].currentText() or "").upper().strip()
        if arr not in ("DIRECTO","PARTIDO"):
            self._set_gm(name, "NO APLICA" if arr=="VARIADOR" else ""); return
        amps = self._get_amps(name)
        if amps is None: self._set_gm(name,""); return
        res = _sel_gm(self._book().as_posix(), amps, arr)
        if res.get("aplica"):
            self._set_gm(name, f"{res['modelo']}".strip())
        else:
            self._set_gm(name, res.get("motivo") or res.get("error","NO ENCONTRADO"))

    def _recalc_contactor(self, name: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return
        arr = (w["arranque"].currentText() or "").upper().strip()
        if arr not in ("DIRECTO","PARTIDO"):
            self._set_cont(name, "NO APLICA" if arr=="VARIADOR" else ""); return
        amps = self._get_amps(name)
        if amps is None: self._set_cont(name,""); return
        res = _sel_cont(self._book().as_posix(), amps, arr)
        if res.get("aplica"):
            self._set_cont(name, f"{res['modelo']}".strip())
        else:
            self._set_cont(name, res.get("motivo") or res.get("error","NO ENCONTRADO"))

    def _recalc_variadores(self, name: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return
        arr = (w["arranque"].currentText() or "").upper().strip()
        if arr != "VARIADOR":
            self._set_vars(name,"",""); self._toggle_var_radios(name); return

        marca_var = (self._cb_marca_var.currentText() or "").upper().strip() if self._cb_marca_var else ""
        if marca_var in ("","NO"):
            self._set_vars(name,"Seleccione MARCA VAR",""); self._toggle_var_radios(name); return

        amps = self._get_amps(name)
        if amps is None: self._set_vars(name,"",""); self._toggle_var_radios(name); return

        df_var = self._load_var_sheet(marca_var)
        if df_var is None: self._set_vars(name,"Sin hoja VAR",""); self._toggle_var_radios(name); return

        tension = (self._cb_t_alim.currentText() or "").strip() if self._cb_t_alim else ""
        t_digits = "".join(ch for ch in tension if ch.isdigit())
        r_col = self._col_letter_to_index("I" if t_digits=="220" else "AI")
        ret_col = self._col_letter_to_index("A" if t_digits=="220" else "AA")
        v1, v2 = self._find_two_var_drives(df_var, r_col, ret_col, amps)
        self._set_vars(name, v1, v2)
        self._toggle_var_radios(name)

    # --- breaker del variador ---
    def _recalc_breaker_vfd(self, name: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return
        arr = (w["arranque"].currentText() or "").upper().strip()
        if arr != "VARIADOR":
            self._set_breaker(name, "NO APLICA" if arr in ("DIRECTO","PARTIDO") else "")
            return
        amps = self._get_amps(name)
        if amps is None:
            self._set_breaker(name, ""); return
        res = _sel_brk_vfd(self._book().as_posix(), amps)
        if res.get("aplica"):
            self._set_breaker(name, f"{res.get('C', res.get('modelo',''))}".strip() or "—")
        else:
            self._set_breaker(name, res.get("motivo") or res.get("error","NO ENCONTRADO"))

    # ---------- helpers de estado / UI ----------
    def _get_amps(self, name: str) -> Optional[float]:
        w = self._step2_widgets.get(name)
        if not w: return None
        s = (w["corriente"].text() or "").replace("A","").replace(",",".").strip()
        try:
            return float(s) if s else None
        except ValueError:
            return None

    def _set_gm(self, name: str, text: str) -> None:
        lab = self._step2_widgets.get(name,{}).get("gm_lbl")
        if isinstance(lab, QLabel): lab.setText(text or "—")

    def _set_cont(self, name: str, text: str) -> None:
        lab = self._step2_widgets.get(name,{}).get("cont_lbl")
        if isinstance(lab, QLabel): lab.setText(text or "—")

    def _set_breaker(self, name: str, text: str) -> None:
        lab = self._step2_widgets.get(name,{}).get("brk_lbl")
        if isinstance(lab, QLabel): lab.setText(text or "—")

    def _set_vars(self, name: str, v1: str, v2: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return
        if isinstance(w.get("v1_lbl"), QLabel): w["v1_lbl"].setText(v1 or "—")
        if isinstance(w.get("v2_lbl"), QLabel): w["v2_lbl"].setText(v2 or "—")

    def _apply_remembered_selection(self, name: str) -> None:
        """Marca la radio guardada en _sel_state si está habilitada."""
        w = self._step2_widgets.get(name)
        if not w: return
        remembered = self._sel_state.get(name, "")
        if remembered == "V1" and w["rb_v1"].isEnabled():
            w["rb_v1"].setChecked(True)
        elif remembered == "V2" and w["rb_v2"].isEnabled():
            w["rb_v2"].setChecked(True)

    def _toggle_var_radios(self, name: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return
        arr = (w["arranque"].currentText() or "").upper().strip()
        v1_ok = (w["v1_lbl"].text() or "—").strip() != "—"
        v2_ok = (w["v2_lbl"].text() or "—").strip() != "—"
        enable = (arr == "VARIADOR") and (v1_ok or v2_ok)
        w["rb_v1"].setEnabled(enable and v1_ok)
        w["rb_v2"].setEnabled(enable and v2_ok)

        if not enable:
            # No borrar _sel_state; solo desmarcar UI
            w["rb_v1"].setAutoExclusive(False); w["rb_v1"].setChecked(False); w["rb_v1"].setAutoExclusive(True)
            w["rb_v2"].setAutoExclusive(False); w["rb_v2"].setChecked(False); w["rb_v2"].setAutoExclusive(True)
            return

        # Si hay recordatorio, respétalo
        remembered = self._sel_state.get(name, "")
        if remembered == "V1" and v1_ok:
            w["rb_v1"].setChecked(True); return
        if remembered == "V2" and v2_ok:
            w["rb_v2"].setChecked(True); return

        # Si no hay recordatorio: auto-seleccionar si sólo hay uno disponible
        if v1_ok and not v2_ok:
            w["rb_v1"].setChecked(True)
        elif v2_ok and not v1_ok:
            w["rb_v2"].setChecked(True)
        # si hay dos, no forzamos nada (el usuario elige)

    # ---------- Excel helpers ----------
    def load_comp_sheet(self, marca: str) -> Optional[pd.DataFrame]:
        return self._load_comp_sheet(marca)

    def _load_comp_sheet(self, marca: str) -> Optional[pd.DataFrame]:
        key = "".join(ch for ch in (marca or "").upper() if ch.isalnum())
        if not key: return None
        if key in self._sheet_cache_comp: return self._sheet_cache_comp[key]
        try:
            df = pd.read_excel(self._book(), sheet_name=f"COMP{key}", header=None, dtype=str)
            self._sheet_cache_comp[key] = df; return df
        except Exception:
            self._sheet_cache_comp[key] = None; return None

    def _buscar_corriente_en_hoja(self, marca: str, modelo: str, tension: str) -> str:
        mu = (marca or "").upper().strip()
        if mu not in ("BITZER","DORIN"): return ""
        df = self._load_comp_sheet(mu)
        if df is None or df.shape[0] < 3: return ""
        t = "".join(ch for ch in tension if ch.isdigit())
        col_idx = None
        if mu=="BITZER":
            if t=="220": col_idx=4   # E
            elif t=="460": col_idx=7 # H
        elif mu=="DORIN":
            if t=="220": col_idx=2   # C
            elif t=="460": col_idx=3 # D
        if col_idx is None or col_idx >= df.shape[1]: return ""
        colA = df.iloc[2:,0].astype(str).str.strip()
        idxs = colA[colA.str.lower()==modelo.lower()].index.tolist()
        if not idxs: idxs = colA[colA.str.lower().str.contains(modelo.lower())].index.tolist()
        if not idxs: return ""
        val = df.iat[idxs[0], col_idx]
        return "" if pd.isna(val) else str(val).strip()

    def _load_var_sheet(self, marca_var: str) -> Optional[pd.DataFrame]:
        def norm(s: str) -> str:
            s = (s or "").upper()
            trans = str.maketrans("ÁÉÍÓÚÄËÏÖÜáéíóúäëïöü","AEIOUAEIOUAEIOUAEIOU")
            s = s.translate(trans)
            return "".join(ch for ch in s if ch.isalnum())
        key = norm(marca_var)
        if not key: return None
        if key in self._sheet_cache_var: return self._sheet_cache_var[key]
        target = norm(f"VAR{key}")
        try:
            xls = pd.ExcelFile(self._book())
            sh = None
            for s in xls.sheet_names:
                if norm(s) == target: sh = s; break
            if sh is None:
                self._sheet_cache_var[key] = None; return None
            df = pd.read_excel(xls, sheet_name=sh, header=None, dtype=str)
            self._sheet_cache_var[key] = df; return df
        except Exception:
            self._sheet_cache_var[key] = None; return None

    @staticmethod
    def _col_letter_to_index(letter: str) -> int:
        lt = (letter or "").upper().strip(); n = 0
        for ch in lt:
            if not ('A' <= ch <= 'Z'): return 0
            n = n * 26 + (ord(ch) - ord('A') + 1)
        return max(0, n - 1)

    @staticmethod
    def _find_two_var_drives(df: pd.DataFrame, rating_col: int, ret_col: int, amps: float) -> Tuple[str, str]:
        col = df.iloc[2:, rating_col] if rating_col < df.shape[1] else pd.Series(dtype=object)
        ret = df.iloc[2:, ret_col] if ret_col < df.shape[1] else pd.Series(dtype=object)

        def to_float(x) -> Optional[float]:
            if pd.isna(x): return None
            s = str(x).strip().replace(",", ".")
            try: return float(s)
            except ValueError: return None

        models: List[str] = []
        in_block = False; found = False
        for rv, mv in zip(col, ret):
            v = to_float(rv)
            if v is None:
                if in_block: in_block = False; found = False
                continue
            if not in_block: in_block = True; found = False
            if (not found) and (v >= amps):
                m = "" if pd.isna(mv) else str(mv).strip()
                if m: models.append(m); found = True
                if len(models) >= 2: break
        if len(models) == 0: return ("","")
        if len(models) == 1: return (models[0], "")
        return (models[0], models[1])

    # ---------- util ----------
    def _book(self) -> Path:
        return Path(__file__).resolve().parents[2] / "data" / "basedatos.xlsx"

    # ---------- helpers públicos para orquestador ----------
    def get_group_keys(self) -> List[str]:
        keys = list(self._step2_widgets.keys())
        import re
        def sort_key(k: str):
            g = {"G": 0, "B": 1, "F": 2}.get(k[:1].upper(), 9)
            n = int(re.sub(r"\D+", "", k) or "999")
            return (g, n)
        return sorted(keys, key=sort_key)

    def set_fields(self, key: str, *, modelo: str | None = None,
                   corriente: str | None = None, arranque: str | None = None,
                   variador1: str | None = None, variador2: str | None = None,
                   variador_sel: str | None = None, **_extras) -> None:
        row = self._step2_widgets.get(key, {})
        if not row:
            return

        modelo_cbo   = row.get("modelo")
        arranque_cbo = row.get("arranque")
        corriente_le = row.get("corriente")

        if isinstance(modelo_cbo, QComboBox) and modelo:
            self._set_combo_text(modelo_cbo, modelo)

        if isinstance(arranque_cbo, QComboBox) and arranque:
            self._set_combo_text(arranque_cbo, arranque)

        if isinstance(corriente_le, QLineEdit) and corriente:
            corriente_le.setText(corriente)

        if isinstance(row.get("v1_lbl"), QLabel) and variador1:
            row["v1_lbl"].setText(variador1 or "—")
        if isinstance(row.get("v2_lbl"), QLabel) and variador2:
            row["v2_lbl"].setText(variador2 or "—")

        if variador_sel:
            self.set_variador_selection(key, variador_sel)

        # actualizar habilitación acorde a arranque/variadores
        self._toggle_var_radios(key)

    def set_variador_selection(self, key: str, sel: str) -> None:
        row = self._step2_widgets.get(key, {})
        if not row:
            return
        vsel = self._normalize_vsel(sel)
        if not vsel:
            return
        self._sel_state[key] = vsel
        if vsel == "V1":
            row["rb_v1"].setChecked(True)
        elif vsel == "V2":
            row["rb_v2"].setChecked(True)

    # ---------- utilidades pequeñas ----------
    @staticmethod
    def _normalize_vsel(val: str) -> str:
        s = (val or "").strip().upper()
        if s in {"1","V1","VARIADOR 1","UNO","PRIMERO"}: return "V1"
        if s in {"2","V2","VARIADOR 2","DOS","SEGUNDO"}: return "V2"
        return ""

    @staticmethod
    def _set_combo_text(cb: QComboBox, value: str) -> None:
        txt = (value or "").strip()
        if not txt:
            return
        i = cb.findText(txt, Qt.MatchFixedString)
        if i >= 0:
            cb.setCurrentIndex(i); return
        for j in range(cb.count()):
            if cb.itemText(j).lower() == txt.lower():
                cb.setCurrentIndex(j); return
        for j in range(cb.count()):
            if txt.lower() in cb.itemText(j).lower():
                cb.setCurrentIndex(j); return
        cb.addItem(txt); cb.setCurrentIndex(cb.count() - 1)
