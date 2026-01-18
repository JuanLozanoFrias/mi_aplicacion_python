# -*- coding: utf-8 -*-
# gui/pages/materiales_step2.py
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple
import json
import unicodedata

import pandas as pd
from PySide6.QtCore import Qt, Signal, QTimer, QObject, QEvent
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QFrame,
    QLabel, QLineEdit, QComboBox, QCompleter,
    QRadioButton, QButtonGroup
)

# ---- imports de lÃ³gica con fallback a stubs para que nunca rompa la GUI ----
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

# --- NUEVO: selecciÃ³n conjunta GM+CT exigiendo puente (Q) ---
try:
    from logic.seleccion_gm_ct_con_puente import seleccionar_par_gm_ct_con_puente_abb as _sel_gm_ct_bridge
except Exception:  # pragma: no cover
    def _sel_gm_ct_bridge(book_path: str, corriente: float, arranque: str) -> Dict:
        return {"aplica": False, "motivo": "BRIDGE stub"}

# >>> breaker para variador (ABB) <<<
try:
    from logic.breaker_vfd import seleccionar_breaker_vfd_abb as _sel_brk_vfd
except Exception:  # pragma: no cover
    def _sel_brk_vfd(book_path: str, corriente: float) -> Dict:
        return {"aplica": False, "motivo": "BREAKER stub", "modelo": "", "valor_col_D": 0.0}


# Texto neutro para celdas vacias
EMPTY_CELL = "--"


class _NoWheelFilter(QObject):
    def eventFilter(self, obj, event):  # type: ignore[override]
        if event.type() == QEvent.Wheel:
            return True
        if event.type() == QEvent.MouseButtonPress and event.button() == Qt.MiddleButton:
            return True
        return False


class Step2Panel(QWidget):
    """
    Paso 2 â€” ConfiguraciÃ³n de Compresores.
    Guarda/restaura la selecciÃ³n del variador (bolita) y no la pierde al navegar.
    """

    changed = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._sheet_cache_comp: Dict[str, Optional[pd.DataFrame]] = {}
        self._sheet_cache_var: Dict[str, Optional[pd.DataFrame]] = {}
        self._perf_cache: Optional[Dict[str, object]] = None
        self._corrientes_cache: Optional[Dict[str, object]] = None
        self._pending_corr_updates: set[str] = set()
        self._last_corr_key: Dict[str, tuple] = {}
        self._corr_timer = QTimer(self)
        self._corr_timer.setSingleShot(True)
        self._corr_timer.timeout.connect(self._apply_pending_corr_updates)
        self._step2_state: Dict[str, Dict[str, str]] = {}     # Ãºltimo estado exportado
        self._sel_state: Dict[str, str] = {}                   # **recordatorio de V1/V2 por grupo**
        self._puente_state: Dict[str, Dict[str, str]] = {}     # **NUEVO**: {comp: {"modelo":..., "codigo":...}}
        self._step2_widgets: Dict[str, Dict[str, object]] = {}
        self._no_wheel_filter = _NoWheelFilter(self)

        # fuentes globales (combos del Paso 1)
        self._cb_tipo_comp: Optional[QComboBox] = None
        self._cb_t_alim: Optional[QComboBox] = None
        self._cb_refrig: Optional[QComboBox] = None
        self._cb_marca_var: Optional[QComboBox] = None

        # UI base (sin scroll interno; el contenedor superior ya scrollea)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)
        self.host = QFrame()
        self.host_layout = QVBoxLayout(self.host)
        self.host_layout.setContentsMargins(0, 0, 0, 0)
        self.host_layout.setSpacing(10)
        root.addWidget(self.host, 1)

    # ---------- API pÃºblica ----------
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

        # conserva estado/selecciÃ³n de nombres existentes
        self._step2_state = {k: v for k, v in self._step2_state.items() if k in names_all}
        self._sel_state   = {k: v for k, v in self._sel_state.items()   if k in names_all}
        self._puente_state = {k: v for k, v in self._puente_state.items() if k in names_all}

        # limpia UI
        while self.host_layout.count():
            w = self.host_layout.takeAt(0).widget()
            if w: w.deleteLater()
        self._step2_widgets.clear()

        # secciones
        self._build_comp_section("COMPRESORES BAJA TEMPERATURA", names_g, modelos)
        self._build_comp_section("COMPRESORES MEDIA TEMPERATURA", names_b, modelos)
        self._build_comp_section("COMPRESORES PARALELO", names_f, modelos)

        # refresca corrientes y cÃ¡lculos
        self.refresh_all()

    def refresh_all(self) -> None:
        for name in list(self._step2_widgets.keys()):
            self._update_corriente_for(name)

    def export_state(self) -> Dict[str, Dict[str, str]]:
        # captura del UI actual a dict
        for name, w in self._step2_widgets.items():
            remembered = self._sel_state.get(name, "")
            ui_sel = "V1" if w["rb_v1"].isChecked() else "V2" if w["rb_v2"].isChecked() else ""
            variador_sel = remembered or ui_sel

            pstate = self._puente_state.get(name, {}) or {}
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
                # NUEVO: puente
                "puente_modelo": pstate.get("modelo", ""),
                "puente_codigo": pstate.get("codigo", ""),
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

            # Breaker/GM/Contactor (si venÃ­an)
            if rec.get("guardamotor"):
                w["gm_lbl"].setText(rec["guardamotor"].strip() or "â€”")
            if rec.get("contactor"):
                w["cont_lbl"].setText(rec["contactor"].strip() or "â€”")
            if rec.get("breaker"):
                w["brk_lbl"].setText(rec["breaker"].strip() or "â€”")

            # Puente (si venÃ­a)
            pm = (rec.get("puente_modelo") or "").strip()
            pc = (rec.get("puente_codigo") or "").strip()
            if pm:
                self._puente_state[name] = {"modelo": pm, "codigo": pc}

            # SelecciÃ³n del variador
            vsel = self._normalize_vsel(rec.get("variador_sel", ""))
            if vsel:
                self._sel_state[name] = vsel
                # se marcarÃ¡ al habilitar radios:
                self._apply_remembered_selection(name)

        # No llamo a refresh_all() para no sobrescribir labels; solo
        # aseguro habilitaciÃ³n/selecciÃ³n acorde a arranque actual:
        for name in state.keys():
            self._toggle_var_radios(name)

    def clear(self) -> None:
        self._step2_state.clear()
        self._sel_state.clear()
        self._puente_state.clear()
        self._step2_widgets.clear()
        while self.host_layout.count():
            w = self.host_layout.takeAt(0).widget()
            if w: w.deleteLater()

    # ---------- construcciÃ³n de secciones ----------
    def _build_comp_section(self, titulo: str, names: List[str], modelos: List[str]) -> None:
        if not names:
            return

        box = QFrame()
        box.setStyleSheet("QFrame{background:#ffffff;border:1px solid #d7e3f8;border-radius:10px;}")
        bl = QVBoxLayout(box)
        bl.setContentsMargins(14, 14, 14, 14)
        bl.setSpacing(10)

        lab = QLabel(titulo)
        lab.setStyleSheet("color:#0f172a;font-weight:800;font-size:13px;letter-spacing:0.4px;")
        bl.addWidget(lab)

        grid = QGridLayout()
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(10)
        bl.addLayout(grid)

        headers = [
            "", "MODELO", "CORRIENTE", "ARRANQUE",
            "VARIADOR 1", "VARIADOR 2", "GUARDAMOTOR",
            "MODELO CONTACTOR", "BREAKER (VFD)"
        ]
        for c, h in enumerate(headers):
            hl = QLabel(h)
            hl.setStyleSheet("color:#1f2937;font-weight:800;padding:4px 2px;")
            grid.addWidget(hl, 0, c, alignment=Qt.AlignVCenter)

        combo_style = (
            "QComboBox{background:#ffffff;color:#0f172a;border:1px solid #c8d4eb;"
            "border-radius:6px;padding:4px 8px;padding-right:12px;min-height:24px;}"
            "QComboBox QAbstractItemView{background:#ffffff;color:#0f172a;"
            "selection-background-color:#e6efff;selection-color:#0f172a;}"
            "QComboBox QAbstractItemView::item{color:#0f172a;}"
            "QComboBox:editable{background:#ffffff;color:#0f172a;}"
            "QComboBox QLineEdit{color:#0f172a;padding:0 4px;border:0;}"
        )
        line_style = (
            "QLineEdit{background:#ffffff;color:#0f172a;border:1px solid #c8d4eb;"
            "border-radius:6px;padding:4px 8px;min-height:24px;}"
            "QLineEdit:focus{border:1px solid #5b8bea;}"
            "QLineEdit::selection{background:#e6efff;color:#0f172a;}"
        )
        # Estilo simple y centrado, con fondo transparente cuando está apagado
        radio_style = (
            "QRadioButton{color:#0f172a;background:transparent;padding:0;margin:0;}"
            "QRadioButton:focus{outline:0;}"
            "QRadioButton::indicator{width:14px;height:14px;border-radius:7px;"
            "border:0;background:transparent;subcontrol-position:center;margin:0;padding:0;}"
            "QRadioButton::indicator:unchecked{border:0;background:transparent;}"
            "QRadioButton::indicator:checked{border:0;background:qradialgradient(cx:0.5, cy:0.5, fx:0.5, fy:0.5, radius:0.9, stop:0 #1ecbff, stop:1 #0b6cff);border-radius:7px;}"
            "QRadioButton::indicator:hover{border:0;background:transparent;}"
            "QRadioButton::indicator:checked:hover{border:0;background:qradialgradient(cx:0.5, cy:0.5, fx:0.5, fy:0.5, radius:0.9, stop:0 #22d7ff, stop:1 #0b5de0);}"
            "QRadioButton::indicator:disabled{border:0;background:transparent;}"
            "QRadioButton::indicator:checked:disabled{border:0;background:#9adfff;}"
        )

        for i, name in enumerate(names, start=1):
            row_label = QLabel(f"{name}:")
            row_label.setStyleSheet("color:#0f172a;font-weight:700;padding-right:6px;")
            grid.addWidget(row_label, i, 0, alignment=Qt.AlignVCenter)

            cb_modelo = QComboBox()
            cb_modelo.addItem("")
            cb_modelo.addItems(modelos)
            cb_modelo.setEditable(True)
            cb_modelo.setInsertPolicy(QComboBox.NoInsert)
            cb_modelo.setMinimumWidth(220)
            cb_modelo.setStyleSheet(combo_style)
            cb_modelo.setFixedHeight(26)
            cb_modelo.installEventFilter(self._no_wheel_filter)
            if cb_modelo.lineEdit():
                cb_modelo.lineEdit().setStyleSheet(line_style)
                cb_modelo.lineEdit().setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                cb_modelo.lineEdit().setContentsMargins(0, 0, 0, 0)
                cb_modelo.lineEdit().setFixedHeight(24)
            compl = QCompleter(modelos)
            compl.setFilterMode(Qt.MatchContains)
            compl.setCaseSensitivity(Qt.CaseInsensitive)
            cb_modelo.setCompleter(compl)
            grid.addWidget(cb_modelo, i, 1)

            le_corr = QLineEdit()
            le_corr.setReadOnly(True)
            le_corr.setPlaceholderText("A")
            le_corr.setFixedWidth(120)
            le_corr.setStyleSheet(
                "QLineEdit{background:#ffffff;color:#0f172a;border:1px solid #c8d4eb;"
                "border-radius:8px;padding:6px 8px;}"
            )
            le_corr.setFixedHeight(26)
            grid.addWidget(le_corr, i, 2)

            cb_arranque = QComboBox()
            cb_arranque.addItems(["", "VARIADOR", "DIRECTO", "PARTIDO"])
            cb_arranque.setStyleSheet(combo_style)
            cb_arranque.setMinimumWidth(130)
            cb_arranque.setFixedHeight(26)
            cb_arranque.installEventFilter(self._no_wheel_filter)
            if cb_arranque.lineEdit():
                cb_arranque.lineEdit().setStyleSheet(line_style)
                cb_arranque.lineEdit().setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                cb_arranque.lineEdit().setContentsMargins(0, 0, 0, 0)
                cb_arranque.lineEdit().setFixedHeight(24)
            grid.addWidget(cb_arranque, i, 3)

            v1_layout = QHBoxLayout()
            v1_layout.setSpacing(6)
            rb_v1 = QRadioButton()
            rb_v1.setFocusPolicy(Qt.NoFocus)
            rb_v1.setEnabled(False)
            rb_v1.setStyleSheet(radio_style)
            v1_lbl = QLabel(EMPTY_CELL)
            v1_lbl.setStyleSheet("color:#0f172a;font-weight:600;padding-left:4px;")
            v1_w = QFrame()
            v1_w.setLayout(v1_layout)
            v1_w.setStyleSheet("background:transparent;")
            v1_layout.addWidget(rb_v1)
            v1_layout.addWidget(v1_lbl, 1)
            grid.addWidget(v1_w, i, 4)

            v2_layout = QHBoxLayout()
            v2_layout.setSpacing(6)
            rb_v2 = QRadioButton()
            rb_v2.setFocusPolicy(Qt.NoFocus)
            rb_v2.setEnabled(False)
            rb_v2.setStyleSheet(radio_style)
            v2_lbl = QLabel(EMPTY_CELL)
            v2_lbl.setStyleSheet("color:#0f172a;font-weight:600;padding-left:4px;")
            v2_w = QFrame()
            v2_w.setLayout(v2_layout)
            v2_w.setStyleSheet("background:transparent;")
            v2_layout.addWidget(rb_v2)
            v2_layout.addWidget(v2_lbl, 1)
            grid.addWidget(v2_w, i, 5)
            grp = QButtonGroup(self)
            grp.addButton(rb_v1)
            grp.addButton(rb_v2)

            gm_lbl = QLabel(EMPTY_CELL)
            gm_lbl.setStyleSheet("color:#0f172a;font-weight:600;")
            cont_lbl = QLabel(EMPTY_CELL)
            cont_lbl.setStyleSheet("color:#0f172a;font-weight:600;")
            brk_lbl = QLabel(EMPTY_CELL)
            brk_lbl.setStyleSheet("color:#0f172a;font-weight:600;")

            grid.addWidget(gm_lbl, i, 6)
            grid.addWidget(cont_lbl, i, 7)
            grid.addWidget(brk_lbl, i, 8)

            if name in self._step2_state:
                prev = self._step2_state[name]
                ixm = cb_modelo.findText(prev.get("modelo", ""), Qt.MatchFixedString)
                if ixm >= 0:
                    cb_modelo.setCurrentIndex(ixm)
                    cb_modelo.lineEdit().setText(cb_modelo.itemText(ixm))
                ixa = cb_arranque.findText(prev.get("arranque", ""), Qt.MatchFixedString)
                if ixa >= 0:
                    cb_arranque.setCurrentIndex(ixa)
                if prev.get("corriente"):
                    le_corr.setText(prev["corriente"])
                if prev.get("guardamotor"):
                    gm_lbl.setText(prev["guardamotor"])
                if prev.get("variador1"):
                    v1_lbl.setText(prev["variador1"])
                if prev.get("variador2"):
                    v2_lbl.setText(prev["variador2"])
                if prev.get("contactor"):
                    cont_lbl.setText(prev["contactor"])
                if prev.get("breaker"):
                    brk_lbl.setText(prev["breaker"])
                vsel = self._normalize_vsel(prev.get("variador_sel", ""))
                if vsel:
                    self._sel_state[name] = vsel
                    if vsel == "V1":
                        rb_v1.setChecked(True)
                    elif vsel == "V2":
                        rb_v2.setChecked(True)
                pm = (prev.get("puente_modelo") or "").strip()
                pc = (prev.get("puente_codigo") or "").strip()
                if pm:
                    self._puente_state[name] = {"modelo": pm, "codigo": pc}

            def _ensure_valid_and_current(cmb=cb_modelo, n=name, le=le_corr):
                txt = cmb.lineEdit().text().strip()
                if not txt:
                    cmb.setCurrentIndex(0)
                    cmb.lineEdit().clear()
                    le.setText("")
                    self._set_gm(n, "")
                    self._set_cont(n, "")
                    self._set_vars(n, "", "")
                    self._set_breaker(n, "")
                    self._puente_state.pop(n, None)
                    self._toggle_var_radios(n)
                    return
                ix = cmb.findText(txt, Qt.MatchFixedString)
                if ix < 0:
                    for j in range(1, cmb.count()):
                        if txt.lower() in cmb.itemText(j).lower():
                            ix = j
                            break
                if ix >= 0:
                    cmb.setCurrentIndex(ix)
                    cmb.lineEdit().setText(cmb.itemText(ix))
                else:
                    cmb.setCurrentIndex(0)
                    cmb.lineEdit().clear()
                self._schedule_corr_update(n)

            cb_modelo.lineEdit().editingFinished.connect(_ensure_valid_and_current); self.changed.emit()
            cb_modelo.currentTextChanged.connect(lambda _=None, n=name: (self._schedule_corr_update(n), self.changed.emit()))
            cb_arranque.currentIndexChanged.connect(lambda _=None, n=name: (self._on_arranque_changed(n), self.changed.emit()))
            rb_v1.toggled.connect(lambda _=None, n=name: (self._on_var_selection_changed(n), self.changed.emit()))
            rb_v2.toggled.connect(lambda _=None, n=name: (self._on_var_selection_changed(n), self.changed.emit()))

            self._step2_widgets[name] = {
                "modelo": cb_modelo,
                "corriente": le_corr,
                "arranque": cb_arranque,
                "rb_v1": rb_v1,
                "rb_v2": rb_v2,
                "v1_lbl": v1_lbl,
                "v2_lbl": v2_lbl,
                "grp": grp,
                "gm_lbl": gm_lbl,
                "cont_lbl": cont_lbl,
                "brk_lbl": brk_lbl,
            }

        self.host_layout.addWidget(box)
        grid.setColumnStretch(1, 2)
        for c in (2, 3, 4, 5, 6, 7, 8):
            grid.setColumnStretch(c, 1)

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

    def _schedule_corr_update(self, name: str) -> None:
        self._pending_corr_updates.add(name)
        # debounce rápido para evitar recalcular dos veces por el mismo cambio
        self._corr_timer.start(120)

    def _apply_pending_corr_updates(self) -> None:
        pending = list(self._pending_corr_updates)
        self._pending_corr_updates.clear()
        for name in pending:
            self._update_corriente_for(name)

    def _update_corriente_for(self, name: str) -> None:
        w = self._step2_widgets.get(name)
        if not w: return

        modelo = w["modelo"].currentText().strip()
        le = w["corriente"]
        marca = (self._cb_tipo_comp.currentText() or "").upper().strip() if self._cb_tipo_comp else ""
        tension = (self._cb_t_alim.currentText() or "").strip() if self._cb_t_alim else ""
        refrig = (self._cb_refrig.currentText() or "").upper().strip() if self._cb_refrig else ""
        key = (modelo, marca, tension, refrig)
        if key == self._last_corr_key.get(name):
            return
        if not modelo:
            le.setText("")
            self._set_gm(name,""); self._set_cont(name,""); self._set_vars(name,"","")
            self._set_breaker(name,""); self._puente_state.pop(name, None); self._toggle_var_radios(name); return

        amps_raw = ""
        if refrig in ("R744", "CO2", "R-744", "R507", "R-507", "R507A", "R-507A"):
            amps_raw = self._buscar_corriente_json(marca, modelo, tension, refrig)
            if not amps_raw and refrig in ("R744", "CO2", "R-744"):
                amps_raw = self._buscar_corriente_en_hoja(marca, modelo, tension)
        else:
            le.setText("")
            self._set_gm(name,""); self._set_cont(name,""); self._set_vars(name,"","")
            self._set_breaker(name,""); self._puente_state.pop(name, None); self._toggle_var_radios(name); return
        if not amps_raw:
            le.setText("")
            self._set_gm(name,""); self._set_cont(name,""); self._set_vars(name,"","")
            self._set_breaker(name,""); self._puente_state.pop(name, None); self._toggle_var_radios(name); return

        try:
            val = float(str(amps_raw).replace(",", ".")); le.setText(f"{val:.1f} A")
        except ValueError:
            le.setText(""); self._set_gm(name,""); self._set_cont(name,""); self._set_vars(name,"","")
            self._set_breaker(name,""); self._puente_state.pop(name, None); self._toggle_var_radios(name); return

        self._last_corr_key[name] = key
        # despuÃ©s de corriente: recalcular
        self._recalc_guardamotor(name)
        self._recalc_contactor(name)
        self._recalc_variadores(name)
        self._recalc_breaker_vfd(name)

    def _recalc_guardamotor(self, name: str) -> None:
        """NUEVO: selecciona GM y CT en conjunto exigiendo puente (Q)."""
        w = self._step2_widgets.get(name)
        if not w: return
        arr = (w["arranque"].currentText() or "").upper().strip()
        if arr not in ("DIRECTO","PARTIDO"):
            self._set_gm(name, "NO APLICA" if arr=="VARIADOR" else "")
            self._puente_state.pop(name, None)
            return
        amps = self._get_amps(name)
        if amps is None:
            self._set_gm(name,""); self._puente_state.pop(name, None); return

        res = _sel_gm_ct_bridge(self._book().as_posix(), amps, arr)
        if res.get("aplica"):
            self._set_gm(name, f"{res['guardamotor']['modelo']}".strip())
            # tambiÃ©n fijamos el contactor y puente:
            self._set_cont(name, f"{res['contactor']['modelo']}".strip())
            self._puente_state[name] = {
                "modelo": res["puente"]["modelo"],
                "codigo": res["puente"]["codigo"],
            }
        else:
            self._puente_state.pop(name, None)
            self._set_gm(name, res.get("motivo") or res.get("error","NO ENCONTRADO"))

    def _recalc_contactor(self, name: str) -> None:
        """NUEVO: espejo de guardamotor para mantener coherencia en pantalla."""
        w = self._step2_widgets.get(name)
        if not w: return
        arr = (w["arranque"].currentText() or "").upper().strip()
        if arr not in ("DIRECTO","PARTIDO"):
            self._set_cont(name, "NO APLICA" if arr=="VARIADOR" else "")
            return
        amps = self._get_amps(name)
        if amps is None:
            self._set_cont(name,""); return

        res = _sel_gm_ct_bridge(self._book().as_posix(), amps, arr)
        if res.get("aplica"):
            self._set_cont(name, f"{res['contactor']['modelo']}".strip())
            self._set_gm(name, f"{res['guardamotor']['modelo']}".strip())
            self._puente_state[name] = {
                "modelo": res["puente"]["modelo"],
                "codigo": res["puente"]["codigo"],
            }
        else:
            self._puente_state.pop(name, None)
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
            self._set_breaker(name, f"{res.get('C', res.get('modelo',''))}".strip() or EMPTY_CELL)
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
        lab = self._step2_widgets.get(name, {}).get("gm_lbl")
        if isinstance(lab, QLabel):
            lab.setText(text or EMPTY_CELL)

    def _set_cont(self, name: str, text: str) -> None:
        lab = self._step2_widgets.get(name, {}).get("cont_lbl")
        if isinstance(lab, QLabel):
            lab.setText(text or EMPTY_CELL)

    def _set_breaker(self, name: str, text: str) -> None:
        lab = self._step2_widgets.get(name, {}).get("brk_lbl")
        if isinstance(lab, QLabel):
            lab.setText(text or EMPTY_CELL)

    def _set_vars(self, name: str, v1: str, v2: str) -> None:
        w = self._step2_widgets.get(name)
        if not w:
            return
        if isinstance(w.get("v1_lbl"), QLabel):
            w["v1_lbl"].setText(v1 or EMPTY_CELL)
        if isinstance(w.get("v2_lbl"), QLabel):
            w["v2_lbl"].setText(v2 or EMPTY_CELL)

    def _apply_remembered_selection(self, name: str) -> None:
        """Marca la radio guardada en _sel_state si estÃ¡ habilitada."""
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
        v1_ok = (w["v1_lbl"].text() or EMPTY_CELL).strip() != EMPTY_CELL
        v2_ok = (w["v2_lbl"].text() or EMPTY_CELL).strip() != EMPTY_CELL
        enable = (arr == "VARIADOR") and (v1_ok or v2_ok)
        w["rb_v1"].setEnabled(enable and v1_ok)
        w["rb_v2"].setEnabled(enable and v2_ok)

        if not enable:
            # No borrar _sel_state; solo desmarcar UI
            w["rb_v1"].setAutoExclusive(False); w["rb_v1"].setChecked(False); w["rb_v1"].setAutoExclusive(True)
            w["rb_v2"].setAutoExclusive(False); w["rb_v2"].setChecked(False); w["rb_v2"].setAutoExclusive(True)
            return

        # Si hay recordatorio, respÃ©talo
        remembered = self._sel_state.get(name, "")
        if remembered == "V1" and v1_ok:
            w["rb_v1"].setChecked(True); return
        if remembered == "V2" and v2_ok:
            w["rb_v2"].setChecked(True); return

        # Si no hay recordatorio: auto-seleccionar si sÃ³lo hay uno disponible
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

    def _buscar_corriente_json(self, marca: str, modelo: str, tension: str, refrigerante: str) -> str:
        data = self._load_comp_corrientes()
        ref_key = self._normalize_refrig(refrigerante).replace("-", "")
        if ref_key in ("R507", "R507A"):
            ref = "R507"
        elif ref_key in ("R744", "CO2"):
            ref = "R744"
        else:
            return ""

        brands = data.get("refrigerants", {}).get(ref, {}).get("brands", {})
        bdata = brands.get((marca or "").upper().strip())
        if not isinstance(bdata, dict):
            return ""
        models = bdata.get("models", {})
        mdata = models.get(modelo)
        if not isinstance(mdata, dict):
            return ""

        if ref == "R744":
            currents = mdata.get("currents", {})
            val = self._pick_current_for_tension(currents, tension)
            if val is None:
                return ""
            try:
                return str(float(val))
            except ValueError:
                return ""

        # R507: primero intenta corrientes directas por tensión (si existen)
        currents = mdata.get("currents", {})
        val = self._pick_current_for_tension(currents, tension)
        if val is not None:
            try:
                return str(float(val))
            except ValueError:
                return ""

        # R507: fallback a points con RLA (nearest)
        ref_norm = self._normalize_refrig(refrigerante)
        points = mdata.get("points", {})
        if not isinstance(points, dict) or not points:
            return ""

        tcond_target = bdata.get("tcond_f")
        tevaps = bdata.get("tevaps_f") or []
        try:
            tevap_target = min(float(x) for x in tevaps) if tevaps else None
        except Exception:
            tevap_target = None

        best = None
        best_dist = None
        for key, pt in points.items():
            parts = str(key).split("|")
            if len(parts) < 3:
                continue
            ref, tcond_s, tevap_s = parts[0], parts[1], parts[2]
            if self._normalize_refrig(ref) != ref_norm:
                continue
            try:
                tcond_v = float(tcond_s)
                tevap_v = float(tevap_s)
            except ValueError:
                continue
            dist = 0.0
            if tcond_target is not None:
                dist += abs(tcond_v - float(tcond_target))
            if tevap_target is not None:
                dist += abs(tevap_v - float(tevap_target))
            if best is None or dist < (best_dist or 0.0):
                best = pt
                best_dist = dist

        if not isinstance(best, dict):
            return ""
        rla = best.get("rla")
        if rla is None:
            return ""
        try:
            return str(float(rla))
        except ValueError:
            return ""

    @staticmethod
    def _normalize_refrig(refrig: str) -> str:
        raw = (refrig or "").upper().replace(" ", "").replace("-", "")
        if raw in ("R507", "R507A"):
            return "R-507"
        if raw in ("R744", "CO2"):
            return "R-744"
        return (refrig or "").upper().strip()

    @staticmethod
    def _pick_current_for_tension(currents: object, tension: str):
        if not isinstance(currents, dict) or not currents:
            return None
        t = "".join(ch for ch in str(tension) if ch.isdigit())
        candidates = []
        if t in ("220", "230", "208"):
            candidates = ["220", "230", "208"]
        elif t in ("460", "480", "440"):
            candidates = ["460", "480", "440"]
        elif t:
            candidates = [t]
        for c in candidates:
            if c in currents and currents[c] not in ("", None):
                return currents[c]
        # fallback: si solo hay un voltaje, usarlo
        if len(currents) == 1:
            return next(iter(currents.values()))
        return None

    def _load_comp_perf(self) -> Dict[str, object]:
        if self._perf_cache is not None:
            return self._perf_cache
        base = Path(__file__).resolve().parents[3] / "data"
        path = base / "LEGEND" / "compresores_perf.json"
        if not path.exists():
            path = base / "legend" / "compresores_perf.json"
        try:
            with open(path, "r", encoding="utf-8") as f:
                self._perf_cache = json.load(f)
        except Exception:
            self._perf_cache = {}
        return self._perf_cache

    def _load_comp_corrientes(self) -> Dict[str, object]:
        if self._corrientes_cache is not None:
            return self._corrientes_cache
        base = Path(__file__).resolve().parents[3] / "data"
        path = base / "tableros_electricos" / "compresores_corrientes.json"
        if not path.exists():
            path = base / "legend" / "compresores_corrientes.json"
        if not path.exists():
            path = base / "LEGEND" / "compresores_corrientes.json"
        try:
            with open(path, "r", encoding="utf-8") as f:
                self._corrientes_cache = json.load(f)
        except Exception:
            self._corrientes_cache = {}
        return self._corrientes_cache

    def _load_var_sheet(self, marca_var: str) -> Optional[pd.DataFrame]:
        def norm(s: str) -> str:
            raw = (s or "").strip().upper()
            normed = unicodedata.normalize("NFKD", raw).encode("ascii", "ignore").decode("ascii")
            return "".join(ch for ch in normed if ch.isalnum())
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
        # archivo Excel en la raíz del proyecto (data/basedatos.xlsx)
        return Path(__file__).resolve().parents[3] / "data" / "tableros_electricos" / "basedatos.xlsx"

    # ---------- helpers pÃºblicos para orquestador ----------
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
            row["v1_lbl"].setText(variador1 or "â€”")
        if isinstance(row.get("v2_lbl"), QLabel) and variador2:
            row["v2_lbl"].setText(variador2 or "â€”")

        if variador_sel:
            self.set_variador_selection(key, variador_sel)

        # actualizar habilitaciÃ³n acorde a arranque/variadores
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

    # ---------- utilidades pequeÃ±as ----------
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
