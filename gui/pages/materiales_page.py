# gui/pages/materiales_page.py
from __future__ import annotations

import json
import random
from pathlib import Path
from typing import Dict, List, Optional, Any

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QFrame,
    QLabel, QLineEdit, QComboBox, QPushButton, QStackedWidget,
    QFileDialog, QMessageBox
)

# Subpaneles
from .materiales_step2 import Step2Panel
from .materiales_step3 import Step3OptionsPanel
from .materiales_step4 import MaterialesStep4Page  # ← unificado

# Loader de “programación” (.ecalc.json)
from logic.programacion_loader import load_programacion_snapshot

# Flag de autofill (Página 1 y 2). Lo tomamos del módulo dev si existe.
try:
    from logic.dev_autofill import DEV_AUTOFILL  # True = activa autorrelleno
except Exception:  # pragma: no cover
    DEV_AUTOFILL = False

# --- Validación laxa del Paso 1 ---
STRICT_VALIDATION = False


class MaterialesPage(QWidget):
    """
    Página 'Materiales' (orquestador)
    """

    def __init__(self) -> None:
        super().__init__()
        self._df_master: Optional[pd.DataFrame] = None
        self._loaded_step2_state: Optional[Dict[str, Dict[str, str]]] = None
        self._loaded_step3_state: Optional[Dict[str, Any]] = None
        self._build_ui()

    # --------------------------------------------------------------------- UI
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(40, 40, 40, 40)
        root.setSpacing(20)

        self.step_stack = QStackedWidget()
        root.addWidget(self.step_stack, 1)

        # ========================= PASO 1 =========================
        p1 = QFrame()
        p1.setStyleSheet("background:#2a2d3c;border-radius:12px;")
        p1l = QVBoxLayout(p1)
        p1l.setContentsMargins(24, 20, 24, 20)
        p1l.setSpacing(16)

        # Barra superior
        tools = QHBoxLayout(); tools.setSpacing(10)
        self.btn_load = QPushButton("CARGAR INFORMACIÓN DEL PROYECTO")
        self.btn_new  = QPushButton("PROYECTO NUEVO")
        ghost_qss = """
            QPushButton{
                background:transparent;color:#cfd7ff;border:1px solid #46517a;
                border-radius:8px;padding:8px 12px;font-weight:700;
            }
            QPushButton:hover{ background:#3a425e; color:#ffffff; }
        """
        self.btn_load.setStyleSheet(ghost_qss)
        self.btn_new.setStyleSheet(ghost_qss)
        self.btn_load.clicked.connect(self._on_load_project)
        self.btn_new.clicked.connect(self._on_new_project)
        tools.addWidget(self.btn_load); tools.addWidget(self.btn_new); tools.addStretch(1)
        p1l.addLayout(tools)

        # Título
        title = QLabel("CÁLCULO DE MATERIALES (TABLEROS ELÉCTRICOS)")
        title.setStyleSheet("font-size:22px;font-weight:700;color:#e0e0e0;")
        title.setAlignment(Qt.AlignCenter)
        p1l.addWidget(title)

        # Fila: nombre del proyecto (a lo ancho)
        line1 = QHBoxLayout()
        self.proj = QLineEdit()
        self.city = QLineEdit()  # usado en grilla izquierda

        # Forzar MAYÚSCULAS
        self.proj.textChanged.connect(lambda t: self._force_upper(self.proj, t))
        self.city.textChanged.connect(lambda t: self._force_upper(self.city, t))

        lab_proj = QLabel("NOMBRE DEL PROYECTO:")
        lab_proj.setStyleSheet("color:#e0e0e0; font-weight:600;")
        line1.addWidget(lab_proj); line1.addWidget(self.proj, 1)
        p1l.addLayout(line1)

        # Grilla 2xN compacta
        g = QGridLayout()
        g.setHorizontalSpacing(18)
        g.setVerticalSpacing(10)

        # Controles
        self.resp = QComboBox(); self.resp.addItems(["", "JUAN DAVID LOZANO"])
        self.t_alim = QComboBox(); self.t_alim.addItems(["", "220", "460"])
        self.refrig = QComboBox(); self.refrig.addItems(["", "R744", "R290", "R507", "R404"])
        self.t_ctl  = QComboBox(); self.t_ctl.addItems(["", "120", "220"])
        self.norma_ap = QComboBox(); self.norma_ap.addItems(["", "UL", "IEC"])
        self.tipo_comp = QComboBox(); self.tipo_comp.addItems(
            ["", "BITZER", "COPELAND", "TECUMSEH", "FRASCOLD", "DORIN", "OTRO"]
        )

        self.n_comp_media = QLineEdit(); self.n_comp_paral = QLineEdit(); self.n_comp_baja = QLineEdit()
        for le in (self.n_comp_media, self.n_comp_paral, self.n_comp_baja):
            le.setPlaceholderText("0")

        self.marca_elem = QComboBox(); self.marca_elem.addItems(
            ["", "ABB", "SIEMENS", "SCHNEIDER", "LS", "WEG", "CHINT", "DELIXY", "RITTAL", "MURRELEKTRONIK", "GENERICO"]
        )
        self.marca_var  = QComboBox(); self.marca_var.addItems(
            ["", "NO", "ABB", "DANFOSS", "SCHNEIDER", "YASKAWA", "DELTA", "CHINT", "DELIXY"]
        )

        def add_pair(row: int, right: bool, text: str, widget: QWidget) -> None:
            c0 = 2 if right else 0
            c1 = 3 if right else 1
            lab = QLabel(text); lab.setStyleSheet("color:#e0e0e0; font-weight:600;")
            g.addWidget(lab, row, c0)
            g.addWidget(widget, row, c1)

        left_pairs = [
            ("CIUDAD:", self.city),
            ("TENSIÓN ALIMENTACIÓN:", self.t_alim),
            ("TENSIÓN DE CONTROL:", self.t_ctl),
            ("TIPO DE COMPRESORES (MARCA):", self.tipo_comp),
            ("Nº COMPRESORES MEDIA:", self.n_comp_media),
            ("Nº COMPRESORES BAJA:", self.n_comp_baja),
        ]
        right_pairs = [
            ("RESPONSABLE:", self.resp),
            ("REFRIGERANTE:", self.refrig),
            ("NORMA APLICABLE:", self.norma_ap),
            ("Nº COMPRESORES PARALELO:", self.n_comp_paral),
            ("MARCA DE ELEMENTOS:", self.marca_elem),
            ("MARCA VARIADORES:", self.marca_var),
        ]
        for r, (txt, w) in enumerate(left_pairs):
            add_pair(r, False, txt, w)
        for r, (txt, w) in enumerate(right_pairs):
            add_pair(r, True, txt, w)

        g.setColumnStretch(0, 0)
        g.setColumnStretch(1, 1)
        g.setColumnStretch(2, 0)
        g.setColumnStretch(3, 1)
        p1l.addLayout(g)

        # Navegación Paso 1
        bottom = QHBoxLayout(); bottom.addStretch(1)
        self.btn_next_p1 = QPushButton("SIGUIENTE ▶")
        self.btn_next_p1.setStyleSheet("""
            QPushButton{background:#5b8bea;color:#fff;font-weight:700;border:none;border-radius:8px;padding:10px 18px;}
            QPushButton:hover{ background:#6c9cf0; }
        """)
        self.btn_next_p1.clicked.connect(self._on_next_from_form)
        bottom.addWidget(self.btn_next_p1, 0, Qt.AlignRight)
        p1l.addLayout(bottom)

        self.step_stack.addWidget(p1)  # index 0

        if DEV_AUTOFILL:
            self._dev_prefill_form()

        # ========================= PASO 2 =========================
        p2 = QFrame(); p2.setStyleSheet("background:#2a2d3c;border-radius:12px;")
        p2l = QVBoxLayout(p2); p2l.setContentsMargins(24, 20, 24, 20); p2l.setSpacing(12)

        t2 = QLabel("PASO 2 — CONFIGURACIÓN DE COMPRESORES")
        t2.setStyleSheet("font-size:20px;font-weight:700;color:#e0e0e0;")
        p2l.addWidget(t2)

        self.step2_panel = Step2Panel()
        # Conecta combos globales para que Step2Panel pueda leerlos/escucharlos
        self.step2_panel.set_globals(self.tipo_comp, self.t_alim, self.refrig, self.marca_var)
        p2l.addWidget(self.step2_panel, 1)

        nav2 = QHBoxLayout()
        self.btn_prev2 = QPushButton("◀ ATRÁS")
        self.btn_next2 = QPushButton("SIGUIENTE ▶")
        for b in (self.btn_prev2, self.btn_next2):
            b.setStyleSheet("""
                QPushButton{background:#5b8bea;color:#fff;font-weight:700;border:none;border-radius:8px;padding:10px 18px;}
                QPushButton:hover{ background:#6c9cf0; }
            """)
        self.btn_prev2.clicked.connect(lambda: self.step_stack.setCurrentIndex(0))
        self.btn_next2.clicked.connect(self._on_next_from_calc)
        nav2.addWidget(self.btn_prev2); nav2.addStretch(1); nav2.addWidget(self.btn_next2)
        p2l.addLayout(nav2)

        self.step_stack.addWidget(p2)  # index 1

        # ========================= PASO 3 =========================
        p3 = QFrame(); p3.setStyleSheet("background:#2a2d3c;border-radius:12px;")
        p3l = QVBoxLayout(p3); p3l.setContentsMargins(24, 20, 24, 20); p3l.setSpacing(12)

        t3 = QLabel("PASO 3 — OPCIONES CO2")
        t3.setStyleSheet("font-size:20px;font-weight:700;color:#e0e0e0;")
        p3l.addWidget(t3)

        self.step3_panel = Step3OptionsPanel()
        p3l.addWidget(self.step3_panel, 1)

        nav3 = QHBoxLayout()
        self.btn_prev3 = QPushButton("◀ ATRÁS")
        self.btn_to_step4 = QPushButton("RESUMEN ▶")  # va al Paso 4
        for b in (self.btn_prev3, self.btn_to_step4):
            b.setStyleSheet("""
                QPushButton{background:#5b8bea;color:#fff;font-weight:700;border:none;border-radius:8px;padding:10px 18px;}
                QPushButton:hover{ background:#6c9cf0; }
            """)
        self.btn_prev3.clicked.connect(lambda: self.step_stack.setCurrentIndex(1))
        self.btn_to_step4.clicked.connect(self._go_to_step4)
        nav3.addWidget(self.btn_prev3); nav3.addStretch(1); nav3.addWidget(self.btn_to_step4)
        p3l.addLayout(nav3)

        self.step_stack.addWidget(p3)  # index 2

        # ========================= PASO 4 =========================
        p4 = QFrame(); p4.setStyleSheet("background:#2a2d3c;border-radius:12px;")
        p4l = QVBoxLayout(p4); p4l.setContentsMargins(24, 20, 24, 20); p4l.setSpacing(12)

        t4 = QLabel("PASO 4 — RESUMEN")
        t4.setStyleSheet("font-size:20px;font-weight:700;color:#e0e0e0;")
        p4l.addWidget(t4)

        # Paso 4 unificado
        self.page4 = MaterialesStep4Page(
            get_step2_state=lambda: self.step2_panel.export_state(),
            get_globals=self._globals_for_step4,
        )
        p4l.addWidget(self.page4, 1)

        nav4 = QHBoxLayout()
        self.btn_prev4 = QPushButton("◀ ATRÁS")
        for b in (self.btn_prev4,):
            b.setStyleSheet("""
                QPushButton{background:#5b8bea;color:#fff;font-weight:700;border:none;border-radius:8px;padding:10px 18px;}
                QPushButton:hover{ background:#6c9cf0; }
            """)
        self.btn_prev4.clicked.connect(lambda: self.step_stack.setCurrentIndex(2))
        nav4.addWidget(self.btn_prev4); nav4.addStretch(1)
        p4l.addLayout(nav4)

        self.step_stack.addWidget(p4)  # index 3

        # Cambios globales → refrescar Step2Panel
        self.t_alim.currentIndexChanged.connect(self.step2_panel.refresh_all)
        self.refrig.currentIndexChanged.connect(self.step2_panel.refresh_all)
        self.tipo_comp.currentIndexChanged.connect(self.step2_panel.refresh_all)
        self.marca_var.currentIndexChanged.connect(self.step2_panel.refresh_all)

    # -------------------------------------------------------- Navegación
    def _to_int(self, s: str) -> int:
        s = "".join(ch for ch in (s or "") if ch.isdigit())
        return int(s) if s else 0

    def _on_next_from_form(self) -> None:
        if not self._validate_required():
            return

        nb = self._to_int(self.n_comp_baja.text())
        nm = self._to_int(self.n_comp_media.text())
        np = self._to_int(self.n_comp_paral.text())

        marca = (self.tipo_comp.currentText() or "OTRO").upper().strip()
        tension = (self.t_alim.currentText() or "220").strip()
        refrigerante = (self.refrig.currentText() or "").upper().strip()

        if refrigerante == "R744" and self.step2_panel.load_comp_sheet(marca) is None:
            self._warn(f"No hay base de datos para compresores de la marca '{marca}' (hoja COMP{marca}). Se usará una lista genérica.")

        modelos = self._get_modelos_r744_por_hoja(marca) if refrigerante == "R744" \
                  else self._get_modelos_master(marca, tension)

        self.step2_panel.rebuild(nb, nm, np, modelos)

        # Aplicar import de Step2 si venimos de cargar una programación
        if self._loaded_step2_state:
            self._apply_loaded_step2(self._loaded_step2_state)
            self._loaded_step2_state = None

        if DEV_AUTOFILL and not self.step2_panel.export_state():
            try:
                plan_arranques = ["PARTIDO", "PARTIDO", "VARIADOR", "VARIADOR", "DIRECTO", "DIRECTO"]
                keys = self.step2_panel.get_group_keys() if hasattr(self.step2_panel, "get_group_keys") else []
                n = min(len(keys), len(plan_arranques))
                if modelos:
                    pool = modelos[:]
                    random.shuffle(pool)
                    picks = (pool[:n] if len(pool) >= n else (pool * (n // max(1, len(pool)) + 1))[:n])
                else:
                    picks = [""] * n

                for i in range(n):
                    self.step2_panel.set_fields(
                        keys[i],
                        modelo=picks[i] or None,
                        arranque=plan_arranques[i],
                        corriente=None
                    )
                self.step2_panel.refresh_all()
            except Exception:
                pass

        self.step_stack.setCurrentIndex(1)

    def _on_next_from_calc(self) -> None:
        _ = self.step2_panel.export_state()
        # Cargar opciones una sola vez y, si venimos con snapshot, aplicarlas
        if hasattr(self.step3_panel, "load_options"):
            try:
                self.step3_panel.load_options(initial_state=self._loaded_step3_state, force=False)
            except TypeError:
                # Compat con versiones que no aceptan initial_state/force
                self.step3_panel.load_options()
                if self._loaded_step3_state and hasattr(self.step3_panel, "import_state"):
                    try:
                        self.step3_panel.import_state(self._loaded_step3_state)
                    except Exception:
                        pass
        self._loaded_step3_state = None
        self.step_stack.setCurrentIndex(2)

    def _go_to_step4(self) -> None:
        try:
            _ = self.step2_panel.export_state()
            self.page4.reload_and_render()
        except Exception as e:
            print(f"[Paso4] error renderizando: {e}")
        self.step_stack.setCurrentIndex(3)

    # ------------------------------------------------------ Validación
    def _validate_required(self) -> bool:
        if not STRICT_VALIDATION:
            return True
        if not self.proj.text().strip() or not self.city.text().strip():
            self._warn("Favor completar todos los campos."); return False
        for cb in [self.resp, self.t_alim, self.refrig, self.t_ctl, self.norma_ap,
                   self.tipo_comp, self.marca_elem, self.marca_var]:
            if cb.currentIndex() <= 0:
                self._warn("Favor completar todos los campos."); return False
        for le in [self.n_comp_media, self.n_comp_paral, self.n_comp_baja]:
            t = le.text().strip()
            if t == "" or not t.isdigit():
                self._warn("Favor completar todos los campos."); return False
        return True

    # ----------------------------------------------- Archivo / Proyecto
    def _on_new_project(self) -> None:
        self.proj.clear(); self.city.clear()
        for cb in (self.resp, self.t_alim, self.refrig, self.t_ctl, self.norma_ap,
                   self.tipo_comp, self.marca_elem, self.marca_var):
            cb.setCurrentIndex(0)
        for le in (self.n_comp_media, self.n_comp_paral, self.n_comp_baja):
            le.clear()
        self.step2_panel.clear()
        self._loaded_step2_state = None
        self._loaded_step3_state = None
        self.step_stack.setCurrentIndex(0)
        if DEV_AUTOFILL:
            self._dev_prefill_form()

    def _on_load_project(self) -> None:
        fname, _ = QFileDialog.getOpenFileName(
            self,
            "Cargar información del proyecto",
            str(Path.home() / "Documents" / "ElectroCalc_Exports"),
            "Programación (*.ecalc.json);;Proyecto JSON (*.json);;Excel (*.xlsx *.xls)"
        )
        if not fname:
            return

        try:
            # Prioridad: snapshot nativo .ecalc.json
            if fname.lower().endswith(".ecalc.json"):
                step2, globs = load_programacion_snapshot(fname)

                # Limpiar primero
                self._on_new_project()

                # Campos del Paso 1 desde globs
                self.proj.setText(self._get_any(globs, ["nombre_proyecto", "proyecto", "project_name"]))
                self.city.setText(self._get_any(globs, ["ciudad", "city"]))

                self._set_combo(self.resp, self._get_any(globs, ["responsable"]))
                self._set_combo(self.t_alim, self._digits(self._get_any(globs, ["t_alim", "tension_alimentacion"])))
                self._set_combo(self.t_ctl,  self._digits(self._get_any(globs, ["t_ctl", "tension_control"])))
                self._set_combo(self.refrig, self._get_any(globs, ["refrigerante"]))
                self._set_combo(self.norma_ap, (self._get_any(globs, ["norma_ap", "norma aplicable"]) or "").upper())
                self._set_combo(self.tipo_comp, (self._get_any(globs, ["tipo_compresor", "tipo_compresores", "marca_compresor"]) or "").upper())
                self._set_combo(self.marca_elem, (self._get_any(globs, ["marca_elem", "marca_elementos"]) or "").upper())
                self._set_combo(self.marca_var,  (self._get_any(globs, ["marca_variadores", "marca_var"]) or "").upper())

                # Estimar cantidades desde step2 (prefijos B y M/G)
                nb = sum(1 for k in (step2 or {}) if str(k).upper().startswith("B"))
                nm_M = sum(1 for k in (step2 or {}) if str(k).upper().startswith("M"))
                nm_G = sum(1 for k in (step2 or {}) if str(k).upper().startswith("G"))
                nm = nm_M if nm_M else nm_G
                self.n_comp_baja.setText(str(nb))
                self.n_comp_media.setText(str(nm))
                npar = self._get_any(globs, ["n_comp_paralelo", "n paralelos", "paralelo", "comp paralelos"])
                if npar:
                    self.n_comp_paral.setText(self._digits(npar))

                # Guardamos estados para aplicarlos más adelante
                self._loaded_step2_state = step2
                self._loaded_step3_state = globs.get("step3_state", {}) if isinstance(globs, dict) else {}

                self._info("✅ Programación cargada. Pulsa ‘Siguiente’ para continuar.")
                return

            # --------- Compat: JSON simple o Excel de key/valor ---------
            data: Dict[str, str] = {}
            if fname.lower().endswith(".json"):
                with open(fname, "r", encoding="utf-8") as f:
                    raw = json.load(f)
                data = {self._norm_key(k): (v if isinstance(v, str) else str(v)) for k, v in raw.items()}
            else:
                df = pd.read_excel(fname, sheet_name=0, usecols=[0, 1], header=None)
                for _, row in df.iterrows():
                    k = self._norm_key(str(row[0]))
                    v = "" if pd.isna(row[1]) else str(row[1])
                    data[k] = v

            self._on_new_project()

            self.proj.setText(data.get("NOMBRE DEL PROYECTO", data.get("PROYECTO", "")))
            self.city.setText(data.get("CIUDAD", ""))
            self._set_combo(self.resp, data.get("RESPONSABLE", ""))

            self._set_combo(self.t_alim, data.get("TENSION ALIMENTACION", data.get("TENSIÓN ALIMENTACIÓN", "")))
            self._set_combo(self.refrig, data.get("REFRIGERANTE", ""))
            self._set_combo(self.t_ctl, data.get("TENSION CONTROL", data.get("TENSIÓN DE CONTROL", "")))
            self._set_combo(self.norma_ap, (data.get("NORMA APLICABLE", "") or "").upper())
            self._set_combo(self.tipo_comp, (data.get("TIPO DE COMPRESORES", data.get("TIPO COMPRESORES", "")) or "").upper())

            self._set_combo(self.marca_elem, (data.get("MARCA DE ELEMENTOS", data.get("MARCA ELEMENTOS", "")) or "").upper())
            self._set_combo(self.marca_var,  (data.get("MARCA VARIADORES", data.get("MARCA DE VARIADORES", "")) or "").upper())

            self.n_comp_media.setText(self._only_int(data.get("NUMERO DE COMPRESORES MEDIA", data.get("Nº COMPRESORES MEDIA", ""))))
            self.n_comp_paral.setText(self._only_int(data.get("NUMERO DE COMPRESORES PARALELO", data.get("Nº COMPRESORES PARALELO", ""))))
            self.n_comp_baja.setText(self._only_int(data.get("NUMERO DE COMPRESORES BAJA", data.get("Nº COMPRESORES BAJA", ""))))

        except Exception as e:
            self._warn(f"No pude leer el archivo.\n\nDetalle: {e}")

    # ---------------------------------------------------- Catálogo modelos
    def _get_modelos_r744_por_hoja(self, marca: str) -> List[str]:
        df = self.step2_panel.load_comp_sheet(marca)
        if df is None or df.shape[0] < 3:
            return self._fallback_por_marca(marca)
        colA = df.iloc[2:, 0]
        modelos = [str(v).strip() for v in colA if pd.notna(v) and str(v).strip()]
        seen, result = set(), []
        for m in modelos:
            if m not in seen:
                seen.add(m); result.append(m)
        return result if result else self._fallback_por_marca(marca)

    def _get_modelos_master(self, marca: str, tension: str) -> List[str]:
        book = Path(__file__).resolve().parents[2] / "data" / "basedatos.xlsx"
        try:
            if self._df_master is None:
                df = pd.read_excel(book, sheet_name="COMPRESORES")

                def norm(s: str) -> str:
                    t = (s or "").strip().upper().translate(str.maketrans("ÁÉÍÓÚ", "AEIOU"))
                    return t.replace(" ", "").replace("-", "")

                df.columns = [norm(c) for c in df.columns]
                c_marca = next((c for c in df.columns if c in ("MARCA", "BRAND", "FABRICANTE")), None)
                c_tens = next((c for c in df.columns if c in ("TENSION", "VOLTAGE", "VOLTAJE", "ALIMENTACION", "TENSIONALIMENTACION")), None)
                c_modelo = next((c for c in df.columns if c in ("MODELO", "MODEL", "REFERENCIA", "CODIGO", "NOMBRE")), None)
                if not (c_marca and c_tens and c_modelo):
                    raise ValueError("Faltan columnas esperadas en hoja COMPRESORES.")

                def tens_clean(x) -> str:
                    return "".join(ch for ch in str(x) if ch.isdigit())

                self._df_master = pd.DataFrame({
                    "MARCA": df[c_marca].astype(str).str.upper().str.strip(),
                    "TENSION": df[c_tens].map(tens_clean),
                    "MODELO": df[c_modelo].astype(str).str.strip()
                })
            mm = (marca or "").upper().strip()
            tt = "".join(ch for ch in tension if ch.isdigit())
            subset = self._df_master[(self._df_master["MARCA"] == mm) & (self._df_master["TENSION"] == tt)]
            modelos = subset["MODELO"].dropna().astype(str).str.strip().unique().tolist()
            if modelos:
                return modelos
        except Exception:
            pass
        return self._fallback_por_marca(marca)

    @staticmethod
    def _fallback_por_marca(marca: str) -> List[str]:
        fallback = {
            "BITZER": ["4EES-3Y", "4FES-5Y", "4CES-6Y"],
            "COPELAND": ["ZB15KQE", "ZR61KCE", "ZR94KCE"],
            "DORIN": ["H1500CS", "H2000CC", "K1500CS"],
            "FRASCOLD": ["F20B", "F30B", "Z40A"],
            "TECUMSEH": ["TAJ4513", "AJ5510", "AG5512"],
            "OTRO": ["MODELO-1", "MODELO-2"],
        }
        return fallback.get((marca or "OTRO").upper().strip(), ["MODELO-1", "MODELO-2"])

    # ----------------------------------------------------------- Mensajes
    def _warn(self, text: str) -> None:
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle("Validación")
        box.setText(text)
        box.setStyleSheet("""
            QMessageBox { background-color:#f7f7f7; }
            QLabel { color:#101828; font-size:14px; }
            QPushButton {
                background:#5b8bea; color:#fff; font-weight:700;
                border:none; border-radius:6px; padding:6px 12px;
            }
            QPushButton:hover { background:#6c9cf0; }
        """)
        box.exec()

    def _info(self, text: str) -> None:
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Information)
        box.setWindowTitle("Información")
        box.setText(text)
        box.setStyleSheet("""
            QMessageBox { background-color:#f7f7f7; }
            QLabel { color:#101828; font-size:14px; }
            QPushButton { background:#5b8bea;color:#fff;font-weight:700;border:none;border-radius:6px;padding:6px 12px; }
            QPushButton:hover { background:#6c9cf0; }
        """)
        box.exec()

    # ----------------------------------------------------------- Helpers
    @staticmethod
    def _norm_key(k: str) -> str:
        return (k or "").strip().upper().translate(str.maketrans("ÁÉÍÓÚ", "AEIOU"))

    @staticmethod
    def _only_int(v: str) -> str:
        s = "".join(ch for ch in str(v) if ch.isdigit())
        return s or "0"

    @staticmethod
    def _set_combo(cb: QComboBox, value: str) -> None:
        txt = (value or "").strip()
        if not txt:
            cb.setCurrentIndex(0); return
        i = cb.findText(txt, Qt.MatchFixedString)
        if i >= 0:
            cb.setCurrentIndex(i); return
        for j in range(cb.count()):
            if txt.lower() == cb.itemText(j).lower():
                cb.setCurrentIndex(j); return
        for j in range(cb.count()):
            if txt.lower() in cb.itemText(j).lower():
                cb.setCurrentIndex(j); return
        cb.setCurrentIndex(0)

    @staticmethod
    def _digits(s: str) -> str:
        return "".join(ch for ch in str(s) if ch.isdigit())

    @staticmethod
    def _get_any(d: Dict[str, Any], names: List[str]) -> str:
        if not isinstance(d, dict):
            return ""
        def norm(x: str) -> str:
            return (x or "").strip().upper().translate(str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN"))
        for name in names:
            for k, v in d.items():
                if norm(k) == norm(name):
                    return "" if v is None else str(v)
        return ""

    def _globals_for_step4(self) -> Dict[str, object]:
        return {
            "nombre_proyecto": self.proj.text().strip(),
            "ciudad": self.city.text().strip(),
            "marca_elem": self.marca_elem.currentText().strip(),
            "norma_ap": self.norma_ap.currentText().strip(),
            "t_ctl": self.t_ctl.currentText().strip(),
            "t_alim": self.t_alim.currentText().strip(),
            "refrigerante": self.refrig.currentText().strip(),
            "responsable": self.resp.currentText().strip(),
            "marca_variadores": self.marca_var.currentText().strip(),
            "tipo_compresores": self.tipo_comp.currentText().strip(),
            "step3_state": self.step3_panel.export_state() if hasattr(self, "step3_panel") else {},
        }

    # ----- Aplicar snapshot de Step2 tras reconstruir la grilla -----
    def _apply_loaded_step2(self, st2: Dict[str, Dict[str, str]]) -> None:
        """
        Aplica modelo/arranque/corriente + variadores (modelo y selección).
        Tolera claves 'G1'/'G 1' y mayúsc./minúsc.; acepta muchos alias.
        """
        try:
            # Si Step2Panel trae import_state nativo, úsalo.
            if hasattr(self.step2_panel, "import_state"):
                self.step2_panel.import_state(st2)
                self.step2_panel.refresh_all()
                return

            def canon(s: str) -> str:
                return (s or "").upper().replace(" ", "").replace("-", "")

            # Normalizamos claves guardadas: {"G1": {...}, "B2": {...}}
            saved = {canon(k): (v or {}) for k, v in (st2 or {}).items()}

            keys = self.step2_panel.get_group_keys() if hasattr(self.step2_panel, "get_group_keys") else []
            for k in keys:
                rec = saved.get(canon(k), {}) or {}

                def any_of(names: List[str]) -> str:
                    # Busca por alias dentro de 'rec'
                    def _norm(x: str) -> str:
                        return (x or "").strip().upper().translate(str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN"))
                    idx = {_norm(a): a for a in rec.keys()}
                    for nm in names:
                        a = idx.get(_norm(nm))
                        if a is not None:
                            val = rec[a]
                            return "" if val is None else str(val)
                    return ""

                modelo = any_of(["modelo","modelo_compresor","model","compressor model"])
                arranq = any_of(["arranque","tipo_arranque","start","tipo de arranque"])
                corr   = any_of(["corriente","amps","amperaje","i","icond","icorriente"])

                # Variadores: modelos y selección
                v1 = any_of(["variador1","variador 1","v1","modelo_v1","variador1_modelo","ref variador 1","vfd1"])
                v2 = any_of(["variador2","variador 2","v2","modelo_v2","variador2_modelo","ref variador 2","vfd2"])
                vsel_raw = any_of([
                    "variador_sel","variador seleccionado","seleccion variador","sel variador",
                    "vfd_sel","vsel","seleccion","selección vfd","vfd seleccionado"
                ])

                kwargs: Dict[str, Any] = {}
                if modelo: kwargs["modelo"] = modelo
                if arranq: kwargs["arranque"] = arranq
                if corr:   kwargs["corriente"] = corr
                if v1:     kwargs["variador1"] = v1
                if v2:     kwargs["variador2"] = v2

                # Primer intento: set_fields con todo lo que tengamos
                applied = False
                try:
                    self.step2_panel.set_fields(k, **kwargs)
                    applied = True
                except TypeError:
                    # Mínimo viable
                    try:
                        self.step2_panel.set_fields(
                            k, modelo=kwargs.get("modelo"),
                            arranque=kwargs.get("arranque"),
                            corriente=kwargs.get("corriente")
                        )
                        applied = True
                    except Exception:
                        pass

                # Aplicar selección de variador (1 o 2)
                sel = self._parse_variador_sel(vsel_raw)
                if sel:
                    ok = False
                    try:
                        self.step2_panel.set_fields(k, variador_sel=sel)
                        ok = True
                    except TypeError:
                        pass
                    if not ok and hasattr(self.step2_panel, "set_variador_selection"):
                        try:
                            self.step2_panel.set_variador_selection(k, sel)
                            ok = True
                        except Exception:
                            pass

                if not applied:
                    # Nada grave, seguimos con los demás
                    pass

            self.step2_panel.refresh_all()
        except Exception as e:
            print(f"[apply_loaded_step2] no crítico: {e}")

    @staticmethod
    def _parse_variador_sel(val: str) -> str:
        s = (val or "").strip().upper()
        if s in {"1","V1","VARIADOR 1","UNO","PRIMERO"}: return "1"
        if s in {"2","V2","VARIADOR 2","DOS","SEGUNDO"}: return "2"
        return ""

    # Forzar mayúsculas en QLineEdit
    @staticmethod
    def _force_upper(le: QLineEdit, text: str) -> None:
        up = (text or "").upper()
        if up != text:
            pos = le.cursorPosition()
            le.blockSignals(True)
            le.setText(up)
            le.setCursorPosition(pos)
            le.blockSignals(False)

    # ----- DEV: prefijo del Paso 1 -----
    def _dev_prefill_form(self) -> None:
        self._set_combo(self.t_alim, "220")
        self._set_combo(self.refrig, "R744")
        self._set_combo(self.t_ctl, "220")
        self._set_combo(self.norma_ap, "IEC")
        self._set_combo(self.tipo_comp, "BITZER")
        self.n_comp_paral.setText("2")
        self.n_comp_media.setText("2")
        self.n_comp_baja.setText("2")
        self._set_combo(self.marca_elem, "ABB")
        self._set_combo(self.marca_var, "SCHNEIDER")
