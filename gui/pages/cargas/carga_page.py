from __future__ import annotations

from pathlib import Path
from typing import List, Optional
import json
from datetime import datetime

from PySide6.QtCore import Qt, QEvent
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFrame,
    QLabel,
    QLineEdit,
    QDoubleSpinBox,
    QSpinBox,
    QComboBox,
    QPushButton,
    QScrollArea,
    QGridLayout,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
    QSizePolicy,
    QLayout,
    QFileDialog,
)

from openpyxl import Workbook

from logic.cuartos_frios_engine import ColdRoomEngine, ColdRoomInputs
from gui.pages.cargas.biblioteca_cuartos import BibliotecaCuartosDialog


def _card() -> QFrame:
    f = QFrame()
    f.setStyleSheet("QFrame{background:#fff;border:1px solid #e2e8f5;border-radius:12px;}")
    return f


def _fmt_num(value: float, decimals: int = 2) -> str:
    s = f"{value:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _slug(text: str) -> str:
    import re

    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_") or "proyecto"


class CargaPage(QWidget):
    """
    CÁLCULO DE CARGAS TÉRMICAS (CUARTOS FRÍOS).
    Permite varios cuartos, selección de evaporadores y validación automática.
    """

    def __init__(self) -> None:
        super().__init__()
        data_json = Path(__file__).resolve().parents[3] / "data" / "cuartos_frios" / "cuartos_frios_data.json"
        self.engine = ColdRoomEngine(data_json)
        self._row_clip: dict | None = None
        self._suppress_compute: bool = False
        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        primary_qss = (
            "QPushButton{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #1e3aef, stop:1 #14c9f5);"
            "color:#ffffff;font-weight:700;border:none;border-radius:10px;padding:10px 14px;}"
            "QPushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #2749ff, stop:1 #29d6fa);}" )

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        root.addWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)
        outer = QVBoxLayout(container)
        outer.setContentsMargins(4, 4, 4, 4)
        outer.setSpacing(10)
        outer.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)

        title = QLabel("CÁLCULO DE CARGAS TÉRMICAS (CUARTOS FRÍOS)")
        title.setStyleSheet("font-size:20px;font-weight:800;color:#0f172a;")
        title.setAlignment(Qt.AlignLeft)
        outer.addWidget(title, alignment=Qt.AlignLeft)

        # --- acciones superiores
        actions = QHBoxLayout()
        actions.setSpacing(8)
        actions.setAlignment(Qt.AlignLeft)
        self.btn_export = QPushButton("EXPORTAR PROYECTO"); self.btn_export.setStyleSheet(primary_qss); self.btn_export.clicked.connect(self._export_project)
        self.btn_library = QPushButton("BIBLIOTECA"); self.btn_library.setStyleSheet(primary_qss); self.btn_library.clicked.connect(self._open_library)
        self.btn_new = QPushButton("NUEVO PROYECTO")
        # estilo pastel rojo (igual a Carga Eléctrica)
        self.btn_new.setStyleSheet(
            "QPushButton{background:#fdecec;color:#c53030;font-weight:700;border:none;border-radius:10px;padding:10px 14px;}"
            "QPushButton:hover{background:#fbd5d5;}"
        )
        self.btn_new.clicked.connect(self._new_project)
        actions.addWidget(self.btn_export)
        actions.addWidget(self.btn_library)
        actions.addStretch(1)
        actions.addWidget(self.btn_new)
        outer.addLayout(actions)

        # --- entrada principal
        card_in = _card()
        cin = QVBoxLayout(card_in)
        cin.setContentsMargins(18, 16, 18, 16)
        cin.setSpacing(10)

        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)

        grid.addWidget(QLabel("PROYECTO:"), 0, 0)
        self.proj_edit = QLineEdit()
        self.proj_edit.setPlaceholderText("NOMBRE DEL PROYECTO (OPCIONAL)")
        self.proj_edit.textChanged.connect(lambda t: self.proj_edit.setText(t.upper()))
        grid.addWidget(self.proj_edit, 0, 1, 1, 3)

        grid.addWidget(QLabel("FACTOR DE SEGURIDAD:"), 0, 4)
        self.sf_spin = QDoubleSpinBox()
        self.sf_spin.setRange(1.0, 1.5)
        self.sf_spin.setDecimals(2)
        self.sf_spin.setSingleStep(0.01)
        self.sf_spin.setValue(float(self.engine.config.get("safety_factor_default", 1.15)))
        self._disable_wheel(self.sf_spin)
        grid.addWidget(self.sf_spin, 0, 5)

        grid.addWidget(QLabel("N° CUARTOS:"), 1, 0)
        self.rooms_spin = QSpinBox()
        self.rooms_spin.setRange(1, 20)
        self.rooms_spin.setValue(1)
        self.pending_rooms = self.rooms_spin.value()
        self.rooms_spin.valueChanged.connect(lambda v: setattr(self, "pending_rooms", v))
        grid.addWidget(self.rooms_spin, 1, 1)

        self.btn_calc = QPushButton("CALCULAR")
        self.btn_calc.setStyleSheet(primary_qss)
        self.btn_calc.clicked.connect(self._on_calcular)
        grid.addWidget(self.btn_calc, 1, 5, alignment=Qt.AlignRight)

        cin.addLayout(grid)
        outer.addWidget(card_in)

        # --- tabla de cuartos
        self.room_table = QTableWidget()
        self.room_table.setColumnCount(7)
        self.room_table.setHorizontalHeaderLabels([
            "CUARTO", "LARGO (m)", "ANCHO (m)", "ALTURA (m)",
            "USO / PERFIL", "N EVAPORADORES", "FAMILIA",
        ])
        self.room_table.verticalHeader().setVisible(False)
        self.room_table.setStyleSheet(
            "QTableWidget{gridline-color:#e2e8f5; selection-background-color:#e0f2fe; selection-color:#0f172a;}"
            "QTableWidget::item{background:#ffffff;}"
            "QHeaderView::section{background:#f8fafc;}"
        )
        header = self.room_table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(QHeaderView.Stretch)
        self.room_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.room_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.room_table.itemChanged.connect(lambda *_: self._compute())
        self.room_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.room_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        outer.addWidget(self.room_table)

        # --- resultados
        self.card_out = _card()
        self.card_out.setVisible(False)
        self.card_out.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        cout = QVBoxLayout(self.card_out)
        cout.setContentsMargins(18, 16, 18, 16)
        cout.setSpacing(10)
        cout.setAlignment(Qt.AlignTop)

        res_title = QLabel("RESULTADOS")
        res_title.setStyleSheet("font-size:16px;font-weight:800;color:#0f172a;")
        cout.addWidget(res_title)

        totals_row = QHBoxLayout()
        totals_row.setSpacing(12)
        box_btu, self.lbl_total_btu = self._kpi("CARGA TOTAL", "BTU/H", wide=True)
        box_kw, self.lbl_total_kw = self._kpi("POTENCIA", "kW", wide=True)
        totals_row.addWidget(box_btu)
        totals_row.addWidget(box_kw)
        cout.addLayout(totals_row)

        self.result_table = QTableWidget()
        self.result_table.setColumnCount(6)
        self.result_table.setHorizontalHeaderLabels([
            "CUARTO", "CARGA (BTU/H)", "POT (kW)", "N EVAPS",
            "EVAPORADOR", "UTIL (%)",
        ])
        self.result_table.verticalHeader().setVisible(False)
        self.result_table.setStyleSheet("QTableWidget{gridline-color:#e2e8f5;} QHeaderView::section{background:#f8fafc;}")
        rheader = self.result_table.horizontalHeader()
        rheader.setStretchLastSection(True)
        rheader.setSectionResizeMode(QHeaderView.Stretch)
        self.result_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.result_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.result_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        cout.addWidget(self.result_table)

        outer.addWidget(self.card_out)

        # construye la tabla inicial
        self._rebuild_rooms(self.rooms_spin.value())

    # ------------------------------------------------------------------ helpers UI
    def _disable_wheel(self, widget):
        def _ef(obj, ev):
            if ev.type() == QEvent.Wheel:
                return True
            return False
        widget.installEventFilter(self)
        widget._wheel_filter = _ef  # keep ref

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.Wheel and isinstance(obj, (QDoubleSpinBox, QSpinBox, QComboBox)):
            return True
        return super().eventFilter(obj, ev)

    def _kpi(self, title: str, unit: str, wide: bool = False) -> tuple[QFrame, QLabel]:
        box = QFrame()
        box.setStyleSheet("QFrame{background:#f8fafc;border:1px solid #e2e8f5;border-radius:10px;}")
        box.setFixedHeight(90)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(4)
        lbl_t = QLabel(title)
        lbl_t.setStyleSheet("font-size:11px;font-weight:700;color:#0f172a;")
        lbl_v = QLabel("--")
        lbl_v.setStyleSheet("font-size:22px;font-weight:800;color:#0f172a;")
        lbl_u = QLabel(unit)
        lbl_u.setStyleSheet("color:#475569;font-size:10px;")
        lay.addWidget(lbl_t)
        lay.addWidget(lbl_v)
        lay.addWidget(lbl_u)
        if wide:
            box.setMinimumWidth(400)
        return box, lbl_v

    # ------------------------------------------------------------------ acciones
    def _on_calcular(self) -> None:
        # Regenera la tabla solo cuando se presiona CALCULAR
        self._rebuild_rooms(self.pending_rooms)
        self._compute()

    def _new_project(self) -> None:
        self.proj_edit.clear()
        self.rooms_spin.setValue(1)
        self.pending_rooms = 1
        self._rebuild_rooms(1)
        self.lbl_total_btu.setText("--")
        self.lbl_total_kw.setText("--")
        self.result_table.setRowCount(1)
        self.card_out.setVisible(True)

    def _export_project(self) -> None:
        inputs = self._collect_inputs()
        if not inputs:
            QMessageBox.warning(self, "Exportar", "No hay datos para exportar. Ingresa cuartos y presiona CALCULAR.")
            return
        stamp = datetime.now().strftime("%Y%m%d")
        proj_name = self.proj_edit.text().strip() or "PROYECTO"
        base_name = f"{stamp}_{_slug(proj_name)}"
        target_dir = QFileDialog.getExistingDirectory(self, "Selecciona carpeta destino", str(Path.cwd()))
        if not target_dir:
            return
        target_dir = Path(target_dir)
        excel_out = target_dir / f"{base_name}.xlsx"
        json_out = target_dir / f"{base_name}.json"

        # JSON
        results = []
        for r in range(self.result_table.rowCount()):
            row = []
            for c in range(self.result_table.columnCount()):
                item = self.result_table.item(r, c)
                row.append(item.text() if item else "")
            results.append(row)
        payload = {
            "proyecto": proj_name,
            "factor_seguridad": self.sf_spin.value(),
            "inputs": inputs,
            "resultados": results,
            "totales": {
                "btu_h": self.lbl_total_btu.text(),
                "kw": self.lbl_total_kw.text(),
            },
        }
        json_out.write_text(json.dumps(payload, indent=2), encoding="utf-8")

        # Excel sencillo
        wb = Workbook()
        ws_in = wb.active
        ws_in.title = "Entradas"
        ws_in.append(["PROYECTO", proj_name])
        ws_in.append(["FACTOR SEGURIDAD", self.sf_spin.value()])
        ws_in.append([])
        ws_in.append(["CUARTO", "LARGO (m)", "ANCHO (m)", "ALTURA (m)", "USO", "N EVAPS", "FAMILIA"])
        for row in inputs:
            ws_in.append(
                [
                    row["cuarto"],
                    row["largo_m"],
                    row["ancho_m"],
                    row["altura_m"],
                    row["uso"],
                    row["n_ev_txt"],
                    row["familia"],
                ]
            )
        ws_out = wb.create_sheet("Resultados")
        ws_out.append(["TOTAL BTU/H", self.lbl_total_btu.text(), "TOTAL kW", self.lbl_total_kw.text()])
        ws_out.append([])
        headers = ["CUARTO", "CARGA (BTU/H)", "POT (kW)", "N EVAPS", "EVAPORADOR", "UTIL (%)"]
        ws_out.append(headers)
        for r in range(self.result_table.rowCount()):
            row = []
            for c in range(self.result_table.columnCount()):
                item = self.result_table.item(r, c)
                row.append(item.text() if item else "")
            ws_out.append(row)
        wb.save(excel_out)

        # copia interna en data/proyectos/cuartos
        lib_dir = Path(__file__).resolve().parents[3] / "data" / "proyectos" / "cuartos"
        lib_dir.mkdir(parents=True, exist_ok=True)
        lib_json = lib_dir / json_out.name
        lib_json.write_text(json_out.read_text(encoding="utf-8"), encoding="utf-8")

        QMessageBox.information(
            self,
            "Exportado",
            f"Archivos guardados:\n{excel_out}\n{json_out}\nCopia interna: {lib_json}",
        )

    def _open_library(self) -> None:
        lib_dir = Path(__file__).resolve().parents[3] / "data" / "proyectos" / "cuartos"
        dlg = BibliotecaCuartosDialog(
            parent=self,
            library_dir=lib_dir,
            on_load=self._load_project_file,
            primary_qss="QPushButton{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #1e3aef, stop:1 #14c9f5);color:#fff;font-weight:700;border:none;border-radius:8px;padding:8px 12px;} QPushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #2749ff, stop:1 #29d6fa);}",
            danger_qss="QPushButton{background:#f87171;color:#fff;font-weight:700;border:none;border-radius:8px;padding:8px 12px;} QPushButton:hover{background:#fb7185;}",
        )
        dlg.exec()

    def _load_project_file(self, path: Path) -> None:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        self.proj_edit.setText(data.get("proyecto", ""))
        self.sf_spin.setValue(float(data.get("factor_seguridad", self.sf_spin.value())))
        inputs = data.get("inputs", [])
        n_rooms = max(len(inputs), 1)
        self.pending_rooms = n_rooms
        self.rooms_spin.setValue(n_rooms)
        self._rebuild_rooms(n_rooms)
        # cargar valores
        for r, row in enumerate(inputs):
            self.room_table.cellWidget(r, 1).setValue(float(row.get("largo_m", 0)))
            self.room_table.cellWidget(r, 2).setValue(float(row.get("ancho_m", 0)))
            self.room_table.cellWidget(r, 3).setValue(float(row.get("altura_m", 0)))
            self.room_table.cellWidget(r, 4).setCurrentText(row.get("uso", ""))
            self.room_table.cellWidget(r, 5).setCurrentText(row.get("n_ev_txt", "AUTO (1-4)"))
            self.room_table.cellWidget(r, 6).setCurrentText(row.get("familia", "AUTO"))
        self._compute()

    # ------------------------------------------------------------------ rooms table
    def _rebuild_rooms(self, count: int) -> None:
        # conservar valores existentes
        snapshots = [self._get_row_snapshot(r) for r in range(self.room_table.rowCount())]
        self._suppress_compute = True
        self.room_table.blockSignals(True)
        self.room_table.setRowCount(count)
        usages = list(self.engine.usage_profiles.keys())
        families = ["AUTO", "FRONTAL BAJA", "FRONTAL MEDIA", "DUAL"]
        for r in range(count):
            # cuarto label
            item = QTableWidgetItem(f"CUARTO {r+1}")
            item.setFlags(Qt.ItemIsEnabled)
            self.room_table.setItem(r, 0, item)
            # largo, ancho, altura
            for c in (1, 2, 3):
                spin = QDoubleSpinBox()
                spin.setRange(0.0, 15.0)
                spin.setDecimals(2)
                spin.setSingleStep(0.1)
                spin.setValue(0.0)
                spin.setStyleSheet("QDoubleSpinBox{background:#ffffff;}")
                max_val = 12.8 if c in (1, 2) else 3.7  # ~42 ft, altura ~12 ft (3.66m)
                label = "LARGO" if c == 1 else "ANCHO" if c == 2 else "ALTURA"
                spin.valueChanged.connect(lambda _, s=spin, m=max_val, l=label: self._spin_guard(s, m, l))
                spin.valueChanged.connect(lambda *_: self._compute())
                self._disable_wheel(spin)
                spin.setToolTip(f"Máximo permitido: {max_val:.2f} m \n Si excede use CUADRO DE CARGAS CUARTOS INDUSTRIALES.")
                self.room_table.setCellWidget(r, c, spin)
            # uso
            cbo = QComboBox(); cbo.addItem(""); cbo.addItems(usages); cbo.currentIndexChanged.connect(lambda *_: self._compute()); self._disable_wheel(cbo)
            self.room_table.setCellWidget(r, 4, cbo)
            # n evap
            cbo_ev = QComboBox(); cbo_ev.addItems(["AUTO (1-4)", "1", "2", "3", "4"]); cbo_ev.currentIndexChanged.connect(lambda *_: self._compute()); self._disable_wheel(cbo_ev)
            self.room_table.setCellWidget(r, 5, cbo_ev)
            # familia
            cbo_fam = QComboBox(); cbo_fam.addItems(families); cbo_fam.currentIndexChanged.connect(lambda *_: self._compute()); self._disable_wheel(cbo_fam)
            self.room_table.setCellWidget(r, 6, cbo_fam)
            # restablecer valores si existían
            if r < len(snapshots):
                snap = snapshots[r]
                self.room_table.cellWidget(r, 1).setValue(snap.get("largo", 0.0))
                self.room_table.cellWidget(r, 2).setValue(snap.get("ancho", 0.0))
                self.room_table.cellWidget(r, 3).setValue(snap.get("altura", 0.0))
                self.room_table.cellWidget(r, 4).setCurrentText(snap.get("uso", ""))
                self.room_table.cellWidget(r, 5).setCurrentText(snap.get("nev", "AUTO (1-4)"))
                self.room_table.cellWidget(r, 6).setCurrentText(snap.get("fam", "AUTO"))
        self.room_table.blockSignals(False)
        self._suppress_compute = False

    # ------------------------------------------------------------------ compute
    def _collect_row(self, r: int) -> Optional[ColdRoomInputs]:
        length = self.room_table.cellWidget(r, 1).value()
        width = self.room_table.cellWidget(r, 2).value()
        height = self.room_table.cellWidget(r, 3).value()
        use = self.room_table.cellWidget(r, 4).currentText().strip()
        n_ev_txt = self.room_table.cellWidget(r, 5).currentText()
        fam_txt = self.room_table.cellWidget(r, 6).currentText()

        if length <= 0 or width <= 0 or height <= 0 or not use:
            return None

        max_m = 12.8  # aprox 42 ft
        if any(v > max_m for v in (length, width, height)):
            QMessageBox.warning(self, "Dimensiones fuera de rango", "USE CUADRO DE CARGAS CUARTOS INDUSTRIALES.")
            return None

        n_ev = None if "AUTO" in n_ev_txt.upper() else int(n_ev_txt)
        fam_map = {
            "AUTO": "auto",
            "FRONTAL BAJA": "frontal_wef",
            "FRONTAL MEDIA": "frontal_wefm",
            "DUAL": "dual",
        }
        fam = fam_map.get(fam_txt.upper(), "auto")

        return ColdRoomInputs(
            length_m=length,
            width_m=width,
            height_m=height,
            usage=use,
            n_evaporators=n_ev,
            safety_factor=self.sf_spin.value(),
            family_override=fam,
        )

    def _collect_inputs(self) -> List[dict]:
        data = []
        for r in range(self.room_table.rowCount()):
            length = self.room_table.cellWidget(r, 1).value()
            width = self.room_table.cellWidget(r, 2).value()
            height = self.room_table.cellWidget(r, 3).value()
            use = self.room_table.cellWidget(r, 4).currentText().strip()
            n_ev_txt = self.room_table.cellWidget(r, 5).currentText()
            fam_txt = self.room_table.cellWidget(r, 6).currentText()
            if length <= 0 or width <= 0 or height <= 0 or not use:
                continue
            data.append(
                {
                    "cuarto": r + 1,
                    "largo_m": length,
                    "ancho_m": width,
                    "altura_m": height,
                    "uso": use,
                    "n_ev_txt": n_ev_txt,
                    "familia": fam_txt,
                }
            )
        return data

    # ------------------------------------------------------------------ copy/paste filas
    def _get_row_snapshot(self, r: int) -> dict:
        return {
            "largo": self.room_table.cellWidget(r, 1).value(),
            "ancho": self.room_table.cellWidget(r, 2).value(),
            "altura": self.room_table.cellWidget(r, 3).value(),
            "uso": self.room_table.cellWidget(r, 4).currentText(),
            "nev": self.room_table.cellWidget(r, 5).currentText(),
            "fam": self.room_table.cellWidget(r, 6).currentText(),
        }

    def _apply_row_snapshot(self, r: int, snap: dict) -> None:
        self.room_table.cellWidget(r, 1).setValue(snap.get("largo", 0.0))
        self.room_table.cellWidget(r, 2).setValue(snap.get("ancho", 0.0))
        self.room_table.cellWidget(r, 3).setValue(snap.get("altura", 0.0))
        self.room_table.cellWidget(r, 4).setCurrentText(snap.get("uso", ""))
        self.room_table.cellWidget(r, 5).setCurrentText(snap.get("nev", "AUTO (1-4)"))
        self.room_table.cellWidget(r, 6).setCurrentText(snap.get("fam", "AUTO"))
        self._compute()

    def _compute(self) -> None:
        if self._suppress_compute:
            return
        results: List = []
        total_btu = 0.0
        total_kw = 0.0
        for r in range(self.room_table.rowCount()):
            inp = self._collect_row(r)
            if not inp:
                results.append(None)
                continue
            try:
                res = self.engine.compute(inp)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error calculando cuarto {r+1}: {e}")
                results.append(None)
                continue
            results.append(res)
            if res.valid and res.load_btu_hr:
                total_btu += res.load_btu_hr
                total_kw += res.load_btu_hr * 0.000293071

        if total_btu > 0:
            self.lbl_total_btu.setText(_fmt_num(total_btu, 2))
            self.lbl_total_kw.setText(_fmt_num(total_kw, 2))
            self.card_out.setVisible(True)
        else:
            self.lbl_total_btu.setText("--")
            self.lbl_total_kw.setText("--")
            self.card_out.setVisible(True)

        self.result_table.setRowCount(len(results))
        for r, res in enumerate(results):
            room_name = f"CUARTO {r+1}"
            usage = self.room_table.cellWidget(r, 4).currentText().strip()
            self.result_table.setItem(r, 0, QTableWidgetItem(f"{room_name} - {usage}" if usage else room_name))
            if not res or not res.valid or not res.load_btu_hr:
                for c in range(1, 6):
                    self.result_table.setItem(r, c, QTableWidgetItem("--"))
                continue
            load = res.load_btu_hr
            kw = load * 0.000293071
            n_ev = res.n_used or res.n_requested or (res.n_requested if res.n_requested else None) or 1
            evap = res.evap_model or "--"
            util = ""
            if res.load_per_evap_btu_hr and res.evap_capacity_btu_hr:
                util_pct = res.load_per_evap_btu_hr / res.evap_capacity_btu_hr * 100
                util = _fmt_num(util_pct, 1)

            self.result_table.setItem(r, 1, QTableWidgetItem(_fmt_num(load, 0)))
            self.result_table.setItem(r, 2, QTableWidgetItem(_fmt_num(kw, 2)))
            self.result_table.setItem(r, 3, QTableWidgetItem(str(n_ev)))
            self.result_table.setItem(r, 4, QTableWidgetItem(evap))
            self.result_table.setItem(r, 5, QTableWidgetItem(util or "--"))

        self.result_table.resizeRowsToContents()
        self._fit_table_height(self.room_table)
        self._fit_table_height(self.result_table)

    # ------------------------------------------------------------------
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_C and event.modifiers() & Qt.ControlModifier:
            idx = self._selected_room_row()
            if idx >= 0:
                self._row_clip = self._get_row_snapshot(idx)
            event.accept()
            return
        if event.key() == Qt.Key_V and event.modifiers() & Qt.ControlModifier:
            if self._row_clip is not None:
                idx = self._selected_room_row()
                if idx >= 0:
                    self._apply_row_snapshot(idx, self._row_clip)
            event.accept()
            return
        super().keyPressEvent(event)

    # ------------------------------------------------------------------
    def _spin_guard(self, spin: QDoubleSpinBox, max_val: float, label: str) -> None:
        v = spin.value()
        if v > max_val:
            QMessageBox.warning(
                self,
                "Dimensión fuera de rango",
                f"{label} máximo {max_val:.2f} m.\nUSE CUADRO DE CARGAS CUARTOS INDUSTRIALES.",
            )
            spin.setValue(max_val)

    def _fit_table_height(self, table: QTableWidget) -> None:
        table.resizeRowsToContents()
        header_h = table.horizontalHeader().height() if table.horizontalHeader() else 0
        rows = table.rowCount()
        if rows == 0:
            total_h = header_h + 6
        else:
            row_h = table.rowHeight(0)
            total_h = header_h + rows * row_h + 6
        table.setMinimumHeight(total_h)
        table.setMaximumHeight(total_h)

    def _selected_room_row(self) -> int:
        sel = self.room_table.selectionModel()
        if sel:
            rows = sel.selectedRows()
            if rows:
                return rows[0].row()
        return self.room_table.currentRow()
