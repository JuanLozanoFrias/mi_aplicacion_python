from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Dict, List
import traceback

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QMessageBox,
    QHBoxLayout,
    QSizePolicy,
    QTableView,
    QAbstractItemView,
    QStyledItemDelegate,
    QComboBox,
    QLineEdit,
    QGridLayout,
    QFileDialog,
    QSpinBox,
    QInputDialog,
    QScrollArea,
)

try:
    from logic.legend_jd import LegendJDService
    from logic.legend_jd.project_service import LegendProjectService
except Exception as exc:  # pragma: no cover
    LegendJDService = None  # type: ignore
    LegendProjectService = None  # type: ignore
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None


PROJ_COLUMNS = [
    "loop",
    "ramal",
    "dim_m",
    "equipo",
    "uso",
    "carga_btu_h",
    "tevap_f",
    "evap_qty",
    "evap_modelo",
    "control",
    "succion",
    "liquida",
    "direccion",
    "deshielo",
]


class LegendPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.service = LegendJDService() if LegendJDService else None
        self.project_dir: Path | None = None
        self.project_data: Dict[str, Any] = {}
        self._dirty = False
        self._suspend_updates = False
        self.total_bt = 0.0
        self.total_mt = 0.0
        self._equipos_catalog: List[str] = []
        self._usos_bt: List[str] = []
        self._usos_mt: List[str] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(12, 12, 12, 12)
        outer.setSpacing(6)

        title = QLabel("LEGEND - PROYECTO (EDITABLE)")
        title.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        outer.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        outer.addWidget(scroll, 1)

        content = QWidget()
        content.setMinimumWidth(0)
        content.setSizePolicy(QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred))
        scroll.setWidget(content)

        proj_outer = QVBoxLayout(content)
        proj_outer.setContentsMargins(4, 4, 4, 4)
        proj_outer.setSpacing(6)

        self.btn_open_proj = QPushButton("ABRIR PROYECTO...")
        self.btn_new_proj = QPushButton("NUEVO PROYECTO")
        self.btn_save_proj = QPushButton("GUARDAR")
        self.btn_save_proj.setEnabled(False)
        self.lbl_proj_folder = QLabel("CARPETA: --")
        self.lbl_proj_folder.setToolTip("")
        self.lbl_total_general = QLabel("TOTAL GENERAL: 0.0")
        proj_hdr = QHBoxLayout()
        proj_hdr.addWidget(self.btn_open_proj)
        proj_hdr.addWidget(self.btn_new_proj)
        proj_hdr.addWidget(self.btn_save_proj)
        proj_hdr.addWidget(self.lbl_proj_folder, 1)
        proj_hdr.addWidget(self.lbl_total_general)
        proj_outer.addLayout(proj_hdr)

        # Specs form
        self.spec_fields: Dict[str, QWidget] = {}
        specs_layout = QGridLayout()
        specs_layout.setHorizontalSpacing(8)
        specs_layout.setVerticalSpacing(4)
        label_width = 220
        specs_layout.setColumnMinimumWidth(0, label_width)
        specs_layout.setColumnMinimumWidth(2, label_width)
        specs_layout.setColumnStretch(1, 1)
        specs_layout.setColumnStretch(3, 1)

        row_idx = 0

        def _label(text: str) -> QLabel:
            lbl = QLabel(text)
            lbl.setWordWrap(False)
            lbl.setFixedWidth(label_width)
            lbl.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
            return lbl

        def _expand(widget: QWidget) -> None:
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        def _add_single(label_text: str, widget: QWidget) -> None:
            nonlocal row_idx
            _expand(widget)
            specs_layout.addWidget(_label(label_text), row_idx, 0)
            specs_layout.addWidget(widget, row_idx, 1, 1, 3)
            row_idx += 1

        def _add_double(label1: str, w1: QWidget, label2: str, w2: QWidget) -> None:
            nonlocal row_idx
            _expand(w1)
            _expand(w2)
            specs_layout.addWidget(_label(label1), row_idx, 0)
            specs_layout.addWidget(w1, row_idx, 1)
            lbl2 = _label(label2)
            specs_layout.addWidget(lbl2, row_idx, 2)
            specs_layout.addWidget(w2, row_idx, 3)
            row_idx += 1

        # PROYECTO
        proj_edit = QLineEdit()
        proj_edit.textEdited.connect(lambda txt, w=proj_edit: self._on_spec_changed_upper("proyecto", w, txt))
        self.spec_fields["proyecto"] = proj_edit
        _add_single("PROYECTO", proj_edit)

        # CIUDAD
        ciudad_edit = QLineEdit()
        ciudad_edit.textEdited.connect(lambda txt, w=ciudad_edit: self._on_spec_changed_upper("ciudad", w, txt))
        self.spec_fields["ciudad"] = ciudad_edit
        _add_single("CIUDAD", ciudad_edit)

        # TIPO DE SISTEMA
        tipo_cb = QComboBox()
        tipo_cb.addItems(["", "RACK", "WATERLOOP"])
        tipo_cb.setEditable(False)
        tipo_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("tipo_sistema"))
        self.spec_fields["tipo_sistema"] = tipo_cb
        _add_single("TIPO DE SISTEMA", tipo_cb)
        # TCOND F/C
        self.tcond_f_edit = QLineEdit()
        self.tcond_c_edit = QLineEdit()
        self.tcond_f_edit.setPlaceholderText("F")
        self.tcond_c_edit.setPlaceholderText("C")
        self.tcond_f_edit.textEdited.connect(lambda txt: self._on_temp_changed("f", txt))
        self.tcond_c_edit.textEdited.connect(lambda txt: self._on_temp_changed("c", txt))
        self.spec_fields["tcond_f"] = self.tcond_f_edit
        self.spec_fields["tcond_c"] = self.tcond_c_edit
        _add_double("TCOND (F)", self.tcond_f_edit, "TCOND (C)", self.tcond_c_edit)

        # Voltajes en la misma fila
        self.voltaje_principal_cb = QComboBox()
        self.voltaje_principal_cb.addItems(["", "460", "400", "220", "120"])
        self.voltaje_principal_cb.setEditable(False)
        self.voltaje_principal_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("voltaje_principal"))
        self.voltaje_control_cb = QComboBox()
        self.voltaje_control_cb.addItems(["", "220", "120", "24"])
        self.voltaje_control_cb.setEditable(False)
        self.voltaje_control_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("voltaje_control"))
        self.spec_fields["voltaje_principal"] = self.voltaje_principal_cb
        self.spec_fields["voltaje_control"] = self.voltaje_control_cb
        _add_double("VOLTAJE PRINCIPAL (V AC)", self.voltaje_principal_cb, "VOLTAJE CONTROL (V AC)", self.voltaje_control_cb)

        # Refrigerante + Controlador
        self.refrigerante_cb = QComboBox()
        self.refrigerante_cb.addItems(["", "R22", "R404", "R449", "R507N", "R290", "R744"])
        self.refrigerante_cb.setEditable(False)
        self.controlador_cb = QComboBox()
        self.controlador_cb.addItems(["", "AKSM 800", "AKPC 782", "URACK", "PRACK", "CPC 300", "COPELAND 1015"])
        self.controlador_cb.setEditable(False)
        self.refrigerante_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("refrigerante"))
        self.controlador_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("controlador"))
        self.spec_fields["refrigerante"] = self.refrigerante_cb
        self.spec_fields["controlador"] = self.controlador_cb
        _add_double("REFRIGERANTE", self.refrigerante_cb, "CONTROLADOR", self.controlador_cb)

        specs_widget = QWidget()
        specs_widget.setLayout(specs_layout)
        proj_outer.addWidget(QLabel("ESPECIFICACIONES TECNICAS"))
        proj_outer.addWidget(specs_widget)

        self.btn_open_proj.clicked.connect(self._on_open_project)
        self.btn_new_proj.clicked.connect(self._on_new_project)
        self.btn_save_proj.clicked.connect(self._on_save_project)

        self.bt_model = LegendItemsTableModel([])
        self.mt_model = LegendItemsTableModel([])

        self.bt_view = QTableView()
        self.bt_view.setModel(self.bt_model)
        self.bt_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.bt_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.bt_view.horizontalHeader().setStretchLastSection(True)
        self.bt_view.verticalHeader().setVisible(False)
        self.bt_model.modelReset.connect(self._recalc_totals)
        self.bt_model.dataChanged.connect(self._mark_dirty_and_totals)
        self.bt_model.rowsInserted.connect(self._mark_dirty_and_totals)
        self.bt_model.rowsRemoved.connect(self._mark_dirty_and_totals)
        self.bt_view.setItemDelegateForColumn(PROJ_COLUMNS.index("equipo"), ComboDelegate(lambda: self._equipos_catalog))
        self.bt_view.setItemDelegateForColumn(PROJ_COLUMNS.index("uso"), ComboDelegate(lambda: self._usos_bt))
        self.bt_view.setItemDelegateForColumn(PROJ_COLUMNS.index("deshielo"), ComboDelegate(lambda: ["", "DESHIELO ELECTRICO", "DESHIELO POR TIEMPO"], editable=False))
        self.bt_view.keyPressEvent = lambda event, v=self.bt_view: self._table_key_press(event, self.bt_model, v)

        self.mt_view = QTableView()
        self.mt_view.setModel(self.mt_model)
        self.mt_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.mt_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.mt_view.horizontalHeader().setStretchLastSection(True)
        self.mt_view.verticalHeader().setVisible(False)
        self.mt_model.modelReset.connect(self._recalc_totals)
        self.mt_model.dataChanged.connect(self._mark_dirty_and_totals)
        self.mt_model.rowsInserted.connect(self._mark_dirty_and_totals)
        self.mt_model.rowsRemoved.connect(self._mark_dirty_and_totals)
        self.mt_view.setItemDelegateForColumn(PROJ_COLUMNS.index("equipo"), ComboDelegate(lambda: self._equipos_catalog))
        self.mt_view.setItemDelegateForColumn(PROJ_COLUMNS.index("uso"), ComboDelegate(lambda: self._usos_mt))
        self.mt_view.setItemDelegateForColumn(PROJ_COLUMNS.index("deshielo"), ComboDelegate(lambda: ["", "DESHIELO ELECTRICO", "DESHIELO POR TIEMPO"], editable=False))
        self.mt_view.keyPressEvent = lambda event, v=self.mt_view: self._table_key_press(event, self.mt_model, v)
        # Ramales y tablas BT
        self.spin_bt = QSpinBox()
        self.spin_bt.setRange(1, 20)
        self.spin_bt.setEnabled(False)
        self.spin_bt_add = QSpinBox()
        self.spin_bt_add.setRange(1, 20)
        self.spin_bt_add.setValue(1)
        self.btn_add_ramales_bt = QPushButton("AGREGAR")
        self.btn_add_ramales_bt.clicked.connect(self._add_ramales_bt_bulk)

        proj_outer.addWidget(QLabel("BAJA (BT)"))
        ram_bt_layout = QHBoxLayout()
        ram_bt_layout.addWidget(QLabel("RAMALES BT:"))
        ram_bt_layout.addWidget(self.spin_bt)
        ram_bt_layout.addSpacing(12)
        ram_bt_layout.addWidget(QLabel("AGREGAR:"))
        ram_bt_layout.addWidget(self.spin_bt_add)
        ram_bt_layout.addWidget(self.btn_add_ramales_bt)
        ram_bt_layout.addStretch(1)
        proj_outer.addLayout(ram_bt_layout)

        btn_bt_row = QHBoxLayout()
        btn_add_bt = QPushButton("AGREGAR FILA")
        btn_del_bt = QPushButton("ELIMINAR FILA")
        btn_dup_bt = QPushButton("DUPLICAR FILA")
        btn_add_bt.clicked.connect(lambda: self._add_row(self.bt_model, self.bt_view))
        btn_del_bt.clicked.connect(lambda: self._del_row(self.bt_view, self.bt_model))
        btn_dup_bt.clicked.connect(lambda: self._dup_row(self.bt_view, self.bt_model))
        for b in (btn_add_bt, btn_del_bt, btn_dup_bt):
            btn_bt_row.addWidget(b)
        btn_bt_row.addStretch(1)
        proj_outer.addLayout(btn_bt_row)
        proj_outer.addWidget(self.bt_view)
        self.lbl_total_bt = QLabel("TOTAL BT: 0.0")
        proj_outer.addWidget(self.lbl_total_bt)

        # Ramales y tablas MT
        self.spin_mt = QSpinBox()
        self.spin_mt.setRange(1, 20)
        self.spin_mt.setEnabled(False)
        btn_ramal_mt_add = QPushButton("+ RAMAL")
        btn_ramal_mt_sub = QPushButton("- RAMAL")
        btn_ramal_mt_add.clicked.connect(lambda: self._add_ramal("mt"))
        btn_ramal_mt_sub.clicked.connect(lambda: self._del_ramal("mt"))

        proj_outer.addWidget(QLabel("MEDIA (MT)"))
        ram_mt_layout = QHBoxLayout()
        ram_mt_layout.addWidget(QLabel("RAMALES MT:"))
        ram_mt_layout.addWidget(self.spin_mt)
        ram_mt_layout.addWidget(btn_ramal_mt_add)
        ram_mt_layout.addWidget(btn_ramal_mt_sub)
        ram_mt_layout.addStretch(1)
        proj_outer.addLayout(ram_mt_layout)

        btn_mt_row = QHBoxLayout()
        btn_add_mt = QPushButton("AGREGAR FILA")
        btn_del_mt = QPushButton("ELIMINAR FILA")
        btn_dup_mt = QPushButton("DUPLICAR FILA")
        btn_dup_ramal_mt = QPushButton("DUPLICAR RAMAL...")
        btn_add_mt.clicked.connect(lambda: self._add_row(self.mt_model, self.mt_view))
        btn_del_mt.clicked.connect(lambda: self._del_row(self.mt_view, self.mt_model))
        btn_dup_mt.clicked.connect(lambda: self._dup_row(self.mt_view, self.mt_model))
        btn_dup_ramal_mt.clicked.connect(lambda: self._duplicate_ramal_dialog("mt"))
        for b in (btn_add_mt, btn_del_mt, btn_dup_mt, btn_dup_ramal_mt):
            btn_mt_row.addWidget(b)
        btn_mt_row.addStretch(1)
        proj_outer.addLayout(btn_mt_row)
        proj_outer.addWidget(self.mt_view)
        self.lbl_total_mt = QLabel("TOTAL MT: 0.0")
        proj_outer.addWidget(self.lbl_total_mt)

        if _IMPORT_ERROR:
            QMessageBox.critical(self, "LEGEND", f"No se pudo importar LegendJDService:\n{_IMPORT_ERROR}")
        else:
            self.refresh()

    def _on_open_project(self) -> None:
        if self._dirty and not self._confirm_discard():
            return
        dir_path = QFileDialog.getExistingDirectory(self, "Selecciona carpeta de proyecto", str(Path.cwd()))
        if not dir_path or not LegendProjectService:
            return
        self._load_project(Path(dir_path))

    def _on_new_project(self) -> None:
        if self._dirty and not self._confirm_discard():
            return
        dir_path = QFileDialog.getExistingDirectory(self, "Selecciona carpeta para nuevo proyecto", str(Path.cwd()))
        if not dir_path or not LegendProjectService:
            return
        self.project_dir = Path(dir_path)
        self.project_data = LegendProjectService.load(self.project_dir)
        self._render_project()
        self.btn_save_proj.setEnabled(True)
        self.lbl_proj_folder.setText(f"CARPETA: {self.project_dir}")
        self.lbl_proj_folder.setToolTip(str(self.project_dir))
        self._set_dirty(True)

    def _on_save_project(self) -> None:
        if not LegendProjectService or not self.project_dir:
            return
        self.project_data["bt_items"] = self.bt_model.items
        self.project_data["mt_items"] = self.mt_model.items
        specs = self.project_data.get("specs", {}) if isinstance(self.project_data, dict) else {}
        for k, widget in self.spec_fields.items():
            if isinstance(widget, QComboBox):
                specs[k] = widget.currentText()
            else:
                specs[k] = widget.text()
        self.project_data["specs"] = specs
        LegendProjectService.save(self.project_dir, self.project_data)
        QMessageBox.information(self, "LEGEND", "GUARDADO OK.")
        self._set_dirty(False)

    def _load_project(self, dir_path: Path) -> None:
        if not LegendProjectService:
            return
        try:
            self.project_dir = dir_path
            self.project_data = LegendProjectService.load(dir_path)
            self.lbl_proj_folder.setText(f"CARPETA: {dir_path}")
            self.lbl_proj_folder.setToolTip(str(dir_path))
            self.btn_save_proj.setEnabled(True)
            self._render_project()
            self._set_dirty(False)
        except Exception as exc:
            QMessageBox.critical(self, "LEGEND", f"No se pudo cargar proyecto:\n{exc}")
    def _on_spec_changed(self, key: str) -> None:
        if not isinstance(self.project_data, dict):
            self.project_data = {}
        specs = self.project_data.get("specs", {}) if isinstance(self.project_data, dict) else {}
        widget = self.spec_fields.get(key)
        if isinstance(widget, QComboBox):
            specs[key] = widget.currentText()
        else:
            specs[key] = widget.text() if widget else ""
        self.project_data["specs"] = specs
        self._set_dirty(True)

    def _on_spec_changed_upper(self, key: str, widget: QLineEdit, text: str) -> None:
        upper = text.upper()
        if upper != text:
            cursor = widget.cursorPosition()
            widget.setText(upper)
            widget.setCursorPosition(cursor)
        self._on_spec_changed(key)

    def _on_temp_changed(self, scale: str, text: str) -> None:
        try:
            val = float(text)
        except Exception:
            val = None
        if scale == "f":
            if val is not None:
                c = (val - 32) * 5.0 / 9.0
                self.tcond_c_edit.blockSignals(True)
                self.tcond_c_edit.setText(f"{c:.2f}")
                self.tcond_c_edit.blockSignals(False)
                self.project_data.setdefault("specs", {})["tcond_c"] = f"{c:.2f}"
            self.project_data.setdefault("specs", {})["tcond_f"] = text
        else:
            if val is not None:
                f = val * 9.0 / 5.0 + 32
                self.tcond_f_edit.blockSignals(True)
                self.tcond_f_edit.setText(f"{f:.2f}")
                self.tcond_f_edit.blockSignals(False)
                self.project_data.setdefault("specs", {})["tcond_f"] = f"{f:.2f}"
            self.project_data.setdefault("specs", {})["tcond_c"] = text
        self._set_dirty(True)

    def _render_project(self) -> None:
        bt_items = self.project_data.get("bt_items", []) if isinstance(self.project_data, dict) else []
        mt_items = self.project_data.get("mt_items", []) if isinstance(self.project_data, dict) else []
        specs = self.project_data.get("specs", {}) if isinstance(self.project_data, dict) else {}
        for k, widget in self.spec_fields.items():
            widget.blockSignals(True)
            val = str(specs.get(k, ""))
            if isinstance(widget, QComboBox):
                idx = widget.findText(val)
                widget.setCurrentIndex(idx if idx >= 0 else 0)
            else:
                widget.setText(val)
            widget.blockSignals(False)
        f_val = specs.get("tcond_f")
        c_val = specs.get("tcond_c")
        if f_val and not c_val:
            try:
                f_num = float(f_val)
                c_calc = (f_num - 32) * 5.0 / 9.0
                self.tcond_c_edit.setText(f"{c_calc:.2f}")
            except Exception:
                pass
        if c_val and not f_val:
            try:
                c_num = float(c_val)
                f_calc = c_num * 9.0 / 5.0 + 32
                self.tcond_f_edit.setText(f"{f_calc:.2f}")
            except Exception:
                pass
        self.spin_bt.blockSignals(True)
        self.spin_mt.blockSignals(True)
        self.spin_bt.setValue(self._to_int(self.project_data.get("bt_ramales", 1)))
        self.spin_mt.setValue(self._to_int(self.project_data.get("mt_ramales", 1)))
        self.spin_bt.setEnabled(True)
        self.spin_mt.setEnabled(True)
        self.spin_bt.blockSignals(False)
        self.spin_mt.blockSignals(False)
        self._update_bt_add_controls()
        self.bt_model.set_items(bt_items)
        self.mt_model.set_items(mt_items)
        self._fit_table_to_contents_view(self.bt_view)
        self._fit_table_to_contents_view(self.mt_view)
        self._recalc_totals()

    def _add_row(self, model, view: QTableView) -> None:
        model.add_row({})
        self._fit_table_to_contents_view(view)
        self._set_dirty(True)

    def _del_row(self, view: QTableView, model) -> None:
        sel = view.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "LEGEND", "SELECCIONA UNA FILA PRIMERO.")
            return
        model.del_row(sel[0].row())
        self._fit_table_to_contents_view(view)
        self._set_dirty(True)

    def _dup_row(self, view: QTableView, model) -> None:
        sel = view.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "LEGEND", "SELECCIONA UNA FILA PRIMERO.")
            return
        model.dup_row(sel[0].row())
        self._fit_table_to_contents_view(view)
        self._set_dirty(True)

    def _fit_table_to_contents_view(self, view: QTableView) -> None:
        view.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        if not hasattr(view, "_legend_fixed_height"):
            view.setMinimumHeight(260)
            view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            setattr(view, "_legend_fixed_height", True)

    def _schedule_after_insert(self, view: QTableView) -> None:
        def _apply() -> None:
            try:
                view.reset()
            except Exception:
                pass
            try:
                view.viewport().update()
            except Exception:
                pass
            try:
                self._fit_table_to_contents_view(view)
            except Exception:
                pass
        from PySide6.QtCore import QTimer
        QTimer.singleShot(0, _apply)

    def _mark_dirty_and_totals(self, *args, **kwargs) -> None:
        if self._suspend_updates:
            return
        self._set_dirty(True)

    def _set_dirty(self, value: bool) -> None:
        self._dirty = value
        self._recalc_totals()

    def _confirm_discard(self) -> bool:
        resp = QMessageBox.question(
            self,
            "LEGEND",
            "HAY CAMBIOS SIN GUARDAR. ¿DESEAS DESCARTARLOS?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        return resp == QMessageBox.Yes

    def _to_int(self, val, default: int = 1) -> int:
        try:
            return int(val)
        except Exception:
            return default

    def _default_row(self, bloque: str) -> dict:
        return dict(self.bt_model.default_row if bloque == "bt" else self.mt_model.default_row)

    def _write_error_log(self, text: str) -> None:
        try:
            log_path = Path.cwd() / "legend_error.log"
            with open(log_path, "a", encoding="utf-8") as fh:
                fh.write(text + "\n")
        except Exception:
            pass

    def _update_bt_add_controls(self) -> None:
        if not hasattr(self, "spin_bt_add"):
            return
        current = self._to_int(self.project_data.get("bt_ramales", 1))
        remaining = max(0, 20 - current)
        if remaining <= 0:
            self.spin_bt_add.setEnabled(False)
            self.btn_add_ramales_bt.setEnabled(False)
            self.spin_bt_add.setMaximum(1)
            self.spin_bt_add.setValue(1)
            return
        self.spin_bt_add.setEnabled(True)
        self.btn_add_ramales_bt.setEnabled(True)
        self.spin_bt_add.setMaximum(remaining)
        if self.spin_bt_add.value() > remaining:
            self.spin_bt_add.setValue(remaining)

    def _add_ramales_bt_bulk(self) -> None:
        try:
            count = self.spin_bt_add.value() if hasattr(self, "spin_bt_add") else 1
            self._add_ramales_bulk("bt", count)
        except Exception:
            msg = traceback.format_exc()
            self._write_error_log(msg)
            QMessageBox.critical(self, "LEGEND", f"ERROR AL AGREGAR RAMALES:\n{msg}")

    def _add_ramales_bulk(self, bloque: str, count: int) -> None:
        try:
            self._write_error_log(f"[ADD_RAMALES] START bloque={bloque} count={count}")
            if count <= 0:
                return
            if not isinstance(self.project_data, dict):
                self.project_data = {}
            key_items = "bt_items" if bloque == "bt" else "mt_items"
            key_ramales = "bt_ramales" if bloque == "bt" else "mt_ramales"
            model = self.bt_model if bloque == "bt" else self.mt_model
            view = self.bt_view if bloque == "bt" else self.mt_view
            items = model.items
            current = self._to_int(self.project_data.get(key_ramales, 1))
            if current >= 20:
                QMessageBox.information(self, "LEGEND", "MAXIMO 20 RAMALES.")
                return
            if current + count > 20:
                count = 20 - current
            n = current
            temp_items = [r for r in items if isinstance(r, dict)]
            new_rows_all: List[dict] = []
            for _ in range(count):
                base_rows = [dict(r) for r in temp_items if self._to_int(r.get("ramal", 1)) == n]
                new_rows = []
                if not base_rows:
                    new_row = self._default_row(bloque)
                    new_row["ramal"] = n + 1
                    new_rows.append(new_row)
                else:
                    for r in base_rows:
                        nr = dict(r)
                        nr["ramal"] = n + 1
                        nr["dim_m"] = ""
                        nr["carga_btu_h"] = 0
                        nr["evap_qty"] = 0
                        nr["evap_modelo"] = ""
                        new_rows.append(nr)
                new_rows_all.extend(new_rows)
                temp_items.extend(new_rows)
                n += 1
            view.setUpdatesEnabled(False)
            model.blockSignals(True)
            self._suspend_updates = True
            model.add_rows(new_rows_all)
            self._suspend_updates = False
            model.blockSignals(False)
            view.setUpdatesEnabled(True)
            self._schedule_after_insert(view)
            self.project_data[key_items] = model.items
            self.project_data[key_ramales] = n
            if bloque == "bt":
                self.spin_bt.setValue(n)
                self._write_error_log("[ADD_RAMALES] BT add_rows OK")
                self._fit_table_to_contents_view(self.bt_view)
                self._update_bt_add_controls()
            else:
                self.spin_mt.setValue(n)
                self._write_error_log("[ADD_RAMALES] MT add_rows OK")
                self._fit_table_to_contents_view(self.mt_view)
            self._set_dirty(True)
            self._write_error_log("[ADD_RAMALES] DONE")
        except Exception:
            msg = traceback.format_exc()
            self._write_error_log(msg)
            QMessageBox.critical(self, "LEGEND", f"ERROR INTERNO:\n{msg}")

    def _add_ramal(self, bloque: str) -> None:
        if not isinstance(self.project_data, dict):
            self.project_data = {}
        key_items = "bt_items" if bloque == "bt" else "mt_items"
        key_ramales = "bt_ramales" if bloque == "bt" else "mt_ramales"
        model = self.bt_model if bloque == "bt" else self.mt_model
        view = self.bt_view if bloque == "bt" else self.mt_view
        items = [r for r in model.items if isinstance(r, dict)]
        n = self._to_int(self.project_data.get(key_ramales, 1))
        n2 = n + 1
        base_rows = [dict(r) for r in items if isinstance(r, dict) and self._to_int(r.get("ramal", 1)) == n]
        new_rows = []
        if not base_rows:
            new_row = self._default_row(bloque)
            new_row["ramal"] = n2
            new_rows.append(new_row)
        else:
            for r in base_rows:
                nr = dict(r)
                nr["ramal"] = n2
                nr["dim_m"] = ""
                nr["carga_btu_h"] = 0
                nr["evap_qty"] = 0
                nr["evap_modelo"] = ""
                new_rows.append(nr)
        view.setUpdatesEnabled(False)
        model.blockSignals(True)
        self._suspend_updates = True
        model.add_rows(new_rows)
        self._suspend_updates = False
        model.blockSignals(False)
        view.setUpdatesEnabled(True)
        self._schedule_after_insert(view)
        self.project_data[key_items] = model.items
        self.project_data[key_ramales] = n2
        if bloque == "bt":
            self.spin_bt.setValue(n2)
            self._fit_table_to_contents_view(self.bt_view)
        else:
            self.spin_mt.setValue(n2)
            self._fit_table_to_contents_view(self.mt_view)
        self._set_dirty(True)

    def _del_ramal(self, bloque: str) -> None:
        if not isinstance(self.project_data, dict):
            return
        key_items = "bt_items" if bloque == "bt" else "mt_items"
        key_ramales = "bt_ramales" if bloque == "bt" else "mt_ramales"
        items = self.project_data.get(key_items, [])
        if not isinstance(items, list):
            items = []
        items = [r for r in items if isinstance(r, dict)]
        n = self._to_int(self.project_data.get(key_ramales, 1))
        if n <= 1:
            return
        resp = QMessageBox.question(
            self,
            "LEGEND",
            f"SE ELIMINARAN TODAS LAS FILAS DEL RAMAL {n}. ¿CONTINUAR?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return
        items = [dict(r) for r in items if isinstance(r, dict) and self._to_int(r.get("ramal", 1)) != n]
        self.project_data[key_items] = items
        self.project_data[key_ramales] = n - 1
        if bloque == "bt":
            self.bt_view.setUpdatesEnabled(False)
            self.bt_model.blockSignals(True)
            self._suspend_updates = True
            self.bt_model.set_items(items)
            self.bt_model.blockSignals(False)
            self.bt_view.setUpdatesEnabled(True)
            self._suspend_updates = False
            self.spin_bt.setValue(n - 1)
            self._fit_table_to_contents_view(self.bt_view)
        else:
            self.mt_view.setUpdatesEnabled(False)
            self.mt_model.blockSignals(True)
            self._suspend_updates = True
            self.mt_model.set_items(items)
            self.mt_model.blockSignals(False)
            self.mt_view.setUpdatesEnabled(True)
            self._suspend_updates = False
            self.spin_mt.setValue(n - 1)
            self._fit_table_to_contents_view(self.mt_view)
        self._set_dirty(True)

    def _recalc_totals(self) -> None:
        def _sum(model: "LegendItemsTableModel") -> float:
            total = 0.0
            for r in model.items:
                try:
                    total += float(r.get("carga_btu_h", 0) or 0)
                except Exception:
                    pass
            return total

        self.total_bt = _sum(self.bt_model)
        self.total_mt = _sum(self.mt_model)
        fmt = lambda v: f"{v:,.1f}"
        self.lbl_total_bt.setText(f"TOTAL BT (BTU/H): {fmt(self.total_bt)}")
        self.lbl_total_mt.setText(f"TOTAL MT (BTU/H): {fmt(self.total_mt)}")
        self.lbl_total_general.setText(f"TOTAL GENERAL: {fmt(self.total_bt + self.total_mt)}")

    def _table_key_press(self, event, model: "LegendItemsTableModel", view: QTableView) -> None:
        if event.matches(event.StandardKey.Copy):
            self._copy_selection(model, view)
            return
        if event.matches(event.StandardKey.Paste):
            self._paste_selection(model, view)
            return
        if event.matches(event.StandardKey.Duplicate):
            self._dup_selection(model, view)
            return
        QTableView.keyPressEvent(view, event)

    def _copy_selection(self, model: "LegendItemsTableModel", view: QTableView) -> None:
        rows = sorted({i.row() for i in view.selectionModel().selectedRows()})
        if not rows:
            return
        lines = []
        for r in rows:
            item = model.items[r]
            vals = [str(item.get(k, "")) for k in model.headers]
            lines.append("\t".join(vals))
        QGuiApplication.clipboard().setText("\n".join(lines))

    def _paste_selection(self, model: "LegendItemsTableModel", view: QTableView) -> None:
        text = QGuiApplication.clipboard().text()
        if not text:
            return
        for line in text.splitlines():
            parts = line.split("\t")
            row = model.default_row.copy()
            for idx, key in enumerate(model.headers):
                if idx < len(parts):
                    row[key] = parts[idx]
            model.add_row(row)
        self._fit_table_to_contents_view(view)
        self._set_dirty(True)

    def _dup_selection(self, model: "LegendItemsTableModel", view: QTableView) -> None:
        rows = sorted({i.row() for i in view.selectionModel().selectedRows()})
        if not rows:
            return
        for r in rows:
            model.dup_row(r)
        self._fit_table_to_contents_view(view)
        self._set_dirty(True)

    def _duplicate_ramal_dialog(self, bloque: str) -> None:
        key_ramales = "bt_ramales" if bloque == "bt" else "mt_ramales"
        max_val = self._to_int(self.project_data.get(key_ramales, 1))
        src, ok = QInputDialog.getInt(self, "Ramal origen", "Ramal origen:", 1, 1, max(max_val, 20))
        if not ok:
            return
        dest, ok = QInputDialog.getInt(self, "Ramal destino", "Ramal destino:", src + 1, 1, 20)
        if not ok:
            return
        self._duplicate_ramal(bloque, src, dest)

    def _duplicate_ramal(self, bloque: str, src: int, dest: int) -> None:
        key_items = "bt_items" if bloque == "bt" else "mt_items"
        key_ramales = "bt_ramales" if bloque == "bt" else "mt_ramales"
        items = self.project_data.get(key_items, [])
        if not isinstance(items, list):
            items = []
        items = [r for r in items if isinstance(r, dict)]
        base_rows = [dict(r) for r in items if isinstance(r, dict) and self._to_int(r.get("ramal", 1)) == src]
        new_rows = []
        for r in base_rows:
            nr = dict(r)
            nr["ramal"] = dest
            nr["dim_m"] = ""
            nr["carga_btu_h"] = 0
            nr["evap_qty"] = 0
            nr["evap_modelo"] = ""
            new_rows.append(nr)
        if not new_rows:
            base = self._default_row(bloque)
            base["ramal"] = dest
            new_rows.append(base)
        items = [r for r in items if isinstance(r, dict) and self._to_int(r.get("ramal", 1)) != dest] + new_rows
        self.project_data[key_items] = items
        current = self._to_int(self.project_data.get(key_ramales, 1))
        self.project_data[key_ramales] = max(current, dest)
        if bloque == "bt":
            self.bt_view.setUpdatesEnabled(False)
            self.bt_model.blockSignals(True)
            self._suspend_updates = True
            self.bt_model.set_items(items)
            self.bt_model.blockSignals(False)
            self.bt_view.setUpdatesEnabled(True)
            self._suspend_updates = False
            self.spin_bt.setValue(self.project_data[key_ramales])
            self._fit_table_to_contents_view(self.bt_view)
        else:
            self.mt_view.setUpdatesEnabled(False)
            self.mt_model.blockSignals(True)
            self._suspend_updates = True
            self.mt_model.set_items(items)
            self.mt_model.blockSignals(False)
            self.mt_view.setUpdatesEnabled(True)
            self._suspend_updates = False
            self.spin_mt.setValue(self.project_data[key_ramales])
            self._fit_table_to_contents_view(self.mt_view)
        self._set_dirty(True)

    def refresh(self) -> None:
        if not self.service:
            return
        try:
            data = self.service.load_all()
        except Exception:
            return
        usos = data.get("usos", {}) if isinstance(data, dict) else {}
        equipos = data.get("equipos", []) if isinstance(data, dict) else []
        self._equipos_catalog = [getattr(e, "equipo", "") for e in equipos]
        self._usos_bt = usos.get("BT", []) if isinstance(usos, dict) else []
        self._usos_mt = usos.get("MT", []) if isinstance(usos, dict) else []


class ComboDelegate(QStyledItemDelegate):
    def __init__(self, items_fn: Callable[[], List[str]], editable: bool = True, parent=None):
        super().__init__(parent)
        self.items_fn = items_fn
        self.editable = editable

    def createEditor(self, parent, option, index):
        items = self.items_fn() or []
        if not items:
            return super().createEditor(parent, option, index)
        cb = QComboBox(parent)
        cb.setEditable(self.editable)
        cb.addItems(items)
        cb.setInsertPolicy(QComboBox.NoInsert)
        return cb

    def setEditorData(self, editor, index):
        if isinstance(editor, QComboBox):
            val = str(index.data(Qt.EditRole) or "")
            idx = editor.findText(val)
            if idx >= 0:
                editor.setCurrentIndex(idx)
            else:
                if self.editable:
                    editor.setEditText(val)
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentText(), Qt.EditRole)
            return
        super().setModelData(editor, model, index)


class LegendItemsTableModel(QAbstractTableModel):
    headers = PROJ_COLUMNS

    def __init__(self, items: List[dict] | None = None):
        super().__init__()
        self.items: List[dict] = items or []
        self.default_row = {
            "loop": 1,
            "ramal": 1,
            "dim_m": "",
            "equipo": "",
            "uso": "",
            "carga_btu_h": 0.0,
            "tevap_f": 0.0,
            "evap_qty": 0,
            "evap_modelo": "",
            "control": "",
            "succion": "",
            "liquida": "",
            "direccion": "",
            "deshielo": "",
        }

    def rowCount(self, parent=QModelIndex()):
        return len(self.items)

    def columnCount(self, parent=QModelIndex()):
        return len(self.headers)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        item = self.items[index.row()]
        key = self.headers[index.column()]
        val = item.get(key, "")
        if role in (Qt.DisplayRole, Qt.EditRole):
            return val
        if role == Qt.TextAlignmentRole:
            if key in ("loop", "ramal", "dim_m", "carga_btu_h", "tevap_f", "evap_qty"):
                return Qt.AlignRight | Qt.AlignVCenter
            return Qt.AlignLeft | Qt.AlignVCenter
        return None

    def setData(self, index: QModelIndex, value, role=Qt.EditRole):
        if role == Qt.EditRole and index.isValid():
            key = self.headers[index.column()]
            self.items[index.row()][key] = value
            self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
            return True
        return False

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsEnabled
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def headerData(self, section: int, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.headers[section].upper()
        return None

    def add_row(self, row: dict) -> None:
        self.beginInsertRows(QModelIndex(), len(self.items), len(self.items))
        if not row:
            row = dict(self.default_row)
        else:
            norm = dict(self.default_row)
            norm.update(row)
            row = norm
        self.items.append(row)
        self.endInsertRows()

    def add_rows(self, rows: List[dict]) -> None:
        if not rows:
            return
        start = len(self.items)
        end = start + len(rows) - 1
        self.beginInsertRows(QModelIndex(), start, end)
        for row in rows:
            if not row:
                row = dict(self.default_row)
            else:
                norm = dict(self.default_row)
                norm.update(row)
                row = norm
            self.items.append(row)
        self.endInsertRows()
        try:
            top = self.index(start, 0)
            bottom = self.index(end, self.columnCount() - 1)
            self.dataChanged.emit(top, bottom, [Qt.DisplayRole, Qt.EditRole])
            self.layoutChanged.emit()
        except Exception:
            pass

    def del_row(self, idx: int) -> None:
        if 0 <= idx < len(self.items):
            self.beginRemoveRows(QModelIndex(), idx, idx)
            self.items.pop(idx)
            self.endRemoveRows()

    def dup_row(self, idx: int) -> None:
        if 0 <= idx < len(self.items):
            row = dict(self.items[idx])
            self.add_row(row)

    def set_items(self, items: List[dict]) -> None:
        self.beginResetModel()
        if not isinstance(items, list):
            items = []
        self.items = [r for r in items if isinstance(r, dict)]
        self.endResetModel()


if __name__ == "__main__":
    from PySide6.QtWidgets import QApplication
    import sys

    app = QApplication(sys.argv)
    w = LegendPage()
    w.resize(1000, 700)
    w.show()
    sys.exit(app.exec())
