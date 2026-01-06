from __future__ import annotations

from pathlib import Path
import json
from typing import Any, Callable, Dict, List
import traceback

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer
from PySide6.QtGui import QGuiApplication, QKeySequence, QShortcut, QIntValidator, QColor
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
    QHeaderView,
    QMenu,
    QDialog,
    QTableWidget,
    QTableWidgetItem,
)

try:
    from logic.legend_jd import LegendJDService
    from logic.legend_jd.project_service import LegendProjectService
    from logic.legend_jd.legend_exporter import build_legend_workbook
except Exception as exc:  # pragma: no cover
    LegendJDService = None  # type: ignore
    LegendProjectService = None  # type: ignore
    build_legend_workbook = None  # type: ignore
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None

try:
    from logic.cuartos_frios_engine import ColdRoomEngine, ColdRoomInputs
except Exception:
    ColdRoomEngine = None  # type: ignore
    ColdRoomInputs = None  # type: ignore


PROJ_COLUMNS = [
    "loop",
    "ramal",
    "equipo",
    "uso",
    "largo_m",
    "ancho_m",
    "alto_m",
    "dim_ft",
    "btu_ft",
    "btu_hr",
    "evap_modelo",
    "evap_qty",
    "familia",
]


class LegendPreviewDialog(QDialog):
    def __init__(self, wb, parent=None):
        super().__init__(parent)
        self.setWindowTitle("VISTA PREVIA LEGEND")
        self.resize(1100, 700)
        layout = QVBoxLayout(self)
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        self.table.setShowGrid(True)
        layout.addWidget(self.table, 1)

        ws = wb.active
        try:
            from openpyxl.utils import range_boundaries, get_column_letter
        except Exception:
            return
        dim = ws.calculate_dimension()
        min_col, min_row, max_col, max_row = range_boundaries(dim)
        rows = max_row - min_row + 1
        cols = max_col - min_col + 1
        self.table.setRowCount(rows)
        self.table.setColumnCount(cols)

        # tamaños de columnas/filas
        for c in range(min_col, max_col + 1):
            letter = get_column_letter(c)
            width = ws.column_dimensions[letter].width
            if width:
                self.table.setColumnWidth(c - min_col, int(width * 7 + 12))
        for r in range(min_row, max_row + 1):
            height = ws.row_dimensions[r].height
            if height:
                self.table.setRowHeight(r - min_row, int(height * 1.33))

        # merges
        for merge in ws.merged_cells.ranges:
            if merge.min_row < min_row or merge.max_row > max_row:
                continue
            r = merge.min_row - min_row
            c = merge.min_col - min_col
            self.table.setSpan(r, c, merge.max_row - merge.min_row + 1, merge.max_col - merge.min_col + 1)

        # celdas
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = ws.cell(r, c)
                text = "" if cell.value is None else str(cell.value)
                item = QTableWidgetItem(text)

                # fuente
                if cell.font:
                    f = item.font()
                    f.setBold(bool(cell.font.bold))
                    f.setItalic(bool(cell.font.italic))
                    if cell.font.sz:
                        try:
                            f.setPointSize(int(cell.font.sz))
                        except Exception:
                            pass
                    item.setFont(f)
                    if cell.font.color and cell.font.color.type == "rgb":
                        try:
                            item.setForeground(QColor(f"#{cell.font.color.rgb[-6:]}"))
                        except Exception:
                            pass

                # fondo
                if cell.fill and cell.fill.patternType == "solid" and cell.fill.fgColor and cell.fill.fgColor.type == "rgb":
                    try:
                        item.setBackground(QColor(f"#{cell.fill.fgColor.rgb[-6:]}"))
                    except Exception:
                        pass

                # alineación
                align = cell.alignment
                if align and align.horizontal:
                    h = align.horizontal
                    if h == "center":
                        item.setTextAlignment(Qt.AlignCenter)
                    elif h == "right":
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    else:
                        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

                self.table.setItem(r - min_row, c - min_col, item)


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
        self._equipos_bt: List[str] = []
        self._equipos_mt: List[str] = []
        self._usos_bt: List[str] = []
        self._usos_mt: List[str] = []
        self._equipos_btu_ft: Dict[str, float] = {}
        self._familia_options = ["AUTO", "FRONTAL BAJA", "FRONTAL MEDIA", "DUAL"]
        self._cold_engine = None

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
        self.btn_preview = QPushButton("VISUALIZAR LEGEND")
        self.btn_export = QPushButton("EXPORTAR LEGEND")
        self.btn_save_proj.setEnabled(False)
        self.btn_preview.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.lbl_proj_folder = QLabel("CARPETA: --")
        self.lbl_proj_folder.setToolTip("")
        self.lbl_total_general = QLabel("TOTAL GENERAL: 0.0")
        proj_hdr = QHBoxLayout()
        proj_hdr.addWidget(self.btn_open_proj)
        proj_hdr.addWidget(self.btn_new_proj)
        proj_hdr.addWidget(self.btn_save_proj)
        proj_hdr.addWidget(self.btn_preview)
        proj_hdr.addWidget(self.btn_export)
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

        # VENDEDOR
        vendedor_edit = QLineEdit()
        vendedor_edit.textEdited.connect(lambda txt, w=vendedor_edit: self._on_spec_changed_upper("vendedor", w, txt))
        self.spec_fields["vendedor"] = vendedor_edit
        _add_single("VENDEDOR", vendedor_edit)

        # TIPO DE SISTEMA + DISTRIBUCIÓN TUBERÍA
        tipo_cb = QComboBox()
        tipo_cb.addItems(["", "RACK", "WATERLOOP"])
        tipo_cb.setEditable(False)
        tipo_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("tipo_sistema"))
        self.spec_fields["tipo_sistema"] = tipo_cb
        distrib_cb = QComboBox()
        distrib_cb.addItems(["", "AMERICANA", "LOOP"])
        distrib_cb.setEditable(False)
        distrib_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("distribucion_tuberia"))
        self.spec_fields["distribucion_tuberia"] = distrib_cb
        _add_double("TIPO DE SISTEMA", tipo_cb, "DISTRIBUCIÓN TUBERÍA", distrib_cb)
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

        # Deshielos + Expansión
        self.deshielos_cb = QComboBox()
        self.deshielos_cb.addItems(["", "ELÉCTRICO", "GAS CALIENTE", "GAS TIBIO"])
        self.deshielos_cb.setEditable(False)
        self.expansion_cb = QComboBox()
        self.expansion_cb.addItems(["", "TERMOSTÁTICA", "ELECTRÓNICA"])
        self.expansion_cb.setEditable(False)
        self.deshielos_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("deshielos"))
        self.expansion_cb.currentTextChanged.connect(lambda _t: self._on_spec_changed("expansion"))
        self.spec_fields["deshielos"] = self.deshielos_cb
        self.spec_fields["expansion"] = self.expansion_cb
        _add_double("DESHIELOS", self.deshielos_cb, "EXPANSIÓN", self.expansion_cb)

        specs_widget = QWidget()
        specs_widget.setLayout(specs_layout)
        proj_outer.addWidget(QLabel("ESPECIFICACIONES TECNICAS"))
        proj_outer.addWidget(specs_widget)

        self.btn_open_proj.clicked.connect(self._on_open_project)
        self.btn_new_proj.clicked.connect(self._on_new_project)
        self.btn_save_proj.clicked.connect(self._on_save_project)
        self.btn_preview.clicked.connect(self._on_preview_legend)
        self.btn_export.clicked.connect(self._on_export_legend)

        self.bt_model = LegendItemsTableModel([])
        self.mt_model = LegendItemsTableModel([])
        self.bt_model.set_limit_callback(self._on_dim_limit)
        self.mt_model.set_limit_callback(self._on_dim_limit)
        self._init_cold_engine()
        self._apply_usage_map()

        self.bt_view = QTableView()
        self.bt_view.setModel(self.bt_model)
        self.bt_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.bt_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.bt_view.setFocusPolicy(Qt.StrongFocus)
        self.bt_view.setEditTriggers(QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed)
        self.bt_view.setStyleSheet(
            "QTableView::item:selected { background: #eaf2ff; color: #0f172a; }"
        )
        self.bt_view.horizontalHeader().setStretchLastSection(True)
        self.bt_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.bt_view.verticalHeader().setVisible(False)
        self.bt_view.verticalHeader().setDefaultSectionSize(28)
        self.bt_view.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.bt_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.bt_view.customContextMenuRequested.connect(
            lambda pos, v=self.bt_view, m=self.bt_model: self._show_table_menu(v, m, "bt", pos)
        )
        self.bt_model.modelReset.connect(self._recalc_totals)
        self.bt_model.dataChanged.connect(self._mark_dirty_and_totals)
        self.bt_model.rowsInserted.connect(self._mark_dirty_and_totals)
        self.bt_model.rowsRemoved.connect(self._mark_dirty_and_totals)
        self.bt_model.modelReset.connect(lambda *_: self._apply_dim_spans(self.bt_view, self.bt_model))
        self.bt_model.rowsInserted.connect(lambda *_: self._apply_dim_spans(self.bt_view, self.bt_model))
        self.bt_model.rowsRemoved.connect(lambda *_: self._apply_dim_spans(self.bt_view, self.bt_model))
        self.bt_model.dataChanged.connect(lambda *_: self._apply_dim_spans(self.bt_view, self.bt_model))
        self.bt_view.keyPressEvent = lambda event, v=self.bt_view: self._table_key_press(event, self.bt_model, v)
        self._install_shortcuts(self.bt_view, self.bt_model)
        self.bt_model.rowsInserted.connect(lambda *_: self._ensure_combo_editors(self.bt_view))
        self.bt_model.modelReset.connect(lambda *_: self._ensure_combo_editors(self.bt_view))

        self.mt_view = QTableView()
        self.mt_view.setModel(self.mt_model)
        self.mt_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.mt_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.mt_view.setFocusPolicy(Qt.StrongFocus)
        self.mt_view.setEditTriggers(QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed)
        self.mt_view.setStyleSheet(
            "QTableView::item:selected { background: #eaf2ff; color: #0f172a; }"
        )
        self.mt_view.horizontalHeader().setStretchLastSection(True)
        self.mt_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.mt_view.verticalHeader().setVisible(False)
        self.mt_view.verticalHeader().setDefaultSectionSize(28)
        self.mt_view.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.mt_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.mt_view.customContextMenuRequested.connect(
            lambda pos, v=self.mt_view, m=self.mt_model: self._show_table_menu(v, m, "mt", pos)
        )
        self.mt_model.modelReset.connect(self._recalc_totals)
        self.mt_model.dataChanged.connect(self._mark_dirty_and_totals)
        self.mt_model.rowsInserted.connect(self._mark_dirty_and_totals)
        self.mt_model.rowsRemoved.connect(self._mark_dirty_and_totals)
        self.mt_model.modelReset.connect(lambda *_: self._apply_dim_spans(self.mt_view, self.mt_model))
        self.mt_model.rowsInserted.connect(lambda *_: self._apply_dim_spans(self.mt_view, self.mt_model))
        self.mt_model.rowsRemoved.connect(lambda *_: self._apply_dim_spans(self.mt_view, self.mt_model))
        self.mt_model.dataChanged.connect(lambda *_: self._apply_dim_spans(self.mt_view, self.mt_model))
        self.mt_view.keyPressEvent = lambda event, v=self.mt_view: self._table_key_press(event, self.mt_model, v)
        self._install_shortcuts(self.mt_view, self.mt_model)
        self.mt_model.rowsInserted.connect(lambda *_: self._ensure_combo_editors(self.mt_view))
        self.mt_model.modelReset.connect(lambda *_: self._ensure_combo_editors(self.mt_view))

        # Delegates para equipos/usos desde data/LEGEND
        self._equipos_bt_delegate = ComboDelegate(lambda: self._equipos_bt, editable=True, parent=self)
        self._equipos_mt_delegate = ComboDelegate(lambda: self._equipos_mt, editable=True, parent=self)
        self._usos_bt_delegate = ComboDelegate(lambda: self._usos_bt, editable=True, parent=self)
        self._usos_mt_delegate = ComboDelegate(lambda: self._usos_mt, editable=True, parent=self)
        self._familia_delegate = ComboDelegate(lambda: self._familia_options, editable=False, parent=self)
        self._dim_delegate = DimDelegate(parent=self)
        col_equipo = PROJ_COLUMNS.index("equipo")
        col_uso = PROJ_COLUMNS.index("uso")
        col_largo = PROJ_COLUMNS.index("largo_m")
        col_familia = PROJ_COLUMNS.index("familia")
        self.bt_view.setItemDelegateForColumn(col_equipo, self._equipos_bt_delegate)
        self.mt_view.setItemDelegateForColumn(col_equipo, self._equipos_mt_delegate)
        self.bt_view.setItemDelegateForColumn(col_uso, self._usos_bt_delegate)
        self.mt_view.setItemDelegateForColumn(col_uso, self._usos_mt_delegate)
        self.bt_view.setItemDelegateForColumn(col_familia, self._familia_delegate)
        self.mt_view.setItemDelegateForColumn(col_familia, self._familia_delegate)
        self.bt_view.setItemDelegateForColumn(col_largo, self._dim_delegate)
        self.mt_view.setItemDelegateForColumn(col_largo, self._dim_delegate)
        self._ensure_combo_editors(self.bt_view)
        self._ensure_combo_editors(self.mt_view)
        self._apply_dim_spans(self.bt_view, self.bt_model)
        self._apply_dim_spans(self.mt_view, self.mt_model)
        # Ramales y tablas BT
        self.spin_bt = QSpinBox()
        self.spin_bt.setRange(1, 20)
        self.spin_bt.setEnabled(True)
        self.btn_apply_bt = QPushButton("APLICAR")
        self.btn_apply_bt.clicked.connect(lambda: self._on_bt_ramales_changed(self.spin_bt.value()))

        proj_outer.addWidget(QLabel("BAJA (BT)"))
        ram_bt_layout = QHBoxLayout()
        ram_bt_layout.addWidget(QLabel("RAMALES BT:"))
        ram_bt_layout.addWidget(self.spin_bt)
        ram_bt_layout.addWidget(self.btn_apply_bt)
        ram_bt_layout.addSpacing(12)
        ram_bt_layout.addStretch(1)
        proj_outer.addLayout(ram_bt_layout)

        # BT fila controls removed per request
        proj_outer.addWidget(self.bt_view)
        self.lbl_total_bt = QLabel("TOTAL BT: 0.0")
        proj_outer.addWidget(self.lbl_total_bt)

        # Ramales y tablas MT
        self.spin_mt = QSpinBox()
        self.spin_mt.setRange(1, 20)
        self.spin_mt.setEnabled(True)
        self.btn_apply_mt = QPushButton("APLICAR")
        self.btn_apply_mt.clicked.connect(lambda: self._on_mt_ramales_changed(self.spin_mt.value()))

        proj_outer.addWidget(QLabel("MEDIA (MT)"))
        ram_mt_layout = QHBoxLayout()
        ram_mt_layout.addWidget(QLabel("RAMALES MT:"))
        ram_mt_layout.addWidget(self.spin_mt)
        ram_mt_layout.addWidget(self.btn_apply_mt)
        ram_mt_layout.addSpacing(12)
        ram_mt_layout.addStretch(1)
        proj_outer.addLayout(ram_mt_layout)

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
        self.btn_preview.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.lbl_proj_folder.setText(f"CARPETA: {self.project_dir}")
        self.lbl_proj_folder.setToolTip(str(self.project_dir))
        self._set_dirty(True)

    def _init_cold_engine(self) -> None:
        if not ColdRoomEngine:
            return
        try:
            data_path = Path("data/cuartos_frios/cuartos_frios_data.json")
            self._cold_engine = ColdRoomEngine(data_path)
            self.bt_model.set_cold_engine(self._cold_engine)
            self.mt_model.set_cold_engine(self._cold_engine)
        except Exception:
            self._cold_engine = None

    def _apply_usage_map(self) -> None:
        usage_map = {
            "IC": "HELADOS",
            "HELADO": "HELADOS",
            "HELADOS": "HELADOS",
            "CC": "COMIDA CONGELADA",
            "COMIDA CONGELADA": "COMIDA CONGELADA",
            "PESCADO": "COMIDA CONGELADA",
            "POLLO FRIZADO": "COMIDA CONGELADA",
            "CARNE": "CARNES",
            "CERDO": "CARNES",
            "POLLO": "CARNES",
            "CARNES": "CARNES",
            "BEBIDAS": "LACTEOS",
            "C. FRIAS": "LACTEOS",
            "C.FRIAS": "LACTEOS",
            "DELI": "LACTEOS",
            "LACTEOS": "LACTEOS",
            "PANADERIA": "LACTEOS",
            "VARIOS": "LACTEOS",
            "PREP": "PROCESO",
            "FRUVER": "PROCESO",
            "PROCESO": "PROCESO",
        }
        self.bt_model.set_usage_map(usage_map)
        self.mt_model.set_usage_map(usage_map)

    def _on_save_project(self) -> None:
        if not LegendProjectService:
            return
        if not self.project_dir:
            dir_path = QFileDialog.getExistingDirectory(self, "Selecciona carpeta para guardar", str(Path.cwd()))
            if not dir_path:
                return
            self.project_dir = Path(dir_path)
            self.lbl_proj_folder.setText(f"CARPETA: {self.project_dir}")
            self.lbl_proj_folder.setToolTip(str(self.project_dir))
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
            self.btn_preview.setEnabled(True)
            self.btn_export.setEnabled(True)
            self._render_project()
            self._set_dirty(False)
        except Exception as exc:
            QMessageBox.critical(self, "LEGEND", f"No se pudo cargar proyecto:\n{exc}")

    def _collect_project_data(self) -> Dict[str, Any]:
        data = dict(self.project_data) if isinstance(self.project_data, dict) else {}
        data["bt_items"] = list(self.bt_model.items)
        data["mt_items"] = list(self.mt_model.items)
        specs = data.get("specs", {}) if isinstance(data, dict) else {}
        for k, widget in self.spec_fields.items():
            if isinstance(widget, QComboBox):
                specs[k] = widget.currentText()
            else:
                specs[k] = widget.text()
        data["specs"] = specs
        return data

    def _on_preview_legend(self) -> None:
        if not build_legend_workbook:
            QMessageBox.warning(self, "LEGEND", "Exportador no disponible.")
            return
        if not self._has_any_legend_data():
            QMessageBox.information(self, "LEGEND", "NO HAY DATOS PARA VISUALIZAR.")
            return
        template_path = self._resolve_legend_template()
        if not template_path:
            QMessageBox.warning(self, "LEGEND", "No se encontró la plantilla legend_template.xlsx en data/legend/plantillas/")
            return
        try:
            wb = build_legend_workbook(template_path, self._collect_project_data())
            dlg = LegendPreviewDialog(wb, self)
            dlg.setWindowModality(Qt.ApplicationModal)
            dlg.raise_()
            dlg.activateWindow()
            dlg.exec()
        except Exception as exc:
            QMessageBox.critical(self, "LEGEND", f"No se pudo generar vista previa:\n{exc}")

    def _on_export_legend(self) -> None:
        if not build_legend_workbook:
            QMessageBox.warning(self, "LEGEND", "Exportador no disponible.")
            return
        if not self._has_any_legend_data():
            QMessageBox.information(self, "LEGEND", "NO HAY DATOS PARA EXPORTAR.")
            return
        template_path = self._resolve_legend_template()
        if not template_path:
            QMessageBox.warning(self, "LEGEND", "No se encontró la plantilla legend_template.xlsx en data/legend/plantillas/")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exportar Legend", str(Path.cwd()), "Excel (*.xlsx)")
        if not path:
            return
        try:
            wb = build_legend_workbook(template_path, self._collect_project_data())
            wb.save(path)
            QMessageBox.information(self, "LEGEND", "EXPORTACIÓN LISTA.")
        except Exception as exc:
            QMessageBox.critical(self, "LEGEND", f"No se pudo exportar:\n{exc}")

    def _resolve_legend_template(self) -> Path | None:
        candidates = [
            Path("data/LEGEND/plantillas/legend_template.xlsx"),
            Path("data/legend/plantillas/legend_template.xlsx"),
        ]
        for p in candidates:
            if p.exists():
                return p
        return None

    def _has_any_legend_data(self) -> bool:
        try:
            if self.bt_model.items or self.mt_model.items:
                return True
        except Exception:
            pass
        try:
            specs = self.project_data.get("specs", {}) if isinstance(self.project_data, dict) else {}
            return any(str(v).strip() for v in specs.values())
        except Exception:
            return False
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
        if key == "distribucion_tuberia":
            self._update_loop_column_visibility()
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
        self._update_loop_column_visibility()
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
        self.bt_model.set_items(bt_items)
        self.mt_model.set_items(mt_items)
        self.bt_model.recompute_all()
        self.mt_model.recompute_all()
        self._fit_table_to_contents_view(self.bt_view)
        self._fit_table_to_contents_view(self.mt_view)
        self._recalc_totals()

    def _add_row(self, model, view: QTableView) -> None:
        model.add_row({})
        self._fit_table_to_contents_view(view)
        bloque = self._block_for_model(model)
        if bloque:
            self._renumber_ramales(bloque)
        else:
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

    def _show_table_menu(self, view: QTableView, model: "LegendItemsTableModel", bloque: str, pos) -> None:
        index = view.indexAt(pos)
        if index.isValid():
            if not view.selectionModel().isRowSelected(index.row(), QModelIndex()):
                view.selectRow(index.row())
        menu = QMenu(view)
        act_add = menu.addAction("AGREGAR FILA")
        act_del = menu.addAction("ELIMINAR FILA")
        menu.addSeparator()
        act_copy = menu.addAction("COPIAR")
        act_paste = menu.addAction("PEGAR")
        action = menu.exec(view.mapToGlobal(pos))
        if action == act_add:
            self._add_row_after(view, model, bloque)
            self._renumber_ramales(bloque)
        elif action == act_del:
            self._delete_selected_rows(view, model, bloque)
            self._renumber_ramales(bloque)
        elif action == act_copy:
            self._copy_selection(model, view)
        elif action == act_paste:
            self._paste_selection_at(model, view)
            self._renumber_ramales(bloque)

    def _install_shortcuts(self, view: QTableView, model: "LegendItemsTableModel") -> None:
        if not hasattr(self, "_shortcuts"):
            self._shortcuts = []
        sc_copy = QShortcut(QKeySequence.Copy, view)
        sc_copy.setContext(Qt.WidgetWithChildrenShortcut)
        sc_copy.activated.connect(lambda m=model, v=view: self._copy_selection(m, v))
        self._shortcuts.append(sc_copy)
        sc_paste = QShortcut(QKeySequence.Paste, view)
        sc_paste.setContext(Qt.WidgetWithChildrenShortcut)
        sc_paste.activated.connect(lambda m=model, v=view: self._paste_selection(m, v))
        self._shortcuts.append(sc_paste)
        sc_dup = QShortcut(QKeySequence(Qt.CTRL | Qt.Key_D), view)
        sc_dup.setContext(Qt.WidgetWithChildrenShortcut)
        sc_dup.activated.connect(lambda m=model, v=view: self._dup_selection(m, v))
        self._shortcuts.append(sc_dup)

    def _add_row_after(self, view: QTableView, model: "LegendItemsTableModel", bloque: str) -> None:
        sel = view.selectionModel().selectedRows()
        insert_at = model.rowCount()
        ramal_val = 1
        if sel:
            row_idx = sel[0].row()
            insert_at = row_idx + 1
            try:
                ramal_val = self._to_int(model.items[row_idx].get("ramal", 1))
            except Exception:
                ramal_val = 1
        row = self._default_row(bloque)
        row["ramal"] = ramal_val
        model.insert_rows_at(insert_at, [row])
        self._fit_table_to_contents_view(view)
        self._set_dirty(True)

    def _delete_selected_rows(self, view: QTableView, model: "LegendItemsTableModel", bloque: str) -> None:
        sel = view.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "LEGEND", "SELECCIONA UNA FILA PRIMERO.")
            return
        rows = sorted((i.row() for i in sel), reverse=True)
        for r in rows:
            model.del_row(r)
        self._fit_table_to_contents_view(view)
        self._set_dirty(True)

    def _paste_selection_at(self, model: "LegendItemsTableModel", view: QTableView) -> None:
        text = QGuiApplication.clipboard().text()
        if not text:
            return
        sel = view.selectionModel().selectedRows()
        insert_at = sel[0].row() if sel else model.rowCount()
        rows: List[dict] = []
        for line in text.splitlines():
            parts = line.split("\t")
            row = model.default_row.copy()
            for idx, key in enumerate(model.headers):
                if idx < len(parts):
                    row[key] = parts[idx]
            rows.append(row)
        model.insert_rows_at(insert_at, rows)
        self._fit_table_to_contents_view(view)
        bloque = self._block_for_model(model)
        if bloque:
            self._renumber_ramales(bloque)
        else:
            self._set_dirty(True)

    def _selected_row_indices(self, view: QTableView) -> List[int]:
        rows = sorted({i.row() for i in view.selectionModel().selectedRows()})
        if rows:
            return rows
        rows = sorted({i.row() for i in view.selectionModel().selectedIndexes()})
        if rows:
            return rows
        current = view.currentIndex()
        if current.isValid():
            return [current.row()]
        return []

    def _renumber_ramales(self, bloque: str) -> None:
        model = self.bt_model if bloque == "bt" else self.mt_model
        view = self.bt_view if bloque == "bt" else self.mt_view
        key_items = "bt_items" if bloque == "bt" else "mt_items"
        key_ramales = "bt_ramales" if bloque == "bt" else "mt_ramales"
        items = model.items
        row_count = model.rowCount()
        for idx, row in enumerate(items):
            if isinstance(row, dict):
                row["ramal"] = idx + 1
        if row_count > 0:
            col = PROJ_COLUMNS.index("ramal")
            top = model.index(0, col)
            bottom = model.index(row_count - 1, col)
            model.dataChanged.emit(top, bottom, [Qt.DisplayRole, Qt.EditRole])
        self.project_data[key_items] = model.items
        self.project_data[key_ramales] = max(row_count, 1)
        if bloque == "bt":
            self.spin_bt.blockSignals(True)
            self.spin_bt.setValue(self.project_data[key_ramales])
            self.spin_bt.blockSignals(False)
        else:
            self.spin_mt.blockSignals(True)
            self.spin_mt.setValue(self.project_data[key_ramales])
            self.spin_mt.blockSignals(False)
        self._fit_table_to_contents_view(view)
        self._set_dirty(True)

    def _fit_table_to_contents_view(self, view: QTableView) -> None:
        model = view.model()
        row_count = model.rowCount() if model else 0
        header_h = view.horizontalHeader().height()
        if model and row_count > 0:
            try:
                view.resizeRowsToContents()
            except Exception:
                pass
            total_rows_h = view.verticalHeader().length()
        else:
            total_rows_h = view.verticalHeader().defaultSectionSize()
        scroll_h = view.horizontalScrollBar().sizeHint().height()
        h = header_h + total_rows_h + view.frameWidth() * 2 + scroll_h
        header = view.horizontalHeader()
        if model:
            for col in range(model.columnCount()):
                header.setSectionResizeMode(col, QHeaderView.Interactive)
        view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        view.setFixedHeight(h)
        view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        view.setAlternatingRowColors(True)
        self._apply_combo_column_widths(view)
        self._apply_dim_spans(view, model)
        try:
            view.scrollToTop()
        except Exception:
            pass

    def _refresh_table_view(self, view: QTableView) -> None:
        try:
            view.doItemsLayout()
            view.viewport().update()
        except Exception:
            pass

    def _notify_model_changed(self, model: "LegendItemsTableModel") -> None:
        try:
            model.layoutChanged.emit()
            if model.rowCount() > 0:
                top = model.index(0, 0)
                bottom = model.index(model.rowCount() - 1, model.columnCount() - 1)
                model.dataChanged.emit(top, bottom, [Qt.DisplayRole, Qt.EditRole])
        except Exception:
            pass

    def _schedule_after_insert(self, view: QTableView) -> None:
        def _apply() -> None:
            try:
                model = view.model()
                if model:
                    try:
                        model.beginResetModel()
                        model.endResetModel()
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                view.viewport().update()
            except Exception:
                pass
            try:
                view.resizeRowsToContents()
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

    def _on_dim_limit(self, label: str, max_val: float) -> None:
        QMessageBox.warning(
            self,
            "Dimensión fuera de rango",
            f"{label} máximo {max_val:.2f} m.\nUSE CUADRO DE CARGAS CUARTOS INDUSTRIALES.",
        )

    def _update_loop_column_visibility(self) -> None:
        try:
            distrib = ""
            widget = self.spec_fields.get("distribucion_tuberia")
            if isinstance(widget, QComboBox):
                distrib = widget.currentText().strip().upper()
            col_loop = PROJ_COLUMNS.index("loop")
            show_loop = distrib == "LOOP"
            self.bt_view.setColumnHidden(col_loop, not show_loop)
            self.mt_view.setColumnHidden(col_loop, not show_loop)
        except Exception:
            pass

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

    def _block_for_model(self, model: "LegendItemsTableModel") -> str:
        if model is self.bt_model:
            return "bt"
        if model is self.mt_model:
            return "mt"
        return ""

    def _default_row(self, bloque: str) -> dict:
        return dict(self.bt_model.default_row if bloque == "bt" else self.mt_model.default_row)

    def _write_error_log(self, text: str) -> None:
        try:
            log_path = Path.cwd() / "legend_error.log"
            with open(log_path, "a", encoding="utf-8") as fh:
                fh.write(text + "\n")
        except Exception:
            pass

    def _on_bt_ramales_changed(self, value: int) -> None:
        if not isinstance(self.project_data, dict):
            return
        current = self._to_int(self.project_data.get("bt_ramales", 1))
        if value == current:
            if self.bt_model.rowCount() == 0 and value >= 1:
                row = self._default_row("bt")
                row["ramal"] = 1
                self._suspend_updates = True
                self.bt_model.add_rows([row])
                self._suspend_updates = False
                self.project_data["bt_items"] = self.bt_model.items
                self._fit_table_to_contents_view(self.bt_view)
                self._set_dirty(True)
            return
        if value > current and self.bt_model.rowCount() == 0:
            row = self._default_row("bt")
            row["ramal"] = 1
            self._suspend_updates = True
            self.bt_model.add_rows([row])
            self._suspend_updates = False
            self.project_data["bt_items"] = self.bt_model.items
            current = 1
        if value > current:
            self._add_ramales_bulk("bt", value - current)
        else:
            self._trim_ramales_to("bt", value)

    def _on_mt_ramales_changed(self, value: int) -> None:
        if not isinstance(self.project_data, dict):
            return
        current = self._to_int(self.project_data.get("mt_ramales", 1))
        if value == current:
            if self.mt_model.rowCount() == 0 and value >= 1:
                row = self._default_row("mt")
                row["ramal"] = 1
                self._suspend_updates = True
                self.mt_model.add_rows([row])
                self._suspend_updates = False
                self.project_data["mt_items"] = self.mt_model.items
                self._fit_table_to_contents_view(self.mt_view)
                self._set_dirty(True)
            return
        if value > current and self.mt_model.rowCount() == 0:
            row = self._default_row("mt")
            row["ramal"] = 1
            self._suspend_updates = True
            self.mt_model.add_rows([row])
            self._suspend_updates = False
            self.project_data["mt_items"] = self.mt_model.items
            current = 1
        if value > current:
            self._add_ramales_bulk("mt", value - current)
        else:
            self._trim_ramales_to("mt", value)

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
                        nr["largo_m"] = ""
                        nr["ancho_m"] = ""
                        nr["alto_m"] = ""
                        nr["btu_hr"] = 0
                        nr["evap_qty"] = 0
                        nr["evap_modelo"] = ""
                        nr["familia"] = r.get("familia", "AUTO")
                        new_rows.append(nr)
                new_rows_all.extend(new_rows)
                temp_items.extend(new_rows)
                n += 1
            self._suspend_updates = True
            model.add_rows(new_rows_all)
            self._suspend_updates = False
            self._write_error_log(f"[ADD_RAMALES] rows={model.rowCount()} items={len(model.items)}")
            self.project_data[key_items] = model.items
            self.project_data[key_ramales] = n
            if bloque == "bt":
                self.spin_bt.setValue(n)
                self._write_error_log("[ADD_RAMALES] BT add_rows OK")
                self._fit_table_to_contents_view(self.bt_view)
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
                nr["largo_m"] = ""
                nr["ancho_m"] = ""
                nr["alto_m"] = ""
                nr["btu_hr"] = 0
                nr["evap_qty"] = 0
                nr["evap_modelo"] = ""
                nr["familia"] = r.get("familia", "AUTO")
                new_rows.append(nr)
        self._suspend_updates = True
        model.add_rows(new_rows)
        self._suspend_updates = False
        self._write_error_log(f"[ADD_RAMALES] rows={model.rowCount()} items={len(model.items)}")
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

    def _trim_ramales_to(self, bloque: str, target: int) -> None:
        if not isinstance(self.project_data, dict):
            return
        key_items = "bt_items" if bloque == "bt" else "mt_items"
        key_ramales = "bt_ramales" if bloque == "bt" else "mt_ramales"
        model = self.bt_model if bloque == "bt" else self.mt_model
        view = self.bt_view if bloque == "bt" else self.mt_view
        items = [r for r in model.items if isinstance(r, dict)]
        target = max(1, min(20, self._to_int(target)))
        if items:
            new_items = items[:target]
        else:
            new_items = []
        if not new_items and target >= 1:
            row = self._default_row(bloque)
            row["ramal"] = 1
            new_items = [row]
        model.set_items(new_items)
        self.project_data[key_items] = model.items
        self.project_data[key_ramales] = target
        if bloque == "bt":
            self.spin_bt.blockSignals(True)
            self.spin_bt.setValue(target)
            self.spin_bt.blockSignals(False)
            self._fit_table_to_contents_view(self.bt_view)
        else:
            self.spin_mt.blockSignals(True)
            self.spin_mt.setValue(target)
            self.spin_mt.blockSignals(False)
            self._fit_table_to_contents_view(self.mt_view)
        self._renumber_ramales(bloque)

    def _recalc_totals(self) -> None:
        def _sum(model: "LegendItemsTableModel") -> float:
            total = 0.0
            for r in model.items:
                try:
                    total += float(r.get("btu_hr", r.get("carga_btu_h", 0)) or 0)
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
        if event.matches(QKeySequence.StandardKey.Copy):
            self._copy_selection(model, view)
            event.accept()
            return
        if event.matches(QKeySequence.StandardKey.Paste):
            self._paste_selection(model, view)
            event.accept()
            return
        QTableView.keyPressEvent(view, event)

    def _copy_selection(self, model: "LegendItemsTableModel", view: QTableView) -> None:
        rows = self._selected_row_indices(view)
        if not rows:
            return
        lines = []
        for r in rows:
            item = model.items[r]
            vals = [str(item.get(k, "")) for k in model.headers]
            lines.append("\t".join(vals))
        QGuiApplication.clipboard().setText("\n".join(lines))

    def _paste_selection(self, model: "LegendItemsTableModel", view: QTableView) -> None:
        self._paste_selection_at(model, view)

    def _dup_selection(self, model: "LegendItemsTableModel", view: QTableView) -> None:
        rows = self._selected_row_indices(view)
        if not rows:
            return
        for r in rows:
            model.dup_row(r)
        self._fit_table_to_contents_view(view)
        bloque = self._block_for_model(model)
        if bloque:
            self._renumber_ramales(bloque)
        else:
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
            nr["largo_m"] = ""
            nr["ancho_m"] = ""
            nr["alto_m"] = ""
            nr["btu_hr"] = 0
            nr["evap_qty"] = 0
            nr["evap_modelo"] = ""
            nr["familia"] = r.get("familia", "AUTO")
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
        equipos_bt = data.get("equipos_bt", []) if isinstance(data, dict) else []
        equipos_mt = data.get("equipos_mt", []) if isinstance(data, dict) else []
        btu_map: Dict[str, float] = {}
        for e in equipos:
            try:
                name = getattr(e, "equipo", "")
                val = float(getattr(e, "btu_hr_ft", 0.0) or 0.0)
                key = self._normalize_equipo_name(name)
                if key:
                    btu_map[key] = val
            except Exception:
                continue
        self._equipos_btu_ft = btu_map
        self.bt_model.set_btu_map(self._equipos_btu_ft)
        self.mt_model.set_btu_map(self._equipos_btu_ft)
        if equipos_bt:
            self._equipos_bt = sorted(
                [getattr(e, "equipo", "") for e in equipos_bt if getattr(e, "equipo", "")],
                key=lambda x: x.upper(),
            )
        else:
            self._equipos_bt = sorted(
                [getattr(e, "equipo", "") for e in equipos if getattr(e, "equipo", "")],
                key=lambda x: x.upper(),
            )
        if equipos_mt:
            self._equipos_mt = sorted(
                [getattr(e, "equipo", "") for e in equipos_mt if getattr(e, "equipo", "")],
                key=lambda x: x.upper(),
            )
        else:
            self._equipos_mt = list(self._equipos_bt)
        self._usos_bt = sorted(
            usos.get("BT", []) if isinstance(usos, dict) else [],
            key=lambda x: str(x).upper(),
        )
        self._usos_mt = sorted(
            usos.get("MT", []) if isinstance(usos, dict) else [],
            key=lambda x: str(x).upper(),
        )
        tevap_bt, tevap_mt, def_bt, def_mt = self._load_tevap_maps()
        self.bt_model.set_tevap_map(tevap_bt, def_bt)
        self.mt_model.set_tevap_map(tevap_mt, def_mt)
        self._apply_combo_column_widths(self.bt_view)
        self._apply_combo_column_widths(self.mt_view)
        self._refresh_combo_editors(self.bt_view)
        self._refresh_combo_editors(self.mt_view)

    def _load_tevap_maps(self) -> tuple[Dict[str, float], Dict[str, float], float, float]:
        tevap_bt: Dict[str, float] = {}
        tevap_mt: Dict[str, float] = {}
        def_bt = -22.0
        def_mt = 15.0
        config_path = Path("data/legend/config.json")
        if config_path.exists():
            try:
                cfg = json.loads(config_path.read_text(encoding="utf-8"))
                def_bt = float(cfg.get("tevap_bt", def_bt))
                def_mt = float(cfg.get("tevap_mt", def_mt))
            except Exception:
                pass
        tevap_path = Path("data/legend/tevap_por_uso.json")
        if tevap_path.exists():
            try:
                data = json.loads(tevap_path.read_text(encoding="utf-8"))
                tevap_bt = data.get("BT", {}) if isinstance(data, dict) else {}
                tevap_mt = data.get("MT", {}) if isinstance(data, dict) else {}
            except Exception:
                pass
        return tevap_bt, tevap_mt, def_bt, def_mt

    def _apply_combo_column_widths(self, view: QTableView) -> None:
        # Mantener mismas columnas y tamaños en BT/MT, basado en el mayor texto de ambos
        self._sync_table_column_widths()
        try:
            self._refresh_combo_editors(view)
        except Exception:
            pass

    def _sync_table_column_widths(self) -> None:
        if not self.bt_view or not self.mt_view:
            return
        bt_model = self.bt_view.model()
        mt_model = self.mt_view.model()
        if not bt_model or not mt_model:
            return
        fm = self.bt_view.fontMetrics()

        equipos_all = list(self._equipos_bt) + list(self._equipos_mt)
        usos_all = list(self._usos_bt) + list(self._usos_mt)

        min_widths = {
            "equipo": 140,
            "uso": 120,
            "familia": 140,
        }
        combo_cols = {"equipo", "uso", "familia"}

        for col in range(len(PROJ_COLUMNS)):
            key = PROJ_COLUMNS[col]
            header_text = bt_model.headerData(col, Qt.Horizontal, Qt.DisplayRole) or ""
            max_w = fm.horizontalAdvance(str(header_text))

            if key == "equipo":
                for t in equipos_all:
                    max_w = max(max_w, fm.horizontalAdvance(str(t)))
            elif key == "uso":
                for t in usos_all:
                    max_w = max(max_w, fm.horizontalAdvance(str(t)))
            elif key == "familia":
                for t in self._familia_options:
                    max_w = max(max_w, fm.horizontalAdvance(str(t)))

            for model in (bt_model, mt_model):
                for row in range(model.rowCount()):
                    try:
                        val = model.index(row, col).data(Qt.DisplayRole)
                    except Exception:
                        val = None
                    if val:
                        max_w = max(max_w, fm.horizontalAdvance(str(val)))

            pad = 70 if key in combo_cols else 24
            width = max_w + pad
            if key in min_widths:
                width = max(width, min_widths[key])
            self.bt_view.setColumnWidth(col, width)
            self.mt_view.setColumnWidth(col, width)

    def _normalize_equipo_name(self, text: str) -> str:
        return " ".join(str(text or "").strip().upper().split())

    def _is_cuarto(self, equipo: str) -> bool:
        return "CUARTO" in (equipo or "").upper()

    def _apply_dim_spans(self, view: QTableView, model: "LegendItemsTableModel" | None) -> None:
        if not view or not model:
            return
        try:
            col_largo = PROJ_COLUMNS.index("largo_m")
        except ValueError:
            return
        view.clearSpans()
        for row in range(model.rowCount()):
            try:
                equipo = str(model.items[row].get("equipo", ""))
            except Exception:
                equipo = ""
            if not self._is_cuarto(equipo):
                view.setSpan(row, col_largo, 1, 3)
        try:
            self._refresh_combo_editors(view)
        except Exception:
            pass

    def _ensure_combo_editors(self, view: QTableView) -> None:
        model = view.model()
        if not model:
            return
        col_equipo = PROJ_COLUMNS.index("equipo")
        col_uso = PROJ_COLUMNS.index("uso")
        col_familia = PROJ_COLUMNS.index("familia")
        for row in range(model.rowCount()):
            for col in (col_equipo, col_uso):
                idx = model.index(row, col)
                view.openPersistentEditor(idx)
            try:
                if hasattr(model, "_row_is_cuarto") and model._row_is_cuarto(row):
                    idx = model.index(row, col_familia)
                    view.openPersistentEditor(idx)
            except Exception:
                pass

    def _refresh_combo_editors(self, view: QTableView) -> None:
        model = view.model()
        if not model:
            return
        col_equipo = PROJ_COLUMNS.index("equipo")
        col_uso = PROJ_COLUMNS.index("uso")
        col_familia = PROJ_COLUMNS.index("familia")
        for row in range(model.rowCount()):
            for col in (col_equipo, col_uso):
                idx = model.index(row, col)
                view.closePersistentEditor(idx)
                view.openPersistentEditor(idx)
            try:
                idx_f = model.index(row, col_familia)
                view.closePersistentEditor(idx_f)
                if hasattr(model, "_row_is_cuarto") and model._row_is_cuarto(row):
                    view.openPersistentEditor(idx_f)
            except Exception:
                pass


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
        cb.view().setStyleSheet(
            "background: #ffffff; color: #0f172a; selection-background-color: #dbeafe; selection-color: #0f172a;"
        )
        cb.setMaxVisibleItems(max(len(items), 10))
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

    def updateEditorGeometry(self, editor, option, index):
        if editor:
            editor.setGeometry(option.rect)
        else:
            super().updateEditorGeometry(editor, option, index)


class DimDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        model = index.model()
        editor = QLineEdit(parent)
        editor.setAlignment(Qt.AlignCenter)
        try:
            if hasattr(model, "_row_is_multipuerta") and model._row_is_multipuerta(index.row()):
                editor.setValidator(QIntValidator(0, 9999, editor))
        except Exception:
            pass
        return editor

    def updateEditorGeometry(self, editor, option, index):
        if editor:
            editor.setGeometry(option.rect)
        else:
            super().updateEditorGeometry(editor, option, index)


class LegendItemsTableModel(QAbstractTableModel):
    headers = PROJ_COLUMNS
    MAX_LW_M = 12.8
    MAX_H_M = 3.7

    def __init__(self, items: List[dict] | None = None):
        super().__init__()
        self.items: List[dict] = items or []
        self._btu_map: Dict[str, float] = {}
        self._cold_engine = None
        self._usage_map: Dict[str, str] = {}
        self._tevap_map: Dict[str, float] = {}
        self._tevap_default: float = 0.0
        self._limit_cb: Callable[[str, float], None] | None = None
        self.default_row = {
            "loop": 1,
            "ramal": 1,
            "largo_m": "",
            "ancho_m": "",
            "alto_m": "",
            "dim_ft": "",
            "equipo": "",
            "uso": "",
            "btu_ft": 0.0,
            "btu_hr": 0.0,
            "evap_qty": 0,
            "evap_modelo": "",
            "familia": "AUTO",
        }

    def set_limit_callback(self, cb: Callable[[str, float], None] | None) -> None:
        self._limit_cb = cb

    def rowCount(self, parent=QModelIndex()):
        return len(self.items)

    def columnCount(self, parent=QModelIndex()):
        return len(self.headers)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        item = self.items[index.row()]
        key = self.headers[index.column()]
        if role in (Qt.DisplayRole, Qt.EditRole):
            if key == "largo_m":
                return item.get("largo_m", item.get("dim_m", ""))
            if key == "btu_ft":
                return item.get("btu_ft", item.get("btu_hr_ft", ""))
            if key == "btu_hr":
                return item.get("btu_hr", item.get("carga_btu_h", ""))
            if key == "dim_ft":
                return self._dim_ft_text(index.row())
            if key == "familia" and not self._row_is_cuarto(index.row()):
                return ""
            if key in ("ancho_m", "alto_m") and not self._row_is_cuarto(index.row()):
                return ""
            return item.get(key, "")
        if role == Qt.TextAlignmentRole:
            if key in ("largo_m", "ancho_m", "alto_m") and not self._row_is_cuarto(index.row()):
                return Qt.AlignHCenter | Qt.AlignVCenter
            if key == "dim_ft":
                return Qt.AlignHCenter | Qt.AlignVCenter
            if key in ("loop", "ramal", "largo_m", "ancho_m", "alto_m", "btu_ft", "btu_hr", "evap_qty"):
                return Qt.AlignRight | Qt.AlignVCenter
            return Qt.AlignLeft | Qt.AlignVCenter
        return None

    def setData(self, index: QModelIndex, value, role=Qt.EditRole):
        if role == Qt.EditRole and index.isValid():
            key = self.headers[index.column()]
            if key in ("largo_m", "ancho_m", "alto_m"):
                val = self._parse_float(value)
                if self._row_is_cuarto(index.row()):
                    max_val = self.MAX_LW_M if key in ("largo_m", "ancho_m") else self.MAX_H_M
                    if val > max_val:
                        val = max_val
                        if self._limit_cb:
                            label = "LARGO" if key == "largo_m" else "ANCHO" if key == "ancho_m" else "ALTO"
                            self._limit_cb(label, max_val)
                if key == "largo_m":
                    self.items[index.row()]["largo_m"] = val
                    self.items[index.row()]["dim_m"] = val
                else:
                    self.items[index.row()][key] = val
                self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                self._recompute_row(index.row())
                return True
            if key in ("ancho_m", "alto_m", "uso", "evap_qty", "familia"):
                self.items[index.row()][key] = value
                self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                self._recompute_row(index.row())
                return True
            if key == "equipo":
                self.items[index.row()][key] = value
                self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                self._recompute_row(index.row())
                return True
            elif key == "btu_hr":
                self.items[index.row()]["btu_hr"] = value
                self.items[index.row()]["carga_btu_h"] = value
            elif key == "btu_ft":
                self.items[index.row()]["btu_ft"] = value
                self.items[index.row()]["btu_hr_ft"] = value
            else:
                self.items[index.row()][key] = value
            self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
            return True
        return False

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.ItemIsEnabled
        key = self.headers[index.column()]
        if key in ("btu_ft", "btu_hr", "dim_ft"):
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if key == "familia" and not self._row_is_cuarto(index.row()):
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if key in ("ancho_m", "alto_m") and not self._row_is_cuarto(index.row()):
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def headerData(self, section: int, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            key = self.headers[section]
            if key == "largo_m":
                return "LARGO"
            if key == "ancho_m":
                return "ANCHO"
            if key == "alto_m":
                return "ALTO"
            if key == "btu_ft":
                return "BTU/FT"
            if key == "btu_hr":
                return "BTU/HR"
            if key == "dim_ft":
                return "DIMENSIONES FT"
            if key == "evap_modelo":
                return "MODELO"
            if key == "evap_qty":
                return "# EVAPORADORES"
            if key == "familia":
                return "FAMILIA"
            return key.upper()
        return None

    def _row_is_cuarto(self, row: int) -> bool:
        try:
            equipo = str(self.items[row].get("equipo", "")).upper()
            return "CUARTO" in equipo
        except Exception:
            return False

    def _row_is_multipuerta(self, row: int) -> bool:
        try:
            equipo = str(self.items[row].get("equipo", "")).upper()
            if "MULTIPUERTA" not in equipo:
                return False
            return "CONGEL" in equipo or "REFRIG" in equipo or "CONSERV" in equipo
        except Exception:
            return False

    def _m_to_ft_round(self, m: float) -> int:
        try:
            if self._cold_engine and hasattr(self._cold_engine, "m_to_ft_round"):
                return int(self._cold_engine.m_to_ft_round(m))
        except Exception:
            pass
        return int(round(m * 3.28))

    def _auto_evap_qty_furniture(self, row: int, largo_m: float) -> int:
        if largo_m <= 0:
            return 0
        if self._row_is_multipuerta(row):
            doors = int(round(largo_m))
            if doors <= 0:
                return 0
            if doors == 1:
                return 1
            # Preferir juegos de 4, luego 3 y 2 (evitar 1 si se puede)
            max4 = doors // 4
            for n4 in range(max4, -1, -1):
                rem4 = doors - (4 * n4)
                if rem4 == 0:
                    return n4
                max3 = rem4 // 3
                for n3 in range(max3, -1, -1):
                    rem = rem4 - (3 * n3)
                    if rem == 0:
                        return n4 + n3
                    if rem % 2 == 0:
                        n2 = rem // 2
                        return n4 + n3 + n2
            return doors
        # Preferencias en metros (misma lógica que Carga Eléctrica)
        preferred = [3.75, 2.5, 1.9]
        rem = largo_m
        modules = 0
        limit = 0
        while rem > 0.2 and limit < 200:
            pick = None
            for p in preferred:
                if p <= rem + 0.05:
                    pick = p
                    break
            if pick is None:
                pick = preferred[-1]
            rem -= pick
            modules += 1
            limit += 1
        if modules == 0:
            modules = 1
        return modules

    def _dim_ft_text(self, row: int) -> str:
        if row < 0 or row >= len(self.items):
            return ""
        item = self.items[row]
        largo = self._parse_float(item.get("largo_m", item.get("dim_m", 0)))
        if self._row_is_cuarto(row):
            ancho = self._parse_float(item.get("ancho_m", 0))
            alto = self._parse_float(item.get("alto_m", 0))
            if largo <= 0 or ancho <= 0 or alto <= 0:
                return ""
            l_ft = self._m_to_ft_round(largo)
            w_ft = self._m_to_ft_round(ancho)
            h_ft = self._m_to_ft_round(alto)
            return f"{l_ft}x{w_ft}x{h_ft}"
        if self._row_is_multipuerta(row):
            puertas = int(round(largo)) if largo > 0 else 0
            return f"{puertas} PUERTAS" if puertas > 0 else ""
        if largo <= 0:
            return ""
        ft = largo * 3.28084
        return f"{ft:.2f}"

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
        self._recompute_row(len(self.items) - 1)

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
        for r in range(start, end + 1):
            self._recompute_row(r)

    def insert_rows_at(self, index: int, rows: List[dict]) -> None:
        if not rows:
            return
        index = max(0, min(index, len(self.items)))
        start = index
        end = start + len(rows) - 1
        self.beginInsertRows(QModelIndex(), start, end)
        insert_list: List[dict] = []
        for row in rows:
            if not row:
                row = dict(self.default_row)
            else:
                norm = dict(self.default_row)
                norm.update(row)
                row = norm
            insert_list.append(row)
        self.items[index:index] = insert_list
        self.endInsertRows()
        for r in range(start, end + 1):
            self._recompute_row(r)

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
        self.recompute_all()

    def set_btu_map(self, btu_map: Dict[str, float]) -> None:
        self._btu_map = btu_map or {}
        self.recompute_all()

    def set_cold_engine(self, engine) -> None:
        self._cold_engine = engine
        self.recompute_all()

    def set_usage_map(self, usage_map: Dict[str, str]) -> None:
        self._usage_map = usage_map or {}
        self.recompute_all()

    def set_tevap_map(self, tevap_map: Dict[str, float], default_val: float) -> None:
        norm: Dict[str, float] = {}
        for k, v in (tevap_map or {}).items():
            key = " ".join(str(k or "").strip().upper().split())
            try:
                norm[key] = float(v)
            except Exception:
                continue
        self._tevap_map = norm
        try:
            self._tevap_default = float(default_val)
        except Exception:
            self._tevap_default = 0.0
        self.recompute_all()

    def recompute_all(self) -> None:
        if not self.items:
            return
        for i in range(len(self.items)):
            self._recompute_row(i, emit=False)
        try:
            col_ft = self.headers.index("btu_ft")
            col_hr = self.headers.index("btu_hr")
            col_model = self.headers.index("evap_modelo")
            col_qty = self.headers.index("evap_qty")
            col_dim = self.headers.index("dim_ft")
            col_start = min(col_ft, col_hr, col_model, col_qty, col_dim)
            col_end = max(col_ft, col_hr, col_model, col_qty, col_dim)
            top = self.index(0, col_start)
            bottom = self.index(len(self.items) - 1, col_end)
            self.dataChanged.emit(top, bottom, [Qt.DisplayRole, Qt.EditRole])
        except Exception:
            pass

    def _normalize_equipo(self, text: str) -> str:
        return " ".join(str(text or "").strip().upper().split())

    def _parse_float(self, val) -> float:
        if val is None:
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        try:
            s = str(val).strip().replace(",", ".")
            return float(s) if s else 0.0
        except Exception:
            return 0.0

    def _map_usage(self, usage: str) -> str:
        u = " ".join(str(usage or "").strip().upper().split())
        return self._usage_map.get(u, u)

    def _tevap_for_usage(self, usage: str) -> float:
        u = self._map_usage(usage)
        return float(self._tevap_map.get(u, self._tevap_default))

    def _recompute_row(self, row: int, emit: bool = True) -> None:
        if row < 0 or row >= len(self.items):
            return
        item = self.items[row]
        equipo = self._normalize_equipo(item.get("equipo", ""))
        btu_ft = float(self._btu_map.get(equipo, 0.0))
        raw_largo = self._parse_float(item.get("largo_m", item.get("dim_m", 0)))
        if self._row_is_cuarto(row):
            largo = min(raw_largo, self.MAX_LW_M)
        else:
            largo = raw_largo
        item["largo_m"] = largo
        item["dim_m"] = largo
        if self._row_is_cuarto(row) and self._cold_engine and ColdRoomInputs:
            ancho = min(self._parse_float(item.get("ancho_m", 0)), self.MAX_LW_M)
            alto = min(self._parse_float(item.get("alto_m", 0)), self.MAX_H_M)
            item["ancho_m"] = ancho
            item["alto_m"] = alto
            evap_model = ""
            evap_qty_out = 0
            if largo > 0 and ancho > 0 and alto > 0:
                uso = self._map_usage(item.get("uso", ""))
                item["tevap_f"] = self._tevap_for_usage(uso)
                evap_qty = int(self._parse_float(item.get("evap_qty", 0)))
                n_evap = evap_qty if evap_qty > 0 else None
                fam_raw = " ".join(str(item.get("familia", "")).strip().upper().split())
                fam_override = None
                if fam_raw and fam_raw != "AUTO":
                    if "BAJA" in fam_raw:
                        fam_override = "frontal_wef"
                    elif "MEDIA" in fam_raw:
                        fam_override = "frontal_wefm"
                    elif "DUAL" in fam_raw:
                        fam_override = "dual"
                try:
                    inp = ColdRoomInputs(
                        length_m=largo,
                        width_m=ancho,
                        height_m=alto,
                        usage=uso,
                        n_evaporators=n_evap,
                        family_override=fam_override,
                    )
                    res = self._cold_engine.compute(inp)
                    if res and res.valid:
                        btu_hr = float(res.load_btu_hr or 0.0)
                        evap_model = str(res.evap_model or "")
                        evap_qty_out = int(res.n_used or n_evap or 0)
                    else:
                        btu_hr = 0.0
                except Exception:
                    btu_hr = 0.0
            else:
                btu_hr = 0.0
            btu_ft = 0.0
            item["evap_modelo"] = evap_model
            item["evap_qty"] = evap_qty_out
        elif self._row_is_multipuerta(row):
            largo_int = int(round(largo))
            item["largo_m"] = largo_int
            item["dim_m"] = largo_int
            btu_hr = btu_ft * largo_int
            item["tevap_f"] = self._tevap_for_usage(item.get("uso", ""))
            item["evap_modelo"] = ""
            if int(self._parse_float(item.get("evap_qty", 0))) <= 0:
                item["evap_qty"] = self._auto_evap_qty_furniture(row, largo_int)
        else:
            btu_hr = btu_ft * (largo * 3.28084)
            item["tevap_f"] = self._tevap_for_usage(item.get("uso", ""))
            item["evap_modelo"] = ""
            if int(self._parse_float(item.get("evap_qty", 0))) <= 0:
                item["evap_qty"] = self._auto_evap_qty_furniture(row, largo)
        item["btu_ft"] = btu_ft
        item["btu_hr"] = btu_hr
        item["btu_hr_ft"] = btu_ft
        item["carga_btu_h"] = btu_hr
        item["dim_ft"] = self._dim_ft_text(row)
        if emit:
            try:
                col_ft = self.headers.index("btu_ft")
                col_hr = self.headers.index("btu_hr")
                col_model = self.headers.index("evap_modelo")
                col_qty = self.headers.index("evap_qty")
                col_l = self.headers.index("largo_m")
                col_dim = self.headers.index("dim_ft")
                col_start = min(col_l, col_ft, col_hr, col_model, col_qty, col_dim)
                col_end = max(col_l, col_ft, col_hr, col_model, col_qty, col_dim)
                top = self.index(row, col_start)
                bottom = self.index(row, col_end)
                self.dataChanged.emit(top, bottom, [Qt.DisplayRole, Qt.EditRole])
            except Exception:
                pass


if __name__ == "__main__":
    from PySide6.QtWidgets import QApplication
    import sys

    app = QApplication(sys.argv)
    w = LegendPage()
    w.resize(1000, 700)
    w.show()
    sys.exit(app.exec())
