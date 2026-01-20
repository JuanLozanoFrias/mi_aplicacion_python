from __future__ import annotations

import random
from datetime import datetime
from pathlib import Path
from typing import List

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QPushButton,
    QTabWidget, QTableWidget, QTableWidgetItem, QAbstractItemView,
    QListWidget, QListWidgetItem, QGridLayout, QLineEdit, QFileDialog,
    QMessageBox, QScrollArea, QDialog, QDialogButtonBox, QComboBox
)
from PySide6.QtWidgets import QHeaderView, QSizePolicy
from PySide6.QtGui import QColor, QDesktopServices, QPainter
try:
    from PySide6.QtCharts import QChart, QChartView, QPieSeries
except Exception:
    QChart = QChartView = QPieSeries = None
from PySide6.QtCore import QUrl

from logic.produccion.demo_store import DemoStore, STAGES
from logic.produccion.loaders import load_orders_excel, load_personal_excel
from logic.produccion.models import ProductionOrder, Technician, MaterialItem

TASKS = [
    "ENSAMBLAJE DE CAJA",
    "CABLEADO DE COMPONENTES",
    "CABLEADO DE ILUMINACION",
    "CABLEADO TABLERO A COMPRESOR",
    "PROGRAMACION EQUIPO",
    "PRUEBAS ELECTRICAS",
    "ETIQUETADO",
    "LIMPIEZA FINAL",
]

AREAS_DEFAULT = [
    "ELECTRICIDAD",
    "HORNOS",
    "VITRINAS",
    "CAFETERIAL",
    "NORTE",
    "RACK",
    "TERMINACION",
]

def _card() -> QFrame:
    f = QFrame()
    f.setStyleSheet("QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;}")
    return f


def _kpi_card(title: str, value: str) -> QFrame:
    card = _card()
    lay = QVBoxLayout(card)
    lay.setContentsMargins(12, 10, 12, 10)
    lay.setSpacing(4)
    t = QLabel(title)
    t.setStyleSheet("font-size:11px;font-weight:700;color:#475569;")
    v = QLabel(value)
    v.setStyleSheet("font-size:20px;font-weight:800;color:#0f172a;")
    lay.addWidget(t)
    lay.addWidget(v)
    return card


def _clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in ("", "nan", "none"):
        return ""
    return text


class _NewOpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("NUEVA OP (DEMO)")
        self.setMinimumWidth(420)
        root = QVBoxLayout(self)
        self.ed_cliente = QLineEdit()
        self.ed_cats = QLineEdit()
        self.ed_notes = QLineEdit()
        for lab, w in (
            ("CLIENTE/REF", self.ed_cliente),
            ("CATEGORIAS (CTRL, RESISTENCIAS, VARIADOR...)", self.ed_cats),
            ("NOTAS", self.ed_notes),
        ):
            row = QVBoxLayout()
            l = QLabel(lab); l.setStyleSheet("font-weight:700;color:#0f172a;")
            row.addWidget(l); row.addWidget(w)
            root.addLayout(row)
        btns = QDialogButtonBox()
        self.btn_manual = QPushButton("CREAR MANUAL")
        self.btn_excel = QPushButton("CREAR DESDE EXCEL")
        btns.addButton(self.btn_manual, QDialogButtonBox.AcceptRole)
        btns.addButton(self.btn_excel, QDialogButtonBox.ActionRole)
        self.btn_manual.clicked.connect(self.accept)
        self.btn_excel.clicked.connect(self.reject)
        root.addWidget(btns)


class ProduccionPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        orders = load_orders_excel()
        techs = load_personal_excel()
        if not orders:
            orders = [
                ProductionOrder("OP-1001", "2025-01-10", "ABIERTA", "CLIENTE A", "REF-01", "", "CONTROL, VARIADOR"),
                ProductionOrder("OP-1002", "2025-01-12", "MATERIALES", "CLIENTE B", "REF-02", "", "RESISTENCIAS"),
            ]
        if not techs:
            techs = [
                Technician("1", "JUAN PEREZ", "123", "DIA", "TABLEROS", "", "", ""),
                Technician("2", "MARIA LOPEZ", "456", "NOCHE", "TABLEROS", "", "", ""),
            ]
        self.store = DemoStore(orders, techs)
        self._build_ui()
        self._refresh_orders()
        self._refresh_personal()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        root.addWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)
        cl = QVBoxLayout(container)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(12)

        card = _card()
        cl.addWidget(card)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        top = QHBoxLayout()
        self.btn_new = QPushButton("NUEVA OP (DEMO)")
        self.btn_gen = QPushButton("GENERAR MATERIALES")
        self.btn_exp = QPushButton("EXPORTAR")
        for b in (self.btn_exp,):
            b.setStyleSheet(
                "QPushButton{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #0f62fe, stop:1 #22d3ee);"
                "color:#fff;font-weight:700;border:none;border-radius:16px;padding:8px 16px;}"
            )
        self.btn_exp.clicked.connect(self._on_export_materials)
        top.addWidget(self.btn_exp)
        top.addStretch(1)
        title = QLabel("PRODUCCION")
        title.setStyleSheet("font-size:22px;font-weight:800;color:#0f172a;")
        top.addWidget(title)
        layout.addLayout(top)

        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet("background:#e2e8f5;")
        layout.addWidget(sep)

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("QTabBar::tab{padding:6px 14px;font-weight:700;}")
        layout.addWidget(self.tabs)

        self._build_tab_orders()
        self._build_tab_personal()

    # ---------------------------------------------------------------- Orders
    def _build_tab_orders(self) -> None:
        tab = QWidget()
        self.tabs.addTab(tab, "ORDENES (OP)")
        lay = QVBoxLayout(tab)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(10)

        kpis = QHBoxLayout()
        self.kpi_open = _kpi_card("OPS ABIERTAS", "0")
        self.kpi_materials = _kpi_card("MATERIALES", "0")
        self.kpi_personal = _kpi_card("PERSONAL", "0")
        self.kpi_exec = _kpi_card("EJECUCION", "0")
        self.kpi_closed = _kpi_card("CERRADAS", "0")
        for c in (
            self.kpi_open, self.kpi_materials, self.kpi_personal,
            self.kpi_exec, self.kpi_closed,
        ):
            kpis.addWidget(c)
        lay.addLayout(kpis)

        body = QHBoxLayout()
        lay.addLayout(body)

        self.orders_table = QTableWidget(0, 4)
        self.orders_table.setHorizontalHeaderLabels(["OP", "FECHA", "ESTADO", "% AVANCE"])
        self.orders_table.verticalHeader().setVisible(False)
        self.orders_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.orders_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.orders_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.orders_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.orders_table.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.orders_table.setStyleSheet(
            "QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QHeaderView::section{background:#f8fafc;border:1px solid #e2e8f5;padding:6px;font-weight:700;}"
        )
        self.orders_table.itemSelectionChanged.connect(self._on_select_order)
        body.addWidget(self.orders_table, 3)

        self.detail_panel = _card()
        dlay = QVBoxLayout(self.detail_panel)
        dlay.setContentsMargins(14, 12, 14, 12)
        dlay.setSpacing(8)
        dtitle = QLabel("DETALLE OP")
        dtitle.setStyleSheet("font-weight:800;color:#0f172a;")
        dlay.addWidget(dtitle)

        self.lbl_summary = QLabel("--")
        self.lbl_summary.setStyleSheet("color:#475569;")
        self.lbl_summary.setWordWrap(True)
        dlay.addWidget(self.lbl_summary)

        self.stage_tabs = QTabWidget()
        self.stage_tabs.setStyleSheet(
            "QTabBar::tab{padding:6px 10px;font-weight:700;}"
            "QTabBar::tab:selected{background:#e8f2ff;border-radius:6px;}"
        )
        self.stage_tabs.currentChanged.connect(self._on_stage_tab_changed)
        dlay.addWidget(self.stage_tabs)

        self.material_files_list = QListWidget()
        self.material_files_list.setStyleSheet(
            "QListWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QListWidget::indicator{width:16px;height:16px;}"
            "QListWidget::indicator:unchecked{border:1px solid #cbd5e1;border-radius:4px;background:#fff;}"
            "QListWidget::indicator:checked{border:1px solid #0f62fe;border-radius:4px;background:#0f62fe;}"
        )
        self.material_files_list.itemChanged.connect(self._on_material_file_toggle)
        self.material_files_list.itemDoubleClicked.connect(self._on_material_file_open)

        tab0 = QWidget()
        t0l = QVBoxLayout(tab0)
        t0l.setContentsMargins(8, 8, 8, 8)
        t0l.addWidget(QLabel("SOLICITUD DE MATERIAL"))
        t0l.addWidget(self.material_files_list)
        self.stage_tabs.addTab(tab0, "SOLICITUD MATERIAL")

        self.personal_list = QListWidget()
        self.personal_list.setStyleSheet(
            "QListWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QListWidget::indicator{width:16px;height:16px;}"
            "QListWidget::indicator:unchecked{border:1px solid #cbd5e1;border-radius:4px;background:#fff;}"
            "QListWidget::indicator:checked{border:1px solid #0f62fe;border-radius:4px;background:#0f62fe;}"
        )
        self.personal_list.itemChanged.connect(self._on_personal_toggle)

        tab1 = QWidget()
        t1l = QVBoxLayout(tab1)
        t1l.setContentsMargins(8, 8, 8, 8)
        t1l.addWidget(QLabel("PERSONAL"))
        t1l.addWidget(self.personal_list)
        self.stage_tabs.addTab(tab1, "PERSONAL")

        self.tasks_table = QTableWidget(0, 2)
        self.tasks_table.setHorizontalHeaderLabels(["TAREA", "RESPONSABLE"])
        self.tasks_table.verticalHeader().setVisible(False)
        self.tasks_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tasks_table.setStyleSheet(
            "QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QHeaderView::section{background:#f8fafc;border:1px solid #e2e8f5;padding:6px;font-weight:700;}"
        )
        self.tasks_table.itemChanged.connect(self._on_task_toggle)

        tab2 = QWidget()
        t2l = QVBoxLayout(tab2)
        t2l.setContentsMargins(8, 8, 8, 8)
        t2l.addWidget(QLabel("TAREAS"))
        t2l.addWidget(self.tasks_table)
        self.stage_tabs.addTab(tab2, "TAREAS")

        body.addWidget(self.detail_panel, 2)

    # ------------------------------------------------------------- Personal
    def _build_tab_personal(self) -> None:
        tab = QWidget()
        self.tabs.addTab(tab, "PERSONAL / PRODUCTIVIDAD")
        lay = QVBoxLayout(tab)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(10)

        kpis = QHBoxLayout()
        self.kpi_disp = _kpi_card("DISPONIBLES", "0")
        self.kpi_asig = _kpi_card("ASIGNADOS", "0")
        for c in (self.kpi_disp, self.kpi_asig):
            kpis.addWidget(c)
        lay.addLayout(kpis)

        kpi_turnos = QHBoxLayout()
        self.kpi_morning = _kpi_card("MAÑANA", "0")
        self.kpi_afternoon = _kpi_card("TARDE", "0")
        kpi_turnos.addWidget(self.kpi_morning, 1)
        kpi_turnos.addWidget(self.kpi_afternoon, 1)
        lay.addLayout(kpi_turnos)

        body = QHBoxLayout()
        lay.addLayout(body)

        self.tech_table = QTableWidget(0, 7)
        self.tech_table.setHorizontalHeaderLabels([
            "TECNICO", "TURNO", "AREA", "ESTADO", "OP ACTUAL",
            "TAREA", "PRODUCTIVIDAD",
        ])
        self.tech_table.verticalHeader().setVisible(False)
        self.tech_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tech_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tech_table.setStyleSheet("QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}")
        self.tech_table.itemSelectionChanged.connect(self._on_select_tech)
        body.addWidget(self.tech_table, 3)

        self.pie_card = _card()
        p_lay = QVBoxLayout(self.pie_card)
        p_lay.setContentsMargins(12, 12, 12, 12)
        p_lay.setSpacing(8)
        title_row = QHBoxLayout()
        p_title = QLabel("PERSONAL")
        p_title.setStyleSheet("font-weight:700;color:#0f172a;")
        title_row.addWidget(p_title)
        title_row.addStretch(1)
        self.pie_mode = QComboBox()
        self.pie_mode.addItems(["POR OP", "POR AREA"])
        self.pie_mode.setStyleSheet(
            "QComboBox{background:#fff;border:1px solid #e2e8f5;border-radius:8px;padding:2px 8px;}"
        )
        self.pie_mode.currentIndexChanged.connect(self._refresh_pie)
        title_row.addWidget(self.pie_mode)
        p_lay.addLayout(title_row)
        if QChartView and QPieSeries:
            self.pie_series = QPieSeries()
            self.pie_chart = QChart()
            self.pie_chart.addSeries(self.pie_series)
            legend = self.pie_chart.legend()
            if legend is not None:
                legend.setVisible(True)
                legend.setAlignment(Qt.AlignBottom)
            self.pie_chart.setBackgroundVisible(False)
            self.pie_view = QChartView(self.pie_chart)
            self.pie_view.setRenderHint(QPainter.Antialiasing)
            p_lay.addWidget(self.pie_view)
        else:
            self.pie_series = None
            self.pie_view = QLabel("QtCharts no disponible")
            self.pie_view.setStyleSheet("color:#64748b;")
            p_lay.addWidget(self.pie_view)
        body.addWidget(self.pie_card, 2)

    # ---------------------------------------------------------------- Logic
    def _refresh_orders(self) -> None:
        self.orders_table.setRowCount(0)
        for o in self.store.orders:
            row = self.orders_table.rowCount()
            self.orders_table.insertRow(row)
            vals = [
                o.op_number, o.date, o.state, f"{o.avance_pct:.0f}%",
            ]
            for c, v in enumerate(vals):
                self.orders_table.setItem(row, c, QTableWidgetItem(str(v)))
        self._fit_orders_table_width()
        self._refresh_kpis()
        # tech list removed from detail panel

    def _fit_orders_table_width(self) -> None:
        header = self.orders_table.horizontalHeader()
        header.resizeSections(QHeaderView.ResizeToContents)
        last = self.orders_table.columnCount() - 1
        if last >= 0:
            w = header.sectionSize(last)
            header.resizeSection(last, int(w + 60))
        width = header.length()
        width += self.orders_table.verticalHeader().width()
        width += self.orders_table.frameWidth() * 2
        if self.orders_table.verticalScrollBar().isVisible():
            width += self.orders_table.verticalScrollBar().sizeHint().width()
        self.orders_table.setFixedWidth(width)

    def _refresh_kpis(self) -> None:
        opens = [o for o in self.store.orders if o.state.upper() not in ("CERRADA", "CERRADO")]
        in_mat = [o for o in self.store.orders if o.material_files]
        in_personal = [o for o in self.store.orders if o.technicians]
        in_exec = [o for o in self.store.orders if 0 < o.avance_pct < 100]
        closed = [o for o in self.store.orders if o.avance_pct >= 100 or "CERR" in o.state.upper()]
        self.kpi_open.layout().itemAt(1).widget().setText(str(len(opens)))
        self.kpi_materials.layout().itemAt(1).widget().setText(str(len(in_mat)))
        self.kpi_personal.layout().itemAt(1).widget().setText(str(len(in_personal)))
        self.kpi_exec.layout().itemAt(1).widget().setText(str(len(in_exec)))
        self.kpi_closed.layout().itemAt(1).widget().setText(str(len(closed)))

    def _refresh_personal(self) -> None:
        self.tech_table.setRowCount(0)
        for idx, t in enumerate(self.store.techs):
            shift = _clean_text(t.shift)
            if not shift:
                shift = "MAÑANA" if idx % 2 == 0 else "TARDE"
                t.shift = shift
            area = _clean_text(t.area) or random.choice(AREAS_DEFAULT)
            t.area = area
            op_text = _clean_text(t.op)
            activity = _clean_text(t.activity)
            row = self.tech_table.rowCount()
            self.tech_table.insertRow(row)
            vals = [
                t.name,
                shift,
                area,
                t.status,
                op_text,
                activity or "--",
                f"{t.productivity:.0f}%",
            ]
            for c, v in enumerate(vals):
                item = QTableWidgetItem(str(v))
                if op_text or t.status == "ASIGNADO":
                    item.setBackground(QColor("#dcfce7"))
                    item.setForeground(QColor("#166534"))
                self.tech_table.setItem(row, c, item)
        self._refresh_personal_kpis()
        self._refresh_pie()

    def _refresh_personal_kpis(self) -> None:
        disp = sum(1 for t in self.store.techs if t.status == "DISPONIBLE")
        asig = sum(1 for t in self.store.techs if t.status == "ASIGNADO")
        morning = sum(1 for t in self.store.techs if _clean_text(t.shift).upper() == "MAÑANA")
        afternoon = sum(1 for t in self.store.techs if _clean_text(t.shift).upper() == "TARDE")
        self.kpi_disp.layout().itemAt(1).widget().setText(str(disp))
        self.kpi_asig.layout().itemAt(1).widget().setText(str(asig))
        self.kpi_morning.layout().itemAt(1).widget().setText(str(morning))
        self.kpi_afternoon.layout().itemAt(1).widget().setText(str(afternoon))

    def _refresh_pie(self) -> None:
        if not self.pie_series:
            return
        self.pie_series.clear()
        counts: dict[str, int] = {}
        mode = self.pie_mode.currentText() if hasattr(self, "pie_mode") else "POR OP"
        for t in self.store.techs:
            if mode == "POR AREA":
                key = _clean_text(t.area)
            else:
                key = _clean_text(t.op)
            if not key:
                continue
            counts[key] = counts.get(key, 0) + 1
        if not counts:
            self.pie_series.append("SIN ASIGNAR", 1)
            return
        for op, cnt in sorted(counts.items()):
            self.pie_series.append(op, cnt)

    def _on_select_order(self) -> None:
        rows = self.orders_table.selectionModel().selectedRows()
        if not rows:
            return
        op = self.orders_table.item(rows[0].row(), 0).text()
        order = self.store.get_order(op)
        if not order:
            return
        self._current_order = order
        notes = order.notes.strip() if order.notes else "--"
        self.lbl_summary.setText(notes)
        self._update_stage_tabs(order.stage_index)
        self._load_material_files(order.material_files)
        self._load_personal_checks(order.technicians)
        self._load_task_checks(order.task_checks if hasattr(order, "task_checks") else [])

    def _update_stage_tabs(self, active: int) -> None:
        if active >= self.stage_tabs.count():
            active = self.stage_tabs.count() - 1
        for i in range(self.stage_tabs.count()):
            color = QColor("#16a34a") if i <= active else QColor("#94a3b8")
            self.stage_tabs.tabBar().setTabTextColor(i, color)
        self._setting_stage_tab = True
        self.stage_tabs.setCurrentIndex(active)
        self._setting_stage_tab = False

    def _on_stage_tab_changed(self, idx: int) -> None:
        if getattr(self, "_setting_stage_tab", False):
            return
        if not getattr(self, "_current_order", None):
            return
        order = self._current_order
        if idx == 1 and not order.technicians:
            return
        if idx == 2 and not getattr(order, "task_checks", []):
            return
        self.store.advance_stage(order, idx)
        self._update_stage_tabs(order.stage_index)
        self._refresh_orders()

    def _load_material_files(self, selected: List[str]) -> None:
        self.material_files_list.blockSignals(True)
        self.material_files_list.clear()
        base = Path("data/produccion/MATERIALES")
        files = []
        if base.exists():
            for ext in ("*.xlsx", "*.xlsm", "*.xls"):
                files.extend(base.glob(ext))
        if not files:
            item = QListWidgetItem("(SIN ARCHIVOS EN MATERIALES)")
            item.setFlags(Qt.ItemIsEnabled)
            self.material_files_list.addItem(item)
        else:
            for f in sorted(files):
                item = QListWidgetItem(f.name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked if f.name in selected else Qt.Unchecked)
                item.setForeground(QColor("#16a34a") if f.name in selected else QColor("#0f172a"))
                item.setData(Qt.UserRole, str(f))
                self.material_files_list.addItem(item)
        self.material_files_list.blockSignals(False)

    def _load_personal_checks(self, selected: List[str]) -> None:
        self.personal_list.blockSignals(True)
        self.personal_list.clear()
        selected_set = {n.upper() for n in selected}
        for t in self.store.techs:
            item = QListWidgetItem(t.name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            is_checked = t.name.upper() in selected_set
            item.setCheckState(Qt.Checked if is_checked else Qt.Unchecked)
            item.setForeground(QColor("#16a34a") if is_checked else QColor("#0f172a"))
            self.personal_list.addItem(item)
        self.personal_list.blockSignals(False)

    def _load_task_checks(self, selected: List[str]) -> None:
        self.tasks_table.blockSignals(True)
        self.tasks_table.setRowCount(0)
        selected_set = {n.upper() for n in selected}
        assigned = getattr(self._current_order, "task_assignments", {})
        for task in TASKS:
            row = self.tasks_table.rowCount()
            self.tasks_table.insertRow(row)
            item = QTableWidgetItem(task)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            is_checked = task.upper() in selected_set
            item.setCheckState(Qt.Checked if is_checked else Qt.Unchecked)
            item.setForeground(QColor("#16a34a") if is_checked else QColor("#0f172a"))
            self.tasks_table.setItem(row, 0, item)

            cb = QComboBox()
            cb.setStyleSheet("QComboBox{background:#fff;border:1px solid #e2e8f5;border-radius:8px;padding:2px 6px;}")
            cb.addItem("--")
            for name in self._current_order.technicians:
                cb.addItem(name)
            assigned_name = assigned.get(task, "")
            if assigned_name:
                idx = cb.findText(assigned_name)
                if idx >= 0:
                    cb.setCurrentIndex(idx)
            cb.currentIndexChanged.connect(
                lambda _=None, t=task, box=cb: self._on_task_assign_changed(t, box)
            )
            self.tasks_table.setCellWidget(row, 1, cb)
        self.tasks_table.blockSignals(False)

    def _on_material_file_toggle(self, item: QListWidgetItem) -> None:
        if not getattr(self, "_current_order", None):
            return
        path = item.data(Qt.UserRole)
        if not path:
            return
        name = Path(path).name
        selected = set(self._current_order.material_files)
        if item.checkState() == Qt.Checked:
            selected.add(name)
            item.setForeground(QColor("#16a34a"))
        else:
            selected.discard(name)
            item.setForeground(QColor("#0f172a"))
        self._current_order.material_files = sorted(selected)
        self._update_materials_progress(self._current_order)
        self._refresh_kpis()

    def _on_personal_toggle(self, item: QListWidgetItem) -> None:
        if not getattr(self, "_current_order", None):
            return
        name = item.text().strip()
        if not name:
            return
        selected = set(self._current_order.technicians)
        if item.checkState() == Qt.Checked:
            selected.add(name)
            item.setForeground(QColor("#16a34a"))
        else:
            selected.discard(name)
            item.setForeground(QColor("#0f172a"))
        self._apply_personal_selection(self._current_order, sorted(selected))

    def _on_task_toggle(self, item: QTableWidgetItem) -> None:
        if not getattr(self, "_current_order", None):
            return
        if item.column() != 0:
            return
        name = item.text().strip()
        if not name:
            return
        selected = set(getattr(self._current_order, "task_checks", []))
        if item.checkState() == Qt.Checked:
            selected.add(name)
            item.setForeground(QColor("#16a34a"))
        else:
            selected.discard(name)
            item.setForeground(QColor("#0f172a"))
        self._current_order.task_checks = sorted(selected)
        self._update_tasks_progress(self._current_order)
        self._refresh_kpis()

    def _on_task_assign_changed(self, task_name: str, combo: QComboBox) -> None:
        if not getattr(self, "_current_order", None):
            return
        name = combo.currentText().strip()
        if name == "--":
            name = ""
        assigns = getattr(self._current_order, "task_assignments", {})
        if name:
            assigns[task_name] = name
        else:
            assigns.pop(task_name, None)
        self._current_order.task_assignments = assigns
        self._refresh_personal_task_assignments(self._current_order)

    def _on_material_file_open(self, item: QListWidgetItem) -> None:
        path = item.data(Qt.UserRole)
        if not path:
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def _update_materials_progress(self, order: ProductionOrder) -> None:
        base = Path("data/produccion/MATERIALES")
        total = 0
        if base.exists():
            for ext in ("*.xlsx", "*.xlsm", "*.xls"):
                total += len(list(base.glob(ext)))
        done = len(order.material_files)
        pct = 0.0 if total <= 0 else (done / total) * 100.0
        if order.stage_index == 0:
            order.avance_pct = round(pct, 0)
        # actualizar fila en tabla
        for r in range(self.orders_table.rowCount()):
            if self.orders_table.item(r, 0).text() == order.op_number:
                self.orders_table.setItem(r, 3, QTableWidgetItem(f"{order.avance_pct:.0f}%"))
                break

    def _update_tasks_progress(self, order: ProductionOrder) -> None:
        total = len(TASKS)
        done = len(getattr(order, "task_checks", []))
        pct = 0.0 if total == 0 else (done / total) * 100.0
        if done:
            if order.stage_index < 2:
                order.stage_index = 2
            order.avance_pct = max(order.avance_pct, round(pct, 0))
        else:
            if order.stage_index == 2:
                order.stage_index = 1 if order.technicians else 0
        self._update_stage_tabs(order.stage_index)
        for r in range(self.orders_table.rowCount()):
            if self.orders_table.item(r, 0).text() == order.op_number:
                self.orders_table.setItem(r, 3, QTableWidgetItem(f"{order.avance_pct:.0f}%"))
                break

    def _refresh_personal_task_assignments(self, order: ProductionOrder) -> None:
        tech_tasks: dict[str, List[str]] = {}
        assigns = getattr(order, "task_assignments", {})
        for task, name in assigns.items():
            if not name:
                continue
            tech_tasks.setdefault(name, []).append(task)
        for t in self.store.techs:
            if t.op == order.op_number:
                tasks = tech_tasks.get(t.name, [])
                t.activity = " / ".join(tasks) if tasks else ""
        self._refresh_personal()

    def _apply_personal_selection(self, order: ProductionOrder, techs: List[str]) -> None:
        order.technicians = techs
        assigns = getattr(order, "task_assignments", {})
        order.task_assignments = {k: v for k, v in assigns.items() if v in techs}
        for t in self.store.techs:
            if t.name in techs:
                t.status = "ASIGNADO"
                t.op = order.op_number
            elif t.op == order.op_number:
                t.op = ""
                t.status = "DISPONIBLE"
        if techs and order.stage_index < 1:
            self.store.advance_stage(order, 1)
        elif not techs and order.stage_index == 1:
            order.stage_index = 0
            order.avance_pct = 0.0
        self._update_stage_tabs(order.stage_index)
        self._load_personal_checks(order.technicians)
        self._load_task_checks(getattr(order, "task_checks", []))
        self._refresh_orders()
        self._refresh_personal()

    def _on_export_materials(self) -> None:
        rows = self.orders_table.selectionModel().selectedRows()
        if not rows:
            return
        op = self.orders_table.item(rows[0].row(), 0).text()
        order = self.store.get_order(op)
        if not order or not order.materials:
            return
        out_path, _ = QFileDialog.getSaveFileName(self, "Exportar materiales", f"{op}_materiales.xlsx", "Excel (*.xlsx)")
        if not out_path:
            return
        df = pd.DataFrame([{
            "ITEM": m.item,
            "CODIGO": m.code,
            "NOMBRE": m.name,
            "DESCRIPCION": m.description,
            "CANTIDAD": m.qty,
            "CATEGORIA": m.category,
            "STATUS": m.status,
        } for m in order.materials])
        df.to_excel(out_path, index=False)

    def _on_new_op(self) -> None:
        dlg = _NewOpDialog(self)
        res = dlg.exec()
        if res == QDialog.Accepted:
            op_id = f"OP-DEMO-{len(self.store.orders)+1:04d}"
            order = ProductionOrder(
                op_number=op_id,
                date=datetime.now().strftime("%Y-%m-%d"),
                state="ABIERTA",
                ref1=dlg.ed_cliente.text().strip(),
                ref2="",
                notes=dlg.ed_notes.text().strip(),
                categories=dlg.ed_cats.text().strip(),
            )
            self.store.add_order(order)
        else:
            # crear desde excel (random)
            if self.store.orders:
                base = random.choice(self.store.orders)
                op_id = f"{base.op_number}-DEMO"
                self.store.add_order(ProductionOrder(
                    op_number=op_id,
                    date=base.date,
                    state=base.state,
                    ref1=base.ref1,
                    ref2=base.ref2,
                    notes=base.notes,
                    categories=base.categories,
                ))
        self._refresh_orders()

    def _on_select_tech(self) -> None:
        pass
