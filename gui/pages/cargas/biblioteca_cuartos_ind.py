from __future__ import annotations

import json
from pathlib import Path
from datetime import datetime
from typing import Callable, Optional

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableView,
    QLineEdit,
    QMessageBox,
    QFileDialog,
)


PROJECTS_DIR = Path("data/proyectos/cuartos_industriales")
PROJECTS_DIR.mkdir(parents=True, exist_ok=True)


class _ProjectsModel(QAbstractTableModel):
    def __init__(self, rows):
        super().__init__()
        self.rows = rows
        self.headers = ["ARCHIVO", "PROYECTO", "# CUARTOS", "MODIFICADO"]

    def rowCount(self, parent=QModelIndex()):
        return len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(self.headers)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row = self.rows[index.row()]
        if role in (Qt.DisplayRole, Qt.EditRole):
            return row[index.column()]
        if role == Qt.TextAlignmentRole:
            return Qt.AlignCenter
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.headers[section]
        return None


class BibliotecaCuartosIndDialog(QDialog):
    def __init__(self, parent=None, on_load: Optional[Callable[[dict], None]] = None):
        super().__init__(parent)
        self.setWindowTitle("BIBLIOTECA DE PROYECTOS - CUARTOS INDUSTRIALES")
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        self.on_load = on_load
        self.rows = []
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        self.setMinimumWidth(900)
        self.setMinimumHeight(600)
        lay = QVBoxLayout(self)

        title = QLabel(self.windowTitle())
        title.setStyleSheet("font-size:18px;font-weight:800;")
        lay.addWidget(title)

        search = QLineEdit()
        search.setPlaceholderText("Buscar proyecto / archivo...")
        search.textChanged.connect(self._filter)
        lay.addWidget(search)
        self.search = search

        self.table = QTableView()
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setStyleSheet(
            "QTableView::item:selected{background:#cbe4ff;color:#0f172a;}"
        )
        lay.addWidget(self.table, 1)

        btns = QHBoxLayout()
        self.btn_open_folder = QPushButton("ABRIR CARPETA")
        self.btn_open_folder.clicked.connect(lambda: QFileDialog.getExistingDirectory(self, "Abrir carpeta", str(PROJECTS_DIR)))
        self.btn_refresh = QPushButton("REFRESCAR")
        self.btn_refresh.clicked.connect(self.refresh)
        self.btn_load = QPushButton("CARGAR")
        self.btn_load.clicked.connect(self._load_selected)
        self.btn_delete = QPushButton("ELIMINAR")
        self.btn_delete.clicked.connect(self._delete_selected)
        for b in (self.btn_open_folder, self.btn_refresh, self.btn_load, self.btn_delete):
            btns.addWidget(b)
        lay.addLayout(btns)

    def _filter(self, text: str):
        text = text.lower()
        filtered = [
            r for r in self.rows if text in r[0].lower() or text in str(r[1]).lower()
        ]
        self._set_model(filtered)

    def _set_model(self, rows):
        model = _ProjectsModel(rows)
        self.table.setModel(model)
        self.table.resizeColumnsToContents()

    def refresh(self):
        PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
        rows = []
        for f in PROJECTS_DIR.glob("*.json"):
            try:
                data = json.load(open(f, encoding="utf-8"))
                proj = data.get("project_name", f.name)
                n = len(data.get("rooms", []))
                ts = datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
                rows.append([f.name, proj, n, ts])
            except Exception:
                continue
        self.rows = sorted(rows, key=lambda r: r[0])
        self._set_model(self.rows)

    def _selected_file(self) -> Optional[Path]:
        idx = self.table.currentIndex()
        if not idx.isValid():
            return None
        fname = self.table.model().data(self.table.model().index(idx.row(), 0))
        return PROJECTS_DIR / fname

    def _load_selected(self):
        p = self._selected_file()
        if not p:
            return
        try:
            data = json.load(open(p, encoding="utf-8"))
            if self.on_load:
                self.on_load(data)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cargar: {e}")

    def _delete_selected(self):
        p = self._selected_file()
        if not p:
            return
        if QMessageBox.question(self, "Eliminar", f"¿Eliminar {p.name}?") == QMessageBox.Yes:
            try:
                p.unlink()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
        self.refresh()

