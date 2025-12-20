from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, List

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QHeaderView,
)


@dataclass(frozen=True)
class _Entry:
    path: Path
    file_name: str
    proyecto: str
    n_cuartos: int
    updated_at: datetime


def _safe_meta(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    except Exception:
        return {}
    return {}


class BibliotecaCuartosDialog(QDialog):
    """Biblioteca de proyectos de cuartos fríos."""

    def __init__(
        self,
        *,
        parent=None,
        library_dir: Path,
        on_load: Callable[[Path], None],
        primary_qss: str,
        danger_qss: str,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("BIBLIOTECA DE PROYECTOS - CUARTOS FRÍOS")
        self.setModal(True)
        # permitir maximizar
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)

        self._library_dir = library_dir
        self._on_load = on_load
        self._primary_qss = primary_qss
        self._danger_qss = danger_qss

        self._entries: List[_Entry] = []
        self._build_ui()
        self._refresh()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        header = QFrame()
        header.setStyleSheet("background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;")
        hl = QVBoxLayout(header)
        hl.setContentsMargins(16, 14, 16, 14)
        hl.setSpacing(10)

        title = QLabel("BIBLIOTECA DE PROYECTOS - CUARTOS FRÍOS")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        hl.addWidget(title)

        search_row = QHBoxLayout()
        search_row.setSpacing(10)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Buscar proyecto / archivo...")
        self.search.textChanged.connect(self._apply_filter)
        btn_open = QPushButton("ABRIR CARPETA")
        btn_open.setStyleSheet(self._primary_qss)
        btn_open.clicked.connect(self._open_folder)
        search_row.addWidget(self.search, 1)
        search_row.addWidget(btn_open, 0)
        hl.addLayout(search_row)

        root.addWidget(header)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["ARCHIVO", "PROYECTO", "# CUARTOS", "MODIFICADO"])
        self.table.verticalHeader().setVisible(False)
        h = self.table.horizontalHeader()
        h.setStretchLastSection(True)
        for i in range(4):
            h.setSectionResizeMode(i, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setStyleSheet(
            "QTableWidget{background:#fff;border:1px solid #e2e8f5;}"
            "QHeaderView::section{background:#f8fafc;}"
            "QTableWidget::item:selected{background:#d6e4ff;color:#0f172a;}"
        )
        root.addWidget(self.table, 1)

        btns = QHBoxLayout()
        btns.setSpacing(8)
        self.btn_refresh = QPushButton("REFRESCAR"); self.btn_refresh.setStyleSheet(self._primary_qss); self.btn_refresh.clicked.connect(self._refresh)
        self.btn_load = QPushButton("CARGAR"); self.btn_load.setStyleSheet(self._primary_qss); self.btn_load.clicked.connect(self._load_selected)
        self.btn_duplicate = QPushButton("DUPLICAR"); self.btn_duplicate.setStyleSheet(self._primary_qss); self.btn_duplicate.clicked.connect(self._duplicate_selected)
        # Botón rojo pastel (coherente con Carga Eléctrica)
        self.btn_delete = QPushButton("ELIMINAR")
        self.btn_delete.setStyleSheet(
            "QPushButton{background:#fdecec;color:#c53030;font-weight:700;border:none;border-radius:8px;padding:8px 12px;}"
            "QPushButton:hover{background:#fbd5d5;}"
        )
        self.btn_delete.clicked.connect(self._delete_selected)
        btns.addWidget(self.btn_refresh)
        btns.addStretch(1)
        btns.addWidget(self.btn_load)
        btns.addWidget(self.btn_duplicate)
        btns.addWidget(self.btn_delete)
        root.addLayout(btns)

    # ------------------------------------------------------------------
    def _refresh(self) -> None:
        self._library_dir.mkdir(parents=True, exist_ok=True)
        entries: List[_Entry] = []
        for path in sorted(self._library_dir.glob("*.json")):
            meta = _safe_meta(path)
            proyecto = meta.get("proyecto", "") or meta.get("nombre_proyecto", "")
            n_cuartos = len(meta.get("inputs", []))
            updated = datetime.fromtimestamp(path.stat().st_mtime)
            entries.append(_Entry(path, path.name, proyecto, n_cuartos, updated))
        self._all_entries = entries
        self._apply_filter()

    def _apply_filter(self) -> None:
        text = self.search.text().strip().lower()
        if not text:
            self._entries = self._all_entries
        else:
            self._entries = [
                e for e in self._all_entries
                if text in e.file_name.lower() or text in (e.proyecto or "").lower()
            ]
        self._populate()

    def _populate(self) -> None:
        self.table.setRowCount(len(self._entries))
        for r, e in enumerate(self._entries):
            self.table.setItem(r, 0, QTableWidgetItem(e.file_name))
            self.table.setItem(r, 1, QTableWidgetItem(e.proyecto))
            self.table.setItem(r, 2, QTableWidgetItem(str(e.n_cuartos)))
            self.table.setItem(r, 3, QTableWidgetItem(e.updated_at.strftime("%Y-%m-%d %H:%M")))
        if self._entries:
            self.table.selectRow(0)

    def _selected_entry(self) -> _Entry | None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self._entries):
            return None
        return self._entries[row]

    def _load_selected(self) -> None:
        entry = self._selected_entry()
        if not entry:
            return
        self._on_load(entry.path)
        self.accept()

    def _delete_selected(self) -> None:
        entry = self._selected_entry()
        if not entry:
            return
        if QMessageBox.question(self, "Eliminar", f"¿Eliminar {entry.file_name}?") != QMessageBox.Yes:
            return
        try:
            entry.path.unlink()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return
        self._refresh()

    def _open_folder(self) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._library_dir)))

    def _duplicate_selected(self) -> None:
        entry = self._selected_entry()
        if not entry:
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_path = entry.path.with_name(f"{entry.path.stem}_copy_{ts}{entry.path.suffix}")
        try:
            new_path.write_text(entry.path.read_text(encoding="utf-8"), encoding="utf-8")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return
        self._refresh()
