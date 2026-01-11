from __future__ import annotations

import json
import shutil
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
)


@dataclass(frozen=True)
class _LegendEntry:
    path: Path
    file_name: str
    proyecto: str
    bt_ramales: int
    mt_ramales: int
    updated_at: datetime


def _safe_read_meta(path: Path) -> dict:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        return {}
    return {}


def _as_str(val: object) -> str:
    if val is None:
        return ""
    try:
        return str(val).strip()
    except Exception:
        return ""


def _unique_path(p: Path) -> Path:
    if not p.exists():
        return p
    for i in range(2, 10_000):
        cand = p.with_name(f"{p.stem}_{i}{p.suffix}")
        if not cand.exists():
            return cand
    return p


class BibliotecaLegendDialog(QDialog):
    """
    Biblioteca de proyectos LEGEND.
    Lista archivos .json en data/proyectos/legend/.
    """

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
        self.setWindowTitle("BIBLIOTECA LEGEND")
        self.setModal(True)

        self._library_dir = library_dir
        self._on_load = on_load
        self._primary_qss = primary_qss
        self._danger_qss = danger_qss

        self._all_entries: List[_LegendEntry] = []
        self._entries: List[_LegendEntry] = []

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

        title = QLabel("BIBLIOTECA DE PROYECTOS - LEGEND")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        hl.addWidget(title)

        search_row = QHBoxLayout()
        search_row.setSpacing(10)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Buscar proyecto / archivo...")
        self.search.textChanged.connect(self._apply_filter)
        btn_open_folder = QPushButton("ABRIR CARPETA")
        btn_open_folder.setStyleSheet(self._primary_qss)
        btn_open_folder.clicked.connect(self._open_folder)
        search_row.addWidget(self.search, 1)
        search_row.addWidget(btn_open_folder, 0)
        hl.addLayout(search_row)
        root.addWidget(header)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ARCHIVO", "PROYECTO", "BT", "MT", "MODIFICADO"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.itemDoubleClicked.connect(lambda _it: self._load_selected())
        self.table.setStyleSheet(
            "QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;}"
            "QHeaderView::section{background:#f1f5ff;color:#0f172a;font-weight:800;padding:6px;border:none;}"
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        root.addWidget(self.table, 1)

        actions = QFrame()
        actions.setStyleSheet("background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;")
        al = QHBoxLayout(actions)
        al.setContentsMargins(16, 12, 16, 12)
        al.setSpacing(10)

        self.btn_refresh = QPushButton("REFRESCAR")
        self.btn_refresh.setStyleSheet(self._primary_qss)
        self.btn_refresh.clicked.connect(self._refresh)

        self.btn_load = QPushButton("CARGAR")
        self.btn_load.setStyleSheet(self._primary_qss)
        self.btn_load.clicked.connect(self._load_selected)

        self.btn_dup = QPushButton("DUPLICAR")
        self.btn_dup.setStyleSheet(self._primary_qss)
        self.btn_dup.clicked.connect(self._duplicate_selected)

        self.btn_delete = QPushButton("ELIMINAR")
        self.btn_delete.setStyleSheet(self._danger_qss)
        self.btn_delete.clicked.connect(self._delete_selected)

        al.addWidget(self.btn_refresh)
        al.addStretch(1)
        al.addWidget(self.btn_load)
        al.addWidget(self.btn_dup)
        al.addWidget(self.btn_delete)
        root.addWidget(actions)

        self.resize(820, 520)

    def _refresh(self) -> None:
        self._library_dir.mkdir(parents=True, exist_ok=True)
        entries: List[_LegendEntry] = []
        for p in sorted(self._library_dir.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True):
            meta = _safe_read_meta(p)
            specs = meta.get("specs", {}) if isinstance(meta, dict) else {}
            proyecto = _as_str(specs.get("proyecto") or meta.get("proyecto"))
            try:
                bt_ramales = int(meta.get("bt_ramales", 0) or 0)
            except Exception:
                bt_ramales = 0
            try:
                mt_ramales = int(meta.get("mt_ramales", 0) or 0)
            except Exception:
                mt_ramales = 0
            try:
                updated = datetime.fromtimestamp(p.stat().st_mtime)
            except Exception:
                updated = datetime.now()
            entries.append(
                _LegendEntry(
                    path=p,
                    file_name=p.name,
                    proyecto=proyecto,
                    bt_ramales=bt_ramales,
                    mt_ramales=mt_ramales,
                    updated_at=updated,
                )
            )
        self._all_entries = entries
        self._apply_filter(self.search.text() if hasattr(self, "search") else "")

    def _apply_filter(self, text: str) -> None:
        q = (text or "").strip().upper()
        base = self._all_entries
        filtered = []
        if not q:
            filtered = base
        else:
            for e in base:
                blob = f"{e.file_name} {e.proyecto}".upper()
                if q in blob:
                    filtered.append(e)
        self._entries = filtered
        self.table.setRowCount(len(filtered))
        for r, e in enumerate(filtered):
            self.table.setItem(r, 0, QTableWidgetItem(e.file_name))
            self.table.setItem(r, 1, QTableWidgetItem(e.proyecto))
            self.table.setItem(r, 2, QTableWidgetItem(str(e.bt_ramales)))
            self.table.setItem(r, 3, QTableWidgetItem(str(e.mt_ramales)))
            self.table.setItem(r, 4, QTableWidgetItem(e.updated_at.strftime("%Y-%m-%d %H:%M")))

    def _selected_entry(self) -> _LegendEntry | None:
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

    def _duplicate_selected(self) -> None:
        entry = self._selected_entry()
        if not entry:
            return
        target = _unique_path(entry.path.with_name(f"{entry.path.stem}_COPIA.json"))
        try:
            shutil.copy(entry.path, target)
            QMessageBox.information(self, "LEGEND", f"Proyecto duplicado:\n{target.name}")
            self._refresh()
        except Exception as exc:
            QMessageBox.critical(self, "LEGEND", f"No se pudo duplicar:\n{exc}")

    def _delete_selected(self) -> None:
        entry = self._selected_entry()
        if not entry:
            return
        resp = QMessageBox.question(self, "Eliminar", f"Eliminar {entry.file_name}?")
        if resp != QMessageBox.Yes:
            return
        try:
            entry.path.unlink(missing_ok=True)
            self._refresh()
        except Exception as exc:
            QMessageBox.critical(self, "LEGEND", f"No se pudo eliminar:\n{exc}")

    def _open_folder(self) -> None:
        self._library_dir.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._library_dir)))
