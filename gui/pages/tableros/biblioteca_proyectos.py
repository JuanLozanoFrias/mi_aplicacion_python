from __future__ import annotations

import json
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional

import pandas as pd
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
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from logic.tableros.step4_engine import Step4Engine
from logic.tableros.export_step4 import (
    _ensure_item_col,
    _inject_marca,
    _collect_totales,
    _load_inventario_map,
    _build_brand_map,
)
from logic.programacion_loader import load_programacion_snapshot


@dataclass(frozen=True)
class _ProjectEntry:
    path: Path
    file_name: str
    project_name: str
    city: str
    norma: str
    updated_at: datetime


def _safe_read_globs(path: Path) -> dict:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            globs = data.get("globs")
            return globs if isinstance(globs, dict) else {}
    except Exception:
        return {}
    return {}


def _as_str(v: object) -> str:
    if v is None:
        return ""
    try:
        return str(v).strip()
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


class BibliotecaProyectosDialog(QDialog):
    """
    Biblioteca local para proyectos de Tableros Eléctricos.
    Lista los archivos `.ecalc.json` almacenados en `data/proyectos/tableros/`.
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
        self.setWindowTitle("BIBLIOTECA DE PROYECTOS")
        self.setModal(True)

        self._library_dir = library_dir
        self._on_load = on_load
        self._primary_qss = primary_qss
        self._danger_qss = danger_qss

        self._all_entries: List[_ProjectEntry] = []
        self._entries: List[_ProjectEntry] = []  # entries visibles (filtradas)

        self._build_ui()
        self._refresh()

    # ---------------------------------------------------------------- UI
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        header = QFrame()
        header.setStyleSheet("background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;")
        hl = QVBoxLayout(header)
        hl.setContentsMargins(16, 14, 16, 14)
        hl.setSpacing(10)

        title = QLabel("BIBLIOTECA DE PROYECTOS")
        title.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        title.setAlignment(Qt.AlignCenter)
        hl.addWidget(title)

        search_row = QHBoxLayout()
        search_row.setSpacing(10)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Buscar por proyecto / ciudad / norma / archivo…")
        self.search.textChanged.connect(self._apply_filter)
        btn_open_folder = QPushButton("ABRIR CARPETA")
        btn_open_folder.setStyleSheet(self._primary_qss)
        btn_open_folder.clicked.connect(self._open_folder)
        search_row.addWidget(self.search, 1)
        search_row.addWidget(btn_open_folder, 0)
        hl.addLayout(search_row)
        root.addWidget(header)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ARCHIVO", "PROYECTO", "CIUDAD", "NORMA", "MODIFICADO"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.itemDoubleClicked.connect(lambda _it: self._load_selected())
        self.table.setStyleSheet(
            "QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;}"
            "QHeaderView::section{background:#f1f5ff;color:#0f172a;font-weight:800;padding:6px;border:none;}"
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        try:
            self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        except Exception:
            pass
        root.addWidget(self.table, 1)

        # Botones
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

        self.btn_consol = QPushButton("CONSOLIDAR")
        self.btn_consol.setStyleSheet(self._primary_qss)
        self.btn_consol.clicked.connect(self._consolidate_selected)

        self.btn_dup = QPushButton("DUPLICAR")
        self.btn_dup.setStyleSheet(self._primary_qss)
        self.btn_dup.clicked.connect(self._duplicate_selected)

        self.btn_delete = QPushButton("ELIMINAR")
        self.btn_delete.setStyleSheet(self._danger_qss)
        self.btn_delete.clicked.connect(self._delete_selected)

        al.addWidget(self.btn_refresh)
        al.addStretch(1)
        al.addWidget(self.btn_load)
        al.addWidget(self.btn_consol)
        al.addWidget(self.btn_dup)
        al.addWidget(self.btn_delete)
        root.addWidget(actions)

        self.resize(980, 560)

    # ------------------------------------------------------------ Helpers
    def _selected_entry(self) -> Optional[_ProjectEntry]:
        sel = self.table.selectionModel()
        if sel is None or not sel.hasSelection():
            return None
        row = sel.selectedRows()[0].row()
        if 0 <= row < len(self._entries):
            return self._entries[row]
        return None

    def _selected_entries(self) -> List[_ProjectEntry]:
        sel = self.table.selectionModel()
        if sel is None or not sel.hasSelection():
            return []
        rows = [i.row() for i in sel.selectedRows()]
        out = []
        for r in rows:
            if 0 <= r < len(self._entries):
                out.append(self._entries[r])
        return out

    def _refresh(self) -> None:
        self._library_dir.mkdir(parents=True, exist_ok=True)

        entries: List[_ProjectEntry] = []
        for p in sorted(self._library_dir.glob("*.ecalc.json"), key=lambda x: x.stat().st_mtime, reverse=True):
            globs = _safe_read_globs(p)
            project_name = _as_str(globs.get("nombre_proyecto") or globs.get("proyecto") or globs.get("project_name"))
            city = _as_str(globs.get("ciudad") or globs.get("city"))
            norma = _as_str(globs.get("norma_ap") or globs.get("norma") or globs.get("norma aplicable")).upper()
            try:
                updated_at = datetime.fromtimestamp(p.stat().st_mtime)
            except Exception:
                updated_at = datetime.now()

            entries.append(
                _ProjectEntry(
                    path=p,
                    file_name=p.name,
                    project_name=project_name,
                    city=city,
                    norma=norma,
                    updated_at=updated_at,
                )
            )

        self._all_entries = entries
        self._apply_filter(self.search.text() if hasattr(self, "search") else "")

    def _apply_filter(self, text: str) -> None:
        q = (text or "").strip().upper()
        all_entries = self._all_entries
        filtered = []
        if not q:
            filtered = all_entries
        else:
            for e in all_entries:
                blob = f"{e.file_name} {e.project_name} {e.city} {e.norma}".upper()
                if q in blob:
                    filtered.append(e)

        self.table.setRowCount(len(filtered))
        for r, e in enumerate(filtered):
            self.table.setItem(r, 0, QTableWidgetItem(e.file_name))
            self.table.setItem(r, 1, QTableWidgetItem(e.project_name))
            self.table.setItem(r, 2, QTableWidgetItem(e.city))
            self.table.setItem(r, 3, QTableWidgetItem(e.norma))
            self.table.setItem(r, 4, QTableWidgetItem(e.updated_at.strftime("%Y-%m-%d %H:%M")))

        for c in range(self.table.columnCount()):
            self.table.resizeColumnToContents(c)

        # Mantener el mapping para acciones (solo filas visibles)
        self._entries = filtered

    def _open_folder(self) -> None:
        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._library_dir.resolve())))
        except Exception:
            QMessageBox.information(self, "Biblioteca", f"Carpeta:\n{self._library_dir}")

    # ----------------------------------------------------------- Actions
    def _load_selected(self) -> None:
        entry = self._selected_entry()
        if entry is None:
            QMessageBox.information(self, "Biblioteca", "Selecciona un proyecto para cargar.")
            return
        try:
            self._on_load(entry.path)
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "Biblioteca", f"No se pudo cargar el proyecto:\n{e}")

    def _duplicate_selected(self) -> None:
        entry = self._selected_entry()
        if entry is None:
            QMessageBox.information(self, "Biblioteca", "Selecciona un proyecto para duplicar.")
            return
        try:
            src = entry.path
            dst = _unique_path(src.with_name(f"{src.stem}_COPIA{src.suffix}"))
            shutil.copy2(src, dst)
            self._refresh()
            QMessageBox.information(self, "Biblioteca", f"Duplicado:\n{dst.name}")
        except Exception as e:
            QMessageBox.critical(self, "Biblioteca", f"No se pudo duplicar:\n{e}")

    def _delete_selected(self) -> None:
        entry = self._selected_entry()
        if entry is None:
            QMessageBox.information(self, "Biblioteca", "Selecciona un proyecto para eliminar.")
            return
        resp = QMessageBox.question(
            self,
            "Eliminar",
            f"¿Eliminar este proyecto?\n\n{entry.file_name}",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return
        try:
            entry.path.unlink(missing_ok=True)
            self._refresh()
        except Exception as e:
            QMessageBox.critical(self, "Biblioteca", f"No se pudo eliminar:\n{e}")

    # ----------------------------------------------------------- Consolidado
    def _consolidate_selected(self) -> None:
        entries = self._selected_entries()
        if not entries:
            QMessageBox.information(self, "Consolidar", "Selecciona uno o varios proyectos.")
            return
        try:
            # usar Excel base por defecto (igual que en export)
            base_dir = Path(__file__).resolve().parents[3] / "data"
            candidates = [
                base_dir / "BD_TABLEROS.xlsx",
                base_dir / "basedatos.xlsx",
                base_dir / "CUADRO DE CARGAS -INDIVIDUAL JUAN LOZANO.xlsx",
            ]
            basedatos = next((c for c in candidates if c.exists()), candidates[0])

            engine = Step4Engine(basedatos)
            brand_map = _build_brand_map(basedatos)
            inv_map = _load_inventario_map(base_dir / "inventarios.xlsx")

            all_rows = []
            for e in entries:
                step2, globs = load_programacion_snapshot(e.path)
                res = engine.calcular(step2, globs)
                tables = res.get("tables_compresores", []) or []
                otros_rows = _ensure_item_col(res.get("otros_rows", []) or [])
                marca_global = str(globs.get("marca_elem", "")).strip()
                for t in tables:
                    filas = _inject_marca(t.rows, marca_global, brand_map)
                    all_rows.extend(filas)
                if otros_rows:
                    otros_fixed = _inject_marca(otros_rows, marca_global, brand_map)
                    all_rows.extend(otros_fixed)

            # totales consolidados
            marca_global = "SIN MARCA"
            tot_rows = _collect_totales(all_rows, marca_global)

            # construir filas con inventario
            rows_view = []
            for r in tot_rows:
                if len(r) < 10:
                    continue
                code = str(r[0])
                qty = float(r[9]) if r[9] not in ("", None) else 0.0
                inv = inv_map.get(code, {"disp": 0.0, "solic": 0.0})
                disp = inv.get("disp", 0.0)
                solic = inv.get("solic", 0.0)
                stock = disp + solic
                falt = max(qty - stock, 0.0)
                rows_view.append(r[:10] + [disp, solic, stock, falt])

            dlg = QDialog(self)
            dlg.setWindowTitle("Consolidado de materiales")
            lay = QVBoxLayout(dlg)
            table = QTableWidget(0, 14)
            headers = ["CÓDIGO", "MODELO", "NOMBRE", "DESCRIPCIÓN", "MARCA",
                       "C 240 V (kA)", "C 480 V (kA)", "REFERENCIA", "TORQUE", "CANTIDAD",
                       "DISPONIBLE", "SOLICITADO", "STOCK", "FALTANTE"]
            table.setHorizontalHeaderLabels(headers)
            table.setSelectionMode(QAbstractItemView.NoSelection)
            table.verticalHeader().setVisible(False)
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setAlternatingRowColors(True)
            table.setRowCount(len(rows_view))
            for i, rr in enumerate(rows_view):
                for j, val in enumerate(rr):
                    table.setItem(i, j, QTableWidgetItem(str(val)))
            table.resizeColumnsToContents()
            lay.addWidget(table)

            def _export_rows() -> None:
                path, _ = QFileDialog.getSaveFileName(
                    self, "Guardar consolidado",
                    str(self._library_dir / "consolidado_tableros.xlsx"),
                    "Excel (*.xlsx)"
                )
                if not path:
                    return
                try:
                    df = pd.DataFrame(rows_view, columns=[
                        "CODIGO", "MODELO", "NOMBRE", "DESCRIPCION", "MARCA",
                        "C240", "C480", "REFERENCIA", "TORQUE", "CANTIDAD",
                        "DISPONIBLE", "SOLICITADO", "STOCK", "FALTANTE"
                    ])
                    meta = pd.DataFrame([{
                        "ARCHIVO": e.file_name, "PROYECTO": e.project_name,
                        "CIUDAD": e.city, "NORMA": e.norma, "MODIFICADO": e.updated_at
                    } for e in entries])
                    with pd.ExcelWriter(path, engine="openpyxl") as writer:
                        df.to_excel(writer, sheet_name="CONSOLIDADO", index=False)
                        meta.to_excel(writer, sheet_name="PROYECTOS", index=False)
                    QMessageBox.information(self, "Exportar", f"Consolidado guardado en:\n{path}")
                except Exception as e:
                    QMessageBox.critical(self, "Exportar", f"No se pudo exportar:\n{e}")

            buttons_row = QHBoxLayout()
            btn_export = QPushButton("EXPORTAR CONSOLIDADO")
            btn_export.setStyleSheet(self._primary_qss)
            btn_export.clicked.connect(_export_rows)
            buttons_row.addStretch(1)
            buttons_row.addWidget(btn_export)
            lay.addLayout(buttons_row)
            dlg.resize(1100, 520)
            dlg.exec()

            # guardar data para export
            self._last_consolidated_rows = rows_view
            self._last_consolidated_projects = entries
        except Exception as e:
            QMessageBox.critical(self, "Consolidar", f"No se pudo consolidar:\n{e}")
