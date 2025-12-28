from __future__ import annotations

from pathlib import Path
from typing import Dict, List

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication, QFontMetrics
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QGroupBox,
    QListWidget,
    QListWidgetItem,
    QScrollArea,
    QHBoxLayout,
    QAbstractScrollArea,
    QSizePolicy,
)

try:
    from logic.legend_jd import LegendJDService
except Exception as exc:
    LegendJDService = None  # type: ignore
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None


class LegendPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.service = LegendJDService() if LegendJDService else None
        self.data: Dict[str, object] = {}

        outer = QVBoxLayout(self)
        outer.setContentsMargins(12, 12, 12, 12)
        outer.setSpacing(8)

        header = QHBoxLayout()
        header.setSpacing(8)
        title = QLabel("LEGEND (solo lectura)")
        title.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        header.addWidget(title, 1)
        self.btn_refresh = QPushButton("Refrescar")
        self.btn_refresh.clicked.connect(self.refresh)
        header.addWidget(self.btn_refresh, 0)
        outer.addLayout(header)

        # Scroll area para el contenido
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        content = QWidget()
        content.setMinimumWidth(0)
        from PySide6.QtWidgets import QSizePolicy
        sp = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        content.setSizePolicy(sp)
        scroll.setWidget(content)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(6)

        # Folder + acciones
        self._folder_full_path = ""
        self.lbl_folder = QLabel("Carpeta: --")
        self.lbl_folder.setToolTip("")
        self.btn_copy_folder = QPushButton("Copiar ruta")
        self.btn_copy_folder.clicked.connect(self._copy_folder)
        self.btn_files = QPushButton("Ver archivos")
        self.btn_files.clicked.connect(self._show_files)
        top_folder = QHBoxLayout()
        top_folder.addWidget(self.lbl_folder, 1)
        top_folder.addWidget(self.btn_copy_folder, 0)
        top_folder.addWidget(self.btn_files, 0)
        layout.addLayout(top_folder)

        self.lbl_summary = QLabel("Resumen: --")
        layout.addWidget(self.lbl_summary)

        # Secciones
        self.tbl_config = self._make_table(["Campo", "Valor"])
        layout.addWidget(self._wrap_group("Config", self.tbl_config))

        self.tbl_equipos = self._make_table(["Equipo", "BTU/h ft"])
        layout.addWidget(self._wrap_group("Equipos", self.tbl_equipos))

        self.list_usos_bt = self._make_list()
        layout.addWidget(self._wrap_group("Usos BT", self.list_usos_bt))

        self.list_usos_mt = self._make_list()
        layout.addWidget(self._wrap_group("Usos MT", self.list_usos_mt))

        self.tbl_variadores = self._make_table(["Modelo", "Potencia"])
        layout.addWidget(self._wrap_group("Variadores FC102", self.tbl_variadores))

        self.tbl_wcr = self._make_table(["Modelo", "Capacidad"])
        layout.addWidget(self._wrap_group("WCR", self.tbl_wcr))

        self.tbl_plantillas = {
            "Rack Loop": self._make_table(["Qty", "Descripción"]),
            "Minisistema": self._make_table(["Qty", "Descripción"]),
            "Rack Americano": self._make_table(["Qty", "Descripción"]),
        }
        for name, table in self.tbl_plantillas.items():
            layout.addWidget(self._wrap_group(f"Plantilla: {name}", table))

        layout.addStretch(1)
        outer.addWidget(scroll, 1)

        if _IMPORT_ERROR:
            QMessageBox.critical(self, "Legend", f"No se pudo importar LegendJDService:\n{_IMPORT_ERROR}")
        else:
            self.refresh()

    def _make_table(self, headers: List[str]) -> QTableWidget:
        tbl = QTableWidget(0, len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        tbl.setSelectionMode(QTableWidget.NoSelection)
        tbl.setAlternatingRowColors(True)
        tbl.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tbl.verticalHeader().setDefaultSectionSize(22)
        return tbl

    def _make_list(self) -> QListWidget:
        lst = QListWidget()
        lst.setAlternatingRowColors(True)
        lst.setSelectionMode(QListWidget.NoSelection)
        lst.setMaximumHeight(150)
        return lst

    def _wrap_group(self, title: str, widget: QWidget) -> QGroupBox:
        box = QGroupBox(title)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.addWidget(widget)
        return box

    def _copy_folder(self) -> None:
        folder = self.data.get("sources", {}).get("folder") if isinstance(self.data, dict) else ""
        if not folder:
            return
        cb = QGuiApplication.clipboard()
        cb.setText(str(folder))

    def _show_files(self) -> None:
        files = self.data.get("sources", {}).get("files_found", []) if isinstance(self.data, dict) else []
        text = "\n".join(files) if files else "(ninguno)"
        QMessageBox.information(self, "Archivos Legend", text)

    def refresh(self) -> None:
        if not self.service:
            return
        try:
            self.data = self.service.load_all()
        except Exception as exc:
            QMessageBox.critical(self, "Legend", f"No se pudo cargar LEGEND:\n{exc}")
            return

        cfg = self.data.get("config")
        usos = self.data.get("usos", {}) or {}
        equipos = self.data.get("equipos", []) or []
        variadores = self.data.get("variadores", []) or []
        wcr = self.data.get("wcr", []) or []
        plantillas = self.data.get("plantillas", {}) or {}
        sources = self.data.get("sources", {}) or {}

        folder = sources.get("folder", "data/LEGEND")
        self.lbl_folder.setText(f"Carpeta: {folder}")
        self.lbl_folder.setToolTip(str(folder))
        files = sources.get("files_found", [])
        self._folder_full_path = str(folder)
        self.lbl_folder.setToolTip(self._folder_full_path)
        self._update_folder_elide()

        resumen = (
            f"Archivos: {len(files)} (ver lista) | "
            f"Equipos: {len(equipos)} | Usos BT: {len(usos.get('BT', []))} | Usos MT: {len(usos.get('MT', []))} | "
            f"Variadores: {len(variadores)} | WCR: {len(wcr)}"
        )
        self.lbl_summary.setText(resumen)

        self._fill_config(cfg)
        self._fill_table(self.tbl_equipos, [(e.equipo, e.btu_hr_ft) for e in equipos], numeric_cols={1}, stretch_cols={0}, resize_cols={1})
        self._fill_list(self.list_usos_bt, usos.get("BT", []))
        self._fill_list(self.list_usos_mt, usos.get("MT", []))
        self._fill_table(self.tbl_variadores, [(v.modelo, v.potencia) for v in variadores], numeric_cols={1}, stretch_cols={0}, resize_cols={1})
        self._fill_table(self.tbl_wcr, [(w.modelo, w.capacidad) for w in wcr], numeric_cols={1}, stretch_cols={0}, resize_cols={1})
        for name, tbl in self.tbl_plantillas.items():
            items = plantillas.get(name, []) or []
            self._fill_table(tbl, [(p.qty, p.descripcion) for p in items], numeric_cols={0}, stretch_cols={1}, resize_cols={0})

    def _fill_config(self, cfg) -> None:
        rows = []
        if cfg:
            for field in (
                "proyecto",
                "ciudad",
                "tipo_sistema",
                "refrigerante",
                "tcond",
                "tevap_bt",
                "tevap_mt",
                "marca_compresores",
                "tipo_instalacion",
                "factor_seguridad",
            ):
                rows.append((field, getattr(cfg, field, None)))
            if getattr(cfg, "extras", None):
                for k, v in cfg.extras.items():
                    rows.append((k, v))
        self._fill_table(self.tbl_config, rows)

    def _fill_table(self, tbl: QTableWidget, rows: List[tuple], numeric_cols: set[int] | None = None, stretch_cols: set[int] | None = None, resize_cols: set[int] | None = None) -> None:
        numeric_cols = numeric_cols or set()
        stretch_cols = stretch_cols or set()
        resize_cols = resize_cols or set()
        tbl.setRowCount(0)
        for r_idx, row in enumerate(rows):
            tbl.insertRow(r_idx)
            for c_idx, val in enumerate(row):
                item = QTableWidgetItem("" if val is None else str(val))
                align = Qt.AlignRight if c_idx in numeric_cols else Qt.AlignCenter
                item.setTextAlignment(align)
                tbl.setItem(r_idx, c_idx, item)
        header = tbl.horizontalHeader()
        for c in range(tbl.columnCount()):
            if c in stretch_cols:
                header.setSectionResizeMode(c, QHeaderView.Stretch)
            elif c in resize_cols:
                header.setSectionResizeMode(c, QHeaderView.ResizeToContents)
            else:
                header.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
        tbl.resizeRowsToContents()
        self._fit_table_to_contents(tbl)

    def _fill_list(self, lst: QListWidget, items: List[str]) -> None:
        lst.clear()
        if not items:
            lst.addItem("(sin datos)")
            return
        for it in items:
            lst.addItem(QListWidgetItem(str(it)))
        self._fit_list_to_contents(lst)

    def _update_folder_elide(self) -> None:
        if not self._folder_full_path:
            return
        fm = QFontMetrics(self.lbl_folder.font())
        text = fm.elidedText(f"Carpeta: {self._folder_full_path}", Qt.ElideMiddle, self.lbl_folder.width())
        self.lbl_folder.setText(text)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_folder_elide()

    # Ajuste de tamaños para evitar scroll interno cuando hay pocas filas
    def _fit_table_to_contents(self, tbl: QTableWidget) -> None:
        max_rows_no_scroll = 40
        tbl.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        tbl.resizeColumnsToContents()
        tbl.resizeRowsToContents()
        row_count = tbl.rowCount()
        if row_count == 0:
            h = tbl.horizontalHeader().height() + tbl.frameWidth() * 2 + 6 + tbl.verticalHeader().defaultSectionSize()
        else:
            h = tbl.horizontalHeader().height() + sum(tbl.rowHeight(r) for r in range(row_count)) + tbl.frameWidth() * 2 + 6
        tbl.setMinimumHeight(h)
        tbl.setMaximumHeight(h if row_count <= max_rows_no_scroll else tbl.sizeHint().height())
        tbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed if row_count <= max_rows_no_scroll else QSizePolicy.Preferred)

    def _fit_list_to_contents(self, lst: QListWidget) -> None:
        max_rows_no_scroll = 40
        lst.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        lst.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        count = lst.count()
        if count == 0:
            h = lst.sizeHintForRow(0) + lst.frameWidth() * 2 + 8
        else:
            h = lst.sizeHintForRow(0) * count + lst.frameWidth() * 2 + 8
        if count <= max_rows_no_scroll:
            lst.setMinimumHeight(h)
            lst.setMaximumHeight(h)
            lst.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        else:
            lst.setMinimumHeight(min(h, 300))
            lst.setMaximumHeight(400)
            lst.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)


if __name__ == "__main__":
    from PySide6.QtWidgets import QApplication
    import sys

    app = QApplication(sys.argv)
    w = LegendPage()
    w.resize(800, 600)
    w.show()
    sys.exit(app.exec())
