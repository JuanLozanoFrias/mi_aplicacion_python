from __future__ import annotations

from pathlib import Path
from typing import Dict, List

from PySide6.QtCore import Qt
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

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        title = QLabel("LEGEND (solo lectura)")
        title.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        layout.addWidget(title)

        self.btn_refresh = QPushButton("Refrescar")
        self.btn_refresh.clicked.connect(self.refresh)
        layout.addWidget(self.btn_refresh)

        self.lbl_summary = QLabel("Resumen: --")
        layout.addWidget(self.lbl_summary)

        # Secciones
        self.tbl_config = self._make_table(["Campo", "Valor"])
        layout.addWidget(self._wrap_group("Config", self.tbl_config))

        self.tbl_equipos = self._make_table(["Equipo", "BTU/h ft"])
        layout.addWidget(self._wrap_group("Equipos", self.tbl_equipos))

        self.tbl_usos_bt = self._make_table(["Usos BT"])
        layout.addWidget(self._wrap_group("Usos BT", self.tbl_usos_bt))

        self.tbl_usos_mt = self._make_table(["Usos MT"])
        layout.addWidget(self._wrap_group("Usos MT", self.tbl_usos_mt))

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

        if _IMPORT_ERROR:
            QMessageBox.critical(self, "Legend", f"No se pudo importar LegendJDService:\n{_IMPORT_ERROR}")
        else:
            self.refresh()

    def _make_table(self, headers: List[str]) -> QTableWidget:
        tbl = QTableWidget(0, len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.horizontalHeader().setStretchLastSection(True)
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        tbl.setSelectionMode(QTableWidget.NoSelection)
        tbl.setAlternatingRowColors(True)
        return tbl

    def _wrap_group(self, title: str, widget: QWidget) -> QGroupBox:
        box = QGroupBox(title)
        lay = QVBoxLayout(box)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.addWidget(widget)
        return box

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

        resumen = (
            f"Folder: {sources.get('folder', 'data/LEGEND')} | "
            f"Archivos: {', '.join(sources.get('files_found', [])) or 'ninguno'} | "
            f"Equipos: {len(equipos)} | Usos BT: {len(usos.get('BT', []))} | Usos MT: {len(usos.get('MT', []))} | "
            f"Variadores: {len(variadores)} | WCR: {len(wcr)}"
        )
        self.lbl_summary.setText(resumen)

        self._fill_config(cfg)
        self._fill_simple_table(self.tbl_equipos, [(e.equipo, e.btu_hr_ft) for e in equipos])
        self._fill_simple_table(self.tbl_usos_bt, [(u,) for u in usos.get("BT", [])])
        self._fill_simple_table(self.tbl_usos_mt, [(u,) for u in usos.get("MT", [])])
        self._fill_simple_table(self.tbl_variadores, [(v.modelo, v.potencia) for v in variadores])
        self._fill_simple_table(self.tbl_wcr, [(w.modelo, w.capacidad) for w in wcr])
        for name, tbl in self.tbl_plantillas.items():
            items = plantillas.get(name, []) or []
            self._fill_simple_table(tbl, [(p.qty, p.descripcion) for p in items])

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
        self._fill_simple_table(self.tbl_config, rows)

    def _fill_simple_table(self, tbl: QTableWidget, rows: List[tuple]) -> None:
        tbl.setRowCount(0)
        for r_idx, row in enumerate(rows):
            tbl.insertRow(r_idx)
            for c_idx, val in enumerate(row):
                item = QTableWidgetItem("" if val is None else str(val))
                item.setTextAlignment(Qt.AlignCenter)
                tbl.setItem(r_idx, c_idx, item)


if __name__ == "__main__":
    # Prueba mínima de standalone
    from PySide6.QtWidgets import QApplication
    import sys

    app = QApplication(sys.argv)
    w = LegendPage()
    w.resize(800, 600)
    w.show()
    sys.exit(app.exec())
