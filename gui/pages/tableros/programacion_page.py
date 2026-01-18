# gui/pages/programacion_page.py
from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFrame, QLabel, QPushButton, QFileDialog, QHBoxLayout
)

from logic.export_step4 import export_step4


class ProgramacionPage(QWidget):
    """
    Página de 'Programación / Exportar'.
    Genera Excel del Paso 4 + Snapshot (.ecalc.json) sin tocar la lógica existente.
    """

    def __init__(
        self,
        get_step2_state: Callable[[], Dict] = lambda: {},
        get_globals: Callable[[], Dict] = lambda: {},
        basedatos_path: Path | None = None,
    ) -> None:
        super().__init__()
        self._get_step2 = get_step2_state
        self._get_globals = get_globals
        self._book = basedatos_path or Path(__file__).resolve().parents[3] / "data" / "tableros_electricos" / "basedatos.xlsx"

        self._out_dir = (Path.home() / "Documents" / "ElectroCalc_Exports")
        self._out_dir.mkdir(parents=True, exist_ok=True)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        # Card: título + selector carpeta
        card = QFrame()
        card.setStyleSheet("QFrame{background:#2f3242;border:1px solid #42465a;border-radius:8px;}")
        cl = QVBoxLayout(card); cl.setContentsMargins(12, 12, 12, 12); cl.setSpacing(8)

        title = QLabel("PROGRAMACIÓN / EXPORTAR PASO 4")
        title.setStyleSheet("color:#e0e0e0;font-weight:700;")
        cl.addWidget(title)

        desc = QLabel("Genera un Excel del Resumen del Paso 4 y un snapshot (.ecalc.json) para recargar el estado.")
        desc.setStyleSheet("color:#c7cee8;")
        desc.setWordWrap(True)
        cl.addWidget(desc)

        hl = QHBoxLayout()
        self.dir_lab = QLabel(str(self._out_dir))
        self.dir_lab.setStyleSheet("color:#aeb7d6;")
        btn_dir = QPushButton("Cambiar carpeta…")
        btn_dir.setStyleSheet("QPushButton{background:#3a3f55;color:#e0e0e0;border:1px solid #4a4f67;border-radius:6px;padding:6px 10px;}")
        btn_dir.clicked.connect(self._pick_dir)
        hl.addWidget(self.dir_lab, 1); hl.addWidget(btn_dir, 0)
        cl.addLayout(hl)

        # Botón principal
        self.btn_export = QPushButton("Imprimir / Exportar Excel + Snapshot")
        self.btn_export.setStyleSheet("QPushButton{background:#4a8bff;color:white;border:none;border-radius:8px;padding:10px 14px;font-weight:700;} QPushButton:pressed{background:#3c72d1;}")
        self.btn_export.clicked.connect(self._run_export)
        cl.addWidget(self.btn_export, 0, Qt.AlignLeft)

        # Resultado
        self.result_lab = QLabel("")
        self.result_lab.setStyleSheet("color:#9ed28a;")
        self.result_lab.setWordWrap(True)
        cl.addWidget(self.result_lab)

        root.addWidget(card)
        root.addStretch(1)

    # ---------------------------- acciones ----------------------------

    def _pick_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de salida", str(self._out_dir))
        if path:
            self._out_dir = Path(path)
            self.dir_lab.setText(path)

    def _run_export(self) -> None:
        try:
            step2 = self._get_step2() or {}
            globs = self._get_globals() or {}
            paths = export_step4(self._book, step2, globs, self._out_dir)
            self.result_lab.setText(
                f"✅ Archivos generados:\n"
                f"• Excel: {paths.xlsx_path}\n"
                f"• Snapshot: {paths.snapshot_path}"
            )
        except Exception as e:
            self.result_lab.setStyleSheet("color:#ff9e9e;")
            self.result_lab.setText(f"❌ Error al exportar: {e}")

