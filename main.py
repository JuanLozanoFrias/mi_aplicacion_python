# main.py — Arranque de la app (no cambia tu UI)
# Carga el QSS base de siempre (fusion_dark.qss) y, si existen,
# aplica overlays suaves: soft_bg_override.qss (solo fondo)
# y readability_override.qss (texto más nítido).

import sys
from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QFile, QTextStream

from gui.main_window import ElectroCalcApp  # ← tu ventana principal

# -----------------------------------------------
# Utilidad: leer y concatenar archivos QSS en orden
# -----------------------------------------------
def _read_qss(path: Path) -> str:
    try:
        f = QFile(str(path))
        if f.open(QFile.ReadOnly | QFile.Text):
            css = QTextStream(f).readAll()
            f.close()
            return css
    except Exception:
        pass
    return ""

def load_styles(app: QApplication) -> None:
    """
    Aplica estilos en este orden (si existen):
      1) resources/fusion_dark.qss              ← tu tema actual
      2) resources/soft_bg_override.qss         ← opcional, SOLO fondo
      3) resources/readability_override.qss     ← opcional, mejora contraste de texto
    """
    resources = Path(__file__).resolve().parent / "resources"
    order = [
        "fusion_dark.qss",
        "soft_bg_override.qss",
        "readability_override.qss",
    ]
    qss_total = ""
    for name in order:
        p = resources / name
        if p.exists():
            qss_total += "\n\n" + _read_qss(p)
    if qss_total:
        app.setStyleSheet(qss_total)

# -----------------------------------------------
# Punto de entrada
# -----------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    load_styles(app)

    win = ElectroCalcApp()
    win.show()

    sys.exit(app.exec())
