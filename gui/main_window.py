from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QFrame, QListWidget, QListWidgetItem, QStackedWidget,
    QGraphicsDropShadowEffect, QLabel
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtSvgWidgets import QSvgWidget

from gui.pages.carga_page import CargaPage
#  >> aquí podrás importar tus páginas futuras
# from gui.pages.materiales_page import MaterialesPage
# from gui.pages.voltaje_page    import VoltajePage

class ElectroCalcApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CALVOBOT")
        self.resize(1024, 700)

        self._apply_base_style()
        self._build_ui()

    # ------------------------------------------------------------------  
    #  ESTILO GLOBAL (QSS BASE + EXTRA)  
    # ------------------------------------------------------------------  
    def _apply_base_style(self):
        project_root = Path(__file__).resolve().parents[2]
        qss_file     = project_root / "resources" / "fusion_dark.qss"
        base_qss     = qss_file.read_text(encoding="utf-8") if qss_file.exists() else ""

        extra_qss = """
/* ---- Colores generales ---- */
QMainWindow { background: #232634; }
QFrame      { background: transparent; }
QLineEdit   { background: #2a2d3c; color: #e0e0e0; border: 1px solid #3a3d4f; }
QLabel      { color: #e0e0e0; }
QComboBox   { background: #2a2d3c; color: #e0e0e0; border: 1px solid #3a3d4f; }
QComboBox QAbstractItemView { background: #2a2d3c; color: #ffffff; }

/* ---- Scrollbars globales ---- */
QScrollBar:vertical, QTableView QScrollBar:vertical {
    background: #2a2d3c;
    width: 12px;
}
QScrollBar::handle:vertical, QTableView QScrollBar::handle:vertical {
    background: #5b8bea;
    min-height: 20px;
    border-radius: 6px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QTableView QScrollBar::add-line:vertical, QTableView QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
QTableView QScrollBar::add-page:vertical, QTableView QScrollBar::sub-page:vertical {
    background: none;
}

QScrollBar:horizontal, QTableView QScrollBar:horizontal {
    background: #2a2d3c;
    height: 12px;
}
QScrollBar::handle:horizontal, QTableView QScrollBar::handle:horizontal {
    background: #5b8bea;
    min-width: 20px;
    border-radius: 6px;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,
QTableView QScrollBar::add-line:horizontal, QTableView QScrollBar::sub-line:horizontal {
    width: 0;
}
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal,
QTableView QScrollBar::add-page:horizontal, QTableView QScrollBar::sub-page:horizontal {
    background: none;
}

/* ---- Sidebar ---- */
QListWidget {
    border: none;
    background: #2d2f3e;
    font-family: 'Segoe UI';
    font-size: 16px;
}
QListWidget::item {
    height: 56px;
    padding-left: 24px;
    font-size: 20px;
    font-weight: 600;
    color: #f0f0f4;
    border-left: 4px solid transparent;
}
QListWidget::item:hover         { background: #383a4a; color:#ffffff; }
QListWidget::item:selected      { background: #36406a; color:#becfff; font-weight:700;
                                   border-left:4px solid #8fb0ff; }

/* ---- Tablas ---- */
QTableWidget {
    background: #2a2d3c;
    color: #ffffff;
    gridline-color: #3a3d4f;
    border: none;
    selection-background-color: #3c6fd7;
    selection-color: #ffffff;
}
QHeaderView::section {
    background: #2d2f3e;
    color: #ffffff;
    padding-left: 4px;
    border: 1px solid #3a3d4f;
    height: 24px;
}
QTableCornerButton::section {
    background: #2d2f3e;
    border: 1px solid #3a3d4f;
}
"""
        self.setStyleSheet(base_qss + extra_qss)

    # ------------------------------------------------------------------  
    #  CONSTRUCCIÓN DE LA INTERFAZ  
    # ------------------------------------------------------------------  
    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        root    = QHBoxLayout(central); root.setContentsMargins(0,0,0,0)

        # -------- Sidebar (logo + navegación) --------
        side           = QFrame(); side.setFixedWidth(220)
        side.setStyleSheet("background:#2d2f3e;")
        side_layout    = QVBoxLayout(side); side_layout.setContentsMargins(0,0,0,0)
        side_layout.setSpacing(0)

        side_layout.addWidget(self._create_logo_section())

        self.nav = QListWidget()
        for txt in ["Cargas Térmicas", "Materiales", "Voltaje", "Corriente", "Resistencia"]:
            self.nav.addItem(QListWidgetItem(txt))
        side_layout.addWidget(self.nav, stretch=1)
        root.addWidget(side)

        # -------- Área de páginas --------
        self.stack = QStackedWidget()
        self.stack.addWidget(CargaPage())
        # self.stack.addWidget(MaterialesPage())
        # self.stack.addWidget(VoltajePage())
        root.addWidget(self.stack, 1)

        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)
        self.nav.setCurrentRow(0)

    # ------------------------------------------------------------------  
    #  Logo en la parte superior de la barra lateral  
    # ------------------------------------------------------------------  
    def _create_logo_section(self):
        profile = QFrame(); profile.setFixedHeight(240)
        p_layout = QVBoxLayout(profile); p_layout.setContentsMargins(0,8,0,16)

        logo_frame = QFrame(); logo_frame.setFixedSize(200,200)
        logo_frame.setStyleSheet(
            "QFrame{background:#1f2028;border:2px solid #3e4ac1;border-radius:14px;}")
        logo_frame.setGraphicsEffect(QGraphicsDropShadowEffect(
            blurRadius=14, xOffset=0, yOffset=3, color=QColor(0,0,0,150)))
        lf_layout = QVBoxLayout(logo_frame); lf_layout.setAlignment(Qt.AlignCenter)

        svg_path = next((p/"resources"/"logo.svg" for p in Path(__file__).resolve().parents
                         if (p/"resources"/"logo.svg").exists()), None)
        if svg_path:
            svg = QSvgWidget(str(svg_path)); svg.setFixedSize(180,180)
            lf_layout.addWidget(svg)
        else:
            lf_layout.addWidget(QLabel())

        p_layout.addWidget(logo_frame, alignment=Qt.AlignHCenter)
        p_layout.addStretch()
        return profile
