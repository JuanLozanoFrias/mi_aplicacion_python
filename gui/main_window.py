from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QFrame, QListWidget, QListWidgetItem, QStackedWidget,
    QGraphicsDropShadowEffect, QLabel, QPushButton, QSizePolicy, QSpacerItem
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtSvgWidgets import QSvgWidget

from gui.pages.carga_page import CargaPage
from gui.pages.cargas.industrial_page import IndustrialPage
from gui.pages.materiales_page import MaterialesPage
from gui.pages.carga_electrica.carga_electrica_page import CargaElectricaPage
from gui.pages.creditos_page import CreditosPage


class ElectroCalcApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WESTON")
        self.resize(1220, 800)

        self._apply_base_style()
        self._build_ui()

    # ------------------------------------------------------------------
    #  ESTILO GLOBAL (QSS BASE + EXTRA)
    # ------------------------------------------------------------------
    def _apply_base_style(self):
        extra_qss = """
QMainWindow { background: #f3f6fb; }
QFrame { background: transparent; }
QLineEdit, QComboBox {
    background: #ffffff;
    color: #0f172a;
    border: 1px solid #d8deeb;
    border-radius: 8px;
    padding: 6px 8px;
    selection-background-color: #0f62fe;
}
QComboBox::drop-down { width: 18px; }
QLabel { color: #0f172a; }

#SideBar {
    background: #0c1220;
    border: none;
}
#SideBar QListWidget { background: transparent; border: none; outline: 0; }

#TopBar {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 #0f62fe, stop:1 #22d3ee);
    border-radius: 14px;
}
#TopTitle { color: #ffffff; font-size: 22px; font-weight: 800; }
#TopSubtitle { color: #e7eefc; font-size: 12px; font-weight: 600; }
#TopUser {
    color: #0f172a;
    background: #ffffff;
    padding: 6px 10px;
    border-radius: 10px;
    font-weight: 700;
}

#MainArea { background: transparent; }

QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 #0f62fe, stop:1 #22d3ee);
    color: #ffffff;
    border: none;
    border-radius: 10px;
    padding: 10px 16px;
    font-weight: 700;
}
QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 #1a6eff, stop:1 #38d7ef); }
QPushButton:pressed { background: #0d4fcc; }
QPushButton[variant="ghost"] {
    background: rgba(255,255,255,0.12);
    color: #e6efff;
    border: 1px solid rgba(255,255,255,0.24);
}

QScrollBar:vertical, QTableView QScrollBar:vertical {
    background: #e7ecf6;
    width: 12px;
    margin: 2px;
    border-radius: 6px;
}
QScrollBar::handle:vertical, QTableView QScrollBar::handle:vertical {
    background: #9bc6ff;
    min-height: 24px;
    border-radius: 6px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QTableView QScrollBar::add-line:vertical, QTableView QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical,
QTableView QScrollBar::add-page:vertical, QTableView QScrollBar::sub-page:vertical { background: none; }

QScrollBar:horizontal, QTableView QScrollBar:horizontal {
    background: #e7ecf6;
    height: 12px;
    margin: 2px;
    border-radius: 6px;
}
QScrollBar::handle:horizontal, QTableView QScrollBar::handle:horizontal {
    background: #9bc6ff;
    min-width: 24px;
    border-radius: 6px;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,
QTableView QScrollBar::add-line:horizontal, QTableView QScrollBar::sub-line:horizontal { width: 0; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal,
QTableView QScrollBar::add-page:horizontal, QTableView QScrollBar::sub-page:horizontal { background: none; }
"""
        self.setStyleSheet(extra_qss)

    # ------------------------------------------------------------------
    #  CONSTRUCCIÃ“N DE LA INTERFAZ
    # ------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # -------- Sidebar (logo + navegaciÃ³n) --------
        side = QFrame()
        side.setObjectName("SideBar")
        side.setFixedWidth(240)
        side_layout = QVBoxLayout(side)
        side_layout.setContentsMargins(18, 18, 18, 18)
        side_layout.setSpacing(12)

        side_layout.addWidget(self._create_logo_section())

        self.nav = QListWidget()
        self.nav.setSpacing(6)
        self.nav.setUniformItemSizes(True)
        self.nav.setSelectionMode(QListWidget.SingleSelection)
        self.nav.setSelectionRectVisible(False)
        self.nav.setFocusPolicy(Qt.NoFocus)
        nav_qss = """
QListWidget {
    background: transparent;
    border: none;
    outline: 0;
}
QListWidget::item {
    height: 38px;
    padding: 8px 12px;
    margin: 4px 8px;
    border-radius: 8px;
    font-size: 15px;
    letter-spacing: 0.3px;
    font-weight: 700;
    color: #e5edff;
    border-left: 3px solid transparent;
    background: transparent;
}
QListWidget::item:selected,
QListWidget::item:selected:active,
QListWidget::item:selected:!active {
    background: rgba(76,201,255,0.32);
    color: #ffffff;
    border-left: 3px solid #4cc9ff;
}
QListWidget::item:hover {
    background: rgba(255,255,255,0.14);
    color: #ffffff;
}
QListWidget::item:selected:hover {
    background: rgba(76,201,255,0.38);
    color: #ffffff;
}
"""
        self.nav.setStyleSheet(nav_qss)
        self.stack = QStackedWidget()
        self.pages = []
        self._sections = {}

        def add_nav_item(text: str, page: QWidget | None, *, header: bool = False, indent: bool = False, section: str | None = None):
            prefix = "    " if indent else ""
            item = QListWidgetItem(prefix + text)
            if header:
                item.setData(Qt.UserRole, None)
                item.setData(Qt.UserRole + 1, "section")
                item.setFlags(Qt.ItemIsEnabled)
            else:
                idx = len(self.pages) if page is not None else None
                item.setData(Qt.UserRole, idx)
                if section:
                    self._sections.setdefault(section, []).append(self.nav.count())
                if page is not None:
                    self.stack.addWidget(page)
                    self.pages.append(page)
            self.nav.addItem(item)

        add_nav_item("CARGAS TERMICAS", None, header=True)
        add_nav_item("CUARTOS FRIOS", CargaPage(), indent=True, section="cargas")
        add_nav_item("CUARTOS INDUSTRIALES", IndustrialPage(), indent=True, section="cargas")
        add_nav_item("LEGEND", LegendPage(self), indent=True, section="cargas")

        add_nav_item("TABLEROS ELECTRICOS", MaterialesPage())
        add_nav_item("CARGA ELECTRICA", CargaElectricaPage())
        add_nav_item("CREDITOS", CreditosPage())
        side_layout.addWidget(self.nav, stretch=1)
        root.addWidget(side)

        # -------- Ãrea de pÃ¡ginas --------
        main = QFrame()
        main.setObjectName("MainArea")
        main_layout = QVBoxLayout(main)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(12)

        main_layout.addWidget(self._create_top_bar())

        main_layout.addWidget(self.stack, 1)
        root.addWidget(main, 1)
        # ocultar submenus inicialmente
        for sec_rows in self._sections.values():
            for r in sec_rows:
                self.nav.item(r).setHidden(True)

        self.nav.currentRowChanged.connect(self._on_nav_changed)
        self.nav.itemClicked.connect(self._on_nav_clicked)
        # seleccionar primer elemento útil
        for i in range(self.nav.count()):
            idx = self.nav.item(i).data(Qt.UserRole)
            if idx is not None:
                self.nav.setCurrentRow(i)
                break

    # ------------------------------------------------------------------
    #  Header superior
    # ------------------------------------------------------------------
    def _create_placeholder_page(self, title: str) -> QWidget:
        page = QFrame()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24,24,24,24)
        layout.setSpacing(12)
        lab = QLabel(title)
        lab.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        layout.addWidget(lab)
        sub = QLabel("Seccion en construccion")
        sub.setStyleSheet("color:#475569;")
        layout.addWidget(sub)
        layout.addStretch(1)
        return page

    def _on_nav_changed(self, row: int):
        item = self.nav.item(row)
        if not item:
            return
        idx = item.data(Qt.UserRole)
        if idx is None:
            return
        self.stack.setCurrentIndex(idx)

    def _on_nav_clicked(self, item: QListWidgetItem):
        # si es sección, toggle de submenus
        if item.data(Qt.UserRole + 1) == "section":
            sec = "cargas"  # único por ahora
            rows = self._sections.get(sec, [])
            if not rows:
                return
            any_visible = any(not self.nav.item(r).isHidden() for r in rows)
            new_state = True if any_visible else False
            for r in rows:
                self.nav.item(r).setHidden(new_state)
            if not new_state:
                # al expandir, seleccionar el primero
                self.nav.setCurrentRow(rows[0])

    def _create_top_bar(self) -> QWidget:
        top = QFrame()
        top.setObjectName("TopBar")
        top.setFixedHeight(96)
        layout = QHBoxLayout(top)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(14)

        title_wrap = QVBoxLayout()
        title = QLabel("WESTON")
        title.setObjectName("TopTitle")
        subtitle = QLabel("DIVISIÓN DE INGENIERÍA")
        subtitle.setObjectName("TopSubtitle")
        title_wrap.addWidget(title)
        title_wrap.addWidget(subtitle)

        layout.addLayout(title_wrap)
        layout.addStretch(1)

        chip = QLabel("WESTON - CALVOBOT")
        chip.setObjectName("TopUser")
        chip.setAlignment(Qt.AlignCenter)
        chip.setMinimumWidth(150)
        layout.addWidget(chip, alignment=Qt.AlignVCenter)

        return top

    # ------------------------------------------------------------------
    #  Logo en la parte superior de la barra lateral
    # ------------------------------------------------------------------
    def _create_logo_section(self):
        profile = QFrame()
        profile.setFixedHeight(220)
        p_layout = QVBoxLayout(profile)
        p_layout.setContentsMargins(0, 4, 0, 8)

        logo_frame = QFrame()
        logo_frame.setFixedSize(180, 180)
        logo_frame.setStyleSheet(
            "QFrame{background:#111827;border:1px solid #1f2937;border-radius:16px;}")
        logo_frame.setGraphicsEffect(QGraphicsDropShadowEffect(
            blurRadius=18, xOffset=0, yOffset=6, color=QColor(0, 0, 0, 180)))
        lf_layout = QVBoxLayout(logo_frame)
        lf_layout.setAlignment(Qt.AlignCenter)

        svg_path = next((p / "resources" / "logo.svg" for p in Path(__file__).resolve().parents
                         if (p / "resources" / "logo.svg").exists()), None)
        if svg_path:
            svg = QSvgWidget(str(svg_path))
            svg.setFixedSize(150, 150)
            lf_layout.addWidget(svg)
        else:
            placeholder = QLabel("E")
            placeholder.setStyleSheet("color:#8dd5ff;font-size:48px;font-weight:800;")
            placeholder.setAlignment(Qt.AlignCenter)
            lf_layout.addWidget(placeholder)

        p_layout.addWidget(logo_frame, alignment=Qt.AlignHCenter)
        p_layout.addStretch()
        return profile
