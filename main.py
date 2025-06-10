import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QFrame, QLabel, QListWidget, QListWidgetItem, QStackedWidget,
    QLineEdit, QFileDialog, QGraphicsDropShadowEffect, QPushButton,
    QProgressBar, QComboBox, QFormLayout
)
from PySide6.QtGui import QIcon, QColor, QAction, QPixmap, QPainter, QPainterPath
from PySide6.QtCore import Qt, QSize, QPropertyAnimation, QEasingCurve
from pathlib import Path
from PySide6.QtSvgWidgets import QSvgWidget


class FancyButton(QPushButton):
    def __init__(self, text: str, icon: str | None = None):
        super().__init__(text.upper())
        if icon:
            self.setIcon(QIcon(icon)); self.setIconSize(QSize(28, 28))
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumSize(180, 60)
        self.setStyleSheet(
            "QPushButton{"
            "background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #5b8bea,stop:1 #3c6fd7);"
            "color:#fff;font-size:16px;font-weight:600;font-family:'Segoe UI',sans-serif;"
            "border:none;border-radius:8px;padding:12px 24px;}"
            "QPushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #6c9cf0,stop:1 #4a80dd);}"  
            "QPushButton:pressed{background:#375fb8;}"
        )
        effect = QGraphicsDropShadowEffect(self)
        effect.setBlurRadius(20); effect.setOffset(0, 4); effect.setColor(QColor(0, 0, 0, 120))
        self.setGraphicsEffect(effect)
        self._anim = QPropertyAnimation(self, b"geometry", self)
        self._anim.setDuration(150); self._anim.setEasingCurve(QEasingCurve.OutQuad)

    def enterEvent(self, e):
        g = self.geometry().adjusted(-2, -2, 2, 2)
        self._anim.setStartValue(self.geometry()); self._anim.setEndValue(g); self._anim.start()
        super().enterEvent(e)

    def leaveEvent(self, e):
        g = self.geometry().adjusted(2, 2, -2, -2)
        self._anim.setStartValue(self.geometry()); self._anim.setEndValue(g); self._anim.start()
        super().leaveEvent(e)


class ElectroCalcApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CALVOBOT")
        self.resize(1024, 700)
        self.setStyleSheet(
            "QMainWindow{background:#232634;}"
            " QFrame{background:transparent;}"
            " QLineEdit{background:#2a2d3c;color:#e0e0e0;border:1px solid #3a3d4f;border-radius:6px;padding:8px;}"
            " QLabel{color:#e0e0e0;}"
            " QComboBox{background:#2a2d3c;color:#e0e0e0;border:1px solid #3a3d4f;border-radius:6px;padding:6px;}"
            " QComboBox QAbstractItemView{background:#2a2d3c;color:#ffffff;selection-background-color:#5b8bea;}"
        )
        self._build_ui()

    def _build_ui(self):
        self._create_toolbar()
        self._create_main_layout()
        self._create_pages()
        self.nav.setCurrentRow(0)

    def _create_toolbar(self):
        tb = self.addToolBar("Main")
        tb.setMovable(False)
        tb.setIconSize(QSize(24, 24))
        tb.setStyleSheet(
            "QToolBar{background:#1f2028;border-bottom:1px solid #3a3d4f;}"
            " QToolButton{color:#e0e0e0;} QToolButton:hover{background:rgba(255,255,255,0.05);}"  
        )
        exit_act = QAction(QIcon("resources/exit.png"), "Salir", self)
        exit_act.triggered.connect(self.close)
        tb.addAction(exit_act)

    def _create_main_layout(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Sidebar
        side = QFrame()
        side.setFixedWidth(220)
        side.setStyleSheet("background:#2d2f3e;")
        side_layout = QVBoxLayout(side)
        side_layout.setContentsMargins(0, 0, 0, 0)
        side_layout.setSpacing(0)

        # Profile logo
        profile = QFrame()
        profile.setFixedHeight(240)
        profile.setStyleSheet("background:transparent;")
        p_layout = QVBoxLayout(profile)
        p_layout.setContentsMargins(0, 8, 0, 16)
        p_layout.setSpacing(0)
        logo_frame = QFrame()
        logo_frame.setFixedSize(200, 200)
        logo_frame.setStyleSheet("QFrame{background:#1f2028; border:2px solid #3e4ac1; border-radius:14px;}")
        logo_frame.setGraphicsEffect(QGraphicsDropShadowEffect(blurRadius=14, xOffset=0, yOffset=3, color=QColor(0, 0, 0, 150)))
        lf_layout = QVBoxLayout(logo_frame)
        lf_layout.setContentsMargins(0, 0, 0, 0)
        lf_layout.setAlignment(Qt.AlignCenter)
        # load logo
        svg_path = next((p / "resources" / "logo.svg" for p in Path(__file__).resolve().parents if (p / "resources" / "logo.svg").exists()), None)
        if svg_path:
            svg = QSvgWidget(str(svg_path)); svg.setFixedSize(180, 180); lf_layout.addWidget(svg)
        else:
            png_path = next((p / "resources" / "logo.png" for p in Path(__file__).resolve().parents if (p / "resources" / "logo.png").exists()), None)
            lbl = QLabel()
            if png_path:
                pix = QPixmap(str(png_path)); pix = pix.scaled(180, 180, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                rounded = QPixmap(pix.size()); rounded.fill(Qt.transparent)
                painter = QPainter(rounded); painter.setRenderHint(QPainter.Antialiasing)
                path = QPainterPath(); path.addRoundedRect(0, 0, pix.width(), pix.height(), 12, 12)
                painter.setClipPath(path); painter.drawPixmap(0, 0, pix); painter.end()
                lbl.setPixmap(rounded)
            lf_layout.addWidget(lbl)
        p_layout.addWidget(logo_frame, alignment=Qt.AlignTop | Qt.AlignHCenter)
        p_layout.addStretch()
        side_layout.addWidget(profile)
        # Navigation
        self.nav = QListWidget()
        self.nav.setStyleSheet(
            "QListWidget{border:none;background:#2d2f3e;font-family:'Segoe UI';font-size:16px;}"
            " QListWidget::item{height:56px;padding-left:24px;font-size:20px;font-weight:600;color:#f0f0f4;border-left:4px solid transparent;}"
            " QListWidget::item:hover{background:#383a4a;color:#ffffff;}"
            " QListWidget::item:selected{background:#36406a;color:#becfff;font-weight:700;border-left:4px solid #8fb0ff;}"
            " QScrollBar:vertical{width:0px;}"
        )
        entries = ["Cargas Termicas", "Materiales", "Voltaje", "Corriente", "Resistencia"]
        icons = ["icon1.png", "icon2.png", "icon3.png", "icon4.png", "icon5.png"]
        for txt, ico in zip(entries, icons):
            path = Path("resources") / ico
            if path.exists():
                item = QListWidgetItem(QIcon(str(path)), txt)
            else:
                item = QListWidgetItem(txt)
            self.nav.addItem(item)
        self.nav.currentRowChanged.connect(lambda i: self.stack.setCurrentIndex(i))
        side_layout.addWidget(self.nav, stretch=1)
        main_layout.addWidget(side)
        # Pages area
        self.stack = QStackedWidget()
        main_layout.addWidget(self.stack)

    def _create_pages(self):
        self.stack.addWidget(self._page_cargas())
        self.stack.addWidget(self._simple_page("CÁLCULO DE MATERIALES", "Seleccionar archivo de Materiales"))
        self.stack.addWidget(self._simple_page("CÁLCULO DE VOLTAJE", "FUENTE DE VOLTAJE (V):"))
        self.stack.addWidget(self._simple_page("CÁLCULO DE CORRIENTE", "RESISTENCIA (Ω):"))
        self.stack.addWidget(self._simple_page("CÁLCULO DE RESISTENCIA", "VOLTAJE (V):"))

    def _page_cargas(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        card = QFrame()
        card.setStyleSheet("background:#2a2d3c;border-radius:12px;")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(30, 30, 30, 30)
        cl.setSpacing(16)

        # Título
        title = QLabel("CARGA TÉRMICA")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:24px;font-weight:700;")
        cl.addWidget(title)

        # Selección de tipo
        row = QHBoxLayout()
        subtitle = QLabel("Seleccione su tipo de configuración:")
        subtitle.setStyleSheet("font-size:16px;")
        combo = QComboBox()
        combo.addItems(["Racks", "Autocontenidas", "Waterloop"])
        combo.setFixedWidth(200)
        combo.setCurrentIndex(-1)
        combo.currentTextChanged.connect(self.on_carga_selected)
        row.addWidget(subtitle)
        row.addStretch()
        row.addWidget(combo)
        cl.addLayout(row)

        # Datos adicionales
        form = QFormLayout()
        form.addRow("Nombre del proyecto:", QLineEdit())
        form.addRow("Ciudad:", QLineEdit())
        form.addRow("Responsable:", QLineEdit())
        form.addRow("Nº de ramales baja:", QLineEdit())
        form.addRow("Nº de ramales media:", QLineEdit())
        form.addRow("Nº de loops baja:", QLineEdit())
        form.addRow("Nº de loops media:", QLineEdit())
        form.addRow("Temperatura ambiente:", QLineEdit())
        form.addRow("Tensión de equipos:", QLineEdit())
        cl.addLayout(form)

        layout.addWidget(card)
        return page

    def _simple_page(self, title, label):
        w = QWidget()
        l = QVBoxLayout(w)
        l.setContentsMargins(40, 40, 40, 40)
        l.setSpacing(20)
        header = QLabel(title)
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("font-size:22px;font-weight:600;")
        l.addWidget(header)
        l.addWidget(QLabel(label))
        l.addWidget(QLineEdit())
        l.addStretch()
        return w

    def on_carga_selected(self, text):
        print(f"Opción seleccionada: {text}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    try:
        app.setStyleSheet(open("resources/fusion_dark.qss", encoding="utf-8").read())
    except FileNotFoundError:
        pass
    win = ElectroCalcApp()
    win.show()
    sys.exit(app.exec())
