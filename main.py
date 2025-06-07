# main.py – Juan David Bot

import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QListWidget, QStackedWidget,
    QWidget, QHBoxLayout, QVBoxLayout, QFrame, QLabel, QLineEdit,
    QPushButton, QFileDialog, QGraphicsDropShadowEffect
)
from PySide6.QtGui import QIcon, QColor, QAction
from PySide6.QtCore import Qt, QSize, QPropertyAnimation, QEasingCurve


class FancyButton(QPushButton):
    def __init__(self, text: str, icon: str | None = None):
        super().__init__(text.upper())
        if icon:
            self.setIcon(QIcon(icon)); self.setIconSize(QSize(28, 28))
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumSize(180, 70)
        self.setStyleSheet(
            "QPushButton{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #4e88c2,stop:1 #3b6da8);color:#fff;font-size:18px;font-weight:600;font-family:'Segoe UI',sans-serif;border:none;border-radius:10px;}"
            "QPushButton:hover{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #5b96d8,stop:1 #4b7fb1);}"
            "QPushButton:pressed{background:qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #3b6ba0,stop:1 #2e5784);}"
        )
        shadow = QGraphicsDropShadowEffect(self); shadow.setBlurRadius(24); shadow.setOffset(0, 6); shadow.setColor(QColor(0, 0, 0, 160)); self.setGraphicsEffect(shadow)
        self._anim = QPropertyAnimation(self, b"geometry", self); self._anim.setDuration(150); self._anim.setEasingCurve(QEasingCurve.OutQuad)

    def enterEvent(self, e):
        g = self.geometry().adjusted(-3, -3, 3, 3); self._anim.stop(); self._anim.setStartValue(self.geometry()); self._anim.setEndValue(g); self._anim.start(); super().enterEvent(e)

    def leaveEvent(self, e):
        g = self.geometry().adjusted(3, 3, -3, -3); self._anim.stop(); self._anim.setStartValue(self.geometry()); self._anim.setEndValue(g); self._anim.start(); super().leaveEvent(e)


class ElectroCalcApp(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle("Juan David Bot"); self.resize(1000, 650); self._build_ui()

    # Toolbar
    def _create_toolbar(self):
        tb = self.addToolBar("TB"); tb.setMovable(False); tb.setIconSize(QSize(24, 24))
        tb.setStyleSheet("QToolBar{background:#282a36;border-bottom:1px solid #44475a;}QToolButton{color:#f8f8f2;padding:4px;border-radius:4px;}QToolButton:hover{background:rgba(255,255,255,.1);}")
        exit_act = QAction(QIcon("resources/exit.png"), "Salir", self); exit_act.triggered.connect(self.close); tb.addAction(exit_act)

    # Central layout
    def _create_central(self):
        central = QWidget(); central.setStyleSheet("background:#282a36;")
        self.main = QHBoxLayout(central); self.main.setContentsMargins(0, 0, 0, 0); self.setCentralWidget(central)

    # Sidebar
    def _create_sidebar(self):
        self.nav = QListWidget(); self.nav.setFixedWidth(220)
        self.nav.setStyleSheet("QListWidget{background:#21232f;color:#f8f8f2;border:none;}QListWidget::item{padding:10px;}QListWidget::item:selected{background:#6272a4;border-radius:4px;color:#fff;}")
        for p in ["Cuadros de carga", "Materiales", "Voltaje", "Corriente", "Resistencia"]:
            self.nav.addItem(p)
        self.nav.currentRowChanged.connect(lambda i: self.stack.setCurrentIndex(i)); self.main.addWidget(self.nav)

    # Pages stack
    def _create_pages(self):
        self.stack = QStackedWidget(); self.main.addWidget(self.stack)
        self.stack.addWidget(self._page_cuadros())
        self.stack.addWidget(self._page_materiales())
        self.stack.addWidget(self._simple_page("CÁLCULO DE VOLTAJE", "FUENTE DE VOLTAJE (V):"))
        self.stack.addWidget(self._simple_page("CÁLCULO DE CORRIENTE", "RESISTENCIA (Ω):"))
        self.stack.addWidget(self._simple_page("CÁLCULO DE RESISTENCIA", "VOLTAJE (V):"))

    # Page Cuadros
    def _page_cuadros(self):
        page = QWidget(); page.setStyleSheet("background:#282a36;")
        outer = QVBoxLayout(page); outer.setContentsMargins(0, 40, 0, 40)
        card = QFrame(); card.setMaximumWidth(560); card.setStyleSheet("QFrame{background:#313244;border-radius:16px;border:1px solid #44475a;}")
        shadow = QGraphicsDropShadowEffect(card); shadow.setBlurRadius(20); shadow.setOffset(0, 10); shadow.setColor(QColor(0,0,0,180)); card.setGraphicsEffect(shadow)
        v = QVBoxLayout(card); v.setContentsMargins(40, 40, 40, 40); v.setSpacing(32)
        title = QLabel("CUADROS DE CARGA"); title.setAlignment(Qt.AlignCenter); title.setStyleSheet("color:#f8f8f2;font-size:30px;font-weight:700;")
        subtitle = QLabel("Selecciona un modo de operación para gestionar tus cuadros de carga."); subtitle.setAlignment(Qt.AlignCenter); subtitle.setWordWrap(True); subtitle.setStyleSheet("color:#b0b0c0;font-size:15px;")
        v.addWidget(title); v.addWidget(subtitle)
        row = QHBoxLayout(); row.setSpacing(40); row.addStretch(); row.addWidget(FancyButton("MANUAL","resources/manual.png")); row.addWidget(FancyButton("LEGEND","resources/legend.png")); row.addStretch(); v.addLayout(row)
        outer.addStretch(); outer.addWidget(card, alignment=Qt.AlignCenter); outer.addStretch(); return page

    # Page Materiales
    def _page_materiales(self):
        return self._upload_page("CÁLCULO DE MATERIALES", "input_mat", "Seleccionar archivo de Materiales")

    # Reusable pages
    def _upload_page(self, title, attr, caption):
        w = QWidget(); w.setStyleSheet("background:#282a36;")
        l = QVBoxLayout(w); l.setContentsMargins(40,40,40,40); l.setSpacing(24)
        lbl = QLabel(title); lbl.setAlignment(Qt.AlignCenter); lbl.setStyleSheet("color:#f8f8f2;font-size:22px;font-weight:bold;")
        line = QLineEdit(); line.setReadOnly(True); setattr(self, attr, line); line.setStyleSheet("background:#44475a;color:#f8f8f2;border:1px solid #6272a4;border-radius:6px;padding:10px;")
        btn = FancyButton("SELECCIONAR ARCHIVO…"); btn.clicked.connect(lambda: self._browse(line, caption))
        l.addWidget(lbl); l.addWidget(line); l.addWidget(btn); l.addStretch(); return w

    def _simple_page(self, title, label_text):
        w = QWidget(); w.setStyleSheet("background:#282a36;")
        l = QVBoxLayout(w); l.setContentsMargins(40,40,40,40); l.setSpacing(24)
        lbl = QLabel(title); lbl.setAlignment(Qt.AlignCenter); lbl.setStyleSheet("color:#f8f8f2;font-size:22px;font-weight:bold;")
        lab = QLabel(label_text); lab.setStyleSheet("color:#f8f8f2;font-size:16px;")
        l.addWidget(lbl); l.addWidget(lab); l.addWidget(QLineEdit()); l.addWidget(FancyButton("CALCULAR")); l.addStretch(); return w

    # File dialog
    def _browse(self, line_edit, caption):
        path, _ = QFileDialog.getOpenFileName(self, caption, "", "Excel (*.xlsx *.xls)");
        if path: line_edit.setText(path)

    def _build_ui(self):
        self._create_toolbar(); self._create_central(); self._create_sidebar(); self._create_pages(); self.nav.setCurrentRow(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    try:
        with open("resources/fusion_dark.qss", encoding="utf-8") as f:
            app.setStyleSheet(f.read())
    except FileNotFoundError:
        pass
    win = ElectroCalcApp(); win.show(); sys.exit(app.exec())
