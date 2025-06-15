# gui/widgets.py
from PySide6.QtWidgets import QPushButton, QGraphicsDropShadowEffect
from PySide6.QtGui     import QIcon, QColor
from PySide6.QtCore    import QSize, Qt, QPropertyAnimation, QEasingCurve

class FancyButton(QPushButton):
    """Custom button with a blue gradient and subtle scale animation."""

    def __init__(self, text: str, icon: str | None = None):
        super().__init__(text.upper())
        if icon:
            self.setIcon(QIcon(icon))
            self.setIconSize(QSize(28, 28))
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
        effect.setBlurRadius(20)
        effect.setOffset(0, 4)
        effect.setColor(QColor(0, 0, 0, 120))
        self.setGraphicsEffect(effect)
        self._anim = QPropertyAnimation(self, b"geometry", self)
        self._anim.setDuration(150)
        self._anim.setEasingCurve(QEasingCurve.OutQuad)

    def enterEvent(self, e):
        g = self.geometry().adjusted(-2, -2, 2, 2)
        self._anim.setStartValue(self.geometry())
        self._anim.setEndValue(g)
        self._anim.start()
        super().enterEvent(e)

    def leaveEvent(self, e):
        g = self.geometry().adjusted(2, 2, -2, -2)
        self._anim.setStartValue(self.geometry())
        self._anim.setEndValue(g)
        self._anim.start()
        super().leaveEvent(e)

