# main.py
import sys
from PySide6.QtWidgets import QApplication
from gui.main_window import ElectroCalcApp

if __name__ == '__main__':
    app = QApplication(sys.argv)
    try:
        qss = open("resources/fusion_dark.qss", encoding="utf-8").read()
        app.setStyleSheet(qss)
    except:
        pass

    win = ElectroCalcApp()
    win.show()
    sys.exit(app.exec())
