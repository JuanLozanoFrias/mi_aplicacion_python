# src/gui/pages/carga_page.py
# ——————————————————————————————————————————————————————————
# Esta versión añade un botón «GENERAR TABLA».
# La tabla solo se reconstruye cuando el usuario pulsa el botón,
# evitando el refresco constante que provocaba lentitud.
# • Todo lo previamente escrito en la tabla se conserva.
# • Conversión °C ↔ °F funciona inmediatamente, pero NO refresca la tabla
#   (solo actualiza el otro campo de temperatura).
# ——————————————————————————————————————————————————————————
from pathlib import Path
from typing import Dict
import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QFrame,
    QLabel, QLineEdit, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QGraphicsDropShadowEffect, QPushButton
)
from PySide6.QtGui import QFont, QColor
from PySide6.QtCore import Qt
from logic.calculo_cargas import calcular_tabla_cargas


# ────────── lee la columna A del Excel ──────────
def load_equipo_options(path: Path | None = None, sheet: str = "MUEBLESFRIOS") -> list[str]:
    book = path or Path(__file__).resolve().parents[2] / "data" / "basedatos.xlsx"
    if not book.exists():
        print(f"[WARN] {book} no encontrado – combos vacíos"); return []
    try:
        df = pd.read_excel(book, sheet_name=sheet, usecols="A", header=None)
    except ValueError:
        print(f"[INFO] Hoja “{sheet}” no hallada – usando primera hoja")
        df = pd.read_excel(book, sheet_name=0, usecols="A", header=None)
    return (df.iloc[:, 0].dropna().astype(str).str.strip().unique().tolist())


class CargaPage(QWidget):
    # ——————————————— INIT ———————————————
    def __init__(self):
        super().__init__()
        self.equipo_opts = load_equipo_options()
        self.table = None
        self._lock = False          # evita bucles temp
        self._build_ui()

    # ———————————— UI ————————————
    def _build_ui(self):
        root = QVBoxLayout(self); root.setContentsMargins(40, 40, 40, 40); root.setSpacing(20)

        # === Formulario oscuro ===
        form = QFrame(); form.setStyleSheet("background:#2a2d3c;border-radius:12px;")
        fl = QVBoxLayout(form); fl.setContentsMargins(30, 30, 30, 30); fl.setAlignment(Qt.AlignTop); fl.setSpacing(8)

        title = QLabel("CARGA TÉRMICA"); title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:24px;font-weight:700;color:#e0e0e0;"); fl.addWidget(title)

        top = QHBoxLayout()
        top.addWidget(QLabel("Nombre del proyecto:")); self.proj = QLineEdit(); top.addWidget(self.proj)
        fl.addLayout(top)

        g = QGridLayout(); g.setHorizontalSpacing(20)
        self.city, self.resp = QLineEdit(), QLineEdit()
        self.rb, self.rm    = QLineEdit(), QLineEdit()
        self.long_cbo = QComboBox(); self.long_cbo.addItems(["SI", "NO"])
        self.unit_cbo = QComboBox(); self.unit_cbo.addItems(["Sistema Internacional", "Sistema Inglés"])

        g.addWidget(QLabel("Ciudad:"), 0, 0); g.addWidget(self.city, 0, 1)
        g.addWidget(QLabel("Responsable:"), 0, 2); g.addWidget(self.resp, 0, 3)
        g.addWidget(QLabel("Nº de ramales baja:"), 1, 0); g.addWidget(self.rb, 1, 1)
        g.addWidget(QLabel("Nº de ramales media:"), 1, 2); g.addWidget(self.rm, 1, 3)
        g.addWidget(QLabel("Longitudes ramales:"), 2, 0); g.addWidget(self.long_cbo, 2, 1)
        g.addWidget(QLabel("Tipo de unidades:"), 2, 2); g.addWidget(self.unit_cbo, 2, 3)

        g.addWidget(QLabel("Temperatura ambiente:"), 3, 0)
        tmp = QWidget(); tl = QHBoxLayout(tmp); tl.setContentsMargins(0, 0, 0, 0); tl.setSpacing(4)
        tl.addWidget(QLabel("°C")); self.t_c = QLineEdit(); self.t_c.setMaximumWidth(60); tl.addWidget(self.t_c)
        tl.addWidget(QLabel("°F")); self.t_f = QLineEdit(); self.t_f.setMaximumWidth(60); tl.addWidget(self.t_f)
        g.addWidget(tmp, 3, 1)

        fl.addLayout(g)

        # BOTÓN GENERAR TABLA
        gen_btn = QPushButton("GENERAR TABLA")
        gen_btn.setStyleSheet("""
            QPushButton{background:#5b8bea;color:#fff;font-weight:600;
                        border:none;border-radius:6px;padding:8px 18px;}
            QPushButton:hover{background:#6c9cf0;}""")
        gen_btn.clicked.connect(self._refresh_table)
        fl.addWidget(gen_btn, alignment=Qt.AlignRight)

        root.addWidget(form); root.addSpacing(16)

        # === Card blanco para la tabla ===
        self.table_box = QFrame()
        self.table_box.setStyleSheet(
            "QFrame{background:#ffffff;border:1px solid #d0d0d0;border-radius:8px;}")
        self.table_box.setGraphicsEffect(
            QGraphicsDropShadowEffect(blurRadius=12,xOffset=0,yOffset=3,color=QColor(0,0,0,60)))
        self.table_box.setLayout(QVBoxLayout()); self.table_box.layout().setContentsMargins(0,0,0,0)
        root.addWidget(self.table_box, 1)

        # — eventos solo para temperatura (sin refrescar tabla) —
        self.t_c.editingFinished.connect(lambda: self._convert_temp("C"))
        self.t_f.editingFinished.connect(lambda: self._convert_temp("F"))

    # ————————— helpers rápidos —————————
    @staticmethod
    def _int_or_none(txt: str): return int(txt) if txt.strip().isdigit() else None
    def _suffix(self): return "(m)" if self.unit_cbo.currentText()=="Sistema Internacional" else "(ft)"

    def _convert_temp(self, changed: str):
        if self._lock: return
        try:
            if changed == "C" and self.t_c.text().strip():
                c = float(self.t_c.text()); f = c * 9/5 + 32
                self._lock = True; self.t_f.setText(f"{f:.1f}"); self._lock = False
            elif changed == "F" and self.t_f.text().strip():
                f = float(self.t_f.text()); c = (f - 32) * 5/9
                self._lock = True; self.t_c.setText(f"{c:.1f}"); self._lock = False
        except ValueError:
            pass   # números inválidos → ignorar

    # combo EQUIPO
    def _combo_equipo(self, val=""):
        cb = QComboBox(); cb.addItems(self.equipo_opts)
        cb.setStyleSheet("""
            QComboBox{background:#ffffff;color:#000;border:1px solid #b0b0b0;padding:2px 6px;}
            QComboBox::drop-down{width:18px;}
            QComboBox QAbstractItemView{background:#ffffff;color:#000;selection-background-color:#cbe4ff;}""")
        if val: ix = cb.findText(val); cb.setCurrentIndex(ix if ix >= 0 else 0)
        return cb

    def _set_ramal(self, row, idx, grp, equipo_prev):
        self.table.setItem(row, 0, QTableWidgetItem(f"{idx} ({grp})"))
        self.table.item(row, 0).setFlags(Qt.ItemIsEnabled)
        self.table.setCellWidget(row, 1, self._combo_equipo(equipo_prev))

    # snapshot para conservar datos
    def _snapshot(self) -> Dict[str, Dict[int, str]]:
        snap = {}
        if not self.table: return snap
        for r in range(5, self.table.rowCount()):
            key_item = self.table.item(r, 0)
            if not key_item: continue
            k = key_item.text()
            row = {}
            combo = self.table.cellWidget(r, 1)
            if combo: row[1] = combo.currentText()
            for c in range(2, self.table.columnCount()):
                it = self.table.item(r, c)
                if it and it.text(): row[c] = it.text()
            if row: snap[k] = row
        return snap

    # ————————— genera / refresca tabla (solo con el botón) —————————
    def _refresh_table(self):
        nb, nm = self._int_or_none(self.rb.text()), self._int_or_none(self.rm.text())
        if nb is None or nm is None:
            return   # no hace nada si campos vacíos

        old = self._snapshot()
        filas = calcular_tabla_cargas(nb, nm)
        long_flag = self.long_cbo.currentText() == "SI"

        headers = ["RAMAL", "EQUIPO", f"DIMENSIONES {self._suffix()}",
                   "USO", "CARGA (BTU/h)", "CARGA (kW)"] + (["LONGITUD RAMAL A RACK"] if long_flag else [])

        rows = 5 + len(filas) + 1
        self.table = QTableWidget(rows, len(headers))
        self.table.verticalHeader().setVisible(False); self.table.horizontalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget{background:#ffffff;color:#000;gridline-color:#b0b0b0;border:none;}
            QTableWidget::item{padding:4px;} QTableWidget::item:alternate{background:#f5f7fb;}
            QTableWidget::item:selected{background:#cbe4ff;color:#000;}""")
        bold = QFont(); bold.setBold(True)

        # filas de información
        info = [
            self.proj.text().upper(),
            f"CIUDAD: {self.city.text().upper()}",
            f"RESPONSABLE: {self.resp.text().upper()}",
            f"TEMPERATURA DE CONDENSACIÓN: {(self.t_f.text() if self.unit_cbo.currentText()=='Sistema Inglés' else self.t_c.text())} "
            f"{'°F' if self.unit_cbo.currentText()=='Sistema Inglés' else '°C'}"
        ]
        for r, txt in enumerate(info):
            self.table.setSpan(r, 0, 1, len(headers))
            it = QTableWidgetItem(txt); it.setFont(bold); it.setFlags(Qt.ItemIsEnabled)
            if r == 0: it.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 0, it)

        # cabeceras
        for c, htxt in enumerate(headers):
            h = QTableWidgetItem(htxt); h.setFont(bold); h.setFlags(Qt.ItemIsEnabled); h.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(4, c, h)

        row = 5
        for idx, grp in filas[:nb]:
            prev_eq = old.get(f"{idx} ({grp})", {}).get(1, "")
            self._set_ramal(row, idx, grp, prev_eq); row += 1

        # separador TOTAL BAJA
        self.table.setSpan(row, 0, 1, len(headers))
        sep = QTableWidgetItem("—  TOTAL BAJA  —"); sep.setFont(bold); sep.setFlags(Qt.ItemIsEnabled); sep.setTextAlignment(Qt.AlignCenter)
        self.table.setItem(row, 0, sep); row += 1

        for idx, grp in filas[nb:]:
            prev_eq = old.get(f"{idx} ({grp})", {}).get(1, "")
            self._set_ramal(row, idx, grp, prev_eq); row += 1

        # restaura texto en otras columnas
        for r in range(5, self.table.rowCount()):
            k_item = self.table.item(r, 0)
            if not k_item: continue
            key = k_item.text()
            if key not in old: continue
            for c, txt in old[key].items():
                if c == 1: continue  # combo ya restaurado
                if c < self.table.columnCount():
                    self.table.setItem(r, c, QTableWidgetItem(txt))

        # muestra tabla
        lay = self.table_box.layout()
        while lay.count(): lay.takeAt(0).widget().deleteLater()
        lay.addWidget(self.table)
