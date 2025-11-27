# src/gui/pages/carga_page.py
# ——————————————————————————————————————————————————————————
# Página «Cargas Térmicas»  ·  versión 18-jun-2025
#
#  ✔ Botón «GENERAR TABLA» (sin refresco continuo).
#  ✔ Conversión °C ↔ °F instantánea.
#  ✔ Numeración visible: 1B… nB —TOTAL BAJA— (n+1)M…
#  ✔ Al reducir Nº de ramales (Baja / Media) solo se descarta el último
#    del grupo modificado; se conservan los datos restantes.
#  ✔ Combos EQUIPO buscables (editable + autocomplete *MatchContains*)
#      – arrancan vacíos;
#      – el texto se valida también al salir con **TAB**:
#        si hay coincidencia parcial toma el primer resultado,
#        si no hay ninguna coincidencia se borra.
#  ✔ Copiar / pegar (Ctrl +C / Ctrl +V)
#      – si el portapapeles contiene **un solo valor**, se replica en
#        todas las celdas/combos seleccionados;
#      – si contiene un bloque TSV (filas × columnas) se pega como Excel.
# ——————————————————————————————————————————————————————————
from __future__ import annotations
from logic.dev_autofill import apply_autofill_globals
from pathlib import Path
from typing import Dict

import pandas as pd
from PySide6.QtCore import Qt, QEvent
from PySide6.QtGui import (
    QColor,
    QFont,
    QKeySequence,
    QClipboard,
)
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QCompleter,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QHeaderView,
)

from logic.calculo_cargas import calcular_tabla_cargas


# ————————————————— utilidades —————————————————
def load_equipo_options(
    path: Path | None = None,
    sheet: str = "MUEBLESFRIOS",
) -> list[str]:
    book = (
        path
        or Path(__file__).resolve().parents[3] / "data" / "basedatos.xlsx"
    )
    if not book.exists():
        print(f"[WARN] {book} no encontrado – combos vacíos")
        return []
    try:
        df = pd.read_excel(book, sheet_name=sheet, usecols="A", header=None)
    except ValueError:
        df = pd.read_excel(book, sheet_name=0, usecols="A", header=None)
    return (
        df.iloc[:, 0]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )


def _int(txt: str | None) -> int | None:
    return int(txt) if txt and txt.strip().isdigit() else None


# ————————————————— clase página —————————————————
class CargaPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._lock = False
        self._snapshot: Dict[str, Dict[int, str]] = {}
        self.equipo_opts = load_equipo_options()
        self.table: QTableWidget | None = None
        self._build_ui()
        self.installEventFilter(self)          # para Ctrl-C / Ctrl-V

    # ——————————— UI ———————————
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(40, 40, 40, 40)
        root.setSpacing(20)

        # ===== formulario =====
        form = QFrame()
        form.setStyleSheet("background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;")
        fl = QVBoxLayout(form); fl.setAlignment(Qt.AlignTop); fl.setSpacing(8)
        fl.setContentsMargins(30, 30, 30, 30)

        title = QLabel("CARGA TÉRMICA")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:24px;font-weight:700;color:#0f172a;")
        fl.addWidget(title)

        top = QHBoxLayout()
        top.addWidget(QLabel("Nombre del proyecto:")); self.proj = QLineEdit(); top.addWidget(self.proj)
        fl.addLayout(top)

        g = QGridLayout(); g.setHorizontalSpacing(20)
        self.city, self.resp = QLineEdit(), QLineEdit()
        self.rb, self.rm = QLineEdit(), QLineEdit()
        self.long_cbo = QComboBox(); self.long_cbo.addItems(["SI", "NO"])
        self.unit_cbo = QComboBox(); self.unit_cbo.addItems(["Sistema Internacional", "Sistema Inglés"])

        g.addWidget(QLabel("Ciudad:"),              0, 0); g.addWidget(self.city, 0, 1)
        g.addWidget(QLabel("Responsable:"),         0, 2); g.addWidget(self.resp, 0, 3)
        g.addWidget(QLabel("Nº de ramales baja:"),  1, 0); g.addWidget(self.rb,   1, 1)
        g.addWidget(QLabel("Nº de ramales media:"), 1, 2); g.addWidget(self.rm,   1, 3)
        g.addWidget(QLabel("Longitudes ramales:"),  2, 0); g.addWidget(self.long_cbo, 2, 1)
        g.addWidget(QLabel("Tipo de unidades:"),    2, 2); g.addWidget(self.unit_cbo, 2, 3)

        g.addWidget(QLabel("Temperatura ambiente:"), 3, 0)
        tmp = QWidget(); tl = QHBoxLayout(tmp); tl.setContentsMargins(0, 0, 0, 0); tl.setSpacing(4)
        tl.addWidget(QLabel("°C")); self.t_c = QLineEdit(); self.t_c.setMaximumWidth(60); tl.addWidget(self.t_c)
        tl.addWidget(QLabel("°F")); self.t_f = QLineEdit(); self.t_f.setMaximumWidth(60); tl.addWidget(self.t_f)
        g.addWidget(tmp, 3, 1)
        fl.addLayout(g)

        gen_btn = QPushButton("GENERAR TABLA"); gen_btn.setStyleSheet("""
            QPushButton{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #0f62fe, stop:1 #22d3ee);color:#fff;font-weight:700;border:none;border-radius:8px;padding:10px 18px;}
            QPushButton:hover{background:#6c9cf0;}""")
        gen_btn.clicked.connect(self._refresh_table)
        fl.addWidget(gen_btn, alignment=Qt.AlignRight)

        root.addWidget(form); root.addSpacing(16)

        # ===== contenedor blanco =====
        self.table_box = QFrame(); self.table_box.setStyleSheet(
            "QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;}")
        self.table_box.setGraphicsEffect(
            QGraphicsDropShadowEffect(blurRadius=12, xOffset=0, yOffset=3, color=QColor(0, 0, 0, 60)))
        self.table_box.setLayout(QVBoxLayout()); self.table_box.layout().setContentsMargins(0, 0, 0, 0)
        root.addWidget(self.table_box, 1)

        self.t_c.editingFinished.connect(lambda: self._convert_temp("C"))
        self.t_f.editingFinished.connect(lambda: self._convert_temp("F"))

    # ————————— conversión °C/°F —————————
    def _convert_temp(self, changed: str) -> None:
        if self._lock: return
        try:
            if changed == "C" and self.t_c.text().strip():
                f = float(self.t_c.text()) * 9 / 5 + 32
                self._lock = True; self.t_f.setText(f"{f:.1f}"); self._lock = False
            elif changed == "F" and self.t_f.text().strip():
                c = (float(self.t_f.text()) - 32) * 5 / 9
                self._lock = True; self.t_c.setText(f"{c:.1f}"); self._lock = False
        except ValueError:
            pass

    # ————————— combo EQUIPO buscable —————————
    def _combo_equipo(self, val: str = "") -> QComboBox:
        cb = QComboBox(); cb.addItems(self.equipo_opts)
        cb.setEditable(True); cb.setInsertPolicy(QComboBox.NoInsert)

        comp: QCompleter = cb.completer()
        comp.setCaseSensitivity(Qt.CaseInsensitive); comp.setFilterMode(Qt.MatchContains)

        cb.setStyleSheet("""
            QComboBox{background:#ffffff;color:#0f172a;border:1px solid #b0b0b0;padding:2px 6px;}
            QComboBox::drop-down{width:18px;}
            QComboBox QAbstractItemView{background:#ffffff;color:#0f172a;selection-background-color:#cbe4ff;}
        """)

        # — validación —
        def _best_match(txt: str) -> int:
            """Devuelve índice de la primera opción que *contenga* txt (case-insens)."""
            t = txt.lower()
            for i in range(cb.count()):
                if t in cb.itemText(i).lower():
                    return i
            return -1

        def _validate() -> None:
            txt = cb.lineEdit().text().strip()
            if not txt:
                cb.setCurrentIndex(-1); return
            ix = cb.findText(txt, Qt.MatchFixedString | Qt.MatchCaseSensitive)
            if ix < 0:
                ix = _best_match(txt)
            if ix >= 0:
                cb.setCurrentIndex(ix)
                cb.lineEdit().setText(cb.itemText(ix))
            else:
                cb.setCurrentIndex(-1); cb.lineEdit().clear()

        cb.lineEdit().editingFinished.connect(_validate)

        def _event_filter(obj, ev):
            if obj is cb.lineEdit() and ev.type() == QEvent.KeyPress and ev.key() == Qt.Key_Tab:
                _validate()
            return False
        cb.lineEdit().installEventFilter(cb); cb.eventFilter = _event_filter

        ix = cb.findText(val, Qt.MatchFixedString)
        cb.setCurrentIndex(ix if ix >= 0 else -1)
        return cb

    # ————————— snapshot —————————
    def _take_snapshot(self) -> None:
        self._snapshot.clear()
        if not self.table: return
        for r in range(5, self.table.rowCount()):
            it = self.table.item(r, 0)
            if not it: continue
            canon = it.data(Qt.UserRole)
            if not canon: continue
            row: Dict[int, str] = {}
            combo = self.table.cellWidget(r, 1)
            if combo: row[1] = combo.currentText()
            for c in range(2, self.table.columnCount()):
                cell = self.table.item(r, c)
                if cell and cell.text(): row[c] = cell.text()
            if row: self._snapshot[canon] = row

    # ————————— fila de ramal —————————
    def _set_ramal(self, row: int, vis: str, canon: str) -> None:
        it = QTableWidgetItem(vis); it.setFlags(Qt.ItemIsEnabled); it.setData(Qt.UserRole, canon)
        self.table.setItem(row, 0, it)
        prev = self._snapshot.get(canon, {}).get(1, "")
        self.table.setCellWidget(row, 1, self._combo_equipo(prev))

    # ————————— reconstrucción de la tabla —————————
    def _refresh_table(self) -> None:
        nb, nm = _int(self.rb.text()), _int(self.rm.text())
        if nb is None or nm is None: return

        self._take_snapshot()
        self._snapshot = {k: v for k, v in self._snapshot.items()
                          if not ((k.startswith('B') and int(k[1:]) > nb) or
                                  (k.startswith('M') and int(k[1:]) > nm))}

        long_flag = self.long_cbo.currentText() == "SI"
        headers = ["RAMAL", "EQUIPO",
                   f"DIMENSIONES {'(m)' if self.unit_cbo.currentText() == 'Sistema Internacional' else '(ft)'}",
                   "USO", "CARGA (BTU/h)", "CARGA (kW)"] + (["LONGITUD RAMAL A RACK"] if long_flag else [])

        sep_row = 1 if nb and nm else 0
        rows = 5 + nb + nm + sep_row
        self.table = QTableWidget(rows, len(headers))
        self.table.verticalHeader().setVisible(False); self.table.horizontalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget{background:#ffffff;color:#0f172a;gridline-color:#d7deeb;border:1px solid #e2e8f5;}
            QTableWidget::item{padding:4px;} QTableWidget::item:alternate{background:#f5f7fb;}
            QTableWidget::item:selected{background:#cbe4ff;color:#000;}""")
        bold = QFont(); bold.setBold(True)

        info = [self.proj.text().upper(),
                f"CIUDAD: {self.city.text().upper()}",
                f"RESPONSABLE: {self.resp.text().upper()}",
                f"TEMPERATURA DE CONDENSACIÓN: "
                f"{(self.t_f.text() if self.unit_cbo.currentText() == 'Sistema Inglés' else self.t_c.text())} "
                f"{'°F' if self.unit_cbo.currentText() == 'Sistema Inglés' else '°C'}"]
        for r, txt in enumerate(info):
            self.table.setSpan(r, 0, 1, len(headers))
            it = QTableWidgetItem(txt); it.setFont(bold); it.setFlags(Qt.ItemIsEnabled)
            if r == 0: it.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 0, it)

        for c, htxt in enumerate(headers):
            h = QTableWidgetItem(htxt); h.setFont(bold); h.setFlags(Qt.ItemIsEnabled); h.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(4, c, h)

        seq, row = 1, 5
        for idx in range(1, nb + 1):
            self._set_ramal(row, f"{seq}B", f"B{idx}"); row += 1; seq += 1
        if nb and nm:
            self.table.setSpan(row, 0, 1, len(headers))
            sep = QTableWidgetItem("—  TOTAL BAJA  —"); sep.setFont(bold); sep.setFlags(Qt.ItemIsEnabled); sep.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, sep); row += 1
        for idx in range(1, nm + 1):
            self._set_ramal(row, f"{seq}M", f"M{idx}"); row += 1; seq += 1

        for r in range(5, self.table.rowCount()):
            canon = self.table.item(r, 0).data(Qt.UserRole)
            if canon in self._snapshot:
                for c, txt in self._snapshot[canon].items():
                    if c == 1: continue
                    if c < self.table.columnCount():
                        if not self.table.item(r, c): self.table.setItem(r, c, QTableWidgetItem())
                        self.table.item(r, c).setText(txt)

        lay = self.table_box.layout()
        while lay.count(): (lay.takeAt(0).widget() or QWidget()).deleteLater()
        lay.addWidget(self.table)

    # ————————— Ctrl-C / Ctrl-V —————————
    def eventFilter(self, obj, ev):                    # pylint: disable=invalid-name
        if obj is self and self.table and ev.type() == QEvent.KeyPress:
            if ev.matches(QKeySequence.Copy):
                self._copy_selection(); return True
            if ev.matches(QKeySequence.Paste):
                self._paste_selection(); return True
        return super().eventFilter(obj, ev)

    def _copy_selection(self) -> None:
        sel = self.table.selectedIndexes()
        if not sel: return
        rows = sorted({i.row() for i in sel}); cols = sorted({i.column() for i in sel})
        data_lines = []
        for r in rows:
            line = []
            for c in cols:
                if c == 1:
                    combo: QComboBox = self.table.cellWidget(r, 1)
                    line.append(combo.currentText() if combo else "")
                else:
                    it = self.table.item(r, c)
                    line.append(it.text() if it else "")
            data_lines.append("\t".join(line))
        QApplication.clipboard().setText("\n".join(data_lines), QClipboard.Clipboard)

    def _paste_selection(self) -> None:
        clip = QApplication.clipboard().text()
        if not clip: return
        data_rows = [row.split("\t") for row in clip.splitlines()]
        sel = self.table.selectedIndexes()
        if not sel: return
        sel_rows = sorted({i.row() for i in sel}); sel_cols = sorted({i.column() for i in sel})

        # Si el portapapeles es un único valor, réplica en cada celda seleccionada
        if len(data_rows) == 1 and len(data_rows[0]) == 1:
            val = data_rows[0][0]
            for i in sel:
                self._set_cell(i.row(), i.column(), val)
            return

        # De lo contrario, pega como Excel starting at top-left of selection
        start_r, start_c = sel_rows[0], sel_cols[0]
        for r_off, row_vals in enumerate(data_rows):
            for c_off, val in enumerate(row_vals):
                self._set_cell(start_r + r_off, start_c + c_off, val)

    def _set_cell(self, r: int, c: int, val: str) -> None:
        """Coloca `val` en (r,c) respetando combos de EQUIPO."""
        if r >= self.table.rowCount() or c >= self.table.columnCount(): return
        if c == 1:
            combo: QComboBox = self.table.cellWidget(r, 1)
            ix = combo.findText(val, Qt.MatchFixedString)
            combo.setCurrentIndex(ix if ix >= 0 else -1)
        else:
            if not self.table.item(r, c):
                self.table.setItem(r, c, QTableWidgetItem())
            self.table.item(r, c).setText(val)

