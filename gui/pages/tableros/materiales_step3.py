from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGridLayout, QFrame, QLabel, QComboBox,
    QScrollArea, QHBoxLayout, QRadioButton, QButtonGroup, QSpinBox
)


def _norm(s: str) -> str:
    t = (s or "").strip().upper()
    return t.translate(str.maketrans("ÁÉÍÓÚÜÑ", "AEIOUUN"))


class Step3OptionsPanel(QWidget):
    """
    PASO 3 — OPCIONES CO2 (dos columnas).
    Lee hoja 'OPCIONES CO2' y genera:
      - radio (SI/NO)       -> 2 RadioButtons
      - spin ('#' en hoja)  -> QSpinBox 0..5
      - combo               -> QComboBox
    """
    def __init__(self) -> None:
        super().__init__()

        self._inputs: Dict[str, Dict[str, object]] = {}
        self._loaded: bool = False  # ← para no reconstruir al volver

        # ----- contenedor con scroll -----
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)

        host = QFrame()
        host.setStyleSheet("QFrame{background:#ffffff;border:1px solid #d7e3f8;border-radius:12px;}")

        host_layout = QVBoxLayout(host)
        host_layout.setContentsMargins(18, 18, 18, 18)
        host_layout.setSpacing(14)

        title = QLabel("OPCIONES CO2")
        title.setStyleSheet("color:#0f172a;font-weight:800;font-size:16px;")
        host_layout.addWidget(title)

        # Panel con dos columnas
        self._panel = QFrame()
        self._panel.setStyleSheet("QFrame{background:#f7f9fd;border:1px solid #e2e8f5;border-radius:10px;}")
        self._gridCols = QGridLayout(self._panel)
        self._gridCols.setContentsMargins(12, 12, 12, 12)
        self._gridCols.setHorizontalSpacing(14)
        self._gridCols.setVerticalSpacing(14)

        self._colL = QFrame(); self._colR = QFrame()
        self._colL_layout = QVBoxLayout(self._colL); self._colR_layout = QVBoxLayout(self._colR)
        for lay in (self._colL_layout, self._colR_layout):
            lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(12)

        self._gridCols.addWidget(self._colL, 0, 0)
        self._gridCols.addWidget(self._colR, 0, 1)
        self._gridCols.setColumnStretch(0, 1); self._gridCols.setColumnStretch(1, 1)

        host_layout.addWidget(self._panel, 1)
        self._scroll.setWidget(host)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(self._scroll, 1)

    # ---------- API pública ----------
    def is_loaded(self) -> bool:
        return self._loaded

    def load_options(self, initial_state: Dict[str, str] | None = None, force: bool = False) -> None:
        """
        Construye los controles **solo la primera vez**.
        Si 'force=True' sí reconstruye (y reimporta si viene initial_state).
        """
        if self._loaded and not force:
            # Si pasan estado inicial aunque ya esté cargado, lo aplico igual
            if initial_state:
                try: self.import_state(initial_state)
                except Exception: pass
            return

        preguntas = self._read_from_excel()
        self._rebuild_form(preguntas)
        self._loaded = True

        if initial_state:
            try: self.import_state(initial_state)
            except Exception: pass

    def export_state(self) -> Dict[str, str]:
        out: Dict[str, str] = {}
        for q, meta in self._inputs.items():
            typ = meta.get("type")
            if typ == "radio":
                rb_si: QRadioButton = meta["rb_si"]  # type: ignore
                rb_no: QRadioButton = meta["rb_no"]  # type: ignore
                out[q] = "SI" if rb_si.isChecked() else ("NO" if rb_no.isChecked() else "")
            elif typ == "spin":
                sp: QSpinBox = meta["spin"]  # type: ignore
                out[q] = str(sp.value())
            else:
                cb: QComboBox = meta["combo"]  # type: ignore
                out[q] = cb.currentText().strip()
        return out

    def import_state(self, state: Dict[str, str]) -> None:
        if not isinstance(state, dict) or not self._inputs:
            return
        inputs_by_key = { _norm(k): (k, v) for k, v in self._inputs.items() }
        for k, raw in state.items():
            key_norm = _norm(k)
            pair = inputs_by_key.get(key_norm)
            if not pair:
                continue
            _, meta = pair
            typ = meta.get("type")
            if typ == "radio":
                val = _norm(str(raw))
                rb_si: QRadioButton = meta["rb_si"]  # type: ignore
                rb_no: QRadioButton = meta["rb_no"]  # type: ignore
                if val in ("SI", "SÍ", "TRUE", "1"): rb_si.setChecked(True)
                elif val in ("NO", "FALSE", "0"): rb_no.setChecked(True)
            elif typ == "spin":
                try: n = int(str(raw).strip() or "0")
                except Exception: n = 0
                sp: QSpinBox = meta["spin"]  # type: ignore
                sp.setValue(max(sp.minimum(), min(sp.maximum(), n)))
            else:
                cb: QComboBox = meta["combo"]  # type: ignore
                txt = "" if raw is None else str(raw)
                i = cb.findText(txt, Qt.MatchFixedString)
                if i < 0:
                    for j in range(cb.count()):
                        if txt.lower() == cb.itemText(j).lower(): i = j; break
                if i < 0:
                    for j in range(cb.count()):
                        if txt.lower() in cb.itemText(j).lower(): i = j; break
                if i >= 0: cb.setCurrentIndex(i)

    # ---------- lectura Excel ----------
    def _read_from_excel(self) -> List[Tuple[str, List[str], str]]:
        book = Path(__file__).resolve().parents[3] / "data" / "basedatos.xlsx"
        items: List[Tuple[str, List[str], str]] = []
        try:
            df = pd.read_excel(book, sheet_name="OPCIONES CO2", header=None, dtype=object)
        except Exception:
            return [("No encontré la hoja 'OPCIONES CO2' en basedatos.xlsx", [""], "combo")]

        for i in range(len(df.index)):
            q = self._cell(df, i, 0)
            if not q: continue
            b = self._cell(df, i, 1); c = self._cell(df, i, 2)

            if ("#" in b) or ("#" in c):
                items.append((q, [str(k) for k in range(6)], "spin")); continue

            raw: List[str] = []
            for s in (b, c):
                if not s: continue
                raw += [p.strip() for p in str(s).replace("|", ",").replace("/", ",").split(",") if p.strip()]

            if not raw:
                items.append((q, ["SI", "NO"], "radio")); continue

            seen, opts = set(), []
            for x in raw:
                up = x.upper()
                if up not in seen:
                    seen.add(up); opts.append(x)
            up_opts = [o.upper() for o in opts]
            if up_opts in (["SI", "NO"], ["NO", "SI"]):
                items.append((q, ["SI", "NO"], "radio"))
            else:
                items.append((q, opts, "combo"))
        if not items:
            items = [("La hoja 'OPCIONES CO2' está vacía", [""], "combo")]
        return items

    @staticmethod
    def _cell(df: pd.DataFrame, r: int, c: int) -> str:
        if r >= len(df.index) or c >= df.shape[1]: return ""
        v = df.iat[r, c]
        return "" if pd.isna(v) else str(v).strip()

    # ---------- UI ----------
    def _clear_columns(self) -> None:
        for lay in (self._colL_layout, self._colR_layout):
            while lay.count():
                item = lay.takeAt(0)
                w = item.widget()
                if w: w.deleteLater()
        self._inputs.clear()

    def _rebuild_form(self, preguntas: List[Tuple[str, List[str], str]]) -> None:
        self._clear_columns()
        to_left = True
        for q, opts, typ in preguntas:
            row = self._build_row(q, opts, typ)
            (self._colL_layout if to_left else self._colR_layout).addWidget(row)
            to_left = not to_left
        self._colL_layout.addStretch(1); self._colR_layout.addStretch(1)

    def _build_row(self, pregunta: str, opts: List[str], typ: str) -> QWidget:
        row = QFrame()
        row.setStyleSheet(
            "QFrame{ background:#ffffff; border:1px solid #d7e3f8; border-radius:10px; }"
        )
        h = QHBoxLayout(row); h.setContentsMargins(12, 10, 12, 10); h.setSpacing(14)

        lab = QLabel(pregunta)
        lab.setStyleSheet("color:#0f172a; font-weight:700;")
        lab.setMinimumWidth(240); lab.setMaximumWidth(360)
        h.addWidget(lab, 0, Qt.AlignVCenter)

        radio_style = (
            "QRadioButton{color:#0f172a;font-weight:600;}"
            "QRadioButton::indicator{width:18px;height:18px;border-radius:9px;"
            "border:2px solid #5b8bea;background:#ffffff;}"
            "QRadioButton::indicator:checked{background:#4cc9ff;border-color:#4cc9ff;}"
        )

        if typ == "radio":
            rb_si = QRadioButton("SI"); rb_no = QRadioButton("NO")
            for rb in (rb_si, rb_no): rb.setStyleSheet(radio_style)
            grp = QButtonGroup(row); grp.addButton(rb_si); grp.addButton(rb_no)
            wrap = QFrame(); hw = QHBoxLayout(wrap); hw.setContentsMargins(0,0,0,0); hw.setSpacing(12)
            hw.addWidget(rb_si); hw.addWidget(rb_no); hw.addStretch(1)
            h.addWidget(wrap, 1)
            self._inputs[pregunta] = {"type":"radio","rb_si":rb_si,"rb_no":rb_no,"group":grp}
        elif typ == "spin":
            sp = QSpinBox(); sp.setRange(0,5); sp.setFixedWidth(90)
            sp.setStyleSheet("QSpinBox{background:#ffffff;color:#0f172a;border:1px solid #c8d4eb;border-radius:8px;padding:4px 6px;}")
            h.addWidget(sp, 0, Qt.AlignLeft)
            self._inputs[pregunta] = {"type":"spin","spin":sp}
        else:
            cb = QComboBox(); cb.addItem(""); cb.addItems([str(o) for o in opts]); cb.setMaximumWidth(240)
            cb.setStyleSheet("QComboBox{background:#ffffff;color:#0f172a;border:1px solid #c8d4eb;border-radius:8px;padding:6px 8px;}")
            h.addWidget(cb, 0, Qt.AlignLeft)
            self._inputs[pregunta] = {"type":"combo","combo":cb}
        return row

