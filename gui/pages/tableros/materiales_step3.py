from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Tuple
from PySide6.QtCore import Signal
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGridLayout, QFrame, QLabel, QComboBox,
    QHBoxLayout, QButtonGroup, QSpinBox, QPushButton
)


def _norm(s: str) -> str:
    return (s or "").strip().upper()


class _NoWheelCombo(QComboBox):
    def wheelEvent(self, event) -> None:  # type: ignore[override]
        event.ignore()

    def mousePressEvent(self, event) -> None:  # type: ignore[override]
        if event.button() == Qt.MiddleButton:
            event.ignore()
            return
        super().mousePressEvent(event)


class _NoWheelSpin(QSpinBox):
    def wheelEvent(self, event) -> None:  # type: ignore[override]
        event.ignore()

    def mousePressEvent(self, event) -> None:  # type: ignore[override]
        if event.button() == Qt.MiddleButton:
            event.ignore()
            return
        super().mousePressEvent(event)
class YesNoToggle(QFrame):
    """Boton doble SI/NO con estilo similar al modulo de carga electrica."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        self.btn_si = QPushButton("SI")
        self.btn_no = QPushButton("NO")
        for b in (self.btn_si, self.btn_no):
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.setMinimumWidth(52)

        self.group = QButtonGroup(self)
        self.group.setExclusive(True)
        self.group.addButton(self.btn_si)
        self.group.addButton(self.btn_no)

        self.btn_si.clicked.connect(lambda: self._set(True))
        self.btn_no.clicked.connect(lambda: self._set(False))

        lay.addWidget(self.btn_si)
        lay.addWidget(self.btn_no)
        lay.addStretch(1)

        self._set(False)

    def _refresh(self) -> None:
        active = "background:#5b8bea;color:#fff;font-weight:700;border:none;border-radius:8px;padding:6px 10px;"
        inactive = "background:#eef3ff;color:#0f172a;font-weight:600;border:1px solid #b9d1ff;border-radius:8px;padding:6px 10px;"
        self.btn_si.setStyleSheet(active if self.btn_si.isChecked() else inactive)
        self.btn_no.setStyleSheet(active if self.btn_no.isChecked() else inactive)

    def _set(self, val: bool) -> None:
        self.btn_si.setChecked(bool(val))
        self.btn_no.setChecked(not bool(val))
        self._refresh()

    def set_value(self, val) -> None:
        s = str(val).strip().upper()
        if s in {"SI", "S?", "TRUE", "1", "YES"}:
            self._set(True)
        elif s in {"NO", "FALSE", "0"}:
            self._set(False)
        else:
            self._set(False)

    def value(self) -> str:
        return "SI" if self.btn_si.isChecked() else ("NO" if self.btn_no.isChecked() else "")


class Step3OptionsPanel(QWidget):
    """
    PASO 3 - OPCIONES CO2 (dos columnas).
    Lee archivo `data/preguntas_opciones_co2.json` (cache) y genera:
      - radio (SI/NO)       -> 2 RadioButtons
      - spin ('#' en hoja)  -> QSpinBox 0..5
      - combo               -> QComboBox
    """

    changed = Signal()
    def __init__(self) -> None:
        super().__init__()

        self._inputs: Dict[str, Dict[str, object]] = {}
        self._loaded: bool = False  # para no reconstruir al volver
        self._pending_all_yes: bool = True

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

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(host, 1)

    # ---------- API pública ----------
    def is_loaded(self) -> bool:
        return self._loaded

    def clear_all(self) -> None:
        """
        Limpia controles y marca para reconstruir en la siguiente carga.
        Deja pendiente el 'todo en SI' para el arranque inicial.
        """
        self._clear_columns()
        self._loaded = False
        self._pending_all_yes = True
        self.update()

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

        preguntas = self._read_from_cache()
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
                toggle: YesNoToggle = meta["toggle"]  # type: ignore
                out[q] = toggle.value()
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
                toggle: YesNoToggle = meta["toggle"]  # type: ignore
                toggle.set_value(raw)
            elif typ == "spin":
                try:
                    n = int(str(raw).strip() or "0")
                except Exception:
                    n = 0
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
    # ---------- Reset helpers ----------
    def set_all_yes(self) -> None:
        """
        Coloca todos los toggles SI/NO en "SI".
        Si a?n no est? cargado, marca la acci?n para aplicarla tras load_options.
        """
        if not self._loaded or not self._inputs:
            self._pending_all_yes = True
            return
        changed_any = False
        for meta in self._inputs.values():
            if meta.get("type") == "radio":
                toggle: YesNoToggle = meta.get("toggle")  # type: ignore
                if toggle is None:
                    continue
                prev = toggle.value()
                toggle.set_value("SI")
                changed_any = changed_any or (prev != toggle.value())
        if changed_any:
            try:
                self.changed.emit()
            except Exception:
                pass

    def _read_from_cache(self) -> List[Tuple[str, List[str], str]]:
        """
        Lee preguntas desde `data/preguntas_opciones_co2.json`.

        Formato esperado:
            {"version":1,"items":[{"pregunta":"...","tipo":"radio|spin|combo","opciones":[...]}]}
        """
        cache_path = Path(__file__).resolve().parents[3] / "data" / "tableros_electricos" / "preguntas_opciones_co2.json"
        try:
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
            items_in = payload.get("items", []) if isinstance(payload, dict) else []
        except Exception:
            return [(f"No encontré `preguntas_opciones_co2.json` en {cache_path.parent}", [""], "combo")]

        items: List[Tuple[str, List[str], str]] = []
        for it in items_in:
            if not isinstance(it, dict):
                continue
            q = str(it.get("pregunta") or "").strip()
            if not q:
                continue
            tipo = str(it.get("tipo") or "").strip().lower()
            opciones = it.get("opciones", [])

            if tipo == "radio":
                items.append((q, ["SI", "NO"], "radio"))
                continue
            if tipo == "spin":
                items.append((q, [str(x) for x in (opciones or [0, 1, 2, 3, 4, 5])], "spin"))
                continue

            # combo
            opts = [str(x).strip() for x in (opciones or []) if str(x).strip()]
            items.append((q, opts, "combo"))

        if not items:
            items = [("`preguntas_opciones_co2.json` está vacío", [""], "combo")]
        return items

    # ---------- UI ----------
    def _clear_columns(self) -> None:
        for lay in (self._colL_layout, self._colR_layout):
            while lay.count():
                item = lay.takeAt(0)
                w = item.widget()
                if w: w.deleteLater()
        self._inputs.clear()
        self._pending_all_yes = False

    def _rebuild_form(self, preguntas: List[Tuple[str, List[str], str]]) -> None:
        self._clear_columns()
        to_left = True
        for q, opts, typ in preguntas:
            row = self._build_row(q, opts, typ)
            (self._colL_layout if to_left else self._colR_layout).addWidget(row)
            to_left = not to_left
        self._colL_layout.addStretch(1); self._colR_layout.addStretch(1)
        if self._pending_all_yes:
            self.set_all_yes()
            self._pending_all_yes = False

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

        if typ == "radio":
            toggle = YesNoToggle()
            toggle.btn_si.clicked.connect(lambda: self.changed.emit())
            toggle.btn_no.clicked.connect(lambda: self.changed.emit())
            h.addWidget(toggle, 1)
            self._inputs[pregunta] = {"type":"radio","toggle":toggle}
        elif typ == "spin":
            sp = _NoWheelSpin(); sp.setRange(0,5); sp.setFixedWidth(90)
            sp.setStyleSheet("QSpinBox{background:#ffffff;color:#0f172a;border:1px solid #c8d4eb;border-radius:8px;padding:4px 6px;}")
            h.addWidget(sp, 0, Qt.AlignLeft)
            sp.valueChanged.connect(lambda *_: self.changed.emit())
            self._inputs[pregunta] = {"type":"spin","spin":sp}
        else:
            cb = _NoWheelCombo(); cb.addItem(""); cb.addItems([str(o) for o in opts]); cb.setMaximumWidth(240)
            cb.setStyleSheet("QComboBox{background:#ffffff;color:#0f172a;border:1px solid #c8d4eb;border-radius:8px;padding:6px 8px;}")
            h.addWidget(cb, 0, Qt.AlignLeft)
            cb.currentIndexChanged.connect(lambda *_: self.changed.emit())
            self._inputs[pregunta] = {"type":"combo","combo":cb}
        return row
