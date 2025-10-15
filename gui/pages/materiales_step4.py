# gui/pages/materiales_step4.py
from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict, List, Optional
import re
import unicodedata

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFrame, QLabel, QScrollArea, QGridLayout,
    QPushButton, QFileDialog, QHBoxLayout
)

from logic.step4_engine import Step4Engine
from logic.step4_compresores import extract_comp_meta
from logic.opciones_co2_engine import ResumenTable
from logic.resumen_cables import build_cable_badges
from logic.export_step4 import export_step4_excel_only, save_programacion_exact


class MaterialesStep4Page(QWidget):
    """
    Paso 4 — Resumen unificado (COMPRESORES + OTROS + BORNERAS).
    ITEM | CÓDIGO | MODELO | NOMBRE | DESCRIPCIÓN | C 240 V (kA) | C 480 V (kA) | REFERENCIA | TORQUE
    """

    def __init__(
        self,
        get_step2_state: Callable[[], Dict[str, Dict[str, str]]] = lambda: {},
        get_globals: Callable[[], Dict[str, object]] = lambda: {},
    ) -> None:
        super().__init__()
        self._get_step2_state = get_step2_state
        self._get_globals = get_globals

        self._book = Path(__file__).resolve().parents[2] / "data" / "basedatos.xlsx"
        self.engine = Step4Engine(self._book)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)

        self.host = QFrame()
        self.host.setStyleSheet("QFrame{background:transparent;}")
        self.host_layout = QVBoxLayout(self.host)
        self.host_layout.setContentsMargins(0, 0, 0, 0)
        self.host_layout.setSpacing(12)

        self.scroll.setWidget(self.host)
        root.addWidget(self.scroll)

    # -------------------------------------------------------------- API
    def reload_and_render(self) -> None:
        # limpia
        while self.host_layout.count():
            w = self.host_layout.takeAt(0).widget()
            if w:
                w.deleteLater()

        step2 = self._get_step2_state() or {}
        globs = self._get_globals() or {}

        try:
            res = self.engine.calcular(step2, globs)
        except Exception as e:
            self._add_info_box(f"No pude calcular el Resumen del Paso 4. Detalle: {e}")
            self.host_layout.addStretch(1)
            return

        # badges de cable (una vez)
        cable_badges = build_cable_badges(step2, self._book)

        # =========== 1) Tablas por compresor ===========
        tables: List[ResumenTable] = res.get("tables_compresores", [])
        if not tables:
            self._add_info_box("NO HAY COMPRESORES CONFIGURADOS EN EL PASO 2.")
        else:
            for t in tables:
                st = step2.get(t.comp_key, {}) or {}
                brand, model, amps = extract_comp_meta(st, globs)

                header_main = f"{t.comp_key} — ARRANQUE: {t.arranque}".upper()
                subtitle_parts = []
                name_part = f"{(brand or '').strip()} {(model or '').strip()}".strip()
                if name_part:
                    subtitle_parts.append(f"COMPRESOR: {name_part}")
                if amps:
                    subtitle_parts.append(f"CORRIENTE: {amps} A")

                base = "   |   ".join(subtitle_parts) if subtitle_parts else ""
                extra = cable_badges.get(t.comp_key, "")
                subtitle = (base + extra) if (base or extra) else None

                self._render_group_table(header_main, t.rows, subtitle=subtitle)

        # =========== 2) OTROS ELEMENTOS (FIJOS) ===========
        otros_rows: List[List[str]] = self._ensure_item_col(res.get("otros_rows", []))
        if otros_rows:
            self._render_group_table("OTROS ELEMENTOS (FIJOS)".upper(), otros_rows)

        # =========== 3) BORNERAS ===========
        b_comp = res.get("borneras_compresores", {"fase": 0, "neutro": 0, "tierra": 0})
        b_otros = res.get("borneras_otros", {"fase": 0, "neutro": 0, "tierra": 0})
        b_total = res.get("borneras_total", {"fase": 0, "neutro": 0, "tierra": 0})

        self._render_borneras_card("BORNERAS — COMPRESORES (AX/AY/AZ)", b_comp["fase"], b_comp["neutro"], b_comp["tierra"])
        self._render_borneras_card("BORNERAS — OTROS (W/X/Y)", b_otros["fase"], b_otros["neutro"], b_otros["tierra"])
        self._render_borneras_card("BORNERAS — TOTAL PROYECTO", b_total["fase"], b_total["neutro"], b_total["tierra"])

        # =========== 4) Botones: Imprimir Excel / Guardar programación ===========
        self._render_export_buttons(step2, globs)

        if not tables and not otros_rows:
            self._add_info_box("SIN RESULTADOS PARA MOSTRAR.")

        self.host_layout.addStretch(1)

    # ---------------- UI helpers ----------------
    def _ensure_item_col(self, rows: List[List[str]]) -> List[List[str]]:
        """
        Si una fila de 'otros' trae 8 columnas (sin ITEM al inicio),
        usa la REFERENCIA (col 7 en ese formato) como ITEM.
        """
        fixed: List[List[str]] = []
        for r in rows or []:
            if len(r) == 8:
                item = r[6] if len(r) > 6 else ""
                fixed.append([item] + r)
            else:
                fixed.append(r)
        return fixed

    def _render_group_table(self, title: str, rows: List[List[str]], subtitle: Optional[str] = None) -> None:
        card = QFrame()
        card.setStyleSheet("QFrame{background:#2f3242;border:1px solid #42465a;border-radius:8px;}")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(12, 12, 12, 12)
        cl.setSpacing(8)

        tlab = QLabel(title)
        tlab.setStyleSheet("color:#e0e0e0;font-weight:700;")
        tlab.setWordWrap(True)
        cl.addWidget(tlab)

        if subtitle:
            slab = QLabel(subtitle)
            slab.setStyleSheet("color:#000000;")  # NEGRO para impresión legible
            slab.setWordWrap(True)
            cl.addWidget(slab)

        table = QFrame()
        table.setStyleSheet("QFrame{background:#2a2d3c;border:1px solid #3d4154;border-radius:6px;}")
        grid = QGridLayout(table)
        grid.setContentsMargins(8, 8, 8, 8)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)

        headers = ["ITEM", "CÓDIGO", "MODELO", "NOMBRE", "DESCRIPCIÓN", "C 240 V (kA)", "C 480 V (kA)", "REFERENCIA", "TORQUE"]
        for c, h in enumerate(headers):
            lab = QLabel(h)
            lab.setStyleSheet("color:#aeb7d6;font-weight:700;")
            grid.addWidget(lab, 0, c)

        for r, row in enumerate(rows, start=1):
            for c, val in enumerate(row):
                lab = QLabel(val or "")
                lab.setStyleSheet("color:#e0e0e0;")
                grid.addWidget(lab, r, c)

        grid.setColumnStretch(0, 1); grid.setColumnStretch(1, 1); grid.setColumnStretch(2, 2)
        grid.setColumnStretch(3, 2); grid.setColumnStretch(4, 3); grid.setColumnStretch(5, 1)
        grid.setColumnStretch(6, 1); grid.setColumnStretch(7, 2); grid.setColumnStretch(8, 1)

        cl.addWidget(table)
        self.host_layout.addWidget(card)

    def _render_borneras_card(self, title: str, fase: int, neutro: int, tierra: int) -> None:
        card = QFrame(); card.setStyleSheet("QFrame{background:#2f3242;border:1px solid #42465a;border-radius:8px;}")
        cl = QVBoxLayout(card); cl.setContentsMargins(12, 12, 12, 12); cl.setSpacing(8)

        tlab = QLabel(title); tlab.setStyleSheet("color:#e0e0e0;font-weight:700;"); tlab.setWordWrap(True)
        cl.addWidget(tlab)

        table = QFrame(); table.setStyleSheet("QFrame{background:#2a2d3c;border:1px solid #3d4154;border-radius:6px;}")
        grid = QGridLayout(table); grid.setContentsMargins(8, 8, 8, 8); grid.setHorizontalSpacing(14); grid.setVerticalSpacing(6)

        h1 = QLabel("TIPO"); h1.setStyleSheet("color:#aeb7d6;font-weight:700;")
        h2 = QLabel("TOTAL"); h2.setStyleSheet("color:#aeb7d6;font-weight:700;")
        grid.addWidget(h1, 0, 0); grid.addWidget(h2, 0, 1)

        def add_row(r: int, label: str, value: int) -> None:
            l1 = QLabel(label); l1.setStyleSheet("color:#e0e0e0;")
            l2 = QLabel(str(value)); l2.setStyleSheet("color:#e0e0e0; font-weight:700;")
            grid.addWidget(l1, r, 0); grid.addWidget(l2, r, 1)

        add_row(1, "FASE",   fase)
        add_row(2, "NEUTRO", neutro)
        add_row(3, "TIERRA", tierra)

        grid.setColumnStretch(0, 3); grid.setColumnStretch(1, 1)
        cl.addWidget(table)
        self.host_layout.addWidget(card)

    def _render_export_buttons(self, step2: Dict[str, Dict[str, str]], globs: Dict[str, object]) -> None:
        card = QFrame()
        card.setStyleSheet("QFrame{background:#2f3242;border:1px solid #42465a;border-radius:8px;}")
        cl = QVBoxLayout(card); cl.setContentsMargins(12, 12, 12, 12); cl.setSpacing(8)

        tlab = QLabel("Exportar / Programación")
        tlab.setStyleSheet("color:#e0e0e0;font-weight:700;")
        cl.addWidget(tlab)

        row = QHBoxLayout()
        btn_excel = QPushButton("Imprimir (Excel Resumen)")
        btn_excel.setStyleSheet(
            "QPushButton{background:#4a8bff;color:white;border:none;border-radius:8px;padding:10px 14px;font-weight:700;}"
            "QPushButton:pressed{background:#3c72d1;}"
        )
        btn_snap = QPushButton("Guardar programación")
        btn_snap.setStyleSheet(
            "QPushButton{background:#3a3f55;color:#e0e0e0;border:1px solid #4a4f67;border-radius:8px;padding:10px 14px;font-weight:700;}"
            "QPushButton:pressed{background:#2f3346;}"
        )

        def do_excel():
            base_dir = Path.home() / "Documents" / "ElectroCalc_Exports"
            base_dir.mkdir(parents=True, exist_ok=True)
            out_dir = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta para Excel", str(base_dir))
            if not out_dir:
                return
            try:
                path = export_step4_excel_only(self._book, step2, globs, out_dir)
                self._add_info_box(f"✅ Excel generado:\n{path}")
            except Exception as e:
                self._add_info_box(f"❌ Error al generar Excel: {e}")

        def _slug_project_name(raw: str) -> str:
            if not raw:
                return "PROYECTO"
            s = unicodedata.normalize("NFKD", raw).encode("ascii", "ignore").decode("ascii")
            s = s.upper().strip()
            s = re.sub(r"[^A-Z0-9._-]+", "_", s)
            s = re.sub(r"__+", "_", s).strip("_")
            return s or "PROYECTO"

        def do_snapshot():
            base_dir = Path.home() / "Documents" / "ElectroCalc_Exports"
            base_dir.mkdir(parents=True, exist_ok=True)

            # nombre sugerido a partir del proyecto
            raw_proj = (globs.get("nombre_proyecto") or globs.get("proyecto") or "").strip()
            suggested = _slug_project_name(raw_proj) + ".ecalc.json"

            fname, _ = QFileDialog.getSaveFileName(
                self, "Guardar programación",
                str(base_dir / suggested),
                "Snapshot (*.ecalc.json)"
            )
            if not fname:
                return
            p = Path(fname)
            if not p.name.lower().endswith(".ecalc.json"):
                p = p.with_name(p.stem + ".ecalc.json")

            try:
                save_programacion_exact(step2, globs, p)
                self._add_info_box(f"✅ Programación guardada:\n{p}\n\n(Úsalo en Paso 1 para recargar)")
            except Exception as e:
                self._add_info_box(f"❌ Error al guardar programación: {e}")

        btn_excel.clicked.connect(do_excel)
        btn_snap.clicked.connect(do_snapshot)
        row.addWidget(btn_excel)
        row.addWidget(btn_snap)
        cl.addLayout(row)

        self.host_layout.addWidget(card)

    def _add_info_box(self, text: str) -> None:
        box = QFrame()
        box.setStyleSheet("QFrame{background:#2f3242;border:1px solid #42465a;border-radius:8px;}")
        bl = QVBoxLayout(box); bl.setContentsMargins(12, 12, 12, 12); bl.setSpacing(6)
        lab = QLabel(text); lab.setStyleSheet("color:#e0e0e0;"); lab.setWordWrap(True); bl.addWidget(lab)
        self.host_layout.addWidget(box)
