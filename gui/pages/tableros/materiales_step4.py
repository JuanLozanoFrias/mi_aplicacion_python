from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict, List, Optional
from datetime import date
import re
import unicodedata

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFrame, QLabel, QScrollArea, QGridLayout,
    QPushButton, QFileDialog, QHBoxLayout, QMessageBox, QInputDialog
)

from logic.step4_engine import Step4Engine
from logic.step4_compresores import extract_comp_meta
from logic.opciones_co2_engine import ResumenTable
from logic.resumen_cables import build_cable_badges
from logic.export_step4 import export_step4_excel_only, save_programacion_exact


class MaterialesStep4Page(QWidget):
    """
    Paso 4 – Resumen unificado (COMPRESORES + OTROS + BORNERAS).
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

        self._book = Path(__file__).resolve().parents[3] / "data" / "basedatos.xlsx"
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

        cable_badges = build_cable_badges(step2, self._book)

        # =========== 1) Tablas por compresor ===========
        tables: List[ResumenTable] = res.get("tables_compresores", [])
        if not tables:
            self._add_info_box("NO HAY COMPRESORES CONFIGURADOS EN EL PASO 2.")
        else:
            for t in tables:
                st = step2.get(t.comp_key, {}) or {}
                brand, model, amps = extract_comp_meta(st, globs)

                header_main = f"{t.comp_key} · ARRANQUE: {t.arranque}".upper()
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

        self._render_borneras_card("BORNERAS - COMPRESORES (AX/AY/AZ)", b_comp["fase"], b_comp["neutro"], b_comp["tierra"])
        self._render_borneras_card("BORNERAS - OTROS (W/X/Y)", b_otros["fase"], b_otros["neutro"], b_otros["tierra"])
        self._render_borneras_card("BORNERAS - TOTAL PROYECTO", b_total["fase"], b_total["neutro"], b_total["tierra"])

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
        card.setStyleSheet("QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:10px;}")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(12, 12, 12, 12)
        cl.setSpacing(8)

        tlab = QLabel(title)
        tlab.setStyleSheet("color:#0f172a;font-weight:800;font-size:15px;")
        tlab.setWordWrap(True)
        cl.addWidget(tlab)

        if subtitle:
            slab = QLabel(subtitle)
            slab.setStyleSheet("color:#1f2937;")
            slab.setWordWrap(True)
            cl.addWidget(slab)

        table = QFrame()
        table.setStyleSheet("QFrame{background:#f7f9fd;border:1px solid #e2e8f5;border-radius:8px;}")
        grid = QGridLayout(table)
        grid.setContentsMargins(8, 8, 8, 8)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)

        headers = ["ITEM", "CÓDIGO", "MODELO", "NOMBRE", "DESCRIPCIÓN", "C 240 V (kA)", "C 480 V (kA)", "REFERENCIA", "TORQUE"]
        for c, h in enumerate(headers):
            lab = QLabel(h)
            lab.setStyleSheet("color:#475569;font-weight:700;")
            grid.addWidget(lab, 0, c)

        for r, row in enumerate(rows, start=1):
            for c, val in enumerate(row):
                lab = QLabel(val or "")
                lab.setStyleSheet("color:#0f172a;")
                grid.addWidget(lab, r, c)

        grid.setColumnStretch(0, 1); grid.setColumnStretch(1, 1); grid.setColumnStretch(2, 2)
        grid.setColumnStretch(3, 2); grid.setColumnStretch(4, 3); grid.setColumnStretch(5, 1)
        grid.setColumnStretch(6, 1); grid.setColumnStretch(7, 2); grid.setColumnStretch(8, 1)

        cl.addWidget(table)
        self.host_layout.addWidget(card)

    def _render_borneras_card(self, title: str, fase: int, neutro: int, tierra: int) -> None:
        card = QFrame(); card.setStyleSheet("QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:10px;}")
        cl = QVBoxLayout(card); cl.setContentsMargins(12, 12, 12, 12); cl.setSpacing(8)

        tlab = QLabel(title); tlab.setStyleSheet("color:#0f172a;font-weight:800;"); tlab.setWordWrap(True)
        cl.addWidget(tlab)

        table = QFrame(); table.setStyleSheet("QFrame{background:#f7f9fd;border:1px solid #e2e8f5;border-radius:8px;}")
        grid = QGridLayout(table); grid.setContentsMargins(8, 8, 8, 8); grid.setHorizontalSpacing(14); grid.setVerticalSpacing(6)

        h1 = QLabel("TIPO"); h1.setStyleSheet("color:#475569;font-weight:700;")
        h2 = QLabel("TOTAL"); h2.setStyleSheet("color:#475569;font-weight:700;")
        grid.addWidget(h1, 0, 0); grid.addWidget(h2, 0, 1)

        def add_row(r: int, label: str, value: int) -> None:
            l1 = QLabel(label); l1.setStyleSheet("color:#0f172a;")
            l2 = QLabel(str(value)); l2.setStyleSheet("color:#0f172a; font-weight:800;")
            grid.addWidget(l1, r, 0); grid.addWidget(l2, r, 1)

        add_row(1, "FASE",   fase)
        add_row(2, "NEUTRO", neutro)
        add_row(3, "TIERRA", tierra)

        grid.setColumnStretch(0, 3); grid.setColumnStretch(1, 1)
        cl.addWidget(table)
        self.host_layout.addWidget(card)

    def _render_export_buttons(
        self, step2: Dict[str, Dict[str, str]], globs: Dict[str, object]
    ) -> None:
        card = QFrame()
        card.setStyleSheet("QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:10px;}")
        cl = QVBoxLayout(card); cl.setContentsMargins(12, 12, 12, 12); cl.setSpacing(8)

        tlab = QLabel("Exportar / Programacion")
        tlab.setStyleSheet("color:#0f172a;font-weight:800;")
        cl.addWidget(tlab)

        row = QHBoxLayout()
        btn_export_all = QPushButton("Exportar Excel + Programacion")
        btn_export_all.setMinimumHeight(42)
        btn_style = (
            "QPushButton{background:#5b8bea;color:#ffffff;font-weight:700;"
            "border:none;border-radius:10px;padding:10px 16px;}"
            "QPushButton:hover{background:#6c9cf0;}"
            "QPushButton:pressed{background:#4a7bd9;}"
        )
        btn_export_all.setStyleSheet(btn_style)

        def do_export_all():
            base_dir = Path.home() / "Documents" / "ElectroCalc_Exports"
            base_dir.mkdir(parents=True, exist_ok=True)
            out_dir = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de destino", str(base_dir))
            if not out_dir:
                return

            raw_proj = (globs.get("nombre_proyecto") or globs.get("proyecto") or "").strip()
            def _slug_project_name(raw: str) -> str:
                s = unicodedata.normalize("NFKD", raw or "").encode("ascii", "ignore").decode("ascii")
                s = s.upper().strip()
                s = re.sub(r"[^A-Z0-9._-]+", "_", s)
                s = re.sub(r"__+", "_", s).strip("_")
                return s or "PROYECTO"

            today = date.today().strftime("%Y%m%d")
            base_name = f"{today}_{_slug_project_name(raw_proj)}"

            out_dir_path = Path(out_dir)

            try:
                excel_generated = Path(export_step4_excel_only(self._book, step2, globs, out_dir_path))
                target_excel = excel_generated.with_name(f"{base_name}{excel_generated.suffix or '.xlsx'}")
                if excel_generated != target_excel:
                    try:
                        target_excel.unlink(missing_ok=True)
                    except Exception:
                        pass
                    excel_generated.replace(target_excel)
            except Exception as e:
                msg = f"Ups... Error al generar Excel: {e}"
                self._add_info_box(msg)
                QMessageBox.critical(self, "Exportar", msg)
                return

            try:
                snap_path = out_dir_path / f"{base_name}.ecalc.json"
                save_programacion_exact(step2, globs, snap_path)
            except Exception as e:
                msg = f"Ups... Error al guardar programacion: {e}"
                self._add_info_box(msg)
                QMessageBox.critical(self, "Exportar", msg)
                return

            msg = (
                "Listo. Se exportaron ambos archivos:\n"
                f"Excel: {target_excel}\n"
                f"Programacion: {snap_path}\n\n"
                "Usa el .ecalc en Paso 1 para recargar"
            )
            self._add_info_box(msg)
            QMessageBox.information(self, "Exportar", msg)

        btn_export_all.clicked.connect(do_export_all)
        row.addWidget(btn_export_all)
        cl.addLayout(row)

        self.host_layout.addWidget(card)

    def _add_info_box(self, text: str) -> None:
        box = QFrame()
        box.setStyleSheet("QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:10px;}")
        bl = QVBoxLayout(box); bl.setContentsMargins(12, 12, 12, 12); bl.setSpacing(6)
        lab = QLabel(text); lab.setStyleSheet("color:#0f172a;"); lab.setWordWrap(True); bl.addWidget(lab)

        insert_at = self.host_layout.count()
        if insert_at:
            last = self.host_layout.itemAt(insert_at - 1)
            if last is not None and last.spacerItem() is not None:
                insert_at -= 1
        self.host_layout.insertWidget(insert_at, box)

