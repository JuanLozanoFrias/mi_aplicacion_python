from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict, List, Optional
from datetime import date
import re
import shutil
import unicodedata

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFrame, QLabel, QGridLayout,
    QPushButton, QFileDialog, QHBoxLayout, QMessageBox, QInputDialog,
    QCheckBox
)
from PySide6.QtCore import Qt

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
        self._comp_enabled: Dict[str, bool] = {}
        self._comp_rows: Dict[str, Dict[str, QLabel]] = {}
        self._current_comp_det: List[Dict[str, object]] = []
        self._current_globs: Dict[str, object] = {}
        self._total_label: Optional[QLabel] = None
        self._simple_label: Optional[QLabel] = None
        self._breaker_label: Optional[QLabel] = None

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        self.host = QFrame()
        self.host.setStyleSheet("QFrame{background:transparent;}")
        self.host_layout = QVBoxLayout(self.host)
        self.host_layout.setContentsMargins(0, 0, 0, 0)
        self.host_layout.setSpacing(12)

        root.addWidget(self.host)

    # -------------------------------------------------------------- API
    def reload_and_render(self) -> None:
        # limpia
        while self.host_layout.count():
            w = self.host_layout.takeAt(0).widget()
            if w:
                w.deleteLater()
        # reset de switches de compresores para evitar estados arrastrados
        self._comp_enabled.clear()
        self._comp_rows = {}
        self._current_comp_det = []
        self._amps_map = {}

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

        # =========== 3) CORRIENTE TOTAL DEL SISTEMA + BREAKER ==============
        corriente_info = res.get("corriente_total", {})
        if corriente_info.get("found"):
            self._render_corriente_card(corriente_info)

        # =========== 3) BORNERAS ===========
        b_comp = res.get("borneras_compresores", {"fase": 0, "neutro": 0, "tierra": 0})
        b_otros = res.get("borneras_otros", {"fase": 0, "neutro": 0, "tierra": 0})
        b_total = res.get("borneras_total", {"fase": 0, "neutro": 0, "tierra": 0})

        self._render_borneras_card("BORNERAS - COMPRESORES (AX/AY/AZ)", b_comp["fase"], b_comp["neutro"], b_comp["tierra"])
        self._render_borneras_card("BORNERAS - OTROS (W/X/Y)", b_otros["fase"], b_otros["neutro"], b_otros["tierra"])
        self._render_borneras_card("BORNERAS - TOTAL PROYECTO", b_total["fase"], b_total["neutro"], b_total["tierra"])

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

    def _render_corriente_card(self, info: Dict[str, object]) -> None:
        det = info.get("detalle") or {}
        breaker = info.get("breaker") or {}
        comp_det = info.get("comp_detalles") or []
        self._current_comp_det = comp_det
        self._current_globs = self._get_globals() or {}
        self._comp_rows = {}
        # mapa de corrientes originales
        self._amps_map = {str(d.get("comp_key", "")): float(d.get("amps", 0.0)) for d in comp_det}

        card = QFrame()
        card.setStyleSheet("QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:10px;}")
        card.setMaximumWidth(820)
        cl = QVBoxLayout(card)
        cl.setContentsMargins(12, 12, 12, 12)
        cl.setSpacing(8)

        tlab = QLabel("CORRIENTE TOTAL DEL SISTEMA")
        tlab.setStyleSheet("color:#0f172a;font-weight:800;font-size:15px;")
        cl.addWidget(tlab)

        # Prefijar estado de switches
        for d in comp_det:
            k = d.get("comp_key", "")
            if k and k not in self._comp_enabled:
                self._comp_enabled[k] = True

        # Aplicar filtros de compresores activos y recalcular totales locales
        det, breaker = self._recalc_totales_local(comp_det, self._current_globs)

        # Tabla de corrientes por compresor
        if comp_det:
            comp_frame = QFrame()
            comp_frame.setStyleSheet("QFrame{background:#f7f9fd;border:1px solid #e2e8f5;border-radius:8px;}")
            cg = QGridLayout(comp_frame)
            cg.setContentsMargins(8, 8, 8, 8)
            cg.setHorizontalSpacing(10)
            cg.setVerticalSpacing(4)

            headers = ["USAR", "COMPRESOR", "CORRIENTE (A)", "CORR. AJUSTADA (A)"]
            for c, h in enumerate(headers):
                lab = QLabel(h)
                lab.setStyleSheet("color:#475569;font-weight:700;")
                cg.addWidget(lab, 0, c)

            for r, d in enumerate(comp_det, start=1):
                k = d.get("comp_key", "")
                a = d.get("amps", 0.0)

                chk = QCheckBox()
                chk.setChecked(self._comp_enabled.get(k, True))
                chk.setStyleSheet(
                    "QCheckBox::indicator {width:18px; height:18px; border:2px solid #2563eb; "
                    "border-radius:6px; background:#f8fafc;} "
                    "QCheckBox::indicator:checked {background:qlineargradient(x1:0,y1:0,x2:1,y2:1,"
                    "stop:0 #38bdf8, stop:1 #0ea5e9); border-color:#1d4ed8;}"
                )
                chk.stateChanged.connect(lambda state, key=k: self._on_toggle_comp(key, state))

                lbl_corr = QLabel(f"{a:.2f}")
                lbl_corr_adj = QLabel("")  # se llena en _update

                cg.addWidget(chk, r, 0, alignment=Qt.AlignCenter)
                modelo = str(d.get("modelo", "") or "").strip()
                comp_txt = f"{k.upper()} - {modelo}".upper() if modelo else k.upper()
                cg.addWidget(QLabel(comp_txt), r, 1)
                cg.addWidget(lbl_corr, r, 2)
                cg.addWidget(lbl_corr_adj, r, 3)

                self._comp_rows[k] = {"base": lbl_corr, "adj": lbl_corr_adj, "chk": chk, "amp": float(a)}

            cg.setColumnStretch(0, 0); cg.setColumnStretch(1, 2); cg.setColumnStretch(2, 1); cg.setColumnStretch(3, 1)
            cl.addWidget(comp_frame)

        # Resumen compacto solo con total ajustado
        resumen = QFrame()
        resumen.setStyleSheet("QFrame{background:#f7f9fd;border:1px solid #e2e8f5;border-radius:8px;}")
        rg = QGridLayout(resumen)
        rg.setContentsMargins(8, 8, 8, 8)
        rg.setHorizontalSpacing(12)
        rg.setVerticalSpacing(6)

        lab_total = QLabel("TOTAL AJUSTADO")
        lab_total.setStyleSheet("color:#475569;font-weight:700;")
        self._total_label = QLabel(f"{det.get('total', 0.0):.2f} A")
        self._total_label.setStyleSheet("color:#0f172a;font-weight:800;font-size:14px;")
        rg.addWidget(lab_total, 0, 0)
        rg.addWidget(self._total_label, 0, 1)
        lab_simple = QLabel("TOTAL SIN AJUSTE")
        lab_simple.setStyleSheet("color:#475569;font-weight:700;")
        self._simple_label = QLabel(f"{det.get('suma_simple', 0.0):.2f} A")
        self._simple_label.setStyleSheet("color:#334155;font-weight:700;")
        rg.addWidget(lab_simple, 1, 0)
        rg.addWidget(self._simple_label, 1, 1)

        rg.setColumnStretch(0, 2); rg.setColumnStretch(1, 1)
        cl.addWidget(resumen)

        # Breaker totalizador
        if breaker.get("found"):
            modelo_txt = (breaker.get('modelo') or "").strip()
            info_txt = f"BREAKER TOTALIZADOR: {modelo_txt} ({breaker.get('amp',0)} A)"
            if breaker.get("codigo"):
                info_txt += f"  - CÓDIGO: {breaker.get('codigo')}"
        else:
            info_txt = breaker.get("motivo", "No se encontró breaker >= corriente calculada")

        self._breaker_label = QLabel(info_txt)
        self._breaker_label.setStyleSheet("color:#0f172a;font-weight:700;")
        self._breaker_label.setWordWrap(True)

        cl.addWidget(self._breaker_label)

        self.host_layout.addWidget(card, alignment=Qt.AlignHCenter)

        # inicializa textos de corrientes ajustadas y totales
        self._update_comp_totals()

    def _on_toggle_comp(self, key: str, state: int) -> None:
        self._comp_enabled[key] = (state == Qt.Checked)
        # Actualiza en caliente sin recalcular el engine
        self._update_comp_totals()

    def _recalc_totales_local(self, comp_det: List[Dict[str, object]], globs: Dict[str, object]) -> tuple[Dict[str, float], Dict[str, object]]:
        """Recalcula totales con el subconjunto de compresores activos."""
        if not comp_det:
            det = {"mayor": 0.0, "ajuste_mayor": 0.0, "suma_restantes": 0.0, "total": 0.0, "suma_simple": 0.0}
            return det, {"found": False, "motivo": "Sin compresores activos"}
        # tomar lista de (amps, es_mayor) respetando marca de mayor original
        if any(d.get("ajustado") for d in comp_det):
            mayor = next((d for d in comp_det if d.get("ajustado")), comp_det[0])
        else:
            mayor = max(comp_det, key=lambda x: x.get("amps", 0.0))
        max_i = float(mayor.get("amps", 0.0))
        suma_rest = sum(float(d.get("amps", 0.0)) for d in comp_det if d is not mayor)
        ajuste = max_i * 0.25
        total = max_i + ajuste + suma_rest
        suma_simple = sum(float(d.get("amps", 0.0)) for d in comp_det)

        norma = (globs.get("norma_ap") or "IEC").upper()
        breaker = self.engine._pick_breaker_total(total, norma)  # reutilizamos lógica existente

        det = {
            "mayor": max_i,
            "ajuste_mayor": ajuste,
            "suma_restantes": suma_rest,
            "total": total,
            "suma_simple": suma_simple,
        }
        return det, breaker

    def _update_comp_totals(self) -> None:
        """Actualiza la tabla en caliente sin reconstruirla completa."""
        # Mapa de corrientes originales por compresor
        amps_map = {str(d.get("comp_key", "")): float(d.get("amps", 0.0)) for d in self._current_comp_det}

        # Datos activos
        activos = []
        for k, info in self._comp_rows.items():
            chk: QCheckBox = info.get("chk")
            enabled = chk.isChecked() if chk else self._comp_enabled.get(k, True)
            self._comp_enabled[k] = enabled
            if not enabled:
                continue
            a = info.get("amp")
            if a is None:
                a = amps_map.get(k, 0.0)
            activos.append((k, float(a or 0.0)))

        if not activos:
            mayor_key = None
            det = {"total": 0.0, "suma_simple": 0.0}
            breaker = {"found": False, "motivo": "Sin compresores activos"}
        else:
            mayor_key, max_i = max(activos, key=lambda x: x[1])
            suma_simple = sum(a for _, a in activos)
            suma_rest = suma_simple - max_i
            ajuste = max_i * 0.25
            total = max_i + ajuste + suma_rest
            det = {"total": total, "suma_simple": suma_simple}
            norma = (self._current_globs.get("norma_ap") or "IEC").upper()
            breaker = self.engine._pick_breaker_total(total, norma)

        # Actualiza filas
        for k, info in self._comp_rows.items():
            lbl_base = info.get("base")
            lbl_adj = info.get("adj")
            amp = info.get("amp")
            if amp is None:
                amp = amps_map.get(k, 0.0)
            amp = float(amp or 0.0)
            enabled = self._comp_enabled.get(k, True)
            if not lbl_base or not lbl_adj:
                continue
            if not enabled:
                lbl_base.setText("-"); lbl_base.setStyleSheet("color:#94a3b8;")
                lbl_adj.setText("-"); lbl_adj.setStyleSheet("color:#94a3b8;")
            else:
                lbl_base.setText(f"{amp:.2f}"); lbl_base.setStyleSheet("color:#0f172a;")
                if k == mayor_key:
                    amp_adj = amp * 1.25
                    lbl_adj.setText(f"{amp_adj:.2f}"); lbl_adj.setStyleSheet("color:#0f172a;font-weight:800;")
                else:
                    lbl_adj.setText(f"{amp:.2f}"); lbl_adj.setStyleSheet("color:#0f172a;")

        if self._total_label:
            self._total_label.setText(f"{det.get('total', 0.0):.2f} A")
        if self._simple_label:
            self._simple_label.setText(f"{det.get('suma_simple', 0.0):.2f} A")

        if self._breaker_label:
            if breaker.get("found"):
                modelo_txt = (breaker.get('modelo') or "").strip()
                info_txt = f"BREAKER TOTALIZADOR: {modelo_txt} ({breaker.get('amp',0)} A)"
                if breaker.get("codigo"):
                    info_txt += f"  - CÓDIGO: {breaker.get('codigo')}"
            else:
                info_txt = breaker.get("motivo", "No se encontró breaker >= corriente calculada")
            self._breaker_label.setText(info_txt)

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

    # -------------------------------------------------------------- Helpers externos
    def export_project(self) -> None:
        """
        Exporta Excel + programación con un diálogo rápido.
        """
        step2 = self._get_step2_state() or {}
        globs = self._get_globals() or {}
        if not step2:
            QMessageBox.information(self, "Exportar", "No hay datos para exportar todavía.")
            return

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

        def _unique_path(p: Path) -> Path:
            if not p.exists():
                return p
            for i in range(2, 10_000):
                cand = p.with_name(f"{p.stem}_{i}{p.suffix}")
                if not cand.exists():
                    return cand
            return p

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
            QMessageBox.critical(self, "Exportar", f"Error al generar Excel:\n{e}")
            return

        # Guardar snapshot SOLO en la biblioteca interna
        lib_snap_path: Optional[Path] = None
        try:
            lib_dir = Path(__file__).resolve().parents[3] / "data" / "proyectos" / "tableros"
            lib_dir.mkdir(parents=True, exist_ok=True)
            lib_snap_path = _unique_path(lib_dir / f"{base_name}.ecalc.json")
            save_programacion_exact(step2, globs, lib_snap_path)
        except Exception as e:
            QMessageBox.critical(self, "Exportar", f"Error al guardar programación:\n{e}")
            return

        QMessageBox.information(
            self,
            "Exportar",
            (
                "Listo. Exportación completada:\n"
                f"Excel: {target_excel}\n"
                + (f"Programación (biblioteca interna): {lib_snap_path}" if lib_snap_path else "Programación: no se pudo guardar")
            ),
        )
