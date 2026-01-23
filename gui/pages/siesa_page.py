from __future__ import annotations

import sys
import subprocess
import os
import json
import unicodedata
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFrame,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QMessageBox,
    QPlainTextEdit,
    QGroupBox,
    QFileDialog,
    QLineEdit,
    QGridLayout,
    QTableWidget,
    QTableWidgetItem,
    QAbstractItemView,
    QHeaderView,
    QSizePolicy,
    QDialog,
    QTreeWidget,
    QTreeWidgetItem,
    QDoubleSpinBox,
    QSpinBox,
    QScrollArea,
)

from logic.siesa_datahub import (
    load_hub,
    get_company_data_dir,
    set_company_data_dir,
)
from logic.cotizador import (
    load_catalog,
    load_rules,
    normalize_project,
    calculate_totals,
    format_cop,
)
from data.siesa.datahub import PackageError
from data.siesa.siesa_uno_client import SiesaConfig, SiesaUNOClient


class SiesaPage(QFrame):
    TEMPLATE_OPM = 50759
    TRAPOS_REF = "129006"
    ITEM_ID = 9044
    TEST_DB = "PRUEBAWESTON"
    TEST_USER = "scaicedo"

    def __init__(self, parent=None):
        super().__init__(parent)
        self._loaded_once = False
        self._orders_cache = []
        self._last_component_rowid = None
        self._last_component_qty = None
        self._last_component_op = ""
        self._last_component_code = ""
        self._last_created_opm = ""
        self._argos_table_updating = False
        self._argos_modules = []
        self._argos_catalog = load_catalog()
        self._argos_rules = load_rules()
        self._argos_project = normalize_project({}, self._argos_rules)
        self._argos_product = self._find_argos_product(self._argos_catalog)
        self._argos_cart = []
        self._ops_selected_item_rowid = None
        self._ops_selected_item_qty = None
        self._ops_rowid_op = None
        self._ops_id_cia = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        layout.addWidget(scroll)

        container = QFrame()
        scroll.setWidget(container)
        content = QVBoxLayout(container)
        content.setContentsMargins(16, 16, 16, 16)
        content.setSpacing(12)

        title = QLabel("SIESA - DATAHUB")
        title.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        content.addWidget(title)

        actions = QHBoxLayout()
        actions.setSpacing(10)
        self.btn_diag = QPushButton("VALIDAR CONEXION")
        actions.addWidget(self.btn_diag)
        actions.addStretch(1)
        content.addLayout(actions)

        self.lbl_folder = QLabel()
        self.lbl_folder.setStyleSheet("color:#334155;")
        self.lbl_folder.setVisible(False)

        info_box = QGroupBox("SIESA")
        info_layout = QVBoxLayout(info_box)
        info_layout.setContentsMargins(12, 10, 12, 12)
        info_layout.setSpacing(8)

        self.lbl_manifest = QLabel("Sin cargar")
        self.lbl_manifest.setStyleSheet("font-weight:700;color:#0f172a;")
        self.lbl_manifest.setVisible(False)

        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMinimumHeight(180)
        self.txt_log.setVisible(False)

        argos_card = QFrame()
        argos_card.setStyleSheet(
            "QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:10px;}"
        )
        argos_layout = QVBoxLayout(argos_card)
        argos_layout.setContentsMargins(12, 10, 12, 12)
        argos_layout.setSpacing(8)

        self.lbl_argos = QLabel("COTIZADOR SIMPLE - AUTOSERVICIO ARGOS")
        self.lbl_argos.setStyleSheet("font-weight:700;color:#0f172a;")
        argos_layout.addWidget(self.lbl_argos)

        argos_grid = QGridLayout()
        argos_grid.setHorizontalSpacing(10)
        argos_grid.setVerticalSpacing(6)

        lbl_argos_eq = QLabel("PROYECTO")
        lbl_argos_dist = QLabel("DISTANCIA (m)")
        lbl_argos_mod = QLabel("MODULACION")
        lbl_argos_qty = QLabel("MODULOS")
        for lbl in (lbl_argos_eq, lbl_argos_dist, lbl_argos_mod, lbl_argos_qty):
            lbl.setStyleSheet("color:#64748b;font-weight:600;")

        self.ed_argos_equipo = QLineEdit()
        self.ed_argos_equipo.setText(
            str(self._argos_product.get("name", "AUTOSERVICIO ARGOS")) if self._argos_product else "AUTOSERVICIO ARGOS"
        )
        self.spin_argos_distance = QDoubleSpinBox()
        self.spin_argos_distance.setRange(0, 9999)
        self.spin_argos_distance.setDecimals(2)
        self.spin_argos_distance.setSingleStep(0.1)
        self.spin_argos_distance.setSpecialValueText("--")
        self.ed_argos_mod = QLineEdit()
        self.ed_argos_mod.setReadOnly(True)
        self.ed_argos_mod.setPlaceholderText("3.75 + 2.5")
        self.ed_argos_qty = QLineEdit()
        self.ed_argos_qty.setReadOnly(True)
        self.ed_argos_qty.setPlaceholderText("0")
        self.btn_argos_validate = QPushButton("VALIDAR ARGOS")
        self.btn_argos_summary = QPushButton("RESUMEN")

        argos_grid.addWidget(lbl_argos_eq, 0, 0)
        argos_grid.addWidget(self.ed_argos_equipo, 0, 1, 1, 2)
        argos_grid.addWidget(lbl_argos_dist, 0, 3)
        argos_grid.addWidget(self.spin_argos_distance, 0, 4)
        argos_grid.addWidget(self.btn_argos_validate, 0, 5)
        argos_grid.addWidget(lbl_argos_mod, 1, 0)
        argos_grid.addWidget(self.ed_argos_mod, 1, 1, 1, 2)
        argos_grid.addWidget(lbl_argos_qty, 1, 3)
        argos_grid.addWidget(self.ed_argos_qty, 1, 4)
        argos_grid.addWidget(self.btn_argos_summary, 1, 5)
        argos_grid.setColumnStretch(1, 1)
        argos_grid.setColumnStretch(2, 1)
        argos_layout.addLayout(argos_grid)

        self.argos_table = QTableWidget(0, 4)
        self.argos_table.setHorizontalHeaderLabels(["DESCRIPCION", "PRECIO", "CANTIDAD", "SUBTOTAL"])
        self.argos_table.verticalHeader().setVisible(False)
        self.argos_table.setAlternatingRowColors(True)
        self.argos_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.argos_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.argos_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.argos_table.setStyleSheet(
            "QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QHeaderView::section{background:#f8fafc;border:1px solid #e2e8f5;padding:6px;font-weight:700;}"
        )
        argos_layout.addWidget(self.argos_table)

        gen_row = QHBoxLayout()
        gen_row.setSpacing(8)
        self.btn_argos_generate = QPushButton("GENERAR OP")
        self.ed_argos_ops_created = QLineEdit()
        self.ed_argos_ops_created.setReadOnly(True)
        self.ed_argos_ops_created.setPlaceholderText("OPS CREADAS")
        gen_row.addWidget(self.btn_argos_generate)
        gen_row.addWidget(self.ed_argos_ops_created, 1)
        argos_layout.addLayout(gen_row)

        info_layout.addWidget(argos_card)

        items_card = QFrame()
        items_card.setStyleSheet(
            "QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:10px;}"
        )
        items_layout = QVBoxLayout(items_card)
        items_layout.setContentsMargins(12, 10, 12, 12)
        items_layout.setSpacing(8)

        self.lbl_ops_items = QLabel("ITEMS OPS (PRUEBAWESTON)")
        self.lbl_ops_items.setStyleSheet("font-weight:700;color:#0f172a;")
        items_layout.addWidget(self.lbl_ops_items)

        items_row = QHBoxLayout()
        items_row.setSpacing(8)
        self.ed_ops_items_op = QLineEdit()
        self.ed_ops_items_op.setPlaceholderText("OP")
        self.btn_ops_items_load = QPushButton("CARGAR ITEMS")
        self.ed_ops_items_qty = QLineEdit()
        self.ed_ops_items_qty.setPlaceholderText("NUEVA CANTIDAD")
        self.btn_ops_items_update = QPushButton("ACTUALIZAR ITEM")
        items_row.addWidget(self.ed_ops_items_op)
        items_row.addWidget(self.btn_ops_items_load)
        items_row.addSpacing(12)
        items_row.addWidget(self.ed_ops_items_qty)
        items_row.addWidget(self.btn_ops_items_update)
        items_row.addStretch(1)
        items_layout.addLayout(items_row)

        self.ops_items_table = QTableWidget(0, 5)
        self.ops_items_table.setHorizontalHeaderLabels([
            "ITEM", "DESCRIPCION", "BODEGA", "UM", "CANTIDAD",
        ])
        self.ops_items_table.verticalHeader().setVisible(False)
        self.ops_items_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.ops_items_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.ops_items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ops_items_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.ops_items_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.ops_items_table.setStyleSheet(
            "QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QHeaderView::section{background:#f8fafc;border:1px solid #e2e8f5;padding:6px;font-weight:700;}"
        )
        items_layout.addWidget(self.ops_items_table)

        info_layout.addWidget(items_card)

        self.lbl_orders = QLabel("ORDENES (ULTIMOS 2 MESES)")
        self.lbl_orders.setStyleSheet("font-weight:700;color:#0f172a;")
        info_layout.addWidget(self.lbl_orders)

        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)
        self.ed_op_filter = QLineEdit()
        self.ed_op_filter.setPlaceholderText("BUSCAR OP")
        filter_row.addWidget(self.ed_op_filter)
        lbl_months = QLabel("MESES")
        lbl_months.setStyleSheet("color:#64748b;font-weight:600;")
        self.spin_orders_months = QSpinBox()
        self.spin_orders_months.setRange(1, 24)
        self.spin_orders_months.setValue(2)
        self.spin_orders_months.setFixedWidth(70)
        self.btn_orders_reload = QPushButton("CARGAR")
        filter_row.addWidget(lbl_months)
        filter_row.addWidget(self.spin_orders_months)
        filter_row.addWidget(self.btn_orders_reload)
        filter_row.addStretch(1)
        info_layout.addLayout(filter_row)

        self.orders_table = QTableWidget(0, 6)
        self.orders_table.setHorizontalHeaderLabels([
            "OP", "FECHA", "ESTADO", "NOTAS", "REF 1", "REF 2",
        ])
        self.orders_table.verticalHeader().setVisible(False)
        self.orders_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.orders_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        header = self.orders_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
        self.orders_table.setMinimumHeight(220)
        self.orders_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.orders_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.orders_table.setStyleSheet(
            "QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QHeaderView::section{background:#f8fafc;border:1px solid #e2e8f5;padding:6px;font-weight:700;}"
        )
        info_layout.addWidget(self.orders_table)

        content.addWidget(info_box)
        content.addStretch(1)

        self.btn_diag.clicked.connect(self._run_diagnostic)
        self.ed_op_filter.textChanged.connect(self._apply_orders_filter)
        self.btn_orders_reload.clicked.connect(self._load_orders_preview)
        self.spin_orders_months.editingFinished.connect(self._load_orders_preview)
        self.btn_argos_validate.clicked.connect(self._argos_validate)
        self.btn_argos_summary.clicked.connect(self._argos_show_summary)
        self.btn_argos_generate.clicked.connect(self._argos_generate_ops)
        self.spin_argos_distance.valueChanged.connect(self._argos_update_from_distance)
        self.argos_table.itemChanged.connect(self._argos_on_table_changed)
        self.btn_ops_items_load.clicked.connect(self._ops_items_load)
        self.btn_ops_items_update.clicked.connect(self._ops_items_update)
        self.ops_items_table.itemSelectionChanged.connect(self._ops_items_select)

        self._update_folder_label()
        if self._argos_product:
            self._argos_populate_table(self._argos_product, [])

    def showEvent(self, event):
        super().showEvent(event)
        if not self._loaded_once:
            self._orders_cache = []
            self.orders_table.setRowCount(0)
            self.lbl_orders.setText(f"ORDENES (ULTIMOS {self._orders_months()} MESES): 0")
            self._loaded_once = True

    def _update_folder_label(self) -> None:
        folder = get_company_data_dir()
        self.lbl_folder.setText(f"CARPETA: {folder}")

    def _append_log(self, text: str) -> None:
        self.txt_log.appendPlainText(text)

    @staticmethod
    def _argos_norm_text(text: str) -> str:
        raw = str(text or "").strip().upper()
        raw = unicodedata.normalize("NFKD", raw).encode("ascii", "ignore").decode("ascii")
        return raw

    def _find_argos_product(self, catalog: dict) -> dict:
        products = catalog.get("_products_by_id", {}) if isinstance(catalog, dict) else {}
        for prod in products.values():
            name = self._argos_norm_text(prod.get("name", ""))
            if "ARGOS" in name and "AUTO" in name:
                return prod
        for prod in products.values():
            name = self._argos_norm_text(prod.get("name", ""))
            if "ARGOS" in name:
                return prod
        return next(iter(products.values()), {}) if products else {}

    def _argos_validate(self) -> None:
        if not self._argos_product:
            QMessageBox.warning(self, "SIESA", "No se encontro producto AUTOSERVICIO ARGOS en el catalogo.")
            return
        name = self._argos_norm_text(self._argos_product.get("name", ""))
        if "ARGOS" not in name:
            QMessageBox.warning(self, "SIESA", "Producto no es AUTOSERVICIO ARGOS.")
            return
        self._argos_update_from_distance()
        QMessageBox.information(self, "SIESA", "Autoservicio Argos validado.")

    def _argos_update_from_distance(self) -> None:
        if not self._argos_product:
            return
        length_m = float(self.spin_argos_distance.value())
        modules = self._argos_compute_modules(length_m)
        self._argos_modules = modules
        if modules:
            mod_text = " + ".join(f"{m:g}" for m in modules)
            self.ed_argos_mod.setText(mod_text)
            self.ed_argos_qty.setText(str(len(modules)))
            self._argos_populate_table(self._argos_product, modules)
        else:
            self.ed_argos_mod.setText("")
            self.ed_argos_qty.setText("0")
            self._argos_populate_table(self._argos_product, [])

    @staticmethod
    def _argos_compute_modules(length_m: float) -> list[float]:
        sizes = [3.75, 2.5]
        if length_m <= 0:
            return []
        rem = length_m
        modules: list[float] = []
        smallest = sizes[-1]
        limit = 0
        while rem > 0.05 and limit < 50:
            pick = None
            for size in sizes:
                if size <= rem + 0.05:
                    pick = size
                    break
            if pick is None:
                pick = smallest
            modules.append(pick)
            rem -= pick
            limit += 1
            if rem < 0.05:
                break
        return modules

    @staticmethod
    def _argos_extract_length(text: str) -> float:
        raw = "".join(ch if ch.isdigit() or ch in ".," else " " for ch in str(text))
        parts = [p for p in raw.replace(",", ".").split() if p]
        for p in parts:
            try:
                return float(p)
            except Exception:
                continue
        return 0.0

    def _argos_get_base_module_price(self, product: dict) -> tuple[float, float]:
        base_modules = product.get("base_modules", []) or []
        base_price = 0.0
        base_len = 2.5
        for base in base_modules:
            desc = self._argos_norm_text(base.get("desc", ""))
            if "MODULO" in desc:
                base_price = float(base.get("unit_price", 0) or 0)
                base_len = self._argos_extract_length(desc) or base_len
                return base_price, base_len
        if base_modules:
            base_price = float(base_modules[0].get("unit_price", 0) or 0)
            base_len = self._argos_extract_length(str(base_modules[0].get("desc", ""))) or base_len
        return base_price, base_len

    def _argos_find_unit_price(self, product: dict, keywords: list[str]) -> float:
        targets = [self._argos_norm_text(k) for k in keywords]
        for base in product.get("base_modules", []) or []:
            desc = self._argos_norm_text(base.get("desc", ""))
            if any(k in desc for k in targets):
                return float(base.get("unit_price", 0) or 0)
        return 0.0

    def _argos_rows_for_modules(self, product: dict, modules: list[float]) -> list[dict]:
        rows: list[dict] = []
        if modules:
            base_price, base_len = self._argos_get_base_module_price(product)
            module_counts: dict[float, int] = {}
            for length in modules:
                key = round(float(length), 2)
                module_counts[key] = module_counts.get(key, 0) + 1
            for length, count in sorted(module_counts.items(), key=lambda x: -x[0]):
                unit_price = self._argos_scale_module_price(base_price, base_len, length)
                rows.append({
                    "desc": f"MODULO BASE {length:g} M",
                    "unit_price": unit_price,
                    "qty": count,
                    "kind": "base",
                    "editable_qty": False,
                })

            module_count = len(modules)
            testero_pairs = max(1, module_count - 1) if module_count > 0 else 0
            illumination_price = self._argos_find_unit_price(product, ["ILUMINACION"])
            control_price = self._argos_find_unit_price(product, ["CONTROL"])

            if illumination_price > 0:
                rows.append({
                    "desc": "ILUMINACION LED",
                    "unit_price": illumination_price,
                    "qty": module_count + testero_pairs,
                    "kind": "base",
                    "editable_qty": True,
                })
            if control_price > 0:
                rows.append({
                    "desc": "CONTROL BASICO",
                    "unit_price": control_price,
                    "qty": max(1, testero_pairs) if module_count else 0,
                    "kind": "base",
                    "editable_qty": True,
                })

            for addon in product.get("optional_addons", []) or []:
                desc = str(addon.get("desc", "")).strip()
                if not desc:
                    continue
                upper = self._argos_norm_text(desc)
                qty = 0
                if "TESTERO" in upper:
                    qty = testero_pairs
                rows.append({
                    "desc": desc.upper(),
                    "unit_price": addon.get("unit_price", 0),
                    "qty": qty,
                    "kind": "addon",
                    "editable_qty": True,
                })
        else:
            for base in product.get("base_modules", []) or []:
                rows.append({
                    "desc": str(base.get("desc", "")).upper(),
                    "unit_price": base.get("unit_price", 0),
                    "qty": 0,
                    "kind": "base",
                    "editable_qty": False,
                })
            for addon in product.get("optional_addons", []) or []:
                rows.append({
                    "desc": str(addon.get("desc", "")).upper(),
                    "unit_price": addon.get("unit_price", 0),
                    "qty": 0,
                    "kind": "addon",
                    "editable_qty": True,
                })
        return rows

    def _argos_notes_for_modules(self, modules: list[float]) -> str:
        if not self._argos_product:
            return ""
        rows = self._argos_rows_for_modules(self._argos_product, modules)
        lines = []
        for row in rows:
            qty = int(row.get("qty", 0) or 0)
            if qty <= 0:
                continue
            desc = str(row.get("desc", "")).upper()
            lines.append(f"{desc} x{qty}")
        return "\r\n".join(lines)

    def _argos_load_ops_templates(self) -> dict | None:
        base = Path(__file__).resolve().parents[2] / "data" / "siesa" / "company_data" / "snapshots"
        general = base / "ops_template_general.json"
        mat_375 = base / "ops_materiales_autoservicio_argos_3_75.json"
        mat_25 = base / "ops_materiales_autoservicio_argos_2_5.json"
        payload_375 = base / "ops_48407_payload.json"
        payload_25 = base / "ops_48405_payload.json"
        if not general.exists() or not mat_375.exists() or not mat_25.exists():
            return None
        if not payload_375.exists() or not payload_25.exists():
            return None
        try:
            return {
                "general": json.loads(general.read_text(encoding="utf-8")),
                "mat_375": json.loads(mat_375.read_text(encoding="utf-8")),
                "mat_25": json.loads(mat_25.read_text(encoding="utf-8")),
                "payload_375": json.loads(payload_375.read_text(encoding="utf-8")),
                "payload_25": json.loads(payload_25.read_text(encoding="utf-8")),
            }
        except Exception:
            return None

    @staticmethod
    def _argos_parse_dt(value):
        if value is None:
            return None
        if isinstance(value, datetime):
            return value
        text = str(value).strip()
        if not text:
            return None
        try:
            return datetime.fromisoformat(text)
        except Exception:
            return value

    @staticmethod
    def _argos_scale_fields(fields: dict, factor: float) -> dict:
        out = {}
        for k, v in fields.items():
            if isinstance(v, (int, float)) and "cant_" in k:
                out[k] = v * factor
            else:
                out[k] = v
        return out

    @staticmethod
    def _safe_int(value, default: int = 0) -> int:
        try:
            return int(value)
        except Exception:
            return default

    def _next_ops_consec(self, cur) -> int:
        cur.execute(
            """
SELECT MAX(f850_consec_docto)
FROM dbo.t850_mf_op_docto
WHERE f850_id_tipo_docto='OPS' AND f850_id_co='001'
  AND f850_consec_docto BETWEEN 40000 AND 90000
"""
        )
        row = cur.fetchone()
        base = row[0] if row and row[0] is not None else None
        if base is None:
            cur.execute(
                """
SELECT f022_cons_proximo
FROM dbo.t022_mm_consecutivos
WHERE f022_id_cia=1 AND f022_id_co='001' AND f022_id_tipo_docto='OPS'
"""
            )
            row = cur.fetchone()
            base = row[0] if row and row[0] else None
        if base is None:
            raise RuntimeError("No se pudo determinar el consecutivo OPS.")
        next_consec = int(base) + 1
        while True:
            cur.execute(
                """
SELECT COUNT(*)
FROM dbo.t850_mf_op_docto
WHERE f850_id_tipo_docto='OPS' AND f850_id_co='001' AND f850_consec_docto=?
""",
                next_consec,
            )
            if cur.fetchone()[0] == 0:
                break
            next_consec += 1
        return next_consec

    def _argos_get_user_rowid(self, cur, user: str) -> int:
        cur.execute(
            "SELECT TOP 1 f552_rowid FROM dbo.t552_ss_usuarios WHERE UPPER(f552_nombre)=UPPER(?)",
            user,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Usuario no encontrado: {user}")
        return int(row[0])

    def _argos_resolve_item_ext(self, cur, item_ref: str) -> int:
        cur.execute(
            """
SELECT TOP 1 e.f121_rowid
FROM dbo.t121_mc_items_extensiones e
JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid
WHERE RTRIM(i.f120_referencia) = ?
""",
            item_ref,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Item no encontrado: {item_ref}")
        return int(row[0])

    def _argos_resolve_instalacion_for_item_ext(
        self,
        cur,
        rowid_item_ext: int,
        prefer: str | None,
        id_cia: int,
    ) -> str:
        pref = str(prefer).strip() if prefer is not None else ""
        if pref:
            cur.execute(
                """
SELECT TOP 1 f135_id_instalacion
FROM dbo.t135_mc_items_instalacion_key
WHERE f135_rowid_item_ext=? AND f135_id_instalacion=?
""",
                rowid_item_ext,
                pref,
            )
            row = cur.fetchone()
            if row:
                return str(row[0]).strip()
        cur.execute(
            """
SELECT TOP 1 f135_id_instalacion
FROM dbo.t135_mc_items_instalacion_key
WHERE f135_rowid_item_ext=?
""",
            rowid_item_ext,
        )
        row = cur.fetchone()
        if row:
            return str(row[0]).strip()
        install = pref or "001"
        cur.execute(
            """
INSERT INTO dbo.t135_mc_items_instalacion_key
    (f135_id_cia, f135_rowid_item_ext, f135_id_instalacion)
VALUES (?, ?, ?)
""",
            id_cia,
            rowid_item_ext,
            install,
        )
        return install

    def _argos_resolve_bodega(self, cur, bodega_id: str) -> int:
        cur.execute(
            "SELECT TOP 1 f150_rowid FROM dbo.t150_mc_bodegas WHERE f150_id = ?",
            bodega_id,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Bodega no encontrada: {bodega_id}")
        return int(row[0])

    def _argos_resolve_ruta(self, cur, ruta_id: str) -> int:
        cur.execute(
            "SELECT TOP 1 f808_rowid FROM dbo.t808_mf_rutas WHERE f808_id = ?",
            ruta_id,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Ruta no encontrada: {ruta_id}")
        return int(row[0])

    def _argos_generate_ops(self) -> None:
        if not self._argos_product:
            QMessageBox.warning(self, "SIESA", "No se encontro producto AUTOSERVICIO ARGOS.")
            return
        proyecto = self.ed_argos_equipo.text().strip()
        if not proyecto:
            QMessageBox.warning(self, "SIESA", "Ingresa el proyecto.")
            return

        if not self._argos_modules:
            self._argos_update_from_distance()
        if not self._argos_modules:
            QMessageBox.warning(self, "SIESA", "No hay modulos para generar OP.")
            return

        templates = self._argos_load_ops_templates()
        if not templates:
            QMessageBox.warning(self, "SIESA", "No se encontraron plantillas OPS.")
            return

        counts = {}
        for m in self._argos_modules:
            key = round(float(m), 2)
            counts[key] = counts.get(key, 0) + 1
        plan = []
        if counts.get(3.75):
            plan.append((3.75, counts[3.75], templates["mat_375"], templates["payload_375"]))
        if counts.get(2.5):
            plan.append((2.5, counts[2.5], templates["mat_25"], templates["payload_25"]))
        if not plan:
            QMessageBox.warning(self, "SIESA", "No hay modulos 3.75 o 2.5 para generar OP.")
            return

        confirm = QMessageBox.question(
            self,
            "SIESA",
            f"Crear {len(plan)} OPS en {self.TEST_DB}?",
        )
        if confirm != QMessageBox.Yes:
            return

        cfg_test = self._get_siesa_test_config()
        client_test = SiesaUNOClient(cfg_test)
        if client_test._import_pyodbc() is None:
            QMessageBox.warning(self, "SIESA", "pyodbc no disponible.")
            return

        created = []
        try:
            import pyodbc
            with pyodbc.connect(client_test._conn_str(), timeout=10) as conn:
                conn.autocommit = False
                cur = conn.cursor()
                cur.execute("SELECT DB_NAME()")
                db_name = str(cur.fetchone()[0]).upper()
                if db_name != self.TEST_DB:
                    raise RuntimeError(f"DB inesperada: {db_name}")
                rowid_usuario = self._argos_get_user_rowid(cur, self.TEST_USER)

                for size, count, mat, payload in plan:
                    header = templates["general"].get("header", {})
                    items = mat.get("items", [])
                    comps = mat.get("components", [])

                    payload_items = payload.get("items", [])
                    payload_comps = payload.get("components", [])
                    item_by_rowid = {i.get("f851_rowid"): i for i in payload_items}
                    parent_keys = []
                    for comp in payload_comps:
                        parent = item_by_rowid.get(comp.get("f860_rowid_op_docto_item"))
                        if not parent:
                            parent_keys.append(None)
                            continue
                        parent_keys.append((
                            str(parent.get("item_ref", "")).strip(),
                            str(parent.get("bodega_id", "")).strip(),
                            str(parent.get("ruta_id", "")).strip(),
                        ))

                    next_consec = self._next_ops_consec(cur)
                    ref1_s = self._safe_text(proyecto, 30)
                    ref2_s = self._safe_text("JAIRO ROA Y OTONIEL", 30)
                    ref3_s = self._safe_text("PEDIDO XXXX", 30)
                    notes_s = self._safe_text(self._argos_notes_for_modules([size] * count), 2000)
                    estado_ops = 1

                    try:
                        cur.execute(
                            """
DECLARE @p_retorno smallint, @p_ts datetime, @p_rowid int;
EXEC dbo.sp_mf_op_docto_eventos
    @p_retorno=@p_retorno OUTPUT,
    @p_ts=@p_ts OUTPUT,
    @p_rowid=@p_rowid OUTPUT,
    @p_opcion=?,
    @p_id_cia=?,
    @p_id_co=?,
    @p_id_tipo_docto=?,
    @p_consec_docto=?,
    @p_id_fecha=?,
    @p_id_grupo_clase_docto=?,
    @p_id_clase_docto=?,
    @p_id_clase_op=?,
    @p_ind_tipo_op=?,
    @p_ind_multiples_items=?,
    @p_ind_consolida_comp_oper=?,
    @p_ind_requiere_lm=?,
    @p_ind_genera_ordenes_comp=?,
    @p_ind_genera_misma_orden=?,
    @p_ind_genera_todos_niveles=?,
    @p_ind_genera_solo_faltantes=?,
    @p_ind_metodo_lista_op=?,
    @p_ind_controla_tep=?,
    @p_ind_genera_consumos_tep=?,
    @p_ind_genera_entradas_tep=?,
    @p_id_clase_op_generar=?,
    @p_ind_confirmar_al_aprobar=?,
    @p_ind_distribucion_costos=?,
    @p_ind_devolucion_comp=?,
    @p_ind_estado=?,
    @p_fecha_cumplida=?,
    @p_rowid_tercero_planif=?,
    @p_id_instalacion=?,
    @p_rowid_op_padre=?,
    @p_referencia_1=?,
    @p_referencia_2=?,
    @p_referencia_3=?,
    @p_ind_transmitido=?,
    @p_ind_impresion=?,
    @p_nro_impresiones=?,
    @p_usuario=?,
    @p_rowid_usuario=?,
    @p_notas=?,
    @p_ind_posdeduccion=?,
    @p_ind_pedido_venta=?,
    @p_rowid_pv_docto=?,
    @p_ind_posdeduccion_tep=?,
    @p_ind_reg_incons_posd=?,
    @p_ind_lote_automatico=?,
    @p_ind_controlar_cant_ext=?,
    @p_ind_incluir_operacion=?,
    @p_ind_entrega_estandar=?,
    @p_ind_valida_consumo_tot=?,
    @p_ind_liq_tep_estandar=?,
    @p_ind_no_liq_tep=?;
SELECT @p_retorno AS retorno, @p_ts AS ts, @p_rowid AS rowid;
""",
                            [
                                0,
                                self._safe_int(header.get("f850_id_cia", 1), 1),
                                header.get("f850_id_co", "001"),
                                header.get("f850_id_tipo_docto", "OPS"),
                                next_consec,
                                self._argos_parse_dt(header.get("f850_id_fecha")),
                                self._safe_int(header.get("f850_id_grupo_clase_docto", 0)),
                                self._safe_int(header.get("f850_id_clase_docto", 0)),
                                header.get("f850_id_clase_op"),
                                self._safe_int(header.get("f850_ind_tipo_op", 0)),
                                self._safe_int(header.get("f850_ind_multiples_items", 0)),
                                self._safe_int(header.get("f850_ind_consolida_comp_oper", 0)),
                                self._safe_int(header.get("f850_ind_requiere_lm", 0)),
                                self._safe_int(header.get("f850_ind_genera_ordenes_comp", 0)),
                                self._safe_int(header.get("f850_ind_genera_misma_orden", 0)),
                                self._safe_int(header.get("f850_ind_genera_todos_niveles", 0)),
                                self._safe_int(header.get("f850_ind_genera_solo_faltantes", 0)),
                                self._safe_int(header.get("f850_ind_metodo_lista_op", 0)),
                                self._safe_int(header.get("f850_ind_controla_tep", 0)),
                                self._safe_int(header.get("f850_ind_genera_consumos_tep", 0)),
                                self._safe_int(header.get("f850_ind_genera_entradas_tep", 0)),
                                header.get("f850_id_clase_op_generar"),
                                self._safe_int(header.get("f850_ind_confirmar_al_aprobar", 0)),
                                self._safe_int(header.get("f850_ind_distribucion_costos", 0)),
                                self._safe_int(header.get("f850_ind_devolucion_comp", 0)),
                                estado_ops,
                                self._argos_parse_dt(header.get("f850_fecha_cumplida")),
                                self._safe_int(header.get("f850_rowid_tercero_planif", 0)),
                                header.get("f850_id_instalacion", "001"),
                                header.get("f850_rowid_op_padre"),
                                ref1_s,
                                ref2_s,
                                ref3_s,
                                self._safe_int(header.get("f850_ind_transmitido", 0)),
                                self._safe_int(header.get("f850_ind_impresion", 0)),
                                self._safe_int(header.get("f850_nro_impresiones", 0)),
                                self.TEST_USER,
                                rowid_usuario,
                                notes_s,
                                self._safe_int(header.get("f850_ind_posdeduccion", 0)),
                                self._safe_int(header.get("f850_ind_pedido_venta", 0)),
                                header.get("f850_rowid_pv_docto"),
                                self._safe_int(header.get("f850_ind_posdeduccion_tep", 0)),
                                self._safe_int(header.get("f850_ind_reg_incons_posd", 0)),
                                self._safe_int(header.get("f850_ind_lote_automatico", 0)),
                                self._safe_int(header.get("f850_ind_controlar_cant_ext", 0)),
                                self._safe_int(header.get("f850_ind_incluir_operacion", 0)),
                                self._safe_int(header.get("f850_ind_entrega_estandar", 0)),
                                self._safe_int(header.get("f850_ind_valida_consumo_tot", 0)),
                                self._safe_int(header.get("f850_ind_liq_tep_estandar", 0)),
                                self._safe_int(header.get("f850_ind_no_liq_tep", 0)),
                            ],
                        )
                        out = cur.fetchone()
                        if not out:
                            raise RuntimeError("Sin respuesta al crear encabezado OPS.")
                        retorno, _, rowid_op = out
                        if retorno not in (0, None):
                            raise RuntimeError(f"Error creando OPS: {retorno}")

                        item_rowids = []
                        item_key_map = {}
                        for item in items:
                            item_ref_val = item.get("item_ref")
                            item_ref = "" if item_ref_val is None else str(item_ref_val).strip()
                            if not item_ref or item_ref.upper() == "NONE":
                                raise RuntimeError("Item no encontrado: None")
                            bodega_val = item.get("bodega_id")
                            bodega_id = "" if bodega_val is None else str(bodega_val).strip()
                            ruta_val = item.get("ruta_id")
                            ruta_id = "" if ruta_val is None else str(ruta_val).strip()
                            fields = item.get("fields", {})
                            scaled = self._argos_scale_fields(fields, float(count))
                            scaled["f851_ind_estado"] = estado_ops

                            rowid_item_ext = self._argos_resolve_item_ext(cur, item_ref)
                            id_instalacion = self._argos_resolve_instalacion_for_item_ext(
                                cur,
                                rowid_item_ext,
                                scaled.get("f851_id_instalacion"),
                                self._safe_int(header.get("f850_id_cia", 1), 1),
                            )
                            rowid_bodega = self._argos_resolve_bodega(cur, bodega_id) if bodega_id else None
                            rowid_ruta = self._argos_resolve_ruta(cur, ruta_id) if ruta_id else None

                            cur.execute(
                                """
DECLARE @p_ts datetime, @p_rowid int;
EXEC dbo.sp_mf_op_movto_eventos
    @p_ts=@p_ts OUTPUT,
    @p_rowid=@p_rowid OUTPUT,
    @p_opcion=?,
    @p_id_cia=?,
    @p_rowid_op_docto=?,
    @p_rowid_item_ext_padre=?,
    @p_rowid_bodega=?,
    @p_id_instalacion=?,
    @p_id_fecha=?,
    @p_ind_estado=?,
    @p_fecha_cumplida=?,
    @p_ind_automatico=?,
    @p_id_metodo_lista_mater=?,
    @p_rowid_ruta=?,
    @p_id_metodo_ruta=?,
    @p_fecha_terminacion=?,
    @p_fecha_inicio=?,
    @p_id_tipo_inv_serv=?,
    @p_ind_tipo_item=?,
    @p_id_unidad_medida=?,
    @p_factor=?,
    @p_cant_planeada_base=?,
    @p_cant_ordenada_base=?,
    @p_cant_completa_base=?,
    @p_cant_desechos_base=?,
    @p_cant_rechazos_base=?,
    @p_cant_planeada_1=?,
    @p_cant_ordenada_1=?,
    @p_cant_completa_1=?,
    @p_cant_desechos_1=?,
    @p_cant_rechazos_1=?,
    @p_cant_parcial_base=?,
    @p_ind_controla_secuencia=?,
    @p_porc_rendimiento=?,
    @p_rowid_bodega_componentes=?,
    @p_notas=?,
    @p_diferente=?,
    @p_ind_condicionar=?,
    @p_estado_condicionar=?,
    @p_id_lote=?,
    @p_ind_afectar_items=?,
    @p_seg_selec_item_por_nivel=?,
    @p_rowid_pv_movto=?;
SELECT @p_rowid AS rowid;
""",
                                [
                                    0,
                                    self._safe_int(header.get("f850_id_cia", 1), 1),
                                    rowid_op,
                                    rowid_item_ext,
                                    rowid_bodega,
                                    id_instalacion,
                                    self._argos_parse_dt(scaled.get("f851_id_fecha")),
                                    self._safe_int(scaled.get("f851_ind_estado", 0)),
                                    self._argos_parse_dt(scaled.get("f851_fecha_cumplida")),
                                    self._safe_int(scaled.get("f851_ind_automatico", 0)),
                                    scaled.get("f851_id_metodo_lista_mater"),
                                    rowid_ruta,
                                    scaled.get("f851_id_metodo_ruta"),
                                    self._argos_parse_dt(scaled.get("f851_fecha_terminacion")),
                                    self._argos_parse_dt(scaled.get("f851_fecha_inicio")),
                                    scaled.get("f851_id_tipo_inv_serv"),
                                    self._safe_int(scaled.get("f851_ind_tipo_item", 0)),
                                    scaled.get("f851_id_unidad_medida"),
                                    float(scaled.get("f851_factor", 0) or 0),
                                    float(scaled.get("f851_cant_planeada_base", 0) or 0),
                                    float(scaled.get("f851_cant_ordenada_base", 0) or 0),
                                    float(scaled.get("f851_cant_completa_base", 0) or 0),
                                    float(scaled.get("f851_cant_desechos_base", 0) or 0),
                                    float(scaled.get("f851_cant_rechazos_base", 0) or 0),
                                    float(scaled.get("f851_cant_planeada_1", 0) or 0),
                                    float(scaled.get("f851_cant_ordenada_1", 0) or 0),
                                    float(scaled.get("f851_cant_completa_1", 0) or 0),
                                    float(scaled.get("f851_cant_desechos_1", 0) or 0),
                                    float(scaled.get("f851_cant_rechazos_1", 0) or 0),
                                    float(scaled.get("f851_cant_parcial_base", 0) or 0),
                                    self._safe_int(scaled.get("f851_ind_controla_secuencia", 0)),
                                    float(scaled.get("f851_porc_rendimiento", 0) or 0),
                                    None,
                                    scaled.get("f851_notas"),
                                    0,
                                    0,
                                    0,
                                    scaled.get("f851_id_lote"),
                                    0,
                                    0,
                                    None,
                                ],
                            )
                            out = cur.fetchone()
                            if not out:
                                raise RuntimeError("Sin respuesta al crear item OPS.")
                            rowid_item = int(out[0])
                            item_rowids.append(rowid_item)
                            item_key_map[(item_ref, bodega_id, ruta_id)] = (rowid_item, rowid_item_ext)

                        for idx, comp in enumerate(comps):
                            comp_ref_val = comp.get("item_ref")
                            comp_ref = "" if comp_ref_val is None else str(comp_ref_val).strip()
                            if not comp_ref or comp_ref.upper() == "NONE":
                                raise RuntimeError("Item no encontrado: None")
                            bodega_val = comp.get("bodega_id")
                            bodega_id = "" if bodega_val is None else str(bodega_val).strip()
                            fields = comp.get("fields", {})
                            scaled = self._argos_scale_fields(fields, float(count))

                            rowid_item_ext_comp = self._argos_resolve_item_ext(cur, comp_ref)
                            comp_inst = self._argos_resolve_instalacion_for_item_ext(
                                cur,
                                rowid_item_ext_comp,
                                scaled.get("f860_id_instalacion"),
                                self._safe_int(header.get("f850_id_cia", 1), 1),
                            )
                            rowid_bodega_comp = self._argos_resolve_bodega(cur, bodega_id) if bodega_id else None
                            sub_val = comp.get("item_ref_sust")
                            sub_ref = "" if sub_val is None else str(sub_val).strip()
                            if sub_ref.upper() == "NONE":
                                sub_ref = ""
                            rowid_item_ext_sub = self._argos_resolve_item_ext(cur, sub_ref) if sub_ref else None

                            parent_key = parent_keys[idx] if idx < len(parent_keys) else None
                            rowid_op_item = None
                            rowid_item_ext_padre = None
                            if parent_key:
                                mapped = item_key_map.get(parent_key)
                                if mapped:
                                    rowid_op_item, rowid_item_ext_padre = mapped
                            if not rowid_op_item:
                                rowid_op_item = item_rowids[0]
                                rowid_item_ext_padre = item_key_map.get(next(iter(item_key_map)), (None, None))[1]
                            if not rowid_item_ext_padre:
                                rowid_item_ext_padre = item_key_map.get(next(iter(item_key_map)), (None, None))[1]

                            cur.execute(
                                """
DECLARE @p_ts datetime, @p_rowid int;
EXEC dbo.sp_mf_op_comp_eventos
    @p_ts=@p_ts OUTPUT,
    @p_rowid=@p_rowid OUTPUT,
    @p_opcion=?,
    @p_id_cia=?,
    @p_rowid_op_docto_item=?,
    @p_rowid_op_docto=?,
    @p_rowid_item_ext_padre=?,
    @p_rowid_item_ext_componente=?,
    @p_numero_operacion=?,
    @p_rowid_ctrabajo=?,
    @p_rowid_bodega=?,
    @p_id_instalacion=?,
    @p_id_unidad_medida=?,
    @p_ind_manual=?,
    @p_factor=?,
    @p_cant_requerida_base=?,
    @p_cant_comprometida_base=?,
    @p_cant_consumida_base=?,
    @p_cant_requerida_1=?,
    @p_cant_comprometida_1=?,
    @p_cant_consumida_1=?,
    @p_cant_requerida_2=?,
    @p_cant_comprometida_2=?,
    @p_cant_consumida_2=?,
    @p_cant_desperdicio_base=?,
    @p_fecha_requerida=?,
    @p_notas=?,
    @p_usuario=?,
    @p_rowid_item_ext_sustituido=?,
    @p_codigo_sustitucion=?,
    @p_cant_equiv_sustitucion=?,
    @p_permiso_costos=?,
    @p_rowid_movto_entidad=?,
    @p_ind_cambio_cantidad=?;
SELECT @p_rowid AS rowid;
""",
                                [
                                    0,
                                    self._safe_int(header.get("f850_id_cia", 1), 1),
                                    rowid_op_item,
                                    rowid_op,
                                    rowid_item_ext_padre,
                                    rowid_item_ext_comp,
                                    self._safe_int(scaled.get("f860_numero_operacion", 0)),
                                    scaled.get("f860_rowid_ctrabajo"),
                                    rowid_bodega_comp,
                                    comp_inst,
                                    scaled.get("f860_id_unidad_medida"),
                                    self._safe_int(scaled.get("f860_ind_manual", 0)),
                                    float(scaled.get("f860_factor", 0) or 0),
                                    float(scaled.get("f860_cant_requerida_base", 0) or 0),
                                    0.0,
                                    0.0,
                                    float(scaled.get("f860_cant_requerida_1", 0) or 0),
                                    0.0,
                                    0.0,
                                    float(scaled.get("f860_cant_requerida_2", 0) or 0),
                                    0.0,
                                    0.0,
                                    float(scaled.get("f860_cant_desperdicio_base", 0) or 0),
                                    self._argos_parse_dt(scaled.get("f860_fecha_requerida")),
                                    scaled.get("f860_notas"),
                                    self.TEST_USER,
                                    rowid_item_ext_sub,
                                    self._safe_int(scaled.get("f860_codigo_sustitucion", 0)),
                                    float(scaled.get("f860_cant_equiv_sustitucion", 0) or 0),
                                    0,
                                    scaled.get("f860_rowid_movto_entidad"),
                                    self._safe_int(scaled.get("f860_ind_cambio_cantidad", 0)),
                                ],
                            )
                            out = cur.fetchone()
                            if not out:
                                raise RuntimeError("Sin respuesta al crear componente OPS.")

                        conn.commit()
                        created.append(next_consec)
                    except Exception:
                        conn.rollback()
                        raise

            if created:
                self.ed_argos_ops_created.setText(", ".join(str(c) for c in created))
                QMessageBox.information(self, "SIESA", f"OPS creadas: {', '.join(str(c) for c in created)}")
        except Exception as e:
            QMessageBox.warning(self, "SIESA", f"Error creando OPS: {e}")

    @staticmethod
    def _argos_scale_module_price(base_price: float, base_len: float, length: float) -> float:
        if base_price <= 0 or base_len <= 0:
            return 0.0
        return round(base_price * (length / base_len) / 1000.0) * 1000.0

    @staticmethod
    def _argos_parse_number(text: str) -> float:
        raw = str(text or "").replace("COP", "").replace("$", "").replace(" ", "")
        raw = raw.replace(".", "")
        raw = raw.replace(",", ".")
        cleaned = "".join(ch for ch in raw if ch.isdigit() or ch in ".-")
        try:
            return float(cleaned) if cleaned else 0.0
        except Exception:
            return 0.0

    @staticmethod
    def _argos_format_money(value: float) -> str:
        return f"{format_cop(float(value or 0))} $"

    def _argos_populate_table(self, product: dict, modules: list[float]) -> None:
        self._argos_table_updating = True
        self.argos_table.setRowCount(0)
        rows = self._argos_rows_for_modules(product, modules)
        self._argos_render_rows(rows)
        self._argos_table_updating = False

    def _argos_render_rows(self, rows: list[dict]) -> None:
        self.argos_table.setRowCount(len(rows) + 1)
        for r, row in enumerate(rows):
            desc = str(row.get("desc", "")).upper()
            desc_item = QTableWidgetItem(desc)
            desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)
            meta = {"kind": row.get("kind", "base")}
            desc_item.setData(Qt.UserRole, meta)

            unit_price = float(row.get("unit_price", 0) or 0)
            price_item = QTableWidgetItem(self._argos_format_money(unit_price))
            price_item.setData(Qt.UserRole, unit_price)
            price_item.setFlags(price_item.flags() & ~Qt.ItemIsEditable)

            qty_item = QTableWidgetItem(str(int(row.get("qty", 0) or 0)))
            if not row.get("editable_qty", False):
                qty_item.setFlags(qty_item.flags() & ~Qt.ItemIsEditable)

            sub_item = QTableWidgetItem("0")
            sub_item.setFlags(sub_item.flags() & ~Qt.ItemIsEditable)

            for item in (price_item, qty_item, sub_item):
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            self.argos_table.setItem(r, 0, desc_item)
            self.argos_table.setItem(r, 1, price_item)
            self.argos_table.setItem(r, 2, qty_item)
            self.argos_table.setItem(r, 3, sub_item)
            self._argos_update_row_subtotal(r)

        total_row = len(rows)
        total_label = QTableWidgetItem("TOTAL")
        total_label.setFlags(total_label.flags() & ~Qt.ItemIsEditable)
        total_label.setData(Qt.UserRole, "total")
        total_label.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        total_value = QTableWidgetItem(self._argos_format_money(0))
        total_value.setFlags(total_value.flags() & ~Qt.ItemIsEditable)
        total_value.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        bold_font = QFont()
        bold_font.setBold(True)
        total_label.setFont(bold_font)
        total_value.setFont(bold_font)
        blank_price = QTableWidgetItem("")
        blank_qty = QTableWidgetItem("")
        blank_price.setFlags(blank_price.flags() & ~Qt.ItemIsEditable)
        blank_qty.setFlags(blank_qty.flags() & ~Qt.ItemIsEditable)
        self.argos_table.setItem(total_row, 0, total_label)
        self.argos_table.setItem(total_row, 1, blank_price)
        self.argos_table.setItem(total_row, 2, blank_qty)
        self.argos_table.setItem(total_row, 3, total_value)
        self._argos_recalc_total()
        self._argos_auto_height()

    def _argos_auto_height(self) -> None:
        rows = max(self.argos_table.rowCount(), 1)
        vh = self.argos_table.verticalHeader()
        row_h = vh.defaultSectionSize()
        vlen = vh.length()
        if vlen <= 0:
            vlen = rows * row_h
        header_h = self.argos_table.horizontalHeader().height() if self.argos_table.horizontalHeader() else 0
        frame = self.argos_table.frameWidth() * 2
        total_h = header_h + vlen + frame + 8
        self.argos_table.setMinimumHeight(total_h)
        self.argos_table.setMaximumHeight(total_h)

    def _argos_is_total_row(self, row: int) -> bool:
        desc_item = self.argos_table.item(row, 0)
        if not desc_item:
            return False
        return desc_item.data(Qt.UserRole) == "total"

    def _argos_update_row_subtotal(self, row: int) -> None:
        if self._argos_is_total_row(row):
            return
        price_item = self.argos_table.item(row, 1)
        qty_item = self.argos_table.item(row, 2)
        sub_item = self.argos_table.item(row, 3)
        if not price_item or not qty_item or not sub_item:
            return
        unit_price = price_item.data(Qt.UserRole)
        if unit_price is None:
            unit_price = self._argos_parse_number(price_item.text())
        qty = self._argos_parse_number(qty_item.text())
        subtotal = float(unit_price) * float(qty)
        sub_item.setText(self._argos_format_money(subtotal))
        sub_item.setData(Qt.UserRole, subtotal)

    def _argos_recalc_total(self) -> None:
        total = 0.0
        for r in range(self.argos_table.rowCount()):
            if self._argos_is_total_row(r):
                continue
            sub_item = self.argos_table.item(r, 3)
            if not sub_item:
                continue
            subtotal = sub_item.data(Qt.UserRole)
            if subtotal is None:
                subtotal = self._argos_parse_number(sub_item.text())
            total += float(subtotal or 0)
        total_row = self.argos_table.rowCount() - 1
        total_item = self.argos_table.item(total_row, 3)
        if total_item:
            total_item.setText(self._argos_format_money(total))
            total_item.setData(Qt.UserRole, total)

    def _argos_on_table_changed(self, item: QTableWidgetItem) -> None:
        if self._argos_table_updating:
            return
        if not item:
            return
        row = item.row()
        if self._argos_is_total_row(row):
            return
        if item.column() == 2:
            self._argos_update_row_subtotal(row)
            self._argos_recalc_total()

    def _argos_build_cart(self) -> list[dict]:
        if not self._argos_product:
            return []
        line_items: list[dict] = []
        for row in range(self.argos_table.rowCount()):
            if self._argos_is_total_row(row):
                continue
            desc_item = self.argos_table.item(row, 0)
            price_item = self.argos_table.item(row, 1)
            qty_item = self.argos_table.item(row, 2)
            if not desc_item or not price_item or not qty_item:
                continue
            desc = str(desc_item.text() or "")
            meta = desc_item.data(Qt.UserRole) or {}
            kind = meta.get("kind", "base")
            unit_price = price_item.data(Qt.UserRole)
            if unit_price is None:
                unit_price = self._argos_parse_number(price_item.text())
            qty = self._argos_parse_number(qty_item.text())
            if qty <= 0:
                continue
            line_items.append({
                "desc": desc,
                "unit_price": float(unit_price or 0),
                "qty": int(qty),
                "kind": kind,
            })
        qty_equipo = max(1, len(self._argos_modules)) if self._argos_modules else 1
        proj_name = self.ed_argos_equipo.text().strip()
        return [{
            "product_id": self._argos_product.get("product_id", "ARGOS"),
            "product_name": proj_name or self._argos_product.get("name", "AUTOSERVICIO ARGOS"),
            "qty": qty_equipo,
            "lead_time_days": int(self._argos_product.get("lead_time_days", 0) or 0),
            "line_items": line_items,
        }]

    def _argos_build_summary_dialog(self) -> None:
        self.argos_summary_dialog = QDialog(self)
        self.argos_summary_dialog.setWindowTitle("Resumen Autoservicio Argos")
        self.argos_summary_dialog.setMinimumSize(900, 600)
        self.argos_summary_dialog.resize(980, 680)
        self.argos_summary_dialog.setModal(False)
        dlg_layout = QVBoxLayout(self.argos_summary_dialog)
        dlg_layout.setContentsMargins(16, 16, 16, 16)
        dlg_layout.setSpacing(10)

        title = QLabel("RESUMEN DEL COTIZADOR")
        title.setStyleSheet("font-size:16px;font-weight:800;color:#0f172a;")
        dlg_layout.addWidget(title)

        self.argos_summary_tree = QTreeWidget()
        self.argos_summary_tree.setColumnCount(4)
        self.argos_summary_tree.setHeaderLabels(["ITEM", "CANTIDAD", "PRECIO", "SUBTOTAL"])
        self.argos_summary_tree.setRootIsDecorated(True)
        self.argos_summary_tree.setAlternatingRowColors(True)
        self.argos_summary_tree.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.argos_summary_tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        for col in (1, 2, 3):
            self.argos_summary_tree.header().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        self.argos_summary_tree.setStyleSheet(
            "QTreeWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QHeaderView::section{background:#f8fafc;border:1px solid #e2e8f5;padding:6px;font-weight:700;}"
        )
        dlg_layout.addWidget(self.argos_summary_tree, 1)

        summary = QGridLayout()
        summary.setHorizontalSpacing(8)
        summary.setVerticalSpacing(6)
        self.argos_summary_subtotal = QLabel("0")
        self.argos_summary_recargo = QLabel("0")
        self.argos_summary_descuento = QLabel("0")
        self.argos_summary_iva = QLabel("0")
        self.argos_summary_total = QLabel("0")
        labels = [
            ("Subtotal", self.argos_summary_subtotal),
            ("Recargo estiba", self.argos_summary_recargo),
            ("Descuento", self.argos_summary_descuento),
            ("IVA", self.argos_summary_iva),
            ("Total", self.argos_summary_total),
        ]
        for r, (name, lab) in enumerate(labels):
            lab.setAlignment(Qt.AlignRight)
            if name == "Total":
                lab.setStyleSheet("font-weight:800;color:#0f172a;")
            summary.addWidget(QLabel(name), r, 0)
            summary.addWidget(lab, r, 1)
        dlg_layout.addLayout(summary)

        close_row = QHBoxLayout()
        close_row.addStretch(1)
        btn_close = QPushButton("Cerrar")
        btn_close.clicked.connect(self.argos_summary_dialog.close)
        close_row.addWidget(btn_close)
        dlg_layout.addLayout(close_row)

    def _argos_refresh_summary(self) -> None:
        self._argos_cart = self._argos_build_cart()
        self.argos_summary_tree.clear()
        for item in self._argos_cart:
            name = str(item.get("product_name") or item.get("product_id") or "ITEM").upper()
            qty_equipo = int(item.get("qty", 1) or 1)
            product_total = 0.0
            for line in item.get("line_items", []) or []:
                product_total += float(line.get("unit_price", 0)) * float(line.get("qty", 0))
            parent = QTreeWidgetItem([
                f"{name} (X{qty_equipo})",
                str(qty_equipo),
                "",
                self._argos_format_money(product_total),
            ])
            bold_font = QFont()
            bold_font.setBold(True)
            for col in range(4):
                parent.setFont(col, bold_font)
            parent.setTextAlignment(1, Qt.AlignRight | Qt.AlignVCenter)
            parent.setTextAlignment(3, Qt.AlignRight | Qt.AlignVCenter)
            self.argos_summary_tree.addTopLevelItem(parent)
            for line in item.get("line_items", []) or []:
                desc = str(line.get("desc", "")).upper()
                qty = float(line.get("qty", 0))
                unit = float(line.get("unit_price", 0))
                subtotal = unit * qty
                child = QTreeWidgetItem([
                    desc,
                    str(int(qty)),
                    self._argos_format_money(unit),
                    self._argos_format_money(subtotal),
                ])
                child.setTextAlignment(1, Qt.AlignRight | Qt.AlignVCenter)
                child.setTextAlignment(2, Qt.AlignRight | Qt.AlignVCenter)
                child.setTextAlignment(3, Qt.AlignRight | Qt.AlignVCenter)
                parent.addChild(child)
        self.argos_summary_tree.expandAll()
        totals = calculate_totals(self._argos_cart, self._argos_project, self._argos_rules)
        self.argos_summary_subtotal.setText(format_cop(totals.get("subtotal_base", 0.0)))
        self.argos_summary_recargo.setText(format_cop(totals.get("recargo_estiba", 0.0)))
        self.argos_summary_descuento.setText(format_cop(totals.get("descuento", 0.0)))
        self.argos_summary_iva.setText(format_cop(totals.get("iva", 0.0)))
        self.argos_summary_total.setText(format_cop(totals.get("total_final_rounded", 0.0)))

    def _argos_show_summary(self) -> None:
        if not hasattr(self, "argos_summary_dialog"):
            self._argos_build_summary_dialog()
        self._argos_refresh_summary()
        self.argos_summary_dialog.show()
        self.argos_summary_dialog.raise_()
        self.argos_summary_dialog.activateWindow()

    def _validate_component_qty(self) -> None:
        op = self.ed_add_op.text().strip()
        code = self.ed_add_code.text().strip()
        if not op or not code:
            QMessageBox.warning(self, "SIESA", "Completa OP y CODIGO.")
            return
        self.ed_current_qty.clear()
        self.ed_current_name.clear()
        self._last_component_rowid = None
        self._last_component_qty = None
        self._last_component_op = ""
        self._last_component_code = ""

        cfg = self._get_siesa_config()
        client = SiesaUNOClient(cfg)
        if client._import_pyodbc() is None:
            QMessageBox.warning(self, "SIESA", "pyodbc no disponible.")
            return
        try:
            import pyodbc
            with pyodbc.connect(client._conn_str(), timeout=5) as conn:
                cur = conn.cursor()
                cur.execute(
                    """
SELECT TOP 1 b850_rowid_op_docto
FROM dbo.vBI_850
WHERE b850_op_numero = ?
ORDER BY b850_rowid_op_docto DESC
""",
                    op,
                )
                row = cur.fetchone()
                if not row:
                    QMessageBox.warning(self, "SIESA", "OP no encontrada.")
                    return
                rowid_op = row[0]
                cur.execute(
                    """
SELECT TOP 1
  c.f860_rowid,
  c.f860_cant_requerida_base,
  c.f860_id_unidad_medida,
  i.f120_descripcion
FROM dbo.t860_mf_op_componentes c
JOIN dbo.t121_mc_items_extensiones ie
  ON c.f860_rowid_item_ext_componente = ie.f121_rowid
JOIN dbo.t120_mc_items i
  ON ie.f121_rowid_item = i.f120_rowid
WHERE c.f860_rowid_op_docto = ?
  AND RTRIM(i.f120_referencia) = ?
ORDER BY c.f860_rowid DESC
""",
                    rowid_op,
                    code,
                )
                comp = cur.fetchone()
                if not comp:
                    QMessageBox.warning(self, "SIESA", "Codigo no encontrado en la OP.")
                    return
                rowid, qty, um, name = comp
                qty_val = float(qty)
                if qty_val.is_integer():
                    qty_val = int(qty_val)
                self.ed_current_qty.setText(str(qty_val))
                self.ed_current_name.setText("" if name is None else str(name).strip())
                self._last_component_rowid = int(rowid)
                self._last_component_qty = qty_val
                self._last_component_op = op
                self._last_component_code = code
                self._append_log(
                    f"VALIDADO: OP={op} CODIGO={code} CANT={qty_val} UM={str(um).strip()} "
                    f"NOMBRE={'' if name is None else str(name).strip()}"
                )
        except Exception as e:
            QMessageBox.warning(self, "SIESA", f"Error validando: {e}")

    def _send_component_qty(self) -> None:
        if not self._last_component_rowid:
            QMessageBox.warning(self, "SIESA", "Primero valida OP y CODIGO.")
            return
        qty_text = self.ed_add_qty.text().strip()
        if not qty_text:
            QMessageBox.warning(self, "SIESA", "Ingresa la nueva cantidad.")
            return
        try:
            qty_val = float(qty_text)
            if qty_val.is_integer():
                qty_val = int(qty_val)
        except Exception:
            QMessageBox.warning(self, "SIESA", "Cantidad invalida.")
            return
        cfg = self._get_siesa_config()
        client = SiesaUNOClient(cfg)
        if client._import_pyodbc() is None:
            QMessageBox.warning(self, "SIESA", "pyodbc no disponible.")
            return
        try:
            import pyodbc
            with pyodbc.connect(client._conn_str(), timeout=5) as conn:
                cur = conn.cursor()
                cur.execute(
                    """
UPDATE dbo.t860_mf_op_componentes
SET f860_cant_requerida_base = ?,
    f860_cant_requerida_1 = ?,
    f860_ind_cambio_cantidad = 1
WHERE f860_rowid = ?
  AND f860_cant_requerida_base = ?
""",
                    qty_val,
                    qty_val,
                    self._last_component_rowid,
                    self._last_component_qty,
                )
                if cur.rowcount <= 0:
                    QMessageBox.warning(self, "SIESA", "No se actualizo (cambio concurrente). Valida de nuevo.")
                    return
                conn.commit()
            payload = {
                "op": self._last_component_op,
                "codigo": self._last_component_code,
                "cantidad_anterior": self._last_component_qty,
                "cantidad_nueva": qty_val,
                "rowid": self._last_component_rowid,
                "timestamp": datetime.now().isoformat(),
            }
            out_path = Path(__file__).resolve().parents[2] / "data" / "siesa" / "company_data" / "local_edits.jsonl"
            out_path.parent.mkdir(parents=True, exist_ok=True)
            with out_path.open("a", encoding="utf-8") as f:
                f.write(json.dumps(payload, ensure_ascii=False) + "\n")
            self._append_log(
                f"ENVIADO: OP={self._last_component_op} CODIGO={self._last_component_code} "
                f"CANT={qty_val} -> {out_path}"
            )
            self.ed_current_qty.setText(str(qty_val))
            self._last_component_qty = qty_val
            self.ed_add_qty.clear()
        except Exception as e:
            QMessageBox.warning(self, "SIESA", f"Error enviando: {e}")

    def _fetch_one_dict(self, cur, sql: str, params=None):
        cur.execute(sql, params or [])
        row = cur.fetchone()
        if not row:
            return None
        cols = [d[0] for d in cur.description]
        return dict(zip(cols, row))

    def _fetch_all_dicts(self, cur, sql: str, params=None) -> list[dict]:
        cur.execute(sql, params or [])
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]

    def _safe_text(self, value: str, limit: int) -> str:
        text = "" if value is None else str(value).strip()
        return text[:limit]

    def _next_opm_consec(self, cur) -> int:
        cur.execute(
            """
SELECT MAX(f850_consec_docto)
FROM dbo.t850_mf_op_docto
WHERE f850_id_tipo_docto='OPM' AND f850_id_co='001'
  AND f850_consec_docto BETWEEN 49000 AND 60000
"""
        )
        row = cur.fetchone()
        base = row[0] if row and row[0] is not None else None
        if base is None:
            cur.execute(
                """
SELECT f022_cons_proximo
FROM dbo.t022_mm_consecutivos
WHERE f022_id_cia=1 AND f022_id_co='001' AND f022_id_tipo_docto='OPM'
"""
            )
            row = cur.fetchone()
            base = row[0] if row and row[0] else None
        if base is None:
            raise RuntimeError("No se pudo determinar el consecutivo.")
        next_consec = int(base) + 1
        while True:
            cur.execute(
                """
SELECT COUNT(*)
FROM dbo.t850_mf_op_docto
WHERE f850_id_tipo_docto='OPM' AND f850_id_co='001' AND f850_consec_docto=?
""",
                next_consec,
            )
            if cur.fetchone()[0] == 0:
                break
            next_consec += 1
        return next_consec

    def _resolve_rowids_test(self, cur, user: str, item_bodega_id: int, comp_bodega_id: int, ruta_id) -> dict:
        ruta_id_text = "" if ruta_id is None else str(ruta_id).strip()
        cur.execute(
            "SELECT TOP 1 f552_rowid FROM dbo.t552_ss_usuarios WHERE UPPER(f552_nombre)=UPPER(?)",
            user,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Usuario no encontrado: {user}")
        rowid_user = int(row[0])

        cur.execute(
            """
SELECT TOP 1 e.f121_rowid
FROM dbo.t121_mc_items_extensiones e
JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid
WHERE i.f120_id = ?
""",
            self.ITEM_ID,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Item base no encontrado: {self.ITEM_ID}")
        rowid_item_ext_padre = int(row[0])

        cur.execute(
            "SELECT TOP 1 f120_id FROM dbo.t120_mc_items WHERE f120_referencia = ?",
            self.TRAPOS_REF,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Referencia de trapos no encontrada: {self.TRAPOS_REF}")
        comp_item_id = int(row[0])

        cur.execute(
            """
SELECT TOP 1 e.f121_rowid
FROM dbo.t121_mc_items_extensiones e
JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid
WHERE i.f120_id = ?
""",
            comp_item_id,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Item trapos sin extension en pruebas: {self.TRAPOS_REF}")
        rowid_item_ext_comp = int(row[0])

        cur.execute(
            "SELECT TOP 1 f150_rowid FROM dbo.t150_mc_bodegas WHERE f150_id = ?",
            item_bodega_id,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Bodega item no encontrada: {item_bodega_id}")
        rowid_bodega_item = int(row[0])

        cur.execute(
            "SELECT TOP 1 f150_rowid FROM dbo.t150_mc_bodegas WHERE f150_id = ?",
            comp_bodega_id,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Bodega componente no encontrada: {comp_bodega_id}")
        rowid_bodega_comp = int(row[0])

        cur.execute(
            "SELECT TOP 1 f808_rowid FROM dbo.t808_mf_rutas WHERE f808_id = ?",
            ruta_id_text,
        )
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"Ruta no encontrada: {ruta_id_text}")
        rowid_ruta = int(row[0])

        return {
            "rowid_usuario": rowid_user,
            "rowid_item_ext_padre": rowid_item_ext_padre,
            "rowid_item_ext_comp": rowid_item_ext_comp,
            "rowid_bodega_item": rowid_bodega_item,
            "rowid_bodega_comp": rowid_bodega_comp,
            "rowid_ruta": rowid_ruta,
        }

    def _create_opm_prueba(self) -> None:
        ref1 = self.ed_ref1.text().strip()
        ref2 = self.ed_ref2.text().strip()
        ref3 = self.ed_ref3.text().strip()
        notes = self.ed_notes.text().strip()
        if not ref1 or not ref2 or not ref3 or not notes:
            QMessageBox.warning(self, "SIESA", "Completa Referencia 1/2/3 y Notas.")
            return
        confirm = QMessageBox.question(
            self,
            "SIESA",
            f"Crear OPM en {self.TEST_DB} con las referencias ingresadas?",
        )
        if confirm != QMessageBox.Yes:
            return

        cfg_test = self._get_siesa_test_config()
        cfg_prod = self._get_siesa_config()
        client_test = SiesaUNOClient(cfg_test)
        client_prod = SiesaUNOClient(cfg_prod)
        if client_test._import_pyodbc() is None:
            QMessageBox.warning(self, "SIESA", "pyodbc no disponible.")
            return
        try:
            import pyodbc
            with pyodbc.connect(client_prod._conn_str(), timeout=10) as prod_conn:
                cur_prod = prod_conn.cursor()
                header = self._fetch_one_dict(
                    cur_prod,
                    """
SELECT f850_id_cia, f850_id_co, f850_id_tipo_docto, f850_id_fecha,
       f850_id_grupo_clase_docto, f850_id_clase_docto, f850_id_clase_op,
       f850_ind_tipo_op, f850_ind_multiples_items, f850_ind_consolida_comp_oper,
       f850_ind_requiere_lm, f850_ind_genera_ordenes_comp, f850_ind_genera_misma_orden,
       f850_ind_genera_todos_niveles, f850_ind_genera_solo_faltantes, f850_ind_metodo_lista_op,
       f850_ind_controla_tep, f850_ind_genera_consumos_tep, f850_ind_genera_entradas_tep,
       f850_id_clase_op_generar, f850_ind_confirmar_al_aprobar, f850_ind_distribucion_costos,
       f850_ind_devolucion_comp, f850_ind_estado, f850_fecha_cumplida,
       f850_rowid_tercero_planif, f850_id_instalacion, f850_rowid_op_padre,
       f850_ind_transmitido, f850_ind_impresion, f850_nro_impresiones,
       f850_ind_posdeduccion, f850_ind_pedido_venta, f850_rowid_pv_docto,
       f850_ind_posdeduccion_tep, f850_ind_reg_incons_posd, f850_ind_lote_automatico,
       f850_ind_controlar_cant_ext, f850_ind_incluir_operacion, f850_ind_entrega_estandar,
       f850_ind_valida_consumo_tot, f850_ind_liq_tep_estandar, f850_ind_no_liq_tep
FROM dbo.t850_mf_op_docto
WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto = ?
""",
                    [self.TEMPLATE_OPM],
                )
                item = self._fetch_one_dict(
                    cur_prod,
                    """
SELECT f851_id_instalacion, f851_id_fecha, f851_ind_estado, f851_fecha_cumplida,
       f851_ind_automatico, f851_id_metodo_lista_mater, f851_id_metodo_ruta,
       f851_fecha_terminacion, f851_fecha_inicio, f851_id_tipo_inv_serv,
       f851_ind_tipo_item, f851_id_unidad_medida, f851_factor,
       f851_cant_planeada_base, f851_cant_ordenada_base, f851_cant_completa_base,
       f851_cant_desechos_base, f851_cant_rechazos_base,
       f851_cant_planeada_1, f851_cant_ordenada_1, f851_cant_completa_1,
       f851_cant_desechos_1, f851_cant_rechazos_1,
       f851_cant_parcial_base, f851_ind_controla_secuencia, f851_porc_rendimiento,
       f851_notas, f851_id_lote
FROM dbo.t851_mf_op_docto_item
WHERE f851_rowid_op_docto IN (
    SELECT f850_rowid
    FROM dbo.t850_mf_op_docto
    WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?
)
""",
                    [self.TEMPLATE_OPM],
                )
                item_ids = self._fetch_one_dict(
                    cur_prod,
                    """
SELECT TOP 1 b.f150_id AS bodega_id, r.f808_id AS ruta_id
FROM dbo.t851_mf_op_docto_item i
LEFT JOIN dbo.t150_mc_bodegas b ON i.f851_rowid_bodega = b.f150_rowid
LEFT JOIN dbo.t808_mf_rutas r ON i.f851_rowid_ruta = r.f808_rowid
WHERE i.f851_rowid_op_docto IN (
    SELECT f850_rowid
    FROM dbo.t850_mf_op_docto
    WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?
)
""",
                    [self.TEMPLATE_OPM],
                )
                comp = self._fetch_one_dict(
                    cur_prod,
                    """
SELECT TOP 1 c.f860_numero_operacion, c.f860_rowid_ctrabajo,
       b.f150_id AS bodega_id, c.f860_id_instalacion, c.f860_id_unidad_medida,
       c.f860_ind_manual, c.f860_factor,
       c.f860_cant_requerida_base, c.f860_cant_requerida_1, c.f860_cant_requerida_2,
       c.f860_cant_desperdicio_base, c.f860_fecha_requerida, c.f860_notas,
       c.f860_codigo_sustitucion, c.f860_cant_equiv_sustitucion,
       c.f860_rowid_movto_entidad, c.f860_ind_cambio_cantidad
FROM dbo.t860_mf_op_componentes c
JOIN dbo.t121_mc_items_extensiones e ON c.f860_rowid_item_ext_componente = e.f121_rowid
JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid
LEFT JOIN dbo.t150_mc_bodegas b ON c.f860_rowid_bodega = b.f150_rowid
WHERE c.f860_rowid_op_docto IN (
    SELECT f850_rowid
    FROM dbo.t850_mf_op_docto
    WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?
)
AND RTRIM(i.f120_referencia) = ?
""",
                    [self.TEMPLATE_OPM, self.TRAPOS_REF],
                )
                if not header or not item or not item_ids or not comp:
                    QMessageBox.warning(self, "SIESA", "No se pudo leer la plantilla base.")
                    return

            with pyodbc.connect(client_test._conn_str(), timeout=10) as test_conn:
                test_conn.autocommit = False
                cur_test = test_conn.cursor()
                cur_test.execute("SELECT DB_NAME()")
                db_name = cur_test.fetchone()[0]
                if str(db_name).upper() != self.TEST_DB:
                    raise RuntimeError(f"DB inesperada: {db_name}")

                next_consec = self._next_opm_consec(cur_test)
                rowids = self._resolve_rowids_test(
                    cur_test,
                    self.TEST_USER,
                    int(item_ids["bodega_id"]),
                    int(comp["bodega_id"]),
                    item_ids["ruta_id"],
                )
                ref1_s = self._safe_text(ref1, 30)
                ref2_s = self._safe_text(ref2, 30)
                ref3_s = self._safe_text(ref3, 30)
                notes_s = self._safe_text(notes, 2000)

                try:
                    cur_test.execute(
                        """
DECLARE @p_retorno smallint, @p_ts datetime, @p_rowid int;
EXEC dbo.sp_mf_op_docto_eventos
    @p_retorno=@p_retorno OUTPUT,
    @p_ts=@p_ts OUTPUT,
    @p_rowid=@p_rowid OUTPUT,
    @p_opcion=?,
    @p_id_cia=?,
    @p_id_co=?,
    @p_id_tipo_docto=?,
    @p_consec_docto=?,
    @p_id_fecha=?,
    @p_id_grupo_clase_docto=?,
    @p_id_clase_docto=?,
    @p_id_clase_op=?,
    @p_ind_tipo_op=?,
    @p_ind_multiples_items=?,
    @p_ind_consolida_comp_oper=?,
    @p_ind_requiere_lm=?,
    @p_ind_genera_ordenes_comp=?,
    @p_ind_genera_misma_orden=?,
    @p_ind_genera_todos_niveles=?,
    @p_ind_genera_solo_faltantes=?,
    @p_ind_metodo_lista_op=?,
    @p_ind_controla_tep=?,
    @p_ind_genera_consumos_tep=?,
    @p_ind_genera_entradas_tep=?,
    @p_id_clase_op_generar=?,
    @p_ind_confirmar_al_aprobar=?,
    @p_ind_distribucion_costos=?,
    @p_ind_devolucion_comp=?,
    @p_ind_estado=?,
    @p_fecha_cumplida=?,
    @p_rowid_tercero_planif=?,
    @p_id_instalacion=?,
    @p_rowid_op_padre=?,
    @p_referencia_1=?,
    @p_referencia_2=?,
    @p_referencia_3=?,
    @p_ind_transmitido=?,
    @p_ind_impresion=?,
    @p_nro_impresiones=?,
    @p_usuario=?,
    @p_rowid_usuario=?,
    @p_notas=?,
    @p_ind_posdeduccion=?,
    @p_ind_pedido_venta=?,
    @p_rowid_pv_docto=?,
    @p_ind_posdeduccion_tep=?,
    @p_ind_reg_incons_posd=?,
    @p_ind_lote_automatico=?,
    @p_ind_controlar_cant_ext=?,
    @p_ind_incluir_operacion=?,
    @p_ind_entrega_estandar=?,
    @p_ind_valida_consumo_tot=?,
    @p_ind_liq_tep_estandar=?,
    @p_ind_no_liq_tep=?;
SELECT @p_retorno AS retorno, @p_ts AS ts, @p_rowid AS rowid;
""",
                        [
                            0,
                            int(header["f850_id_cia"]),
                            header["f850_id_co"],
                            header["f850_id_tipo_docto"],
                            next_consec,
                            header["f850_id_fecha"],
                            int(header["f850_id_grupo_clase_docto"]),
                            int(header["f850_id_clase_docto"]),
                            header["f850_id_clase_op"],
                            int(header["f850_ind_tipo_op"]),
                            int(header["f850_ind_multiples_items"]),
                            int(header["f850_ind_consolida_comp_oper"]),
                            int(header["f850_ind_requiere_lm"]),
                            int(header["f850_ind_genera_ordenes_comp"]),
                            int(header["f850_ind_genera_misma_orden"]),
                            int(header["f850_ind_genera_todos_niveles"]),
                            int(header["f850_ind_genera_solo_faltantes"]),
                            int(header["f850_ind_metodo_lista_op"]),
                            int(header["f850_ind_controla_tep"]),
                            int(header["f850_ind_genera_consumos_tep"]),
                            int(header["f850_ind_genera_entradas_tep"]),
                            header["f850_id_clase_op_generar"],
                            int(header["f850_ind_confirmar_al_aprobar"]),
                            int(header["f850_ind_distribucion_costos"]),
                            int(header["f850_ind_devolucion_comp"]),
                            int(header["f850_ind_estado"]),
                            header["f850_fecha_cumplida"],
                            int(header["f850_rowid_tercero_planif"]),
                            header["f850_id_instalacion"],
                            header["f850_rowid_op_padre"],
                            ref1_s,
                            ref2_s,
                            ref3_s,
                            int(header["f850_ind_transmitido"]),
                            int(header["f850_ind_impresion"]),
                            int(header["f850_nro_impresiones"]),
                            self.TEST_USER,
                            rowids["rowid_usuario"],
                            notes_s,
                            int(header["f850_ind_posdeduccion"]),
                            int(header["f850_ind_pedido_venta"]),
                            header["f850_rowid_pv_docto"],
                            int(header["f850_ind_posdeduccion_tep"]),
                            int(header["f850_ind_reg_incons_posd"]),
                            int(header["f850_ind_lote_automatico"]),
                            int(header["f850_ind_controlar_cant_ext"]),
                            int(header["f850_ind_incluir_operacion"]),
                            int(header["f850_ind_entrega_estandar"]),
                            int(header["f850_ind_valida_consumo_tot"]),
                            int(header["f850_ind_liq_tep_estandar"]),
                            int(header["f850_ind_no_liq_tep"]),
                        ],
                    )
                    out = cur_test.fetchone()
                    if not out:
                        raise RuntimeError("Sin respuesta al crear encabezado.")
                    retorno, _, rowid_op = out
                    if retorno not in (0, None):
                        raise RuntimeError(f"Error creando encabezado: {retorno}")

                    cur_test.execute(
                        """
DECLARE @p_ts datetime, @p_rowid int;
EXEC dbo.sp_mf_op_movto_eventos
    @p_ts=@p_ts OUTPUT,
    @p_rowid=@p_rowid OUTPUT,
    @p_opcion=?,
    @p_id_cia=?,
    @p_rowid_op_docto=?,
    @p_rowid_item_ext_padre=?,
    @p_rowid_bodega=?,
    @p_id_instalacion=?,
    @p_id_fecha=?,
    @p_ind_estado=?,
    @p_fecha_cumplida=?,
    @p_ind_automatico=?,
    @p_id_metodo_lista_mater=?,
    @p_rowid_ruta=?,
    @p_id_metodo_ruta=?,
    @p_fecha_terminacion=?,
    @p_fecha_inicio=?,
    @p_id_tipo_inv_serv=?,
    @p_ind_tipo_item=?,
    @p_id_unidad_medida=?,
    @p_factor=?,
    @p_cant_planeada_base=?,
    @p_cant_ordenada_base=?,
    @p_cant_completa_base=?,
    @p_cant_desechos_base=?,
    @p_cant_rechazos_base=?,
    @p_cant_planeada_1=?,
    @p_cant_ordenada_1=?,
    @p_cant_completa_1=?,
    @p_cant_desechos_1=?,
    @p_cant_rechazos_1=?,
    @p_cant_parcial_base=?,
    @p_ind_controla_secuencia=?,
    @p_porc_rendimiento=?,
    @p_rowid_bodega_componentes=?,
    @p_notas=?,
    @p_diferente=?,
    @p_ind_condicionar=?,
    @p_estado_condicionar=?,
    @p_id_lote=?,
    @p_ind_afectar_items=?,
    @p_seg_selec_item_por_nivel=?,
    @p_rowid_pv_movto=?;
SELECT @p_rowid AS rowid;
""",
                        [
                            0,
                            int(header["f850_id_cia"]),
                            rowid_op,
                            rowids["rowid_item_ext_padre"],
                            rowids["rowid_bodega_item"],
                            item["f851_id_instalacion"],
                            item["f851_id_fecha"],
                            int(item["f851_ind_estado"]),
                            item["f851_fecha_cumplida"],
                            int(item["f851_ind_automatico"]),
                            item["f851_id_metodo_lista_mater"],
                            rowids["rowid_ruta"],
                            item["f851_id_metodo_ruta"],
                            item["f851_fecha_terminacion"],
                            item["f851_fecha_inicio"],
                            item["f851_id_tipo_inv_serv"],
                            int(item["f851_ind_tipo_item"]),
                            item["f851_id_unidad_medida"],
                            float(item["f851_factor"]),
                            float(item["f851_cant_planeada_base"]),
                            float(item["f851_cant_ordenada_base"]),
                            float(item["f851_cant_completa_base"]),
                            float(item["f851_cant_desechos_base"]),
                            float(item["f851_cant_rechazos_base"]),
                            float(item["f851_cant_planeada_1"]),
                            float(item["f851_cant_ordenada_1"]),
                            float(item["f851_cant_completa_1"]),
                            float(item["f851_cant_desechos_1"]),
                            float(item["f851_cant_rechazos_1"]),
                            float(item["f851_cant_parcial_base"]),
                            int(item["f851_ind_controla_secuencia"]),
                            float(item["f851_porc_rendimiento"]),
                            None,
                            item["f851_notas"],
                            0,
                            0,
                            0,
                            item["f851_id_lote"],
                            0,
                            0,
                            None,
                        ],
                    )
                    out = cur_test.fetchone()
                    if not out:
                        raise RuntimeError("Sin respuesta al crear item.")
                    rowid_op_item = out[0]

                    cur_test.execute(
                        """
DECLARE @p_ts datetime, @p_rowid int;
EXEC dbo.sp_mf_op_comp_eventos
    @p_ts=@p_ts OUTPUT,
    @p_rowid=@p_rowid OUTPUT,
    @p_opcion=?,
    @p_id_cia=?,
    @p_rowid_op_docto_item=?,
    @p_rowid_op_docto=?,
    @p_rowid_item_ext_padre=?,
    @p_rowid_item_ext_componente=?,
    @p_numero_operacion=?,
    @p_rowid_ctrabajo=?,
    @p_rowid_bodega=?,
    @p_id_instalacion=?,
    @p_id_unidad_medida=?,
    @p_ind_manual=?,
    @p_factor=?,
    @p_cant_requerida_base=?,
    @p_cant_comprometida_base=?,
    @p_cant_consumida_base=?,
    @p_cant_requerida_1=?,
    @p_cant_comprometida_1=?,
    @p_cant_consumida_1=?,
    @p_cant_requerida_2=?,
    @p_cant_comprometida_2=?,
    @p_cant_consumida_2=?,
    @p_cant_desperdicio_base=?,
    @p_fecha_requerida=?,
    @p_notas=?,
    @p_usuario=?,
    @p_rowid_item_ext_sustituido=?,
    @p_codigo_sustitucion=?,
    @p_cant_equiv_sustitucion=?,
    @p_permiso_costos=?,
    @p_rowid_movto_entidad=?,
    @p_ind_cambio_cantidad=?;
SELECT @p_rowid AS rowid;
""",
                        [
                            0,
                            int(header["f850_id_cia"]),
                            rowid_op_item,
                            rowid_op,
                            rowids["rowid_item_ext_padre"],
                            rowids["rowid_item_ext_comp"],
                            int(comp["f860_numero_operacion"]),
                            comp["f860_rowid_ctrabajo"],
                            rowids["rowid_bodega_comp"],
                            comp["f860_id_instalacion"],
                            comp["f860_id_unidad_medida"],
                            int(comp["f860_ind_manual"]),
                            float(comp["f860_factor"]),
                            float(comp["f860_cant_requerida_base"]),
                            0.0,
                            0.0,
                            float(comp["f860_cant_requerida_1"]),
                            0.0,
                            0.0,
                            float(comp["f860_cant_requerida_2"]),
                            0.0,
                            0.0,
                            float(comp["f860_cant_desperdicio_base"]),
                            comp["f860_fecha_requerida"],
                            comp["f860_notas"],
                            self.TEST_USER,
                            None,
                            int(comp["f860_codigo_sustitucion"]),
                            float(comp["f860_cant_equiv_sustitucion"]),
                            0,
                            comp["f860_rowid_movto_entidad"],
                            int(comp["f860_ind_cambio_cantidad"]),
                        ],
                    )
                    out = cur_test.fetchone()
                    if not out:
                        raise RuntimeError("Sin respuesta al crear componente.")

                    test_conn.commit()
                except Exception:
                    test_conn.rollback()
                    raise

            self.ed_created_opm.setText(str(next_consec))
            self._last_created_opm = str(next_consec)
            self._append_log(f"OPM PRUEBA CREADA: {next_consec} ({self.TEST_DB})")
            QMessageBox.information(self, "SIESA", f"OPM creada en {self.TEST_DB}: {next_consec}")
        except Exception as e:
            QMessageBox.warning(self, "SIESA", f"Error creando OPM prueba: {e}")

    def _load_template_list(self) -> None:
        target_text = self.ed_created_opm.text().strip() or self._last_created_opm
        if not target_text:
            QMessageBox.warning(self, "SIESA", "Primero crea una OPM de prueba.")
            return
        try:
            target_op = int(target_text)
        except Exception:
            QMessageBox.warning(self, "SIESA", "OPM invalida para cargar listado.")
            return

        cfg_test = self._get_siesa_test_config()
        cfg_prod = self._get_siesa_config()
        client_test = SiesaUNOClient(cfg_test)
        client_prod = SiesaUNOClient(cfg_prod)
        if client_test._import_pyodbc() is None:
            QMessageBox.warning(self, "SIESA", "pyodbc no disponible.")
            return

        try:
            import pyodbc
            with pyodbc.connect(client_prod._conn_str(), timeout=10) as prod_conn:
                cur_prod = prod_conn.cursor()
                comps = self._fetch_all_dicts(
                    cur_prod,
                    """
SELECT c.f860_numero_operacion, c.f860_rowid_ctrabajo,
       b.f150_id AS bodega_id, c.f860_id_instalacion, c.f860_id_unidad_medida,
       c.f860_ind_manual, c.f860_factor,
       c.f860_cant_requerida_base, c.f860_cant_requerida_1, c.f860_cant_requerida_2,
       c.f860_cant_desperdicio_base, c.f860_fecha_requerida, c.f860_notas,
       c.f860_rowid_item_ext_sustituido, c.f860_codigo_sustitucion,
       c.f860_cant_equiv_sustitucion, c.f860_rowid_movto_entidad, c.f860_ind_cambio_cantidad,
       RTRIM(i.f120_referencia) AS referencia,
       RTRIM(isub.f120_referencia) AS referencia_sustituido
FROM dbo.t860_mf_op_componentes c
JOIN dbo.t121_mc_items_extensiones e ON c.f860_rowid_item_ext_componente = e.f121_rowid
JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid
LEFT JOIN dbo.t121_mc_items_extensiones esub ON c.f860_rowid_item_ext_sustituido = esub.f121_rowid
LEFT JOIN dbo.t120_mc_items isub ON esub.f121_rowid_item = isub.f120_rowid
LEFT JOIN dbo.t150_mc_bodegas b ON c.f860_rowid_bodega = b.f150_rowid
WHERE c.f860_rowid_op_docto IN (
    SELECT f850_rowid
    FROM dbo.t850_mf_op_docto
    WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?
)
ORDER BY c.f860_rowid
""",
                    [self.TEMPLATE_OPM],
                )
            if not comps:
                QMessageBox.warning(self, "SIESA", "No se encontro listado de la plantilla.")
                return

            template_comps = []
            for comp in comps:
                ref = self._safe_text(comp.get("referencia"), 50)
                if not ref or ref == self.TRAPOS_REF:
                    continue
                template_comps.append(comp)
            if not template_comps:
                QMessageBox.warning(self, "SIESA", "La plantilla no tiene componentes para cargar.")
                return

            confirm = QMessageBox.question(
                self,
                "SIESA",
                f"Cargar {len(template_comps)} componentes en OPM {target_op} ({self.TEST_DB})?",
            )
            if confirm != QMessageBox.Yes:
                return

            with pyodbc.connect(client_test._conn_str(), timeout=10) as test_conn:
                test_conn.autocommit = False
                cur_test = test_conn.cursor()
                cur_test.execute("SELECT DB_NAME()")
                db_name = cur_test.fetchone()[0]
                if str(db_name).upper() != self.TEST_DB:
                    raise RuntimeError(f"DB inesperada: {db_name}")

                op_row = self._fetch_one_dict(
                    cur_test,
                    """
SELECT f850_rowid, f850_id_cia
FROM dbo.t850_mf_op_docto
WHERE f850_id_tipo_docto='OPM' AND f850_consec_docto=?
""",
                    [target_op],
                )
                if not op_row:
                    QMessageBox.warning(self, "SIESA", "No se encontro la OPM en pruebas.")
                    return
                rowid_op = int(op_row["f850_rowid"])
                id_cia = int(op_row["f850_id_cia"])

                item_row = self._fetch_one_dict(
                    cur_test,
                    """
SELECT TOP 1 f851_rowid, f851_rowid_item_ext_padre
FROM dbo.t851_mf_op_docto_item
WHERE f851_rowid_op_docto=?
""",
                    [rowid_op],
                )
                if not item_row:
                    QMessageBox.warning(self, "SIESA", "OPM sin item base en pruebas.")
                    return
                rowid_op_item = int(item_row["f851_rowid"])
                rowid_item_ext_padre = int(item_row["f851_rowid_item_ext_padre"])

                existing_rows = self._fetch_all_dicts(
                    cur_test,
                    """
SELECT RTRIM(i.f120_referencia) AS referencia
FROM dbo.t860_mf_op_componentes c
JOIN dbo.t121_mc_items_extensiones e ON c.f860_rowid_item_ext_componente = e.f121_rowid
JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid
WHERE c.f860_rowid_op_docto=?
""",
                    [rowid_op],
                )
                existing_refs = {
                    self._safe_text(row.get("referencia"), 50)
                    for row in existing_rows
                    if row.get("referencia")
                }

                desired_refs = {
                    self._safe_text(comp.get("referencia"), 50)
                    for comp in template_comps
                    if comp.get("referencia")
                }
                missing_refs = []
                ref_rowid_map = {}
                for ref in sorted(desired_refs):
                    if ref in existing_refs:
                        continue
                    row = self._fetch_one_dict(
                        cur_test,
                        """
SELECT TOP 1 e.f121_rowid
FROM dbo.t121_mc_items_extensiones e
JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid
WHERE RTRIM(i.f120_referencia) = ?
""",
                        [ref],
                    )
                    if not row:
                        missing_refs.append(ref)
                    else:
                        ref_rowid_map[ref] = int(row["f121_rowid"])

                bodega_ids = {
                    int(comp["bodega_id"])
                    for comp in template_comps
                    if comp.get("bodega_id") is not None
                }
                missing_bodegas = []
                bodega_rowid_map = {}
                for bodega_id in sorted(bodega_ids):
                    row = self._fetch_one_dict(
                        cur_test,
                        "SELECT TOP 1 f150_rowid FROM dbo.t150_mc_bodegas WHERE f150_id = ?",
                        [bodega_id],
                    )
                    if not row:
                        missing_bodegas.append(str(bodega_id))
                    else:
                        bodega_rowid_map[bodega_id] = int(row["f150_rowid"])

                sub_refs = {
                    self._safe_text(comp.get("referencia_sustituido"), 50)
                    for comp in template_comps
                    if comp.get("referencia_sustituido")
                }
                missing_subs = []
                sub_ref_rowid_map = {}
                for sub_ref in sorted(sub_refs):
                    row = self._fetch_one_dict(
                        cur_test,
                        """
SELECT TOP 1 e.f121_rowid
FROM dbo.t121_mc_items_extensiones e
JOIN dbo.t120_mc_items i ON e.f121_rowid_item = i.f120_rowid
WHERE RTRIM(i.f120_referencia) = ?
""",
                        [sub_ref],
                    )
                    if not row:
                        missing_subs.append(sub_ref)
                    else:
                        sub_ref_rowid_map[sub_ref] = int(row["f121_rowid"])

                if missing_refs or missing_bodegas or missing_subs:
                    msg = []
                    if missing_refs:
                        msg.append("Items faltantes: " + ", ".join(missing_refs[:8]))
                    if missing_bodegas:
                        msg.append("Bodegas faltantes: " + ", ".join(missing_bodegas[:8]))
                    if missing_subs:
                        msg.append("Sustitutos faltantes: " + ", ".join(missing_subs[:8]))
                    QMessageBox.warning(self, "SIESA", "No se puede cargar. " + " | ".join(msg))
                    return

                insertable = [
                    comp for comp in template_comps
                    if self._safe_text(comp.get("referencia"), 50) not in existing_refs
                ]
                if not insertable:
                    QMessageBox.information(self, "SIESA", "No hay componentes nuevos para cargar.")
                    return

                try:
                    for comp in insertable:
                        ref = self._safe_text(comp.get("referencia"), 50)
                        rowid_item_ext_comp = ref_rowid_map.get(ref)
                        if not rowid_item_ext_comp:
                            raise RuntimeError(f"Item sin extension en pruebas: {ref}")
                        bodega_id = int(comp["bodega_id"]) if comp.get("bodega_id") is not None else None
                        rowid_bodega_comp = bodega_rowid_map.get(bodega_id)
                        if bodega_id is not None and not rowid_bodega_comp:
                            raise RuntimeError(f"Bodega no encontrada en pruebas: {bodega_id}")
                        sub_ref = self._safe_text(comp.get("referencia_sustituido"), 50)
                        rowid_item_ext_sub = sub_ref_rowid_map.get(sub_ref) if sub_ref else None

                        cur_test.execute(
                            """
DECLARE @p_ts datetime, @p_rowid int;
EXEC dbo.sp_mf_op_comp_eventos
    @p_ts=@p_ts OUTPUT,
    @p_rowid=@p_rowid OUTPUT,
    @p_opcion=?,
    @p_id_cia=?,
    @p_rowid_op_docto_item=?,
    @p_rowid_op_docto=?,
    @p_rowid_item_ext_padre=?,
    @p_rowid_item_ext_componente=?,
    @p_numero_operacion=?,
    @p_rowid_ctrabajo=?,
    @p_rowid_bodega=?,
    @p_id_instalacion=?,
    @p_id_unidad_medida=?,
    @p_ind_manual=?,
    @p_factor=?,
    @p_cant_requerida_base=?,
    @p_cant_comprometida_base=?,
    @p_cant_consumida_base=?,
    @p_cant_requerida_1=?,
    @p_cant_comprometida_1=?,
    @p_cant_consumida_1=?,
    @p_cant_requerida_2=?,
    @p_cant_comprometida_2=?,
    @p_cant_consumida_2=?,
    @p_cant_desperdicio_base=?,
    @p_fecha_requerida=?,
    @p_notas=?,
    @p_usuario=?,
    @p_rowid_item_ext_sustituido=?,
    @p_codigo_sustitucion=?,
    @p_cant_equiv_sustitucion=?,
    @p_permiso_costos=?,
    @p_rowid_movto_entidad=?,
    @p_ind_cambio_cantidad=?;
SELECT @p_rowid AS rowid;
""",
                            [
                                0,
                                id_cia,
                                rowid_op_item,
                                rowid_op,
                                rowid_item_ext_padre,
                                rowid_item_ext_comp,
                                int(comp["f860_numero_operacion"]),
                                comp["f860_rowid_ctrabajo"],
                                rowid_bodega_comp,
                                comp["f860_id_instalacion"],
                                comp["f860_id_unidad_medida"],
                                int(comp["f860_ind_manual"]),
                                float(comp["f860_factor"]),
                                float(comp["f860_cant_requerida_base"]),
                                0.0,
                                0.0,
                                float(comp["f860_cant_requerida_1"]),
                                0.0,
                                0.0,
                                float(comp["f860_cant_requerida_2"]),
                                0.0,
                                0.0,
                                float(comp["f860_cant_desperdicio_base"]),
                                comp["f860_fecha_requerida"],
                                comp["f860_notas"],
                                self.TEST_USER,
                                rowid_item_ext_sub,
                                int(comp["f860_codigo_sustitucion"]),
                                float(comp["f860_cant_equiv_sustitucion"]),
                                0,
                                comp["f860_rowid_movto_entidad"],
                                int(comp["f860_ind_cambio_cantidad"]),
                            ],
                        )
                        out = cur_test.fetchone()
                        if not out:
                            raise RuntimeError(f"Sin respuesta insertando {ref}.")

                    test_conn.commit()
                except Exception:
                    test_conn.rollback()
                    raise

            skipped = len(template_comps) - len(insertable)
            self._append_log(
                f"LISTADO CARGADO: OPM={target_op} OK={len(insertable)} SKIP={skipped}"
            )
            QMessageBox.information(
                self,
                "SIESA",
                f"Listado cargado en OPM {target_op}. Agregados: {len(insertable)}.",
            )
        except Exception as e:
            QMessageBox.warning(self, "SIESA", f"Error cargando listado: {e}")

    def _order_date_key(self, order: dict) -> datetime:
        value = order.get("FechaDocto")
        if isinstance(value, datetime):
            return value
        if value is None:
            return datetime.min
        text = str(value).strip()
        if not text:
            return datetime.min
        try:
            return datetime.fromisoformat(text)
        except Exception:
            return datetime.min

    def _orders_months(self) -> int:
        try:
            return int(self.spin_orders_months.value())
        except Exception:
            return 2

    def _load_orders_preview(self) -> None:
        months = self._orders_months()
        cfg = self._get_siesa_config()
        client = SiesaUNOClient(cfg)
        if client._import_pyodbc() is None:
            self._append_log("ERROR ORDENES: pyodbc no disponible.")
            QMessageBox.warning(self, "SIESA", "pyodbc no disponible para cargar ordenes.")
            return
        try:
            import pyodbc
            with pyodbc.connect(client._conn_str(), timeout=10) as conn:
                cur = conn.cursor()
                cur.execute("SELECT DB_NAME()")
                _ = cur.fetchone()

                cur.execute(
                    """
SELECT DISTINCT
  b850_op_numero            AS OpNumero,
  b850_op_fecha_elaboracion AS FechaDocto,
  b850_op_desc_estado       AS Estado,
  b850_notas                AS Notas,
  b850_op_docto_referencia1 AS OpReferencia1,
  b850_op_docto_referencia2 AS OpReferencia2
FROM dbo.vBI_850
WHERE b850_op_co_id = RIGHT('000' + CAST(? AS varchar(3)), 3)
  AND b850_op_fecha_elaboracion >= DATEADD(MONTH, -?, CAST(GETDATE() AS date))
""",
                    1,
                    max(1, months),
                )
                rows = cur.fetchall()
                cols = [d[0] for d in cur.description] if cur.description else []
                orders = [dict(zip(cols, r)) for r in rows]
        except Exception as e:
            self._append_log(f"ERROR ORDENES: {e}")
            QMessageBox.warning(self, "SIESA", f"Error cargando ordenes: {e}")
            return

        if not orders:
            self.lbl_orders.setText(f"ORDENES (ULTIMOS {months} MESES): 0")
            self._orders_cache = []
            self._apply_orders_filter()
            return
        orders_sorted = sorted(orders, key=self._order_date_key, reverse=True)
        self._orders_cache = orders_sorted
        self._apply_orders_filter()

    def _apply_orders_filter(self) -> None:
        self.orders_table.setRowCount(0)
        if not self._orders_cache:
            return
        term = self.ed_op_filter.text().strip()
        if term:
            filtered = [
                o for o in self._orders_cache
                if term in str(o.get("OpNumero", "")).strip()
            ]
            months = self._orders_months()
            self.lbl_orders.setText(
                f"ORDENES (ULTIMOS {months} MESES): {len(filtered)} / {len(self._orders_cache)}"
            )
        else:
            filtered = self._orders_cache
            months = self._orders_months()
            self.lbl_orders.setText(f"ORDENES (ULTIMOS {months} MESES): {len(self._orders_cache)}")
        limit = 200
        for order in filtered[:limit]:
            row = self.orders_table.rowCount()
            self.orders_table.insertRow(row)
            vals = [
                order.get("OpNumero", ""),
                order.get("FechaDocto", ""),
                order.get("Estado", ""),
                order.get("Notas", ""),
                order.get("OpReferencia1", ""),
                order.get("OpReferencia2", ""),
            ]
            for c, v in enumerate(vals):
                text = "" if v is None else str(v)
                item = QTableWidgetItem(text)
                if c == 3 and len(text) > 80:
                    item.setToolTip(text)
                self.orders_table.setItem(row, c, item)
        self._orders_auto_height()

    def _orders_auto_height(self) -> None:
        rows = max(self.orders_table.rowCount(), 1)
        vh = self.orders_table.verticalHeader()
        row_h = vh.defaultSectionSize()
        vlen = vh.length()
        if vlen <= 0:
            vlen = rows * row_h
        header_h = self.orders_table.horizontalHeader().height() if self.orders_table.horizontalHeader() else 0
        frame = self.orders_table.frameWidth() * 2
        total_h = header_h + vlen + frame + 8
        self.orders_table.setMinimumHeight(total_h)
        self.orders_table.setMaximumHeight(total_h)

    def _ops_items_auto_height(self) -> None:
        rows = max(self.ops_items_table.rowCount(), 1)
        vh = self.ops_items_table.verticalHeader()
        row_h = vh.defaultSectionSize()
        vlen = vh.length()
        if vlen <= 0:
            vlen = rows * row_h
        header_h = self.ops_items_table.horizontalHeader().height() if self.ops_items_table.horizontalHeader() else 0
        frame = self.ops_items_table.frameWidth() * 2
        total_h = header_h + vlen + frame + 8
        self.ops_items_table.setMinimumHeight(total_h)
        self.ops_items_table.setMaximumHeight(total_h)

    def _ops_items_load(self) -> None:
        op_text = self.ed_ops_items_op.text().strip()
        if not op_text:
            QMessageBox.warning(self, "SIESA", "Ingresa el numero de OP.")
            return
        try:
            op_num = int(op_text)
        except Exception:
            QMessageBox.warning(self, "SIESA", "OP invalida.")
            return

        cfg_test = self._get_siesa_test_config()
        client_test = SiesaUNOClient(cfg_test)
        if client_test._import_pyodbc() is None:
            QMessageBox.warning(self, "SIESA", "pyodbc no disponible.")
            return

        try:
            import pyodbc
            with pyodbc.connect(client_test._conn_str(), timeout=10) as conn:
                cur = conn.cursor()
                cur.execute("SELECT DB_NAME()")
                db_name = str(cur.fetchone()[0]).upper()
                if db_name != self.TEST_DB:
                    raise RuntimeError(f"DB inesperada: {db_name}")

                cur.execute(
                    """
SELECT f850_rowid, f850_id_cia
FROM dbo.t850_mf_op_docto
WHERE f850_id_tipo_docto='OPS' AND f850_consec_docto=?
""",
                    op_num,
                )
                row = cur.fetchone()
                if not row:
                    QMessageBox.warning(self, "SIESA", "OP no encontrada en PRUEBAWESTON.")
                    return
                self._ops_rowid_op = int(row[0])
                self._ops_id_cia = int(row[1])

                cur.execute(
                    """
SELECT i.f851_rowid,
       i.f851_cant_planeada_base,
       i.f851_id_unidad_medida,
       b.f150_id AS bodega_id,
       RTRIM(it.f120_referencia) AS item_ref,
       it.f120_descripcion AS item_desc
FROM dbo.t851_mf_op_docto_item i
LEFT JOIN dbo.t121_mc_items_extensiones e ON i.f851_rowid_item_ext_padre = e.f121_rowid
LEFT JOIN dbo.t120_mc_items it ON e.f121_rowid_item = it.f120_rowid
LEFT JOIN dbo.t150_mc_bodegas b ON i.f851_rowid_bodega = b.f150_rowid
WHERE i.f851_rowid_op_docto=?
ORDER BY i.f851_rowid
""",
                    self._ops_rowid_op,
                )
                rows = cur.fetchall()
                cols = [d[0] for d in cur.description]
                items = [dict(zip(cols, r)) for r in rows]

            self.ops_items_table.setRowCount(0)
            for item in items:
                row = self.ops_items_table.rowCount()
                self.ops_items_table.insertRow(row)
                ref = str(item.get("item_ref", "")).strip()
                desc = "" if item.get("item_desc") is None else str(item.get("item_desc")).strip()
                bod = "" if item.get("bodega_id") is None else str(item.get("bodega_id")).strip()
                um = "" if item.get("f851_id_unidad_medida") is None else str(item.get("f851_id_unidad_medida")).strip()
                qty = float(item.get("f851_cant_planeada_base") or 0)
                qty_val = int(qty) if qty.is_integer() else qty
                cells = [ref, desc, bod, um, str(qty_val)]
                for c, text in enumerate(cells):
                    cell = QTableWidgetItem(text)
                    if c == 0:
                        cell.setData(Qt.UserRole, {
                            "rowid": int(item.get("f851_rowid")),
                            "qty": qty,
                        })
                    if c == 4:
                        cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.ops_items_table.setItem(row, c, cell)
            self._ops_items_auto_height()
            self._ops_selected_item_rowid = None
            self._ops_selected_item_qty = None
            self.ed_ops_items_qty.clear()
        except Exception as e:
            QMessageBox.warning(self, "SIESA", f"Error cargando items: {e}")

    def _ops_items_select(self) -> None:
        items = self.ops_items_table.selectedItems()
        if not items:
            return
        row = items[0].row()
        ref_item = self.ops_items_table.item(row, 0)
        if not ref_item:
            return
        meta = ref_item.data(Qt.UserRole) or {}
        rowid = meta.get("rowid")
        qty = meta.get("qty")
        if rowid is None:
            return
        self._ops_selected_item_rowid = int(rowid)
        try:
            qty_val = float(qty)
        except Exception:
            qty_val = None
        self._ops_selected_item_qty = qty_val
        if qty_val is not None:
            self.ed_ops_items_qty.setText(str(int(qty_val) if qty_val.is_integer() else qty_val))

    def _ops_items_update(self) -> None:
        if not self._ops_selected_item_rowid:
            QMessageBox.warning(self, "SIESA", "Selecciona un item.")
            return
        qty_text = self.ed_ops_items_qty.text().strip()
        if not qty_text:
            QMessageBox.warning(self, "SIESA", "Ingresa la nueva cantidad.")
            return
        try:
            new_qty = float(qty_text)
            if new_qty.is_integer():
                new_qty = int(new_qty)
        except Exception:
            QMessageBox.warning(self, "SIESA", "Cantidad invalida.")
            return
        old_qty = self._ops_selected_item_qty or 0
        if old_qty <= 0:
            QMessageBox.warning(self, "SIESA", "Cantidad actual invalida para escalar.")
            return
        factor = float(new_qty) / float(old_qty)

        confirm = QMessageBox.question(
            self,
            "SIESA",
            f"Actualizar item de {old_qty} a {new_qty}? Esto ajusta materiales.",
        )
        if confirm != QMessageBox.Yes:
            return

        cfg_test = self._get_siesa_test_config()
        client_test = SiesaUNOClient(cfg_test)
        if client_test._import_pyodbc() is None:
            QMessageBox.warning(self, "SIESA", "pyodbc no disponible.")
            return

        try:
            import pyodbc
            with pyodbc.connect(client_test._conn_str(), timeout=10) as conn:
                conn.autocommit = False
                cur = conn.cursor()
                cur.execute("SELECT DB_NAME()")
                db_name = str(cur.fetchone()[0]).upper()
                if db_name != self.TEST_DB:
                    raise RuntimeError(f"DB inesperada: {db_name}")

                cur.execute(
                    """
SELECT f860_rowid,
       f860_cant_requerida_base,
       f860_cant_requerida_1,
       f860_cant_requerida_2,
       f860_cant_comprometida_base,
       f860_cant_comprometida_1,
       f860_cant_comprometida_2,
       f860_cant_consumida_base,
       f860_cant_consumida_1,
       f860_cant_consumida_2
FROM dbo.t860_mf_op_componentes
WHERE f860_rowid_op_docto_item = ?
ORDER BY f860_rowid
""",
                    self._ops_selected_item_rowid,
                )
                comp_rows = cur.fetchall()
                comp_cols = [d[0] for d in cur.description] if cur.description else []
                comps = [dict(zip(comp_cols, row)) for row in comp_rows]

                def _intish(val: float, tol: float = 1e-6) -> bool:
                    return abs(val - round(val)) < tol

                use_integer_split = False
                new_base_ints: list[int] = []
                old_qty_val = float(old_qty)
                new_qty_val = float(new_qty)
                if comps and _intish(old_qty_val) and _intish(new_qty_val):
                    old_qty_int = int(round(old_qty_val))
                    new_qty_int = int(round(new_qty_val))
                    base_vals = [float(comp.get("f860_cant_requerida_base") or 0) for comp in comps]
                    total_base = sum(base_vals)
                    if _intish(total_base) and int(round(total_base)) == old_qty_int:
                        if all(_intish(val) for val in base_vals):
                            max_idx = max(range(len(base_vals)), key=lambda i: base_vals[i])
                            if base_vals[max_idx] > 0:
                                delta = new_qty_int - old_qty_int
                                if delta == 0 or base_vals[max_idx] + delta >= 0:
                                    new_base_ints = [int(round(val)) for val in base_vals]
                                    new_base_ints[max_idx] += delta
                                    use_integer_split = True

                cur.execute(
                    """
UPDATE dbo.t851_mf_op_docto_item
SET f851_cant_planeada_base = ?,
    f851_cant_ordenada_base = f851_cant_ordenada_base * ?,
    f851_cant_completa_base = f851_cant_completa_base * ?,
    f851_cant_planeada_1 = f851_cant_planeada_1 * ?,
    f851_cant_ordenada_1 = f851_cant_ordenada_1 * ?,
    f851_cant_completa_1 = f851_cant_completa_1 * ?,
      f851_cant_parcial_base = f851_cant_parcial_base * ?
WHERE f851_rowid = ?
""",
                    new_qty,
                    factor,
                    factor,
                    factor,
                    factor,
                    factor,
                    factor,
                    self._ops_selected_item_rowid,
                )
                if use_integer_split:
                    for comp, new_base in zip(comps, new_base_ints):
                        rowid = int(comp.get("f860_rowid"))
                        old_base = float(comp.get("f860_cant_requerida_base") or 0)
                        if old_base > 0:
                            factor_comp = float(new_base) / old_base
                        else:
                            factor_comp = 0.0

                        def _scale(key: str) -> float:
                            return float(comp.get(key) or 0) * factor_comp

                        cur.execute(
                            """
UPDATE dbo.t860_mf_op_componentes
SET f860_cant_requerida_base = ?,
    f860_cant_requerida_1 = ?,
    f860_cant_requerida_2 = ?,
    f860_cant_comprometida_base = ?,
    f860_cant_comprometida_1 = ?,
    f860_cant_comprometida_2 = ?,
    f860_cant_consumida_base = ?,
    f860_cant_consumida_1 = ?,
    f860_cant_consumida_2 = ?,
    f860_ind_cambio_cantidad = 1
WHERE f860_rowid = ?
""",
                            new_base,
                            _scale("f860_cant_requerida_1"),
                            _scale("f860_cant_requerida_2"),
                            _scale("f860_cant_comprometida_base"),
                            _scale("f860_cant_comprometida_1"),
                            _scale("f860_cant_comprometida_2"),
                            _scale("f860_cant_consumida_base"),
                            _scale("f860_cant_consumida_1"),
                            _scale("f860_cant_consumida_2"),
                            rowid,
                        )
                else:
                    cur.execute(
                        """
UPDATE dbo.t860_mf_op_componentes
SET f860_cant_requerida_base = f860_cant_requerida_base * ?,
    f860_cant_requerida_1 = f860_cant_requerida_1 * ?,
    f860_cant_requerida_2 = f860_cant_requerida_2 * ?,
    f860_cant_comprometida_base = f860_cant_comprometida_base * ?,
    f860_cant_comprometida_1 = f860_cant_comprometida_1 * ?,
    f860_cant_comprometida_2 = f860_cant_comprometida_2 * ?,
    f860_cant_consumida_base = f860_cant_consumida_base * ?,
    f860_cant_consumida_1 = f860_cant_consumida_1 * ?,
    f860_cant_consumida_2 = f860_cant_consumida_2 * ?,
    f860_ind_cambio_cantidad = 1
WHERE f860_rowid_op_docto_item = ?
""",
                        factor,
                        factor,
                        factor,
                        factor,
                        factor,
                        factor,
                        factor,
                        factor,
                        factor,
                        self._ops_selected_item_rowid,
                    )
                conn.commit()
            self._ops_items_load()
            QMessageBox.information(self, "SIESA", "Item actualizado.")
        except Exception as e:
            QMessageBox.warning(self, "SIESA", f"Error actualizando item: {e}")

    def _load_package(self) -> None:
        try:
            hub = load_hub(verify_hashes=True)
            man = hub.manifest
            self.lbl_manifest.setText(
                f"PAQUETE: {man.package_name} | FECHA: {man.generated_at} | ARCHIVOS: {len(man.files)}"
            )
            self.txt_log.clear()
            self._append_log("MANIFEST OK")
            self._append_log(f"ARCHIVOS: {len(man.files)}")
            for f in man.files[:50]:
                self._append_log(f"- {f.path} ({f.kind})")
            if len(man.files) > 50:
                self._append_log(f"... ({len(man.files) - 50} más)")
            self._load_orders_preview()
        except PackageError as e:
            self.lbl_manifest.setText("ERROR AL CARGAR PAQUETE")
            self._append_log(f"ERROR: {e}")
            QMessageBox.warning(self, "SIESA", str(e))
        except Exception as e:
            self.lbl_manifest.setText("ERROR AL CARGAR PAQUETE")
            self._append_log(f"ERROR: {e}")
            QMessageBox.warning(self, "SIESA", f"Error inesperado: {e}")

    def _choose_and_load(self) -> None:
        start_dir = str(get_company_data_dir())
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de paquete", start_dir)
        if folder:
            set_company_data_dir(Path(folder))
            self._update_folder_label()
        self._load_package()

    def _run_diagnostic(self) -> None:
        cfg = self._get_siesa_config()
        client = SiesaUNOClient(cfg)
        if cfg.auth == "sql":
            auth_label = "SQL"
        elif cfg.auth == "credman":
            auth_label = "CREDMAN"
        else:
            auth_label = "WINDOWS"
        ok_py, msg_py = client.pyodbc_status()
        ok_ping, msg_ping = client.ping_host()
        ok_tcp, msg_tcp = client.tcp_port_open(1433)
        ok_sql, msg_sql = client.test_sql_login()
        drivers = client.odbc_drivers()

        lines = [
            "=== DIAGNOSTICO SIESA UNO ===",
            f"AUTH: {auth_label}",
            f"PYTHON: {sys.executable}",
            f"PYODBC: {'OK' if ok_py else 'FAIL'} -> {msg_py}",
            f"PING: {'OK' if ok_ping else 'FAIL'} -> {msg_ping}",
            f"TCP 1433: {'OK' if ok_tcp else 'FAIL'} -> {msg_tcp}",
            f"SQL LOGIN: {'OK' if ok_sql else 'FAIL'} -> {msg_sql}",
        ]
        if drivers:
            lines.append(f"DRIVERS ODBC: {', '.join(drivers)}")
        else:
            lines.append("DRIVERS ODBC: No detectados")
        QMessageBox.information(self, "SIESA", "\n".join(lines))

    def _update_snapshots(self) -> None:
        script = Path(__file__).resolve().parents[2] / "data" / "siesa" / "export_siesa_to_package.py"
        if not script.exists():
            QMessageBox.warning(self, "SIESA", f"No se encontro el script: {script}")
            return
        try:
            env = os.environ.copy()
            cfg = self._get_siesa_config()
            if cfg.auth == "sql" and cfg.user and cfg.password:
                env["SIESA_AUTH"] = "sql"
                env["SIESA_USER"] = cfg.user
                env["SIESA_PASSWORD"] = cfg.password
            elif cfg.auth == "credman":
                env["SIESA_AUTH"] = "credman"
                env["SIESA_CRED_TARGET"] = cfg.cred_target or "CalvoSiesaUNOEE"
                env["SIESA_CRED_USER"] = cfg.cred_user or "sa"
            proc = subprocess.run(
                [sys.executable, str(script)],
                cwd=str(script.parent),
                capture_output=True,
                text=True,
                check=False,
                env=env,
            )
            self._append_log("=== ACTUALIZAR SNAPSHOTS ===")
            if proc.stdout:
                self._append_log(proc.stdout.strip())
            if proc.stderr:
                self._append_log(proc.stderr.strip())
            if proc.returncode != 0:
                QMessageBox.warning(self, "SIESA", "Error al actualizar snapshots. Revisa el log.")
                return
            QMessageBox.information(self, "SIESA", "Snapshots actualizados.")
            self._load_package()
        except Exception as e:
            QMessageBox.warning(self, "SIESA", f"Error ejecutando export: {e}")

    def _get_siesa_config(self) -> SiesaConfig:
        return SiesaConfig(auth="credman", cred_target="CalvoSiesaUNOEE", cred_user="sa")

    def _get_siesa_test_config(self) -> SiesaConfig:
        return SiesaConfig(
            auth="credman",
            cred_target="CalvoSiesaUNOEE",
            cred_user="sa",
            database=self.TEST_DB,
        )
