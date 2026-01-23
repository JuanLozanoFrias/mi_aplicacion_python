from __future__ import annotations

import copy
import unicodedata
from typing import Any, Dict, List

import pandas as pd

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QFrame, QLabel, QDialog,
    QPushButton, QScrollArea, QTableWidget, QTableWidgetItem, QHeaderView,
    QListWidget, QListWidgetItem, QLineEdit, QComboBox, QSpinBox,
    QDoubleSpinBox, QCheckBox, QMessageBox, QSizePolicy, QAbstractScrollArea,
    QTreeWidget, QTreeWidgetItem, QAbstractItemView
)

from logic.cotizador import (
    get_base_dir,
    load_catalog,
    load_rules,
    load_clients,
    load_project_seed,
    normalize_project,
    build_cart_from_seed,
    calculate_totals,
    format_cop,
    format_usd,
    export_pdf,
    save_project_draft,
)


def _card() -> QFrame:
    f = QFrame()
    f.setStyleSheet("QFrame{background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;}")
    return f


def _label(text: str) -> QLabel:
    lab = QLabel(text)
    lab.setStyleSheet("color:#0f172a;font-weight:700;")
    return lab


class CotizadorPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.base_dir = get_base_dir()
        self.catalog = load_catalog(self.base_dir)
        self.rules = load_rules(self.base_dir)
        self.clients = load_clients(self.base_dir)
        project_seed, cart_seed = load_project_seed(self.base_dir)
        self.project_seed = project_seed
        self.cart_seed = cart_seed
        self.project = normalize_project(project_seed, self.rules)
        self.cart = []
        self.products = self.catalog.get("_products_by_id", {}) if isinstance(self.catalog, dict) else {}
        self._product_name_index = self._build_product_name_index()
        self._catalog_template = self._pick_template_product()
        self._equipment_measures: Dict[str, List[str]] = {}
        self._equipment_catalog = self._load_equipment_catalog()
        self._contact_emails = {
            "JUAN LOZANO": "jlozano@weston.com.co",
            "NICOLAS LEON": "nileonpa@weston.com.co",
        }
        self._option_costs = {
            "ESTIBA": 480000,
            "TRANSPORTE": 1350000,
            "INSTALACION": 2100000,
        }
        self._clients_by_id = self._index_clients()
        self._selected_product_id: str | None = None
        self._selected_product: Dict[str, Any] | None = None
        self._selected_equipment_name: str = ""
        self._current_modules: List[float] = []
        self._modulation_manual = False
        self._table_updating = False
        self._totals: Dict[str, Any] = {}
        self._build_ui()
        self._load_project_into_form()
        self._refresh_cart()
        if self._equipment_catalog:
            self._select_equipment_name(self._equipment_catalog[0])
        elif self.products:
            pid = next(iter(self.products.keys()))
            self._select_product(pid, self.products[pid])

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        root.addWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)
        cl = QVBoxLayout(container)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(12)

        hero = QFrame()
        hero.setObjectName("CotizadorHero")
        hero.setStyleSheet("""
QFrame#CotizadorHero{
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #0f62fe, stop:1 #22d3ee);
    border-radius:18px;
}
#HeroTitle{color:#ffffff;font-size:28px;font-weight:800;}
#HeroSub{color:#dbeafe;font-size:12px;font-weight:600;}
#HeroBadge{background:rgba(255,255,255,0.16);color:#e2e8f5;border-radius:10px;padding:4px 10px;font-weight:700;}
        """)
        hl = QHBoxLayout(hero)
        hl.setContentsMargins(24, 20, 24, 20)
        hl.setSpacing(16)
        title_col = QVBoxLayout()
        title = QLabel("Cotizador Weston")
        title.setObjectName("HeroTitle")
        sub = QLabel("DEMO GERENCIA - COTIZADOR PREMIUM")
        sub.setObjectName("HeroSub")
        badge = QLabel("DEMO")
        badge.setObjectName("HeroBadge")
        badge.setAlignment(Qt.AlignCenter)
        badge.setFixedWidth(70)
        title_row = QHBoxLayout()
        title_row.addWidget(title)
        title_row.addWidget(badge, alignment=Qt.AlignLeft)
        title_row.addStretch(1)
        title_col.addLayout(title_row)
        title_col.addWidget(sub)
        hl.addLayout(title_col, 1)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_cart = QPushButton("CARRITO")
        self.btn_cart.setProperty("variant", "ghost")
        self.btn_cart.setIcon(self._cart_icon())
        self.btn_cart.setIconSize(QSize(16, 16))
        self.btn_new = QPushButton("NUEVO")
        self.btn_save = QPushButton("GUARDAR BORRADOR")
        self.btn_export = QPushButton("EXPORTAR PDF")
        for b in (self.btn_cart, self.btn_new, self.btn_save):
            b.setProperty("variant", "ghost")
        btn_row.addWidget(self.btn_cart)
        btn_row.addWidget(self.btn_new)
        btn_row.addWidget(self.btn_save)
        btn_row.addWidget(self.btn_export)
        hl.addLayout(btn_row)

        self.btn_new.clicked.connect(self._on_new_project)
        self.btn_save.clicked.connect(self._on_save_draft)
        self.btn_export.clicked.connect(self._on_export_pdf)
        self.btn_cart.clicked.connect(self._on_show_cart)

        cl.addWidget(hero)

        stepper = _card()
        sl = QHBoxLayout(stepper)
        sl.setContentsMargins(16, 12, 16, 12)
        sl.setSpacing(10)
        steps = ["Proyecto", "Equipos", "Resumen", "Exportar"]
        self.btn_step_resumen = None
        for idx, name in enumerate(steps, start=1):
            step = QFrame()
            step.setStyleSheet("QFrame{background:transparent;}")
            step_l = QHBoxLayout(step)
            step_l.setContentsMargins(0, 0, 0, 0)
            step_l.setSpacing(8)
            bubble = QLabel(str(idx))
            bubble.setFixedSize(26, 26)
            bubble.setAlignment(Qt.AlignCenter)
            bubble.setStyleSheet(
                "background:#0f62fe;color:#ffffff;border-radius:13px;font-weight:700;"
            )
            if name.lower() == "resumen":
                self.btn_step_resumen = QPushButton(name.upper())
                self.btn_step_resumen.setCursor(Qt.PointingHandCursor)
                self.btn_step_resumen.setFlat(True)
                self.btn_step_resumen.setStyleSheet(
                    "QPushButton{color:#0f172a;font-weight:700;background:transparent;border:none;padding:0;}"
                    "QPushButton:hover{color:#0f62fe;}"
                )
                self.btn_step_resumen.clicked.connect(self._on_show_summary)
                label = self.btn_step_resumen
            else:
                label = QLabel(name.upper())
                label.setStyleSheet("color:#0f172a;font-weight:700;")
            step_l.addWidget(bubble)
            step_l.addWidget(label)
            sl.addWidget(step)
            if idx < len(steps):
                bar = QFrame()
                bar.setFixedHeight(2)
                bar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                bar.setStyleSheet("background:#cbd5e1;border-radius:1px;")
                sl.addWidget(bar)
        cl.addWidget(stepper)

        kpis = QHBoxLayout()
        self.kpi_total_cop = self._make_kpi("TOTAL COP")
        self.kpi_total_usd = self._make_kpi("TOTAL USD")
        self.kpi_margin = self._make_kpi("MARGEN DEMO")
        for card in (self.kpi_total_cop["card"], self.kpi_total_usd["card"], self.kpi_margin["card"]):
            kpis.addWidget(card, 1)
        cl.addLayout(kpis)

        body = QHBoxLayout()
        body.setSpacing(12)
        cl.addLayout(body)

        # Left column: catalog
        left = _card()
        left.setMinimumWidth(260)
        left_l = QVBoxLayout(left)
        left_l.setContentsMargins(16, 16, 16, 16)
        left_l.setSpacing(12)
        left_l.addWidget(_label("CATALOGO EQUIPOS"))
        self.ed_catalog_search = QLineEdit()
        self.ed_catalog_search.setPlaceholderText("Buscar equipo...")
        left_l.addWidget(self.ed_catalog_search)
        self.lbl_catalog_count = QLabel("")
        self.lbl_catalog_count.setStyleSheet("color:#64748b;font-size:11px;")
        left_l.addWidget(self.lbl_catalog_count)
        self.catalog_list = QListWidget()
        self.catalog_list.setStyleSheet(
            "QListWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QListWidget::item{padding:6px 8px;}"
            "QListWidget::item:selected{background:#e8f2ff;color:#0f172a;}"
        )
        left_l.addWidget(self.catalog_list, 1)
        self._populate_catalog_list(self._equipment_catalog)
        body.addWidget(left)

        # Center column: project + configurator
        center = QFrame()
        center.setStyleSheet("QFrame{background:transparent;}")
        center_l = QVBoxLayout(center)
        center_l.setContentsMargins(0, 0, 0, 0)
        center_l.setSpacing(12)

        proj_card = _card()
        pc = QVBoxLayout(proj_card)
        pc.setContentsMargins(16, 16, 16, 16)
        pc.setSpacing(10)
        pc.addWidget(_label("PROYECTO"))
        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(8)

        self.ed_project = QLineEdit()
        self.cb_client = QComboBox()
        self.cb_contact = QComboBox()
        self.ed_email = QLineEdit()
        self.ed_city = QLineEdit()
        self.ed_trm = QDoubleSpinBox()
        self.ed_trm.setMaximum(9999999)
        self.ed_trm.setDecimals(0)
        self.ed_trm.setSingleStep(10)
        self.ed_discount = QDoubleSpinBox()
        self.ed_discount.setMaximum(25.0)
        self.ed_discount.setDecimals(1)
        self.ed_discount.setSingleStep(0.5)
        self.ed_notes = QLineEdit()

        self.cb_client.setMinimumWidth(220)
        self._load_client_combo()

        self.ed_email.setReadOnly(True)
        self.cb_contact.addItems([""] + list(self._contact_emails.keys()))
        self.ed_project.textChanged.connect(lambda t: self._force_upper(self.ed_project, t))
        self.ed_city.textChanged.connect(lambda t: self._force_upper(self.ed_city, t))
        self.ed_notes.textChanged.connect(lambda t: self._force_upper(self.ed_notes, t))

        grid.addWidget(_label("NOMBRE PROYECTO"), 0, 0)
        grid.addWidget(self.ed_project, 0, 1)
        grid.addWidget(_label("CLIENTE"), 0, 2)
        grid.addWidget(self.cb_client, 0, 3)
        grid.addWidget(_label("CONTACTO"), 1, 0)
        grid.addWidget(self.cb_contact, 1, 1)
        grid.addWidget(_label("EMAIL"), 1, 2)
        grid.addWidget(self.ed_email, 1, 3)
        grid.addWidget(_label("CIUDAD"), 2, 0)
        grid.addWidget(self.ed_city, 2, 1)
        grid.addWidget(_label("TRM"), 2, 2)
        grid.addWidget(self.ed_trm, 2, 3)
        grid.addWidget(_label("DESCUENTO %"), 3, 0)
        grid.addWidget(self.ed_discount, 3, 1)
        grid.addWidget(_label("NOTAS"), 3, 2)
        grid.addWidget(self.ed_notes, 3, 3)
        pc.addLayout(grid)
        center_l.addWidget(proj_card)

        conf_card = _card()
        cc = QVBoxLayout(conf_card)
        cc.setContentsMargins(16, 16, 16, 16)
        cc.setSpacing(10)
        top_row = QHBoxLayout()
        top_row.addWidget(_label("CONFIGURADOR"))
        top_row.addStretch(1)
        self.btn_add = QPushButton("AGREGAR AL CARRITO")
        top_row.addWidget(self.btn_add)
        cc.addLayout(top_row)

        info_row = QHBoxLayout()
        self.lbl_product_name = QLabel("--")
        self.lbl_product_name.setStyleSheet("font-size:16px;font-weight:800;color:#0f172a;")
        self.lbl_product_lead = QLabel("")
        self.lbl_product_lead.setStyleSheet("color:#0f172a;font-weight:700;")
        info_row.addWidget(self.lbl_product_name)
        info_row.addStretch(1)
        info_row.addWidget(self.lbl_product_lead)
        cc.addLayout(info_row)

        metrics_row = QHBoxLayout()
        metrics_row.addWidget(_label("CANTIDAD EQUIPOS"))
        self.spin_qty = QSpinBox()
        self.spin_qty.setMinimum(0)
        self.spin_qty.setMaximum(50)
        self.spin_qty.setSpecialValueText("--")
        metrics_row.addWidget(self.spin_qty)
        metrics_row.addSpacing(12)
        metrics_row.addWidget(_label("DISTANCIA (m)"))
        self.spin_distance = QDoubleSpinBox()
        self.spin_distance.setRange(0, 9999)
        self.spin_distance.setDecimals(2)
        self.spin_distance.setSingleStep(0.1)
        self.spin_distance.setSpecialValueText("--")
        metrics_row.addWidget(self.spin_distance)
        metrics_row.addSpacing(12)
        metrics_row.addWidget(_label("MODULACION"))
        self.ed_modulation = QLineEdit()
        self.ed_modulation.setPlaceholderText("3.75 + 3.75 + 2.5")
        self.ed_modulation.setStyleSheet("color:#0f172a;font-weight:600;")
        metrics_row.addWidget(self.ed_modulation, 1)
        cc.addLayout(metrics_row)

        self.config_table = QTableWidget(0, 4)
        self.config_table.setHorizontalHeaderLabels(["DESCRIPCION", "PRECIO", "CANTIDAD", "SUBTOTAL"])
        self.config_table.verticalHeader().setVisible(False)
        self.config_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.config_table.setAlternatingRowColors(True)
        self.config_table.setStyleSheet(
            "QTableWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QHeaderView::section{background:#f8fafc;border:1px solid #e2e8f5;padding:6px;font-weight:700;}"
        )
        self.config_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.config_table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        self.config_table.itemChanged.connect(self._on_table_item_changed)
        cc.addWidget(self.config_table)

        opts = QHBoxLayout()
        self.chk_estiba = QCheckBox("ESTIBA +3%")
        self.chk_transporte = QCheckBox("TRANSPORTE")
        self.chk_instalacion = QCheckBox("INSTALACION")
        for chk in (self.chk_estiba, self.chk_transporte, self.chk_instalacion):
            chk.setStyleSheet("color:#0f172a;font-weight:700;")
        self.lbl_estiba_cost = QLabel("--")
        self.lbl_transporte_cost = QLabel("--")
        self.lbl_instalacion_cost = QLabel("--")
        for lab in (self.lbl_estiba_cost, self.lbl_transporte_cost, self.lbl_instalacion_cost):
            lab.setStyleSheet("color:#0f172a;font-weight:700;")
        opts.addWidget(self.chk_estiba)
        opts.addWidget(self.lbl_estiba_cost)
        opts.addSpacing(10)
        opts.addWidget(self.chk_transporte)
        opts.addWidget(self.lbl_transporte_cost)
        opts.addSpacing(10)
        opts.addWidget(self.chk_instalacion)
        opts.addWidget(self.lbl_instalacion_cost)
        opts.addStretch(1)
        cc.addLayout(opts)

        center_l.addWidget(conf_card)
        center_l.addStretch(1)
        body.addWidget(center, 1)

        self.btn_add.clicked.connect(self._on_add_to_cart)
        self.spin_qty.valueChanged.connect(self._on_qty_changed)
        self.spin_distance.valueChanged.connect(self._on_distance_changed)
        self.ed_modulation.editingFinished.connect(self._on_modulation_edited)
        self.chk_estiba.stateChanged.connect(self._on_project_options_changed)
        self.chk_transporte.stateChanged.connect(self._on_project_options_changed)
        self.chk_instalacion.stateChanged.connect(self._on_project_options_changed)
        self.cb_client.currentIndexChanged.connect(self._on_client_changed)
        self.cb_contact.currentIndexChanged.connect(self._on_contact_changed)
        self.ed_trm.valueChanged.connect(self._on_project_options_changed)
        self.ed_discount.valueChanged.connect(self._on_project_options_changed)
        self.ed_project.editingFinished.connect(self._on_project_options_changed)
        self.ed_city.editingFinished.connect(self._on_project_options_changed)
        self.ed_notes.editingFinished.connect(self._on_project_options_changed)
        self.ed_catalog_search.textChanged.connect(self._filter_catalog_list)
        self.catalog_list.itemSelectionChanged.connect(self._on_catalog_selected)
        self._build_cart_dialog()
        self._build_summary_dialog()

    def _make_kpi(self, title: str) -> Dict[str, Any]:
        card = _card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)
        t = QLabel(title)
        t.setStyleSheet("font-size:11px;font-weight:700;color:#475569;")
        v = QLabel("0")
        v.setStyleSheet("font-size:20px;font-weight:800;color:#0f172a;")
        layout.addWidget(t)
        layout.addWidget(v)
        return {"card": card, "value": v}

    def _index_clients(self) -> Dict[str, Dict[str, Any]]:
        clients = self.clients.get("clients", []) if isinstance(self.clients, dict) else []
        return {str(c.get("client_id", "")): c for c in clients if isinstance(c, dict)}

    def _load_client_combo(self) -> None:
        self.cb_client.blockSignals(True)
        self.cb_client.clear()
        for cid, data in self._clients_by_id.items():
            name = data.get("name", "")
            self.cb_client.addItem(f"{cid} - {name}".upper(), userData=cid)
        self.cb_client.blockSignals(False)

    def _load_project_into_form(self) -> None:
        self.project = normalize_project(self.project, self.rules)
        self.ed_project.setText(str(self.project.get("project_name", "")))
        contact = str(self.project.get("contact", "")).strip().upper()
        if contact in self._contact_emails:
            idx = self.cb_contact.findText(contact, Qt.MatchFixedString)
            if idx >= 0:
                self.cb_contact.setCurrentIndex(idx)
        else:
            self.cb_contact.setCurrentIndex(1 if self.cb_contact.count() > 1 else 0)
        self._apply_contact_email(self.cb_contact.currentText())
        self.ed_city.setText(str(self.project.get("city", "")))
        self.ed_trm.setValue(float(self.project.get("trm", 0) or 0))
        self.ed_discount.setValue(float(self.project.get("discount_pct", 0) or 0))
        self.ed_notes.setText(str(self.project.get("notes", "")))
        self.chk_estiba.setChecked(bool(self.project.get("estiba", False)))
        options = self.project.get("options", {}) if isinstance(self.project.get("options"), dict) else {}
        self.chk_transporte.setChecked(bool(options.get("transporte", False)))
        self.chk_instalacion.setChecked(bool(options.get("instalacion", False)))
        cid = str(self.project.get("client_id", ""))
        if cid:
            idx = self.cb_client.findData(cid)
            if idx >= 0:
                self.cb_client.setCurrentIndex(idx)
        if not cid and self.cb_client.count() > 0:
            self.cb_client.setCurrentIndex(0)
            cid = str(self.cb_client.currentData() or "")
        self._apply_client_details(cid)
        self._refresh_totals()

    def _apply_client_details(self, client_id: str) -> None:
        client = self._clients_by_id.get(client_id)
        if not client:
            return
        if not self.ed_city.text().strip():
            self.ed_city.setText(str(client.get("city", "")))
        self.project["client_name"] = client.get("name", "")

    def _apply_contact_email(self, contact_name: str) -> None:
        key = str(contact_name or "").strip().upper()
        email = self._contact_emails.get(key, "")
        self.ed_email.setText(email.upper())

    @staticmethod
    def _norm_text(text: str) -> str:
        raw = str(text or "").strip().upper()
        raw = unicodedata.normalize("NFKD", raw).encode("ascii", "ignore").decode("ascii")
        return "".join(ch for ch in raw if ch.isalnum() or ch in " _-").strip()

    def _build_product_name_index(self) -> Dict[str, str]:
        index: Dict[str, str] = {}
        for pid, prod in self.products.items():
            for key in (prod.get("name", ""), prod.get("sku", ""), pid):
                norm = self._norm_text(key)
                if norm and norm not in index:
                    index[norm] = pid
        return index

    def _pick_template_product(self) -> Dict[str, Any]:
        if not self.products:
            return {}
        pid = next(iter(self.products.keys()))
        return copy.deepcopy(self.products[pid])

    def _load_equipment_catalog(self) -> List[str]:
        base_dir = self.base_dir / "data"
        candidates = [
            base_dir / "CUADRO DE CARGAS -INDIVIDUAL JUAN LOZANO.xlsx",
            base_dir / "CUADRO DE CARGAS -INDIVIDUAL JUAN LOZANO.xlsm",
        ]
        book = next((c for c in candidates if c.exists()), None)
        equipos: List[str] = []
        self._equipment_measures = {}
        if book is not None:
            try:
                xls = pd.ExcelFile(book)
                names = list(xls.sheet_names)
                sheet = None
                if "MUEBLES" in names:
                    sheet = "MUEBLES"
                if sheet is None:
                    for name in names:
                        if "MUEB" in str(name).upper() or "EQUIP" in str(name).upper():
                            sheet = name
                            break
                if sheet is None:
                    raise ValueError("No encontre hoja de muebles/equipos.")
                df = pd.read_excel(book, sheet_name=sheet)
                df.columns = [str(c).strip() for c in df.columns]
                col = next((c for c in df.columns if self._norm_text(c) == "NOMBRE"), None)
                if col is None:
                    col = next((c for c in df.columns if self._norm_text(c) == "EQUIPO"), None)
                if col is None and df.columns.size:
                    col = df.columns[0]
                dim_col = next((c for c in df.columns if self._norm_text(c) == "DIMENSION"), None)
                if col is not None:
                    for _, row in df.iterrows():
                        eq = str(row.get(col, "")).strip()
                        dim = str(row.get(dim_col, "")).strip() if dim_col is not None else ""
                        if eq:
                            equipos.append(eq)
                            if dim:
                                key = self._norm_text(eq)
                                self._equipment_measures.setdefault(key, []).append(dim)
                    for k, vals in list(self._equipment_measures.items()):
                        self._equipment_measures[k] = sorted(set(vals))
                    equipos = sorted(set(equipos))
            except Exception:
                equipos = []
        fallback = {str(p.get("name", "")).strip() for p in self.products.values() if str(p.get("name", "")).strip()}
        if equipos:
            equipos = sorted(set(equipos) | fallback)
        else:
            equipos = sorted(fallback)
        equipos = [e for e in equipos if not self._exclude_equipment(e)]
        return equipos

    def _populate_catalog_list(self, items: List[str]) -> None:
        self.catalog_list.blockSignals(True)
        self.catalog_list.clear()
        for name in items:
            self.catalog_list.addItem(QListWidgetItem(name))
        self.catalog_list.blockSignals(False)
        self.lbl_catalog_count.setText(f"{len(items)} equipos")
        if items and not self.catalog_list.currentItem():
            self.catalog_list.setCurrentRow(0)

    def _filter_catalog_list(self, text: str) -> None:
        query = self._norm_text(text)
        if not query:
            self._populate_catalog_list(self._equipment_catalog)
            return
        filtered = [n for n in self._equipment_catalog if query in self._norm_text(n)]
        self._populate_catalog_list(filtered)

    def _on_catalog_selected(self) -> None:
        item = self.catalog_list.currentItem()
        if not item:
            return
        self._select_equipment_name(item.text())

    def _select_equipment_name(self, name: str) -> None:
        norm = self._norm_text(name)
        pid = self._product_name_index.get(norm)
        self._reset_configurator_state()
        if pid and pid in self.products:
            self._select_product(pid, self.products[pid])
            return
        if not self._catalog_template:
            return
        product = self._build_dynamic_product(name, self._catalog_template)
        self._select_product(product["product_id"], product)

    def _build_dynamic_product(self, name: str, template: Dict[str, Any]) -> Dict[str, Any]:
        product = copy.deepcopy(template)
        product_id = self._norm_text(name).replace(" ", "_")
        product["product_id"] = f"EQUIPO_{product_id}" if product_id else "EQUIPO_DEMO"
        product["name"] = name
        product["sku"] = product.get("sku", "") or product["product_id"]
        product["lead_time_days"] = int(product.get("lead_time_days", 15) or 15)
        return product

    def _select_product(self, product_id: str, product: Dict[str, Any]) -> None:
        if not product:
            return
        self._selected_product_id = product_id
        self._selected_product = product
        self._selected_equipment_name = str(product.get("name", ""))
        self.lbl_product_name.setText(str(product.get("name", "--")))
        lead = product.get("lead_time_days", 0)
        self.lbl_product_lead.setText(f"Lead time: {lead} dias")
        self.spin_qty.setValue(1)
        self._update_modulation()
        self._update_modulation()

    def _cart_icon(self) -> QIcon:
        size = 16
        pix = QPixmap(size, size)
        pix.fill(Qt.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.Antialiasing, True)
        pen = QColor("#e2e8f5")
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawRect(3, 4, 10, 6)
        painter.drawLine(3, 4, 1, 2)
        painter.setBrush(QColor("#e2e8f5"))
        painter.drawEllipse(4, 12, 3, 3)
        painter.drawEllipse(10, 12, 3, 3)
        painter.end()
        return QIcon(pix)

    def _build_cart_dialog(self) -> None:
        self.cart_dialog = QDialog(self)
        self.cart_dialog.setWindowTitle("Carrito")
        self.cart_dialog.setMinimumWidth(360)
        self.cart_dialog.setModal(False)
        dlg_layout = QVBoxLayout(self.cart_dialog)
        dlg_layout.setContentsMargins(16, 16, 16, 16)
        dlg_layout.setSpacing(10)

        title = QLabel("CARRITO")
        title.setStyleSheet("font-size:16px;font-weight:800;color:#0f172a;")
        dlg_layout.addWidget(title)

        self.cart_list = QListWidget()
        self.cart_list.setStyleSheet(
            "QListWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
        )
        dlg_layout.addWidget(self.cart_list, 1)

        self.btn_remove = QPushButton("Quitar seleccionado")
        self.btn_remove.setStyleSheet(
            "QPushButton{background:#fee2e2;color:#991b1b;font-weight:700;border-radius:8px;padding:6px 10px;}"
            "QPushButton:hover{background:#fecaca;}"
        )
        dlg_layout.addWidget(self.btn_remove)

        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet("background:#e2e8f5;")
        dlg_layout.addWidget(sep)

        summary = QGridLayout()
        summary.setHorizontalSpacing(8)
        summary.setVerticalSpacing(6)
        self.lbl_subtotal = QLabel("0")
        self.lbl_recargo = QLabel("0")
        self.lbl_descuento = QLabel("0")
        self.lbl_iva = QLabel("0")
        self.lbl_total = QLabel("0")
        labels = [
            ("Subtotal", self.lbl_subtotal),
            ("Recargo estiba", self.lbl_recargo),
            ("Descuento", self.lbl_descuento),
            ("IVA", self.lbl_iva),
            ("Total", self.lbl_total),
        ]
        for r, (name, lab) in enumerate(labels):
            lab.setAlignment(Qt.AlignRight)
            if name == "Total":
                lab.setStyleSheet("font-weight:800;color:#0f172a;")
            summary.addWidget(QLabel(name), r, 0)
            summary.addWidget(lab, r, 1)
        dlg_layout.addLayout(summary)

        margin_wrap = QHBoxLayout()
        margin_wrap.addWidget(_label("Semaforo margen"))
        margin_wrap.addStretch(1)
        self.margin_red = QLabel()
        self.margin_yellow = QLabel()
        self.margin_green = QLabel()
        for dot in (self.margin_red, self.margin_yellow, self.margin_green):
            dot.setFixedSize(12, 12)
            dot.setStyleSheet("background:#e5e7eb;border-radius:6px;")
            margin_wrap.addWidget(dot)
        dlg_layout.addLayout(margin_wrap)

        close_row = QHBoxLayout()
        close_row.addStretch(1)
        btn_close = QPushButton("Cerrar")
        btn_close.clicked.connect(self.cart_dialog.close)
        close_row.addWidget(btn_close)
        dlg_layout.addLayout(close_row)

        self.btn_remove.clicked.connect(self._on_remove_item)

    def _build_summary_dialog(self) -> None:
        self.summary_dialog = QDialog(self)
        self.summary_dialog.setWindowTitle("Resumen")
        self.summary_dialog.setWindowFlags(
            Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMaximizeButtonHint
        )
        self.summary_dialog.setMinimumSize(960, 640)
        self.summary_dialog.resize(1100, 720)
        self.summary_dialog.setModal(False)
        dlg_layout = QVBoxLayout(self.summary_dialog)
        dlg_layout.setContentsMargins(16, 16, 16, 16)
        dlg_layout.setSpacing(10)

        title = QLabel("RESUMEN DEL CARRITO")
        title.setStyleSheet("font-size:16px;font-weight:800;color:#0f172a;")
        dlg_layout.addWidget(title)

        self.summary_tree = QTreeWidget()
        self.summary_tree.setColumnCount(4)
        self.summary_tree.setHeaderLabels(["ITEM", "CANTIDAD", "PRECIO", "SUBTOTAL"])
        self.summary_tree.setRootIsDecorated(True)
        self.summary_tree.setAlternatingRowColors(True)
        self.summary_tree.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.summary_tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        for col in (1, 2, 3):
            self.summary_tree.header().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        self.summary_tree.setStyleSheet(
            "QTreeWidget{background:#ffffff;border:1px solid #e2e8f5;border-radius:8px;}"
            "QHeaderView::section{background:#f8fafc;border:1px solid #e2e8f5;padding:6px;font-weight:700;}"
        )
        dlg_layout.addWidget(self.summary_tree, 1)

        sep = QFrame()
        sep.setFixedHeight(1)
        sep.setStyleSheet("background:#e2e8f5;")
        dlg_layout.addWidget(sep)

        summary = QGridLayout()
        summary.setHorizontalSpacing(8)
        summary.setVerticalSpacing(6)
        self.summary_subtotal = QLabel("0")
        self.summary_recargo = QLabel("0")
        self.summary_descuento = QLabel("0")
        self.summary_iva = QLabel("0")
        self.summary_total = QLabel("0")
        labels = [
            ("Subtotal", self.summary_subtotal),
            ("Recargo estiba", self.summary_recargo),
            ("Descuento", self.summary_descuento),
            ("IVA", self.summary_iva),
            ("Total", self.summary_total),
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
        btn_close.clicked.connect(self.summary_dialog.close)
        close_row.addWidget(btn_close)
        dlg_layout.addLayout(close_row)

    def _on_show_cart(self) -> None:
        if not hasattr(self, "cart_dialog"):
            self._build_cart_dialog()
        self._refresh_cart()
        self.cart_dialog.show()
        self.cart_dialog.raise_()
        self.cart_dialog.activateWindow()

    def _on_show_summary(self) -> None:
        if not hasattr(self, "summary_dialog"):
            self._build_summary_dialog()
        self._refresh_summary()
        self.summary_dialog.show()
        self.summary_dialog.raise_()
        self.summary_dialog.activateWindow()

    def _update_cart_button(self) -> None:
        count = len(self.cart)
        self.btn_cart.setText(f"CARRITO ({count})" if count else "CARRITO")

    def _populate_config_table(self, product: Dict[str, Any]) -> None:
        self._table_updating = True
        self.config_table.setRowCount(0)
        rows: List[Dict[str, Any]] = []
        is_vitrina = self._is_vitrina()
        for base in product.get("base_modules", []) or []:
            rows.append({
                "desc": str(base.get("desc", "")).upper(),
                "unit_price": base.get("unit_price", 0),
                "qty": self.spin_qty.value(),
                "kind": "base",
                "editable_qty": True,
            })
        for addon in product.get("optional_addons", []) or []:
            desc = str(addon.get("desc", "")).upper()
            upper = self._norm_text(desc)
            if is_vitrina and any(k in upper for k in ("TESTERO", "PUERTA", "BUMPER")):
                continue
            rows.append({
                "desc": desc,
                "unit_price": addon.get("unit_price", 0),
                "qty": 0,
                "kind": "addon",
                "editable_qty": True,
            })
        if is_vitrina:
            rows.extend(self._vitrina_parts(max(1, self.spin_qty.value())))
        rows.extend(self._option_rows_from_state())
        self._render_config_rows(rows)
        self._table_updating = False
        self._auto_height_config_table()

    def _populate_config_table_for_modules(self, product: Dict[str, Any], modules: List[float]) -> None:
        self._table_updating = True
        self.config_table.setRowCount(0)
        rows: List[Dict[str, Any]] = []
        is_vitrina = self._is_vitrina()

        base_price, base_len = self._get_base_module_price(product)
        for length in modules:
            unit_price = self._scale_module_price(base_price, base_len, length)
            rows.append({
                "desc": f"MODULO BASE {length:g} M",
                "unit_price": unit_price,
                "qty": 1,
                "kind": "base",
                "editable_qty": True,
            })

        module_count = len(modules)
        testero_pairs = 0 if is_vitrina else (max(1, module_count - 1) if module_count > 0 else 0)
        illumination_price = self._find_unit_price(product, ["ILUMINACION"])
        control_price = self._find_unit_price(product, ["CONTROL"])

        if illumination_price > 0:
            rows.append({
                "desc": "ILUMINACION LED",
                "unit_price": illumination_price,
                "qty": module_count if is_vitrina else (module_count + testero_pairs),
                "kind": "base",
                "editable_qty": True,
            })
        if control_price > 0:
            rows.append({
                "desc": "CONTROL BASICO",
                "unit_price": control_price,
                "qty": max(1, module_count // 2) if is_vitrina and module_count else (max(1, testero_pairs) if module_count else 0),
                "kind": "base",
                "editable_qty": True,
            })

        for addon in product.get("optional_addons", []) or []:
            desc = str(addon.get("desc", "")).strip()
            if not desc:
                continue
            upper = self._norm_text(desc)
            qty = 0
            editable = True
            if is_vitrina and any(k in upper for k in ("TESTERO", "PUERTA", "BUMPER")):
                continue
            if "TESTERO" in upper:
                qty = testero_pairs
                editable = True
            rows.append({
                "desc": desc.upper(),
                "unit_price": addon.get("unit_price", 0),
                "qty": qty,
                "kind": "addon",
                "editable_qty": editable,
            })
        if is_vitrina:
            rows.extend(self._vitrina_parts(module_count))
        rows.extend(self._option_rows_from_state())

        self._render_config_rows(rows)
        self._table_updating = False
        self._auto_height_config_table()

    def _parse_number(self, text: str) -> float:
        raw = str(text or "").replace("COP", "").replace("$", "").replace(" ", "")
        raw = raw.replace(".", "")
        raw = raw.replace(",", ".")
        cleaned = "".join(ch for ch in raw if ch.isdigit() or ch in ".-")
        try:
            return float(cleaned) if cleaned else 0.0
        except Exception:
            return 0.0

    @staticmethod
    def _safe_float(value: Any) -> float:
        if isinstance(value, (int, float)):
            return float(value)
        raw = str(value or "").strip()
        if not raw:
            return 0.0
        raw = raw.replace("COP", "").replace("$", "").replace(" ", "")
        if "." in raw and "," in raw:
            raw = raw.replace(".", "").replace(",", ".")
        elif raw.count(".") > 1:
            raw = raw.replace(".", "")
        elif "," in raw:
            raw = raw.replace(",", ".")
        cleaned = "".join(ch for ch in raw if ch.isdigit() or ch in ".-")
        try:
            return float(cleaned) if cleaned else 0.0
        except Exception:
            return 0.0

    @staticmethod
    def _parse_qty(text: str) -> float:
        raw = str(text or "").replace("COP", "").replace("$", "").replace(" ", "")
        if "." in raw and "," in raw:
            raw = raw.replace(".", "").replace(",", ".")
        elif "," in raw:
            raw = raw.replace(",", ".")
        cleaned = "".join(ch for ch in raw if ch.isdigit() or ch in ".-")
        try:
            return float(cleaned) if cleaned else 0.0
        except Exception:
            return 0.0

    def _update_row_subtotal(self, row: int) -> None:
        price_item = self.config_table.item(row, 1)
        qty_item = self.config_table.item(row, 2)
        if not price_item or not qty_item:
            return
        price = self._parse_number(price_item.text())
        qty = self._parse_qty(qty_item.text())
        subtotal = price * qty
        sub_item = self.config_table.item(row, 3)
        if sub_item:
            sub_item.setText(self._format_money(subtotal))
            sub_item.setData(Qt.UserRole, subtotal)
        self._recalc_total_row()

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._table_updating:
            return
        if self._is_total_row(item.row()):
            return
        if item.column() not in (1, 2):
            return
        self._table_updating = True
        row = item.row()
        if item.column() == 1:
            val = self._parse_number(item.text())
            item.setText(self._format_money(val))
        if item.column() == 2:
            val = max(0.0, self._parse_qty(item.text()))
            item.setText(str(int(val)))
        self._update_row_subtotal(row)
        self._table_updating = False

    def _on_qty_changed(self, value: int) -> None:
        if self._table_updating:
            return
        if self._current_modules:
            return
        _ = value
        self._apply_base_qty()

    def _on_add_to_cart(self) -> None:
        if not self._selected_product_id or not self._selected_product:
            return
        product = self._selected_product
        qty_equipo = self.spin_qty.value()
        line_items: List[Dict[str, Any]] = []
        for row in range(self.config_table.rowCount()):
            if self._is_total_row(row):
                continue
            desc_item = self.config_table.item(row, 0)
            price_item = self.config_table.item(row, 1)
            qty_item = self.config_table.item(row, 2)
            if not (desc_item and price_item and qty_item):
                continue
            desc = desc_item.text().strip()
            price = self._parse_number(price_item.text())
            qty = self._parse_qty(qty_item.text())
            if qty <= 0:
                continue
            meta = desc_item.data(Qt.UserRole) or {}
            line_items.append({
                "desc": desc,
                "unit_price": price,
                "qty": qty,
                "kind": meta.get("kind", "base"),
            })
        if not line_items:
            self._warn("No hay items para agregar.")
            return
        self.cart.append({
            "product_id": self._selected_product_id,
            "product_name": product.get("name", self._selected_product_id),
            "qty": qty_equipo,
            "lead_time_days": int(product.get("lead_time_days", 0) or 0),
            "line_items": line_items,
        })
        self._refresh_cart()
        self._info("ITEM AGREGADO AL CARRITO.")
        self._clear_after_add()

    def _on_distance_changed(self, value: float) -> None:
        _ = value
        self._modulation_manual = False
        self._update_modulation()

    def _on_modulation_edited(self) -> None:
        modules = self._parse_modules_text(self.ed_modulation.text())
        if not modules:
            self._modulation_manual = False
            self._update_modulation()
            return
        self._apply_manual_modules(modules)

    def _on_remove_item(self) -> None:
        row = self.cart_list.currentRow()
        if row < 0 or row >= len(self.cart):
            return
        self.cart.pop(row)
        self._refresh_cart()

    def _refresh_cart(self) -> None:
        self.cart_list.blockSignals(True)
        self.cart_list.clear()
        for item in self.cart:
            name = item.get("product_name", item.get("product_id", "ITEM"))
            qty = int(item.get("qty", 1) or 1)
            total = 0.0
            for line in item.get("line_items", []) or []:
                total += self._safe_float(line.get("unit_price", 0)) * self._safe_float(line.get("qty", 0))
            label = f"{name} x{qty}  COP {format_cop(total)}"
            self.cart_list.addItem(QListWidgetItem(label))
        self.cart_list.blockSignals(False)
        self._update_cart_button()
        self._refresh_totals()

    def _refresh_summary(self) -> None:
        self.summary_tree.clear()
        for item in self.cart:
            name = str(item.get("product_name") or item.get("product_id") or "ITEM").upper()
            qty_equipo = int(item.get("qty", 1) or 1)
            product_total = 0.0
            for line in item.get("line_items", []) or []:
                product_total += self._safe_float(line.get("unit_price", 0)) * self._safe_float(line.get("qty", 0))
            parent = QTreeWidgetItem([
                f"{name} (X{qty_equipo})",
                str(qty_equipo),
                "",
                self._format_money(product_total),
            ])
            bold_font = QFont()
            bold_font.setBold(True)
            for col in range(4):
                parent.setFont(col, bold_font)
            parent.setTextAlignment(1, Qt.AlignRight | Qt.AlignVCenter)
            parent.setTextAlignment(3, Qt.AlignRight | Qt.AlignVCenter)
            self.summary_tree.addTopLevelItem(parent)
            for line in item.get("line_items", []) or []:
                desc = str(line.get("desc", "")).upper()
                qty = self._safe_float(line.get("qty", 0))
                unit = self._safe_float(line.get("unit_price", 0))
                subtotal = unit * qty
                child = QTreeWidgetItem([
                    desc,
                    str(int(qty)),
                    self._format_money(unit),
                    self._format_money(subtotal),
                ])
                child.setTextAlignment(1, Qt.AlignRight | Qt.AlignVCenter)
                child.setTextAlignment(2, Qt.AlignRight | Qt.AlignVCenter)
                child.setTextAlignment(3, Qt.AlignRight | Qt.AlignVCenter)
                parent.addChild(child)
        self.summary_tree.expandAll()
        self._sync_project_from_form()
        totals = calculate_totals(self.cart, self.project, self.rules)
        self.summary_subtotal.setText(format_cop(totals.get("subtotal_base", 0.0)))
        self.summary_recargo.setText(format_cop(totals.get("recargo_estiba", 0.0)))
        self.summary_descuento.setText(format_cop(totals.get("descuento", 0.0)))
        self.summary_iva.setText(format_cop(totals.get("iva", 0.0)))
        self.summary_total.setText(format_cop(totals.get("total_final_rounded", 0.0)))

    def _refresh_totals(self) -> None:
        self._sync_project_from_form()
        self._totals = calculate_totals(self.cart, self.project, self.rules)
        self.lbl_subtotal.setText(format_cop(self._totals.get("subtotal_base", 0.0)))
        self.lbl_recargo.setText(format_cop(self._totals.get("recargo_estiba", 0.0)))
        self.lbl_descuento.setText(format_cop(self._totals.get("descuento", 0.0)))
        self.lbl_iva.setText(format_cop(self._totals.get("iva", 0.0)))
        self.lbl_total.setText(format_cop(self._totals.get("total_final_rounded", 0.0)))
        self.kpi_total_cop["value"].setText(f"COP {format_cop(self._totals.get('total_final_rounded', 0.0))}")
        self.kpi_total_usd["value"].setText(f"USD {format_usd(self._totals.get('total_usd', 0.0))}")
        self.kpi_margin["value"].setText(f"{self._totals.get('margin_pct', 0.0):.1f}%")
        self._update_margin_lights()
        self._update_option_costs()

    def _update_margin_lights(self) -> None:
        margin_cfg = self.rules.get("margin_demo", {}) if isinstance(self.rules, dict) else {}
        target = float(margin_cfg.get("target_pct", 0) or 0)
        warn = float(margin_cfg.get("warn_below_pct", 0) or 0)
        val = float(self._totals.get("margin_pct", 0) or 0)
        red = "#ef4444"
        yellow = "#f59e0b"
        green = "#22c55e"
        off = "#e5e7eb"
        if val >= target:
            colors = (off, off, green)
        elif val >= warn:
            colors = (off, yellow, off)
        else:
            colors = (red, off, off)
        self.margin_red.setStyleSheet(f"background:{colors[0]};border-radius:6px;")
        self.margin_yellow.setStyleSheet(f"background:{colors[1]};border-radius:6px;")
        self.margin_green.setStyleSheet(f"background:{colors[2]};border-radius:6px;")

    def _sync_project_from_form(self) -> None:
        self.project["project_name"] = self.ed_project.text().strip()
        contact = self.cb_contact.currentText().strip()
        self.project["contact"] = contact
        self.project["email"] = self.ed_email.text().strip()
        self.project["city"] = self.ed_city.text().strip()
        self.project["trm"] = float(self.ed_trm.value())
        self.project["discount_pct"] = float(self.ed_discount.value())
        self.project["notes"] = self.ed_notes.text().strip()
        self.project["estiba"] = self.chk_estiba.isChecked()
        options = self.project.get("options", {}) if isinstance(self.project.get("options"), dict) else {}
        options["transporte"] = self.chk_transporte.isChecked()
        options["instalacion"] = self.chk_instalacion.isChecked()
        self.project["options"] = options
        cid = self.cb_client.currentData()
        if cid:
            self.project["client_id"] = cid
            self.project["client_name"] = self._clients_by_id.get(cid, {}).get("name", "")

    def _on_project_options_changed(self) -> None:
        self._sync_option_rows()
        self._refresh_totals()

    def _on_client_changed(self) -> None:
        cid = self.cb_client.currentData()
        if not cid:
            return
        self.project["client_id"] = cid
        self._apply_client_details(cid)
        self._refresh_totals()

    def _on_contact_changed(self) -> None:
        self._apply_contact_email(self.cb_contact.currentText())
        self._refresh_totals()

    def _auto_height_config_table(self) -> None:
        rows = max(self.config_table.rowCount(), 1)
        vh = self.config_table.verticalHeader()
        row_h = vh.defaultSectionSize()
        vlen = vh.length()
        if vlen <= 0:
            vlen = rows * row_h
        header_h = self.config_table.horizontalHeader().height() if self.config_table.horizontalHeader() else 0
        frame = self.config_table.frameWidth() * 2
        total_h = header_h + vlen + frame + 8
        self.config_table.setMinimumHeight(total_h)
        self.config_table.setMaximumHeight(total_h)

    def _apply_base_qty(self) -> None:
        if self._table_updating:
            return
        self._table_updating = True
        qty_base = self.spin_qty.value()
        for row in range(self.config_table.rowCount()):
            if self._is_total_row(row):
                continue
            desc_item = self.config_table.item(row, 0)
            if not desc_item:
                continue
            meta = desc_item.data(Qt.UserRole) or {}
            if meta.get("kind") == "base":
                qty_item = self.config_table.item(row, 2)
                if qty_item and qty_item.flags() & Qt.ItemIsEditable:
                    qty_item.setText(str(int(qty_base)))
                self._update_row_subtotal(row)
        self._table_updating = False

    def _get_base_module_price(self, product: Dict[str, Any]) -> tuple[float, float]:
        base_modules = product.get("base_modules", []) or []
        base_price = 0.0
        base_len = 2.5
        for base in base_modules:
            desc = str(base.get("desc", "")).upper()
            if "MODULO" in desc:
                base_price = float(base.get("unit_price", 0) or 0)
                base_len = self._extract_length(desc) or base_len
                return base_price, base_len
        if base_modules:
            base_price = float(base_modules[0].get("unit_price", 0) or 0)
            base_len = self._extract_length(str(base_modules[0].get("desc", ""))) or base_len
        return base_price, base_len

    def _find_unit_price(self, product: Dict[str, Any], keywords: List[str]) -> float:
        targets = [self._norm_text(k) for k in keywords]
        for base in product.get("base_modules", []) or []:
            desc = self._norm_text(base.get("desc", ""))
            if any(k in desc for k in targets):
                return float(base.get("unit_price", 0) or 0)
        return 0.0

    @staticmethod
    def _extract_length(text: str) -> float:
        raw = "".join(ch if ch.isdigit() or ch in ".,"
                      else " " for ch in str(text))
        parts = [p for p in raw.replace(",", ".").split() if p]
        for p in parts:
            try:
                return float(p)
            except Exception:
                continue
        return 0.0

    @staticmethod
    def _scale_module_price(base_price: float, base_len: float, length: float) -> float:
        if base_price <= 0 or base_len <= 0:
            return 0.0
        return round(base_price * (length / base_len) / 1000.0) * 1000.0

    @staticmethod
    def _parse_modules_text(text: str) -> List[float]:
        import re
        raw = str(text or "")
        parts = re.findall(r"\d+(?:[.,]\d+)?", raw)
        modules = []
        for p in parts:
            try:
                modules.append(float(p.replace(",", ".")))
            except Exception:
                continue
        return modules

    def _render_config_rows(self, rows: List[Dict[str, Any]]) -> None:
        total = 0.0
        self.config_table.setRowCount(len(rows) + 1)
        for r, row in enumerate(rows):
            desc = str(row.get("desc", "")).upper()
            desc_item = QTableWidgetItem(desc)
            desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)
            meta = {"kind": row.get("kind", "base")}
            if row.get("option_key"):
                meta["option_key"] = row.get("option_key")
            desc_item.setData(Qt.UserRole, meta)
            price_item = QTableWidgetItem(self._format_money(row.get("unit_price", 0)))
            price_item.setFlags(price_item.flags() & ~Qt.ItemIsEditable)
            qty_item = QTableWidgetItem(str(int(row.get("qty", 0))))
            if not row.get("editable_qty", False):
                qty_item.setFlags(qty_item.flags() & ~Qt.ItemIsEditable)
            sub_item = QTableWidgetItem("0")
            sub_item.setFlags(sub_item.flags() & ~Qt.ItemIsEditable)
            for item in (price_item, qty_item, sub_item):
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.config_table.setItem(r, 0, desc_item)
            self.config_table.setItem(r, 1, price_item)
            self.config_table.setItem(r, 2, qty_item)
            self.config_table.setItem(r, 3, sub_item)
            self._update_row_subtotal(r)
            total += self._parse_number(sub_item.text())

        total_row = len(rows)
        total_label = QTableWidgetItem("TOTAL")
        total_label.setFlags(total_label.flags() & ~Qt.ItemIsEditable)
        total_label.setData(Qt.UserRole, "total")
        total_label.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        total_value = QTableWidgetItem(self._format_money(total))
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
        self.config_table.setItem(total_row, 0, total_label)
        self.config_table.setItem(total_row, 1, blank_price)
        self.config_table.setItem(total_row, 2, blank_qty)
        self.config_table.setItem(total_row, 3, total_value)
        self._auto_height_config_table()

    def _split_length_generic(self, equipo: str, length_m: float) -> List[str]:
        dims = self._equipment_measures.get(self._norm_text(equipo), [])
        numeric_dims = []
        for d in dims:
            try:
                numeric_dims.append(float(str(d).replace(",", ".")))
            except Exception:
                continue
        numeric_dims = sorted(set(numeric_dims), reverse=True)
        if not numeric_dims or length_m <= 0:
            return []
        modules = []
        rem = length_m
        smallest = numeric_dims[-1]
        limit = 0
        while rem > 0.05 and limit < 200:
            pick = None
            for d in numeric_dims:
                if d <= rem + 0.05:
                    pick = d
                    break
            if pick is None:
                pick = smallest
            modules.append(str(pick))
            rem -= pick
            limit += 1
            if rem < smallest:
                break
        if not modules:
            modules.append(str(numeric_dims[0]))
        return modules

    def _update_modulation(self) -> None:
        if not self._selected_product:
            return
        equip = self._selected_equipment_name or self.lbl_product_name.text().strip()
        length_m = float(self.spin_distance.value())
        modules_raw = self._split_length_generic(equip, length_m)
        modules = []
        for m in modules_raw:
            try:
                modules.append(float(str(m).replace(",", ".")))
            except Exception:
                continue
        self._current_modules = modules
        if modules:
            mod_text = " + ".join(f"{m:g}" for m in modules)
            self.ed_modulation.setText(mod_text)
            self.spin_qty.blockSignals(True)
            self.spin_qty.setValue(len(modules))
            self.spin_qty.blockSignals(False)
            self.spin_qty.setEnabled(False)
            self._populate_config_table_for_modules(self._selected_product, modules)
        else:
            if length_m > 0:
                self.ed_modulation.setText(f"{length_m:.2f}")
            else:
                self.ed_modulation.setText("")
            self.spin_qty.setEnabled(True)
            self._populate_config_table(self._selected_product)

    def _reset_configurator_state(self) -> None:
        self._current_modules = []
        self._modulation_manual = False
        self.spin_distance.blockSignals(True)
        self.spin_distance.setValue(0)
        self.spin_distance.blockSignals(False)
        self.ed_modulation.setText("")
        self.spin_qty.setEnabled(True)
        self.spin_qty.setValue(1)

    def _clear_after_add(self) -> None:
        self._selected_product_id = None
        self._selected_product = None
        self._selected_equipment_name = ""
        self.lbl_product_name.setText("--")
        self.lbl_product_lead.setText("")
        self.config_table.setRowCount(0)
        self.spin_qty.blockSignals(True)
        self.spin_qty.setValue(0)
        self.spin_qty.blockSignals(False)
        self.spin_distance.blockSignals(True)
        self.spin_distance.setValue(0)
        self.spin_distance.blockSignals(False)
        self.ed_modulation.setText("--")

    def _apply_manual_modules(self, modules: List[float]) -> None:
        if not self._selected_product:
            return
        self._modulation_manual = True
        self._current_modules = modules
        mod_text = " + ".join(f"{m:g}" for m in modules)
        self.ed_modulation.setText(mod_text)
        self.spin_qty.blockSignals(True)
        self.spin_qty.setValue(len(modules))
        self.spin_qty.blockSignals(False)
        self.spin_qty.setEnabled(False)
        total_len = sum(modules)
        self.spin_distance.blockSignals(True)
        self.spin_distance.setValue(total_len)
        self.spin_distance.blockSignals(False)
        self._populate_config_table_for_modules(self._selected_product, modules)

    def _option_items(self) -> List[Dict[str, Any]]:
        return [
            {
                "key": "ESTIBA",
                "label": "ESTIBA +3%",
                "cost": self._option_costs.get("ESTIBA", 0.0),
                "checked": self.chk_estiba.isChecked(),
            },
            {
                "key": "TRANSPORTE",
                "label": "TRANSPORTE",
                "cost": self._option_costs.get("TRANSPORTE", 0.0),
                "checked": self.chk_transporte.isChecked(),
            },
            {
                "key": "INSTALACION",
                "label": "INSTALACION",
                "cost": self._option_costs.get("INSTALACION", 0.0),
                "checked": self.chk_instalacion.isChecked(),
            },
        ]

    def _option_rows_from_state(self) -> List[Dict[str, Any]]:
        rows = []
        for opt in self._option_items():
            if not opt["checked"]:
                continue
            key = opt["key"].lower()
            rows.append({
                "desc": opt["label"],
                "unit_price": opt["cost"],
                "qty": 1,
                "kind": f"option_{key}",
                "option_key": opt["key"],
                "editable_qty": False,
            })
        return rows

    def _find_total_row(self) -> int:
        for row in range(self.config_table.rowCount()):
            if self._is_total_row(row):
                return row
        return -1

    def _find_option_row(self, option_key: str) -> int:
        for row in range(self.config_table.rowCount()):
            if self._is_total_row(row):
                continue
            desc_item = self.config_table.item(row, 0)
            if not desc_item:
                continue
            meta = desc_item.data(Qt.UserRole) or {}
            if meta.get("kind", "").startswith("option") and meta.get("option_key") == option_key:
                return row
        return -1

    def _insert_option_row(self, row_index: int, option: Dict[str, Any]) -> None:
        self.config_table.insertRow(row_index)
        desc_text = str(option.get("label", "")).upper()
        desc_item = QTableWidgetItem(desc_text)
        desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)
        desc_item.setData(Qt.UserRole, {
            "kind": f"option_{option.get('key', '').lower()}",
            "option_key": option.get("key"),
        })
        price_item = QTableWidgetItem(self._format_money(option.get("cost", 0)))
        price_item.setFlags(price_item.flags() & ~Qt.ItemIsEditable)
        qty_item = QTableWidgetItem("1")
        qty_item.setFlags(qty_item.flags() & ~Qt.ItemIsEditable)
        sub_item = QTableWidgetItem("0")
        sub_item.setFlags(sub_item.flags() & ~Qt.ItemIsEditable)
        for item in (price_item, qty_item, sub_item):
            item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.config_table.setItem(row_index, 0, desc_item)
        self.config_table.setItem(row_index, 1, price_item)
        self.config_table.setItem(row_index, 2, qty_item)
        self.config_table.setItem(row_index, 3, sub_item)
        self._update_row_subtotal(row_index)

    def _sync_option_rows(self) -> None:
        total_row = self._find_total_row()
        if total_row < 0:
            return
        desired = [opt for opt in self._option_items() if opt["checked"]]
        desired_keys = {opt["key"] for opt in desired}
        self._table_updating = True
        row = total_row - 1
        while row >= 0:
            desc_item = self.config_table.item(row, 0)
            if desc_item:
                meta = desc_item.data(Qt.UserRole) or {}
                if meta.get("kind", "").startswith("option") and meta.get("option_key") not in desired_keys:
                    self.config_table.removeRow(row)
                    total_row -= 1
            row -= 1
        insert_at = total_row
        for opt in desired:
            if self._find_option_row(opt["key"]) < 0:
                self._insert_option_row(insert_at, opt)
                insert_at += 1
                total_row += 1
        self._table_updating = False
        self._recalc_total_row()
        self._auto_height_config_table()

    def _update_option_costs(self) -> None:
        self.lbl_estiba_cost.setText(
            self._format_money(self._option_costs.get("ESTIBA", 0.0)) if self.chk_estiba.isChecked() else "--"
        )
        self.lbl_transporte_cost.setText(
            self._format_money(self._option_costs.get("TRANSPORTE", 0.0)) if self.chk_transporte.isChecked() else "--"
        )
        self.lbl_instalacion_cost.setText(
            self._format_money(self._option_costs.get("INSTALACION", 0.0)) if self.chk_instalacion.isChecked() else "--"
        )

    def _is_vitrina(self) -> bool:
        name = self._selected_equipment_name or self.lbl_product_name.text()
        return "VITRINA" in self._norm_text(name)

    def _vitrina_parts(self, module_count: int) -> List[Dict[str, Any]]:
        count = max(1, module_count)
        return [
            {"desc": "RESISTENCIA MESON", "unit_price": 620000, "qty": count, "kind": "base", "editable_qty": True},
            {"desc": "VIDRIO CURVO", "unit_price": 840000, "qty": count, "kind": "base", "editable_qty": True},
            {"desc": "BANDEJA DRENADO", "unit_price": 310000, "qty": count, "kind": "base", "editable_qty": True},
            {"desc": "MARCO ALUMINIO", "unit_price": 520000, "qty": count, "kind": "base", "editable_qty": True},
        ]

    def _recalc_total_row(self) -> None:
        total_row = self.config_table.rowCount() - 1
        if total_row < 0:
            return
        if not self._is_total_row(total_row):
            return
        total = 0.0
        for r in range(total_row):
            sub_item = self.config_table.item(r, 3)
            if not sub_item:
                continue
            total += self._parse_number(sub_item.text())
        total_item = self.config_table.item(total_row, 3)
        if total_item:
            total_item.setText(self._format_money(total))

    def _is_total_row(self, row: int) -> bool:
        desc_item = self.config_table.item(row, 0)
        if not desc_item:
            return False
        return desc_item.data(Qt.UserRole) == "total"

    @staticmethod
    def _format_money(value: float) -> str:
        return f"{format_cop(float(value or 0))} $"

    @staticmethod
    def _exclude_equipment(name: str) -> bool:
        low = str(name or "").lower()
        if ("2.5" in low or "2,5" in low or "2.50" in low or "2,50" in low) and ("autoservicio" in low or "argos" in low):
            return True
        return False

    @staticmethod
    def _force_upper(le: QLineEdit, text: str) -> None:
        up = (text or "").upper()
        if up != text:
            pos = le.cursorPosition()
            le.blockSignals(True)
            le.setText(up)
            le.setCursorPosition(pos)
            le.blockSignals(False)

    def _on_new_project(self) -> None:
        seed_currency = self.project_seed.get("currency", "COP") if isinstance(self.project_seed, dict) else "COP"
        seed_trm = self.project_seed.get("trm", 0) if isinstance(self.project_seed, dict) else 0
        self.project = normalize_project({"currency": seed_currency, "trm": seed_trm}, self.rules)
        self.cart = []
        self._load_project_into_form()
        self._refresh_cart()

    def _on_save_draft(self) -> None:
        self._sync_project_from_form()
        path = save_project_draft(self.project, self.cart, base_dir=self.base_dir)
        self._info(f"Borrador guardado en:\n{path}")

    def _on_export_pdf(self) -> None:
        self._sync_project_from_form()
        totals = calculate_totals(self.cart, self.project, self.rules)
        path = export_pdf(self.project, self.cart, totals, self.rules, base_dir=self.base_dir)
        self._info(f"PDF exportado en:\n{path}")

    def _warn(self, text: str) -> None:
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle("Cotizador")
        box.setText(text)
        box.exec()

    def _info(self, text: str) -> None:
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Information)
        box.setWindowTitle("Cotizador")
        box.setText(text)
        box.exec()
