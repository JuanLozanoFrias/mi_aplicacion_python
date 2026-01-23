from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from PySide6.QtCore import Qt, QRect
from PySide6.QtGui import QColor, QFont, QPainter, QPageSize, QPdfWriter


DATA_DIR = Path("data/cotizador")


def get_base_dir() -> Path:
    return Path(__file__).resolve().parents[2]


def _safe_read_json(path: Path) -> Dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _safe_number(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def _safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except Exception:
        return default


def load_catalog(base_dir: Path | None = None) -> Dict[str, Any]:
    base = base_dir or get_base_dir()
    data = _safe_read_json(base / DATA_DIR / "catalogo_demo.json")
    families = data.get("families", []) if isinstance(data, dict) else []
    products_by_id: Dict[str, Dict[str, Any]] = {}
    for fam in families:
        if not isinstance(fam, dict):
            continue
        fam_name = fam.get("name", "")
        fam_tag = fam.get("tag", "")
        fam_image = fam.get("image", "")
        for prod in fam.get("products", []) or []:
            if not isinstance(prod, dict):
                continue
            pid = str(prod.get("product_id", "")).strip()
            if not pid:
                continue
            entry = dict(prod)
            entry["family_name"] = fam_name
            entry["family_tag"] = fam_tag
            entry["family_image"] = fam_image
            products_by_id[pid] = entry
    data["_products_by_id"] = products_by_id
    return data


def load_rules(base_dir: Path | None = None) -> Dict[str, Any]:
    base = base_dir or get_base_dir()
    return _safe_read_json(base / DATA_DIR / "reglas_demo.json")


def load_clients(base_dir: Path | None = None) -> Dict[str, Any]:
    base = base_dir or get_base_dir()
    return _safe_read_json(base / DATA_DIR / "clientes_demo.json")


def load_project_seed(base_dir: Path | None = None) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    base = base_dir or get_base_dir()
    data = _safe_read_json(base / DATA_DIR / "proyecto_demo.json")
    if not isinstance(data, dict):
        return {}, []
    project = data.get("project", {}) if isinstance(data.get("project"), dict) else {}
    seed = data.get("cart_seed", []) if isinstance(data.get("cart_seed"), list) else []
    return project, seed


def normalize_project(project: Dict[str, Any], rules: Dict[str, Any]) -> Dict[str, Any]:
    defaults = rules.get("defaults", {}) if isinstance(rules, dict) else {}
    result = dict(project or {})
    if "discount_pct" not in result:
        result["discount_pct"] = defaults.get("discount_pct", 0.0)
    if "price_increase_pct" not in result:
        result["price_increase_pct"] = defaults.get("price_increase_pct", 0.0)
    if "estiba" not in result:
        result["estiba"] = True
    if "trm" not in result:
        result["trm"] = 0.0
    if "options" not in result or not isinstance(result.get("options"), dict):
        result["options"] = {}
    return result


def _build_product_index(catalog: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    return catalog.get("_products_by_id", {}) if isinstance(catalog, dict) else {}


def build_cart_from_seed(seed: List[Dict[str, Any]], catalog: Dict[str, Any]) -> List[Dict[str, Any]]:
    products = _build_product_index(catalog)
    cart: List[Dict[str, Any]] = []
    for entry in seed or []:
        if not isinstance(entry, dict):
            continue
        pid = str(entry.get("product_id", "")).strip()
        if not pid or pid not in products:
            continue
        product = products[pid]
        qty = max(1, _safe_int(entry.get("qty", 1), 1))
        addons = entry.get("selected_addons", []) or []
        line_items: List[Dict[str, Any]] = []
        for base in product.get("base_modules", []) or []:
            line_items.append({
                "desc": base.get("desc", ""),
                "unit_price": _safe_number(base.get("unit_price", 0)),
                "qty": qty,
                "kind": "base",
            })
        optional_addons = {a.get("desc", ""): a for a in product.get("optional_addons", []) or []}
        for addon in addons:
            if not isinstance(addon, dict):
                continue
            desc = addon.get("desc", "")
            addon_qty = max(0, _safe_int(addon.get("qty", 0), 0))
            if addon_qty <= 0:
                continue
            price = _safe_number(optional_addons.get(desc, {}).get("unit_price", 0))
            line_items.append({
                "desc": desc,
                "unit_price": price,
                "qty": addon_qty,
                "kind": "addon",
            })
        cart.append({
            "product_id": pid,
            "product_name": product.get("name", pid),
            "qty": qty,
            "lead_time_days": _safe_int(product.get("lead_time_days", 0), 0),
            "line_items": line_items,
        })
    return cart


def _sum_line_items(cart: Iterable[Dict[str, Any]]) -> float:
    total = 0.0
    for item in cart:
        for line in item.get("line_items", []) or []:
            total += _safe_number(line.get("unit_price", 0)) * _safe_number(line.get("qty", 0))
    return total


def _max_lead_time(cart: Iterable[Dict[str, Any]]) -> int:
    lead = 0
    for item in cart:
        lead = max(lead, _safe_int(item.get("lead_time_days", 0), 0))
    return lead


def round_amount(value: float, step: float, mode: str = "nearest") -> float:
    if step <= 0:
        return value
    if mode != "nearest":
        return value
    return round(value / step) * step


def calculate_totals(
    cart: List[Dict[str, Any]],
    project: Dict[str, Any],
    rules: Dict[str, Any],
) -> Dict[str, Any]:
    project = normalize_project(project, rules)
    subtotal_base = 0.0
    recargo_from_items = 0.0
    for item in cart:
        for line in item.get("line_items", []) or []:
            line_total = _safe_number(line.get("unit_price", 0)) * _safe_number(line.get("qty", 0))
            if str(line.get("kind", "")).lower() == "option_estiba":
                recargo_from_items += line_total
            else:
                subtotal_base += line_total
    defaults = rules.get("defaults", {}) if isinstance(rules, dict) else {}
    estiba_pct = _safe_number(defaults.get("estiba_surcharge_pct", 0.0))
    estiba = bool(project.get("estiba", False))
    recargo = recargo_from_items
    if recargo <= 0 and estiba:
        recargo = subtotal_base * (estiba_pct / 100.0)
    discount_pct = _safe_number(project.get("discount_pct", defaults.get("discount_pct", 0.0)))
    descuento = (subtotal_base + recargo) * (discount_pct / 100.0)
    iva_pct = 0.0
    iva_apply = False
    for tax in rules.get("taxes", []) or []:
        if str(tax.get("name", "")).upper() == "IVA":
            iva_pct = _safe_number(tax.get("pct", 0.0))
            iva_apply = bool(tax.get("apply", False))
            break
    iva = (subtotal_base + recargo - descuento) * (iva_pct / 100.0) if iva_apply else 0.0
    total_final = subtotal_base + recargo - descuento + iva
    rounding = rules.get("rounding", {}) if isinstance(rules, dict) else {}
    rounded_total = round_amount(
        total_final,
        _safe_number(rounding.get("step", 0.0)),
        str(rounding.get("mode", "nearest")).lower(),
    )
    trm = _safe_number(project.get("trm", 0.0))
    total_usd = rounded_total / trm if trm > 0 else 0.0
    margin_cfg = rules.get("margin_demo", {}) if isinstance(rules, dict) else {}
    cost_pct = _safe_number(margin_cfg.get("cost_assumption_pct_of_sales", 0.0))
    cost_value = rounded_total * (cost_pct / 100.0)
    margin_value = rounded_total - cost_value
    margin_pct = (margin_value / rounded_total * 100.0) if rounded_total > 0 else 0.0
    estiba_pct_applied = estiba_pct if estiba and recargo_from_items <= 0 else 0.0
    return {
        "subtotal_base": subtotal_base,
        "recargo_estiba": recargo,
        "descuento": descuento,
        "iva": iva,
        "total_final": total_final,
        "total_final_rounded": rounded_total,
        "total_usd": total_usd,
        "lead_time_days": _max_lead_time(cart),
        "margin_pct": margin_pct,
        "margin_value": margin_value,
        "iva_pct": iva_pct if iva_apply else 0.0,
        "discount_pct": discount_pct,
        "estiba_pct": estiba_pct_applied,
    }


def format_cop(value: float) -> str:
    return f"{value:,.0f}".replace(",", ".")


def format_usd(value: float) -> str:
    return f"{value:,.2f}"


def _slugify(text: str) -> str:
    cleaned = "".join(ch for ch in (text or "") if ch.isalnum() or ch in (" ", "_", "-")).strip()
    cleaned = cleaned.replace(" ", "_")
    return cleaned or "cotizacion"


def save_project_draft(
    project: Dict[str, Any],
    cart: List[Dict[str, Any]],
    out_dir: Path | None = None,
    base_dir: Path | None = None,
) -> Path:
    base = base_dir or get_base_dir()
    target_dir = out_dir or (base / "data" / "proyectos" / "cotizador")
    target_dir.mkdir(parents=True, exist_ok=True)
    name = _slugify(str(project.get("project_name", "cotizacion")))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = target_dir / f"{stamp}_{name}.json"
    payload = {
        "project": project,
        "cart": cart,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _build_export_rows(cart: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for item in cart or []:
        product = item.get("product_name") or item.get("product_id") or "ITEM"
        for line in item.get("line_items", []) or []:
            desc = str(line.get("desc", "")).strip()
            label = f"{product} - {desc}" if desc else product
            rows.append({
                "desc": label,
                "qty": _safe_number(line.get("qty", 0)),
                "unit_price": _safe_number(line.get("unit_price", 0)),
            })
    return rows


def export_pdf(
    project: Dict[str, Any],
    cart: List[Dict[str, Any]],
    totals: Dict[str, Any],
    rules: Dict[str, Any],
    out_dir: Path | None = None,
    base_dir: Path | None = None,
) -> Path:
    base = base_dir or get_base_dir()
    export_cfg = rules.get("export", {}) if isinstance(rules, dict) else {}
    target_dir = out_dir or (base / str(export_cfg.get("output_folder", "exports/cotizaciones")))
    target_dir.mkdir(parents=True, exist_ok=True)
    name = _slugify(str(project.get("project_name", "cotizacion")))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = target_dir / f"{stamp}_{name}.pdf"

    writer = QPdfWriter(str(path))
    writer.setPageSize(QPageSize(QPageSize.A4))
    writer.setResolution(120)
    painter = QPainter(writer)

    page_rect = writer.pageLayout().fullRectPixels(writer.resolution())
    page_w = int(page_rect.width())
    page_h = int(page_rect.height())
    margin = 48
    header_h = 70
    footer_reserved = 34

    title_font = QFont()
    title_font.setPointSize(18)
    title_font.setBold(True)

    header_font = QFont()
    header_font.setPointSize(10)
    header_font.setBold(True)

    body_font = QFont()
    body_font.setPointSize(9)

    small_font = QFont()
    small_font.setPointSize(8)

    light_gray = QColor("#e2e8f5")
    header_bg = QColor("#f1f5f9")
    accent = QColor("#0f62fe")
    stripe = QColor("#f8fafc")
    group_bg = QColor("#e8f2ff")
    text_dark = QColor("#0f172a")
    text_muted = QColor("#64748b")

    date_str = datetime.now().strftime("%d/%m/%Y")
    quote_id = datetime.now().strftime("COT-%Y%m%d-%H%M")

    def money(value: float) -> str:
        val = float(value or 0.0)
        s = f"{val:,.2f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"$ {s}"

    page_num = 1

    def draw_header() -> None:
        painter.fillRect(0, 0, page_w, 12, accent)
        painter.setPen(text_dark)
        painter.setFont(title_font)
        painter.drawText(margin, 40, "WESTON")
        painter.setFont(body_font)
        painter.drawText(margin, 58, "DIVISION DE INGENIERIA")
        painter.setFont(header_font)
        painter.setPen(accent)
        painter.drawText(page_w - margin - 200, 36, "COTIZACION")
        painter.setPen(text_dark)
        painter.setFont(small_font)
        painter.drawText(page_w - margin - 200, 52, f"FECHA: {date_str}")
        painter.drawText(page_w - margin - 200, 64, f"COT: {quote_id}")
        painter.setPen(light_gray)
        painter.drawLine(margin, header_h, page_w - margin, header_h)

    def draw_footer(current_page: int) -> None:
        painter.setFont(small_font)
        painter.setPen(text_muted)
        footer = (
            "CALLE 16 # 65B-82 | BOGOTA, COLOMBIA | PBX: +57 (601) 290 7700 | "
            "WESTON@WESTON.COM.CO | WWW.WESTON.COM.CO"
        )
        footer_rect = QRect(margin, page_h - margin + 6, page_w - margin * 2, 14)
        painter.drawText(footer_rect, Qt.AlignCenter | Qt.AlignVCenter, f"{footer} | PAG. {current_page}")

    def draw_project_block(y_pos: int) -> int:
        client = str(project.get("client_name", "")) or str(project.get("client_id", ""))
        validity = _safe_int(rules.get("defaults", {}).get("validity_days", 0), 0)
        nit = str(project.get("nit", "") or "N/A")
        phone = str(project.get("phone", "") or "N/A")
        email = str(project.get("email", "") or "N/A")
        branch = str(project.get("branch", "") or "No aplica")
        ship_point = str(project.get("shipping_point", "") or "No aplica")
        ship_addr = str(project.get("shipping_address", "") or "No especificada")
        left = [
            f"CLIENTE: {client}",
            f"CONTACTO: {project.get('contact', '') or 'N/A'}",
            f"NIT/CC: {nit}",
            f"TELEFONO: {phone}",
            f"CORREO: {email}",
            f"CIUDAD: {project.get('city', '')}",
        ]
        right = [
            f"COTIZACION NRO: {quote_id}",
            f"FECHA: {date_str}",
            f"SUCURSAL: {branch}",
            f"PUNTO DE ENVIO: {ship_point}",
            f"DIRECCION ENVIO: {ship_addr}",
            f"MONEDA: {project.get('currency', 'COP')}",
            f"TRM: {project.get('trm', '')}",
            f"VIGENCIA: {validity} DIAS",
        ]
        row_h = 14
        rows = max(len(left), len(right))
        box_h = 20 + rows * row_h + 12
        box_w = page_w - margin * 2
        painter.setPen(light_gray)
        painter.drawRect(margin, y_pos, box_w, box_h)
        painter.setPen(text_dark)
        painter.setFont(header_font)
        painter.drawText(margin + 8, y_pos + 16, "DATOS DEL CLIENTE Y COTIZACION")
        painter.setFont(body_font)
        left_x = margin + 8
        right_x = margin + box_w // 2 + 8
        base_y = y_pos + 32
        for idx, line in enumerate(left):
            painter.drawText(left_x, base_y + idx * row_h, line)
        for idx, line in enumerate(right):
            painter.drawText(right_x, base_y + idx * row_h, line)
        return y_pos + box_h + 14

    def draw_table_header(y_pos: int, table_x: int, table_w: int, cols: Dict[str, int]) -> int:
        painter.fillRect(table_x, y_pos, table_w, 20, header_bg)
        painter.setFont(header_font)
        painter.setPen(text_dark)
        painter.drawText(QRect(cols["item_x"], y_pos, cols["item_w"], 20), Qt.AlignLeft | Qt.AlignVCenter, "ITEM")
        painter.drawText(QRect(cols["desc_x"], y_pos, cols["desc_w"], 20), Qt.AlignLeft | Qt.AlignVCenter, "DESCRIPCION DEL MODELO")
        painter.drawText(QRect(cols["qty_x"], y_pos, cols["qty_w"], 20), Qt.AlignRight | Qt.AlignVCenter, "CANT.")
        painter.drawText(QRect(cols["unit_x"], y_pos, cols["unit_w"], 20), Qt.AlignRight | Qt.AlignVCenter, "VALOR UNIT.")
        painter.drawText(QRect(cols["total_x"], y_pos, cols["total_w"], 20), Qt.AlignRight | Qt.AlignVCenter, "TOTAL")
        painter.setPen(light_gray)
        painter.drawLine(table_x, y_pos + 20, table_x + table_w, y_pos + 20)
        painter.setPen(text_dark)
        return y_pos + 24

    def new_page(with_table_header: bool = False, cols: Dict[str, int] | None = None) -> int:
        nonlocal page_num
        draw_footer(page_num)
        writer.newPage()
        page_num += 1
        draw_header()
        y_pos = header_h + margin
        if with_table_header and cols:
            y_pos = draw_table_header(y_pos, margin, page_w - margin * 2, cols)
        return y_pos

    def ensure_space(y_pos: int, needed: int, cols: Dict[str, int] | None = None, table_mode: bool = False) -> int:
        if y_pos + needed < page_h - margin - footer_reserved:
            return y_pos
        return new_page(with_table_header=table_mode, cols=cols)

    rows: List[Dict[str, Any]] = []
    for idx, item in enumerate(cart or [], start=1):
        product = str(item.get("product_name") or item.get("product_id") or "ITEM").upper()
        qty_equipo = _safe_int(item.get("qty", 1), 1)
        product_total = 0.0
        for line in item.get("line_items", []) or []:
            product_total += _safe_number(line.get("unit_price", 0)) * _safe_number(line.get("qty", 0))
        rows.append({
            "item": f"{idx}",
            "desc": f"{product}",
            "qty": qty_equipo,
            "unit": "",
            "total": product_total,
            "is_group": True,
        })
        for sub_idx, line in enumerate(item.get("line_items", []) or [], start=1):
            desc = str(line.get("desc", "")).upper()
            qty = _safe_number(line.get("qty", 0))
            unit = _safe_number(line.get("unit_price", 0))
            rows.append({
                "item": f"{idx}.{sub_idx}",
                "desc": f"  {desc}",
                "qty": qty,
                "unit": unit,
                "total": unit * qty,
                "is_group": False,
            })

    draw_header()
    y = header_h + margin
    y = draw_project_block(y)

    table_x = margin
    table_w = page_w - margin * 2
    cols = {
        "item_x": table_x,
        "item_w": int(table_w * 0.08),
    }
    cols["desc_x"] = cols["item_x"] + cols["item_w"]
    cols["desc_w"] = int(table_w * 0.47)
    cols["qty_x"] = cols["desc_x"] + cols["desc_w"]
    cols["qty_w"] = int(table_w * 0.10)
    cols["unit_x"] = cols["qty_x"] + cols["qty_w"]
    cols["unit_w"] = int(table_w * 0.17)
    cols["total_x"] = cols["unit_x"] + cols["unit_w"]
    cols["total_w"] = table_w - cols["item_w"] - cols["desc_w"] - cols["qty_w"] - cols["unit_w"]

    y = draw_table_header(y, table_x, table_w, cols)
    row_h = 20
    for idx, row in enumerate(rows):
        y = ensure_space(y, row_h + 80, cols=cols, table_mode=True)
        if row.get("is_group"):
            painter.fillRect(table_x, y, table_w, row_h, group_bg)
        elif idx % 2 == 0:
            painter.fillRect(table_x, y, table_w, row_h, stripe)
        painter.setFont(header_font if row.get("is_group") else body_font)
        painter.setPen(text_dark)
        painter.drawText(QRect(cols["item_x"], y, cols["item_w"], row_h), Qt.AlignLeft | Qt.AlignVCenter, row.get("item", ""))
        painter.drawText(QRect(cols["desc_x"], y, cols["desc_w"], row_h), Qt.AlignLeft | Qt.AlignVCenter, row.get("desc", ""))
        qty = row.get("qty", "")
        unit = row.get("unit", "")
        total = row.get("total", "")
        if qty != "":
            painter.drawText(QRect(cols["qty_x"], y, cols["qty_w"], row_h), Qt.AlignRight | Qt.AlignVCenter, f"{qty:.0f}")
        if unit != "":
            painter.drawText(QRect(cols["unit_x"], y, cols["unit_w"], row_h), Qt.AlignRight | Qt.AlignVCenter, money(unit))
        if total != "":
            painter.drawText(QRect(cols["total_x"], y, cols["total_w"], row_h), Qt.AlignRight | Qt.AlignVCenter, money(total))
        y += row_h

    y = ensure_space(y + 12, 140, cols=cols, table_mode=False)
    summary_w = 270
    summary_x = page_w - margin - summary_w
    summary_y = y
    painter.fillRect(summary_x, summary_y, summary_w, 118, header_bg)
    painter.setPen(light_gray)
    painter.drawRect(summary_x, summary_y, summary_w, 118)
    painter.setPen(text_dark)
    painter.setFont(header_font)
    painter.drawText(summary_x + 10, summary_y + 18, "RESUMEN")
    painter.setFont(body_font)
    lines = [
        ("SUBTOTAL", money(totals.get("subtotal_base", 0.0))),
        ("RECARGO ESTIBA", money(totals.get("recargo_estiba", 0.0))),
        ("DESCUENTO", money(totals.get("descuento", 0.0))),
        ("IVA", money(totals.get("iva", 0.0))),
    ]
    line_y = summary_y + 36
    for label, val in lines:
        painter.drawText(QRect(summary_x + 10, line_y, summary_w - 20, 14), Qt.AlignLeft | Qt.AlignVCenter, label)
        painter.drawText(QRect(summary_x + 10, line_y, summary_w - 20, 14), Qt.AlignRight | Qt.AlignVCenter, val)
        line_y += 14
    painter.setFont(header_font)
    painter.setPen(accent)
    painter.drawText(QRect(summary_x + 10, line_y + 4, summary_w - 20, 16), Qt.AlignLeft | Qt.AlignVCenter, "TOTAL")
    painter.drawText(
        QRect(summary_x + 10, line_y + 4, summary_w - 20, 16),
        Qt.AlignRight | Qt.AlignVCenter,
        money(totals.get("total_final_rounded", 0.0)),
    )
    painter.setPen(text_dark)

    usd_total = _safe_number(totals.get("total_usd", 0.0))
    if usd_total > 0:
        line_y += 22
        painter.setFont(body_font)
        painter.drawText(QRect(summary_x + 8, line_y, summary_w - 16, 14), Qt.AlignLeft | Qt.AlignVCenter, "TOTAL USD")
        painter.drawText(
            QRect(summary_x + 8, line_y, summary_w - 16, 14),
            Qt.AlignRight | Qt.AlignVCenter,
            f"{format_usd(usd_total)}",
        )

    y = summary_y + 132
    y = ensure_space(y, 60, cols=cols, table_mode=False)
    notes = str(project.get("notes", "")).strip()
    validity = _safe_int(rules.get("defaults", {}).get("validity_days", 0), 0)
    painter.setFont(body_font)
    painter.setPen(text_dark)
    painter.drawText(margin, y, f"VIGENCIA: {validity} DIAS")
    y += 16
    painter.drawText(margin, y, f"NOTAS: {notes or 'N/A'}")
    y += 24

    def wrap_text(text: str, max_width: int) -> List[str]:
        words = text.split()
        if not words:
            return [""]
        lines: List[str] = []
        current = words[0]
        for word in words[1:]:
            candidate = f"{current} {word}"
            if painter.fontMetrics().horizontalAdvance(candidate) <= max_width:
                current = candidate
            else:
                lines.append(current)
                current = word
        lines.append(current)
        return lines

    def draw_section(y_pos: int, title: str, paragraphs: List[str]) -> int:
        y_pos = ensure_space(y_pos, 30, table_mode=False)
        painter.setFont(header_font)
        painter.setPen(text_dark)
        painter.drawText(margin, y_pos, title)
        y_pos += 16
        painter.setFont(body_font)
        max_width = page_w - margin * 2
        line_h = painter.fontMetrics().height() + 2
        for para in paragraphs:
            lines = wrap_text(para, max_width)
            for line in lines:
                y_pos = ensure_space(y_pos, line_h + 10, table_mode=False)
                painter.drawText(margin, y_pos, line)
                y_pos += line_h
            y_pos += 6
        return y_pos + 6

    terms_sections = [
        ("NOTAS", [
            "ESTA COTIZACION NO INCLUYE TRANSPORTE, MATERIALES DE INSTALACION Y/O MANO DE OBRA DE INSTALACION A MENOS QUE SE ESPECIFIQUE LO CONTRARIO.",
            "ESTA COTIZACION NO INCLUYE CAPACITACION A MENOS DE QUE ESTE COTIZADO EXPLICITAMENTE. DE LLEGAR A INCLUIR LA CAPACITACION, LOS VIATICOS CORREN POR CUENTA DEL CLIENTE.",
            "LOS VALORES SON BAJO CONDICIONES IDEALES DE FUNCIONAMIENTO. EL CLIENTE DEBE GARANTIZAR CONDICIONES AMBIENTALES ADECUADAS PARA LA OPERACION DEL EQUIPO.",
        ]),
        ("ALCANCE Y CONDICIONES GENERALES", [
            "LOS ALCANCES Y CONDICIONES TECNICAS DEL EQUIPO SE DEFINEN DE ACUERDO CON LA SOLICITUD DEL CLIENTE Y LOS ESTANDARES DE INGENIERIA DE WESTON.",
            "CUALQUIER MODIFICACION DEL PROYECTO, DIMENSIONES O ACCESORIOS DESPUES DE APROBADA LA COTIZACION PUEDE GENERAR AJUSTES EN COSTO Y PLAZOS.",
        ]),
        ("INSTALACION", [
            "WESTON SAS NO REALIZA CONEXIONES DE DESAGUES. CONTRA ORDEN DE COMPRA SE HARA ENTREGA DE LOS PLANOS CON LAS ESPECIFICACIONES.",
            "SI AL MOMENTO DE LA INSTALACION EL CLIENTE NO CUMPLE CON LAS EXIGENCIAS TECNICAS Y LOCATIVAS, WESTON PODRA REALIZARLA BAJO LAS CONDICIONES EXISTENTES SIN GARANTIA.",
            "EL CLIENTE DEBE CONTAR CON UN PROGRAMA DE MANTENIMIENTO PREVENTIVO QUE GARANTICE LAS OPTIMAS CONDICIONES DE OPERACION DE LOS EQUIPOS.",
        ]),
        ("ENTREGA Y RECIBO A SATISFACCION", [
            "LOS TIEMPOS DE ENTREGA PUEDEN VARIAR POR ESCASEZ DE MATERIA PRIMA, CONGESTION EN PUERTOS, PROBLEMAS DE TRANSPORTE O FUERZA MAYOR.",
            "AL MOMENTO DE LA ENTREGA, EL CLIENTE DEBE REPORTAR POR ESCRITO CUALQUIER DEFECTO ESTETICO O FALTA DE PARTES. DESPUES DE 15 DIAS SE ENTIENDE RECIBO A SATISFACCION.",
            "SE DEBEN MANTENER LOS PARAMETROS, REGULACIONES Y AJUSTES QUE SE HAGAN EN LA ENTREGA.",
        ]),
        ("TRM", [
            "LA PRESENTE COTIZACION SE REALIZO CON UNA TRM A LA FECHA DE EMISION. SI LA TRM AUMENTA IGUAL O SUPERIOR AL 5%, SE AJUSTARAN LOS PRECIOS.",
            "EN NINGUN CASO EL PAGO SERA LIQUIDADO CON UNA TRM INFERIOR A LA DE LA FECHA DE LA PRESENTE COTIZACION.",
        ]),
        ("GARANTIA LIMITADA WESTON", [
            "WESTON GARANTIZA EL BUEN FUNCIONAMIENTO DEL EQUIPO DURANTE UN (1) ANO A PARTIR DE LA FECHA DE ENTREGA.",
            "LOS PERIODOS DE GARANTIA SE MANTIENEN SIEMPRE QUE EL EQUIPO HAYA SIDO INSTALADO Y OPERADO SEGUN LAS RECOMENDACIONES DEL FABRICANTE.",
            "LA GARANTIA ES EX WORKS; LAS PARTES Y/O PIEZAS A REEMPLAZAR SE ENTREGAN EN LAS INSTALACIONES DE WESTON. EL COSTO DEL TRANSPORTE, SEGUROS Y MANO DE OBRA CORRE POR CUENTA DEL CLIENTE.",
        ]),
        ("EXCLUSIONES DE LA GARANTIA", [
            "PROBLEMAS ELECTRICOS DEBIDO A SUMINISTRO DE NIVELES DE TENSION DIFERENTES A LOS ESPECIFICADOS EN LA PLACA DEL EQUIPO.",
            "MERMA Y/O PERECEDEROS QUE SE PUEDAN PRESENTAR EN EL PRODUCTO ALMACENADO.",
            "DANOS POR INSTALACION O MONTAJE REALIZADO POR TERCEROS SIN SUPERVISION DE WESTON.",
            "LA GARANTIA NO CUBRE VITRINAS Y MUEBLES QUE NO CUMPLAN CON CONDICIONES DE TEMPERATURA Y HUMEDAD DEL AMBIENTE (HUMEDAD MAXIMA 55% Y TEMPERATURA MAXIMA 25 C).",
            "ADECUACIONES LOCATIVAS NECESARIAS PARA CUMPLIR CON CONDICIONES DE OPERACION (AIRE ACONDICIONADO, DUCTOS, DESAGUES, ETC.).",
        ]),
        ("EQUIPOS CON R290", [
            "PARA EQUIPOS CON REFRIGERANTE R290, EL CLIENTE DEBE TENER CONOCIMIENTO SOBRE EL MANEJO DE ESTE REFRIGERANTE Y UTILIZAR HERRAMIENTAS Y REPUESTOS APTOS.",
        ]),
        ("INVITACION A MANTENIMIENTO PREVENTIVO", [
            "WESTON RECOMIENDA CONTRATAR UN PROGRAMA DE MANTENIMIENTO PREVENTIVO PARA EXTENDER LA VIDA UTIL Y ASEGURAR EL DESEMPENO DEL EQUIPO.",
        ]),
    ]

    y = new_page()
    painter.setFont(header_font)
    painter.setPen(text_dark)
    painter.drawText(margin, header_h + margin, "CONDICIONES COMERCIALES Y GARANTIA")
    y = header_h + margin + 18

    for title, paragraphs in terms_sections:
        y = draw_section(y, title, paragraphs)

    draw_footer(page_num)

    painter.end()
    return path
