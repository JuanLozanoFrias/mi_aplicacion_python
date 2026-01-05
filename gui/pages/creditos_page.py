from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)


def _repo_root() -> Path:
    # gui/pages/creditos_page.py -> repo root
    return Path(__file__).resolve().parents[2]


def _find_image(stem: str) -> Optional[Path]:
    exts = [".png", ".jpg", ".jpeg", ".bmp", ".webp"]
    root = _repo_root()
    preferred_dirs = [
        root / "resources" / "creditos",
        root / "resources",
        root / "data",
    ]
    for base in preferred_dirs:
        for ext in exts:
            p = base / f"{stem}{ext}"
            if p.exists():
                return p
        for p in base.glob(f"{stem}.*"):
            if p.suffix.lower() in exts:
                return p
    return None


class _ScaledPixmapLabel(QLabel):
    def __init__(self, pixmap: QPixmap | None, fallback_text: str) -> None:
        super().__init__()
        self._pix = pixmap if pixmap and not pixmap.isNull() else None
        self.setAlignment(Qt.AlignCenter)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMinimumSize(180, 220)
        self.setStyleSheet("background:#f8fafc;border:1px solid #e2e8f5;border-radius:10px;")
        if self._pix is None:
            self.setText(fallback_text)
            self.setStyleSheet(
                "background:#f8fafc;border:1px dashed #cbd5e1;border-radius:10px;"
                "color:#64748b;font-weight:800;"
            )

    def resizeEvent(self, event) -> None:  # noqa: N802 - Qt override
        super().resizeEvent(event)
        if self._pix is None:
            return
        dpr = self.devicePixelRatioF()
        target_w = max(1, int(self.width() * dpr))
        target_h = max(1, int(self.height() * dpr))
        max_w = self._pix.width()
        max_h = self._pix.height()
        target_w = min(target_w, max_w)
        target_h = min(target_h, max_h)
        scaled = self._pix.scaled(target_w, target_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        scaled.setDevicePixelRatio(dpr)
        self.setPixmap(scaled)


@dataclass(frozen=True)
class _PersonPhoto:
    name: str
    path: Optional[Path]


class CreditosPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(40, 40, 40, 40)
        root.setSpacing(20)

        header = QFrame()
        header.setStyleSheet("background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;")
        hl = QVBoxLayout(header)
        hl.setContentsMargins(30, 26, 30, 26)
        hl.setSpacing(6)

        title = QLabel("CRÉDITOS")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:24px;font-weight:800;color:#0f172a;")
        hl.addWidget(title)

        subtitle = QLabel("EQUIPO DE DESARROLLO")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color:#475569;font-weight:700;letter-spacing:1px;")
        hl.addWidget(subtitle)
        root.addWidget(header)

        content = QFrame()
        content.setStyleSheet("background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;")
        cl = QVBoxLayout(content)
        cl.setContentsMargins(30, 24, 30, 24)
        cl.setSpacing(14)

        row = QHBoxLayout()
        row.setSpacing(16)

        people = [
            _PersonPhoto("JUAN LOZANO", _find_image("1")),
            _PersonPhoto("JAVIER PARDO", _find_image("2")),
            _PersonPhoto("NICOLÁS LEÓN", _find_image("3")),
        ]

        for i, p in enumerate(people, start=1):
            card = QFrame()
            card.setStyleSheet("background:#ffffff;border:1px solid #e2e8f5;border-radius:12px;")
            card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            cxl = QVBoxLayout(card)
            cxl.setContentsMargins(14, 14, 14, 14)
            cxl.setSpacing(10)

            photo_pix = QPixmap(str(p.path)) if p.path else QPixmap()
            photo = _ScaledPixmapLabel(photo_pix, fallback_text=f"FOTO {i}\nNO ENCONTRADA")
            cxl.addWidget(photo, 1)

            name = QLabel(p.name)
            name.setAlignment(Qt.AlignCenter)
            name.setStyleSheet("color:#0f172a;font-weight:800;")
            cxl.addWidget(name)

            row.addWidget(card, 1)

        cl.addLayout(row)

        quote = QFrame()
        quote.setStyleSheet(
            "background:#f8fafc;border:1px solid #e2e8f5;border-radius:12px;"
        )
        ql = QVBoxLayout(quote)
        ql.setContentsMargins(18, 14, 18, 14)
        ql.setSpacing(6)

        quote_title = QLabel("DEDICATORIA")
        quote_title.setAlignment(Qt.AlignCenter)
        quote_title.setStyleSheet("color:#0f172a;font-weight:900;letter-spacing:1px;")
        ql.addWidget(quote_title)

        quote_text = QLabel(
            "DEDICAMOS ESTA PLATAFORMA A NUESTRAS FAMILIAS.\n"
            "GRACIAS POR SU PACIENCIA, APOYO Y AMOR MIENTRAS INNOVÁBAMOS.\n"
            "SI ALGUNA VEZ NOS PERDIMOS MOMENTOS BONITOS, ESTE LOGRO TAMBIÉN ES DE USTEDES."
        )
        quote_text.setAlignment(Qt.AlignCenter)
        quote_text.setWordWrap(True)
        quote_text.setStyleSheet("color:#334155;font-weight:700;")
        ql.addWidget(quote_text)

        cl.addWidget(quote)
        root.addWidget(content, 1)
        root.addStretch(1)
