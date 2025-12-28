# -*- coding: utf-8 -*-
from pathlib import Path
from PySide6.QtWidgets import QSizePolicy
p=Path('gui/pages/cargas/industrial_page.py')
text=p.read_text(encoding='utf-8')
text=text.replace('self.ed_nombre_cuarto = QLineEdit(); self.ed_nombre_cuarto.setPlaceholderText("Nombre del cuarto")',
'self.ed_nombre_cuarto = QLineEdit(); self.ed_nombre_cuarto.setPlaceholderText("Nombre del cuarto"); self.ed_nombre_cuarto.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)')
text=text.replace('self.cb_perfil_head = QComboBox()', 'self.cb_perfil_head = QComboBox(); self.cb_perfil_head.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)')
lines=text.splitlines()
new=[]
for line in lines:
    if 'form_cuarto.addRow("LARGO (M):"' in line or 'form_cuarto.addRow("ANCHO (M):"' in line or 'form_cuarto.addRow("ALTO (M):"' in line:
        continue
    new.append(line)
p=Path('gui/pages/cargas/industrial_page.py')
p.write_text("\n".join(new),encoding='utf-8')
