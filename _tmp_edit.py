from pathlib import Path
p=Path('gui/pages/cargas/industrial_page.py')
text=p.read_text()
old="        # barra superior de acciones\n\n        actions = QHBoxLayout()\n\n        self.btn_save_room = QPushButton(\"GUARDAR CUARTO\")\n\n        self.btn_dup_room = QPushButton(\"DUPLICAR CUARTO\")\n\n        self.btn_clean_room = QPushButton(\"LIMPIAR CUARTO\")\n\n        self.btn_golden_room = QPushButton(\"CARGAR GOLDEN\")\n\n        self.btn_save_room.clicked.connect(self._save_editor_to_room)\n\n        self.btn_dup_room.clicked.connect(self._duplicate_room)\n\n        self.btn_clean_room.clicked.connect(self._clear_editor)\n\n        self.btn_golden_room.clicked.connect(self._load_golden)\n\n        for b in (self.btn_save_room, self.btn_dup_room, self.btn_clean_room, self.btn_golden_room):\n\n            actions.addWidget(b)\n\n        actions.addStretch()\n\n        vbox.addLayout(actions)\n\n\n\n        tabs = QTabWidget()\n"
new="        # Acciones simplificadas: eliminadas de la parte superior\n\n        tabs = QTabWidget()\n"
if old not in text:
    raise SystemExit('old block not found')
text=text.replace(old,new)
p.write_text(text)
print('removed editor top buttons')
