from __future__ import annotations

import sys
import subprocess
import os
from pathlib import Path

from PySide6.QtCore import Qt
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
)

from logic.siesa_datahub import (
    load_hub,
    get_company_data_dir,
    set_company_data_dir,
)
from data.siesa.datahub import PackageError
from data.siesa.siesa_uno_client import SiesaConfig, SiesaUNOClient


class SiesaPage(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._loaded_once = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel("SIESA - DATAHUB")
        title.setStyleSheet("font-size:18px;font-weight:800;color:#0f172a;")
        layout.addWidget(title)

        actions = QHBoxLayout()
        actions.setSpacing(10)
        self.btn_load = QPushButton("CARGAR PAQUETE LOCAL")
        self.btn_diag = QPushButton("DIAGNOSTICO SIESA")
        self.btn_update = QPushButton("ACTUALIZAR SNAPSHOTS")
        actions.addWidget(self.btn_load)
        actions.addWidget(self.btn_diag)
        actions.addWidget(self.btn_update)
        actions.addStretch(1)
        layout.addLayout(actions)

        self.lbl_folder = QLabel()
        self.lbl_folder.setStyleSheet("color:#334155;")
        layout.addWidget(self.lbl_folder)

        info_box = QGroupBox("ESTADO DEL PAQUETE")
        info_layout = QVBoxLayout(info_box)
        info_layout.setContentsMargins(12, 10, 12, 12)
        info_layout.setSpacing(8)

        self.lbl_manifest = QLabel("Sin cargar")
        self.lbl_manifest.setStyleSheet("font-weight:700;color:#0f172a;")
        info_layout.addWidget(self.lbl_manifest)

        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMinimumHeight(180)
        info_layout.addWidget(self.txt_log)

        layout.addWidget(info_box)
        layout.addStretch(1)

        self.btn_load.clicked.connect(self._choose_and_load)
        self.btn_diag.clicked.connect(self._run_diagnostic)
        self.btn_update.clicked.connect(self._update_snapshots)

        self._update_folder_label()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._loaded_once:
            self._load_package()
            self._loaded_once = True

    def _update_folder_label(self) -> None:
        folder = get_company_data_dir()
        self.lbl_folder.setText(f"CARPETA: {folder}")

    def _append_log(self, text: str) -> None:
        self.txt_log.appendPlainText(text)

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
        self._append_log("=== DIAGNOSTICO SIESA UNO ===")
        if cfg.auth == "sql":
            auth_label = "SQL"
        elif cfg.auth == "credman":
            auth_label = "CREDMAN"
        else:
            auth_label = "WINDOWS"
        self._append_log(f"AUTH: {auth_label}")
        self._append_log(f"PYTHON: {sys.executable}")
        ok_py, msg_py = client.pyodbc_status()
        self._append_log(f"PYODBC: {'OK' if ok_py else 'FAIL'} -> {msg_py}")
        ok_ping, msg_ping = client.ping_host()
        self._append_log(f"PING: {'OK' if ok_ping else 'FAIL'} -> {msg_ping}")
        ok_tcp, msg_tcp = client.tcp_port_open(1433)
        self._append_log(f"TCP 1433: {'OK' if ok_tcp else 'FAIL'} -> {msg_tcp}")
        ok_sql, msg_sql = client.test_sql_login()
        self._append_log(f"SQL LOGIN: {'OK' if ok_sql else 'FAIL'} -> {msg_sql}")
        drivers = client.odbc_drivers()
        if drivers:
            self._append_log(f"DRIVERS ODBC: {', '.join(drivers)}")
        else:
            self._append_log("DRIVERS ODBC: No detectados")
        QMessageBox.information(self, "SIESA", "Diagnostico finalizado. Revisa el panel de estado.")

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
