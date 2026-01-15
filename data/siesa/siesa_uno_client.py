# siesa_uno_client.py
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Literal, Dict, Tuple
import socket
import subprocess

pyodbc = None
_pyodbc_error: str | None = None
try:
    import pyodbc as _pyodbc  # type: ignore
    pyodbc = _pyodbc
except Exception as e:
    _pyodbc_error = str(e)


CIA_MAP: Dict[str, int] = {"Weston": 1, "WBR": 5, "TEKOAM": 6}

# Consulta base equivalente a fnInventario(IdCia) en Power Query.
# Ajusta el schema si el nombre real difiere (dbo.fnInventario).
SQL_FN_INVENTARIO = r"""
SELECT
    Item,
    Referencia,
    [Desc corta item] AS Desc_corta_item,
    [Desc item] AS Desc_item,
    [Unidad inventario] AS Unidad_inventario,
    [Unidad orden] AS Unidad_orden,
    [Tipo inv serv] AS Tipo_inv_serv,
    [Cant existencia] AS Cant_existencia,
    [Cant requerida] AS Cant_requerida,
    [Cant OC/OP] AS Cant_OC_OP,
    [Fecha creacion] AS Fecha_creacion,
    Estado,
    Notas
FROM dbo.fnInventario(?);
"""


@dataclass(frozen=True)
class SiesaConfig:
    server: str = "192.168.155.93"
    database: str = "UNOEE"
    auth: Literal["windows", "sql", "credman"] = "windows"
    user: Optional[str] = None
    password: Optional[str] = None
    cred_target: Optional[str] = "CalvoSiesaUNOEE"
    cred_user: Optional[str] = "sa"
    driver: str = "ODBC Driver 18 for SQL Server"
    encrypt: bool = False
    trust_server_certificate: bool = True


class SiesaUNOClient:
    def __init__(self, cfg: SiesaConfig):
        self.cfg = cfg

    @staticmethod
    def _import_pyodbc():
        global pyodbc, _pyodbc_error
        if pyodbc is not None:
            return pyodbc
        try:
            import pyodbc as _pyodbc  # type: ignore
            pyodbc = _pyodbc
            _pyodbc_error = None
        except Exception as e:
            _pyodbc_error = str(e)
        return pyodbc

    @staticmethod
    def pyodbc_status() -> tuple[bool, str]:
        if pyodbc is None:
            return False, (_pyodbc_error or 'pyodbc no esta instalado.')
        return True, f"pyodbc {getattr(pyodbc, 'version', '')}".strip()

    @staticmethod
    def odbc_drivers() -> list[str]:
        if SiesaUNOClient._import_pyodbc() is None:
            return []
        return list(pyodbc.drivers())

    def _conn_str(self) -> str:
        c = self.cfg
        base = f"DRIVER={{{c.driver}}};SERVER={c.server};DATABASE={c.database};"

        tls = ""
        if c.encrypt:
            tls += "Encrypt=yes;"
            if c.trust_server_certificate:
                tls += "TrustServerCertificate=yes;"
        else:
            tls += "Encrypt=no;"
            if c.trust_server_certificate:
                tls += "TrustServerCertificate=yes;"

        if c.auth == "windows":
            return base + "Trusted_Connection=yes;" + tls

        if c.auth == "credman":
            user, password = self._read_credman()
            return base + f"UID={user};PWD={password};" + tls

        if not c.user or not c.password:
            raise ValueError("auth='sql' requiere user y password.")
        return base + f"UID={c.user};PWD={c.password};" + tls

    def _read_credman(self) -> Tuple[str, str]:
        target = self.cfg.cred_target or "CalvoSiesaUNOEE"
        user_hint = self.cfg.cred_user or None
        errors: list[str] = []

        try:
            import keyring  # type: ignore
            if user_hint:
                pwd = keyring.get_password(target, user_hint)
                if pwd:
                    return user_hint, pwd
            get_cred = getattr(keyring, "get_credential", None)
            if get_cred:
                cred = get_cred(target, user_hint)
                if cred and cred.username and cred.password:
                    return cred.username, cred.password
        except Exception as e:
            errors.append(f"keyring: {e}")

        try:
            import win32cred  # type: ignore
            cred = win32cred.CredRead(target, win32cred.CRED_TYPE_GENERIC, 0)
            username = cred.get("UserName") or user_hint
            blob = cred.get("CredentialBlob")
            password = ""
            if isinstance(blob, bytes):
                try:
                    password = blob.decode("utf-16-le")
                except Exception:
                    password = blob.decode("utf-8", errors="ignore")
            elif blob is not None:
                password = str(blob)
            if username and password:
                return username, password
        except Exception as e:
            errors.append(f"win32cred: {e}")

        detail = f" Detalles: {', '.join(errors)}" if errors else ""
        raise ValueError(
            "No se pudo leer la credencial CalvoSiesaUNOEE. "
            "Cree una credencial generica en Windows o instale pywin32/keyring."
            + detail
        )

    # --------- Diagnóstico ---------
    def ping_host(self, timeout_ms: int = 900) -> Tuple[bool, str]:
        try:
            r = subprocess.run(
                ["ping", "-n", "1", "-w", str(timeout_ms), self.cfg.server],
                capture_output=True,
                text=True,
                check=False,
            )
            ok = (r.returncode == 0)
            return ok, (r.stdout.strip() or r.stderr.strip())
        except Exception as e:
            return False, str(e)

    def tcp_port_open(self, port: int = 1433, timeout_s: float = 1.5) -> Tuple[bool, str]:
        try:
            with socket.create_connection((self.cfg.server, port), timeout=timeout_s):
                return True, f"TCP {self.cfg.server}:{port} OK"
        except Exception as e:
            return False, f"TCP {self.cfg.server}:{port} FAIL -> {e}"

    def test_sql_login(self, timeout_s: int = 5) -> Tuple[bool, str]:
        if self._import_pyodbc() is None or pyodbc is None:
            return False, f"pyodbc no esta disponible: {_pyodbc_error or 'instale pyodbc'}"
        try:
            with pyodbc.connect(self._conn_str(), timeout=timeout_s) as _:
                return True, "Conexion SQL OK"
        except Exception as e:
            return False, f"Conexion SQL FAIL -> {e}"

    def fetch_df(self, sql: str, params: list | None = None):
        if self._import_pyodbc() is None or pyodbc is None:
            raise RuntimeError(f"pyodbc no esta disponible: {_pyodbc_error or 'instale pyodbc'}")
        import pandas as pd  # local import para no exigirlo en diagnostico
        with pyodbc.connect(self._conn_str(), timeout=10) as conn:
            conn.timeout = 120
            return pd.read_sql(sql, conn, params=params or [])
