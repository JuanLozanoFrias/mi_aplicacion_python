# Siesa DataHub (credenciales)

Este modulo usa **SQL Login** via Windows Credential Manager.

## Credencial requerida

Cree una credencial generica en Windows:

1) Panel de control -> Administrador de credenciales
2) Credenciales de Windows -> Agregar credencial generica
3) Target: **CalvoSiesaUNOEE**
4) Usuario: **sa**
5) Contrasena: (su clave)

Alternativa por CMD:

```
cmdkey /generic:CalvoSiesaUNOEE /user:sa /pass:SU_CLAVE
```

## Dependencias

```
pip install pyodbc pywin32
```

Alternativa:

```
pip install keyring
```

## Notas

- Servidor: 192.168.155.93
- Base: UNOEE
- Driver: ODBC Driver 18 for SQL Server
- TLS: Encrypt=no, TrustServerCertificate=yes

